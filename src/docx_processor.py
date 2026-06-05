"""Orchestrate báo giá .docx processing: AMIS merge, discount, delivery."""

from __future__ import annotations

import io
import re
from dataclasses import dataclass

from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from code_extractor import extract_product_codes
from docx_utils import (
    column_index,
    column_index_any,
    find_items_table,
    format_vnd,
    is_discount_summary_row,
    is_summary_row,
    iter_paragraph_texts,
    parse_vnd,
    set_cell_text,
)
from mock_context import build_mock_context
from placeholder_filler import replace_placeholders
from product_catalog import Product, resolve_codes
from product_rows import apply_product_images, expand_and_fill_product_rows

_CUSTOMER_ID_RE = re.compile(
    r"(?:Mã|Ma)\s*KH\s*[:：]\s*(\S+)",
    re.IGNORECASE | re.UNICODE,
)


@dataclass
class DiscountResult:
    lines_updated: int = 0
    discount_total: float = 0.0
    subtotal_before: float = 0.0
    subtotal_after: float = 0.0


@dataclass
class ProcessPreview:
    customer_id: str | None
    product_codes: list[str]
    products: list[Product]
    missing_codes: list[str]


class ProcessingError(Exception):
    def __init__(self, message: str, *, missing_codes: list[str] | None = None):
        super().__init__(message)
        self.missing_codes = missing_codes or []


def extract_customer_id(doc: Document) -> str | None:
    for text in iter_paragraph_texts(doc):
        m = _CUSTOMER_ID_RE.search(text)
        if m:
            return m.group(1).strip()
    return None


def extract_customer_id_from_bytes(data: bytes) -> str | None:
    return extract_customer_id(Document(io.BytesIO(data)))


def preview_from_bytes(data: bytes) -> ProcessPreview:
    doc = Document(io.BytesIO(data))
    codes = extract_product_codes(doc)
    products, missing = resolve_codes(codes)
    return ProcessPreview(
        customer_id=extract_customer_id(doc),
        product_codes=codes,
        products=products,
        missing_codes=missing,
    )


def _discount_column_index(table) -> int | None:
    header = table.rows[0]
    for i, cell in enumerate(header.cells):
        text = cell.text.lower()
        if ("chiết" in text or "chiet" in text) and ("khấu" in text or "khau" in text):
            return i
    return None


def _find_summary_row(table, prefix: str):
    for row in table.rows[1:]:
        if row.cells[0].text.strip().lower().startswith(prefix):
            return row
    return None


def _insert_discount_row_before(table, before_row, discount_total: float) -> None:
    thanh_tien_col = column_index(table, "thành", "tiền")
    if thanh_tien_col is None:
        thanh_tien_col = max(len(table.rows[0].cells) - 2, 0)
    ncols = len(table.rows[0].cells)
    amount = f"-{format_vnd(discount_total)}"

    new_tr = OxmlElement("w:tr")
    before_row._tr.addprevious(new_tr)
    for i in range(ncols):
        tc = OxmlElement("w:tc")
        p = OxmlElement("w:p")
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        if i == 0:
            t.text = "Chiết khấu"
        elif i == thanh_tien_col:
            t.text = amount
        else:
            t.text = ""
        r.append(t)
        p.append(r)
        tc.append(p)
        new_tr.append(tc)


def apply_discount_to_items(doc: Document, discount_percent: float) -> DiscountResult:
    result = DiscountResult()
    if discount_percent <= 0:
        return result

    table = find_items_table(doc)
    if table is None:
        return result

    qty_col = column_index_any(table, ("sl",), ("số", "lượng"))
    price_col = column_index_any(table, ("đơn", "giá"))
    total_col = column_index_any(table, ("thành", "tiền"))
    discount_col = _discount_column_index(table)

    if total_col is None:
        return result

    factor = 1.0 - (discount_percent / 100.0)
    line_totals_after: list[float] = []

    for r in range(1, len(table.rows)):
        row = table.rows[r]
        if is_summary_row(row) or is_discount_summary_row(row):
            break

        qty = parse_vnd(row.cells[qty_col].text) if qty_col is not None else None
        unit = parse_vnd(row.cells[price_col].text) if price_col is not None else None
        line_total = parse_vnd(row.cells[total_col].text)

        if qty is None and unit is None and line_total is None:
            continue

        if line_total is not None:
            before = line_total
        elif qty is not None and unit is not None:
            before = qty * unit
        else:
            continue

        after = before * factor
        line_discount = before - after

        result.subtotal_before += before
        result.subtotal_after += after
        result.discount_total += line_discount
        line_totals_after.append(after)

        set_cell_text(row.cells[total_col], format_vnd(after))
        if discount_col is not None:
            set_cell_text(row.cells[discount_col], format_vnd(line_discount))
        result.lines_updated += 1

    if result.lines_updated == 0:
        return result

    subtotal_after = sum(line_totals_after)
    cong_row = _find_summary_row(table, "cộng tiền")
    if cong_row is not None:
        current = parse_vnd(cong_row.cells[total_col].text)
        if current is not None and result.subtotal_before > 0:
            ratio = subtotal_after / result.subtotal_before
            set_cell_text(cong_row.cells[total_col], format_vnd(current * ratio))
        else:
            set_cell_text(cong_row.cells[total_col], format_vnd(subtotal_after))

    tax_row = _find_summary_row(table, "tổng tiền thuế")
    total_row = _find_summary_row(table, "tổng tiền thanh toán")

    if cong_row is not None and tax_row is not None and total_row is not None:
        tax_before = parse_vnd(tax_row.cells[total_col].text)
        if tax_before is not None and result.subtotal_before > 0:
            ratio = subtotal_after / result.subtotal_before
            new_tax = tax_before * ratio
            set_cell_text(tax_row.cells[total_col], format_vnd(new_tax))
            set_cell_text(total_row.cells[total_col], format_vnd(subtotal_after + new_tax))

    if discount_col is None and result.discount_total > 0 and total_row is not None:
        has_discount_row = any(is_discount_summary_row(row) for row in table.rows)
        if not has_discount_row:
            _insert_discount_row_before(table, total_row, result.discount_total)

    return result


def apply_delivery_time(doc: Document, days: int) -> bool:
    delivery_text = (
        f"Trong vòng {days} ngày tính từ ngày Công ty chúng tôi nhận được tiền đặt cọc."
    )
    found = False
    for table in doc.tables:
        for row in table.rows:
            if not any("thời gian giao hàng" in c.text.lower() for c in row.cells):
                continue
            for cell in row.cells:
                txt = cell.text.strip()
                lower = txt.lower()
                if "thời gian giao hàng" in lower and ":" in lower:
                    continue
                if "trong vòng" in lower or "....." in txt or (
                    "ngày" in lower and "giao hàng" not in lower
                ):
                    set_cell_text(cell, delivery_text)
                    found = True
    return found


def append_discount_note(doc: Document, customer_id: str, discount_percent: float) -> None:
    table = find_items_table(doc)
    customer_text = f"Mã KH: {customer_id}"
    discount_text = f"Chiết khấu áp dụng: {discount_percent:.1f}%"
    if table is None:
        p1 = doc.add_paragraph()
        p1.add_run(customer_text).bold = True
        p2 = doc.add_paragraph()
        p2.add_run(discount_text).bold = True
        return
    new_p = OxmlElement("w:p")
    table._tbl.addnext(new_p)
    paragraph = Paragraph(new_p, doc.element.body)
    paragraph.add_run(customer_text).bold = True

    new_p2 = OxmlElement("w:p")
    paragraph._p.addnext(new_p2)
    paragraph2 = Paragraph(new_p2, doc.element.body)
    paragraph2.add_run(discount_text).bold = True


def process(
    input_bytes: bytes,
    discount_percent: float,
    delivery_days: int = 15,
    customer_id: str | None = None,
    qty: int = 1,
) -> tuple[bytes, DiscountResult]:
    doc = Document(io.BytesIO(input_bytes))
    cid = customer_id or extract_customer_id(doc) or "UNKNOWN"

    codes = extract_product_codes(doc)
    if not codes:
        raise ProcessingError(
            "Không tìm thấy mã hàng. Thêm dòng: Mã hàng: CODE1; CODE2"
        )

    products, missing = resolve_codes(codes)
    if missing:
        raise ProcessingError(
            f"Không tìm thấy mã hàng trong AMIS: {', '.join(missing)}",
            missing_codes=missing,
        )

    expand_and_fill_product_rows(doc, products, qty=qty)
    context = build_mock_context(cid, products, qty=qty)
    replace_placeholders(doc, context)
    apply_product_images(doc, products)

    discount_result = apply_discount_to_items(doc, discount_percent)
    apply_delivery_time(doc, delivery_days)
    append_discount_note(doc, cid, discount_percent)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue(), discount_result
