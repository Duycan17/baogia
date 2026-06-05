"""Build placeholder values for non-product merge fields."""

from __future__ import annotations

from datetime import date

from num2words import num2words

from docx_utils import format_vnd
from product_catalog import Product


def _amount_in_words_vi(amount: float) -> str:
    n = int(round(amount))
    try:
        words = num2words(n, lang="vi")
    except NotImplementedError:
        words = num2words(n, lang="en")
    return f"{words.capitalize()} đồng"


def compute_totals(products: list[Product], qty: int = 1) -> tuple[float, float, float]:
    subtotal = sum(p.sale_price * qty for p in products)
    if not products:
        return 0.0, 0.0, 0.0
    weighted_vat = sum(p.sale_price * qty * p.vat_percent for p in products) / subtotal
    tax = subtotal * (weighted_vat / 100.0)
    grand = subtotal + tax
    return subtotal, tax, grand


def build_mock_context(
    customer_id: str,
    products: list[Product],
    qty: int = 1,
) -> dict[str, str]:
    today = date.today()
    subtotal, tax, grand = compute_totals(products, qty)
    grand_fmt = format_vnd(grand)
    amount_words = _amount_in_words_vi(grand)

    cid = customer_id.strip() or "UNKNOWN"
    quote_no = f"BG-{cid}-{today.strftime('%Y%m%d')}"
    quote_date = today.strftime("%d/%m/%Y")

    return {
        "##Báo giá.Khách hàng##": f"Công ty {cid}",
        "##Báo giá.Địa chỉ##": "123 Nguyễn Văn Linh, Q. Hải Châu, TP. Đà Nẵng",
        "##Báo giá.Số báo giá##": quote_no,
        "##Báo giá.Ngày báo giá##": quote_date,
        "##Người thực hiện.Họ và tên##": "Nguyễn Văn A",
        "##Người thực hiện.Điện thoại di động##": "0901 234 567",
        "##Người thực hiện.Email công ty##": "sales@vihem1.com.vn",
        "##Liên hệ.Họ và tên##": "Trần Thị B",
        "##Liên hệ.Chức danh##": "Trưởng phòng Mua hàng",
        "##Liên hệ.Phòng ban##": "Phòng Kỹ thuật",
        "##Liên hệ.ĐT di động##": "0912 345 678",
        "##Liên hệ.Email cá nhân##": f"contact.{cid.lower()}@example.com",
        "##Báo giá.Thành tiền##": format_vnd(subtotal),
        "##Báo giá.Tiền thuế##": format_vnd(tax),
        "##Báo giá.Tổng tiền##": grand_fmt,
        "##Hàm chuyển số thành chữ(##Báo giá.Tổng tiền##)##": amount_words,
        "##Hàm chuyển số thành chữ(##": "",
        "##)##": "",
    }
