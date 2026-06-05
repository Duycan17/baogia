"""Streamlit POC: Báo giá .docx + AMIS catalog merge."""

from __future__ import annotations

import sys
from pathlib import Path

import streamlit as st

_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from bill_store import list_bills, load_bill, save_bill  # noqa: E402
from customer_store import get_discount, save_discount  # noqa: E402
from docx_processor import (  # noqa: E402
    ProcessingError,
    preview_from_bytes,
    process,
)
from docx_utils import format_vnd  # noqa: E402

MIME_DOCX = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)


def _format_vnd_short(value: float) -> str:
    return format_vnd(value)


def _render_product_preview(data: bytes) -> bool:
    preview = preview_from_bytes(data)

    if not preview.product_codes:
        st.error("Không tìm thấy mã hàng. Thêm: **Mã hàng: CODE1; CODE2**")
        return False

    if preview.missing_codes:
        st.error(
            "Mã không có trong AMIS: **"
            + "**, **".join(preview.missing_codes)
            + "**"
        )
        return False

    rows = [
        {
            "Mã": p.code,
            "Tên": p.name[:40],
            "Giá": _format_vnd_short(p.sale_price),
            "ĐVT": p.unit,
        }
        for p in preview.products
    ]
    st.dataframe(rows, hide_index=True, use_container_width=True)
    st.caption(f"Đã khớp {len(preview.products)} sản phẩm từ AMIS")
    return True


def _render_create_tab() -> None:
    uploaded = st.file_uploader("File báo giá (.docx)", type=["docx"])

    if uploaded is not None:
        st.session_state["upload_bytes"] = uploaded.read()
        st.session_state["upload_name"] = uploaded.name
        st.session_state.pop("report_bytes", None)
        st.session_state.pop("report_download_name", None)

    data = st.session_state.get("upload_bytes")
    upload_name = st.session_state.get("upload_name", "report.docx")

    if not data:
        st.info("Chọn file .docx")
        return

    preview = preview_from_bytes(data)

    if not preview.customer_id:
        st.error("Không thấy mã KH. Thêm dòng: **Mã KH: VHM001**")
        return

    st.success(f"Mã KH: **{preview.customer_id}**")

    if not _render_product_preview(data):
        return

    saved = get_discount(preview.customer_id)
    default_discount = float(saved) if saved is not None else 0.0
    if saved is not None:
        st.info(f"CK đã lưu: **{saved:g}%**")
    else:
        st.info("Chưa có CK lưu")

    col1, col2 = st.columns(2)
    with col1:
        discount = float(
            st.number_input(
                "Chiết khấu (%)",
                min_value=0.0,
                max_value=100.0,
                value=default_discount,
                step=0.5,
            )
        )
    with col2:
        delivery_days = int(
            st.number_input(
                "Giao hàng (ngày)",
                min_value=1,
                max_value=365,
                value=15,
                step=1,
            )
        )

    save_for_future = st.checkbox(
        "Lưu CK cho lần sau",
        value=(saved is None or discount != default_discount),
    )

    if st.button("Tạo báo giá", type="primary"):
        if save_for_future:
            save_discount(preview.customer_id, discount)
        try:
            out, discount_result = process(
                data,
                discount,
                delivery_days=delivery_days,
                customer_id=preview.customer_id,
            )
        except ProcessingError as exc:
            st.error(str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            st.exception(exc)
            return

        output_filename = f"BaoGia_{preview.customer_id}.docx"
        save_bill(
            preview.customer_id,
            out,
            source_filename=upload_name,
            output_filename=output_filename,
            discount_percent=discount,
            delivery_days=delivery_days,
        )

        st.session_state["report_bytes"] = out
        st.session_state["report_download_name"] = output_filename

        if discount_result.lines_updated:
            st.caption(
                f"CK {discount_result.lines_updated} dòng · "
                f"Giảm {_format_vnd_short(discount_result.discount_total)} VND"
            )

    out_bytes = st.session_state.get("report_bytes")
    dl_name = st.session_state.get("report_download_name")
    if out_bytes and dl_name:
        st.download_button(
            label="Tải file",
            data=out_bytes,
            file_name=dl_name,
            mime=MIME_DOCX,
            type="primary",
        )


def _render_history_tab() -> None:
    st.caption("Tìm báo giá đã tạo theo mã KH.")
    search_id = st.text_input("Mã KH", key="history_customer_id").strip()

    if not search_id:
        st.info("Nhập mã KH")
        return

    bills = list_bills(search_id)
    if not bills:
        st.warning(f"Không có báo giá cho **{search_id}**")
        return

    st.success(f"Tìm thấy **{len(bills)}** báo giá")

    for bill in bills:
        created = bill.created_at.replace("T", " ")[:19]
        label = (
            f"{created} UTC · CK {bill.discount_percent:g}% · "
            f"Giao {bill.delivery_days} ngày · {bill.source_filename}"
        )
        file_bytes = load_bill(bill.id)
        if file_bytes is None:
            st.error(f"Không đọc được: {bill.output_filename}")
            continue
        st.download_button(
            label=f"Tải — {label}",
            data=file_bytes,
            file_name=bill.output_filename,
            mime=MIME_DOCX,
            key=f"dl_{bill.id}",
        )


def main() -> None:
    st.set_page_config(page_title="Báo Giá", layout="centered")
    st.title("Báo Giá")

    tab_create, tab_history = st.tabs(["Tạo báo giá", "Lịch sử"])

    with tab_create:
        _render_create_tab()

    with tab_history:
        _render_history_tab()


if __name__ == "__main__":
    main()
