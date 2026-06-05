"""Streamlit POC: Price Report (.docx) + discount lookup + placeholder images."""

from __future__ import annotations

import sys
from pathlib import Path

import streamlit as st

_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from bill_store import list_bills, load_bill, save_bill  # noqa: E402
from customer_store import get_discount, save_discount  # noqa: E402
from docx_processor import extract_customer_id_from_bytes, process  # noqa: E402

MIME_DOCX = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)


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

    customer_id = extract_customer_id_from_bytes(data)

    if not customer_id:
        st.error("Không thấy mã KH. Thêm dòng: **Mã KH: VHM001**")
        return

    st.success(f"Mã KH: **{customer_id}**")

    saved = get_discount(customer_id)
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
                help="Áp dụng vào cột Thành tiền khi file có số thực",
            )
        )
    with col2:
        delivery_days = int(
            st.number_input(
                "Thời gian giao hàng (ngày)",
                min_value=1,
                max_value=365,
                value=15,
                step=1,
                help="Ghi vào mục Thời gian giao hàng trong báo giá",
            )
        )

    save_for_future = st.checkbox(
        "Lưu CK cho lần sau",
        value=(saved is None or discount != default_discount),
    )

    if st.button("Tạo báo giá", type="primary"):
        if save_for_future:
            save_discount(customer_id, discount)
        try:
            out, discount_result = process(
                data,
                discount,
                delivery_days=delivery_days,
                customer_id=customer_id,
            )
        except Exception as exc:  # noqa: BLE001 — POC surface errors in UI
            st.exception(exc)
            return

        output_filename = f"BaoGia_{customer_id}.docx"
        save_bill(
            customer_id,
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
                f"Đã CK {discount_result.lines_updated} dòng · "
                f"Giảm {_format_vnd_short(discount_result.discount_total)} VND"
            )
        elif discount > 0:
            st.caption(
                "CK % đã ghi chú; không đổi số (file còn placeholder ##…## hoặc không có giá)."
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


def _format_vnd_short(value: float) -> str:
    return f"{round(value):,}".replace(",", ".")


def _render_history_tab() -> None:
    st.caption("Tìm và tải lại báo giá đã tạo theo mã khách hàng.")
    search_id = st.text_input("Mã KH", key="history_customer_id").strip()

    if not search_id:
        st.info("Nhập mã KH để tìm.")
        return

    bills = list_bills(search_id)
    if not bills:
        st.warning(f"Không có báo giá đã lưu cho **{search_id}**.")
        return

    st.success(f"Tìm thấy **{len(bills)}** báo giá.")

    for bill in bills:
        created = bill.created_at.replace("T", " ")[:19]
        label = (
            f"{created} UTC · CK {bill.discount_percent:g}% · "
            f"Giao {bill.delivery_days} ngày · {bill.source_filename}"
        )
        file_bytes = load_bill(bill.id)
        if file_bytes is None:
            st.error(f"Không đọc được file: {bill.output_filename}")
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

    tab_create, tab_history = st.tabs(["Tạo báo giá", "Lịch sử báo giá"])

    with tab_create:
        _render_create_tab()

    with tab_history:
        _render_history_tab()


if __name__ == "__main__":
    main()
