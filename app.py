"""Streamlit POC: Price Report (.docx) + discount lookup + placeholder images."""

from __future__ import annotations

import sys
from pathlib import Path

import streamlit as st

_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from customer_store import get_discount, save_discount  # noqa: E402
from docx_processor import extract_customer_id_from_bytes, process  # noqa: E402

MIME_DOCX = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)


def main() -> None:
    st.set_page_config(page_title="Báo Giá", layout="centered")
    st.title("Báo Giá")
    st.caption("Tải file .docx")

    uploaded = st.file_uploader("File báo giá (.docx)", type=["docx"])

    if uploaded is not None:
        st.session_state["upload_bytes"] = uploaded.read()
        st.session_state["upload_name"] = uploaded.name
        st.session_state.pop("report_bytes", None)
        st.session_state.pop("report_download_name", None)

    data = st.session_state.get("upload_bytes")

    if not data:
        st.info("Chọn file .docx")
        return

    customer_id = extract_customer_id_from_bytes(data)

    if not customer_id:
        st.error(
            "Không thấy mã KH. Thêm dòng: **Mã KH: VHM001**"
        )
        return

    st.success(f"Mã KH: **{customer_id}**")

    saved = get_discount(customer_id)
    default_discount = float(saved) if saved is not None else 0.0
    if saved is not None:
        st.info(f"CK đã lưu: **{saved:g}%**")
    else:
        st.info("Chưa có CK lưu")

    discount = float(
        st.number_input(
            "Chiết khấu (%)",
            min_value=0.0,
            max_value=100.0,
            value=default_discount,
            step=0.5,
            help="Giá trị áp dụng cho file này",
        )
    )
    save_for_future = st.checkbox(
        "Lưu cho lần sau",
        value=(saved is None or discount != default_discount),
    )

    if st.button("Tạo báo giá", type="primary"):
        if save_for_future:
            save_discount(customer_id, discount)
        try:
            out = process(data, discount)
        except Exception as exc:  # noqa: BLE001 — POC surface errors in UI
            st.exception(exc)
            return
        st.session_state["report_bytes"] = out
        st.session_state["report_download_name"] = f"BaoGia_{customer_id}.docx"

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


if __name__ == "__main__":
    main()
