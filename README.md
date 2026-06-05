# Price Report Discount POC

Small Streamlit app + Python helpers that:

1. Accept a **Báo Giá** Word report (`.docx`).
2. Read **customer id** from a line like `Mã KH: VHM001` anywhere in the file.
3. Look up a **discount percent** in `data/customers.json`; first-time customers enter a value and it is saved for later.
4. Insert a **fixed real furniture image** (`assets/furniture.jpg`) into each line-item row (before totals).
5. Apply **Chiết khấu (%)** to line **Thành tiền** when the file has real numbers (post-AMIS merge); recalc subtotals and add a **Chiết khấu** summary row when needed.
6. Fill **Thời gian giao hàng** (days) in the terms table.
7. Append **Mã KH: ...** and **Chiết khấu áp dụng: X%** immediately after the items table.
8. **Save bill history** per customer and re-download from the **Lịch sử báo giá** tab.

## Prerequisites

- [uv](https://docs.astral.sh/uv/)
- Python **3.12+** (see `.python-version`)

## Setup

```bash
uv sync
```

## Template change (one-time)

Open your price report template in Word and add a plain text line (not a merge field), for example under **Kính gửi**:

```text
Mã KH: VHM001
```

ASCII `Ma KH: VHM001` is also accepted.

If your template has a **Chiết khấu** column, the app fills per-line discount amounts there automatically.

## Run the app

```bash
uv run streamlit run app.py
```

### Tạo báo giá

Upload the `.docx`, set **Chiết khấu (%)** and **Thời gian giao hàng (ngày)**, click **Tạo báo giá**, and download the result.

Numeric discount only applies when **SL**, **Đơn giá**, and **Thành tiền** contain real numbers. Raw `##Báo giá…##` merge placeholders are left unchanged (discount % is still noted in the footer).

### Lịch sử báo giá

Open the **Lịch sử báo giá** tab, enter a **Mã KH**, and download any previously generated file for that customer.

## Files

| Path | Role |
|------|------|
| `app.py` | Streamlit UI (create + history tabs) |
| `src/customer_store.py` | JSON read/write for `customer_id` → `discount_percent` |
| `src/bill_store.py` | Save/list/load generated `.docx` by customer |
| `src/image_provider.py` | Loads fixed furniture image (fallback to generated placeholder) |
| `src/docx_processor.py` | Parse id, images, discount math, delivery time, notes |
| `assets/furniture.jpg` | Fixed product image |
| `data/customers.json` | Discount defaults; gitignored |
| `data/bills/` | Saved báo giá files; gitignored |
| `data/bills_index.json` | Bill metadata index; gitignored |

## Docs during development

Library APIs were checked with Context7 (`npx ctx7@latest`) for **python-docx** and **Streamlit** widgets (`file_uploader`, `download_button`, `number_input`).

## Deploy globally

### Option 1: Render (recommended for this repo)

This repo now includes:

- `requirements.txt` (for dependency install)
- `.streamlit/config.toml` (headless server config)
- `render.yaml` (infrastructure-as-code service config)

Steps:

1. Push this repository to GitHub.
2. In Render, create a new **Blueprint** and select this repository.
3. Render will read `render.yaml` and create a web service using:
   - Build command: `pip install -r requirements.txt`
   - Start command: `streamlit run app.py --server.address 0.0.0.0 --server.port $PORT`
4. After deploy finishes, you get a public URL that is reachable globally.

Note: this app stores discounts in `data/customers.json` and bills in `data/bills/`. On free cloud instances, local disk may be ephemeral; data can reset on redeploy/restart.

### Option 2: Streamlit Community Cloud

1. Push repository to GitHub.
2. In Streamlit Community Cloud, create an app from the repo.
3. Set main file path to `app.py`.
4. It will install dependencies from `requirements.txt` automatically.

# baogia
