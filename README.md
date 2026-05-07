# Price Report Discount POC

Small Streamlit app + Python helpers that:

1. Accept a **Báo Giá** Word report (`.docx`).
2. Read **customer id** from a line like `Mã KH: VHM001` anywhere in the file.
3. Look up a **discount percent** in `data/customers.json`; first-time customers enter a value and it is saved for later.
4. Insert a **fixed real furniture image** (`assets/furniture.jpg`) into each line-item row (before totals).
5. Append **Mã KH: ...** and **Chiết khấu áp dụng: X%** immediately after the items table.
6. Offer the modified `.docx` for download.

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

## Run the app

```bash
uv run streamlit run app.py
```

Then upload the `.docx`, set discount if prompted, click **Generate report**, and download the result.

## Files

| Path | Role |
|------|------|
| `app.py` | Streamlit UI |
| `src/customer_store.py` | JSON read/write for `customer_id` → `discount_percent` |
| `src/image_provider.py` | Loads fixed furniture image (fallback to generated placeholder) |
| `src/docx_processor.py` | Parse id, inject images, append discount note |
| `assets/furniture.jpg` | Fixed product image downloaded via Firecrawl search |
| `data/customers.json` | Created automatically; gitignored |

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

Note: this app stores discounts in `data/customers.json`. On free cloud instances, local disk may be ephemeral; data can reset on redeploy/restart.

### Option 2: Streamlit Community Cloud

1. Push repository to GitHub.
2. In Streamlit Community Cloud, create an app from the repo.
3. Set main file path to `app.py`.
4. It will install dependencies from `requirements.txt` automatically.

# baogia
