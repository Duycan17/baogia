# Báo Giá — AMIS Catalog Merge

Streamlit app that fills a Word báo giá template from AMIS product master data.

## Flow

1. Upload `.docx` with **Mã KH** and **Mã hàng** lines.
2. Match product codes against [`file nhap quat AMIS.xls`](file nhap quat AMIS.xls).
3. Fill line items (tên, mô tả, ĐVT, SL, đơn giá, thành tiền, ảnh).
4. Mock customer/contact/date placeholders — no `##…##` left in output.
5. Apply chiết khấu %, giao hàng, lưu lịch sử.

## Template format

Add plain text lines anywhere in the document:

```text
Mã KH: DEMO001
Mã hàng: CV.P7/560/80/K1-0.75kW4P, SUS316; AVR.H4/6/350/K13-0.75KW2P
```

- Separate multiple codes with `;` (not `,` — codes may contain commas).
- Fallback: codes appearing anywhere in the document text are matched against AMIS.

## AMIS columns used

| Header | Usage |
|--------|-------|
| Mã hàng | Lookup key |
| Tên hàng | Line item name |
| Mô tả | Specs column |
| Đơn vị tính chính | Unit |
| Đơn giá bán | Unit price |
| Thuế suất GTGT (%) | VAT (weighted average) |
| Hình ảnh sản phẩm | Image URL or local path (column AZ when present) |

## Setup

```bash
uv sync
```

## Run

```bash
uv run streamlit run app.py
```

## Project layout

| Path | Role |
|------|------|
| `app.py` | Streamlit UI |
| `src/product_catalog.py` | Load AMIS `.xls` by header names |
| `src/code_extractor.py` | Parse `Mã hàng` from docx |
| `src/mock_context.py` | Mock quote/customer fields + totals |
| `src/placeholder_filler.py` | Replace `##…##` placeholders |
| `src/product_rows.py` | Clone/fill item rows + images |
| `src/docx_processor.py` | Pipeline orchestration |
| `src/docx_utils.py` | Shared docx helpers |
| `src/image_provider.py` | Product image load/cache |
| `assets/default_product.jpg` | Fallback image |
| `data/product_images/` | Downloaded image cache (gitignored) |

## Config

| Variable | Default |
|----------|---------|
| `CATALOG_PATH` | `file nhap quat AMIS.xls` |

## Deploy

See `render.yaml` and `requirements.txt` for cloud deploy. Local `data/` may reset on redeploy.
