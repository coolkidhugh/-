# AGENTS.md

## Cursor Cloud specific instructions

This repo is a single small **Streamlit** app (Excel report analyzer for 伯爵酒店 hotel team bookings). Core logic is in `analyze_excel.py`; the UI is in `app.py`. Dependencies are pip-based via `requirements.txt` (pandas, streamlit, openpyxl). There is no database, Docker, CI, or automated test suite.

### Running the app (dev)

- Dependencies install to the user site; `streamlit` lives in `~/.local/bin`, which is NOT on `PATH` by default. Export it first:
  `export PATH="$HOME/.local/bin:$PATH"`
- Start the dev server: `streamlit run app.py` (defaults to port 8501, with hot-reload on file save). In headless/remote environments use `streamlit run app.py --server.address 0.0.0.0 --server.port 8501 --server.headless true`.
- Health check: `curl http://localhost:8501/_stcore/health` returns `ok`.

### Testing / lint / build

- There is no lint config, no build step, and no test framework in the repo.
- To exercise the core logic without the UI, call `analyze_reports_ultimate([<xlsx paths>])` from `analyze_excel.py` (running `analyze_excel.py` directly only prints usage guidance).

### Input format gotchas (for creating test `.xlsx` files)

The parser expects specific row/column structure per report:
- A group line containing `团体名称: <name> 市场码：<code>` (market code drives categorization: prefixes `MGM`/`MTC` = meeting/corporate teams, `GTO` = travel agency).
- A header row containing the strings `房号`, `姓名`, and `人数`; that header must also include columns named `状态`, `房数`, `房类`.
- Data rows follow the header; rows containing `小计` are skipped.
- Status filtering depends on the file NAME keywords: `在住` → statuses `R,I`; `离店`/`后天` → `I,R,O`; otherwise → `R`. So a filename like `次日到达_测试.xlsx` only counts rows with status `R`.
- Room-type codes not in the hardcoded `jinling_room_types`/`yatai_room_types` lists are reported as "未知房型代码" (unknown codes).
