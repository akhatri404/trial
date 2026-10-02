# PB strategy report generator

Upload the pet-category Excel workbook (the one with the 「サブカテ」 sheet) and download the report automatically as a **PDF** or an editable **PowerPoint (.pptx)**.

## What it does
1. Reads the workbook (`pbreport/loader.py`), validates the layout and gives a plain-language error if something is missing.
2. Runs the analysis (`pbreport/analysis.py`):
   * **The workbook's own numbers, no derived PB sales:** the workbook has no PB sales column, so PB sales (sales x PB ratio) are never calculated or shown. Sub-categories use the 「サブカテ」 sheet, categories the 「カテゴリー」 sheet, departments and the company total the 「方向性」 sheet, exactly as given (sales, units, PB ratio, PB SKU count). Only for a level the workbook has no figure for (e.g. a file without those sheets) is the PB ratio the sales-weighted average of the rows below, and the data checks say so. Categories that exist only in the category sheet (観賞魚 / 鳥 / 昆虫) are included, and every run reconciles the sub-category rows with the workbook's category, department and total figures.
   * **PB ratio and SKUs:** PB ratio change (pt) by category and sub-category, and PB SKU change vs PB ratio change with a verdict (a minimum PB ratio / sales base so tiny bases don't produce meaningless verdicts).
   * **Department view (部門 → カテゴリー → サブカテ):** a comparison of the five departments (sales, PB ratio, SKUs), then one page per department with its category table (sales rank inside the department, PB ratio change, SKU vs PB ratio verdict) and every sub-category with its direction.
   * **Direction rules** (拡大 Scale / 横展開 Replicate / シェア奪回 Recapture / 小規模テスト Test small / 要点検 Review / 防衛 Defend / 優先度低 Deprioritize / 経過観察 Monitor), based on sales size and PB ratio level and change. Data checks and detailed caveats are shown on the app page, not in the report; the report carries only a short 前提・注意事項 box.
3. Builds the report in the chosen format:
   * **PDF** (`pbreport/pdf.py`): about 7 pages with an embedded Japanese font; one page per department lists every sub-category.
   * **PowerPoint** (`pbreport/pptx_report.py`): 14-slide 16:9 deck with real slide titles, *native* charts and tables (editable in PowerPoint), speaker notes on every slide, Calibri / Yu Gothic fonts.

The PDF (日本語), PowerPoint, app screen and error messages are in Japanese; an English PDF (`-f pdf_en`) is also available (category names and other Excel labels stay Japanese); amounts use 億円 / 万円. The analysis is deterministic: the same file always gives the same report. トライアル direction markers and the 過去資料 sheet are ignored.

## Run it
```bash
./run.sh                                  # installs requirements, then starts the app on port 8501
# or by hand:
pip install -r requirements.txt
streamlit run app.py                      # upload page -> choose PDF (日本語) / PDF (English) / PowerPoint -> download
python -m pbreport input.xlsx -f pptx -o out    # command line; -f pdf | pdf_en | pptx | both
python -m pbreport --watch inbox/ reports/ -f both   # drop an .xlsx in inbox/, reports appear in reports/
```

## Deploy (Streamlit Community Cloud)
1. Push this folder to a GitHub repository.
2. At share.streamlit.io choose **Create app**, select the repository and branch, and set the main file to `app.py`.
   The platform installs `requirements.txt` and runs `streamlit run app.py` itself; `run.sh` is not used there.
3. Files are processed in memory; nothing is stored on the server. The data is commercially sensitive, so keep the app
   **private** and invite only the people who need it (app **Settings → Sharing**).

`run.sh` is for running the app on your own machine or a Linux server (`PORT=8080 ./run.sh` to change the port).

## Expected layout
Sheet name contains 「サブカテ」; a header row with 部門 / カテゴリ / サブカテ and, for each period, columns named like
`25/6期 全体売上`, `全体売れ数`, `PB比率(売上)`, `PB比率(売れ数)`, `PB SKU数`. Two or more periods. If the newest period covers clearly less than the others (a 「※Nヶ月分」 note on the 「方向性」 sheet, or sales under 60% of the previous period), growth, the bridge and the directions use the last two comparable periods. The newest period is shown next to them as it is: PB ratio and PB SKU count are compared, sales and PB sales are shown but their change is not. Nothing is scaled up and no number of months is assumed or printed.

## Changing the rules
All thresholds are fixed defaults in `Thresholds` (`pbreport/analysis.py`); they are not shown in the report. Fonts: bundled IPAex Gothic (see `fonts/`, IPA Font License);
set `PBREPORT_FONT` to use another TrueType font.

## Tests
`PB_SAMPLE=/path/to/workbook.xlsx pytest -q`  (also checks the PPTX has native charts, tables and slide titles)
