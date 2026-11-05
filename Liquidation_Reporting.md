# Liquidation Reporting

Builds the **FSN × DS × City** liquidation risk report from the IPC NPD-risk raw file, enriched with NLC, live prices and last-7-day sales. Spec: `spec_files_Liq_reporting.xlsx` (tabs `Reference_Files`, `Ref_Output_Format_liq`, `Vert_List_for_query`).

## Run

```bash
python main.py          # prompts for sales period, then confirm
python main.py --yes    # no prompts; sales period = last 7 days in the sales file
```

Requires Python 3 with `pandas` and `openpyxl`. Runtime is about 5–6 min (NLC loading is the slowest part). Needs access to the Google Drive `G:` shortcut.

## Folders

Root: `G:\.shortcut-targets-by-id\15JLkN3qz_KXLVTnsFWiCIXCbhkgPWc4q\SMNM_Minutes` (constants at the top of `main.py`).

| Input | Folder | File pattern |
|---|---|---|
| IPC NPD risk (main source) | `Liquidation_Raw_Data` | `ipc_DataSense_Dry_NPD_ds_npd_risk_<YYYYMMDD>.csv` |
| DS → City | `Mapping_Tables` | `*_DS_City_Map.csv` |
| Vertical → Sell BU | `Mapping_Tables` | `*_Minutes_BU_Tags.csv` |
| Day-level sales | `Sales_Data\day_level_sales_fsn_ds` | `Sales_FSN_DS_<range>.csv` |
| Live prices | `Live_Prices` | `*_Live_Prices_Minutes.csv` |
| NLC (weekly Bizfin dump) | `NLC_Dump\nlc_fsn_city_weekly` | `Category_WEEK <n> NLC_*.xlsx` |

The Liquidation (IPC NPD risk) file is chosen at run time from a numbered list (Enter = newest; `--yes` = newest). The latest file of every other type is picked automatically (date parsed from the file name). Output goes to `Output_files_Liquidation\<IPC date>_liquidation_output_sales_<from>-<to>_<HHMMSS>.csv`.

## Logic

1. **Base:** IPC file, rows with `atp > 0`, aggregated to FSN × DS.
2. **Risk ATP:** Total = sum of `atp`. It is split into buckets by the IPC `risk_bucket` label: 0-3, 4-7, 8-15, 16-31 days. The buckets sum to the total.
3. **NLC:** per FSN. Only the weekly files are used, latest week/version first (`nlc_sales_bizfin` is not used). The first file containing the FSN wins. Multiple rows per FSN are averaged. Tax-exclusive NLC is preferred.
4. **Risk Value:** each bucket's ATP × NLC. Total Risk Value is Total Risk ATP × NLC.
5. **Live LP / MRP:** latest live-price file (read in chunks, several GB). Live LP = mean `final_price` per FSN × DS (`business_zone` = DS). MRP = mean `mrp` per FSN × DS, rounded to 0 decimals.
6. **Daily units:** period chosen at run time. Prompt shows the dates available in the latest sales file: [1] last 7 days (default) or [2] custom range (YYYY-MM-DD, within the file range). Units are summed per FSN × DS × day, missing days are 0. **Sales Sum (<n> days)** = total units in the period. **Avg DRR (<n> days)** = Sales Sum ÷ number of days selected. **% Achievement** = Sales Sum ÷ Total Risk ATP, stored as a numeric ratio (0.85 = 85%, 4 decimals); apply % format in Excel. Blank when Total Risk ATP is 0.
7. **Sell BU:** `analytic_vertical` mapped to `Minutes_BU_Tags`. **City:** DS mapped via `DS_City_Map` (`Warehouse ID`).
8. **Export:** sorted by Total Risk ATP, descending.

## Output columns

`fsn, ds, City, analytic_vertical, Sell BU, brand, product_title, NLC, Total Risk ATP, Risk ATP (0-3 / 4-7 / 8-15 / 16-31 days), Total Risk Value, Risk Value (0-3 / 4-7 / 8-15 / 16-31 days), Live LP, MRP, <daily unit columns for the selected period>, Sales Sum (<n> days), Avg DRR (<n> days), % Achievement`

## Notes and gotchas

- To change the risk basis, edit `RISK_COL` (currently `atp`) and `BUCKETS` at the top of `main.py`.
- The spec names the output `<date>_smnm_output`. That looks like a copy-paste from SMNM, so the file is named `*_liquidation_output_*`.
- The `Next 31 Days Expiry Tracker` CSV in `Liquidation_Raw_Data` is **not** used by this pipeline.
- The sales default window ends on the newest date in the sales file, not the IPC date. If the sales file lags the IPC file, the default shifts with it.
- IPC buckets outside 0-31 days (if any) count only in Total, and the run logs them.
- No automated tests. Sanity checks to do after a run: no duplicate fsn × ds keys, and the buckets sum to the total.
- Each run adds a new timestamped file, so clean up old outputs in `Output_files_Liquidation` as needed.







