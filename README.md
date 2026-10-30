# Stock Report App (Laporan Stock Ecer)

A Windows desktop app that builds a retail **stock request report**, prints it on paper, saves it as a
protected Excel file and keeps a searchable history of every printout.

It was built for a small retail shop that used to fill this report by hand. Staff pick items from the shop's
item master, type the quantities, and print. Every printout is stored, so the office can later search what was
requested, when, and by whom.

> This is a cleaned public copy for my portfolio. All shop data, names and passwords were removed.
> The included item master (`master_barang_contoh.xlsx`) is **made-up sample data**.
> The app's interface is in **Bahasa Indonesia**, because that is what its users speak.

## Features

- **Report table**: pick items from the item master (xlsx) with a search-as-you-type picker; type the
  quantity, warehouse (drop-down) and stock-count columns directly in the table, Excel-style
  (double-click a cell, Enter / Tab to move).
- **Print**: one click saves the report as xlsx and opens it in Excel, or sends it straight to the default
  printer ("otomatis" tick). A4 landscape, at most 20 item rows per paper, the header rows and
  "Halaman x dari y" repeat on every page.
- **Excel export with protection**: optional edit lock (sheet + workbook structure protection), so anyone can
  open and print the file but editing needs the password. Only Excel's own password hash is stored.
- **Watermark**: a light, repeated shop-name pattern drawn with Pillow and injected into the xlsx as a page
  header picture (printed, edge to edge) and as a sheet background (on screen). openpyxl cannot do this, so
  the app edits the saved xlsx package (VML drawing + relationships) itself. The text is configurable.
- **Approval ("Disetujui")**: approvers tick a box and type their own password; their name is printed on the
  report and stored in the history.
- **History ("Riwayat")**: every printout goes into SQLite. Search, date range, sort, a "total per item" view,
  open any printout, edit its quantities (approved reports need a password), reprint, and import old report
  files. The history and the xlsx files stay in sync (changed files are re-read, edits rewrite the file).
- **Password-protected settings**: the gear window (approvers, Excel lock, settings password) needs the
  settings password. All app passwords are stored only as **PBKDF2-SHA256 hashes with a random salt**
  (100,000 rounds, constant-time compare); the settings can set passwords but never show them. Older plain-text
  configs are upgraded automatically on start.
- **Self-upgrading database**: new columns are added with `ALTER TABLE` on start, so old history files keep
  working.

## Tech

- Python 3, **tkinter / ttk** (GUI), **openpyxl** (xlsx), **Pillow** (watermark picture), **sqlite3** (history)
- `hashlib.pbkdf2_hmac` + `hmac.compare_digest` for password hashing
- Packaged as a single Windows exe with **PyInstaller**
- One source file: `laporan_stock_ecer.py`

## Run from source

```
pip install -r requirements.txt
python laporan_stock_ecer.py
```

On first start the app loads `master_barang_contoh.xlsx` from its own folder. To use another item master,
click the folder icon (top right). The item master is an xlsx with sheet `ACCDBRG` (else the first sheet),
header in row 1, and columns B = KODEBARANG, C = NAMABARANG, D = SATUAN, F = STNPERPAK.

Files the app creates next to itself (ignored by git): `laporan_config.json` (settings), `laporan.db`
(history) and the `Laporan\` folder (saved reports).

### Default password: change it

The settings password starts as **`admin123`** (demo default). Open the gear window with it and set your own
under "Password Pengaturan". After that only its hash is stored.

### Watermark text

The watermark says `NamaToko` by default. To show your shop name, add this to `laporan_config.json`:

```json
"watermark_text": "Your Shop Name"
```

## Build the exe

```
python -m PyInstaller --noconfirm --onefile --windowed --name LaporanStockEcer ^
   --distpath dist --workpath build --specpath build laporan_stock_ecer.py
```

Copy `master_barang_contoh.xlsx` (or your own item master) next to `dist\LaporanStockEcer.exe`.
Keep `--workpath` on the same drive as the project, and delete the `build\` folder afterwards.

## Screenshots

_Screenshots coming soon._

<!--
![Main window](docs/main.png)
![History (Riwayat)](docs/riwayat.png)
![Printed report with watermark](docs/print.png)
-->

## License

MIT, see [LICENSE](LICENSE).
