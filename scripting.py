from pathlib import Path
import xlwings as xw
xlsx_path = Path(__file__).resolve().parent / "Data case gas_cross_border - 2026.xlsx"

if not xlsx_path.is_file():
	raise FileNotFoundError(f"Workbook not found at: {xlsx_path}")

book = xw.Book(str(xlsx_path))

print(book.sheets)

