
from openpyxl import load_workbook
from pathlib import Path
from calls.views import normalize_header

folder = Path(r"C:\Users\Admin\Downloads")

files = list(folder.glob("goldcoin_dealer_calls*.xlsx"))
files += list(folder.glob("Yokohama Gold Coin Audit*.xlsx"))
files += list(folder.glob("Yokohama Gold Coin Disbursal*.xlsx"))

for file in files:
    print("\n================================")
    print("FILE:", file.name)

    wb = load_workbook(file, read_only=True, data_only=True)

    for ws in wb.worksheets:
        print("\nSHEET:", ws.title)

        for row_num, row in enumerate(
            ws.iter_rows(min_row=1, max_row=20, values_only=True), 1
        ):
            headers = [normalize_header(v) for v in row]

            if any("dealername" in h for h in headers):
                print("HEADER ROW:", row_num)
                print("HEADERS:", row)

                for col, header in enumerate(headers):
                    if any(x in header for x in [
                        "contact", "mobile", "phone", "telephone", "whatsapp"
                    ]):
                        count = sum(
                            1 for data_row in ws.iter_rows(
                                min_row=row_num + 1,
                                min_col=col + 1,
                                max_col=col + 1,
                                values_only=True
                            )
                            if data_row[0] not in (None, "")
                        )
                        print("CONTACT COLUMN:", row[col])
                        print("NON-EMPTY VALUES:", count)

                break

    wb.close()
