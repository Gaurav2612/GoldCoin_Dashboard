
from pathlib import Path
from openpyxl import load_workbook
from calls.models import DealerCall

def clean(value):
    if value is None:
        return ""
    return "".join(str(value).lower().split())

db_codes = set()
db_names = set()

for dealer in DealerCall.objects.all():
    if dealer.dealer_code:
        db_codes.add(clean(dealer.dealer_code))
    if dealer.dealer_name:
        db_names.add(clean(dealer.dealer_name))

print("DATABASE TOTAL:", DealerCall.objects.count())
print("DATABASE WITH PHONE:", DealerCall.objects.exclude(phone="").count())

folder = Path(r"C:\Users\Admin\Downloads")

files = [
    ("goldcoin_dealer_calls (2).xlsx", "Dealer Calls", 1),
    ("Yokohama Gold Coin Disbursal Audit for Q2'26 V2.xlsx",
     "Calls_Verification", 4),
]

for filename, sheet_name, header_row in files:
    path = folder / filename

    print("\n==============================")
    print("FILE:", filename)

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name]

    headers = [
        clean(value)
        for value in next(
            ws.iter_rows(
                min_row=header_row,
                max_row=header_row,
                values_only=True
            )
        )
    ]

    code_col = headers.index("dealercode")
    name_col = headers.index("dealername")

    phone_col = (
        headers.index("phone")
        if "phone" in headers
        else headers.index("contactnumber")
    )

    total = 0
    with_phone = 0
    code_matches = 0
    name_matches = 0

    for row in ws.iter_rows(
        min_row=header_row + 1,
        values_only=True
    ):
        if not row or not row[name_col]:
            continue

        total += 1

        phone = row[phone_col] if phone_col < len(row) else None
        code = clean(row[code_col])
        name = clean(row[name_col])

        if phone not in (None, ""):
            with_phone += 1

            if code and code in db_codes:
                code_matches += 1
            elif name and name in db_names:
                name_matches += 1

    print("TOTAL DEALERS:", total)
    print("DEALERS WITH PHONE:", with_phone)
    print("MATCHED BY DEALER CODE:", code_matches)
    print("MATCHED BY DEALER NAME:", name_matches)

    wb.close()
