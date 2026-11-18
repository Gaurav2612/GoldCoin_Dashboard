import re
import logging
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from django.conf import settings
from django.core.cache import cache
import hashlib
from django.core.paginator import Paginator
from django.contrib import messages
from django.shortcuts import redirect
from django.contrib.auth import login, logout
from django.contrib.auth.decorators import login_required
from django.contrib.auth.forms import UserCreationForm
from django.db import transaction
from django.db.models import Count, Q
from django.http import HttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.utils.dateparse import parse_date, parse_time
from django.views.decorators.http import require_POST

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .models import (
    DealerCall,
    DealerFollowUp,
    DealerCallAttempt,
    PODTracking,
    ExcelImportLog,
)

logger = logging.getLogger(__name__)


# ============================================================
# CONSTANTS
# ============================================================

EXCEL_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

HEADER_COLOR = "1F4E78"

BASE_DIR = Path(__file__).resolve().parent.parent

EXCEL_IMPORT_DIR = settings.BASE_DIR / "data" / "excel_import"


def get_excel_files():

    EXCEL_IMPORT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    excel_files = [
        file_path
        for file_path in EXCEL_IMPORT_DIR.iterdir()
        if (
            file_path.is_file()
            and file_path.suffix.lower() in {".xlsx", ".xlsm"}
            and not file_path.name.startswith("~$")
        )
    ]

    return sorted(
        excel_files,
        key=lambda file_path: file_path.name.lower()
    )
ALIASES = {
    "sr_no": ["sr. no.", "sr no", "serial no"],
    "scheme_period": ["scheme period"],
    "city": ["cities", "city"],
    "location_code": ["location code"],
    "location_name": ["location name"],
    "document_no": ["document no.", "document no", "document number"],
    "posting_date": ["posting date"],
    "dealer_code": ["dealer code", "code", "dealer id"],
    "dealer_name": ["dealer name", "name", "dealer"],
    "sales_in_charge": ["sales in charge"],
    "qty": ["qty", "quantity"],
    "territory": [
        "territory",
        "territory name",
        "sales territory",
        "territory code",
    ],
    "address": ["address"],
    "phone": [
        "phone", "contact number", "contact no", "contact",
        "mobile", "mobile number", "mobile no", "phone number",
        "telephone", "whatsapp number",
    ],
    "email": ["email", "email address"],
    "contact_name": ["contact name"],
    "recipient_relationship": [
        "recipient relationship (with dealer)",
        "recipient relationship",
        "relationship with dealer",
    ],
    "pod_received": [
        "pod received from yokohama",
        "received from yokohama",
    ],
    "audited_by_cofiser": ["audited by cofiser"],
    "call_status": ["calling status", "call status", "status"],
    "call_1_remarks": ["call 1 (remarks)", "call 1 remarks"],
    "coin_receive_remark": ["coin receive remark"],
    "call_2_remarks": ["call 2 (remarks)", "call 2 remarks"],
    "final_call_remarks": ["final call (remarks)", "final call remarks"],
    "remarks": ["remarks", "remark", "comments"],
    "pod_number": ["pod number", "pod"],
}


# ============================================================
# EXCEL AND REMARK HELPERS
# ============================================================

def normalize_header(value):
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def clean_value(value, phone=False, number_format=None):
    if value is None or isinstance(value, bool):
        return ""

    if phone:
        if isinstance(value, (int, float, Decimal)):
            try:
                number = Decimal(str(value))

                if not number.is_finite():
                    return ""

                result = format(number, "f")

                if "." in result:
                    result = result.rstrip("0").rstrip(".")

                if number_format and re.fullmatch(r"0+", number_format):
                    result = result.zfill(len(number_format))

                return result

            except (InvalidOperation, ValueError):
                return ""

        result = str(value).strip()

        if result.lower() in {"none", "nan", "null"}:
            return ""

        if re.fullmatch(r"\d+\.0+", result):
            result = result.split(".")[0]

        return result

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


def normalize_remark(value):
    return re.sub(
        r"\s+",
        " ",
        str(value or "").replace("\xa0", " "),
    ).strip().lower()



def is_gold_coin_received(*remarks):
    positive_pattern = re.compile(
        r"\b(?:received|receive|receiving|recevied|receieved)\b",
        re.IGNORECASE,
    )

    negative_pattern = re.compile(
        r"\b(?:not|never)\s+(?:yet\s+)?(?:been\s+)?"
        r"(?:received|receive|receiving|recevied|receieved)\b"
        r"|\bnot\s+received\s+yet\b"
        r"|\bnot\s+receive\s+yet\b",
        re.IGNORECASE,
    )

    for remark in remarks:
        text = normalize_remark(remark)

        if not text:
            continue

        if negative_pattern.search(text):
            continue

        if positive_pattern.search(text):
            return True

    return False


def contains_received(*remarks):
    return is_gold_coin_received(*remarks)


def is_successful_connection(remark):
    text = normalize_remark(remark)

    if not text:
        return False

    negative_patterns = [
        r"\bnot\s+connected\b",
        r"\bnot\s+connect(?:ed)?\b",
        r"\bcould\s+not\s+connect\b",
        r"\bunable\s+to\s+connect\b",
        r"\bno\s+answer\b",
        r"\bnot\s+answer(?:ed)?\b",
        r"\bnot\s+receiv(?:e|ed)\b",
        r"\bbusy\b",
        r"\bswitched\s+off\b",
        r"\bwrong\s+number\b",
        r"\bnot\s+available\b",
        r"\bout\s+of\s+service\b",
        r"\binvalid\s+number\b",
        r"\bdisconnected\b",
    ]

    if any(re.search(pattern, text) for pattern in negative_patterns):
        return False

    positive_patterns = [
        r"\bcall\s+connected\b",
        r"\bconnected\s+successfully\b",
        r"\bsuccessfully\s+connected\b",
        r"\bconnected\b",
    ]

    return any(re.search(pattern, text) for pattern in positive_patterns)



def classify_goldcoin_record(record):
    call_1 = record.get("call_1_remarks", "")
    call_2 = record.get("call_2_remarks", "")
    coin_remark = record.get("coin_receive_remark", "")
    excel_status = str(record.get("call_status") or "").strip().lower()

    call_1_connected = is_successful_connection(call_1)
    call_2_connected = is_successful_connection(call_2)
    coin_received = is_gold_coin_received(coin_remark)

    connected = call_1_connected or call_2_connected

    # Completion is separate from successful connection.
    completed = (
        excel_status == "completed"
        or coin_received
    )

    if completed:
        status = "completed"
    elif excel_status == "unclassified":
        status = "unclassified"
    elif excel_status in {
        "pending", "connected", "no_answer",
        "busy", "wrong_number"
    }:
        status = "pending"
    elif connected:
        status = "pending"
    else:
        status = "pending"

    return {
        "status": status,
        "connected": connected,
        "call_1_connected": call_1_connected,
        "call_2_connected": call_2_connected,
        "coin_received": coin_received,
        "successful_attempts": (
            int(call_1_connected) + int(call_2_connected)
        ),
    }



def classify_call_1_remark(remark):
    remark = normalize_remark(remark)

    if not remark or remark in {"none", "nan", "null"}:
        return "No Remark"

    if "call back" in remark or "callback" in remark:
        return "Callback"

    if "busy" in remark:
        return "Busy"

    if "invalid number" in remark:
        return "Invalid Number"

    if "wrong number" in remark:
        return "Wrong Number"

    not_connected_keywords = [
        "not receive",
        "not received",
        "not connected",
        "not answered",
        "not responding",
        "not available",
        "switched off",
        "out of network",
        "out of netork",
        "out of service",
        "no incoming call",
        "no incomming call",
    ]

    if any(keyword in remark for keyword in not_connected_keywords):
        return "Not Connected"

    if "disconnect" in remark:
        return "Disconnected"

    if "call connected" in remark or remark == "connected":
        return "Connected"

    return "Other"


def detect_headers(ws, required=("dealer_name", "phone"), max_rows=30):

    normalized_aliases = {
        field: {
            normalize_header(alias)
            for alias in aliases
        }
        for field, aliases in ALIASES.items()
    }

    last_row = ws.max_row or max_rows
    rows_to_check = min(max_rows, last_row)

    for row_number in range(1, rows_to_check + 1):

        mapping = {}
        headers = []

        for cell in ws[row_number]:

            header = normalize_header(cell.value)
            headers.append(str(cell.value or "").strip())

            if not header:
                continue

            matched_field = None

            for field, aliases in normalized_aliases.items():

                if header in aliases:
                    matched_field = field
                    break

            if matched_field is None and "territory" in header:
                matched_field = "territory"

            if matched_field is None:

                phone_patterns = (
                    "contactnumber",
                    "contactno",
                    "contactmobileno",
                    "contactpersonnumber",
                    "mobilenumber",
                    "mobileno",
                    "phonenumber",
                    "phone",
                    "telephone",
                    "whatsapp",
                    "mobile",
                )

                if any(pattern in header for pattern in phone_patterns):
                    matched_field = "phone"

            if matched_field:
                mapping[matched_field] = cell.column - 1

        if set(required).issubset(mapping):
            return row_number, mapping, headers

    return None, {}, []

def redirect_with_error(request, message, url_name):
    messages.error(request, message)
    return redirect(url_name)


def convert_pod_date(value):
    if not value:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    value = str(value).strip()

    formats = [
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%d.%m.%Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue

    return None


# ============================================================
# EXCEL EXPORT FORMATTING
# ============================================================

def excel_phone_text(ws, row, column):
    cell = ws.cell(row=row, column=column)

    if cell.value is not None:
        cell.value = str(cell.value).strip()

    cell.number_format = "@"
    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
    )


def prepare_sheet(ws, max_width=40):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    header_fill = PatternFill(
        start_color="C9A227",
        end_color="C9A227",
        fill_type="solid",
    )

    header_font = Font(
        bold=True,
        color="FFFFFF",
        size=11,
    )

    header_alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    ws.row_dimensions[1].height = 30

    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)

        for cell in column_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[column_letter].width = min(
            max(max_length + 3, 15),
            max_width,
        )

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True,
            )


def style_header(ws, row=1):
    header_fill = PatternFill(
        start_color=HEADER_COLOR,
        end_color=HEADER_COLOR,
        fill_type="solid",
    )

    for cell in ws[row]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )


def excel_response(workbook, filename="calling_summary.xlsx"):
    output = BytesIO()
    workbook.save(output)
    content = output.getvalue()
    output.close()

    response = HttpResponse(
        content,
        content_type=EXCEL_CONTENT_TYPE,
    )

    response["Content-Disposition"] = (
        f'attachment; filename="{filename}"'
    )
    response["Content-Length"] = str(len(content))

    return response


# ============================================================
# 1. DEALER EXCEL IMPORT
# ============================================================

@login_required
@require_POST
def import_excel(request):
    excel_file = request.FILES.get("excel_file")

    if not excel_file:
        return redirect_with_error(
            request,
            "Please select an Excel file.",
            "dashboard",
        )

    if not excel_file.name.lower().endswith((".xlsx", ".xlsm")):
        return redirect_with_error(
            request,
            "Please upload XLSX or XLSM.",
            "dashboard",
        )

    workbook = None

    try:
        workbook = load_workbook(excel_file, data_only=True)

        worksheet = None
        header_row = None
        mapping = {}

        for sheet in workbook.worksheets:
            detected_row, detected_mapping, headers = detect_headers(
                sheet,
                required=("dealer_name", "phone"),
            )

            if detected_row:
                worksheet = sheet
                header_row = detected_row
                mapping = detected_mapping
                break

        if worksheet is None:
            messages.error(
                request,
                "Dealer Name and Contact Number columns were not found.",
            )
            return redirect("dashboard")

        imported = 0
        updated = 0
        skipped = 0
        errors = 0
        error_details = []

        status_map = {}

        for key, label in DealerCall.STATUS_CHOICES:
            status_map[str(key).strip().lower()] = key
            status_map[str(label).strip().lower()] = key

        blank_values = {"", "na", "n/a", "nan", "none", "null"}

        for row_number, row in enumerate(
            worksheet.iter_rows(min_row=header_row + 1),
            start=header_row + 1,
        ):
            try:
                values = {}

                for field, index in mapping.items():
                    if index >= len(row):
                        continue

                    cell = row[index]

                    if field == "posting_date":
                        values[field] = convert_pod_date(cell.value)
                    else:
                        values[field] = clean_value(
                            cell.value,
                            phone=(field == "phone"),
                            number_format=cell.number_format,
                        )

                if "call_status" in mapping:
                    raw_value = row[mapping["call_status"]].value
                    raw_status = (
                        "" if raw_value is None else str(raw_value).strip()
                    )

                    values["excel_calling_status"] = raw_status

                    if raw_status.lower() in blank_values:
                        values["call_status"] = "unclassified"
                    else:
                        values["call_status"] = status_map.get(
                            raw_status.lower(),
                            "unclassified",
                        )

                coin_receive = values.get("coin_receive_remark", "")
                existing_final = values.get("final_call_remarks", "")

                values["final_call_remarks"] = (
                    "Received"
                    if is_gold_coin_received(coin_receive)
                    else existing_final
                )

                dealer_name = str(values.get("dealer_name") or "").strip()
                dealer_code = str(values.get("dealer_code") or "").strip()
                document_no = str(values.get("document_no") or "").strip()

                if not dealer_name:
                    skipped += 1
                    continue

                if document_no:
                    values["pod_number"] = document_no

                dealer = None

                if dealer_code and document_no:
                    dealer = DealerCall.objects.filter(
                        dealer_code__iexact=dealer_code,
                        document_no__iexact=document_no,
                    ).first()

                elif dealer_code:
                    dealer = DealerCall.objects.filter(
                        dealer_code__iexact=dealer_code,
                        document_no="",
                    ).first()

                elif not document_no:
                    dealer = DealerCall.objects.filter(
                        dealer_name__iexact=dealer_name,
                        document_no="",
                    ).first()

                if dealer:

                    dealer.is_active = True

                    for field, value in values.items():
                        if value not in ("", None) or field in {
                            "call_status",
                            "excel_calling_status",
                        }:
                            setattr(dealer, field, value)

                    dealer.save()
                    updated += 1

                else:
                    DealerCall.objects.create(**values)
                    imported += 1

            except Exception as exc:
                errors += 1
                error_details.append(f"Row {row_number}: {exc}")

        summary = (
            f"Import completed. New: {imported}, Updated: {updated}, "
            f"Skipped: {skipped}, Errors: {errors}"
        )

        if errors:
            messages.warning(request, summary)

            for error in error_details[:5]:
                messages.error(request, error)
        else:
            messages.success(request, summary)

    except Exception as exc:
        messages.error(request, f"Excel import failed: {exc}")

    finally:
        if workbook:
            workbook.close()

    return redirect("dashboard")


# ============================================================
# 2. DEALER EXCEL EXPORT
# ============================================================

@login_required
def export_excel(request):
    workbook = Workbook()
    ws = workbook.active
    ws.title = "GoldCoin Dealer Register"

    ws.append([
        "Sr. No.",
        "Scheme Period",
        "Cities",
        "Location Code",
        "Location Name",
        "Document No.",
        "Posting Date",
        "Dealer Code",
        "Dealer Name",
        "Sales In Charge",
        "Qty",
        "Territory",
        "Address",
        "Contact Number",
        "Email",
        "Contact Name",
        "Recipient Relationship (With Dealer)",
        "POD Received from Yokohama",
        "Audited by Cofiser",
        "Calling Status",
        "Call 1 (Remarks)",
        "Coin Receive Remark",
        "Call 2 (Remarks)",
        "Final Call (Remarks)",
    ])

    for dealer in DealerCall.objects.all().order_by("id"):
        ws.append([
            dealer.sr_no or "",
            dealer.scheme_period or "",
            dealer.city or "",
            dealer.location_code or "",
            dealer.location_name or "",
            dealer.document_no or "",
            dealer.posting_date.strftime("%d-%m-%Y")
            if dealer.posting_date else "",
            dealer.dealer_code or "",
            dealer.dealer_name or "",
            dealer.sales_in_charge or "",
            dealer.qty or "",
            dealer.territory or "",
            dealer.address or "",
            dealer.phone or "",
            dealer.email or "",
            dealer.contact_name or "",
            dealer.recipient_relationship or "",
            dealer.pod_received or "",
            dealer.audited_by_cofiser or "",
            dealer.get_call_status_display(),
            dealer.call_1_remarks or "",
            dealer.coin_receive_remark or "",
            dealer.call_2_remarks or "",
            dealer.final_call_remarks or "",
        ])

        excel_phone_text(ws, ws.max_row, 8)
        excel_phone_text(ws, ws.max_row, 14)

    prepare_sheet(ws, max_width=40)

    return excel_response(
        workbook,
        "GoldCoin_Dealer_24_Columns.xlsx",
    )


# ============================================================
# 3. CALLING SUMMARY REPORT
# ============================================================

@login_required
def export_calling_summary(request):
    workbook = Workbook()
    dealers = DealerCall.objects.all().order_by("dealer_name")
    attempts = DealerCallAttempt.objects.select_related(
        "dealer"
    ).order_by("called_at")

    total_dealers = dealers.count()

    legacy_dealers = dealers.filter(
        called_at__isnull=False,
        call_attempts__isnull=True,
    ).distinct()

    legacy_count = legacy_dealers.count()
    actual_attempt_count = attempts.count()
    total_attempts = actual_attempt_count + legacy_count

    called_dealers = dealers.filter(
        Q(called_at__isnull=False)
        | Q(call_attempts__isnull=False)
    ).distinct()

    unique_called = called_dealers.count()

    pending_count = dealers.filter(
        call_status="pending",
        called_at__isnull=True,
        call_attempts__isnull=True,
    ).distinct().count()

    not_received_count = dealers.filter(
        call_status__in=["no_answer", "busy", "wrong_number"]
    ).distinct().count()

    contacted_count = dealers.filter(
        call_status__in=["connected", "completed"]
    ).distinct().count()

    completion_percentage = (
        round(unique_called / total_dealers * 100, 2)
        if total_dealers else 0
    )

    status_labels = dict(DealerCall.STATUS_CHOICES)

    ws = workbook.active
    ws.title = "Overall Summary"

    ws.append(["GOLDCOIN DEALER CALLING SUMMARY"])
    ws.merge_cells("A1:B1")
    ws["A1"].font = Font(bold=True, size=15, color="FFFFFF")
    ws["A1"].fill = PatternFill(
        start_color=HEADER_COLOR,
        end_color=HEADER_COLOR,
        fill_type="solid",
    )

    ws.append([
        "Generated At",
        timezone.localtime().strftime("%d-%m-%Y %H:%M"),
    ])
    ws.append(["Total Dealers", total_dealers])
    ws.append(["Unique Dealers Called", unique_called])
    ws.append(["Dealers Successfully Contacted", contacted_count])
    ws.append(["Dealers Not Received", not_received_count])
    ws.append(["Pending Dealers", pending_count])
    ws.append(["Total Call Attempts", total_attempts])
    ws.append(["Legacy Recorded Attempts", legacy_count])
    ws.append(["Call Completion Percentage", f"{completion_percentage}%"])
    ws.append([])
    ws.append(["CALLING STATUS", "TOTAL DEALERS"])

    for key, label in DealerCall.STATUS_CHOICES:
        ws.append([label, dealers.filter(call_status=key).count()])

    style_header(ws, row=12)
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 25

    ws2 = workbook.create_sheet("Dealer Details")
    ws2.append([
        "Dealer Name", "Dealer Code", "Phone", "City", "POD Number",
        "Latest Status", "Latest Remark", "Last Called",
        "Total Call Attempts",
    ])

    for dealer in dealers:
        count = dealer.call_attempts.count()

        if count == 0 and dealer.called_at:
            count = 1

        ws2.append([
            dealer.dealer_name or "",
            dealer.dealer_code or "",
            str(dealer.phone or ""),
            dealer.city or "",
            dealer.pod_number or "",
            dealer.get_call_status_display(),
            dealer.remarks or "",
            timezone.localtime(dealer.called_at).strftime(
                "%d-%m-%Y %H:%M"
            ) if dealer.called_at else "Not Called",
            count,
        ])

        excel_phone_text(ws2, ws2.max_row, 3)

    prepare_sheet(ws2)

    ws3 = workbook.create_sheet("Call History")
    ws3.append([
        "Dealer Name", "Dealer Code", "Phone", "Call Status",
        "Remarks", "Call Date", "Call Time", "Record Type",
    ])

    for attempt in attempts:
        if not attempt.dealer:
            continue

        called_at = timezone.localtime(attempt.called_at)

        ws3.append([
            attempt.dealer.dealer_name or "",
            attempt.dealer.dealer_code or "",
            str(attempt.dealer.phone or ""),
            status_labels.get(
                attempt.call_status,
                attempt.call_status,
            ),
            attempt.remarks or "",
            called_at.strftime("%d-%m-%Y"),
            called_at.strftime("%H:%M:%S"),
            "Call Attempt",
        ])

        excel_phone_text(ws3, ws3.max_row, 3)

    for dealer in legacy_dealers:
        called_at = timezone.localtime(dealer.called_at)

        ws3.append([
            dealer.dealer_name or "",
            dealer.dealer_code or "",
            str(dealer.phone or ""),
            dealer.get_call_status_display(),
            dealer.remarks or "",
            called_at.strftime("%d-%m-%Y"),
            called_at.strftime("%H:%M:%S"),
            "Legacy Recorded Attempt",
        ])

        excel_phone_text(ws3, ws3.max_row, 3)

    prepare_sheet(ws3)

    ws4 = workbook.create_sheet("Remark Analysis")
    ws4.append(["Call Status", "Remark", "Total Attempts"])

    remark_counts = {}

    for attempt in attempts:
        key = (
            attempt.call_status,
            (attempt.remarks or "").strip() or "No Remark",
        )
        remark_counts[key] = remark_counts.get(key, 0) + 1

    for dealer in legacy_dealers:
        key = (
            dealer.call_status,
            (dealer.remarks or "").strip() or "No Remark",
        )
        remark_counts[key] = remark_counts.get(key, 0) + 1

    for (status_key, remark), count in sorted(remark_counts.items()):
        ws4.append([
            status_labels.get(status_key, status_key),
            remark,
            count,
        ])

    prepare_sheet(ws4)

    return excel_response(
        workbook,
        "GoldCoin_Calling_Summary.xlsx",
    )





def get_worksheet(wb, *possible_names):

    def normalize(name):
        return re.sub(
            r"[^a-z0-9]",
            "",
            str(name).strip().lower()
        )

    available_sheets = {
        normalize(sheet_name): sheet_name
        for sheet_name in wb.sheetnames
    }

    for name in possible_names:

        normalized = normalize(name)

        if normalized in available_sheets:
            actual_name = available_sheets[normalized]

            return wb[actual_name]

    raise ValueError(
        f"Worksheet not found. "
        f"Expected one of: {possible_names}. "
        f"Available sheets: {wb.sheetnames}"
    )


# 4. AUTOMATIC EXCEL SYNCHRONIZATION
# ============================================================

def sync_goldcoin_excel_file(
    excel_file,
    replace_active_data=False
):

    file_stat = excel_file.stat()

    file_signature = (
        str(excel_file.resolve()),
        file_stat.st_mtime_ns,
        file_stat.st_size,
    )

    file_hash = hashlib.sha256(
        str(excel_file.resolve()).encode()
    ).hexdigest()[:16]

    cache_key = f"goldcoin_excel_last_sync_{file_hash}"

    """ if (
        cache.get(cache_key) == file_signature
        and not DealerCall.objects.filter(is_active=False).exists()
        and not PODTracking.objects.filter(is_active=False).exists()
    ):
        return f"Already synchronized - {excel_file.name} is unchanged." """

    workbook = None

    try:
        workbook = load_workbook(
            excel_file,
            data_only=True,
            read_only=True,
        )

        imported_calls = 0
        updated_calls = 0
        imported_pods = 0
        updated_pods = 0

        status_map = {}

        for key, label in DealerCall.STATUS_CHOICES:
            status_map[str(key).strip().lower()] = key
            status_map[str(label).strip().lower()] = key

        blank_values = {"", "na", "n/a", "nan", "none", "null"}

        with transaction.atomic():

            if replace_active_data:

                DealerCall.objects.filter(
                    is_active=True
                ).update(is_active=False)

                PODTracking.objects.filter(
                    is_active=True
                ).update(is_active=False)

            # ============================================
            # 1. CALLS VERIFICATION SHEET
            # ============================================

            if "Calls_Verification" not in workbook.sheetnames:
                raise ValueError(
                    "Calls_Verification sheet not found."
                )

            ws = workbook["Calls_Verification"]

            header_row, call_mapping, headers = detect_headers(
                ws,
                required=("dealer_name", "dealer_code"),
                max_rows=30,
            )

            if not header_row:
                raise ValueError(
                    "Unable to detect Calls_Verification headers."
                )

            def get_call_value(row, field):

                index = call_mapping.get(field)

                if index is None or index >= len(row):
                    return ""

                cell = row[index]

                if field == "posting_date":
                    return convert_pod_date(cell.value)

                return clean_value(
                    cell.value,
                    phone=(field == "phone"),
                    number_format=cell.number_format,
                )

            for row in ws.iter_rows(min_row=header_row + 1):

                values = {
                    field: get_call_value(row, field)
                    for field in call_mapping
                }

                dealer_code = str(
                    values.get("dealer_code") or ""
                ).strip()

                dealer_name = str(
                    values.get("dealer_name") or ""
                ).strip()

                document_no = str(
                    values.get("document_no") or ""
                ).strip()

                if not dealer_code or not dealer_name:
                    continue

                if document_no:
                    values["pod_number"] = document_no

                # CALL STATUS MAPPING
                if "call_status" in call_mapping:

                    raw_status = str(
                        values.get("call_status") or ""
                    ).strip()

                    if raw_status.lower() not in blank_values:

                        values["excel_calling_status"] = raw_status

                        values["call_status"] = status_map.get(
                            raw_status.lower(),
                            "unclassified",
                        )

                    else:
                        # Do not overwrite an existing CRM status
                        # when the Excel status cell is blank.
                        values.pop("call_status", None)
                        values.pop("excel_calling_status", None)

                # GOLD COIN RECEIVED STATUS
                coin_receive = values.get(
                    "coin_receive_remark",
                    "",
                )

                existing_final = values.get(
                    "final_call_remarks",
                    "",
                )

                values["final_call_remarks"] = (
                    "Received"
                    if is_gold_coin_received(coin_receive)
                    else existing_final
                )

                # FIND EXISTING DEALER RECORD
                if document_no:

                    dealer = DealerCall.objects.filter(
                        dealer_code__iexact=dealer_code,
                        document_no__iexact=document_no,
                    ).first()

                else:

                    dealer = DealerCall.objects.filter(
                        dealer_code__iexact=dealer_code,
                        document_no="",
                    ).first()

                # UPDATE EXISTING RECORD
                if dealer:

                    dealer.is_active = True

                    for field, value in values.items():

                        if (
                            hasattr(dealer, field)
                            and (value not in ("", None)
                                 or field in {
                                     "call_status",
                                     "excel_calling_status",
                                 })
                        ):
                            setattr(dealer, field, value)

                    dealer.save()

                    updated_calls += 1

                # CREATE NEW RECORD
                else:

                    DealerCall.objects.create(**values)

                    imported_calls += 1

            # ============================================
            # 2. REFERENCE POD SHEET
            # ============================================

            ws = get_worksheet(
                workbook,
                "Reference_POD_Sheet",
                "Refference_POD_Sheet",
                "Reference POD Sheet",
                "Refference POD Sheet",
            )

            print("POD WORKSHEET FOUND:", ws.title)


            pod_header_row = None
            pod_mapping = {}

            for row_number in range(
                1,
                min(30, ws.max_row or 30) + 1
            ):

                mapping = {
                    normalize_header(cell.value): cell.column - 1
                    for cell in ws[row_number]
                    if cell.value
                }

                if {
                    "dealercode",
                    "documentno",
                }.issubset(mapping):

                    pod_header_row = row_number
                    pod_mapping = mapping
                    break

            if not pod_header_row:
                raise ValueError(
                    "Reference_POD_Sheet headers not detected."
                )

            def pod_value(row, *names):

                for name in names:

                    index = pod_mapping.get(
                        normalize_header(name)
                    )

                    if index is not None and index < len(row):

                        cell = row[index]

                        return clean_value(
                            cell.value,
                            number_format=cell.number_format,
                        )

                return ""

            # ============================================
            # 3. IMPORT POD DATA
            # ============================================

            for row in ws.iter_rows(
                min_row=pod_header_row + 1
            ):

                dealer_code = pod_value(
                    row,
                    "Dealer Code",
                )

                document_no = pod_value(
                    row,
                    "Document No.",
                    "Document Number",
                )

                if not dealer_code or not document_no:
                    continue

                posting_index = pod_mapping.get("postingdate")
                date_index = pod_mapping.get("date")

                data = {

                    "sr_no": pod_value(
                        row, "Sr. No.", "Sr No"
                    ),

                    "scheme_period": pod_value(
                        row, "Scheme Period"
                    ),

                    "location_code": pod_value(
                        row, "Location Code"
                    ),

                    "posting_date": (
                        convert_pod_date(row[posting_index].value)
                        if posting_index is not None
                        and posting_index < len(row)
                        else None
                    ),

                    "month": pod_value(row, "Month"),

                    "group_code": pod_value(row, "Group Code"),

                    "dealer_name": pod_value(row, "Dealer Name"),

                    "location_name": pod_value(row, "Location Name"),

                    "sales_in_charge": pod_value(
                        row, "Sales In Charge"
                    ),

                    "quarterly_wise": pod_value(
                        row, "Quartly wise", "Quarterly wise"
                    ),

                    "qty": pod_value(row, "Qty"),

                    "received_from_yokohama": pod_value(
                        row, "Received from Yokohama"
                    ),

                    "audited_by_cofiser": pod_value(
                        row, "Audited by Cofiser"
                    ),

                    "sign": pod_value(row, "Sign"),

                    "stamp": pod_value(row, "Stamp"),

                    "date": (
                        convert_pod_date(row[date_index].value)
                        if date_index is not None
                        and date_index < len(row)
                        else None
                    ),

                    "remarks": pod_value(row, "Remarks"),

                    "depots_remark": pod_value(
                        row,
                        "Depot's Remark",
                        "Depots Remark",
                    ),
                }

                # ========================================
                # 4. CREATE OR UPDATE POD RECORD
                # ========================================

                pod, created = PODTracking.objects.get_or_create(
                    dealer_code=dealer_code,
                    document_no=document_no,
                    defaults=data,
                )

                if created:

                    imported_pods += 1

                else:

                    changed = False

                    if not pod.is_active:
                        pod.is_active = True
                        changed = True

                    for field, value in data.items():

                        if (
                            hasattr(pod, field)
                            and value not in ("", None)
                            and getattr(pod, field) != value
                        ):

                            setattr(pod, field, value)

                            changed = True

                    if changed:

                        pod.save()

                        updated_pods += 1

        # ================================================
        # 5. SAVE SUCCESSFUL SYNC STATUS
        # ================================================

        cache.set(
            cache_key,
            file_signature,
            timeout=None,
        )

        return (
            f"Sync successful | "
            f"Calls New: {imported_calls}, "
            f"Calls Updated: {updated_calls}, "
            f"POD New: {imported_pods}, "
            f"POD Updated: {updated_pods}"
        )

    except Exception as exc:

        raise RuntimeError(
            f"Excel synchronization failed: {exc}"
        ) from exc

    finally:

        if workbook is not None:
            workbook.close()


""" def sync_goldcoin_excel():

    excel_files = get_excel_files()

    print("EXCEL FOLDER:", EXCEL_IMPORT_DIR)
    print("EXCEL FILES FOUND:", len(excel_files))

    # No Excel files: archive imported records safely
    if not excel_files:

        with transaction.atomic():

            dealer_count = DealerCall.objects.filter(
                is_active=True
            ).update(is_active=False)

            pod_count = PODTracking.objects.filter(
                is_active=True
            ).update(is_active=False)

        # Invalidate dashboard cache if you use cached statistics.
        cache.delete("goldcoin_dashboard_stats")

        return (
            f"No Excel files found. "
            f"Archived {dealer_count} dealer records and "
            f"{pod_count} POD records."
        )

    results = []

    for excel_file in excel_files:

        try:
            file_result = sync_goldcoin_excel_file(excel_file)

        except Exception as exc:

            file_result = f"ERROR: {str(exc)}"

            print(
                f"Failed to synchronize {excel_file.name}: {exc}"
            )

        print(f"{excel_file.name}: {file_result}")

        results.append(
            f"{excel_file.name}: {file_result}"
        )

    return "\n".join(results) """




def sync_goldcoin_excel():

    excel_files = [
        f for f in get_excel_files()
        if f.suffix.lower() in (".xlsx", ".xlsm")
        and not f.name.startswith("~$")
    ]

    print("EXCEL FOLDER:", EXCEL_IMPORT_DIR)
    print("EXCEL FILES FOUND:", len(excel_files))

    # Archive records only when the folder is genuinely empty.
    if not excel_files:

        with transaction.atomic():

            dealer_count = DealerCall.objects.filter(
                is_active=True
            ).update(is_active=False)

            pod_count = PODTracking.objects.filter(
                is_active=True
            ).update(is_active=False)

        cache.delete("goldcoin_active_source")
        cache.delete("goldcoin_active_signature")

        return (
            f"No Excel files found. "
            f"Archived {dealer_count} dealer records and "
            f"{pod_count} POD records."
        )

    # Required worksheets for dashboard synchronization.
    required_sheets = {
        "Calls_Verification",
        "Reference_POD_Sheet",
    }

    valid_files = []

    # ============================================
    # 1. CHECK ALL EXCEL FILES
    # ============================================

    for excel_file in excel_files:

        wb = None

        try:

            wb = load_workbook(
                excel_file,
                read_only=True,
                data_only=True,
            )

            available_sheets = set(wb.sheetnames)

            missing_sheets = (
                required_sheets - available_sheets
            )

            if missing_sheets:

                print(
                    f"SKIPPING INVALID FILE: {excel_file.name}"
                )

                print(
                    f"MISSING SHEETS: {sorted(missing_sheets)}"
                )

                continue

            print(
                f"VALID EXCEL FILE: {excel_file.name}"
            )

            valid_files.append(excel_file)

        except Exception as exc:

            print(
                f"SKIPPING UNREADABLE FILE: "
                f"{excel_file.name} - {exc}"
            )

        finally:

            if wb is not None:
                wb.close()

    # ============================================
    # 2. STOP IF NO VALID FILE EXISTS
    # ============================================

    if not valid_files:

        raise ValueError(
            "No valid Excel files found. "
            "Required worksheets are "
            "'Calls_Verification' and "
            "'Reference_POD_Sheet'. "
            "Existing dashboard records were not changed."
        )

    # ============================================
    # 3. SELECT LATEST VALID EXCEL FILE
    # ============================================

    latest_file = max(
        valid_files,
        key=lambda f: (
            f.stat().st_mtime_ns,
            f.name.lower()
        )
    )

    latest_path = str(latest_file.resolve())

    print("\nSELECTED LATEST VALID FILE:", latest_file.name)

    # ============================================
    # 4. FILE SIGNATURE
    # ============================================

    file_stat = latest_file.stat()

    latest_signature = (
        file_stat.st_mtime_ns,
        file_stat.st_size,
    )

    active_source = cache.get(
        "goldcoin_active_source"
    )

    active_signature = cache.get(
        "goldcoin_active_signature"
    )

    file_hash = hashlib.sha256(
        latest_path.encode()
    ).hexdigest()[:16]

    file_cache_key = (
        f"goldcoin_excel_last_sync_{file_hash}"
    )

    # ============================================
    # 5. SKIP UNCHANGED ACTIVE FILE
    # ============================================

    if (
        active_source == latest_path
        and active_signature == latest_signature
        and cache.get(file_cache_key) is not None
    ):

        return (
            f"Already synchronized. "
            f"Latest valid file: {latest_file.name}"
        )

    # ============================================
    # 6. REPLACE ACTIVE DATA IF SOURCE CHANGES
    # ============================================

    replace_active_data = (
        active_source != latest_path
        or active_signature != latest_signature
    )

    print("STARTING EXCEL IMPORT...")
    print("REPLACE ACTIVE DATA:", replace_active_data)

    # Import only the latest VALID Excel file.
    result = sync_goldcoin_excel_file(
        latest_file,
        replace_active_data=replace_active_data,
    )

    if str(result).startswith("ERROR"):
        raise RuntimeError(str(result))

    # ============================================
    # 7. SAVE SUCCESSFUL SYNC STATUS
    # ============================================

    cache.set(
        "goldcoin_active_source",
        latest_path,
        timeout=None,
    )

    cache.set(
        "goldcoin_active_signature",
        latest_signature,
        timeout=None,
    )

    return (
        f"Selected latest valid Excel file: "
        f"{latest_file.name}\n"
        f"{result}"
    )


# ============================================================
# 5. DEALER DASHBOARD
# ============================================================


@login_required
def dashboard(request):

    try:
        sync_result = sync_goldcoin_excel()
        print("GOLDCOIN EXCEL SYNC:", sync_result)

    except Exception as exc:
        print("GOLDCOIN EXCEL SYNC ERROR:", str(exc))

    today = timezone.localdate()

    search = request.GET.get("q", "").strip()
    status = request.GET.get("status", "").strip()

    selected_call_date = (
        parse_date(request.GET.get("call_date", "")) or today
    )

    # =====================================================
    # COMPLETE DATABASE RECORDS
    # =====================================================
    all_records = list(
    DealerCall.objects.filter(
        is_active=True
    ).values(
            "id",
            "dealer_code",
            "dealer_name",
            "document_no",
            "call_status",
            "pod_received",
            "call_1_remarks",
            "coin_receive_remark",
            "call_2_remarks",
            "final_call_remarks",
            "called_at",
        )
    )

    total = len(all_records)

    # =====================================================
    # POD + GOLD COIN STATISTICS
    # =====================================================
    pod_received_count = 0
    pending_coin_count = 0
    gold_coin_received_count = 0

    completed = 0
    unclassified = 0
    connected_records = 0

    for record in all_records:

        # ---------------------------------------------
        # POD RECEIVED FROM YOKOHAMA
        # Based on Excel: POD receive from Yokohama
        # ---------------------------------------------
        pod_value = str(
            record.get("pod_received") or ""
        ).strip().lower()

        if pod_value in {
            "yes",
            "received",
            "true",
            "1",
            "y",
            "done",
            "complete",
            "completed",
        }:
            pod_received_count += 1

        # ---------------------------------------------
        # GOLD COIN RECEIVED / PENDING
        # Based ONLY on Coin Receive Remarks
        # ---------------------------------------------
        coin_remark = record.get("coin_receive_remark", "")

        if is_gold_coin_received(coin_remark):
            gold_coin_received_count += 1
        else:
            pending_coin_count += 1

        # ---------------------------------------------
        # CALLING STATISTICS
        # ---------------------------------------------
        call_status = str(
            record.get("call_status") or ""
        ).strip().lower()

        if call_status == "completed":
            completed += 1

        elif call_status == "unclassified":
            unclassified += 1

        classification = classify_goldcoin_record(record)

        if classification.get("connected"):
            connected_records += 1

    # =====================================================
    # CONNECTED CALLS FOR SELECTED DATE
    # =====================================================
    connected_today = DealerCallAttempt.objects.filter(
        call_status="connected",
        called_at__date=selected_call_date,
    ).count()

    # Legacy calls without attempt records
    legacy_connected = DealerCall.objects.filter(
        call_status="connected",
        called_at__date=selected_call_date,
        call_attempts__isnull=True,
    ).count()

    connected_today += legacy_connected

    # =====================================================
    # TABLE FILTERING
    # =====================================================
    queryset = DealerCall.objects.filter(
    is_active=True
).order_by("-id")

    if search:
        queryset = queryset.filter(
            Q(dealer_name__icontains=search)
            | Q(dealer_code__icontains=search)
            | Q(phone__icontains=search)
            | Q(document_no__icontains=search)
            | Q(pod_number__icontains=search)
            | Q(city__icontains=search)
            | Q(location_name__icontains=search)
            | Q(location_code__icontains=search)
            | Q(territory__icontains=search)
            | Q(contact_name__icontains=search)
        )

    valid_statuses = dict(DealerCall.STATUS_CHOICES)

    if status in valid_statuses:
        queryset = queryset.filter(call_status=status)
    else:
        status = ""

    paginator = Paginator(queryset, 20)
    page_obj = paginator.get_page(
        request.GET.get("page")
    )

    # =====================================================
    # FOLLOW-UP STATISTICS
    # =====================================================
    pending_followups = DealerFollowUp.objects.filter(
        status="pending"
    )

    followup_stats = {
        "pending": pending_followups.count(),

        "today": pending_followups.filter(
            followup_date=today
        ).count(),

        "overdue": pending_followups.filter(
            followup_date__lt=today
        ).count(),

        "completed": DealerFollowUp.objects.filter(
            status="completed"
        ).count(),
    }

    upcoming_followups = (
        pending_followups
        .select_related("dealer")
        .order_by("followup_date", "followup_time")[:10]
    )

    # =====================================================
    # DASHBOARD STATISTICS
    # =====================================================
    stats = {
        "total": total,

        # Pending Gold Coin - Coin Receive Remarks
        "pending": pending_coin_count,
        "pending_coin": pending_coin_count,

        # POD Received from Yokohama
        "pod_received": pod_received_count,

        # Gold Coin Received
        "gold_coin_received": gold_coin_received_count,

        # Calling Statistics
        "connected": connected_today,
        "connected_today": connected_today,
        "connected_records": connected_records,
        "completed": completed,
        "unclassified": unclassified,
    }

    # =====================================================
    # CONTEXT
    # =====================================================
    context = {
        "records": page_obj,
        "page_obj": page_obj,
        "paginator": paginator,

        "stats": stats,

        "total_dealers": total,
        "pending_dealers": pending_coin_count,
        "completed_dealers": completed,
        "unclassified_dealers": unclassified,

        "connected_attempts": connected_records,
        "connected_today": connected_today,

        "gold_coin_received": gold_coin_received_count,
        "pod_received": pod_received_count,

        "search": search,
        "status": status,
        "selected_call_date": selected_call_date,

        "followup_stats": followup_stats,
        "upcoming_followups": upcoming_followups,
        "followup_statuses": DealerFollowUp.STATUS_CHOICES,
    }

    return render(
        request,
        "dashboard.html",
        context
    )


# ============================================================
# 6. UPDATE CALL STATUS / CALL HISTORY
# ============================================================

@login_required
@require_POST
def update_call(request, pk):
    try:
        with transaction.atomic():
            dealer = get_object_or_404(
                DealerCall.objects.select_for_update(),
                pk=pk,
            )

            remark_field = request.POST.get("remark_field", "")

            allowed_fields = {
                "call_1_remarks",
                "coin_receive_remark",
                "call_2_remarks",
                "final_call_remarks",
            }

            if remark_field in allowed_fields:
                remark = request.POST.get("remarks", "").strip()

                setattr(dealer, remark_field, remark)
                dealer.save(
                    update_fields=[remark_field, "updated_at"]
                )

                messages.success(
                    request,
                    f"{remark_field.replace('_', ' ').title()} saved.",
                )

            elif request.POST.get("remark_only") == "1":
                dealer.remarks = request.POST.get(
                    "remarks",
                    "",
                ).strip()

                dealer.save(
                    update_fields=["remarks", "updated_at"]
                )

                messages.success(
                    request,
                    "Remark saved successfully.",
                )

            else:
                new_status = request.POST.get(
                    "call_status",
                    "",
                ).strip().lower()

                valid_statuses = {
                    key for key, label in DealerCall.STATUS_CHOICES
                }

                if new_status not in valid_statuses:
                    messages.error(
                        request,
                        "Invalid calling status.",
                    )
                    return redirect("dashboard")

                current_time = timezone.now()

                dealer.call_status = new_status
                dealer.called_at = current_time

                dealer.save(
                    update_fields=[
                        "call_status",
                        "called_at",
                        "updated_at",
                    ]
                )

                DealerCallAttempt.objects.create(
                    dealer=dealer,
                    call_status=new_status,
                    remarks=dealer.remarks or "",
                    called_at=current_time,
                )

                messages.success(
                    request,
                    f"Call updated for {dealer.dealer_name}.",
                )

    except Exception as exc:
        messages.error(
            request,
            f"Unable to update call: {exc}",
        )

    return redirect("dashboard")


# ============================================================
# 7. FOLLOW-UP DASHBOARD AND HELPERS
# ============================================================

@login_required
def dealer_followups(request):
    today = timezone.localdate()

    queryset = DealerFollowUp.objects.select_related(
        "dealer"
    ).order_by(
        "followup_date",
        "followup_time",
    )

    search = request.GET.get("q", "").strip()
    status = request.GET.get("status", "").strip()

    if search:
        queryset = queryset.filter(
            Q(dealer__dealer_name__icontains=search)
            | Q(dealer__dealer_code__icontains=search)
            | Q(dealer__phone__icontains=search)
            | Q(dealer__pod_number__icontains=search)
            | Q(dealer__city__icontains=search)
            | Q(dealer_name__icontains=search)
            | Q(contact_number__icontains=search)
        )

    valid_statuses = dict(DealerFollowUp.STATUS_CHOICES)

    if status in valid_statuses:
        queryset = queryset.filter(status=status)
    else:
        status = ""

    pending = DealerFollowUp.objects.filter(status="pending")

    stats = {
        "pending": pending.count(),
        "today": pending.filter(followup_date=today).count(),
        "overdue": pending.filter(followup_date__lt=today).count(),
        "completed": DealerFollowUp.objects.filter(
            status="completed"
        ).count(),
    }

    return render(
        request,
        "followups.html",
        {
            "followups": queryset,
            "search": search,
            "status": status,
            "today": today,
            "followup_stats": stats,
            "followup_statuses": DealerFollowUp.STATUS_CHOICES,
        },
    )


def parse_followup_datetime(date_value, time_value):
    selected_date = parse_date(date_value)
    selected_time = parse_time(time_value)

    if not selected_date or not selected_time:
        return None

    return timezone.make_aware(
        datetime.combine(selected_date, selected_time),
        timezone.get_current_timezone(),
    )


@login_required
def schedule_followup(request, dealer_id):
    dealer = get_object_or_404(DealerCall, pk=dealer_id)

    if request.method == "POST":
        selected_datetime = parse_followup_datetime(
            request.POST.get("followup_date", "").strip(),
            request.POST.get("followup_time", "").strip(),
        )

        remarks = request.POST.get("remarks", "").strip()

        if not selected_datetime:
            messages.error(
                request,
                "Please enter a valid follow-up date and time.",
            )
            return redirect(
                "schedule_followup",
                dealer_id=dealer.pk,
            )

        if selected_datetime < timezone.now():
            messages.error(
                request,
                "Follow-up date and time cannot be in the past.",
            )
            return redirect(
                "schedule_followup",
                dealer_id=dealer.pk,
            )

        DealerFollowUp.objects.create(
            dealer=dealer,
            dealer_name=dealer.dealer_name or "",
            contact_number=dealer.phone or "",
            calling_status=dealer.call_status or "",
            followup_date=selected_datetime.date(),
            followup_time=selected_datetime.time(),
            remarks=remarks,
            status="pending",
        )

        messages.success(
            request,
            f"Follow-up scheduled for {dealer.dealer_name}.",
        )

        return redirect("dealer_followups")

    return render(
        request,
        "schedule_followup.html",
        {
            "dealer": dealer,
            "today": timezone.localdate(),
        },
    )


@login_required
@require_POST
def complete_followup(request, followup_id):
    followup = get_object_or_404(
        DealerFollowUp,
        pk=followup_id,
    )

    if followup.status != "pending":
        messages.warning(
            request,
            "This follow-up is already closed.",
        )
        return redirect("dealer_followups")

    followup.status = "completed"
    followup.completed_at = timezone.now()

    followup.save(
        update_fields=["status", "completed_at"]
    )

    dealer_name = (
        followup.dealer.dealer_name
        if followup.dealer
        else followup.dealer_name or "dealer"
    )

    messages.success(
        request,
        f"Follow-up completed for {dealer_name}.",
    )

    return redirect("dealer_followups")


@login_required
def reschedule_followup(request, followup_id):
    followup = get_object_or_404(
        DealerFollowUp,
        pk=followup_id,
    )

    if followup.status != "pending":
        messages.error(
            request,
            "Only pending follow-ups can be rescheduled.",
        )
        return redirect("dealer_followups")

    if request.method == "POST":
        selected_datetime = parse_followup_datetime(
            request.POST.get("followup_date", "").strip(),
            request.POST.get("followup_time", "").strip(),
        )

        if not selected_datetime:
            messages.error(
                request,
                "Please enter a valid date and time.",
            )
            return redirect(
                "reschedule_followup",
                followup_id=followup.pk,
            )

        if selected_datetime < timezone.now():
            messages.error(
                request,
                "Follow-up cannot be scheduled in the past.",
            )
            return redirect(
                "reschedule_followup",
                followup_id=followup.pk,
            )

        followup.followup_date = selected_datetime.date()
        followup.followup_time = selected_datetime.time()

        remarks = request.POST.get("remarks", "").strip()

        if remarks:
            followup.remarks = remarks

        followup.save(
            update_fields=[
                "followup_date",
                "followup_time",
                "remarks",
            ]
        )

        messages.success(
            request,
            "Follow-up rescheduled successfully.",
        )

        return redirect("dealer_followups")

    return render(
        request,
        "reschedule_followup.html",
        {
            "followup": followup,
            "today": timezone.localdate(),
        },
    )


@login_required
@require_POST
def cancel_followup(request, followup_id):
    followup = get_object_or_404(
        DealerFollowUp,
        pk=followup_id,
    )

    if followup.status == "pending":
        followup.status = "cancelled"
        followup.save(update_fields=["status"])

        messages.success(
            request,
            "Follow-up cancelled successfully.",
        )
    else:
        messages.warning(
            request,
            "Only pending follow-ups can be cancelled.",
        )

    return redirect("dealer_followups")


@login_required
def dealer_followup_history(request, dealer_id):
    dealer = get_object_or_404(
        DealerCall,
        pk=dealer_id,
    )

    history = DealerFollowUp.objects.filter(
        dealer=dealer
    ).order_by("-created_at")

    return render(
        request,
        "followup_history.html",
        {
            "dealer": dealer,
            "history": history,
        },
    )


# ============================================================
# 8. DEALER EDIT / DELETE
# ============================================================

@login_required
def edit_dealer(request, dealer_id):
    dealer = get_object_or_404(
        DealerCall,
        pk=dealer_id,
    )

    if request.method == "POST":
        dealer.dealer_name = request.POST.get(
            "dealer_name",
            "",
        ).strip()

        dealer.dealer_code = request.POST.get(
            "dealer_code",
            "",
        ).strip()

        dealer.phone = request.POST.get(
            "phone",
            "",
        ).strip()

        dealer.city = request.POST.get(
            "city",
            "",
        ).strip()

        dealer.pod_number = request.POST.get(
            "pod_number",
            "",
        ).strip()

        dealer.save()

        messages.success(
            request,
            "Dealer details updated successfully.",
        )

        return redirect("dashboard")

    return render(
        request,
        "edit_dealer.html",
        {"dealer": dealer},
    )


@login_required
@require_POST
def delete_dealer(request, dealer_id):
    dealer = get_object_or_404(
        DealerCall,
        pk=dealer_id,
    )

    dealer.delete()

    messages.success(
        request,
        "Dealer deleted successfully.",
    )

    return redirect("dashboard")


# ============================================================
# 9. AUTHENTICATION
# ============================================================

@login_required
def logout_user(request):
    if request.method == "POST":
        logout(request)

    return redirect("login")


def register(request):
    if request.user.is_authenticated:
        return redirect("dashboard")

    if request.method == "POST":
        form = UserCreationForm(request.POST)

        if form.is_valid():
            user = form.save()
            login(request, user)

            messages.success(
                request,
                "Registration successful. Welcome!",
            )

            return redirect("dashboard")
    else:
        form = UserCreationForm()

    return render(
        request,
        "register.html",
        {"form": form},
    )


# ============================================================
# 10. POD EXCEL IMPORT
# ============================================================

@login_required
@require_POST
def import_pod_excel(request):
    excel_file = request.FILES.get("excel_file")

    if not excel_file:
        return redirect_with_error(
            request,
            "Please select an Excel file.",
            "pod_tracking",
        )

    if not excel_file.name.lower().endswith((".xlsx", ".xlsm")):
        return redirect_with_error(
            request,
            "Please upload an XLSX or XLSM file.",
            "pod_tracking",
        )

    workbook = None

    try:
        workbook = load_workbook(
            excel_file,
            data_only=True,
        )

        worksheet = None
        header_row = None
        headers = {}

        required = {"dealercode", "documentno"}

        for sheet in workbook.worksheets:
            for row_number in range(
                1,
                min(30, sheet.max_row) + 1,
            ):
                mapping = {
                    normalize_header(cell.value): cell.column - 1
                    for cell in sheet[row_number]
                    if cell.value
                }

                if required.issubset(mapping):
                    worksheet = sheet
                    header_row = row_number
                    headers = mapping
                    break

            if worksheet is not None:
                break

        if worksheet is None:
            messages.error(
                request,
                "Dealer Code or Document No. column not found.",
            )
            return redirect("pod_tracking")

        def get_value(row, *names):
            for name in names:
                index = headers.get(normalize_header(name))

                if index is not None and index < len(row):
                    cell = row[index]

                    return clean_value(
                        cell.value,
                        number_format=cell.number_format,
                    )

            return ""

        imported = 0
        updated = 0
        skipped = 0

        with transaction.atomic():
            for row_number, row in enumerate(
                worksheet.iter_rows(min_row=header_row + 1),
                start=header_row + 1,
            ):
                try:
                    dealer_code = get_value(row, "Dealer Code")
                    document_no = get_value(
                        row,
                        "Document No.",
                        "Document Number",
                    )

                    if not dealer_code or not document_no:
                        skipped += 1
                        continue

                    posting_index = headers.get("postingdate")
                    date_index = headers.get("date")

                    data = {
                        "sr_no": get_value(row, "Sr. No.", "Sr No"),
                        "scheme_period": get_value(row, "Scheme Period"),
                        "location_code": get_value(row, "Location Code"),
                        "posting_date": (
                            convert_pod_date(row[posting_index].value)
                            if posting_index is not None
                            and posting_index < len(row)
                            else None
                        ),
                        "month": get_value(row, "Month"),
                        "group_code": get_value(row, "Group Code"),
                        "dealer_name": get_value(row, "Dealer Name"),
                        "location_name": get_value(row, "Location Name"),
                        "sales_in_charge": get_value(
                            row,
                            "Sales In Charge",
                        ),
                        "quarterly_wise": get_value(
                            row,
                            "Quartly wise",
                            "Quarterly wise",
                        ),
                        "qty": get_value(row, "Qty"),
                        "received_from_yokohama": get_value(
                            row,
                            "Received from Yokohama",
                        ),
                        "audited_by_cofiser": get_value(
                            row,
                            "Audited by Cofiser",
                        ),
                        "sign": get_value(row, "Sign"),
                        "stamp": get_value(row, "Stamp"),
                        "date": (
                            convert_pod_date(row[date_index].value)
                            if date_index is not None
                            and date_index < len(row)
                            else None
                        ),
                        "remarks": get_value(row, "Remarks"),
                        "depots_remark": get_value(
                            row,
                            "Depot's Remark",
                            "Depots Remark",
                        ),
                    }

                    _, created = PODTracking.objects.update_or_create(
                        dealer_code=dealer_code,
                        document_no=document_no,
                        defaults=data,
                    )

                    if created:
                        imported += 1
                    else:
                        updated += 1

                except Exception as exc:
                    skipped += 1
                    print(
                        f"POD import error at row {row_number}: {exc}"
                    )

        messages.success(
            request,
            f"POD import completed. New: {imported}, "
            f"Updated: {updated}, Skipped: {skipped}",
        )

    except Exception as exc:
        messages.error(
            request,
            f"POD Excel import failed: {exc}",
        )

    finally:
        if workbook:
            workbook.close()

    return redirect("pod_tracking")


# ============================================================
# 11. POD TRACKING DASHBOARD
# ============================================================

@login_required
def pod_tracking(request):

    try:
        sync_result = sync_goldcoin_excel()
        print("POD PAGE EXCEL SYNC:", sync_result)

    except Exception as exc:
        print("POD PAGE EXCEL SYNC ERROR:", str(exc))
    all_pods = PODTracking.objects.all()
    records = all_pods.order_by("-id")

    search = request.GET.get("search", "").strip()
    scheme = request.GET.get("scheme", "").strip()
    month = request.GET.get("month", "").strip()
    location = request.GET.get("location", "").strip()
    document_status = request.GET.get("document_status", "").strip()

    if search:
        records = records.filter(
            Q(dealer_code__icontains=search)
            | Q(dealer_name__icontains=search)
            | Q(document_no__icontains=search)
            | Q(location_name__icontains=search)
        )

    if scheme:
        records = records.filter(scheme_period=scheme)

    if month:
        records = records.filter(month=month)

    if location:
        records = records.filter(location_code=location)

    # Calculate POD flags consistently using the model's is_marked().
    all_pod_objects = list(all_pods)

    total = len(all_pod_objects)
    signed_count = sum(pod.is_signed for pod in all_pod_objects)
    stamped_count = sum(pod.is_stamped for pod in all_pod_objects)
    dated_count = sum(pod.is_dated for pod in all_pod_objects)
    complete_count = sum(
        pod.document_complete for pod in all_pod_objects
    )

    # Apply document status filter.
    if document_status == "complete":
        complete_ids = [
            pod.pk for pod in all_pod_objects
            if pod.document_complete
        ]
        records = records.filter(pk__in=complete_ids)

    elif document_status == "pending":
        pending_ids = [
            pod.pk for pod in all_pod_objects
            if not pod.document_complete
        ]
        records = records.filter(pk__in=pending_ids)

    paginator = Paginator(records, 20)
    page_obj = paginator.get_page(request.GET.get("page"))

    context = {
        "records": page_obj,
        "page_obj": page_obj,
        "paginator": paginator,

        "total_pods": total,
        "complete_pods": complete_count,
        "pending_pods": total - complete_count,
        "signed_pods": signed_count,
        "stamped_pods": stamped_count,
        "dated_pods": dated_count,

        "schemes": all_pods.exclude(
            scheme_period=""
        ).values_list(
            "scheme_period",
            flat=True,
        ).distinct(),

        "months": all_pods.exclude(
            month=""
        ).values_list(
            "month",
            flat=True,
        ).distinct(),

        "locations": all_pods.exclude(
            location_code=""
        ).values_list(
            "location_code",
            flat=True,
        ).distinct(),

        "search": search,
        "scheme": scheme,
        "month": month,
        "location": location,
        "document_status": document_status,
    }

    return render(
        request,
        "calls/pod_tracking.html",
        context,
    )


# ============================================================
# 12. UPDATE POD
# ============================================================

@login_required
@require_POST
def update_pod(request, pk):
    pod = get_object_or_404(
        PODTracking,
        pk=pk,
    )

    pod.received_from_yokohama = request.POST.get(
        "received_from_yokohama",
        "",
    ).strip()

    pod.audited_by_cofiser = request.POST.get(
        "audited_by_cofiser",
        "",
    ).strip()

    pod.sign = request.POST.get("sign", "").strip()
    pod.stamp = request.POST.get("stamp", "").strip()
    pod.date = convert_pod_date(request.POST.get("date", ""))
    pod.remarks = request.POST.get("remarks", "").strip()
    pod.depots_remark = request.POST.get(
        "depots_remark",
        "",
    ).strip()

    pod.save()

    messages.success(
        request,
        "POD details updated successfully.",
    )

    return redirect("pod_tracking")


# ============================================================
# 13. EXPORT POD EXCEL
# ============================================================

@login_required
def export_pod_excel(request):
    workbook = Workbook()
    ws = workbook.active
    ws.title = "POD Tracking"

    ws.append([
        "Sr. No.",
        "Scheme Period",
        "Location Code",
        "Document No.",
        "Posting Date",
        "Month",
        "Dealer Code",
        "Group Code",
        "Dealer Name",
        "Location Name",
        "Sales In Charge",
        "Quarterly Wise",
        "Qty",
        "Received from Yokohama",
        "Audited by Cofiser",
        "Sign",
        "Stamp",
        "Date",
        "Remarks",
        "Depot's Remark",
    ])

    pods = PODTracking.objects.all().order_by("id")

    for pod in pods:
        ws.append([
            pod.sr_no or "",
            pod.scheme_period or "",
            pod.location_code or "",
            pod.document_no or "",
            pod.posting_date.strftime("%d-%m-%Y")
            if pod.posting_date else "",
            pod.month or "",
            pod.dealer_code or "",
            pod.group_code or "",
            pod.dealer_name or "",
            pod.location_name or "",
            pod.sales_in_charge or "",
            pod.quarterly_wise or "",
            pod.qty if pod.qty is not None else "",
            pod.received_from_yokohama or "",
            pod.audited_by_cofiser or "",
            pod.sign or "",
            pod.stamp or "",
            pod.date.strftime("%d-%m-%Y")
            if pod.date else "",
            pod.remarks or "",
            pod.depots_remark or "",
        ])

        for column in [3, 4, 7, 8]:
            ws.cell(
                row=ws.max_row,
                column=column,
            ).number_format = "@"

    prepare_sheet(ws, max_width=35)

    return excel_response(
        workbook,
        "GoldCoin_POD_Tracking.xlsx",
    )


