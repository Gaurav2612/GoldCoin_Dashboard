
import re
from decimal import Decimal, InvalidOperation
from django.core.paginator import Paginator
from datetime import datetime
from .models import DealerCall, DealerFollowUp

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.db.models import Q
from django.shortcuts import render, get_object_or_404, redirect
from django.utils import timezone
from django.http import HttpResponse, HttpResponseBadRequest

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from .models import DealerCall
from django.contrib.auth import login
from django.contrib.auth.forms import UserCreationForm
from django.contrib.auth.views import LoginView


# ==================================================
# 1. EXCEL COLUMN ALIASES
# ==================================================


ALIASES = {
    "dealer_name": ["dealer name", "name"],

    "dealer_code": ["dealer code", "code"],

    "phone": [
        "phone",
        "mobile",
        "contact number",
        "contact no",
        "contact",
        "contact num",
        "mobile number",
        "phone number",
        "telephone",
        "dealer phone",
        "dealer contact",
        "whatsapp number"
    ],

    "city": [
        "city",
        "location",
        "location name"
    ],

    "pod_number": [
        "pod number",
        "pod",
        "document no",
        "document number"
    ],

    "call_status": [
        "call status",
        "calling status",
        "status"
    ],

    "remarks": [
        "remarks",
        "remark",
        "comments"
    ],
}


# ==================================================
# 2. NORMALIZE EXCEL HEADERS
# ==================================================

def normalize_header(value):

    return re.sub(
        r"[^a-z0-9]",
        "",
        str(value or "").lower()
    )


# ==================================================
# 3. CLEAN EXCEL VALUES
# ==================================================

def clean_value(value, phone=False, number_format=None):

    if value is None:
        return ""

    if phone:

        if isinstance(value, bool):
            return ""

        if isinstance(value, (int, float, Decimal)):

            try:
                number = Decimal(str(value))

                if not number.is_finite():
                    return ""

                number = format(number, "f")

                if "." in number:
                    number = number.rstrip("0").rstrip(".")

                # Preserve leading zeros where Excel
                # uses a simple zero number format.
                if number_format and re.fullmatch(
                    r"0+",
                    number_format
                ):
                    number = number.zfill(len(number_format))

                return number

            except (InvalidOperation, ValueError):
                return ""

        number = str(value).strip()

        if number.lower() in ("none","nan","null"):
            return ""
        if re.fullmatch(r"\d+\.0+", number):
            number = number.split(".")[0]

        return number

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


# ==================================================
# 4. DETECT EXCEL HEADERS
# ==================================================



def detect_headers(ws):
    """
    Detect dealer Excel headers and return:
    header row number, column mapping, and header names.
    """

    for row_num in range(1, 31):
        mapping = {}
        headers = []

        for cell in ws[row_num]:
            header = normalize_header(cell.value)
            headers.append(str(cell.value or "").strip())

            if not header:
                continue

            for field, aliases in ALIASES.items():
                normalized_aliases = [
                    normalize_header(alias)
                    for alias in aliases
                ]

                if header in normalized_aliases:
                    mapping[field] = cell.column - 1
                    break

        if "dealer_name" in mapping and "phone" in mapping:
            print("EXCEL HEADER ROW:", row_num)
            print("EXCEL HEADERS:", headers)
            print("EXCEL COLUMN MAPPING:", mapping)

            return row_num, mapping, headers

    return None, {}, []
# ==================================================
# 5. IMPORT EXCEL
# ==================================================



@login_required
def import_excel(request):

    if request.method != "POST":
        return HttpResponseBadRequest("Invalid request method.")

    if "excel_file" not in request.FILES:
        messages.error(request, "Please select an Excel file.")
        return redirect("dashboard")

    file = request.FILES["excel_file"]

    if not file.name.lower().endswith((".xlsx", ".xlsm")):
        messages.error(request, "Please upload a valid Excel file.")
        return redirect("dashboard")

    if file.size == 0:
        messages.error(request, "The uploaded file is empty.")
        return redirect("dashboard")

    wb = None

    try:
        wb = load_workbook(
            file,
            data_only=True,
            read_only=False
        )

        ws = None
        header_row = None
        mapping = {}

        # Detect headers in all worksheets
        for sheet in wb.worksheets:

            detected_row, detected_mapping, detected_headers = (
                detect_headers(sheet)
            )

            if (
                detected_row
                and "dealer_name" in detected_mapping
                and "phone" in detected_mapping
            ):
                ws = sheet
                header_row = detected_row
                mapping = detected_mapping
                break

        if ws is None:
            messages.error(
                request,
                "Dealer Name or Contact Number column not found. "
                "Please check your Excel headings."
            )
            return redirect("dashboard")

        print("SELECTED SHEET:", ws.title)
        print("HEADER ROW:", header_row)
        print("COLUMN MAPPING:", mapping)

        imported = 0
        updated = 0
        skipped = 0
        phone_updated = 0
        errors = 0

        
        for row_number, row in enumerate(
            ws.iter_rows(min_row=header_row + 1),
            start=header_row + 1
        ):

            try:
                values = {}

                for field, index in mapping.items():

                    if index >= len(row):
                        continue

                    cell = row[index]

                    values[field] = clean_value(
                        cell.value,
                        phone=(field == "phone"),
                        number_format=cell.number_format
                    )

                # NORMALIZE EXCEL CALLING STATUS
                if "call_status" in values:

                    excel_status = values["call_status"].strip().lower()

                    matched_status = ""

                    for key, label in DealerCall.STATUS_CHOICES:

                        if (
                            excel_status == str(key).lower()
                            or excel_status == str(label).lower()
                        ):
                            matched_status = key
                            break

                    if matched_status:
                        values["call_status"] = matched_status
                    else:
                        values.pop("call_status")

                # CONTINUE YOUR EXISTING IMPORT CODE
                dealer_name = values.get(
                    "dealer_name", ""
                ).strip()


                if not dealer_name:
                    skipped += 1
                    continue

                dealer_code = values.get(
                    "dealer_code", ""
                ).strip()

                pod_number = values.get(
                    "pod_number", ""
                ).strip()

                phone = values.get("phone", "").strip()

                dealer = None

                # 1. Match by dealer code
                if dealer_code:
                    dealer = DealerCall.objects.filter(
                        dealer_code__iexact=dealer_code
                    ).first()

                # 2. Match by dealer name and POD
                if dealer is None:

                    existing = DealerCall.objects.filter(
                        dealer_name__iexact=dealer_name
                    )

                    if pod_number:
                        existing = existing.filter(
                            pod_number__iexact=pod_number
                        )

                    dealer = existing.first()

                # 3. Update existing dealer
                if dealer:

                    old_phone = str(dealer.phone or "").strip()

                    for field, value in values.items():

                        # Do not overwrite existing values with blanks
                        if value == "":
                            continue

                        setattr(dealer, field, value)

                    if phone and old_phone != phone:
                        phone_updated += 1

                    dealer.save()
                    updated += 1

                # 4. Create new dealer
                else:

                    DealerCall.objects.create(**values)
                    imported += 1

            except Exception as row_error:

                errors += 1

                print(
                    f"ROW {row_number} IMPORT ERROR:",
                    row_error
                )

        messages.success(
            request,
            (
                f"Excel import completed! "
                f"New: {imported}, "
                f"Updated: {updated}, "
                f"Phone numbers updated: {phone_updated}, "
                f"Skipped: {skipped}, "
                f"Errors: {errors}"
            )
        )

    except Exception as exc:

        messages.error(
            request,
            f"Excel import failed: {exc}"
        )

        print("IMPORT ERROR:", exc)

    finally:

        if wb is not None:
            wb.close()

    return redirect("dashboard")


# ==================================================
# 6. EXPORT EXCEL
# ==================================================

@login_required
def export_excel(request):

    wb = Workbook()

    ws = wb.active
    ws.title = "Dealer Calls"

    headers = [
        "dealer_name",
        "dealer_code",
        "phone",
        "city",
        "pod_number",
        "call_status",
        "remarks",
        "called_at"
    ]

    ws.append(headers)

    # Header formatting
    for cell in ws[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            start_color="1F4E78",
            end_color="1F4E78",
            fill_type="solid"
        )

        cell.alignment = Alignment(
            horizontal="center"
        )

    dealers = DealerCall.objects.all().order_by("-id")

    for dealer in dealers:

        ws.append([
            dealer.dealer_name or "",
            dealer.dealer_code or "",
            str(dealer.phone or ""),
            dealer.city or "",
            dealer.pod_number or "",
            dealer.get_call_status_display(),
            dealer.remarks or "",
            dealer.called_at.strftime("%d-%m-%Y %H:%M")
            if dealer.called_at else ""
        ])

        # Keep phone numbers as text in Excel
        phone_cell = ws.cell(
            row=ws.max_row,
            column=3
        )

        phone_cell.number_format = "@"

    # Automatic column width
    for column in ws.columns:

        max_length = max(
            len(str(cell.value or ""))
            for cell in column
        )

        column_letter = get_column_letter(
            column[0].column
        )

        ws.column_dimensions[column_letter].width = min(
            max(max_length + 3, 15),
            35
        )

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    response = HttpResponse(
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

    response["Content-Disposition"] = (
        'attachment; filename="goldcoin_dealer_calls.xlsx"'
    )

    wb.save(response)

    return response


# ==================================================
# 7. DASHBOARD
# ==================================================



@login_required
def dashboard(request):

    today = timezone.localdate()

    qs = DealerCall.objects.all().order_by("-id")

    search = request.GET.get("q", "").strip()
    status = request.GET.get("status", "").strip()

    # Search dealers
    if search:
        qs = qs.filter(
            Q(dealer_name__icontains=search) |
            Q(dealer_code__icontains=search) |
            Q(phone__icontains=search) |
            Q(pod_number__icontains=search) |
            Q(city__icontains=search)
        )

    # Filter calling status
    valid_statuses = dict(DealerCall.STATUS_CHOICES)

    if status in valid_statuses:
        qs = qs.filter(call_status=status)
    else:
        status = ""

    # Statistics
    stats = {
        "total": DealerCall.objects.count(),
        "pending": DealerCall.objects.filter(call_status="pending").count(),
        "connected": DealerCall.objects.filter(call_status="connected").count(),
        "completed": DealerCall.objects.filter(call_status="completed").count(),
        "no_answer": DealerCall.objects.filter(call_status="no_answer").count(),
    }

    # PAGINATION: 20 dealers per page
    paginator = Paginator(qs, 20)

    page_number = request.GET.get("page")

    page_obj = paginator.get_page(page_number)

    # Follow-up statistics
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

    return render(
        request,
        "dashboard.html",
        {
            "dealers": page_obj,
            "page_obj": page_obj,
            "paginator": paginator,
            "stats": stats,
            "search": search,
            "status": status,
            "status_choices": DealerCall.STATUS_CHOICES,
            "followup_stats": followup_stats,
            "upcoming_followups": upcoming_followups,
            "today": today,
        }
    )




# ==================================================
# 8. UPDATE DEALER CALL STATUS
# ==================================================

@login_required
def update_call(request, pk):

    if request.method != "POST":
        return redirect("dashboard")

    dealer = get_object_or_404(
        DealerCall,
        pk=pk
    )

    new_status = request.POST.get(
        "call_status", ""
    ).strip()

    valid_statuses = dict(
        DealerCall.STATUS_CHOICES
    )

    if new_status not in valid_statuses:

        messages.error(
            request,
            "Please select a valid call status."
        )

        return redirect("dashboard")

    dealer.call_status = new_status

    dealer.remarks = request.POST.get(
        "remarks", ""
    ).strip()

    dealer.called_at = timezone.now()

    dealer.save()

    messages.success(
        request,
        f"Call update saved for {dealer.dealer_name}."
    )

    return redirect("dashboard")


# ==================================================
# 9. DEALER FOLLOW-UP DASHBOARD
# ==================================================

@login_required
def dealer_followups(request):

    today = timezone.localdate()

    qs = DealerFollowUp.objects.select_related(
        "dealer"
    ).order_by("followup_date", "followup_time")

    search = request.GET.get("q", "").strip()
    status = request.GET.get("status", "").strip()

    # Search dealer
    if search:
        qs = qs.filter(
            Q(dealer__dealer_name__icontains=search) |
            Q(dealer__dealer_code__icontains=search) |
            Q(dealer__phone__icontains=search) |
            Q(dealer__pod_number__icontains=search) |
            Q(dealer__city__icontains=search)
        )

    # Filter status
    valid_statuses = dict(DealerFollowUp.STATUS_CHOICES)

    if status in valid_statuses:
        qs = qs.filter(status=status)
    else:
        status = ""

    # Statistics
    pending = DealerFollowUp.objects.filter(
        status="pending"
    )

    followup_stats = {
        "pending": pending.count(),

        "today": pending.filter(
            followup_date=today
        ).count(),

        "overdue": pending.filter(
            followup_date__lt=today
        ).count(),

        "completed": DealerFollowUp.objects.filter(
            status="completed"
        ).count(),
    }

    return render(
        request,
        "followups.html",
        {
            "followups": qs,
            "search": search,
            "status": status,
            "today": today,
            "followup_stats": followup_stats,
            "followup_statuses": DealerFollowUp.STATUS_CHOICES,
        }
    )


# ==================================================
# 10. SCHEDULE DEALER FOLLOW-UP
# ==================================================

@login_required
def schedule_followup(request, dealer_id):

    dealer = get_object_or_404(
        DealerCall,
        id=dealer_id
    )

    if request.method == "POST":

        followup_date = request.POST.get(
            "followup_date", ""
        ).strip()

        followup_time = request.POST.get(
            "followup_time", ""
        ).strip()

        remarks = request.POST.get(
            "remarks", ""
        ).strip()

        # Validate required fields
        if not followup_date or not followup_time:

            messages.error(
                request,
                "Please select follow-up date and time."
            )

            return redirect(
                "schedule_followup",
                dealer_id=dealer.id
            )

        try:

            selected_date = datetime.strptime(
                followup_date,
                "%Y-%m-%d"
            ).date()

            selected_time = datetime.strptime(
                followup_time,
                "%H:%M"
            ).time()

        except ValueError:

            messages.error(
                request,
                "Invalid date or time format."
            )

            return redirect(
                "schedule_followup",
                dealer_id=dealer.id
            )

        # Prevent scheduling in the past
        selected_datetime = timezone.make_aware(
            datetime.combine(
                selected_date,
                selected_time
            ),
            timezone.get_current_timezone()
        )

        if selected_datetime < timezone.now():

            messages.error(
                request,
                "Follow-up date and time cannot be in the past."
            )

            return redirect(
                "schedule_followup",
                dealer_id=dealer.id
            )

        # Create follow-up
        DealerFollowUp.objects.create(
            dealer=dealer,
            followup_date=selected_date,
            followup_time=selected_time,
            remarks=remarks,
            status="pending"
        )

        messages.success(
            request,
            f"Follow-up scheduled successfully for {dealer.dealer_name}."
        )

        return redirect("dealer_followups")

    return render(
        request,
        "schedule_followup.html",
        {
            "dealer": dealer,
            "today": timezone.localdate(),
        }
    )


# ==================================================
# 11. COMPLETE FOLLOW-UP
# ==================================================

@login_required
def complete_followup(request, followup_id):

    if request.method != "POST":
        return redirect("dealer_followups")

    followup = get_object_or_404(
        DealerFollowUp,
        id=followup_id
    )

    if followup.status != "pending":

        messages.warning(
            request,
            "This follow-up is already closed."
        )

        return redirect("dealer_followups")

    followup.status = "completed"
    followup.completed_at = timezone.now()

    followup.save(
        update_fields=["status", "completed_at"]
    )

    messages.success(
        request,
        f"Follow-up completed for {followup.dealer.dealer_name}."
    )

    return redirect("dealer_followups")


# ==================================================
# 12. RESCHEDULE FOLLOW-UP
# ==================================================

@login_required
def reschedule_followup(request, followup_id):

    followup = get_object_or_404(
        DealerFollowUp,
        id=followup_id
    )

    if followup.status != "pending":

        messages.error(
            request,
            "Only pending follow-ups can be rescheduled."
        )

        return redirect("dealer_followups")

    if request.method == "POST":

        date_value = request.POST.get(
            "followup_date", ""
        ).strip()

        time_value = request.POST.get(
            "followup_time", ""
        ).strip()

        remarks = request.POST.get(
            "remarks", ""
        ).strip()

        try:

            selected_date = datetime.strptime(
                date_value,
                "%Y-%m-%d"
            ).date()

            selected_time = datetime.strptime(
                time_value,
                "%H:%M"
            ).time()

        except ValueError:

            messages.error(
                request,
                "Please enter a valid date and time."
            )

            return redirect(
                "reschedule_followup",
                followup_id=followup.id
            )

        selected_datetime = timezone.make_aware(
            datetime.combine(
                selected_date,
                selected_time
            ),
            timezone.get_current_timezone()
        )

        if selected_datetime < timezone.now():

            messages.error(
                request,
                "Follow-up cannot be scheduled in the past."
            )

            return redirect(
                "reschedule_followup",
                followup_id=followup.id
            )

        followup.followup_date = selected_date
        followup.followup_time = selected_time
        followup.remarks = remarks

        followup.save(
            update_fields=[
                "followup_date",
                "followup_time",
                "remarks"
            ]
        )

        messages.success(
            request,
            "Follow-up rescheduled successfully."
        )

        return redirect("dealer_followups")

    return render(
        request,
        "reschedule_followup.html",
        {
            "followup": followup,
            "today": timezone.localdate(),
        }
    )


# ==================================================
# 13. CANCEL FOLLOW-UP
# ==================================================

@login_required
def cancel_followup(request, followup_id):

    if request.method != "POST":
        return redirect("dealer_followups")

    followup = get_object_or_404(
        DealerFollowUp,
        id=followup_id
    )

    if followup.status == "pending":

        followup.status = "cancelled"

        followup.save(
            update_fields=["status"]
        )

        messages.success(
            request,
            "Follow-up cancelled successfully."
        )

    else:

        messages.warning(
            request,
            "Only pending follow-ups can be cancelled."
        )

    return redirect("dealer_followups")


# ==================================================
# 14. DEALER FOLLOW-UP HISTORY
# ==================================================

@login_required
def dealer_followup_history(request, dealer_id):

    dealer = get_object_or_404(
        DealerCall,
        id=dealer_id
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
        }
    )

@login_required
def edit_dealer(request, dealer_id):
    dealer = get_object_or_404(DealerCall, id=dealer_id)

    if request.method == "POST":
        dealer.dealer_name = request.POST.get("dealer_name", "").strip()
        dealer.dealer_code = request.POST.get("dealer_code", "").strip()
        dealer.phone = request.POST.get("phone", "").strip()
        dealer.city = request.POST.get("city", "").strip()
        dealer.pod_number = request.POST.get("pod_number", "").strip()

        dealer.save()

        messages.success(request, "Dealer details updated successfully!")
        return redirect("dashboard")

    return render(request, "edit_dealer.html", {"dealer": dealer})


@login_required
def delete_dealer(request, dealer_id):
    if request.method == "POST":
        dealer = get_object_or_404(DealerCall, id=dealer_id)
        dealer.delete()

        messages.success(request, "Dealer deleted successfully!")

    return redirect("dashboard")



@login_required
def logout_user(request):
    from django.contrib.auth import logout

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
                "Registration successful. Welcome!"
            )

            return redirect("dashboard")
    else:
        form = UserCreationForm()

    return render(
        request,
        "register.html",
        {"form": form}
    )


