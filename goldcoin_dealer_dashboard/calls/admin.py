from django.contrib import admin
from .models import DealerCall
from .models import PODTracking


@admin.register(DealerCall)
class DealerCallAdmin(admin.ModelAdmin):
    list_display = ("dealer_name", "dealer_code", "phone", "city", "pod_number", "call_status", "called_at")
    list_filter = ("call_status", "city")
    search_fields = ("dealer_name", "dealer_code", "phone", "pod_number")

    list_per_page = 50

@admin.register(PODTracking)
class PODTrackingAdmin(admin.ModelAdmin):
    list_display = (
        "sr_no",
        "dealer_code",
        "dealer_name",
        "document_no",
        "scheme_period",
        "received_from_yokohama",
        "audited_by_cofiser",
    )

    search_fields = (
        "dealer_code",
        "dealer_name",
        "document_no",
        "location_name",
    )

    list_filter = (
        "scheme_period",
        "month",
        "location_code",
    )

