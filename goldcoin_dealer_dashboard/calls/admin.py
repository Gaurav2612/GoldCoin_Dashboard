from django.contrib import admin
from .models import DealerCall
@admin.register(DealerCall)
class DealerCallAdmin(admin.ModelAdmin):
    list_display = ("dealer_name", "dealer_code", "phone", "city", "pod_number", "call_status", "called_at")
    list_filter = ("call_status", "city")
    search_fields = ("dealer_name", "dealer_code", "phone", "pod_number")

    list_per_page = 50
