

from django.db import models
from django.utils import timezone




class DealerCall(models.Model):
    is_active = models.BooleanField(default=True, db_index=True)

    STATUS_CHOICES = [
        ("pending", "Pending"),
        ("connected", "Connected"),
        ("no_answer", "No Answer"),
        ("busy", "Busy"),
        ("wrong_number", "Wrong Number"),
        ("completed", "Completed"),
        ("unclassified", "Blank / Unclassified"),
]

    # Excel columns 1â€“19: Dealer and scheme information
    sr_no = models.CharField(max_length=50, blank=True, default="")
    scheme_period = models.CharField(max_length=100, blank=True, default="")
    city = models.CharField(max_length=100, blank=True, default="")
    location_code = models.CharField(max_length=100, blank=True, default="")
    location_name = models.CharField(max_length=150, blank=True, default="")
    document_no = models.CharField(max_length=100, blank=True, default="")
    posting_date = models.DateField(null=True, blank=True)

    dealer_code = models.CharField(max_length=80, blank=True, default="")
    dealer_name = models.CharField(max_length=150, blank=True, default="")
    sales_in_charge = models.CharField(max_length=150, blank=True, default="")
    qty = models.CharField(max_length=50, blank=True, default="")
    territory = models.CharField(max_length=150, blank=True, default="")
    address = models.TextField(blank=True, default="")
    phone = models.CharField(max_length=30, blank=True, default="")
    email = models.EmailField(max_length=254, blank=True, default="")
    contact_name = models.CharField(max_length=150, blank=True, default="")
    recipient_relationship = models.CharField(
        max_length=150, blank=True, default=""
    )

    pod_received = models.CharField(max_length=100, blank=True, default="")
    audited_by_cofiser = models.CharField(
        max_length=100, blank=True, default=""
    )

    # Existing compatibility field
    pod_number = models.CharField(max_length=100, blank=True, default="")

    # Excel columns 20â€“24: Calling information
    call_status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default="pending",
        db_index=True
    )
    excel_calling_status = models.CharField(
    max_length=100,
    blank=True,
    default=""
    )

    call_1_remarks = models.TextField(blank=True, default="")
    coin_receive_remark = models.TextField(blank=True, default="")
    call_2_remarks = models.TextField(blank=True, default="")
    final_call_remarks = models.TextField(blank=True, default="")

    # Existing compatibility fields
    remarks = models.TextField(blank=True, default="")
    called_at = models.DateTimeField(null=True, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-updated_at"]
        verbose_name = "Dealer Call"
        verbose_name_plural = "Dealer Calls"

    def __str__(self):
        return f"{self.dealer_name} - {self.call_status}"



class DealerFollowUp(models.Model):

    STATUS_CHOICES = [
        ("pending", "Pending"),
        ("completed", "Completed"),
        ("cancelled", "Cancelled"),
    ]

    dealer = models.ForeignKey(
        DealerCall,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="followups"
    )

    dealer_name = models.CharField(
        max_length=255,
        blank=True,
        default=""
    )

    contact_number = models.CharField(
        max_length=20,
        blank=True,
        default=""
    )

    calling_status = models.CharField(
        max_length=100,
        blank=True,
        default=""
    )

    followup_date = models.DateField(
        null=True,
        blank=True,
        db_index=True
    )

    followup_time = models.TimeField(
        null=True,
        blank=True
    )

    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default="pending",
        db_index=True
    )

    remarks = models.TextField(
        blank=True,
        default=""
    )

    completed_at = models.DateTimeField(
        null=True,
        blank=True
    )

    # Retained for existing compatibility
    follow_up_date = models.DateField(
        null=True,
        blank=True
    )

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["followup_date", "followup_time"]
        verbose_name = "Dealer Follow Up"
        verbose_name_plural = "Dealer Follow Ups"

    def __str__(self):
        return (
            self.dealer.dealer_name
            if self.dealer
            else self.dealer_name
        )




# KEEP YOUR EXISTING DealerFollowUp, DealerCallAttempt
# AND PODTracking CLASSES BELOW THIS LINE WITHOUT CHANGES.

class PODTracking(models.Model):
    is_active = models.BooleanField(default=True, db_index=True)

    sr_no = models.CharField(max_length=50, blank=True, default="")
    scheme_period = models.CharField(max_length=100, blank=True, default="")
    location_code = models.CharField(max_length=100, blank=True, default="")
    document_no = models.CharField(max_length=100)
    posting_date = models.DateField(null=True, blank=True)
    month = models.CharField(max_length=50, blank=True, default="")

    dealer_code = models.CharField(max_length=80)
    group_code = models.CharField(max_length=100, blank=True, default="")
    dealer_name = models.CharField(max_length=150, blank=True, default="")
    location_name = models.CharField(max_length=150, blank=True, default="")
    sales_in_charge = models.CharField(max_length=150, blank=True, default="")
    quarterly_wise = models.CharField(max_length=100, blank=True, default="")

    qty = models.CharField(max_length=50, blank=True, default="")
    received_from_yokohama = models.CharField(max_length=100, blank=True, default="")
    audited_by_cofiser = models.CharField(max_length=100, blank=True, default="")
    sign = models.CharField(max_length=100, blank=True, default="")
    stamp = models.CharField(max_length=100, blank=True, default="")

    date = models.DateField(null=True, blank=True)
    remarks = models.TextField(blank=True, default="")
    depots_remark = models.TextField(blank=True, default="")

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)


    @staticmethod
    def is_marked(value):
        return str(value or "").strip().lower() in {
            "yes", "y", "true", "1", "done",
            "complete", "completed", "signed", "stamped"
        }

    @property
    def is_signed(self):
        return self.is_marked(self.sign)

    @property
    def is_stamped(self):
        return self.is_marked(self.stamp)

    @property
    def is_dated(self):
        return self.date is not None

    @property
    def document_complete(self):
        return (
            self.is_marked(self.received_from_yokohama)
            and self.is_marked(self.audited_by_cofiser)
            and self.is_signed
            and self.is_stamped
            and self.is_dated
    )

    @property
    def document_status(self):
        return (
            "Complete"
            if self.document_complete
            else "Pending"
        )


    class Meta:
        ordering = ["-updated_at"]
        constraints = [
            models.UniqueConstraint(
                fields=["dealer_code", "document_no"],
                name="unique_pod_dealer_document"
            )
        ]

    def __str__(self):
        return f"{self.dealer_code} - {self.document_no}"



class DealerCallAttempt(models.Model):

    STATUS_CHOICES = [
        ("pending", "Pending"),
        ("connected", "Connected"),
        ("no_answer", "No Answer"),
        ("busy", "Busy"),
        ("wrong_number", "Wrong Number"),
        ("completed", "Completed"),
    ]

    dealer = models.ForeignKey(
        DealerCall,
        on_delete=models.CASCADE,
        related_name="call_attempts"
    )

    call_status = models.CharField(
        max_length=30,
        choices=STATUS_CHOICES
    )

    remarks = models.TextField(
        blank=True,
        default=""
    )

    called_at = models.DateTimeField(
        db_index=True,
        default=timezone.now
    )

    class Meta:
        verbose_name = "Dealer Call Attempt"
        verbose_name_plural = "Dealer Call Attempts"
        ordering = ["-called_at"]

    def __str__(self):
        return f"{self.dealer.dealer_name} - {self.call_status}"


class ExcelImportLog(models.Model):

    STATUS_CHOICES = [
        ("pending", "Pending Import"),
        ("imported", "Imported"),
        ("error", "Import Error"),
    ]

    filename = models.CharField(max_length=255)

    file_path = models.CharField(
        max_length=1000,
        unique=True
    )

    file_hash = models.CharField(
        max_length=64,
        blank=True,
        default=""
    )

    file_size = models.BigIntegerField(default=0)

    file_modified_at = models.DateTimeField(
        null=True,
        blank=True
    )

    last_synced_at = models.DateTimeField(
        null=True,
        blank=True
    )

    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default="pending",
        db_index=True
    )

    imported_calls = models.PositiveIntegerField(default=0)
    updated_calls = models.PositiveIntegerField(default=0)
    imported_pods = models.PositiveIntegerField(default=0)
    updated_pods = models.PositiveIntegerField(default=0)

    error_message = models.TextField(blank=True, default="")

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-updated_at"]
        verbose_name = "Excel Import Log"
        verbose_name_plural = "Excel Import Logs"

    def __str__(self):
        return f"{self.filename} - {self.status}"

