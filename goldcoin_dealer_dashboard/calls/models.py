
from django.db import models


class DealerCall(models.Model):

    STATUS_CHOICES = [
        ("pending", "Pending"),
        ("connected", "Connected"),
        ("no_answer", "No Answer"),
        ("busy", "Busy"),
        ("wrong_number", "Wrong Number"),
        ("completed", "Completed"),
    ]

    dealer_name = models.CharField(max_length=150)

    dealer_code = models.CharField(
        max_length=80,
        blank=True,
        default=""
    )

    # Keep phone numbers as strings to preserve leading zeros
    phone = models.CharField(
    max_length=20,
    blank=True,
    default=""
)

    city = models.CharField(
        max_length=100,
        blank=True,
        default=""
    )

    pod_number = models.CharField(
        max_length=100,
        blank=True,
        default=""
    )

    call_status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default="pending",
        db_index=True
    )

    remarks = models.TextField(
        blank=True,
        default=""
    )

    called_at = models.DateTimeField(
        null=True,
        blank=True
    )

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

    # Dealer information
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

    # Follow-up scheduling
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

    # Original field retained for compatibility
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
