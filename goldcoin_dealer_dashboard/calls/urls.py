
from django.urls import path
from django.contrib.auth import views as auth_views
from . import views

urlpatterns = [

    # ==========================================
    # MAIN DASHBOARD
    # ==========================================
    path("", views.dashboard, name="dashboard"),

    # ==========================================
    # DEALER CALL MANAGEMENT
    # ==========================================
    path(
        "dealer/<int:pk>/update/",
        views.update_call,
        name="update_call"
    ),

    path(
        "dealer/edit/<int:dealer_id>/",
        views.edit_dealer,
        name="edit_dealer"
    ),

    path(
        "dealer/delete/<int:dealer_id>/",
        views.delete_dealer,
        name="delete_dealer"
    ),

    # ==========================================
    # EXCEL IMPORT / EXPORT
    # ==========================================
    path(
        "import/",
        views.import_excel,
        name="import_excel"
    ),

    path(
        "export/",
        views.export_excel,
        name="export_excel"
    ),

    path(
        "export/calling-summary/",
        views.export_calling_summary,
        name="export_calling_summary"
    ),

    # ==========================================
    # DEALER FOLLOW-UP DASHBOARD
    # ==========================================
    path(
        "followups/",
        views.dealer_followups,
        name="dealer_followups"
    ),

    # ==========================================
    # FOLLOW-UP MANAGEMENT
    # ==========================================
    path(
        "followup/schedule/<int:dealer_id>/",
        views.schedule_followup,
        name="schedule_followup"
    ),

    path(
        "followup/complete/<int:followup_id>/",
        views.complete_followup,
        name="complete_followup"
    ),

    path(
        "followup/reschedule/<int:followup_id>/",
        views.reschedule_followup,
        name="reschedule_followup"
    ),

    path(
        "followup/cancel/<int:followup_id>/",
        views.cancel_followup,
        name="cancel_followup"
    ),

    path(
        "followup/history/<int:dealer_id>/",
        views.dealer_followup_history,
        name="dealer_followup_history"
    ),

    # ==========================================
    # AUTHENTICATION
    # ==========================================
    path(
        "register/",
        views.register,
        name="register"
    ),

    path(
        "login/",
        auth_views.LoginView.as_view(
            template_name="login.html"
        ),
        name="login"
    ),

    path(
        "logout/",
        views.logout_user,
        name="logout"
    ),

    # ==========================================
    # POD TRACKING
    # ==========================================
    path(
        "pod-tracking/",
        views.pod_tracking,
        name="pod_tracking"
    ),

    path(
        "pod-tracking/import/",
        views.import_pod_excel,
        name="import_pod_excel"
    ),

    path(
        "pod-tracking/update/<int:pk>/",
        views.update_pod,
        name="update_pod"
    ),

    path(
        "pod-tracking/export/",
        views.export_pod_excel,
        name="export_pod_excel"
    ),
]
