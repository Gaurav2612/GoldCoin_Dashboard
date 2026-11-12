
from django.urls import path
from . import views
from django.contrib.auth import views as auth_views

urlpatterns = [

    # =================================
    # MAIN DASHBOARD
    # =================================

    path(
        "",
        views.dashboard,
        name="dashboard"
    ),

    # =================================
    # DEALER CALL MANAGEMENT
    # =================================

    path(
        "dealer/<int:pk>/update/",
        views.update_call,
        name="update_call"
    ),

    # =================================
    # EXCEL IMPORT / EXPORT
    # =================================

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

    # =================================
    # DEALER FOLLOW-UP DASHBOARD
    # =================================

    path(
        "followups/",
        views.dealer_followups,
        name="dealer_followups"
    ),

    # =================================
    # SCHEDULE FOLLOW-UP
    # =================================

    path(
        "followup/schedule/<int:dealer_id>/",
        views.schedule_followup,
        name="schedule_followup"
    ),

    # =================================
    # COMPLETE FOLLOW-UP
    # =================================

    path(
        "followup/complete/<int:followup_id>/",
        views.complete_followup,
        name="complete_followup"
    ),

    # =================================
    # RESCHEDULE FOLLOW-UP
    # =================================

    path(
        "followup/reschedule/<int:followup_id>/",
        views.reschedule_followup,
        name="reschedule_followup"
    ),

    # =================================
    # CANCEL FOLLOW-UP
    # =================================

    path(
        "followup/cancel/<int:followup_id>/",
        views.cancel_followup,
        name="cancel_followup"
    ),

    # =================================
    # DEALER FOLLOW-UP HISTORY
    # =================================

    path(
        "followup/history/<int:dealer_id>/",
        views.dealer_followup_history,
        name="dealer_followup_history"
    ),

    path(
        "dealer/edit/<int:dealer_id>/",
        views.edit_dealer,
        name="edit_dealer"),


    path(
        "dealer/delete/<int:dealer_id>/",
        views.delete_dealer,
         name="delete_dealer"),

        
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


]
