from django.urls import path

from . import views

app_name = "merger"

urlpatterns = [
    path("", views.home, name="home"),
    path("merge/", views.merge, name="merge"),
    path("download/<uuid:token>/", views.download, name="download"),
    path("discard/<uuid:token>/", views.discard, name="discard"),
]
