from django.urls import include, path

urlpatterns = [
    path("", include("merger.urls")),
]
