from django.urls import path
from . import views

urlpatterns = [
    path("", views.index, name="index"),
    # Los 3 listados (reportes) del Certamen 2 + el nuevo Reporte 4
    path("listado1", views.listado1, name="listado1"),
    path("listado2", views.listado2, name="listado2"),
    path("listado3", views.listado3, name="listado3"),
    path("listado4", views.listado4, name="listado4"),

    # Rutas de descarga en PDF y Excel para los 4 reportes
    path("listado1/pdf", views.listado1_pdf, name="listado1_pdf"),
    path("listado1/excel", views.listado1_excel, name="listado1_excel"),
    path("listado2/pdf", views.listado2_pdf, name="listado2_pdf"),
    path("listado2/excel", views.listado2_excel, name="listado2_excel"),
    path("listado3/pdf", views.listado3_pdf, name="listado3_pdf"),
    path("listado3/excel", views.listado3_excel, name="listado3_excel"),
    path("listado4/pdf", views.listado4_pdf, name="listado4_pdf"),
    path("listado4/excel", views.listado4_excel, name="listado4_excel"),

    # Los 3 gráficos del Certamen 2 + el nuevo Gráfico 4 dinámico
    path("grafico1", views.grafico1, name="grafico1"),
    path("grafico2", views.grafico2, name="grafico2"),
    path("grafico3", views.grafico3, name="grafico3"),
    path("grafico4", views.grafico4, name="grafico4"),
]