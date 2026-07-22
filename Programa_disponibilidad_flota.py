import json
import os
import smtplib
from email.message import EmailMessage
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import matplotlib.pyplot as plt
import pandas as pd
import requests
from matplotlib.patches import FancyBboxPatch
from openpyxl.styles import Alignment, Font, PatternFill


API_KEY = os.environ["DRIVIN_API_KEY"]

URL_VEHICULOS = "https://external.driv.in/api/external/v2/vehicles"
URL_EVENTOS = (
    "https://external.driv.in/api/external/v2/"
    "schedulable_events/events_abastible"
)

ARCHIVO_GRAFICO = "status_flota_dependencia.png"
ARCHIVO_EXCEL = "eventos_abiertos.xlsx"

CORREO_ORIGEN = os.environ["CORREO_ORIGEN"]
CORREO_DESTINO = os.environ["CORREO_DESTINO"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]

HEADERS = {
    "x-api-key": API_KEY,
    "Accept": "application/json",
}


def consultar(url, params=None):
    respuesta = requests.get(
        url,
        headers=HEADERS,
        params=params,
        timeout=90,
    )
    respuesta.raise_for_status()
    datos = respuesta.json()
    return datos["response"] if isinstance(datos, dict) else datos


def dependencia(fleets):
    nombre = str(fleets or "").split(",")[0].strip()
    return nombre[6:].strip() if nombre.lower().startswith("flota ") else nombre


def tarjeta(ax, titulo, valor, subtitulo):
    ax.axis("off")
    ax.add_patch(
        FancyBboxPatch(
            (0.02, 0.06),
            0.96,
            0.88,
            boxstyle="round,pad=0.02,rounding_size=0.04",
            transform=ax.transAxes,
            facecolor="#F7F8FA",
            edgecolor="#D9DDE3",
            linewidth=1.2,
        )
    )
    ax.text(
        0.5, 0.68, titulo,
        ha="center", va="center",
        transform=ax.transAxes,
        fontsize=11, color="#5B616B",
    )
    ax.text(
        0.5, 0.39, valor,
        ha="center", va="center",
        transform=ax.transAxes,
        fontsize=27, color="#086B82",
        fontweight="bold",
    )
    ax.text(
        0.5, 0.16, subtitulo,
        ha="center", va="center",
        transform=ax.transAxes,
        fontsize=8.5, color="#8A9099",
    )


def generar_grafico():
    df = pd.DataFrame(consultar(URL_VEHICULOS))[
        ["code", "fleets", "is_active"]
    ].copy()

    df["code"] = df["code"].fillna("").astype(str).str.strip()
    df["fleets"] = df["fleets"].fillna("").astype(str).str.strip()

    df = df[
        ~df["code"].str.upper().str.contains("FI", na=False)
        & df["code"].str.len().le(6)
        & ~df["fleets"].str.lower().str.contains("emergencia", na=False)
    ].copy()

    df["Dependencia"] = df["fleets"].apply(dependencia)
    df["Estado"] = df["is_active"].map(
        {True: "Disponible", False: "Indisponible"}
    )
    df = df.dropna(subset=["Estado"])

    resumen = (
        df.groupby(["Dependencia", "Estado"])
        .size()
        .unstack(fill_value=0)
        .reindex(columns=["Indisponible", "Disponible"], fill_value=0)
    )
    resumen["Total"] = resumen.sum(axis=1)
    resumen = resumen.sort_values("Total")

    disponibles = int(resumen["Disponible"].sum())
    indisponibles = int(resumen["Indisponible"].sum())
    total = disponibles + indisponibles
    porcentaje = indisponibles / total * 100 if total else 0

    alto = max(8, 3.2 + len(resumen) * 0.42)
    fig = plt.figure(
        figsize=(13, alto),
        facecolor="white",
        layout="constrained",
    )
    grid = fig.add_gridspec(
        2, 3,
        height_ratios=[1.35, 6],
        hspace=0.12,
        wspace=0.08,
    )

    tarjetas = [fig.add_subplot(grid[0, i]) for i in range(3)]
    ax = fig.add_subplot(grid[1, :])

    tarjeta(
        tarjetas[0],
        "Flota actual",
        str(total),
        "Vehículos considerados",
    )
    tarjeta(
        tarjetas[1],
        "Flota indisponible",
        str(indisponibles),
        "Suma de todos los centros",
    )
    tarjeta(
        tarjetas[2],
        "% indisponibilidad",
        f"{porcentaje:.2f}%".replace(".", ","),
        "Indisponibles / flota actual",
    )

    y = range(len(resumen))
    barras_ind = ax.barh(
        y,
        resumen["Indisponible"],
        color="#08788D",
        height=0.68,
    )
    barras_disp = ax.barh(
        y,
        resumen["Disponible"],
        left=resumen["Indisponible"],
        color="#931B62",
        height=0.68,
    )

    ax.set_yticks(list(y), resumen.index)
    ax.set_title(
        "Estado de la flota por dependencia",
        loc="left",
        fontsize=17,
        fontweight="bold",
        pad=24,
    )
    ax.text(
        0,
        1.015,
        "Vehículos activos e inactivos registrados en Drivin",
        transform=ax.transAxes,
        fontsize=9.5,
        color="#757B85",
    )
    ax.legend(
        [barras_disp, barras_ind],
        ["Disponible", "Indisponible"],
        loc="upper right",
        bbox_to_anchor=(1, 1.11),
        frameon=False,
        ncol=2,
    )

    ax.xaxis.grid(alpha=0.16)
    ax.set_axisbelow(True)
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.spines["bottom"].set_color("#DADDE2")
    ax.tick_params(axis="y", length=0, pad=8)
    ax.tick_params(axis="x", colors="#777D86")
    ax.set_xlabel("Cantidad de vehículos", color="#686E77")

    for posicion, fila in enumerate(resumen.itertuples()):
        ind = int(fila.Indisponible)
        disp = int(fila.Disponible)
        total_centro = int(fila.Total)

        if ind:
            ax.text(
                ind / 2,
                posicion,
                str(ind),
                ha="center",
                va="center",
                color="white",
                fontweight="bold",
            )
        if disp:
            ax.text(
                ind + disp / 2,
                posicion,
                str(disp),
                ha="center",
                va="center",
                color="white",
                fontweight="bold",
            )

        ax.text(
            total_centro + 0.25,
            posicion,
            f"Total {total_centro}",
            va="center",
            color="#5E646D",
            fontsize=8.5,
        )

    maximo = max(int(resumen["Total"].max()), 1)
    ax.set_xlim(0, maximo + max(3, int(maximo * 0.13)))

    fig.savefig(
        ARCHIVO_GRAFICO,
        dpi=200,
        facecolor="white",
        pad_inches=0.18,
    )
    plt.close(fig)


def es_vacio(serie):
    texto = serie.astype(str).str.strip().str.lower()
    return serie.isna() | texto.isin({"", "none", "nan", "nat", "null"})


def ajustar_excel(hoja):
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    relleno = PatternFill("solid", fgColor="0B7185")

    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = relleno
        celda.alignment = Alignment(horizontal="center")

    for columna in hoja.columns:
        letra = columna[0].column_letter
        ancho = max(len(str(c.value or "")) for c in columna) + 2
        hoja.column_dimensions[letra].width = min(max(ancho, 12), 45)


def generar_excel(fecha):
    eventos = consultar(
        URL_EVENTOS,
        {
            "from_datetime": fecha,
            "is_unfinished_event": 1,
            "type": "vehicle",
        },
    )

    df = pd.json_normalize(eventos, sep=".")

    df = df[
        ~es_vacio(df["start_date"])
        & es_vacio(df["end_date"])
    ].copy()


    columnas_eliminar = [
        "vehicle_detail",
        "listado_dependiente_clase_de_detencion_1",
        "listado_dependiente_clase_de_detencion_2",
        "numero_de_sap",
        "flota_original",
        "informacion_adicional",
        "estimated_start_date",
        "end_date",
        "created_by_user_at",
        "closed_by",
        "causa_raiz",
        "listado_dependiente_causa_raiz_1",
        "motivos_de_cierre",
        "listado_dependiente_motivos_de_cierre_1",
    ]

    df = df.drop(
        columns=columnas_eliminar,
        errors="ignore",
    )


    nuevos_nombres = {
        "correlative": "Correlativo",
        "vehicle_code": "Código vehículo",
        "name": "Nombre",
        "description": "Descripción",
        "fleets": "Flota",
        "start_date": "Fecha de inicio",
        "observacion": "Observación",
        "clase_de_detencion": "Clase de detención",
        "estimated_end_date": "Fecha estimada de fin",
        "days_duration": "Duración (días)",
        "created_by_user": "Creado por",
        "opened_by": "Abierto por",
        "naturaleza_detencion": "Naturaleza de la detención",
    }

    df = df.rename(columns=nuevos_nombres)

    cantidad = len(df)

    for columna in df.columns:
        df[columna] = df[columna].apply(
            lambda valor: (
                json.dumps(valor, ensure_ascii=False, default=str)
                if isinstance(valor, (list, dict))
                else valor
            )
        )

    if df.empty:
        df = pd.DataFrame(
            {"Mensaje": ["No se encontraron eventos abiertos."]}
        )

    resumen = pd.DataFrame(
        {
            "Dato": [
                "Fecha consultada",
                "Eventos abiertos",
                "Criterio",
            ],
            "Valor": [
                fecha,
                cantidad,
                "start_date con valor y end_date vacío",
            ],
        }
    )

    with pd.ExcelWriter(ARCHIVO_EXCEL, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            sheet_name="Eventos abiertos",
            index=False,
        )
        resumen.to_excel(
            writer,
            sheet_name="Resumen",
            index=False,
        )
        ajustar_excel(writer.sheets["Eventos abiertos"])
        ajustar_excel(writer.sheets["Resumen"])




def enviar_correo(fecha_hora):
    mensaje = EmailMessage()

    mensaje["From"] = CORREO_ORIGEN
    mensaje["To"] = CORREO_DESTINO
    mensaje["Subject"] = "Reporte indisponibilidad flota granel"

    texto = (
        "Estimados,\n\n"
        "Se adjunta reporte de indisponibilidad flota granel "
        f"con última actualización al {fecha_hora}.\n\n"
        "Saludos."
    )

    mensaje.set_content(texto)

    mensaje.add_alternative(
        f"""
        <html>
            <body style="font-family: Arial, sans-serif; color: #333333;">
                <p>Estimados,</p>

                <p>
                    Se adjunta reporte de indisponibilidad flota granel
                    con última actualización al
                    <strong>{fecha_hora}</strong>.
                </p>

                <p>Saludos.</p>
            </body>
        </html>
        """,
        subtype="html",
    )

    
    with open(ARCHIVO_GRAFICO, "rb") as archivo:
        mensaje.add_attachment(
            archivo.read(),
            maintype="image",
            subtype="png",
            filename="status_flota_dependencia.png",
        )

    
    with open(ARCHIVO_EXCEL, "rb") as archivo:
        mensaje.add_attachment(
            archivo.read(),
            maintype="application",
            subtype=(
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            filename="eventos_abiertos.xlsx",
        )

    with smtplib.SMTP_SSL(
        "smtp.gmail.com",
        465,
    ) as servidor:

        servidor.login(
            CORREO_ORIGEN,
            GMAIL_APP_PASSWORD,
        )

        servidor.send_message(mensaje)

    print(f"Correo enviado a {CORREO_DESTINO}")


def main():
    ahora = datetime.now(ZoneInfo("America/Santiago"))
    fecha_api = ahora.strftime("%Y-%m-%d")
    fecha_correo = ahora.strftime("%d-%m-%Y %H:%M")

    generar_grafico()
    generar_excel(fecha_api)
    enviar_correo(fecha_correo)

    print(
        f"Reporte consultado y enviado con fecha {fecha_correo}"
    )


if __name__ == "__main__":
    main()
