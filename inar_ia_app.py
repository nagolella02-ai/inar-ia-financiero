import hashlib
import io
import json
import os
import re
import time
from datetime import datetime
import pandas as pd
import numpy as np
import streamlit as st

try:
    import plotly.graph_objects as go
    PLOTLY_DISPONIBLE = True
except ImportError:
    PLOTLY_DISPONIBLE = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    DOCX_DISPONIBLE = True
except ImportError:
    DOCX_DISPONIBLE = False

try:
    from google import genai
    GEMINI_DISPONIBLE = True
except ImportError:
    GEMINI_DISPONIBLE = False

st.set_page_config(
    page_title="INAR IA — Tablero Financiero Ejecutivo",
    page_icon="📊",
    layout="wide",
)

# ==============================================================================
# ESTILOS CSS CORPORATIVOS (PALETA INAR ASOCIADOS S.A.S.)
# ==============================================================================
st.markdown(
    """
<style>
    .kpi-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        margin-bottom: 12px;
    }
    .kpi-title {
        font-size: 13px;
        font-weight: 600;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 28px;
        font-weight: 700;
        color: #0f172a;
    }
    .kpi-sub {
        font-size: 13px;
        margin-top: 4px;
    }
    .badge-ok {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 600;
        background-color: #dcfce7;
        color: #166534;
    }
    .badge-warn {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 600;
        background-color: #fef3c7;
        color: #92400e;
    }
    .badge-alert {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 600;
        background-color: #fee2e2;
        color: #991b1b;
    }
    .section-banner {
        background: linear-gradient(90deg, #7B9A29 0%, #0F766E 100%);
        color: white;
        padding: 8px 16px;
        border-radius: 8px;
        font-weight: 600;
        margin: 12px 0 14px 0;
    }
</style>
""",
    unsafe_allow_html=True,
)


# ==============================================================================
# 0. GESTIÓN DE API KEY PERSISTENTE Y HISTÓRICO DE INFORMES IA
# ==============================================================================
RUTA_HISTORIAL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "historial_informes_ia.json")


def cargar_api_key_guardada() -> str:
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY", "").strip()


def guardar_api_key_local(api_key: str) -> None:
    os.makedirs(".streamlit", exist_ok=True)
    ruta_secrets = os.path.join(".streamlit", "secrets.toml")
    clave_limpia = api_key.strip().replace('"', "")
    with open(ruta_secrets, "w", encoding="utf-8") as f:
        f.write(f'GEMINI_API_KEY = "{clave_limpia}"\n')


def calcular_hash_datos(df_comp: pd.DataFrame, df_ind: pd.DataFrame, dias_periodo: int) -> str:
    payload = f"{dias_periodo}|{df_comp.to_json()}|{df_ind.to_json()}"
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


def cargar_historial_ia() -> list[dict]:
    if not os.path.exists(RUTA_HISTORIAL):
        return []
    try:
        with open(RUTA_HISTORIAL, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def guardar_en_historial_ia(registro: dict) -> None:
    historial = cargar_historial_ia()
    # Reemplazar si ya existía exactamente el mismo hash_datos para evitar duplicados innecesarios
    historial = [h for h in historial if h.get("hash_datos") != registro.get("hash_datos")]
    historial.insert(0, registro)
    with open(RUTA_HISTORIAL, "w", encoding="utf-8") as f:
        json.dump(historial, f, ensure_ascii=False, indent=2)


def eliminar_del_historial_ia(registro_id: str) -> None:
    historial = [h for h in cargar_historial_ia() if h.get("id") != registro_id]
    with open(RUTA_HISTORIAL, "w", encoding="utf-8") as f:
        json.dump(historial, f, ensure_ascii=False, indent=2)


# ==============================================================================
# 1. LECTURA Y CLASIFICACIÓN POR CÓDIGO PUC (FORMATO SIIGO)
# ==============================================================================
def cargar_balance_siigo(archivo) -> pd.DataFrame:
    df_raw = pd.read_excel(archivo, header=None)
    header_idx = None
    for idx, row in df_raw.iterrows():
        if str(row.iloc[0]).strip().upper() == "GRUPO":
            header_idx = idx
            break

    df = pd.read_excel(archivo, dtype=str) if header_idx is None else pd.read_excel(archivo, header=header_idx, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    for col in ["GRUPO", "CUENTA", "SUBCUENT", "AUXILIAR", "SUBAUXIL", "NIT", "DESCRIPCION"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.replace(r"\.0$", "", regex=True)
            )
            if col in ["CUENTA", "SUBCUENT", "AUXILIAR", "SUBAUXIL"]:
                df[col] = df[col].apply(lambda x: x.zfill(2) if x.isdigit() and len(x) == 1 else x)
        else:
            df[col] = ""

    for num_col in ["SALDO ANTERIOR", "DEBITOS", "CREDITOS", "NUEVO SALDO"]:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors="coerce").fillna(0.0)
        else:
            df[num_col] = 0.0

    return df


def obtener_saldo_puc(
    df: pd.DataFrame,
    grupo: str,
    cuenta: str = "",
    subcuent: str = "",
    auxiliar: str = "",
    col: str = "NUEVO SALDO",
) -> float:
    mask = (
        (df["GRUPO"] == grupo)
        & (df["CUENTA"] == cuenta)
        & (df["SUBCUENT"] == subcuent)
        & (df["AUXILIAR"] == auxiliar)
        & (df["SUBAUXIL"] == "")
        & (df["NIT"] == "")
    )
    filas = df[mask]
    return float(filas[col].iloc[0]) if not filas.empty else 0.0


def clasificar_cuentas(df: pd.DataFrame, col: str = "NUEVO SALDO") -> dict:
    disponible = obtener_saldo_puc(df, "11", col=col)
    disp_debitos = obtener_saldo_puc(df, "11", col="DEBITOS")
    disp_creditos = obtener_saldo_puc(df, "11", col="CREDITOS")
    inversiones = obtener_saldo_puc(df, "12", col=col)
    cxc_totales = obtener_saldo_puc(df, "13", col=col)
    clientes = obtener_saldo_puc(df, "13", "05", col=col)
    cxc_por_facturar = obtener_saldo_puc(df, "13", "05", "05", "70", col=col)
    inventarios = obtener_saldo_puc(df, "14", col=col)
    ppe = obtener_saldo_puc(df, "15", col=col)
    intangibles = obtener_saldo_puc(df, "16", col=col)
    activo_total = obtener_saldo_puc(df, "1", col=col)

    activo_corriente = disponible + cxc_totales + inventarios
    activo_no_corriente = inversiones + ppe + intangibles

    pasivo_total = abs(obtener_saldo_puc(df, "2", col=col))
    ob_financieras = abs(obtener_saldo_puc(df, "21", col=col))
    cxp = abs(obtener_saldo_puc(df, "23", col=col))
    dividendos_lp = abs(obtener_saldo_puc(df, "23", "60", "05", "02", col=col))
    impuestos_pagar = abs(obtener_saldo_puc(df, "24", col=col))
    ob_laborales = abs(obtener_saldo_puc(df, "25", col=col))
    provisiones = -obtener_saldo_puc(df, "26", col=col)

    pasivo_no_corriente = dividendos_lp
    pasivo_corriente = max(pasivo_total - pasivo_no_corriente, 0.0)
    patrimonio_base = abs(obtener_saldo_puc(df, "3", col=col))

    ingresos_totales = -obtener_saldo_puc(df, "4", col=col)
    ingresos_op = -obtener_saldo_puc(df, "41", col=col)
    ingresos_noop = -obtener_saldo_puc(df, "42", col=col)

    gastos_totales = obtener_saldo_puc(df, "5", col=col)
    gastos_adm = obtener_saldo_puc(df, "51", col=col)
    gastos_ventas_op = obtener_saldo_puc(df, "52", col=col)
    gastos_noop = obtener_saldo_puc(df, "53", col=col)
    gastos_financieros = obtener_saldo_puc(df, "53", "05", col=col)
    costos_ventas = obtener_saldo_puc(df, "6", col=col)

    utilidad_bruta = ingresos_op - costos_ventas
    utilidad_operacional = utilidad_bruta - (gastos_adm + gastos_ventas_op)
    utilidad_neta = ingresos_totales - (costos_ventas + gastos_totales)
    patrimonio_total = patrimonio_base + utilidad_neta

    return {
        "disponible": disponible,
        "disp_debitos": disp_debitos,
        "disp_creditos": disp_creditos,
        "inversiones": inversiones,
        "cxc_totales": cxc_totales,
        "clientes": clientes,
        "cxc_por_facturar": cxc_por_facturar,
        "inventarios": inventarios,
        "ppe": ppe,
        "intangibles": intangibles,
        "activo_corriente": activo_corriente,
        "activo_no_corriente": activo_no_corriente,
        "activo_total": activo_total,
        "ob_financieras": ob_financieras,
        "cxp": cxp,
        "dividendos_lp": dividendos_lp,
        "impuestos_pagar": impuestos_pagar,
        "ob_laborales": ob_laborales,
        "provisiones": provisiones,
        "pasivo_corriente": pasivo_corriente,
        "pasivo_no_corriente": pasivo_no_corriente,
        "pasivo_total": pasivo_total,
        "patrimonio_base": patrimonio_base,
        "patrimonio_total": patrimonio_total,
        "ingresos_op": ingresos_op,
        "ingresos_noop": ingresos_noop,
        "ingresos_totales": ingresos_totales,
        "costos_ventas": costos_ventas,
        "gastos_adm": gastos_adm,
        "gastos_ventas_op": gastos_ventas_op,
        "gastos_noop": gastos_noop,
        "gastos_financieros": gastos_financieros,
        "gastos_totales": gastos_totales,
        "utilidad_bruta": utilidad_bruta,
        "utilidad_operacional": utilidad_operacional,
        "utilidad_neta": utilidad_neta,
    }


def calcular_indicadores(d: dict, dias_periodo: int = 180) -> dict:
    ac, pc = d["activo_corriente"], d["pasivo_corriente"]
    at, pt, pat = d["activo_total"], d["pasivo_total"], d["patrimonio_total"]
    inv = d["inventarios"]
    ing_op = d["ingresos_op"]
    cli = d["clientes"]
    cxc = d["cxc_totales"]

    capital_trabajo = ac - pc
    razon_corriente = ac / pc if pc > 0 else 0.0
    prueba_acida = (ac - inv) / pc if pc > 0 else 0.0

    endeudamiento_total = (pt / at) * 100 if at > 0 else 0.0
    endeudamiento_cp = (pc / pt) * 100 if pt > 0 else 0.0
    apalancamiento = pt / pat if pat > 0 else 0.0
    cobertura_intereses = (
        d["utilidad_operacional"] / d["gastos_financieros"]
        if d["gastos_financieros"] > 0
        else 0.0
    )

    margen_bruto = (d["utilidad_bruta"] / ing_op) * 100 if ing_op > 0 else 0.0
    margen_operacional = (d["utilidad_operacional"] / ing_op) * 100 if ing_op > 0 else 0.0
    margen_neto = (d["utilidad_neta"] / ing_op) * 100 if ing_op > 0 else 0.0
    roa = (d["utilidad_neta"] / at) * 100 if at > 0 else 0.0
    roe = (d["utilidad_neta"] / pat) * 100 if pat > 0 else 0.0

    rot_clientes = ing_op / cli if cli > 0 else 0.0
    dias_cobro_clientes = dias_periodo / rot_clientes if rot_clientes > 0 else 0.0
    rot_cxc = ing_op / cxc if cxc > 0 else 0.0
    dias_cobro_cxc = dias_periodo / rot_cxc if rot_cxc > 0 else 0.0
    base_costo_inv = d["costos_ventas"] if d["costos_ventas"] > 0 else d["gastos_ventas_op"]
    rot_inventarios = base_costo_inv / inv if inv > 0 else 0.0
    rot_activos = ing_op / at if at > 0 else 0.0

    return {
        "Capital de Trabajo Neto ($)": capital_trabajo,
        "Razón Corriente (veces)": razon_corriente,
        "Prueba Ácida (veces)": prueba_acida,
        "Endeudamiento Total (%)": endeudamiento_total,
        "Endeudamiento Corto Plazo (%)": endeudamiento_cp,
        "Apalancamiento (Pasivo/Patrimonio)": apalancamiento,
        "Cobertura de Intereses (veces)": cobertura_intereses,
        "Margen Bruto (%)": margen_bruto,
        "Margen Operacional (%)": margen_operacional,
        "Margen Neto (%)": margen_neto,
        "ROA - Rentabilidad del Activo (%)": roa,
        "ROE - Rentabilidad del Patrimonio (%)": roe,
        "Rotación de Clientes (veces)": rot_clientes,
        "Días Promedio de Cobro - Clientes": dias_cobro_clientes,
        "Rotación de CxC Totales (veces)": rot_cxc,
        "Días Promedio de Cobro - CxC Totales": dias_cobro_cxc,
        "Rotación de Inventarios / Obras (veces)": rot_inventarios,
        "Rotación de Activos Totales (veces)": rot_activos,
    }


def generar_alertas_y_diagnostico(d_ant: dict, d_act: dict, ind_act: dict) -> tuple[list, dict]:
    alertas = []

    var_disp = (
        ((d_act["disponible"] - d_ant["disponible"]) / d_ant["disponible"]) * 100
        if d_ant["disponible"] > 0
        else 0.0
    )
    if var_disp < -10:
        alertas.append(
            f"🔴 **Caída de liquidez inmediata:** El Disponible (Grupo 11) disminuyó un **{var_disp:.2f}%** "
            f"(de ${d_ant['disponible']:,.0f} a ${d_act['disponible']:,.0f})."
        )

    pct_cxc_ac = (d_act["cxc_totales"] / d_act["activo_corriente"] * 100) if d_act["activo_corriente"] > 0 else 0.0
    if pct_cxc_ac > 50:
        alertas.append(
            f"🟠 **Alta concentración en cartera:** Las Cuentas por Cobrar representan el **{pct_cxc_ac:.1f}%** "
            f"del Activo Corriente (${d_act['cxc_totales']:,.0f}), incluyendo **${d_act['cxc_por_facturar']:,.0f}** "
            "en CxC pendientes por facturar (cuenta 13050570)."
        )

    if ind_act["Días Promedio de Cobro - Clientes"] > 90:
        alertas.append(
            f"🟠 **Ciclo de cobro lento:** El periodo promedio de cobro a clientes es de "
            f"**{ind_act['Días Promedio de Cobro - Clientes']:.1f} días**, lo que inmoviliza capital de trabajo."
        )

    if d_act["ingresos_op"] > 0 and d_act["utilidad_operacional"] < 0:
        alertas.append(
            f"🔴 **Pérdida operacional en el periodo:** Los gastos operacionales (51 + 52) superan a los ingresos "
            f"operacionales (41), generando un margen operacional de **{ind_act['Margen Operacional (%)']:.2f}%**."
        )

    if d_act["inventarios"] > 0 and abs(d_act["inventarios"] - d_ant["inventarios"]) < 1:
        alertas.append(
            f"🟡 **Partida sin movimiento:** La cuenta 14 (Inventarios / Contratos en Ejecución) permanece "
            f"inmovilizada en **${d_act['inventarios']:,.0f}** sin variación respecto al saldo anterior."
        )

    if not alertas:
        alertas.append("🟢 **Sin alertas críticas:** Los indicadores principales se mantienen dentro de rangos estables.")

    interpretaciones = {
        "liquidez": (
            f"La razón corriente es de **{ind_act['Razón Corriente (veces)']:.2f}x** y la prueba ácida de "
            f"**{ind_act['Prueba Ácida (veces)']:.2f}x**, con un capital de trabajo neto de "
            f"**${ind_act['Capital de Trabajo Neto ($)']:,.0f}**. Sin embargo, el **{pct_cxc_ac:.1f}%** del activo "
            "corriente está concentrado en cuentas por cobrar."
        ),
        "endeudamiento": (
            f"El endeudamiento total se ubica en **{ind_act['Endeudamiento Total (%)']:.2f}%** "
            f"(apalancamiento de **{ind_act['Apalancamiento (Pasivo/Patrimonio)']:.2f}x**). Solo el "
            f"**{ind_act['Endeudamiento Corto Plazo (%)']:.2f}%** de la deuda es exigible a corto plazo, gracias "
            f"a los dividendos por pagar a largo plazo (${d_act['dividendos_lp']:,.0f})."
        ),
        "rentabilidad": (
            f"El margen operacional es de **{ind_act['Margen Operacional (%)']:.2f}%** y el margen neto de "
            f"**{ind_act['Margen Neto (%)']:.2f}%** (ROA: **{ind_act['ROA - Rentabilidad del Activo (%)']:.2f}%**, "
            f"ROE: **{ind_act['ROE - Rentabilidad del Patrimonio (%)']:.2f}%**). Se requiere priorizar la "
            "facturación de obra ejecutada y el control de costos operativos en la cuenta 52."
        ),
        "actividad": (
            f"La cartera comercial rota **{ind_act['Rotación de Clientes (veces)']:.2f} veces** en el periodo "
            f"(**{ind_act['Días Promedio de Cobro - Clientes']:.1f} días** de cobro promedio) y los activos totales "
            f"rotan **{ind_act['Rotación de Activos Totales (veces)']:.2f} veces**."
        ),
    }

    return alertas, interpretaciones


# ==============================================================================
# 2. GRÁFICOS VISUALES TIPO TACÓMETRO Y SPARKLINES (PLOTLY)
# ==============================================================================
def crear_gauge(valor: float, titulo: str, rango_max: float, steps: list, sufijo: str = "") -> go.Figure:
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=valor,
            number={"suffix": sufijo, "font": {"size": 28, "color": "#0f172a"}},
            title={"text": titulo, "font": {"size": 15, "color": "#334155"}},
            gauge={
                "axis": {"range": [None, rango_max], "tickwidth": 1, "tickcolor": "#94a3b8"},
                "bar": {"color": "#1e293b", "thickness": 0.25},
                "bgcolor": "white",
                "borderwidth": 0,
                "steps": steps,
            },
        )
    )
    fig.update_layout(height=230, margin=dict(l=20, r=20, t=45, b=15), paper_bgcolor="rgba(0,0,0,0)")
    return fig


def crear_sparkline(val_ant: float, val_act: float, color_hex: str) -> go.Figure:
    fig = go.Figure(
        go.Scatter(
            x=["Saldo Anterior", "Nuevo Saldo"],
            y=[val_ant, val_act],
            mode="lines+markers",
            fill="tozeroy",
            line=dict(color=color_hex, width=3),
            marker=dict(size=7, color=color_hex),
        )
    )
    fig.update_layout(
        height=110,
        margin=dict(l=5, r=5, t=5, b=5),
        xaxis=dict(showgrid=False, showticklabels=True, tickfont=dict(size=10)),
        yaxis=dict(showgrid=False, showticklabels=False),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
    )
    return fig


# ==============================================================================
# 3. INTEGRACIÓN CON LA API DE GEMINI
# ==============================================================================
def obtener_modelos_disponibles(client) -> list[str]:
    disponibles = []
    try:
        for m in client.models.list():
            nombre = getattr(m, "name", "")
            acciones = getattr(m, "supported_actions", []) or []
            if "gemini" in nombre.lower() and (not acciones or "generateContent" in acciones):
                disponibles.append(nombre.replace("models/", ""))
    except Exception:
        pass
    return disponibles


def generar_interpretacion_gemini(
    api_key: str,
    modelo: str,
    df_comp: pd.DataFrame,
    df_ind: pd.DataFrame,
    alertas: list,
    dias_periodo: int,
) -> tuple[str, str]:
    client = genai.Client(api_key=api_key.strip())
    prompt = f"""
Eres un Director Financiero (CFO) experto en análisis financiero bajo NIIF/PUC en Colombia.
Analiza los siguientes resultados reales de INAR ASOCIADOS S.A.S. (Bogotá, sector ingeniería/construcción)
correspondientes a un corte de {dias_periodo} días.

1. COMPARATIVO DE CUENTAS PRINCIPALES:
{df_comp.to_string(index=False)}

2. INDICADORES FINANCIEROS:
{df_ind.to_string(index=False)}

3. ALERTAS DETECTADAS AUTOMÁTICAMENTE:
{chr(10).join(alertas)}

Redacta un informe ejecutivo estructurado en español con:
- **Diagnóstico General** (máximo 1 párrafo directo).
- **Análisis por Categoría** (Liquidez, Endeudamiento, Rentabilidad y Actividad, explicando el porqué de cada cifra).
- **Plan de Acción Prioritario** (3 recomendaciones concretas y accionables para la gerencia).
Usa exclusivamente las cifras entregadas, sin inventar datos.
"""
    modelos_disponibles = obtener_modelos_disponibles(client)
    candidatos = [modelo] + [m for m in modelos_disponibles if m != modelo and "flash" in m.lower()]
    candidatos += [m for m in modelos_disponibles if m not in candidatos]

    ultimo_error = None
    for m in candidatos[:5]:
        for _ in range(2):
            try:
                response = client.models.generate_content(model=m, contents=prompt)
                return response.text, m
            except Exception as e:
                ultimo_error = e
                msg = str(e)
                if "503" in msg or "UNAVAILABLE" in msg or "429" in msg:
                    time.sleep(2)
                    continue
                if "404" in msg or "NOT_FOUND" in msg:
                    break
                raise e

    raise ultimo_error


# ==============================================================================
# 4. EXPORTACIÓN A EXCEL Y WORD CON HOJA MEMBRETE AUTOMÁTICA
# ==============================================================================
def exportar_excel_informe(
    df_comp: pd.DataFrame,
    df_ind: pd.DataFrame,
    alertas: list,
    interpretaciones: dict,
    analisis_gemini: str = "",
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_comp.to_excel(writer, sheet_name="Comparativo_Cuentas", index=False)
        df_ind.to_excel(writer, sheet_name="Indicadores_Financieros", index=False)
        df_diag = pd.DataFrame(
            {
                "Sección": ["Liquidez", "Endeudamiento", "Rentabilidad", "Actividad"] + ["Alerta"] * len(alertas),
                "Detalle": [
                    interpretaciones["liquidez"].replace("**", ""),
                    interpretaciones["endeudamiento"].replace("**", ""),
                    interpretaciones["rentabilidad"].replace("**", ""),
                    interpretaciones["actividad"].replace("**", ""),
                ]
                + [a.replace("**", "") for a in alertas],
            }
        )
        df_diag.to_excel(writer, sheet_name="Diagnostico_y_Alertas", index=False)
        if analisis_gemini:
            pd.DataFrame({"Informe_Ejecutivo_Gemini": analisis_gemini.splitlines()}).to_excel(
                writer, sheet_name="Analisis_IA_Gemini", index=False
            )
    return output.getvalue()


def _colorear_celda_word(celda, color_hex: str) -> None:
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def _aplicar_bordes_tabla(tabla, color_hex: str = "CBD5E1") -> None:
    tblPr = tabla._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color_hex)
        borders.append(el)
    tblPr.append(borders)


def _agregar_texto_con_negritas(parrafo, texto: str) -> None:
    partes = re.split(r"(\*\*.*?\*\*)", texto)
    for parte in partes:
        if parte.startswith("**") and parte.endswith("**"):
            run = parrafo.add_run(parte[2:-2])
            run.bold = True
        else:
            parrafo.add_run(parte)


def _agregar_bullet(doc, texto: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    r_b = p.add_run("• ")
    r_b.bold = True
    _agregar_texto_con_negritas(p, texto)


def _buscar_membrete_automatico() -> str | None:
    carpeta_script = os.path.dirname(os.path.abspath(__file__))
    candidatos = [
        os.path.join(carpeta_script, "Hoja membrete nar 2026.docx"),
        "Hoja membrete nar 2026.docx",
    ]
    for c in candidatos:
        if os.path.exists(c):
            return c
    for carpeta in (carpeta_script, os.getcwd()):
        try:
            for nombre in os.listdir(carpeta):
                if nombre.lower().endswith(".docx") and "membrete" in nombre.lower():
                    return os.path.join(carpeta, nombre)
        except Exception:
            pass
    return None


def _crear_documento_con_membrete_por_defecto():
    ruta = _buscar_membrete_automatico()
    if ruta:
        return Document(ruta)

    doc = Document()
    sec = doc.sections[0]
    p_head = sec.header.paragraphs[0]
    r_h = p_head.add_run("INAR ASOCIADOS S.A.S.  |  Consultoría e Ingeniería")
    r_h.bold = True
    r_h.font.size = Pt(10)
    r_h.font.color.rgb = RGBColor(0x7B, 0x9A, 0x29)

    p_foot = sec.footer.paragraphs[0]
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f = p_foot.add_run(
        "INAR ASOCIADOS S.A.S. • PBX: (+57) 601 6248006 • Transversal 60 # 115-58 Torre C Of. 603-605 • "
        "Bogotá, Colombia • contactenos@inarasociados.com.co"
    )
    r_f.font.size = Pt(8)
    r_f.font.color.rgb = RGBColor(0x5A, 0x73, 0x1E)
    return doc


def exportar_word_membrete(
    df_ind: pd.DataFrame,
    alertas: list,
    interpretaciones: dict,
    analisis_gemini: str,
    dias_periodo: int,
) -> bytes:
    doc = _crear_documento_con_membrete_por_defecto()

    if len(doc.paragraphs) == 1 and not doc.paragraphs[0].text.strip():
        p_titulo = doc.paragraphs[0]
    else:
        p_titulo = doc.add_paragraph()

    p_titulo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_tit = p_titulo.add_run("INFORME EJECUTIVO DE ANÁLISIS FINANCIERO — INAR IA")
    r_tit.bold = True
    r_tit.font.size = Pt(14)
    r_tit.font.color.rgb = RGBColor(0x5A, 0x73, 0x1E)

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run(f"INAR ASOCIADOS S.A.S. | Corte analizado: {dias_periodo} días")
    r_sub.font.size = Pt(10)
    r_sub.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    h1 = doc.add_paragraph()
    r_h1 = h1.add_run("1. Resumen de Indicadores Financieros")
    r_h1.bold = True
    r_h1.font.size = Pt(12)
    r_h1.font.color.rgb = RGBColor(0x0F, 0x76, 0x6E)

    tabla = doc.add_table(rows=1, cols=4)
    _aplicar_bordes_tabla(tabla)
    encabezados = ["Indicador", "Saldo Anterior", "Nuevo Saldo", "Variación"]
    for i, texto in enumerate(encabezados):
        celda = tabla.rows[0].cells[i]
        _colorear_celda_word(celda, "7B9A29")
        run = celda.paragraphs[0].add_run(texto)
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for _, fila in df_ind.iterrows():
        row_cells = tabla.add_row().cells
        vals = [
            str(fila["Indicador"]),
            f"{fila['Saldo Anterior']:,.2f}",
            f"{fila['Nuevo Saldo']:,.2f}",
            f"{fila['Variación']:,.2f}",
        ]
        for i, val in enumerate(vals):
            run = row_cells[i].paragraphs[0].add_run(val)
            run.font.size = Pt(8.5)

    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(10)
    r_h2 = h2.add_run("2. Alertas Tempranas Detectadas")
    r_h2.bold = True
    r_h2.font.size = Pt(12)
    r_h2.font.color.rgb = RGBColor(0x0F, 0x76, 0x6E)

    for alerta in alertas:
        limpia = re.sub(r"^[🔴🟠🟡🟢]\s*", "", alerta)
        _agregar_bullet(doc, limpia)

    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(10)
    r_h3 = h3.add_run("3. Diagnóstico Ejecutivo e Interpretación")
    r_h3.bold = True
    r_h3.font.size = Pt(12)
    r_h3.font.color.rgb = RGBColor(0x0F, 0x76, 0x6E)

    if analisis_gemini:
        for linea in analisis_gemini.splitlines():
            linea = linea.strip()
            if not linea:
                continue
            if linea.startswith("#"):
                p_sec = doc.add_paragraph()
                r_sec = p_sec.add_run(linea.lstrip("#").strip())
                r_sec.bold = True
                r_sec.font.size = Pt(11)
            elif linea.startswith("* ") or linea.startswith("- "):
                _agregar_bullet(doc, linea[2:].strip())
            else:
                p_txt = doc.add_paragraph()
                _agregar_texto_con_negritas(p_txt, linea)
    else:
        for clave, nombre in [
            ("liquidez", "Liquidez"),
            ("endeudamiento", "Endeudamiento"),
            ("rentabilidad", "Rentabilidad"),
            ("actividad", "Actividad"),
        ]:
            p_loc = doc.add_paragraph()
            _agregar_texto_con_negritas(p_loc, f"**{nombre}:** {interpretaciones[clave]}")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ==============================================================================
# 5. INTERFAZ STREAMLIT — INAR IA
# ==============================================================================
st.title("📊 Tablero Financiero Ejecutivo — INAR IA")
st.caption("INAR ASOCIADOS S.A.S. • Especialización en Gerencia Financiera • Análisis Automatizado PUC / NIIF")

st.sidebar.header("⚙️ Configuración del Análisis")
archivo = st.sidebar.file_uploader("Cargar Balance de Prueba Siigo (.xlsx)", type=["xlsx"])
dias_periodo = st.sidebar.selectbox(
    "Días del periodo analizado (indicadores de actividad)",
    options=[30, 90, 180, 365],
    index=2,
    help="Usa 180 días para cortes semestrales (Ene-Jun), 30 para mensual o 365 para anual.",
)

st.sidebar.markdown("---")
st.sidebar.header("🤖 Conexión API Gemini")
clave_inicial = cargar_api_key_guardada()
gemini_api_key = st.sidebar.text_input(
    "API Key de Google Gemini",
    value=clave_inicial,
    type="password",
    help="Pega tu clave una vez y guárdala para no volver a escribirla.",
)
if gemini_api_key and gemini_api_key != clave_inicial:
    if st.sidebar.button("💾 Anclar API Key en este equipo"):
        guardar_api_key_local(gemini_api_key)
        st.sidebar.success("✅ API Key guardada en .streamlit/secrets.toml")
elif clave_inicial:
    st.sidebar.caption("🔒 API Key anclada y lista para usar.")

modelo_gemini = st.sidebar.text_input(
    "Modelo principal",
    value="gemini-3.8-flash",
    help="Si está saturado o no disponible, INAR IA buscará otro modelo activo automáticamente.",
)

historial_guardado = cargar_historial_ia()
st.sidebar.markdown("---")
st.sidebar.caption(f"🗂️ Informes IA en histórico local: **{len(historial_guardado)}**")

if archivo is not None:
    df_siigo = cargar_balance_siigo(archivo)
    d_ant = clasificar_cuentas(df_siigo, col="SALDO ANTERIOR")
    d_act = clasificar_cuentas(df_siigo, col="NUEVO SALDO")

    ind_ant = calcular_indicadores(d_ant, dias_periodo=dias_periodo)
    ind_act = calcular_indicadores(d_act, dias_periodo=dias_periodo)
    alertas, interpretaciones = generar_alertas_y_diagnostico(d_ant, d_act, ind_act)

    rubros = [
        ("Disponible (11)", "disponible"),
        ("Inversiones (12)", "inversiones"),
        ("Cuentas por Cobrar (13)", "cxc_totales"),
        ("  - Clientes (1305)", "clientes"),
        ("Inventarios / Obras (14)", "inventarios"),
        ("Propiedad, Planta y Equipo (15)", "ppe"),
        ("ACTIVO TOTAL (1)", "activo_total"),
        ("Obligaciones Financieras (21)", "ob_financieras"),
        ("Cuentas por Pagar (23)", "cxp"),
        ("  - Dividendos Largo Plazo (23600502)", "dividendos_lp"),
        ("Impuestos por Pagar (24)", "impuestos_pagar"),
        ("Obligaciones Laborales (25)", "ob_laborales"),
        ("Provisiones (26)", "provisiones"),
        ("PASIVO TOTAL (2)", "pasivo_total"),
        ("PATRIMONIO TOTAL (3 + Resultado)", "patrimonio_total"),
    ]

    filas_comp = []
    for nombre, clave in rubros:
        v_ant, v_act = d_ant[clave], d_act[clave]
        var_abs = v_act - v_ant
        var_pct = (var_abs / v_ant * 100) if v_ant != 0 else 0.0
        pct_activo = (v_act / d_act["activo_total"] * 100) if d_act["activo_total"] > 0 else 0.0
        filas_comp.append(
            {
                "Rubro PUC": nombre,
                "Saldo Anterior ($)": round(v_ant, 2),
                "Nuevo Saldo ($)": round(v_act, 2),
                "Variación ($)": round(var_abs, 2),
                "Variación (%)": round(var_pct, 2),
                "Análisis Vertical (% Activo)": round(pct_activo, 2),
            }
        )
    df_comp = pd.DataFrame(filas_comp)

    filas_ind = [
        {
            "Indicador": k,
            "Saldo Anterior": round(ind_ant[k], 2),
            "Nuevo Saldo": round(ind_act[k], 2),
            "Variación": round(ind_act[k] - ind_ant[k], 2),
        }
        for k in ind_act
    ]
    df_ind = pd.DataFrame(filas_ind)

    hash_actual = calcular_hash_datos(df_comp, df_ind, dias_periodo)
    if st.session_state.get("hash_activo") != hash_actual:
        st.session_state["hash_activo"] = hash_actual
        coincidencia = next((h for h in historial_guardado if h.get("hash_datos") == hash_actual), None)
        if coincidencia:
            st.session_state["analisis_gemini"] = coincidencia["analisis_gemini"]
            st.session_state["modelo_usado"] = coincidencia["modelo_usado"]
            st.session_state["fecha_analisis"] = coincidencia["fecha_hora"]
            st.session_state["desde_cache"] = True
        else:
            st.session_state["analisis_gemini"] = ""
            st.session_state["desde_cache"] = False

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        [
            "🎯 Tablero Ejecutivo KPI",
            "📈 Rentabilidad y Detalle",
            "📊 Comparativo Horizontal / Vertical",
            "🤖 Diagnóstico IA y Exportación (Word / Excel)",
            "🗂️ Histórico de Informes IA",
        ]
    )

    with tab1:
        if not PLOTLY_DISPONIBLE:
            st.warning("Para ver los tacómetros y gráficos ejecutivos instala Plotly: `py -m pip install plotly`")

        st.markdown('<div class="section-banner">💧 Capital de Trabajo y Posición de Caja</div>', unsafe_allow_html=True)
        col_izq, col_der = st.columns([1.1, 1.4])

        with col_izq:
            pa_act = ind_act["Prueba Ácida (veces)"]
            pa_ant = ind_ant["Prueba Ácida (veces)"]
            badge_pa = '<span class="badge-ok">✓ Saludable</span>' if pa_act >= 1 else '<span class="badge-alert">! Atención</span>'
            st.markdown(
                f"""<div class="kpi-card">
                <div class="kpi-title">Prueba Ácida (Quick Ratio) &nbsp; {badge_pa}</div>
                <div class="kpi-value">{pa_act:.2f}x</div>
                <div class="kpi-sub">Anterior: {pa_ant:.2f}x</div>
                </div>""",
                unsafe_allow_html=True,
            )
            if PLOTLY_DISPONIBLE:
                st.plotly_chart(crear_sparkline(pa_ant, pa_act, "#0F766E"), use_container_width=True)

            rc_act = ind_act["Razón Corriente (veces)"]
            rc_ant = ind_ant["Razón Corriente (veces)"]
            badge_rc = '<span class="badge-ok">✓ Saludable</span>' if rc_act >= 1.5 else '<span class="badge-warn">! Revisar</span>'
            st.markdown(
                f"""<div class="kpi-card">
                <div class="kpi-title">Razón Corriente (Current Ratio) &nbsp; {badge_rc}</div>
                <div class="kpi-value">{rc_act:.2f}x</div>
                <div class="kpi-sub">Anterior: {rc_ant:.2f}x</div>
                </div>""",
                unsafe_allow_html=True,
            )
            if PLOTLY_DISPONIBLE:
                st.plotly_chart(crear_sparkline(rc_ant, rc_act, "#D97706"), use_container_width=True)

        with col_der:
            st.markdown(
                f"""<div class="kpi-card">
                <div class="kpi-title">Disponible en Caja y Bancos (Grupo 11)</div>
                <div class="kpi-value">${d_act['disponible']:,.0f}</div>
                <div class="kpi-sub">
                    <span style="color:#0F766E;font-weight:600;">Entradas (Débitos): ${d_act['disp_debitos']:,.0f}</span> &nbsp;|&nbsp;
                    <span style="color:#D97706;font-weight:600;">Salidas (Créditos): ${d_act['disp_creditos']:,.0f}</span>
                </div>
                </div>""",
                unsafe_allow_html=True,
            )
            if PLOTLY_DISPONIBLE:
                fig_liq = go.Figure(
                    data=[
                        go.Bar(name="Saldo Anterior", x=["Disponible", "CxC Totales", "Pasivo Corriente"], y=[d_ant["disponible"], d_ant["cxc_totales"], d_ant["pasivo_corriente"]], marker_color="#94a3b8"),
                        go.Bar(name="Nuevo Saldo", x=["Disponible", "CxC Totales", "Pasivo Corriente"], y=[d_act["disponible"], d_act["cxc_totales"], d_act["pasivo_corriente"]], marker_color="#7B9A29"),
                    ]
                )
                fig_liq.update_layout(barmode="group", height=295, margin=dict(l=10, r=10, t=30, b=20), title="Composición de Liquidez (COP)")
                st.plotly_chart(fig_liq, use_container_width=True)

        st.markdown('<div class="section-banner">⏱️ Tacómetros de Desempeño (Actividad, Endeudamiento y Margen)</div>', unsafe_allow_html=True)
        if PLOTLY_DISPONIBLE:
            g1, g2, g3 = st.columns(3)
            with g1:
                st.plotly_chart(
                    crear_gauge(
                        ind_act["Días Promedio de Cobro - Clientes"],
                        "Días Promedio de Cobro",
                        max(365, ind_act["Días Promedio de Cobro - Clientes"] * 1.1),
                        [
                            {"range": [0, 60], "color": "#bbf7d0"},
                            {"range": [60, 120], "color": "#fde68a"},
                            {"range": [120, max(365, ind_act["Días Promedio de Cobro - Clientes"] * 1.1)], "color": "#fecaca"},
                        ],
                        " d",
                    ),
                    use_container_width=True,
                )
            with g2:
                st.plotly_chart(
                    crear_gauge(
                        ind_act["Endeudamiento Total (%)"],
                        "Endeudamiento Total (%)",
                        100,
                        [
                            {"range": [0, 50], "color": "#bbf7d0"},
                            {"range": [50, 70], "color": "#fde68a"},
                            {"range": [70, 100], "color": "#fecaca"},
                        ],
                        "%",
                    ),
                    use_container_width=True,
                )
            with g3:
                st.plotly_chart(
                    crear_gauge(
                        ind_act["Margen Operacional (%)"],
                        "Margen Operacional (%)",
                        50,
                        [
                            {"range": [-50, 0], "color": "#fecaca"},
                            {"range": [0, 10], "color": "#fde68a"},
                            {"range": [10, 50], "color": "#bbf7d0"},
                        ],
                        "%",
                    ),
                    use_container_width=True,
                )

    with tab2:
        st.markdown("### 📈 Rentabilidad y Estructura de Resultados")
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Margen Operacional", f"{ind_act['Margen Operacional (%)']:.2f}%")
        r2.metric("Margen Neto", f"{ind_act['Margen Neto (%)']:.2f}%")
        r3.metric("ROA (Semestral/Periodo)", f"{ind_act['ROA - Rentabilidad del Activo (%)']:.2f}%")
        r4.metric("ROE (Semestral/Periodo)", f"{ind_act['ROE - Rentabilidad del Patrimonio (%)']:.2f}%")

        if PLOTLY_DISPONIBLE:
            fig_res = go.Figure(
                go.Bar(
                    x=["Ingresos Op. (41)", "Gastos Admin (51)", "Gastos Op. (52)", "Ingresos No Op. (42)", "Resultado Neto"],
                    y=[d_act["ingresos_op"], d_act["gastos_adm"], d_act["gastos_ventas_op"], d_act["ingresos_noop"], d_act["utilidad_neta"]],
                    marker_color=["#0F766E", "#D97706", "#DC2626", "#7B9A29", "#1E293B"],
                )
            )
            fig_res.update_layout(height=340, title="Desglose de Ingresos, Gastos y Resultado Neto (COP)")
            st.plotly_chart(fig_res, use_container_width=True)

        st.info(interpretaciones["rentabilidad"])
        st.info(interpretaciones["actividad"])

    with tab3:
        st.markdown("### 📊 Análisis Horizontal y Vertical del Balance")
        st.dataframe(df_comp, use_container_width=True)
        st.markdown("### 📋 Resumen Completo de Indicadores")
        st.dataframe(df_ind, use_container_width=True)

    with tab4:
        st.markdown("### 🚨 Alertas Automáticas Detectadas")
        for alerta in alertas:
            st.warning(alerta)

        st.markdown("### 🤖 Informe Gerencial Generado con Gemini")
        if st.session_state.get("desde_cache") and st.session_state.get("analisis_gemini"):
            st.success(
                f"⚡ Informe cargado al instante desde el histórico local (generado el {st.session_state.get('fecha_analisis')}). "
                "No fue necesario esperar nueva respuesta de internet."
            )

        if not GEMINI_DISPONIBLE:
            st.info("Instala la librería oficial ejecutando: `py -m pip install google-genai`")
        elif not gemini_api_key:
            st.info("Pega tu API Key de Gemini en la barra lateral para generar la interpretación con IA.")
        else:
            etiqueta_btn = "🔄 Regenerar Análisis Ejecutivo con Gemini" if st.session_state.get("analisis_gemini") else "✨ Generar Análisis Ejecutivo con Gemini"
            if st.button(etiqueta_btn):
                with st.spinner("Consultando a Gemini y guardando en histórico..."):
                    try:
                        texto_ia, modelo_usado = generar_interpretacion_gemini(
                            gemini_api_key, modelo_gemini, df_comp, df_ind, alertas, dias_periodo
                        )
                        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        st.session_state["analisis_gemini"] = texto_ia
                        st.session_state["modelo_usado"] = modelo_usado
                        st.session_state["fecha_analisis"] = ahora
                        st.session_state["desde_cache"] = False

                        guardar_en_historial_ia(
                            {
                                "id": f"{hash_actual}_{int(time.time())}",
                                "fecha_hora": ahora,
                                "archivo": getattr(archivo, "name", "Balance_Siigo.xlsx"),
                                "dias_periodo": dias_periodo,
                                "modelo_usado": modelo_usado,
                                "hash_datos": hash_actual,
                                "activo_total": d_act["activo_total"],
                                "pasivo_total": d_act["pasivo_total"],
                                "utilidad_neta": d_act["utilidad_neta"],
                                "analisis_gemini": texto_ia,
                                "alertas": alertas,
                                "interpretaciones": interpretaciones,
                                "df_ind": df_ind.to_dict(orient="records"),
                            }
                        )
                        st.success("✅ Informe generado y guardado automáticamente en el histórico.")
                    except Exception as e:
                        st.error(f"Error al consultar la API de Gemini: {e}")

        analisis_guardado = st.session_state.get("analisis_gemini", "")
        if analisis_guardado:
            st.caption(f"Modelo: `{st.session_state.get('modelo_usado', modelo_gemini)}` | Fecha: `{st.session_state.get('fecha_analisis', '')}`")
            st.markdown(analisis_guardado)

        st.markdown("---")
        st.markdown("### 📥 Exportar Informe para Revisión")
        col_exp1, col_exp2 = st.columns(2)

        with col_exp1:
            excel_bytes = exportar_excel_informe(df_comp, df_ind, alertas, interpretaciones, analisis_guardado)
            st.download_button(
                label="📊 Descargar Informe en Excel (.xlsx)",
                data=excel_bytes,
                file_name="Informe_Financiero_INAR_IA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col_exp2:
            if not DOCX_DISPONIBLE:
                st.info("Para exportar a Word instala: `py -m pip install python-docx`")
            else:
                word_bytes = exportar_word_membrete(
                    df_ind, alertas, interpretaciones, analisis_guardado, dias_periodo
                )
                st.download_button(
                    label="📄 Descargar Informe Ejecutivo en Word (.docx)",
                    data=word_bytes,
                    file_name="Informe_Ejecutivo_INAR_IA.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

    with tab5:
        st.markdown("### 🗂️ Histórico de Informes Ejecutivos Generados con IA")
        st.caption("Consulta y descarga al instante informes generados previamente, incluso con internet lento o sin conexión.")
        hist_lista = cargar_historial_ia()
        if not hist_lista:
            st.info("Aún no hay informes guardados en el histórico. Genera un análisis con Gemini en la pestaña anterior.")
        else:
            for item in hist_lista:
                titulo_exp = (
                    f"📄 {item.get('fecha_hora')} | {item.get('archivo')} "
                    f"({item.get('dias_periodo')} días) — Modelo: {item.get('modelo_usado')}"
                )
                with st.expander(titulo_exp):
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Activo Total", f"${item.get('activo_total', 0):,.0f}")
                    c2.metric("Pasivo Total", f"${item.get('pasivo_total', 0):,.0f}")
                    c3.metric("Utilidad Neta", f"${item.get('utilidad_neta', 0):,.0f}")
                    st.markdown(item.get("analisis_gemini", ""))

                    b1, b2 = st.columns([2, 1])
                    with b1:
                        if DOCX_DISPONIBLE and item.get("df_ind"):
                            df_ind_hist = pd.DataFrame(item["df_ind"])
                            word_hist = exportar_word_membrete(
                                df_ind_hist,
                                item.get("alertas", []),
                                item.get("interpretaciones", interpretaciones),
                                item.get("analisis_gemini", ""),
                                item.get("dias_periodo", 180),
                            )
                            st.download_button(
                                label="📥 Descargar este informe en Word (.docx)",
                                data=word_hist,
                                file_name=f"Informe_INAR_IA_{item.get('hash_datos', 'hist')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"dl_{item.get('id')}",
                            )
                    with b2:
                        if st.button("🗑️ Eliminar del histórico", key=f"del_{item.get('id')}"):
                            eliminar_del_historial_ia(item.get("id"))
                            st.rerun()
else:
    st.info("👈 Carga el archivo de Balance de Prueba de Siigo (.xlsx) en la barra lateral para iniciar el análisis.")
    if historial_guardado:
        st.markdown("### 🗂️ Informes Ejecutivos Guardados en Histórico (Acceso Instantáneo)")
        for item in historial_guardado[:5]:
            with st.expander(f"📄 {item.get('fecha_hora')} | {item.get('archivo')} ({item.get('dias_periodo')} días)"):
                st.markdown(item.get("analisis_gemini", ""))
                if DOCX_DISPONIBLE and item.get("df_ind"):
                    word_hist = exportar_word_membrete(
                        pd.DataFrame(item["df_ind"]),
                        item.get("alertas", []),
                        item.get("interpretaciones", {"liquidez": "", "endeudamiento": "", "rentabilidad": "", "actividad": ""}),
                        item.get("analisis_gemini", ""),
                        item.get("dias_periodo", 180),
                    )
                    st.download_button(
                        label="📥 Descargar Word (.docx)",
                        data=word_hist,
                        file_name=f"Informe_INAR_IA_{item.get('hash_datos', 'hist')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_home_{item.get('id')}",
                    )
