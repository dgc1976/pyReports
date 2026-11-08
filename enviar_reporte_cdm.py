#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de Automatización: Cuadro de Mando de Instalaciones (Looker Studio / BigQuery)
Carga los parámetros desde config.json, consulta en vivo la vista de BigQuery,
genera el informe HTML y lo despacha a través del relay SMTP corporativo.
"""

import os
import sys
import json
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, date, timedelta
import pandas as pd
from google.cloud import bigquery

# ==========================================
# 1. CARGA DE CONFIG.JSON & LOGGING
# ==========================================
CONFIG_FILE_PATH = os.getenv("CONFIG_PATH", "config.json")

def load_configuration(path=CONFIG_FILE_PATH):
    """Carga el fichero de configuración con valores por defecto de respaldo."""
    default_cfg = {
        "bigquery": {
            "project_id": "mo-provision-reporting",
            "dataset_table_ir": "mo-provision-reporting.INSTALLATIONS_prod.VIS_IR_prod"
        },
        "smtp": {
            "server": "fsp.si.orange.es",
            "port": 587,
            "use_tls": False,
            "sender": "noreply@es.orange.com",
            "recipients": ["david.gonzalezc@masorange.es"],
            "cc": ["david.gonzalezc@masorange.es"]
        },
        "reports": {
            "ir": {
                "subject_prefix": "Orange - Informe Diario Entradas y Cierres / IR",
                "send_attachment": True,  # True para enviar el .html adjunto
                "local_output_dir": "D:\\Automaticos\\Extraciones Datos Informes\\Informes\\IR"
            }
        },
        "logging": {
            "level": "INFO",
            "log_dir": "D:\\Automaticos\\Logs\\reports",
            "log_file": "report_automation.log"
        }
    }

    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                # Actualización recursiva básica
                for k, v in loaded.items():
                    if isinstance(v, dict) and k in default_cfg:
                        default_cfg[k].update(v)
                    else:
                        default_cfg[k] = v
            print(f"Configuración cargada con éxito desde: {path}")
        except Exception as e:
            print(f"Aviso: Error al leer {path} ({e}). Usando configuración por defecto.")
    else:
        print(f"Aviso: {path} no encontrado en el directorio actual. Usando configuración base.")

    return default_cfg

CFG = load_configuration()

# Configurar Logging
log_dir = CFG["logging"].get("log_dir", "./logs")
log_file = CFG["logging"].get("log_file", "report_automation.log")
log_level = getattr(logging, CFG["logging"].get("level", "INFO").upper(), logging.INFO)

try:
    os.makedirs(log_dir, exist_ok=True)
    full_log_path = os.path.join(log_dir, log_file)
except Exception:
    full_log_path = log_file

logging.basicConfig(
    level=log_level,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(full_log_path, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

LOOKER_URL = "https://datastudio.google.com/reporting/cc244b22-6794-4bd5-bf1f-66d2c4a64317/page/baECG"

# ==========================================
# 2. CONSULTAS A BIGQUERY
# ==========================================
def fetch_bigquery_data(project_id):
    # Cargar la configuración para obtener la ubicación y la tabla correctas
    with open('config.json', 'r') as f:
        config = json.load(f)
    
    bq_config = config.get("bigquery", {})
    # 1. Obtener la ubicación (por defecto 'europe-southwest1' o 'EU')
    location = bq_config.get("location", "europe-southwest1")
    # 2. Obtener la tabla/vista correcta
    table_name = bq_config.get("dataset_table_ir", "mo-provision-reporting.INSTALLATION_prod.VW_ENTRADA_CIERRE")
    
    # Inicializar el cliente con el proyecto y la ubicación de Europa específicos
    client = bigquery.Client(project=project_id, location=location)
    
    # Tu query usando la tabla parametrizada
    query_raw = """
    SELECT
        DATE(FECHA_CIERRE) AS dia_cierre,
        FORMAT_DATE('%d %b', DATE(FECHA_CIERRE)) AS label_dia,
        COALESCE(GRUPO, 'NO_DEFINIDO') AS grupo,
        COALESCE(SEGMENTO, 'NO_DEFINIDO') AS segmento,
        COALESCE(SERVICIO, 'OTROS') AS servicio,
        COALESCE(CONTRATA, 'OTRAS') AS contrata,
        COALESCE(MARCA, 'OTRAS') AS marca,
        COUNT(DISTINCT ID_ORDEN) AS ordenes_cerradas
    FROM `mo-provision-reporting.INSTALLATION_prod.VW_ENTRADA_CIERRE`
    WHERE DATE(FECHA_CIERRE) BETWEEN DATE_SUB(CURRENT_DATE(), INTERVAL 6 DAY) AND CURRENT_DATE()
      AND UPPER(COALESCE(ESTADO, '')) IN ('CERRADA', 'CERRADA_OK', 'FINALIZADA', 'EJECUTADA')
    GROUP BY 1, 2, 3, 4, 5, 6, 7
    ORDER BY dia_cierre ASC;
    """
    
    # Es CRUCIAL pasar también el parámetro 'location' en la llamada .query()
    query_job = client.query(query_raw, location=location)
    df = query_job.to_dataframe()
    return df

# ==========================================
# 3. PROCESAMIENTO ANALÍTICO Y GENERACIÓN HTML
# ==========================================
def build_html_report(df, template_html_content):
    """Procesa los datos agregados e inyecta los valores en la plantilla HTML."""
    today = date.today()
    start_date = today - timedelta(days=6)
    rango_txt = f"{start_date.strftime('%d %b')} – {today.strftime('%d %b %Y')}"

    # Eje de fechas (7 días ordenados)
    dias_unicos = sorted(df['dia_cierre'].unique())
    labels_dias = [pd.to_datetime(d).strftime('%d %b') for d in dias_unicos]

    # KPIs Globales
    total_cierres = int(df['ordenes_cerradas'].sum())
    altas_df = df[df['servicio'].str.upper() == 'ALTA']
    traslados_df = df[df['servicio'].str.upper() == 'TRASLADO']
    migraciones_df = df[df['servicio'].str.upper() == 'MIGRACION']

    altas_tot = int(altas_df['ordenes_cerradas'].sum())
    traslados_tot = int(traslados_df['ordenes_cerradas'].sum())
    migraciones_tot = int(migraciones_df['ordenes_cerradas'].sum())

    pct_altas = round((altas_tot / total_cierres * 100), 1) if total_cierres else 0
    pct_traslados = round((traslados_tot / total_cierres * 100), 1) if total_cierres else 0
    pct_migraciones = round((migraciones_tot / total_cierres * 100), 1) if total_cierres else 0

    # Top Contrata
    contratas_grp = df.groupby('contrata')['ordenes_cerradas'].sum().sort_values(ascending=False)
    top_contrata = contratas_grp.index[0] if not contratas_grp.empty else "N/A"
    top_contrata_val = int(contratas_grp.iloc[0]) if not contratas_grp.empty else 0
    top_contrata_pct = round((top_contrata_val / total_cierres * 100), 1) if total_cierres else 0

    # Serie Temporal Global
    piv_grupo = df.pivot_table(index='dia_cierre', columns='grupo', values='ordenes_cerradas', aggfunc='sum', fill_value=0)
    piv_grupo = piv_grupo.reindex(dias_unicos, fill_value=0)
    global_total = [int(df[df['dia_cierre'] == d]['ordenes_cerradas'].sum()) for d in dias_unicos]
    global_osp = [int(piv_grupo.get('ORANGE', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]
    global_mm = [int(piv_grupo.get('MAS MOVIL', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]
    global_krt = [int(piv_grupo.get('KRTV', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]

    # Barras por Servicio
    piv_serv = df.pivot_table(index='dia_cierre', columns='servicio', values='ordenes_cerradas', aggfunc='sum', fill_value=0)
    piv_serv = piv_serv.reindex(dias_unicos, fill_value=0)
    bar_altas = [int(piv_serv.get('ALTA', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]
    bar_traslados = [int(piv_serv.get('TRASLADO', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]
    bar_migraciones = [int(piv_serv.get('MIGRACION', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]

    # Series por Grupo y Servicio
    def get_serv_series_by_group(grp_name):
        sub_df = df[df['grupo'] == grp_name]
        p = sub_df.pivot_table(index='dia_cierre', columns='servicio', values='ordenes_cerradas', aggfunc='sum', fill_value=0)
        p = p.reindex(dias_unicos, fill_value=0)
        return (
            [int(p.get('ALTA', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos],
            [int(p.get('TRASLADO', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos],
            [int(p.get('MIGRACION', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]
        )

    orange_a, orange_t, orange_m = get_serv_series_by_group('ORANGE')
    mm_a, mm_t, mm_m = get_serv_series_by_group('MAS MOVIL')
    krt_a, krt_t, krt_m = get_serv_series_by_group('KRTV')

    # Ranking Contratas (Top 10)
    top10_contratas = contratas_grp.head(10)
    contratas_labels = list(top10_contratas.index)
    contratas_vals = [int(v) for v in top10_contratas.values]

    # Generación de Filas para Tablas Dinámicas
    table_headers_days = "".join([f"<th>{l}</th>" for l in labels_dias])

    # 1. Filas Tabla Contratas
    piv_eecc = df.pivot_table(index='contrata', columns='dia_cierre', values='ordenes_cerradas', aggfunc='sum', fill_value=0)
    piv_eecc['Total'] = piv_eecc.sum(axis=1)
    piv_eecc = piv_eecc.sort_values(by='Total', ascending=False)

    rows_eecc = ""
    for contrata_name, row in piv_eecc.iterrows():
        cols_html = "".join([f"<td>{int(row.get(d, 0)):,}</td>".replace(",", ".") for d in dias_unicos])
        rows_eecc += f"<tr><td class='col-left'><strong>{contrata_name}</strong></td>{cols_html}<td><strong>{int(row['Total']):,}</strong></td></tr>\n".replace(",", ".")

    # 2. Filas Tabla Marcas
    piv_marca = df.pivot_table(index='marca', columns='dia_cierre', values='ordenes_cerradas', aggfunc='sum', fill_value=0)
    piv_marca['Total'] = piv_marca.sum(axis=1)
    piv_marca = piv_marca.sort_values(by='Total', ascending=False)

    rows_marca = ""
    for marca_name, row in piv_marca.iterrows():
        cols_html = "".join([f"<td>{int(row.get(d, 0)):,}</td>".replace(",", ".") for d in dias_unicos])
        rows_marca += f"<tr><td class='col-left'><strong>{marca_name}</strong></td>{cols_html}<td><strong>{int(row['Total']):,}</strong></td></tr>\n".replace(",", ".")

    # 3. Filas Tabla Dinámica Grupo / Segmento / Servicio
    piv_det = df.pivot_table(index=['grupo', 'segmento', 'servicio'], columns='dia_cierre', values='ordenes_cerradas', aggfunc='sum', fill_value=0)
    piv_det['Total'] = piv_det.sum(axis=1)

    badge_map = {
        'ALTA': '<span class="badge-serv b-alta">ALTA</span>',
        'TRASLADO': '<span class="badge-serv b-traslado">TRASLADO</span>',
        'MIGRACION': '<span class="badge-serv b-migracion">MIGRACION</span>'
    }

    rows_servicio = ""
    for (g, s, serv), row in piv_det.iterrows():
        cols_html = "".join([f"<td>{int(row.get(d, 0)):,}</td>".replace(",", ".") for d in dias_unicos])
        serv_badge = badge_map.get(serv.upper(), serv)
        rows_servicio += f"<tr><td class='col-left'>{g}</td><td class='col-left'>{s}</td><td class='col-left'>{serv_badge}</td>{cols_html}<td><strong>{int(row['Total']):,}</strong></td></tr>\n".replace(",", ".")

    # Reemplazo de marcadores en el HTML
    replacements = {
        "{{FECHA_ACTUALIZACION}}": datetime.now().strftime('%d/%m/%Y %H:%M'),
        "{{FECHA_RANGO}}": rango_txt,
        "{{TOTAL_CIERRES}}": f"{total_cierres:,}".replace(",", "."),
        "{{ALTAS_TOTAL}}": f"{altas_tot:,}".replace(",", "."),
        "{{PCT_ALTAS}}": str(pct_altas),
        "{{TRASLADOS_TOTAL}}": f"{traslados_tot:,}".replace(",", "."),
        "{{PCT_TRASLADOS}}": str(pct_traslados),
        "{{MIGRACIONES_TOTAL}}": f"{migraciones_tot:,}".replace(",", "."),
        "{{PCT_MIGRACIONES}}": str(pct_migraciones),
        "{{TOP_CONTRATA_NOMBRE}}": str(top_contrata),
        "{{TOP_CONTRATA_VALOR}}": f"{top_contrata_val:,}".replace(",", "."),
        "{{TOP_CONTRATA_PCT}}": str(top_contrata_pct),
        "{{TABLE_HEADER_DAYS}}": table_headers_days,
        "{{PIVOT_SERVICIO_ROWS}}": rows_servicio,
        "{{PIVOT_CONTRATAS_ROWS}}": rows_eecc,
        "{{PIVOT_MARCAS_ROWS}}": rows_marca,
        "{{DAYS_LABELS_JSON}}": json.dumps(labels_dias),
        "{{GLOBAL_TOTAL_DATA_JSON}}": json.dumps(global_total),
        "{{GLOBAL_OSP_DATA_JSON}}": json.dumps(global_osp),
        "{{GLOBAL_MM_DATA_JSON}}": json.dumps(global_mm),
        "{{GLOBAL_KRT_DATA_JSON}}": json.dumps(global_krt),
        "{{BAR_ALTAS_DATA_JSON}}": json.dumps(bar_altas),
        "{{BAR_TRASLADOS_DATA_JSON}}": json.dumps(bar_traslados),
        "{{BAR_MIGRACIONES_DATA_JSON}}": json.dumps(bar_migraciones),
        "{{CONTRATAS_LABELS_JSON}}": json.dumps(contratas_labels),
        "{{CONTRATAS_VALUES_JSON}}": json.dumps(contratas_vals),
        "{{ORANGE_ALTAS_JSON}}": json.dumps(orange_a),
        "{{ORANGE_TRASLADOS_JSON}}": json.dumps(orange_t),
        "{{ORANGE_MIGRACIONES_JSON}}": json.dumps(orange_m),
        "{{MM_ALTAS_JSON}}": json.dumps(mm_a),
        "{{MM_TRASLADOS_JSON}}": json.dumps(mm_t),
        "{{MM_MIGRACIONES_JSON}}": json.dumps(mm_m),
        "{{KRTV_ALTAS_JSON}}": json.dumps(krt_a),
        "{{KRTV_TRASLADOS_JSON}}": json.dumps(krt_t),
        "{{KRTV_MIGRACIONES_JSON}}": json.dumps(krt_m)
    }

    html_final = template_html_content
    for tag, val in replacements.items():
        html_final = html_final.replace(tag, val)

    kpis_dict = {
        "rango": rango_txt,
        "total": f"{total_cierres:,}".replace(",", "."),
        "altas": f"{altas_tot:,}".replace(",", "."),
        "traslados": f"{traslados_tot:,}".replace(",", "."),
        "migraciones": f"{migraciones_tot:,}".replace(",", "."),
        "top_contrata": top_contrata,
        "top_contrata_val": f"{top_contrata_val:,}".replace(",", ".")
    }

    return html_final, kpis_dict

# ==========================================
# 4. ENVÍO DE CORREO POR SMTP (CONFIG.JSON)
# ==========================================
def send_email_with_report(html_content, kpis, cfg):
    """Envía el correo utilizando la configuración definida en config.json."""
    smtp_cfg = cfg["smtp"]
    report_cfg = cfg.get("reports", {}).get("ir", {})

    server_host = smtp_cfg.get("server", "fsp.si.orange.es")
    server_port = int(smtp_cfg.get("port", 587))
    use_tls = bool(smtp_cfg.get("use_tls", False))
    sender = smtp_cfg.get("sender", "noreply@es.orange.com")
    recipients = smtp_cfg.get("recipients", ["david.gonzalezc@masorange.es"])
    cc_list = smtp_cfg.get("cc", [])
    all_to = recipients + [c for c in cc_list if c not in recipients]

    subject_prefix = report_cfg.get("subject_prefix", "Orange - Informe Diario Entradas y Cierres / IR")
    today_str = datetime.now().strftime('%d/%m/%Y')
    subject = f"{subject_prefix} - Evolución del Cierre ({kpis['rango']}) - {today_str}"

    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ", ".join(recipients)
    if cc_list:
        msg['Cc'] = ", ".join(cc_list)

    # Cuerpo ejecutivo del correo
    body_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        body {{ font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif; color: #1e293b; background: #f8fafc; margin: 0; padding: 20px; }}
        .card {{ background: #ffffff; border-radius: 8px; border: 1px solid #e2e8f0; padding: 24px; max-width: 680px; margin: auto; }}
        .header {{ border-bottom: 2px solid #ff6600; padding-bottom: 12px; margin-bottom: 16px; }}
        .title {{ font-size: 18px; font-weight: bold; color: #ff6600; margin: 0; }}
        .sub {{ font-size: 13px; color: #64748b; margin-top: 4px; }}
        .kpi-table {{ width: 100%; border-collapse: collapse; margin: 16px 0; font-size: 14px; }}
        .kpi-table th, .kpi-table td {{ padding: 9px 12px; border-bottom: 1px solid #f1f5f9; }}
        .kpi-table th {{ background-color: #fff8ee; color: #334155; text-align: left; }}
        .btn {{ display: inline-block; background: #ff6600; color: #ffffff !important; padding: 10px 20px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 14px; margin-top: 14px; }}
      </style>
    </head>
    <body>
      <div class="card">
        <div class="header">
          <div class="title">Cuadro de Mando: Evolución del Cierre de Instalaciones</div>
          <div class="sub">Periodo: <strong>{kpis['rango']}</strong> | Origen: <code>mo-provision-reporting.INSTALLATION_prod.VW_ENTRADA_CIERRE</code></div>
        </div>
        <p>Estimado equipo,</p>
        <p>Se ha generado el resumen operativo del cierre de instalaciones para los últimos 7 días:</p>

        <table class="kpi-table">
          <thead>
            <tr><th>Métrica Operativa</th><th style="text-align: right;">Volumen</th></tr>
          </thead>
          <tbody>
            <tr><td><strong>Total Órdenes Cerradas</strong></td><td align="right"><strong>{kpis['total']}</strong></td></tr>
            <tr><td>Altas de Instalación</td><td align="right">{kpis['altas']}</td></tr>
            <tr><td>Traslados</td><td align="right">{kpis['traslados']}</td></tr>
            <tr><td>Migraciones</td><td align="right">{kpis['migraciones']}</td></tr>
            <tr><td>Top Contrata</td><td align="right"><strong>{kpis['top_contrata']}</strong> ({kpis['top_contrata_val']})</td></tr>
          </tbody>
        </table>

        <p style="font-size: 13px; color: #64748b;">
          El archivo interactivo adjunto contiene los gráficos detallados de series temporales y las 3 tablas dinámicas navegables (Grupo/Servicio, Contratas y Marcas).
        </p>

        <a href="{LOOKER_URL}" class="btn" target="_blank">Ver Informe en Looker Studio</a>
      </div>
    </body>
    </html>
    """
    msg.attach(MIMEText(body_html, 'html'))

    # Adjuntar archivo HTML interactivo (según configuración)
    send_attachment = report_cfg.get("send_attachment", True)
    if send_attachment:
        filename = f"cdm_evolucion_cierre_{datetime.now().strftime('%Y%m%d')}.html"
        attachment = MIMEBase('application', 'octet-stream')
        attachment.set_payload(html_content.encode('utf-8'))
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(attachment)

    # Envío a través del relay SMTP corporativo
    logging.info(f"Conectando a {server_host}:{server_port} (TLS={use_tls})...")
    server = smtplib.SMTP(server_host, server_port)
    if use_tls:
        server.starttls()

    server.sendmail(sender, all_to, msg.as_string())
    server.quit()
    logging.info(f"Correo enviado correctamente a: {', '.join(all_to)}")

# ==========================================
# 5. PUNTO DE ENTRADA PRINCIPAL
# ==========================================
if __name__ == "__main__":
    logging.info("=== Inicio del proceso de generación y envío de CDM ===")

    # Cargar plantilla HTML
    template_path = os.getenv("TEMPLATE_PATH", "plantilla_cdm.html")
    if os.path.exists(template_path):
        with open(template_path, "r", encoding="utf-8") as f:
            template_code = f.read()
    else:
        logging.warning(f"No se encontró {template_path}. Asegúrate de exportar el HTML de la Sección 3 como plantilla.")
        sys.exit(1)

    # 1. Extracción de datos desde BigQuery
    bq_project = CFG.get("bigquery", {}).get("project_id", "mo-provision-reporting")
    df_data = fetch_bigquery_data(bq_project)

    # 2. Generación del reporte HTML dinámico
    html_final, kpis = build_html_report(df_data, template_code)

    # 3. Guardado en la ruta local definida en config.json
    local_dir = CFG.get("reports", {}).get("ir", {}).get("local_output_dir", "./output")
    try:
        os.makedirs(local_dir, exist_ok=True)
        local_output_file = os.path.join(local_dir, f"cdm_instalaciones_{date.today().strftime('%Y%m%d')}.html")
    except Exception as e:
        logging.warning(f"No fue posible escribir en {local_dir} ({e}). Guardando en directorio local.")
        local_output_file = f"cdm_instalaciones_{date.today().strftime('%Y%m%d')}.html"

    with open(local_output_file, "w", encoding="utf-8") as f:
        f.write(html_final)
    logging.info(f"Reporte HTML guardado localmente en: {local_output_file}")

    # 4. Envío por SMTP
    try:
        send_email_with_report(html_final, kpis, CFG)
    except Exception as e:
        logging.error(f"Error al enviar el correo por SMTP: {e}", exc_info=True)

    logging.info("=== Proceso finalizado ===")