#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de Automatización: Cuadro de Mando de Instalaciones (Looker Studio / BigQuery)
Consulta mo-provision-reporting.INSTALLATIONS_looker.VW_ENTRADA_CIERRE en la región EU,
genera el cuerpo de correo idéntico a cdm_provision.html y lo envía por SMTP.
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
    default_cfg = {
        "bigquery": {
            "project_id": "mo-provision-reporting",
            "table_name": "mo-provision-reporting.INSTALLATIONS_looker.VW_ENTRADA_CIERRE",
            "location": "EU"
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
                "send_attachment": True,
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
                for k, v in loaded.items():
                    if isinstance(v, dict) and k in default_cfg:
                        default_cfg[k].update(v)
                    else:
                        default_cfg[k] = v
            print(f"Configuración cargada con éxito desde: {path}")
        except Exception as e:
            print(f"Aviso: Error al leer {path} ({e}). Usando configuración por defecto.")
    else:
        print(f"Aviso: {path} no encontrado. Usando configuración base.")

    return default_cfg

CFG = load_configuration()

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
# 2. EXTRACCIÓN DIRECTA DESDE BIGQUERY
# ==========================================
def fetch_bigquery_data(cfg):
    bq_cfg = cfg.get("bigquery", {})
    project_id = bq_cfg.get("project_id", "mo-provision-reporting")
    target_table = bq_cfg.get("table_name", "mo-provision-reporting.INSTALLATIONS_looker.VW_ENTRADA_CIERRE")
    loc = bq_cfg.get("location", "EU")

    logging.info(f"Conectando a BigQuery: `{target_table}` (Región: {loc})...")
    client = bigquery.Client(project=project_id, location=loc)

    query_sql = f"""
    SELECT
        DATE(report_date) AS dia_cierre,
        FORMAT_DATE('%d %b', DATE(report_date)) AS label_dia,
        COALESCE(brand_group, 'NO_DEFINIDO') AS grupo,
        COALESCE(activity_group_soho_masivo, activity_group, 'NO_DEFINIDO') AS segmento,
        COALESCE(service_operation, order_category, 'OTROS') AS servicio,
        COALESCE(normalized_contracts, 'OTRAS') AS contrata,
        COALESCE(normalized_brand, 'OTRAS') AS marca,
        SUM(COALESCE(num_done_orders, 0)) AS ordenes_cerradas,
        SUM(COALESCE(num_in_ot, 0)) AS ordenes_entradas
    FROM `{target_table}`
    WHERE DATE(report_date) BETWEEN DATE_SUB(CURRENT_DATE(), INTERVAL 6 DAY) AND CURRENT_DATE()
    GROUP BY 1, 2, 3, 4, 5, 6, 7
    ORDER BY dia_cierre ASC;
    """

    query_job = client.query(query_sql, location=loc)
    df = query_job.to_dataframe()
    logging.info(f"¡Éxito! {len(df)} registros agregados obtenidos desde BigQuery.")
    return df

# ==========================================
# 3. PROCESAMIENTO ANALÍTICO Y GENERACIÓN HTML
# ==========================================
def build_html_report(df, template_html_content):
    today = date.today()
    start_date = today - timedelta(days=6)
    rango_txt = f"{start_date.strftime('%d %b')} – {today.strftime('%d %b %Y')}"

    dias_unicos = sorted(df['dia_cierre'].unique())
    if not dias_unicos:
        dias_unicos = [(today - timedelta(days=i)) for i in range(6, -1, -1)]
        labels_dias = [d.strftime('%d %b') for d in dias_unicos]
    else:
        labels_dias = [pd.to_datetime(d).strftime('%d %b') for d in dias_unicos]

    total_cierres = int(df['ordenes_cerradas'].sum()) if not df.empty else 0
    altas_df = df[df['servicio'].str.upper() == 'ALTA'] if not df.empty else pd.DataFrame()
    traslados_df = df[df['servicio'].str.upper() == 'TRASLADO'] if not df.empty else pd.DataFrame()
    migraciones_df = df[df['servicio'].str.upper() == 'MIGRACION'] if not df.empty else pd.DataFrame()

    altas_tot = int(altas_df['ordenes_cerradas'].sum()) if not altas_df.empty else 0
    traslados_tot = int(traslados_df['ordenes_cerradas'].sum()) if not traslados_df.empty else 0
    migraciones_tot = int(migraciones_df['ordenes_cerradas'].sum()) if not migraciones_df.empty else 0

    pct_altas = round((altas_tot / total_cierres * 100), 1) if total_cierres else 0
    pct_traslados = round((traslados_tot / total_cierres * 100), 1) if total_cierres else 0
    pct_migraciones = round((migraciones_tot / total_cierres * 100), 1) if total_cierres else 0

    if not df.empty:
        contratas_grp = df.groupby('contrata')['ordenes_cerradas'].sum().sort_values(ascending=False)
        top_contrata = contratas_grp.index[0] if not contratas_grp.empty else "N/A"
        top_contrata_val = int(contratas_grp.iloc[0]) if not contratas_grp.empty else 0
        top_contrata_pct = round((top_contrata_val / total_cierres * 100), 1) if total_cierres else 0
    else:
        contratas_grp = pd.Series(dtype=int)
        top_contrata, top_contrata_val, top_contrata_pct = "N/A", 0, 0

    table_headers_days = "".join([f"<th>{l}</th>" for l in labels_dias])

    # Tabla dinámica: Contratas
    if not df.empty:
        piv_eecc = df.pivot_table(index='contrata', columns='dia_cierre', values='ordenes_cerradas', aggfunc='sum', fill_value=0)
        piv_eecc['Total'] = piv_eecc.sum(axis=1)
        piv_eecc = piv_eecc.sort_values(by='Total', ascending=False)
        rows_eecc = ""
        for contrata_name, row in piv_eecc.iterrows():
            cols_html = "".join([f"<td>{int(row.get(d, 0)):,}</td>".replace(",", ".") for d in dias_unicos])
            rows_eecc += f"<tr><td class='col-left'><strong>{contrata_name}</strong></td>{cols_html}<td><strong>{int(row['Total']):,}</strong></td></tr>\n".replace(",", ".")
    else:
        rows_eecc = "<tr><td colspan='9' style='text-align:center;'>Sin datos registrados</td></tr>"

    # Tabla dinámica: Marcas
    if not df.empty:
        piv_marca = df.pivot_table(index='marca', columns='dia_cierre', values='ordenes_cerradas', aggfunc='sum', fill_value=0)
        piv_marca['Total'] = piv_marca.sum(axis=1)
        piv_marca = piv_marca.sort_values(by='Total', ascending=False)
        rows_marca = ""
        for marca_name, row in piv_marca.iterrows():
            cols_html = "".join([f"<td>{int(row.get(d, 0)):,}</td>".replace(",", ".") for d in dias_unicos])
            rows_marca += f"<tr><td class='col-left'><strong>{marca_name}</strong></td>{cols_html}<td><strong>{int(row['Total']):,}</strong></td></tr>\n".replace(",", ".")
    else:
        rows_marca = "<tr><td colspan='9' style='text-align:center;'>Sin datos registrados</td></tr>"

    # Tabla dinámica: Servicio & Segmento (con subtotales idénticos a cdm_provision.html)
    rows_servicio = ""
    if not df.empty:
        grupos_orden = ['ORANGE', 'MAS MOVIL', 'KRTV'] + [g for g in df['grupo'].unique() if g not in ['ORANGE', 'MAS MOVIL', 'KRTV']]
        badge_map = {
            'ALTA': '<span class="badge-serv b-alta">ALTA</span>',
            'TRASLADO': '<span class="badge-serv b-traslado">TRASLADO</span>',
            'MIGRACION': '<span class="badge-serv b-migracion">MIGRACION</span>'
        }
        for grp in grupos_orden:
            df_g = df[df['grupo'] == grp]
            if df_g.empty:
                continue
            segmentos = df_g['segmento'].unique()
            for seg in segmentos:
                df_gs = df_g[df_g['segmento'] == seg]
                servicios = ['ALTA', 'TRASLADO', 'MIGRACION'] + [s for s in df_gs['servicio'].unique() if s.upper() not in ['ALTA', 'TRASLADO', 'MIGRACION']]
                for srv in servicios:
                    df_gss = df_gs[df_gs['servicio'].str.upper() == srv.upper()]
                    if df_gss.empty:
                        continue
                    piv_s = df_gss.groupby('dia_cierre')['ordenes_cerradas'].sum()
                    cols_html = "".join([f"<td>{int(piv_s.get(d, 0)):,}</td>".replace(",", ".") for d in dias_unicos])
                    tot_srv = int(piv_s.sum())
                    badge = badge_map.get(srv.upper(), f'<span class="badge-serv">{srv}</span>')
                    rows_servicio += f"<tr><td class='col-left'><strong>{grp}</strong></td><td class='col-left'>{seg}</td><td class='col-left'>{badge}</td>{cols_html}<td>{tot_srv:,}</td></tr>\n".replace(",", ".")
                
                # Subtotal Segmento
                piv_seg = df_gs.groupby('dia_cierre')['ordenes_cerradas'].sum()
                cols_sub = "".join([f"<td>{int(piv_seg.get(d, 0)):,}</td>".replace(",", ".") for d in dias_unicos])
                rows_servicio += f"<tr class='row-subtotal'><td class='col-left' colspan='3'>Subtotal {seg}</td>{cols_sub}<td>{int(piv_seg.sum()):,}</td></tr>\n".replace(",", ".")
            
            # Subtotal Grupo
            piv_grp = df_g.groupby('dia_cierre')['ordenes_cerradas'].sum()
            cols_grp = "".join([f"<td>{int(piv_grp.get(d, 0)):,}</td>".replace(",", ".") for d in dias_unicos])
            rows_servicio += f"<tr class='row-total-group'><td class='col-left' colspan='3'>Total {grp}</td>{cols_grp}<td><strong>{int(piv_grp.sum()):,}</strong></td></tr>\n".replace(",", ".")
    else:
        rows_servicio = "<tr><td colspan='11' style='text-align:center;'>Sin datos registrados</td></tr>"

    replacements = {
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
        "{{PIVOT_MARCAS_ROWS}}": rows_marca
    }

    html_final = template_html_content
    for tag, val in replacements.items():
        html_final = html_final.replace(tag, val)

    kpis_dict = {
        "rango": rango_txt,
        "total": f"{total_cierres:,}".replace(",", ".")
    }

    return html_final, kpis_dict

# ==========================================
# 4. ENVÍO DE CORREO POR SMTP
# ==========================================
def send_email_with_report(html_content, kpis, cfg):
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

    # El cuerpo del correo es exactamente el HTML del dashboard
    msg.attach(MIMEText(html_content, 'html', 'utf-8'))

    # Adjuntar también el archivo .html para visualización local
    send_attachment = report_cfg.get("send_attachment", True)
    if send_attachment:
        filename = f"cdm_evolucion_cierre_{datetime.now().strftime('%Y%m%d')}.html"
        attachment = MIMEBase('application', 'octet-stream')
        attachment.set_payload(html_content.encode('utf-8'))
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(attachment)

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

    template_path = os.getenv("TEMPLATE_PATH", "plantilla_cdm.html")
    if os.path.exists(template_path):
        with open(template_path, "r", encoding="utf-8") as f:
            template_code = f.read()
    else:
        logging.warning(f"No se encontró {template_path}. Asegúrate de guardar la plantilla de la Sección 3.")
        sys.exit(1)

    df_data = fetch_bigquery_data(CFG)
    html_final, kpis = build_html_report(df_data, template_code)

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

    try:
        send_email_with_report(html_final, kpis, CFG)
    except Exception as e:
        logging.error(f"Error al enviar el correo por SMTP: {e}", exc_info=True)

    logging.info("=== Proceso finalizado ===")