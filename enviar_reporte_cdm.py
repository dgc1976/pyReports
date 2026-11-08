#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de Automatización: Cuadro de Mando de Instalaciones (Looker Studio / BigQuery)
Genera gráficos embebidos (CID) y tablas dinámicas idénticas a cdm_provision.html,
enviándolo por correo electrónico mediante el relay SMTP configurado en config.json.
"""

import os
import sys
import io
import json
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, date, timedelta
import pandas as pd
import numpy as np

# Generador de gráficos para entornos sin servidor X11 / headless
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

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
            print(f"Aviso: Error al leer {path} ({e}). Usando configuración base.")
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
        UPPER(COALESCE(service_operation, order_category, 'OTROS')) AS servicio,
        COALESCE(normalized_contracts, 'OTRAS') AS contrata,
        COALESCE(normalized_brand, 'OTRAS') AS marca,
        SUM(COALESCE(num_done_orders, 0)) AS ordenes_cerradas
    FROM `{target_table}`
    WHERE DATE(report_date) BETWEEN DATE_SUB(CURRENT_DATE(), INTERVAL 6 DAY) AND CURRENT_DATE()
      AND COALESCE(num_done_orders, 0) > 0
    GROUP BY 1, 2, 3, 4, 5, 6, 7
    ORDER BY dia_cierre ASC;
    """

    query_job = client.query(query_sql, location=loc)
    df = query_job.to_dataframe()
    logging.info(f"¡Éxito! {len(df)} filas extraídas desde BigQuery.")
    return df

# ==========================================
# 3. GENERACIÓN DE GRÁFICOS (PNG EN MEMORIA)
# ==========================================
def generate_chart_images(df, dias_unicos, labels_dias):
    """
    Genera en memoria (bytes PNG) los 6 gráficos correspondientes a cdm_provision.html.
    """
    charts = {}
    plt.rcParams['font.sans-serif'] = 'Arial'
    plt.rcParams['axes.edgecolor'] = '#cbd5e1'
    plt.rcParams['axes.linewidth'] = 0.8

    # 1. Serie Temporal Global (Total, OSP, MásMóvil, KRTV)
    piv_grp = df.pivot_table(index='dia_cierre', columns='grupo', values='ordenes_cerradas', aggfunc='sum', fill_value=0) if not df.empty else pd.DataFrame()
    piv_grp = piv_grp.reindex(dias_unicos, fill_value=0)
    total_dia = [int(df[df['dia_cierre'] == d]['ordenes_cerradas'].sum()) if not df.empty else 0 for d in dias_unicos]
    osp_dia = [int(piv_grp.get('ORANGE', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]
    mm_dia = [int(piv_grp.get('MAS MOVIL', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]
    krt_dia = [int(piv_grp.get('KRTV', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]

    fig, ax = plt.subplots(figsize=(6.2, 3.2), dpi=150)
    ax.plot(labels_dias, total_dia, label='Órdenes Cerradas', color='#22c55e', linestyle='--', linewidth=2.5, marker='o')
    ax.plot(labels_dias, osp_dia, label='OSP (Orange)', color='#f97316', linewidth=2, marker='o')
    ax.plot(labels_dias, mm_dia, label='MAS MOVIL', color='#eab308', linewidth=2, marker='o')
    ax.plot(labels_dias, krt_dia, label='KRTV', color='#0ea5e9', linewidth=2, marker='o')
    ax.set_title('Serie Temporal: Órdenes Cerradas por Grupo', fontsize=10, fontweight='bold', color='#1e293b')
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    ax.legend(frameon=False, fontsize=8, loc='upper left')
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    plt.close(fig)
    charts['chart_timeseries'] = buf.getvalue()

    # 2. Barras por Servicio (Alta, Traslado, Migración)
    piv_srv = df.pivot_table(index='dia_cierre', columns='servicio', values='ordenes_cerradas', aggfunc='sum', fill_value=0) if not df.empty else pd.DataFrame()
    piv_srv = piv_srv.reindex(dias_unicos, fill_value=0)
    b_alta = [int(piv_srv.get('ALTA', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]
    b_tras = [int(piv_srv.get('TRASLADO', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]
    b_migr = [int(piv_srv.get('MIGRACION', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos]

    x = np.arange(len(labels_dias))
    width = 0.26
    fig, ax = plt.subplots(figsize=(6.2, 3.2), dpi=150)
    ax.bar(x - width, b_alta, width, label='ALTA', color='#8aa9ff')
    ax.bar(x, b_tras, width, label='TRASLADO', color='#f8b4a6')
    ax.bar(x + width, b_migr, width, label='MIGRACION', color='#df5353')
    ax.set_xticks(x)
    ax.set_xticklabels(labels_dias, fontsize=8)
    ax.set_title('Evolución Diaria por Servicio', fontsize=10, fontweight='bold', color='#1e293b')
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    ax.legend(frameon=False, fontsize=8)
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    plt.close(fig)
    charts['chart_service_bar'] = buf.getvalue()

    # 3. Ranking de Contratas (Horizontal)
    contratas_grp = df.groupby('contrata')['ordenes_cerradas'].sum().sort_values(ascending=True) if not df.empty else pd.Series()
    top_c = contratas_grp.tail(8)
    fig, ax = plt.subplots(figsize=(6.2, 3.2), dpi=150)
    ax.barh(list(top_c.index), list(top_c.values), color='#a3d9a5', edgecolor='#86efac', height=0.6)
    ax.set_title('Cierres Acumulados por Contrata', fontsize=10, fontweight='bold', color='#1e293b')
    ax.grid(axis='x', linestyle=':', alpha=0.6)
    for i, v in enumerate(top_c.values):
        ax.text(v + (max(top_c.values)*0.01), i, f"{v:,}".replace(",", "."), va='center', fontsize=7.5, color='#334155')
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    plt.close(fig)
    charts['chart_contratas'] = buf.getvalue()

    # Función auxiliar para gráficos de líneas de subgrupos
    def line_subgroup(group_name, title_txt):
        sub_df = df[df['grupo'] == group_name] if not df.empty else pd.DataFrame()
        p = sub_df.pivot_table(index='dia_cierre', columns='servicio', values='ordenes_cerradas', aggfunc='sum', fill_value=0) if not sub_df.empty else pd.DataFrame()
        p = p.reindex(dias_unicos, fill_value=0)
        fig, ax = plt.subplots(figsize=(6.2, 3.2), dpi=150)
        ax.plot(labels_dias, [int(p.get('ALTA', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos], label='ALTA', color='#8aa9ff', marker='o')
        ax.plot(labels_dias, [int(p.get('TRASLADO', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos], label='TRASLADO', color='#f8b4a6', marker='o')
        ax.plot(labels_dias, [int(p.get('MIGRACION', pd.Series(0, index=dias_unicos))[d]) for d in dias_unicos], label='MIGRACION', color='#df5353', marker='o')
        ax.set_title(title_txt, fontsize=10, fontweight='bold', color='#1e293b')
        ax.grid(axis='y', linestyle=':', alpha=0.6)
        ax.legend(frameon=False, fontsize=8)
        plt.tight_layout()
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        plt.close(fig)
        return buf.getvalue()

    charts['chart_orange'] = line_subgroup('ORANGE', 'Grupo ORANGE por Servicio')
    charts['chart_masmovil'] = line_subgroup('MAS MOVIL', 'Grupo MAS MOVIL por Servicio')
    charts['chart_krtv'] = line_subgroup('KRTV', 'Grupo KRTV por Servicio')

    return charts

# ==========================================
# 4. CONSTRUCCIÓN DE TABLAS Y CONTENIDO
# ==========================================
def build_html_report(df, template_content):
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
    altas_df = df[df['servicio'] == 'ALTA'] if not df.empty else pd.DataFrame()
    traslados_df = df[df['servicio'] == 'TRASLADO'] if not df.empty else pd.DataFrame()
    migraciones_df = df[df['servicio'] == 'MIGRACION'] if not df.empty else pd.DataFrame()

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
        top_contrata, top_contrata_val, top_contrata_pct = "N/A", 0, 0

    table_headers_days = "".join([f"<th>{l}</th>" for l in labels_dias])

    # 1. Tabla Dinámica Contratas
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

    # 2. Tabla Dinámica Marcas
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

    # 3. Tabla Dinámica Servicio & Segmento (con subtotales idénticos a cdm_provision.html)
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
                    df_gss = df_gs[df_gs['servicio'] == srv]
                    if df_gss.empty:
                        continue
                    piv_s = df_gss.groupby('dia_cierre')['ordenes_cerradas'].sum()
                    cols_html = "".join([f"<td>{int(piv_s.get(d, 0)):,}</td>".replace(",", ".") for d in dias_unicos])
                    tot_srv = int(piv_s.sum())
                    badge = badge_map.get(srv.upper(), f'<span class="badge-serv">{srv}</span>')
                    rows_servicio += f"<tr><td class='col-left'><strong>{grp}</strong></td><td class='col-left'>{seg}</td><td class='col-left'>{badge}</td>{cols_html}<td>{tot_srv:,}</td></tr>\n".replace(",", ".")

                piv_seg = df_gs.groupby('dia_cierre')['ordenes_cerradas'].sum()
                cols_sub = "".join([f"<td>{int(piv_seg.get(d, 0)):,}</td>".replace(",", ".") for d in dias_unicos])
                rows_servicio += f"<tr class='row-subtotal'><td class='col-left' colspan='3'>Subtotal {seg}</td>{cols_sub}<td>{int(piv_seg.sum()):,}</td></tr>\n".replace(",", ".")

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

    html_final = template_content
    for tag, val in replacements.items():
        html_final = html_final.replace(tag, val)

    kpis_dict = {
        "rango": rango_txt,
        "total": f"{total_cierres:,}".replace(",", ".")
    }

    chart_images = generate_chart_images(df, dias_unicos, labels_dias)
    return html_final, kpis_dict, chart_images

# ==========================================
# 5. ENVÍO SMTP CON IMÁGENES EMBEBIDAS
# ==========================================
def send_email_with_report(html_content, kpis, chart_images, cfg):
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

    # Estructura MIMEMultipart('related') para permitir imágenes embebidas (CID)
    msg_root = MIMEMultipart('related')
    msg_root['Subject'] = subject
    msg_root['From'] = sender
    msg_root['To'] = ", ".join(recipients)
    if cc_list:
        msg_root['Cc'] = ", ".join(cc_list)

    msg_alt = MIMEMultipart('alternative')
    msg_root.attach(msg_alt)

    # Adjuntar HTML en el cuerpo
    msg_alt.attach(MIMEText(html_content, 'html', 'utf-8'))

    # Adjuntar cada gráfico referenciado como cid:chart_...
    for cid_name, img_bytes in chart_images.items():
        img_part = MIMEImage(img_bytes)
        img_part.add_header('Content-ID', f'<{cid_name}>')
        img_part.add_header('Content-Disposition', 'inline', filename=f"{cid_name}.png")
        msg_root.attach(img_part)

    # Adjuntar también el archivo HTML para abrir fuera del correo
    if report_cfg.get("send_attachment", True):
        filename = f"cdm_evolucion_cierre_{datetime.now().strftime('%Y%m%d')}.html"
        attachment = MIMEBase('application', 'octet-stream')
        attachment.set_payload(html_content.encode('utf-8'))
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg_root.attach(attachment)

    logging.info(f"Conectando al servidor SMTP {server_host}:{server_port}...")
    server = smtplib.SMTP(server_host, server_port)
    if use_tls:
        server.starttls()
    server.sendmail(sender, all_to, msg_root.as_string())
    server.quit()
    logging.info(f"¡Correo enviado con éxito a {', '.join(all_to)} con los 6 gráficos embebidos!")

# ==========================================
# 6. ENTRADA PRINCIPAL
# ==========================================
if __name__ == "__main__":
    logging.info("=== Inicio de generación de CDM de Instalaciones ===")

    template_path = os.getenv("TEMPLATE_PATH", "plantilla_cdm_mail.html")
    if os.path.exists(template_path):
        with open(template_path, "r", encoding="utf-8") as f:
            template_code = f.read()
    else:
        logging.error(f"No se encontró la plantilla en {template_path}. Guarda la Sección 3 en dicho archivo.")
        sys.exit(1)

    df_data = fetch_bigquery_data(CFG)
    html_final, kpis, charts = build_html_report(df_data, template_code)

    local_dir = CFG.get("reports", {}).get("ir", {}).get("local_output_dir", "./output")
    try:
        os.makedirs(local_dir, exist_ok=True)
        local_output_file = os.path.join(local_dir, f"cdm_instalaciones_{date.today().strftime('%Y%m%d')}.html")
    except Exception:
        local_output_file = f"cdm_instalaciones_{date.today().strftime('%Y%m%d')}.html"

    with open(local_output_file, "w", encoding="utf-8") as f:
        f.write(html_final)
    logging.info(f"Archivo HTML persistido en: {local_output_file}")

    try:
        send_email_with_report(html_final, kpis, charts, CFG)
    except Exception as e:
        logging.error(f"Error al enviar correo por SMTP: {e}", exc_info=True)

    logging.info("=== Proceso finalizado correctamente ===")