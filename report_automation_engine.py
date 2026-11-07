#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Framework de Automatización de Reportes Operativos - MASORANGE
Diseño: Programación Orientada a Objetos (POO) & Arquitectura Modular por Capas.

Capas:
  1. ConfigManager: Carga de configuración JSON y merge con argumentos CLI.
  2. BigQueryExtractor: Conexión y ejecución de consultas en Google Cloud BigQuery.
  3. BaseReportProcessor (ABC): Contrato base para procesar cualquier reporte.
  4. IRReportProcessor: Especialización para el Cuadro de Mando IR y Entradas/Cierres.
  5. SMTPEmailSender: Gestión del transporte SMTP, adjuntos seguros y deduplicación.
  6. ReportPipeline: Orquestador general del ciclo de vida del reporte.

Autor: David González Cogolludo
Departamento: Operaciones MASORANGE
"""

import os
import sys
import json
import logging
import smtplib
import argparse
from abc import ABC, abstractmethod
from datetime import date, timedelta
from typing import Dict, List, Any, Tuple, Optional
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from google.cloud import bigquery
from jinja2 import Template


# ==============================================================================
# 1. GESTOR DE CONFIGURACIÓN Y ARGUMENTOS CLI
# ==============================================================================
class ConfigManager:
    """Gestiona la configuración combinando defaults, fichero JSON y argumentos CLI."""

    DEFAULT_CONFIG = {
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
            "cc": []
        },
        "reports": {
            "ir": {
                "subject_prefix": "Informe Diario Entradas y Cierres / IR",
                "attachment_extension": ".txt"
            }
        },
        "logging": {
            "level": "INFO",
            "log_dir": r"D:\Automaticos\Logs\reports",
            "log_file": "report_automation.log"
        }
    }

    def __init__(self, config_path: Optional[str] = None):
        self.config = self.DEFAULT_CONFIG.copy()
        if config_path and os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                self._deep_update(self.config, loaded)

    def _deep_update(self, base: dict, update: dict):
        for k, v in update.items():
            if isinstance(v, dict) and k in base:
                self._deep_update(base[k], v)
            else:
                base[k] = v

    def override_from_args(self, args: argparse.Namespace):
        if getattr(args, "project_id", None):
            self.config["bigquery"]["project_id"] = args.project_id
        if getattr(args, "to", None):
            self.config["smtp"]["recipients"] = [email.strip() for email in args.to.split(",")]
        if getattr(args, "cc", None):
            self.config["smtp"]["cc"] = [email.strip() for email in args.cc.split(",")]
        if getattr(args, "smtp_server", None):
            self.config["smtp"]["server"] = args.smtp_server
        if getattr(args, "smtp_port", None):
            self.config["smtp"]["port"] = args.smtp_port
        if getattr(args, "use_tls", None):
            self.config["smtp"]["use_tls"] = args.use_tls
        if getattr(args, "log_dir", None):
            self.config["logging"]["log_dir"] = args.log_dir


# ==============================================================================
# 2. EXTRACTOR BIGQUERY
# ==============================================================================
class BigQueryExtractor:
    """Encapsula las operaciones de lectura en Google Cloud BigQuery."""

    def __init__(self, project_id: str):
        self.project_id = project_id
        self.client = bigquery.Client(project=self.project_id)

    def execute_query(self, query: str) -> List[Dict[str, Any]]:
        logging.info("Ejecutando consulta en BigQuery...")
        try:
            query_job = self.client.query(query)
            rows = [dict(row) for row in query_job.result()]
            logging.info(f"Consulta completada exitosamente. Registros obtenidos: {len(rows)}")
            return rows
        except Exception as e:
            logging.error(f"Fallo crítico al ejecutar consulta en BigQuery: {e}")
            raise


# ==============================================================================
# 3. CONTRATO BASE PARA REPORTES (ABSTRACT BASE CLASS)
# ==============================================================================
class BaseReportProcessor(ABC):
    """Contrato base que todo nuevo reporte debe implementar."""

    @property
    @abstractmethod
    def report_key(self) -> str:
        """Identificador único del reporte (ej. 'ir', 'tm_visitas')."""
        pass

    @abstractmethod
    def build_query(self, target_date: date) -> str:
        """Construye y retorna la consulta SQL para la fecha objetivo."""
        pass

    @abstractmethod
    def transform_data(self, raw_rows: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Procesa y estructura los datos crudos devueltos por BigQuery."""
        pass

    @abstractmethod
    def render_email_body(self, transformed_data: Dict[str, Any], target_date: date) -> str:
        """Genera el código HTML para el cuerpo del mensaje."""
        pass

    @abstractmethod
    def build_attachment(self, raw_rows: List[Dict[str, Any]], target_date: date) -> Tuple[str, str]:
        """Retorna una tupla (contenido_archivo, nombre_archivo_con_extension)."""
        pass


# ==============================================================================
# 4. IMPLEMENTACIÓN CONCRETA: REPORTE ÍNDICE DE ROTACIÓN (IR)
# ==============================================================================
class IRReportProcessor(BaseReportProcessor):
    """Procesador específico del Cuadro de Mando de IR y Entradas/Cierres."""

    def __init__(self, table_name: str, attachment_ext: str = ".txt"):
        self.table_name = table_name
        self.attachment_ext = attachment_ext

    @property
    def report_key(self) -> str:
        return "ir"

    @staticmethod
    def _fmt(val, dec=0):
        if val is None or val == "":
            return "-"
        try:
            val_f = float(val)
            if dec == 0:
                return f"{int(round(val_f)):,}".replace(",", ".")
            return f"{val_f:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return str(val)

    def build_query(self, target_date: date) -> str:
        date_str = target_date.strftime("%Y-%m-%d")
        return f"""
        WITH TOTAL_IR AS (
          SELECT 
            *,
            CASE WHEN tipo = 'Rec' THEN total_3_dias ELSE 0 END AS MediaRecibidas,
            CASE WHEN tipo = 'Real' THEN total_3_dias ELSE 0 END AS MediaRealizadas,
            CASE WHEN tipo = 'UltDiaReal' THEN total_3_dias ELSE 0 END AS RealUltDiaHabil,

            -- Serie de Recibidas (Entradas D-8 a D-1)
            CASE WHEN dias_atras = 1 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_1,
            CASE WHEN dias_atras = 2 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_2,
            CASE WHEN dias_atras = 3 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_3,
            CASE WHEN dias_atras = 4 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_4,
            CASE WHEN dias_atras = 5 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_5,
            CASE WHEN dias_atras = 6 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_6,
            CASE WHEN dias_atras = 7 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_7,
            CASE WHEN dias_atras = 8 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_8,

            -- Serie de Realizadas (Cierres D-8 a D-1)
            CASE WHEN dias_atras = 1 THEN total_is_ot_done ELSE 0 END AS Realizadas_D_1,
            CASE WHEN dias_atras = 2 THEN total_is_ot_done ELSE 0 END AS Realizadas_D_2,
            CASE WHEN dias_atras = 3 THEN total_is_ot_done ELSE 0 END AS Realizadas_D_3,
            CASE WHEN dias_atras = 4 THEN total_is_ot_done ELSE 0 END AS Realizadas_D_4,
            CASE WHEN dias_atras = 5 THEN total_is_ot_done ELSE 0 END AS Realizadas_D_5,
            CASE WHEN dias_atras = 6 THEN total_is_ot_done ELSE 0 END AS Realizadas_D_6,
            CASE WHEN dias_atras = 7 THEN total_is_ot_done ELSE 0 END AS Realizadas_D_7,
            CASE WHEN dias_atras = 8 THEN total_is_ot_done ELSE 0 END AS Realizadas_D_8,

            CASE WHEN brand_group IN ('MAS MOVIL', 'KRTV') THEN total_pospuesta ELSE 0 END AS TotalPospuestaMM_KRTV
          FROM `{self.table_name}`
          WHERE fecha_referencia = '{date_str}'
        )
        SELECT 
          brand_group AS grupo,
          normalized_contracts AS contrata,
          normalized_province AS provincia,
          ROUND(SUM(MediaRecibidas) / 3, 1) AS rec_media,
          MAX(max_ots) AS rec_max,
          MAX(Recibidas_D_8) AS rec_d8,
          MAX(Recibidas_D_7) AS rec_d7,
          MAX(Recibidas_D_6) AS rec_d6,
          MAX(Recibidas_D_5) AS rec_d5,
          MAX(Recibidas_D_4) AS rec_d4,
          MAX(Recibidas_D_3) AS rec_d3,
          MAX(Recibidas_D_2) AS rec_d2,
          MAX(Recibidas_D_1) AS rec_d1,

          ROUND(SUM(MediaRealizadas) / 3, 1) AS rea_media,
          MAX(max_ots) AS rea_max,
          MAX(Realizadas_D_8) AS rea_d8,
          MAX(Realizadas_D_7) AS rea_d7,
          MAX(Realizadas_D_6) AS rea_d6,
          MAX(Realizadas_D_5) AS rea_d5,
          MAX(Realizadas_D_4) AS rea_d4,
          MAX(Realizadas_D_3) AS rea_d3,
          MAX(Realizadas_D_2) AS rea_d2,
          MAX(Realizadas_D_1) AS rea_d1,

          SUM(total_bolsa) AS bolsa,
          ROUND(SAFE_DIVIDE(SUM(total_bolsa), (SUM(MediaRealizadas) / 3)), 2) AS ir,
          ROUND(SAFE_DIVIDE(SUM(total_bolsa), NULLIF(SUM(RealUltDiaHabil), 0)), 2) AS ir_ayer,
          ROUND(SAFE_DIVIDE(SUM(total_bolsa), NULLIF(SUM(total_citaD), 0)), 2) AS ir_previsto,
          ROUND(SUM(total_bolsa) / 2.5) AS n_inst_bjetivo,
          SUM(total_citaD) AS citado,
          SUM(total_citaD1) AS citad1,
          SUM(total_citaD2) AS citad2,
          ROUND(100 * SAFE_DIVIDE(SUM(total_pdtecita), NULLIF(SUM(total_bolsa), 0)), 1) AS per_pdte_cita,
          ROUND(100 * SAFE_DIVIDE(SUM(total_pospuesta), NULLIF(SUM(total_bolsa) + SUM(TotalPospuestaMM_KRTV), 0)), 1) AS per_pospuesto,
          SUM(total_pdtecita) AS pdtecita,
          SUM(total_pospuesta) AS pospuesta,
          SUM(TotalPospuestaMM_KRTV) AS pospuesta_mm
        FROM TOTAL_IR
        WHERE brand_group IS NOT NULL
          AND normalized_contracts IS NOT NULL
          AND normalized_province IS NOT NULL
        GROUP BY grupo, contrata, provincia
        ORDER BY grupo, contrata, provincia;
        """

    def transform_data(self, raw_rows: List[Dict[str, Any]]) -> Dict[str, Any]:
        timeline_raw = {}
        tot_tl = {k: 0.0 if "media" in k else 0 for k in [
            "rec_media", "rec_max", "rec_d8", "rec_d7", "rec_d6", "rec_d5",
            "rec_d4", "rec_d3", "rec_d2", "rec_d1"
        ]}
        marcas_agg = {}

        tot_bolsa_g = 0
        tot_rec_media_g = 0.0
        tot_rea_media_g = 0.0
        tot_rea_d1_g = 0
        tot_citado_g = 0

        for r in raw_rows:
            g = str(r.get("grupo") or "OTROS").strip()
            c = str(r.get("contrata") or "SIN CONTRATA").strip()

            b = int(r.get("bolsa") or 0)
            rec_m = float(r.get("rec_media") or 0.0)
            rec_max = int(r.get("rec_max") or 0)
            rea_m = float(r.get("rea_media") or 0.0)
            rea_1 = int(r.get("rea_d1") or 0)
            cit = int(r.get("citado") or 0)

            tot_bolsa_g += b
            tot_rec_media_g += rec_m
            tot_rea_media_g += rea_m
            tot_rea_d1_g += rea_1
            tot_citado_g += cit

            if c not in timeline_raw:
                timeline_raw[c] = {k: 0.0 if "media" in k else 0 for k in tot_tl.keys()}

            timeline_raw[c]["rec_media"] += rec_m
            timeline_raw[c]["rec_max"] = max(timeline_raw[c]["rec_max"], rec_max)
            tot_tl["rec_media"] += rec_m
            tot_tl["rec_max"] = max(tot_tl["rec_max"], rec_max)

            for d_idx in range(1, 9):
                key = f"rec_d{d_idx}"
                val = int(r.get(key) or 0)
                timeline_raw[c][key] += val
                tot_tl[key] += val

            if g not in marcas_agg:
                marcas_agg[g] = {"bolsa": 0, "rec_media": 0.0, "rea_media": 0.0, "rea_d1": 0, "citado": 0}
            marcas_agg[g]["bolsa"] += b
            marcas_agg[g]["rec_media"] += rec_m
            marcas_agg[g]["rea_media"] += rea_m
            marcas_agg[g]["rea_d1"] += rea_1
            marcas_agg[g]["citado"] += cit

        timeline_contratas = {
            c: {k: self._fmt(round(v) if "media" in k else v) for k, v in data.items()}
            for c, data in sorted(timeline_raw.items())
        }

        total_timeline = {
            k: self._fmt(round(v) if "media" in k else v)
            for k, v in tot_tl.items()
        }

        kpis_marcas = {}
        for g, vals in marcas_agg.items():
            kpis_marcas[g] = {
                "bolsa": self._fmt(vals["bolsa"]),
                "rec_media": self._fmt(round(vals["rec_media"])),
                "rea_media": self._fmt(round(vals["rea_media"])),
                "rea_d1": self._fmt(vals["rea_d1"]),
                "citado": self._fmt(vals["citado"]),
                "ir": self._fmt(vals["bolsa"] / vals["rea_media"] if vals["rea_media"] > 0 else 0, 2),
                "ir_ayer": self._fmt(vals["bolsa"] / vals["rea_d1"] if vals["rea_d1"] > 0 else 0, 2),
                "ir_previsto": self._fmt(vals["bolsa"] / vals["citado"] if vals["citado"] > 0 else 0, 2)
            }

        total_global = {
            "bolsa": self._fmt(tot_bolsa_g),
            "rec_media": self._fmt(round(tot_rec_media_g)),
            "rec_d1": self._fmt(tot_tl["rec_d1"]),
            "rea_media": self._fmt(round(tot_rea_media_g)),
            "rea_d1": self._fmt(tot_rea_d1_g),
            "citado": self._fmt(tot_citado_g),
            "ir": self._fmt(tot_bolsa_g / tot_rea_media_g if tot_rea_media_g > 0 else 0, 2),
            "ir_ayer": self._fmt(tot_bolsa_g / tot_rea_d1_g if tot_rea_d1_g > 0 else 0, 2),
            "ir_previsto": self._fmt(tot_bolsa_g / tot_citado_g if tot_citado_g > 0 else 0, 2)
        }

        return {
            "timeline_contratas": timeline_contratas,
            "total_timeline": total_timeline,
            "kpis_marcas": kpis_marcas,
            "total_global": total_global
        }

    def render_email_body(self, transformed_data: Dict[str, Any], target_date: date) -> str:
        template_str = """
        <!DOCTYPE html>
        <html>
        <head><meta charset="utf-8"></head>
        <body style="margin: 0; padding: 15px; background-color: #f8fafc; font-family: sans-serif; color: #1e293b;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 20px;">
            <tr>
              <td>
                <h2 style="margin: 0; font-size: 19px; font-weight: 800; color: #ff5500; text-transform: uppercase;">
                  Entradas y Cierres
                </h2>
                <p style="margin: 4px 0 0 0; font-size: 13px; color: #64748b;">
                  Seguimiento Operativo de Órdenes Recibidas, Producción y Bolsa — MASORANGE
                </p>
              </td>
              <td align="right" valign="top">
                <span style="background-color: #0f172a; color: #ffffff; padding: 6px 14px; border-radius: 20px; font-size: 11px; font-weight: 700;">
                  Foto: {{ fecha_ayer_formateada }}
                </span>
              </td>
            </tr>
          </table>

          <!-- TARJETAS RESUMEN KPI -->
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 22px;">
            <tr>
              <td width="23.5%" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 4px solid #ff6600; border-radius: 8px; padding: 14px;">
                <div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase;">GMM (MásMóvil)</div>
                <div style="font-size: 24px; font-weight: 800; color: #0f172a; margin: 4px 0;">{{ kpis_marcas.get('MAS MOVIL', {}).get('ir', '-') }}</div>
                <div style="font-size: 11.5px; color: #475569;">Bolsa: <strong>{{ kpis_marcas.get('MAS MOVIL', {}).get('bolsa', '-') }}</strong></div>
              </td>
              <td width="2%">&nbsp;</td>
              <td width="23.5%" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 4px solid #0284c7; border-radius: 8px; padding: 14px;">
                <div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase;">OSP (Orange)</div>
                <div style="font-size: 24px; font-weight: 800; color: #0f172a; margin: 4px 0;">{{ kpis_marcas.get('ORANGE', {}).get('ir', '-') }}</div>
                <div style="font-size: 11.5px; color: #475569;">Bolsa: <strong>{{ kpis_marcas.get('ORANGE', {}).get('bolsa', '-') }}</strong></div>
              </td>
              <td width="2%">&nbsp;</td>
              <td width="23.5%" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 4px solid #6366f1; border-radius: 8px; padding: 14px;">
                <div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase;">KRTV (Euskaltel)</div>
                <div style="font-size: 24px; font-weight: 800; color: #0f172a; margin: 4px 0;">{{ kpis_marcas.get('KRTV', {}).get('ir', '-') }}</div>
                <div style="font-size: 11.5px; color: #475569;">Bolsa: <strong>{{ kpis_marcas.get('KRTV', {}).get('bolsa', '-') }}</strong></div>
              </td>
              <td width="2%">&nbsp;</td>
              <td width="23.5%" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 4px solid #0f172a; border-radius: 8px; padding: 14px;">
                <div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase;">Total Proyecto</div>
                <div style="font-size: 24px; font-weight: 800; color: #0f172a; margin: 4px 0;">{{ total_global.get('ir', '-') }}</div>
                <div style="font-size: 11.5px; color: #475569;">Bolsa Total: <strong>{{ total_global.get('bolsa', '-') }}</strong></div>
              </td>
            </tr>
          </table>

          <!-- TABLA 1: ENTRADAS Y CIERRES -->
          <div style="font-size: 13px; font-weight: 800; color: #0f172a; margin-bottom: 8px;">
            1. Entradas de Órdenes (Evolución Recibidas D-8 a D-1 por Contrata)
          </div>
          <table width="100%" cellpadding="6" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #ebdcd0; border-radius: 6px; font-size: 11px; text-align: right; margin-bottom: 24px; border-collapse: collapse;">
            <thead>
              <tr style="border-bottom: 2px solid #ffcc99; color: #475569;">
                <th align="left" style="padding: 8px 10px; color: #0f172a; font-weight: 700;">Contrata</th>
                <th style="padding: 8px 6px; font-weight: 700;">Media</th>
                <th style="padding: 8px 6px; font-weight: 700;">Max</th>
                <th style="padding: 8px 6px;">D-8</th>
                <th style="padding: 8px 6px;">D-7</th>
                <th style="padding: 8px 6px;">D-6</th>
                <th style="padding: 8px 6px;">D-5</th>
                <th style="padding: 8px 6px;">D-4</th>
                <th style="padding: 8px 6px;">D-3</th>
                <th style="padding: 8px 6px;">D-2</th>
                <th style="padding: 8px 6px; background-color: #fff9e6; font-weight: 700; color: #0f172a;">D-1</th>
              </tr>
            </thead>
            <tbody>
              {% for contrata, cdata in timeline_contratas.items() %}
              <tr style="border-bottom: 1px solid #f8eee7;">
                <td align="left" style="padding: 7px 10px; font-weight: 600; color: #1e293b;">{{ contrata }}</td>
                <td style="padding: 7px 6px; font-weight: 600;">{{ cdata.rec_media }}</td>
                <td style="padding: 7px 6px;">{{ cdata.rec_max }}</td>
                <td style="padding: 7px 6px;">{{ cdata.rec_d8 }}</td>
                <td style="padding: 7px 6px;">{{ cdata.rec_d7 }}</td>
                <td style="padding: 7px 6px;">{{ cdata.rec_d6 }}</td>
                <td style="padding: 7px 6px;">{{ cdata.rec_d5 }}</td>
                <td style="padding: 7px 6px;">{{ cdata.rec_d4 }}</td>
                <td style="padding: 7px 6px;">{{ cdata.rec_d3 }}</td>
                <td style="padding: 7px 6px;">{{ cdata.rec_d2 }}</td>
                <td style="padding: 7px 6px; background-color: #fffdf5; font-weight: 700; color: #0f172a;">{{ cdata.rec_d1 }}</td>
              </tr>
              {% endfor %}
              <tr style="background-color: #fffdf9; font-weight: 800; border-top: 2px solid #ffaa66; border-bottom: 2px solid #ffaa66; color: #0f172a;">
                <td align="left" style="padding: 9px 10px;">Total</td>
                <td style="padding: 9px 6px;">{{ total_timeline.rec_media }}</td>
                <td style="padding: 9px 6px;">{{ total_timeline.rec_max }}</td>
                <td style="padding: 9px 6px;">{{ total_timeline.rec_d8 }}</td>
                <td style="padding: 9px 6px;">{{ total_timeline.rec_d7 }}</td>
                <td style="padding: 9px 6px;">{{ total_timeline.rec_d6 }}</td>
                <td style="padding: 9px 6px;">{{ total_timeline.rec_d5 }}</td>
                <td style="padding: 9px 6px;">{{ total_timeline.rec_d4 }}</td>
                <td style="padding: 9px 6px;">{{ total_timeline.rec_d3 }}</td>
                <td style="padding: 9px 6px;">{{ total_timeline.rec_d2 }}</td>
                <td style="padding: 9px 6px; background-color: #fef08a; color: #0f172a;">{{ total_timeline.rec_d1 }}</td>
              </tr>
            </tbody>
          </table>

          <!-- TABLA 2: RESUMEN DE STOCK E IR -->
          <div style="font-size: 13px; font-weight: 800; color: #0f172a; margin-bottom: 8px;">
            2. Resumen de Stock, Bolsa e Índice de Rotación (IR)
          </div>
          <table width="100%" cellpadding="6" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 11px; text-align: right; margin-bottom: 24px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #000000; color: #ffffff;">
                <th align="left" style="padding: 8px 12px;">Grupo</th>
                <th style="padding: 8px 8px;">Rec. Media</th>
                <th style="padding: 8px 8px;">Rea. Media</th>
                <th style="padding: 8px 8px;">Rea. D-1</th>
                <th style="padding: 8px 8px;">Bolsa</th>
                <th style="background-color: #1e293b; color: #fef08a; padding: 8px 10px;">IR</th>
                <th style="padding: 8px 8px;">IR Ayer</th>
                <th style="padding: 8px 8px;">Citado D</th>
                <th style="padding: 8px 8px;">IR Previsto</th>
              </tr>
            </thead>
            <tbody>
              {% for g_nombre in ['MAS MOVIL', 'ORANGE', 'KRTV'] %}
              {% if kpis_marcas.get(g_nombre) %}
              {% set m = kpis_marcas[g_nombre] %}
              <tr style="border-bottom: 1px solid #f1f5f9; background-color: #ffffff;">
                <td align="left" style="padding: 8px 12px; font-weight: bold; color: #0f172a;">{{ g_nombre }}</td>
                <td style="padding: 8px 8px;">{{ m.rec_media }}</td>
                <td style="padding: 8px 8px;">{{ m.rea_media }}</td>
                <td style="padding: 8px 8px;">{{ m.rea_d1 }}</td>
                <td style="padding: 8px 8px; font-weight: bold;">{{ m.bolsa }}</td>
                <td style="background-color: #fff9c3; font-weight: 800; color: #0f172a; padding: 8px 10px;">{{ m.ir }}</td>
                <td style="padding: 8px 8px;">{{ m.ir_ayer }}</td>
                <td style="padding: 8px 8px;">{{ m.citado }}</td>
                <td style="padding: 8px 8px;">{{ m.ir_previsto }}</td>
              </tr>
              {% endif %}
              {% endfor %}
              <tr style="background-color: #0f172a; color: #ffffff; font-weight: 800;">
                <td align="left" style="padding: 9px 12px; color: #ffffff;">TOTAL PROYECTO</td>
                <td style="padding: 9px 8px; color: #ffffff;">{{ total_global.rec_media }}</td>
                <td style="padding: 9px 8px; color: #ffffff;">{{ total_global.rea_media }}</td>
                <td style="padding: 9px 8px; color: #ffffff;">{{ total_global.rea_d1 }}</td>
                <td style="padding: 9px 8px; color: #ffffff;">{{ total_global.bolsa }}</td>
                <td style="background-color: #fef08a; color: #0f172a; font-weight: 900; padding: 9px 10px;">{{ total_global.ir }}</td>
                <td style="padding: 9px 8px; color: #ffffff;">{{ total_global.ir_ayer }}</td>
                <td style="padding: 9px 8px; color: #ffffff;">{{ total_global.citado }}</td>
                <td style="padding: 9px 8px; color: #ffffff;">{{ total_global.ir_previsto }}</td>
              </tr>
            </tbody>
          </table>

          <p style="font-size: 11px; color: #64748b; margin-top: 18px;">
            <em>* Se adjunta el Cuadro de Mando Interactivo completo con extensión permitida <strong><code>{{ attachment_name }}</code></strong>.</em><br>
            Reporte automático de Operaciones MASORANGE — Fuente: BigQuery (<code>{{ table_name }}</code>).
          </p>
        </body>
        </html>
        """
        template = Template(template_str)
        attachment_name = f"Dashboard_IR_{target_date.strftime('%Y-%m-%d')}{self.attachment_ext}"
        return template.render(
            fecha_ayer_formateada=target_date.strftime("%d/%m/%Y"),
            timeline_contratas=transformed_data["timeline_contratas"],
            total_timeline=transformed_data["total_timeline"],
            kpis_marcas=transformed_data["kpis_marcas"],
            total_global=transformed_data["total_global"],
            attachment_name=attachment_name,
            table_name=self.table_name
        )

    def build_attachment(self, raw_rows: List[Dict[str, Any]], target_date: date) -> Tuple[str, str]:
        date_str = target_date.strftime("%Y-%m-%d")
        date_es = target_date.strftime("%d/%m/%Y")
        filename = f"Dashboard_IR_{date_str}{self.attachment_ext}"

        json_data = json.dumps(raw_rows, default=str, ensure_ascii=False)
        html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Cuadro de Mando IR - {date_es}</title>
  <style>
    body {{ font-family: sans-serif; margin: 0; padding: 14px; background: #f8fafc; font-size: 11.5px; }}
    table {{ width: 100%; border-collapse: collapse; text-align: right; background: #fff; }}
    th, td {{ padding: 6px 8px; border-bottom: 1px solid #ebdcd0; white-space: nowrap; }}
    thead th {{ background: #fff8f3; color: #ff6600; font-weight: 700; position: sticky; top: 0; }}
    .text-left {{ text-align: left; }}
    .row-total {{ background: #fffdf9; font-weight: 700; border-top: 2px solid #ffaa66; border-bottom: 2px solid #ffaa66; }}
    select, input {{ padding: 5px 8px; font-size: 11.5px; border: 1px solid #cbd5e0; border-radius: 4px; }}
  </style>
</head>
<body>
  <h2>Cuadro de Mando IR - Foto: {date_es}</h2>
  <div style="margin-bottom: 12px; display: flex; gap: 10px;">
    <select id="selGrupo" onchange="filterData()"><option value="TODOS">Todos los Grupos</option></select>
    <select id="selContrata" onchange="filterData()"><option value="TODAS">Todas las Contratas</option></select>
    <select id="selProvincia" onchange="filterData()"><option value="TODAS">Todas las Provincias</option></select>
    <input type="text" id="txtSearch" placeholder="Buscar..." oninput="filterData()">
  </div>
  <table id="tbl">
    <thead>
      <tr>
        <th class="text-left">Grupo</th>
        <th class="text-left">Contrata</th>
        <th class="text-left">Provincia</th>
        <th>Rec. Media</th><th>Rec. Max</th><th>Rec. D-2</th><th>Rec. D-1</th>
        <th>Rea. Media</th><th>Rea. Max</th><th>Rea. D-2</th><th>Rea. D-1</th>
        <th>Bolsa</th><th>IR</th><th>IR Ayer</th><th>Citado D</th>
      </tr>
    </thead>
    <tbody id="tb"></tbody>
  </table>
  <script>
    const data = {json_data};
    function fmt(n, d=0) {{ return n !== null && n !== undefined ? Number(n).toLocaleString('es-ES', {{minimumFractionDigits:d, maximumFractionDigits:d}}) : '-'; }}
    function init() {{
      const grupos = [...new Set(data.map(d => d.grupo).filter(Boolean))].sort();
      const contratas = [...new Set(data.map(d => d.contrata).filter(Boolean))].sort();
      const provincias = [...new Set(data.map(d => d.provincia).filter(Boolean))].sort();
      grupos.forEach(g => document.getElementById('selGrupo').add(new Option(g, g)));
      contratas.forEach(c => document.getElementById('selContrata').add(new Option(c, c)));
      provincias.forEach(p => document.getElementById('selProvincia').add(new Option(p, p)));
      filterData();
    }}
    function filterData() {{
      const g = document.getElementById('selGrupo').value;
      const c = document.getElementById('selContrata').value;
      const p = document.getElementById('selProvincia').value;
      const q = document.getElementById('txtSearch').value.toUpperCase();
      const rows = data.filter(d => 
        (g === 'TODOS' || d.grupo === g) &&
        (c === 'TODAS' || d.contrata === c) &&
        (p === 'TODAS' || d.provincia === p) &&
        (!q || (d.grupo+d.contrata+d.provincia).toUpperCase().includes(q))
      );
      render(rows);
    }}
    function render(rows) {{
      const tb = document.getElementById('tb');
      tb.innerHTML = '';
      let bTot = 0, reaTot = 0;
      rows.forEach(d => {{
        bTot += Number(d.bolsa || 0);
        reaTot += Number(d.rea_media || 0);
        const tr = document.createElement('tr');
        tr.innerHTML = `<td class="text-left">${{d.grupo}}</td><td class="text-left"><strong>${{d.contrata}}</strong></td><td class="text-left">${{d.provincia}}</td><td>${{fmt(d.rec_media, 1)}}</td><td>${{fmt(d.rec_max)}}</td><td>${{fmt(d.rec_d2)}}</td><td>${{fmt(d.rec_d1)}}</td><td>${{fmt(d.rea_media, 1)}}</td><td>${{fmt(d.rea_max)}}</td><td>${{fmt(d.rea_d2)}}</td><td>${{fmt(d.rea_d1)}}</td><td><strong>${{fmt(d.bolsa)}}</strong></td><td><strong>${{fmt(d.ir, 2)}}</strong></td><td>${{fmt(d.ir_ayer, 2)}}</td><td>${{fmt(d.citado)}}</td>`;
        tb.appendChild(tr);
      }});
    }}
    window.onload = init;
  </script>
</body>
</html>
"""
        return html_content, filename


# ==============================================================================
# 5. TRANSPORTE SMTP CORPORATIVO
# ==============================================================================
class SMTPEmailSender:
    """Gestiona la construcción de mensajes MIME y el envío seguro por relay SMTP."""

    def __init__(self, server: str, port: int, sender: str, use_tls: bool = False):
        self.server = server
        self.port = port
        self.sender = sender
        self.use_tls = use_tls

    def send_report(
        self,
        subject: str,
        body_html: str,
        recipients: List[str],
        cc: Optional[List[str]] = None,
        attachment: Optional[Tuple[str, str]] = None
    ):
        cc = cc or []
        logging.info(f"Preparando correo para {len(recipients)} destinatario(s)...")

        msg = MIMEMultipart("mixed")
        msg["From"] = self.sender
        msg["To"] = ", ".join(recipients)
        if cc:
            msg["Cc"] = ", ".join(cc)
        msg["Subject"] = subject

        msg.attach(MIMEText(body_html, "html", "utf-8"))

        if attachment:
            content, filename = attachment
            logging.info(f"Adjuntando archivo en memoria: '{filename}' ({len(content)} bytes)")
            part = MIMEBase("application", "octet-stream")
            part.set_payload(content.encode("utf-8"))
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        all_to = list(set(recipients + cc))
        logging.info(f"Conectando al servidor relay SMTP {self.server}:{self.port}...")

        try:
            with smtplib.SMTP(self.server, self.port, timeout=30) as smtp:
                if self.use_tls:
                    logging.info("Iniciando negociación STARTTLS...")
                    smtp.starttls()
                smtp.send_message(msg, from_addr=self.sender, to_addrs=all_to)
                logging.info("¡Correo corporativo enviado con éxito!")
        except Exception as e:
            logging.error(f"Fallo en el envío SMTP a través de {self.server}: {e}")
            raise


# ==============================================================================
# 6. ORQUESTADOR DEL PIPELINE
# ==============================================================================
class ReportPipeline:
    """Coordina el flujo completo: Extracción -> Transformación -> Renderizado -> Envío."""

    def __init__(self, config_manager: ConfigManager):
        self.cfg = config_manager.config
        self.extractor = BigQueryExtractor(self.cfg["bigquery"]["project_id"])
        self.sender = SMTPEmailSender(
            server=self.cfg["smtp"]["server"],
            port=self.cfg["smtp"]["port"],
            sender=self.cfg["smtp"]["sender"],
            use_tls=self.cfg["smtp"]["use_tls"]
        )
        self.processors: Dict[str, BaseReportProcessor] = {}
        self._register_default_processors()

    def _register_default_processors(self):
        # Registro del reporte IR oficial
        ir_proc = IRReportProcessor(
            table_name=self.cfg["bigquery"]["dataset_table_ir"],
            attachment_ext=self.cfg["reports"]["ir"]["attachment_extension"]
        )
        self.register_processor(ir_proc)

    def register_processor(self, processor: BaseReportProcessor):
        self.processors[processor.report_key] = processor

    def run(self, report_key: str, target_date: date):
        processor = self.processors.get(report_key)
        if not processor:
            raise ValueError(f"Reporte '{report_key}' no registrado. Disponibles: {list(self.processors.keys())}")

        logging.info(f"=== Iniciando Pipeline para reporte '{report_key}' (Fecha: {target_date}) ===")

        # 1. Extracción
        query = processor.build_query(target_date)
        raw_rows = self.extractor.execute_query(query)

        if not raw_rows:
            logging.warning(f"No se devolvieron registros para la fecha {target_date}. El informe se enviará vacío.")

        # 2. Transformación
        transformed = processor.transform_data(raw_rows)

        # 3. Renderizado y Adjuntos
        body_html = processor.render_email_body(transformed, target_date)
        attachment = processor.build_attachment(raw_rows, target_date)

        # 4. Envío
        subject = f"{self.cfg['reports'][report_key]['subject_prefix']} - Foto {target_date.strftime('%d/%m/%Y')}"
        self.sender.send_report(
            subject=subject,
            body_html=body_html,
            recipients=self.cfg["smtp"]["recipients"],
            cc=self.cfg["smtp"]["cc"],
            attachment=attachment
        )
        logging.info("=== Pipeline finalizado con éxito ===")


# ==============================================================================
# 7. PUNTO DE ENTRADA CLI
# ==============================================================================
def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Framework de Automatización y Envío de Reportes Operativos MASORANGE."
    )
    parser.add_argument(
        "-c", "--config",
        default="config.json",
        help="Ruta al archivo de configuración JSON (por defecto: config.json)"
    )
    parser.add_argument(
        "-r", "--report",
        default="ir",
        choices=["ir"],
        help="Tipo de reporte a procesar (por defecto: ir)"
    )
    parser.add_argument(
        "-d", "--date",
        default=None,
        help="Fecha de la foto en formato YYYY-MM-DD (por defecto: ayer)"
    )
    parser.add_argument(
        "--to",
        default=None,
        help="Lista de destinatarios separados por coma (sobrescribe config.json)"
    )
    parser.add_argument(
        "--cc",
        default=None,
        help="Lista de copias separadas por coma (sobrescribe config.json)"
    )
    parser.add_argument(
        "--smtp-server",
        default=None,
        help="Servidor relay SMTP (sobrescribe config.json)"
    )
    parser.add_argument(
        "--smtp-port",
        type=int,
        default=None,
        help="Puerto relay SMTP (sobrescribe config.json)"
    )
    parser.add_argument(
        "--log-dir",
        default=None,
        help="Directorio donde se guardarán los ficheros de log (sobrescribe config.json)"
    )
    return parser.parse_args()


def main():
    args = parse_arguments()
    config_mgr = ConfigManager(args.config)
    config_mgr.override_from_args(args)

    # Configuración de Logging y creación automática de directorio
    log_cfg = config_mgr.config.get("logging", {})
    log_dir = log_cfg.get("log_dir", r"D:\Automaticos\Logs\reports")
    log_filename = log_cfg.get("log_file", "report_automation.log")

    if log_dir:
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, log_filename)
    else:
        log_path = log_filename

    logging.basicConfig(
        level=getattr(logging, log_cfg.get("level", "INFO")),
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_path, encoding="utf-8")
        ]
    )
    logging.info(f"Sistema de logs inicializado en: {log_path}")

    # Cálculo de la fecha objetivo
    if args.date:
        target_date = date.fromisoformat(args.date)
    else:
        target_date = date.today() - timedelta(days=1)

    # Ejecución del pipeline
    pipeline = ReportPipeline(config_mgr)
    pipeline.run(report_key=args.report, target_date=target_date)


if __name__ == "__main__":
    main()
