#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Framework de Automatización de Reportes Operativos - Orange
Diseño: Programación Orientada a Objetos (POO) & Arquitectura Modular por Capas.

Capas:
  1. Setup & Logging: Configuración prioritaria con FileHandler en D:\\Automaticos\\Logs\\reports.
  2. ConfigManager: Carga configuración JSON y lee credenciales de config.ini (Talend).
  3. GraphAuth & SharePointClient: Cliente robusto Microsoft Graph con saneamiento de rutas.
  4. BigQueryExtractor: Conexión y ejecución en Google Cloud BigQuery.
  5. BaseReportProcessor (ABC): Contrato base para procesar cualquier reporte.
  6. IRReportProcessor: Especialización para IR y Entradas/Cierres con logo oficial de Orange.
  7. SMTPEmailSender: Relay corporativo SMTP (fsp.si.orange.es:587).
  8. ReportPipeline: Orquestador del ciclo de vida (BQ -> HTML -> SharePoint -> Email).

Autor: David González Cogolludo
Departamento: Operaciones Orange
"""

import os
import sys
import json
import time
import logging
import smtplib
import argparse
import configparser
import urllib.parse
from abc import ABC, abstractmethod
from datetime import date, timedelta
from typing import Dict, List, Any, Optional
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests
from google.cloud import bigquery
from jinja2 import Template


# ==============================================================================
# 1. UTILIDADES Y REINTENTOS PARA SHAREPOINT / HTTP
# ==============================================================================
def normalize_path(p: str) -> str:
    """
    Normaliza barras y elimina únicamente el prefijo de la biblioteca raíz
    ('Documentos compartidos' o 'Shared Documents') si viene incluido,
    respetando subcarpetas reales como 'Documentos/Extraciones.../IR'.
    """
    if not p:
        return ""
    p = p.replace("\\", "/").strip("/")
    while "//" in p:
        p = p.replace("//", "/")

    partes = p.split("/")
    if partes and partes[0].lower() in ("documentos compartidos", "shared documents"):
        partes = partes[1:]
    return "/".join(partes)


def encode_path_segments(p: str) -> str:
    p = normalize_path(p)
    if not p:
        return ""
    return "/".join(
        urllib.parse.quote(seg, safe="@!$&'()*,;=._- ")
        for seg in p.split("/")
    )


def build_root_item_path(folder: str, file_name: Optional[str] = None) -> str:
    folder_enc = encode_path_segments(folder)
    if file_name:
        name_enc = urllib.parse.quote(file_name, safe="@!$&'()*,;=._- ")
        path = f"{folder_enc}/{name_enc}" if folder_enc else name_enc
    else:
        path = folder_enc
    return f"/{path}:" if path else ":"


def with_retries(session: requests.Session, method: str, url: str, **kwargs) -> requests.Response:
    max_retries = 5
    backoff = 0.8
    for attempt in range(max_retries):
        try:
            resp = session.request(method, url, **kwargs)
            if resp.status_code in (429, 500, 502, 503, 504):
                raise requests.RequestException(f"HTTP {resp.status_code}: {resp.text}")
            return resp
        except Exception as e:
            sleep = backoff * (2 ** attempt)
            logging.warning(f"Reintento HTTP {attempt + 1}/{max_retries} hacia SharePoint: {e} -> Esperando {sleep:.1f}s")
            time.sleep(sleep)
    raise TimeoutError(f"Se excedió el número máximo de reintentos para: {url}")


# ==============================================================================
# 2. CLIENTE DE AUTENTICACIÓN Y SERVICIO SHAREPOINT (GRAPH API)
# ==============================================================================
class GraphAuth:
    """Gestiona la autenticación OAuth2 con Microsoft Graph (App / Client Credentials)."""

    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        self.tenant_id = tenant_id.strip("'\" \t\r\n")
        self.client_id = client_id.strip("'\" \t\r\n")
        self.client_secret = client_secret.strip("'\" \t\r\n")
        self.token_expiry = 0
        self.access_token = None

    def get_token(self) -> str:
        if self.access_token and time.time() < self.token_expiry - 60:
            return self.access_token

        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default"
        }
        resp = requests.post(url, data=data, timeout=20)
        if resp.status_code >= 400:
            logging.error(f"Error al obtener token de Azure AD: HTTP {resp.status_code} - {resp.text}")
            resp.raise_for_status()

        token_data = resp.json()
        self.access_token = token_data["access_token"]
        self.token_expiry = time.time() + token_data.get("expires_in", 3600)
        logging.info("Token de Microsoft Graph obtenido exitosamente.")
        return self.access_token


class SharePointClient:
    """Cliente para operaciones en bibliotecas de SharePoint mediante Microsoft Graph API."""

    def __init__(self, auth: GraphAuth, host: str, site_name: str):
        self.auth = auth
        self.host = host.strip("'\" \t\r\n")
        self.site_name = site_name.strip("'\" \t\r\n")
        self.session = requests.Session()
        self.site_id = None
        self._initialize_site()

    def _initialize_site(self):
        token = self.auth.get_token()
        self.session.headers.update({
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        })
        url = f"https://graph.microsoft.com/v1.0/sites/{self.host}:/sites/{self.site_name}/"
        logging.info(f"Consultando Site ID en SharePoint: {url}")
        resp = with_retries(self.session, "GET", url, timeout=20)
        resp.raise_for_status()
        self.site_id = resp.json()["id"]
        logging.info(f"Conexión con SharePoint establecida con éxito. Site ID: {self.site_id}")

    def refresh_headers(self):
        token = self.auth.get_token()
        self.session.headers.update({
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        })

    def build_url(self, folder: str = "", file_name: Optional[str] = None, suffix: str = "") -> str:
        clean_folder = normalize_path(folder)
        path = build_root_item_path(clean_folder, file_name)
        if path == ":":
            return f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root{suffix}"
        return f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root:{path}{suffix}"

    def upload_file(self, local_file_path: str, target_folder: str) -> str:
        """Sube un archivo a SharePoint y devuelve la webUrl directa generada."""
        self.refresh_headers()
        file_name = os.path.basename(local_file_path)
        file_size = os.path.getsize(local_file_path)
        clean_folder = normalize_path(target_folder)
        logging.info(f"Iniciando subida a SharePoint de '{file_name}' ({file_size:,} bytes) en carpeta '{clean_folder}'...")

        web_url = None
        if file_size > 4 * 1024 * 1024:
            # Subida por bloques (> 4MB)
            url = self.build_url(clean_folder, file_name, "/createUploadSession")
            resp = with_retries(
                self.session, "POST", url,
                json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
                timeout=30
            )
            resp.raise_for_status()
            upload_url = resp.json()["uploadUrl"]
            chunk_size = 5 * 1024 * 1024

            with open(local_file_path, "rb") as f:
                bytes_uploaded = 0
                while bytes_uploaded < file_size:
                    chunk = f.read(chunk_size)
                    headers = {
                        "Content-Length": str(len(chunk)),
                        "Content-Range": f"bytes {bytes_uploaded}-{bytes_uploaded + len(chunk) - 1}/{file_size}"
                    }
                    upload_resp = with_retries(
                        self.session, "PUT", upload_url,
                        headers=headers, data=chunk, timeout=60
                    )
                    upload_resp.raise_for_status()
                    bytes_uploaded += len(chunk)
                    logging.info(f"Progreso subida: {bytes_uploaded:,}/{file_size:,} bytes ({(bytes_uploaded / file_size) * 100:.1f}%)")
                    if bytes_uploaded >= file_size:
                        web_url = upload_resp.json().get("webUrl")
        else:
            # Subida directa en una sola petición (< 4MB)
            url = self.build_url(clean_folder, file_name, "/content")
            with open(local_file_path, "rb") as f:
                data = f.read()
            resp = with_retries(self.session, "PUT", url, data=data, timeout=45)
            resp.raise_for_status()
            web_url = resp.json().get("webUrl")

        logging.info(f"¡Archivo '{file_name}' subido exitosamente a SharePoint!")
        if web_url:
            logging.info(f"URL de acceso directo en SharePoint: {web_url}")
        return web_url or ""


# ==============================================================================
# 3. GESTOR DE CONFIGURACIÓN (CON CARGA DESDE CONFIG.INI)
# ==============================================================================
class ConfigManager:
    """Gestiona la configuración combinando defaults, config.json y config.ini."""

    DEFAULT_CONFIG = {
        "bigquery": {
            "project_id": "mo-provision-reporting",
            "dataset_table_ir": "mo-provision-reporting.INSTALLATIONS_prod.VIS_IR_prod"
        },
        "sharepoint": {
            "credentials_ini_path": r"D:\Automaticos\Talend\Sharepoint\config.ini",
            "tenant_id": "",
            "client_id": "",
            "client_secret": "",
            "host": "orangees.sharepoint.com",
            "site_name": "TRIBUPROVISIONSQUAPBUSINESSPERFORMANCE",
            "target_folder": "Documentos/Extraciones Datos Informes/Informes/IR",
            "web_url_template": "https://orangees.sharepoint.com/sites/TRIBUPROVISIONSQUAPBUSINESSPERFORMANCE/Documentos%20compartidos/Documentos/Extraciones%20Datos%20Informes/Informes/IR/{file_name}?csf=1&web=1"
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
                "send_attachment": False,
                "local_output_dir": r"D:\Automaticos\Extraciones Datos Informes\Informes\IR"
            }
        },
        "logging": {
            "level": "INFO",
            "log_dir": r"D:\Automaticos\Logs\reports",
            "log_file": "report_automation.log"
        }
    }

    def __init__(self, config_path: Optional[str] = None):
        self.config = json.loads(json.dumps(self.DEFAULT_CONFIG))

        if config_path and os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    self._deep_update(self.config, loaded)
                    logging.info(f"Configuración cargada desde '{config_path}'")
            except Exception as e:
                logging.error(f"Error al leer '{config_path}': {e}")

        self._load_credentials_from_ini()

    def _deep_update(self, base: dict, update: dict):
        for k, v in update.items():
            if isinstance(v, dict) and k in base:
                self._deep_update(base[k], v)
            else:
                base[k] = v

    def _load_credentials_from_ini(self):
        """Lee tenant_id, client_id, client_secret y site de config.ini."""
        ini_path = self.config["sharepoint"].get("credentials_ini_path", r"D:\Automaticos\Talend\Sharepoint\config.ini")
        if os.path.exists(ini_path):
            logging.info(f"Cargando credenciales de Azure AD desde: '{ini_path}'")
            try:
                ini_parser = configparser.ConfigParser()
                ini_parser.read(ini_path, encoding="utf-8")

                target_section = None
                for s in ini_parser.sections():
                    if s.lower() in ("general", "sharepoint", "default"):
                        target_section = s
                        break

                if target_section:
                    sec = ini_parser[target_section]
                    for key, val in sec.items():
                        clean_val = val.strip("'\" \t\r\n")
                        k_lower = key.lower()
                        if k_lower == "tenant_id" and clean_val:
                            self.config["sharepoint"]["tenant_id"] = clean_val
                        elif k_lower == "client_id" and clean_val:
                            self.config["sharepoint"]["client_id"] = clean_val
                        elif k_lower == "client_secret" and clean_val:
                            self.config["sharepoint"]["client_secret"] = clean_val
                        elif k_lower == "host" and clean_val:
                            self.config["sharepoint"]["host"] = clean_val
                        elif k_lower in ("site", "site_name") and clean_val:
                            self.config["sharepoint"]["site_name"] = clean_val

                    sp_c = self.config["sharepoint"]
                    if sp_c.get("tenant_id") and sp_c.get("client_id") and sp_c.get("client_secret"):
                        logging.info("Credenciales de SharePoint cargadas exitosamente desde config.ini.")
                    else:
                        logging.warning("Faltan campos obligatorios (tenant_id/client_id/client_secret) en config.ini.")
            except Exception as e:
                logging.error(f"Error al parsear '{ini_path}': {e}")
        else:
            logging.warning(f"Fichero de credenciales INI no encontrado en '{ini_path}'.")

    def override_from_args(self, args: argparse.Namespace):
        if getattr(args, "to", None):
            self.config["smtp"]["recipients"] = [e.strip() for e in args.to.split(",")]
        if getattr(args, "cc", None):
            self.config["smtp"]["cc"] = [e.strip() for e in args.cc.split(",")]
        if getattr(args, "log_dir", None):
            self.config["logging"]["log_dir"] = args.log_dir
        if getattr(args, "local_dir", None):
            self.config["reports"]["ir"]["local_output_dir"] = args.local_dir
        if getattr(args, "sp_folder", None):
            self.config["sharepoint"]["target_folder"] = args.sp_folder
        if getattr(args, "sp_ini", None):
            self.config["sharepoint"]["credentials_ini_path"] = args.sp_ini
            self._load_credentials_from_ini()


# ==============================================================================
# 4. EXTRACTOR BIGQUERY
# ==============================================================================
class BigQueryExtractor:
    """Encapsula las consultas a Google Cloud BigQuery."""

    def __init__(self, project_id: str):
        self.project_id = project_id
        self.client = bigquery.Client(project=self.project_id)

    def execute_query(self, query: str) -> List[Dict[str, Any]]:
        logging.info("Ejecutando consulta en BigQuery...")
        try:
            job = self.client.query(query)
            rows = [dict(row) for row in job.result()]
            logging.info(f"Consulta finalizada con éxito ({len(rows)} registros obtenidos).")
            return rows
        except Exception as e:
            logging.error(f"Error al ejecutar consulta en BigQuery: {e}")
            raise


# ==============================================================================
# 5. CONTRATO BASE DE PROCESADORES (POO)
# ==============================================================================
class BaseReportProcessor(ABC):
    @property
    @abstractmethod
    def report_key(self) -> str:
        pass

    @abstractmethod
    def build_query(self, target_date: date) -> str:
        pass

    @abstractmethod
    def transform_data(self, raw_rows: List[Dict[str, Any]]) -> Dict[str, Any]:
        pass

    @abstractmethod
    def generate_html_file(self, raw_rows: List[Dict[str, Any]], target_date: date, output_dir: str) -> str:
        pass

    @abstractmethod
    def render_email_body(self, transformed_data: Dict[str, Any], target_date: date, sharepoint_web_url: str) -> str:
        pass


# ==============================================================================
# 6. IMPLEMENTACIÓN: REPORTE IR CON LOGO OFICIAL ORANGE
# ==============================================================================
class IRReportProcessor(BaseReportProcessor):
    def __init__(self, table_name: str):
        self.table_name = table_name

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
          COALESCE(brand_group, 'OTROS') AS grupo,
          COALESCE(normalized_contracts, 'SIN CONTRATA') AS contrata,
          COALESCE(normalized_province, 'SIN PROVINCIA') AS provincia,

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
                k = f"rec_d{d_idx}"
                v = int(r.get(k) or 0)
                timeline_raw[c][k] += v
                tot_tl[k] += v

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

    def generate_html_file(self, raw_rows: List[Dict[str, Any]], target_date: date, output_dir: str) -> str:
        os.makedirs(output_dir, exist_ok=True)
        date_str = target_date.strftime("%Y-%m-%d")
        date_es = target_date.strftime("%d/%m/%Y")
        file_name = f"Dashboard_IR_{date_str}.html"
        output_path = os.path.join(output_dir, file_name)

        json_payload = json.dumps(raw_rows, default=str, ensure_ascii=False)
        html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Orange — Cuadro de Mando IR y Operaciones ({date_es})</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 0; padding: 14px; background: #f8fafc; font-size: 11px; color: #1e293b; }}
    .header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; border-bottom: 2px solid #ff7900; padding-bottom: 8px; }}
    .brand-container {{ display: flex; align-items: center; gap: 12px; }}
    .orange-logo-container {{ display: flex; align-items: center; gap: 8px; }}
    .orange-square {{ width: 30px; height: 30px; background-color: #ff7900; display: inline-block; }}
    .orange-wordmark {{ font-family: -apple-system, BlinkMacSystemFont, "Helvetica Neue", Arial, sans-serif; font-size: 24px; font-weight: 800; color: #000000; letter-spacing: -0.6px; }}
    .header h1 {{ margin: 0; font-size: 16px; color: #0f172a; font-weight: 800; }}
    .header h1 span {{ color: #ff7900; }}
    .filters-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 8px; margin-bottom: 12px; background: #ffffff; padding: 12px; border-radius: 6px; border: 1px solid #e2e8f0; }}
    .filter-item label {{ display: block; font-size: 10px; font-weight: 700; color: #64748b; margin-bottom: 3px; text-transform: uppercase; }}
    select, input {{ width: 100%; box-sizing: border-box; padding: 5px 8px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 11px; background: #fff; }}
    .tbl-container {{ overflow-x: auto; background: #fff; border: 1px solid #e2e8f0; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
    table {{ width: 100%; border-collapse: collapse; text-align: right; }}
    th, td {{ padding: 6px 8px; border-bottom: 1px solid #f1f5f9; white-space: nowrap; }}
    thead th {{ background: #fff8f3; color: #475569; font-weight: 700; border-bottom: 2px solid #ffcc99; }}
    .text-left {{ text-align: left; }}
    .row-total {{ background: #fffdf9; font-weight: 800; border-top: 2px solid #ffaa66; border-bottom: 2px solid #ffaa66; color: #0f172a; position: sticky; bottom: 0; }}
    .pill {{ padding: 2px 6px; border-radius: 8px; font-weight: 700; font-size: 10px; display: inline-block; }}
    .pill-green {{ background: #dcfce7; color: #15803d; }}
    .pill-amber {{ background: #fef9c3; color: #a16207; }}
    .pill-red {{ background: #fee2e2; color: #b91c1c; }}
  </style>
</head>
<body>
  <div class="header">
    <div class="brand-container">
      <div class="orange-logo-container">
        <span class="orange-square"></span>
        <span class="orange-wordmark">orange</span>
      </div>
      <h1 style="border-left: 2px solid #cbd5e1; padding-left: 12px; margin-left: 4px;">
        Cuadro de Mando IR y Operaciones ({date_es})
      </h1>
    </div>
    <span style="font-weight: 700; color: #64748b;">Orange Operaciones</span>
  </div>

  <!-- PANEL DE FILTROS ÁGIL (3 DIMENSIONES) -->
  <div class="filters-grid">
    <div class="filter-item">
      <label>Grupo</label>
      <select id="selGrupo" onchange="filterData()"><option value="TODOS">Todos</option></select>
    </div>
    <div class="filter-item">
      <label>Contrata</label>
      <select id="selContrata" onchange="filterData()"><option value="TODAS">Todas</option></select>
    </div>
    <div class="filter-item">
      <label>Provincia</label>
      <select id="selProvincia" onchange="filterData()"><option value="TODAS">Todas</option></select>
    </div>
    <div class="filter-item">
      <label>Búsqueda General</label>
      <input type="text" id="txtSearch" placeholder="Buscar por texto..." oninput="filterData()">
    </div>
  </div>

  <div class="tbl-container">
    <table id="tbl">
      <thead>
        <tr>
          <th class="text-left">Grupo</th>
          <th class="text-left">Contrata</th>
          <th class="text-left">Provincia</th>
          <th>Rec. Media</th><th>Rec. Max</th><th>Rec. D-2</th><th>Rec. D-1</th>
          <th>Rea. Media</th><th>Rea. Max</th><th>Rea. D-2</th><th>Rea. D-1</th>
          <th>Bolsa</th><th>IR</th><th>IR Ayer</th><th>IR Prev.</th>
          <th>Obj. Inst</th><th>Citado D</th><th>Cita D+1</th><th>Cita D+2</th>
          <th>% Pdte</th><th>% Posp</th>
        </tr>
      </thead>
      <tbody id="tb"></tbody>
      <tfoot><tr id="tfoot" class="row-total"></tr></tfoot>
    </table>
  </div>

  <script>
    const data = {json_payload};

    function fmt(n, d=0) {{
      return n !== null && n !== undefined && !isNaN(n)
        ? Number(n).toLocaleString('es-ES', {{minimumFractionDigits:d, maximumFractionDigits:d}})
        : '-';
    }}

    function getPill(ir) {{
      const v = parseFloat(ir);
      if (isNaN(v)) return '-';
      if (v <= 3.0) return `<span class="pill pill-green">${{fmt(v, 2)}}</span>`;
      if (v <= 4.0) return `<span class="pill pill-amber">${{fmt(v, 2)}}</span>`;
      return `<span class="pill pill-red">${{fmt(v, 2)}}</span>`;
    }}

    function populateSelect(id, field) {{
      const vals = [...new Set(data.map(d => d[field]).filter(Boolean))].sort();
      const sel = document.getElementById(id);
      vals.forEach(v => sel.add(new Option(v, v)));
    }}

    function init() {{
      populateSelect('selGrupo', 'grupo');
      populateSelect('selContrata', 'contrata');
      populateSelect('selProvincia', 'provincia');
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
        (!q || ((d.grupo || '') + (d.contrata || '') + (d.provincia || '')).toUpperCase().includes(q))
      );
      render(rows);
    }}

    function render(rows) {{
      const tb = document.getElementById('tb');
      tb.innerHTML = '';

      let sumRecM = 0, sumRecMax = 0, sumRecD2 = 0, sumRecD1 = 0;
      let sumReaM = 0, sumReaMax = 0, sumReaD2 = 0, sumReaD1 = 0;
      let sumBolsa = 0, sumCitado = 0, sumCitaD1 = 0, sumCitaD2 = 0;
      let sumPdteCita = 0, sumPospuesta = 0, sumPospuestaMM = 0;

      rows.forEach(d => {{
        sumRecM += Number(d.rec_media || 0);
        sumRecMax = Math.max(sumRecMax, Number(d.rec_max || 0));
        sumRecD2 += Number(d.rec_d2 || 0);
        sumRecD1 += Number(d.rec_d1 || 0);

        sumReaM += Number(d.rea_media || 0);
        sumReaMax = Math.max(sumReaMax, Number(d.rea_max || 0));
        sumReaD2 += Number(d.rea_d2 || 0);
        sumReaD1 += Number(d.rea_d1 || 0);

        sumBolsa += Number(d.bolsa || 0);
        sumCitado += Number(d.citado || 0);
        sumCitaD1 += Number(d.citad1 || 0);
        sumCitaD2 += Number(d.citad2 || 0);
        sumPdteCita += Number(d.pdtecita || 0);
        sumPospuesta += Number(d.pospuesta || 0);
        sumPospuestaMM += Number(d.pospuesta_mm || 0);

        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td class="text-left">${{d.grupo}}</td>
          <td class="text-left"><strong>${{d.contrata}}</strong></td>
          <td class="text-left">${{d.provincia}}</td>
          <td>${{fmt(d.rec_media, 1)}}</td><td>${{fmt(d.rec_max)}}</td><td>${{fmt(d.rec_d2)}}</td><td>${{fmt(d.rec_d1)}}</td>
          <td>${{fmt(d.rea_media, 1)}}</td><td>${{fmt(d.rea_max)}}</td><td>${{fmt(d.rea_d2)}}</td><td>${{fmt(d.rea_d1)}}</td>
          <td><strong>${{fmt(d.bolsa)}}</strong></td>
          <td>${{getPill(d.ir)}}</td>
          <td>${{getPill(d.ir_ayer)}}</td>
          <td>${{fmt(d.ir_previsto, 2)}}</td>
          <td>${{fmt(d.n_inst_bjetivo)}}</td>
          <td>${{fmt(d.citado)}}</td>
          <td>${{fmt(d.citad1)}}</td>
          <td>${{fmt(d.citad2)}}</td>
          <td>${{fmt(d.per_pdte_cita, 1)}}%</td>
          <td>${{fmt(d.per_pospuesto, 1)}}%</td>
        `;
        tb.appendChild(tr);
      }});

      const irTot = sumReaM > 0 ? (sumBolsa / sumReaM) : 0;
      const irAyerTot = sumReaD1 > 0 ? (sumBolsa / sumReaD1) : 0;
      const irPrevTot = sumCitado > 0 ? (sumBolsa / sumCitado) : 0;
      const nInstObjTot = Math.round(sumBolsa / 2.5);
      const perPdteTot = sumBolsa > 0 ? (100 * sumPdteCita / sumBolsa) : 0;
      const perPospTot = (sumBolsa + sumPospuestaMM) > 0 ? (100 * sumPospuesta / (sumBolsa + sumPospuestaMM)) : 0;

      document.getElementById('tfoot').innerHTML = `
        <td class="text-left" colspan="3">TOTAL FILTRADO (${{rows.length}} registros)</td>
        <td>${{fmt(sumRecM, 0)}}</td><td>${{fmt(sumRecMax)}}</td><td>${{fmt(sumRecD2)}}</td><td>${{fmt(sumRecD1)}}</td>
        <td>${{fmt(sumReaM, 0)}}</td><td>${{fmt(sumReaMax)}}</td><td>${{fmt(sumReaD2)}}</td><td>${{fmt(sumReaD1)}}</td>
        <td><strong>${{fmt(sumBolsa)}}</strong></td>
        <td>${{getPill(irTot)}}</td>
        <td>${{getPill(irAyerTot)}}</td>
        <td>${{fmt(irPrevTot, 2)}}</td>
        <td>${{fmt(nInstObjTot)}}</td>
        <td>${{fmt(sumCitado)}}</td>
        <td>${{fmt(sumCitaD1)}}</td>
        <td>${{fmt(sumCitaD2)}}</td>
        <td>${{fmt(perPdteTot, 1)}}%</td>
        <td>${{fmt(perPospTot, 1)}}%</td>
      `;
    }}

    window.onload = init;
  </script>
</body>
</html>
"""
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        logging.info(f"Dashboard interactivo HTML generado en disco: '{output_path}' ({len(html_content):,} bytes)")
        return output_path

    def render_email_body(self, transformed_data: Dict[str, Any], target_date: date, sharepoint_web_url: str) -> str:
        template_str = """
        <!DOCTYPE html>
        <html>
        <head><meta charset="utf-8"></head>
        <body style="margin: 0; padding: 15px; background-color: #f8fafc; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; color: #1e293b;">
          <!-- CABECERA PRINCIPAL CON LOGOTIPO OFICIAL DE ORANGE -->
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 18px;">
            <tr>
              <td valign="middle">
                <!-- LOGO OFICIAL ORANGE (CUADRADO SÓLIDO + WORDMARK) -->
                <table cellpadding="0" cellspacing="0" border="0">
                  <tr>
                    <td style="width: 32px; height: 32px; background-color: #ff7900; line-height: 0; font-size: 0;" valign="middle">
                      &nbsp;
                    </td>
                    <td style="padding-left: 10px; vertical-align: middle;">
                      <span style="font-family: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', Arial, sans-serif; font-size: 26px; font-weight: 800; color: #000000; letter-spacing: -0.6px; line-height: 1;">orange</span>
                    </td>
                  </tr>
                </table>
                <h2 style="margin: 10px 0 0 0; font-size: 19px; font-weight: 800; color: #0f172a; text-transform: uppercase;">
                  Entradas y Cierres
                </h2>
                <p style="margin: 4px 0 0 0; font-size: 13px; color: #64748b;">
                  Seguimiento Operativo de Órdenes Recibidas, Producción y Bolsa — Orange Operaciones
                </p>
              </td>
              <td align="right" valign="top">
                <span style="background-color: #0f172a; color: #ffffff; padding: 6px 14px; border-radius: 20px; font-size: 11px; font-weight: 700;">
                  Foto: {{ fecha_ayer_formateada }}
                </span>
              </td>
            </tr>
          </table>

          <!-- BANNER ACCESO DIRECTO A SHAREPOINT (SIN ADJUNTOS) -->
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 20px;">
            <tr>
              <td style="background-color: #eff6ff; border: 1px solid #bfdbfe; border-left: 5px solid #ff7900; border-radius: 6px; padding: 14px 18px;">
                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                  <tr>
                    <td valign="middle">
                      <span style="font-size: 13.5px; font-weight: 800; color: #0f172a;">🌐 Cuadro de Mando Actualizado en SharePoint</span>
                      <p style="margin: 4px 0 0 0; font-size: 12px; color: #334155;">
                        El informe interactivo con filtros por Grupo, Contrata y Provincia ya está publicado y disponible:
                      </p>
                    </td>
                    <td align="right" valign="middle" style="white-space: nowrap; padding-left: 15px;">
                      <a href="{{ sharepoint_web_url }}" target="_blank" style="background-color: #ff7900; color: #ffffff; text-decoration: none; padding: 9px 18px; border-radius: 4px; font-size: 12px; font-weight: 700; display: inline-block;">
                        Abrir Dashboard en SharePoint &rarr;
                      </a>
                    </td>
                  </tr>
                </table>
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
              <td width="23.5%" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 4px solid #ff7900; border-radius: 8px; padding: 14px;">
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
              <tr style="background-color: #fffdf9; font-weight: 800; border-top: 2px solid #ff7900; border-bottom: 2px solid #ff7900; color: #0f172a;">
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
            <em>* Informe sin adjuntos: para evitar bloqueos perimetrales, el archivo HTML interactivo con filtros detallados se encuentra almacenado y accesible directamente en <a href="{{ sharepoint_web_url }}" style="color: #ff7900; text-decoration: underline;">SharePoint Orange</a>.</em><br>
            Reporte automático de Operaciones Orange — Fuente: BigQuery (<code>{{ table_name }}</code>).
          </p>
        </body>
        </html>
        """
        template = Template(template_str)
        return template.render(
            fecha_ayer_formateada=target_date.strftime("%d/%m/%Y"),
            timeline_contratas=transformed_data["timeline_contratas"],
            total_timeline=transformed_data["total_timeline"],
            kpis_marcas=transformed_data["kpis_marcas"],
            total_global=transformed_data["total_global"],
            table_name=self.table_name,
            sharepoint_web_url=sharepoint_web_url
        )


# ==============================================================================
# 7. TRANSPORTE SMTP CORPORATIVO
# ==============================================================================
class SMTPEmailSender:
    """Envío de correo mediante el relay interno de Orange."""

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
        cc: Optional[List[str]] = None
    ):
        cc = cc or []
        logging.info(f"Preparando correo hacia: {', '.join(recipients)} (CC: {', '.join(cc)})...")

        msg = MIMEMultipart("alternative")
        msg["From"] = self.sender
        msg["To"] = ", ".join(recipients)
        if cc:
            msg["Cc"] = ", ".join(cc)
        msg["Subject"] = subject
        msg.attach(MIMEText(body_html, "html", "utf-8"))

        to_addrs = list(set([e.strip() for e in (recipients + cc) if e.strip()]))
        logging.info(f"Conectando al relay SMTP {self.server}:{self.port}...")

        try:
            with smtplib.SMTP(self.server, self.port, timeout=30) as smtp:
                if self.use_tls:
                    logging.info("Negociando STARTTLS...")
                    smtp.starttls()
                smtp.send_message(msg, from_addr=self.sender, to_addrs=to_addrs)
                logging.info("¡Correo informativo enviado exitosamente sin adjuntos!")
        except Exception as e:
            logging.error(f"Error al enviar mensaje por SMTP: {e}")
            raise


# ==============================================================================
# 8. ORQUESTADOR DEL PIPELINE GENERAL
# ==============================================================================
class ReportPipeline:
    """Coordina el flujo completo: BigQuery -> HTML Local -> SharePoint Upload -> Email."""

    def __init__(self, config_manager: ConfigManager):
        self.cfg = config_manager.config
        self.extractor = BigQueryExtractor(self.cfg["bigquery"]["project_id"])
        self.sender = SMTPEmailSender(
            server=self.cfg["smtp"]["server"],
            port=self.cfg["smtp"]["port"],
            sender=self.cfg["smtp"]["sender"],
            use_tls=self.cfg["smtp"]["use_tls"]
        )

        sp_cfg = self.cfg["sharepoint"]
        self.sp_client = None
        has_tenant = bool(sp_cfg.get("tenant_id"))
        has_client = bool(sp_cfg.get("client_id"))
        has_secret = bool(sp_cfg.get("client_secret"))

        if has_tenant and has_client and has_secret:
            try:
                auth = GraphAuth(
                    tenant_id=sp_cfg["tenant_id"],
                    client_id=sp_cfg["client_id"],
                    client_secret=sp_cfg["client_secret"]
                )
                self.sp_client = SharePointClient(
                    auth=auth,
                    host=sp_cfg["host"],
                    site_name=sp_cfg["site_name"]
                )
            except Exception as e:
                logging.exception(f"Fallo al inicializar SharePointClient: {e}")
        else:
            logging.warning(
                f"Credenciales de SharePoint incompletas (tenant={has_tenant}, client={has_client}, secret={has_secret}). "
                f"Comprueba '{sp_cfg.get('credentials_ini_path')}'."
            )

        self.processors: Dict[str, BaseReportProcessor] = {}
        self._register_default_processors()

    def _register_default_processors(self):
        self.processors["ir"] = IRReportProcessor(
            table_name=self.cfg["bigquery"]["dataset_table_ir"]
        )

    def run(self, report_key: str, target_date: date, skip_email: bool = False):
        processor = self.processors.get(report_key)
        if not processor:
            raise ValueError(f"Reporte '{report_key}' desconocido.")

        logging.info(f"=== INICIANDO PIPELINE: {report_key.upper()} (Fecha objetivo: {target_date}) ===")

        # 1. Extracción BigQuery con las 3 dimensiones consolidadas (Grupo, Contrata, Provincia)
        query = processor.build_query(target_date)
        raw_rows = self.extractor.execute_query(query)

        # 2. Transformación de datos para el cuerpo del correo
        transformed = processor.transform_data(raw_rows)

        # 3. Generación del archivo HTML interactivo en disco con logo oficial de Orange
        out_dir = self.cfg["reports"][report_key].get(
            "local_output_dir",
            r"D:\Automaticos\Extraciones Datos Informes\Informes\IR"
        )
        local_html_path = processor.generate_html_file(raw_rows, target_date, out_dir)
        file_name = os.path.basename(local_html_path)

        # 4. Subida a SharePoint y captura de webUrl directa
        sp_folder = self.cfg["sharepoint"].get(
            "target_folder",
            "Documentos/Extraciones Datos Informes/Informes/IR"
        )
        uploaded_url = None
        if self.sp_client:
            try:
                uploaded_url = self.sp_client.upload_file(local_html_path, sp_folder)
            except Exception as e:
                logging.exception(f"Error durante la subida a SharePoint de '{file_name}': {e}")
        else:
            logging.warning("Se omitió la subida a SharePoint porque el cliente Graph no está disponible.")

        if uploaded_url:
            sharepoint_web_url = uploaded_url
        else:
            web_template = self.cfg["sharepoint"].get("web_url_template", "")
            sharepoint_web_url = web_template.format(file_name=file_name)

        # 5. Envío de correo informativo (sin adjuntos)
        if skip_email:
            logging.info("Opción --no-email activada. Se omite el envío por SMTP.")
        else:
            body_html = processor.render_email_body(transformed, target_date, sharepoint_web_url)
            subject = f"{self.cfg['reports'][report_key]['subject_prefix']} - Foto {target_date.strftime('%d/%m/%Y')}"
            self.sender.send_report(
                subject=subject,
                body_html=body_html,
                recipients=self.cfg["smtp"]["recipients"],
                cc=self.cfg["smtp"]["cc"]
            )

        logging.info(f"=== PIPELINE FINALIZADO CON ÉXITO: {file_name} ===")


# ==============================================================================
# 9. CLI Y PUNTO DE ENTRADA
# ==============================================================================
def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Automatización BigQuery -> HTML -> Subida a SharePoint -> Correo Orange"
    )
    parser.add_argument("-c", "--config", default="config.json", help="Ruta al config.json")
    parser.add_argument("-r", "--report", default="ir", choices=["ir"], help="Reporte a procesar")
    parser.add_argument("-d", "--date", default=None, help="Fecha en formato YYYY-MM-DD (por defecto: ayer)")
    parser.add_argument("--to", default=None, help="Destinatarios separados por coma")
    parser.add_argument("--cc", default=None, help="Copias separadas por coma")
    parser.add_argument("--log-dir", default=None, help="Directorio de logs")
    parser.add_argument("--local-dir", default=None, help="Directorio local de salida HTML")
    parser.add_argument("--sp-folder", default=None, help="Carpeta destino en SharePoint")
    parser.add_argument("--sp-ini", default=None, help="Ruta al archivo config.ini de SharePoint (Talend)")
    parser.add_argument("--no-email", action="store_true", help="Generar y subir a SharePoint sin enviar correo")
    return parser.parse_args()


def init_logging(log_dir: str, log_file: str, level_name: str = "INFO") -> str:
    """Configura el logging inmediatamente forzando handlers tanto a consola como a disco."""
    if log_dir:
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, log_file)
    else:
        log_path = log_file

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    stream_handler = logging.StreamHandler(sys.stdout)

    logging.basicConfig(
        level=getattr(logging, level_name.upper(), logging.INFO),
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[stream_handler, file_handler],
        force=True
    )
    logging.info(f"Logging corporativo inicializado. Registrando en: '{log_path}'")
    return log_path


def main():
    args = parse_arguments()

    default_log_dir = r"D:\Automaticos\Logs\reports"
    default_log_file = "report_automation.log"

    if args.config and os.path.exists(args.config):
        try:
            with open(args.config, "r", encoding="utf-8") as f:
                pre_cfg = json.load(f)
                log_sec = pre_cfg.get("logging", {})
                default_log_dir = log_sec.get("log_dir", default_log_dir)
                default_log_file = log_sec.get("log_file", default_log_file)
        except Exception:
            pass

    log_dir = args.log_dir or default_log_dir
    init_logging(log_dir, default_log_file)

    config_mgr = ConfigManager(args.config)
    config_mgr.override_from_args(args)

    if args.date:
        target_date = date.fromisoformat(args.date)
    else:
        target_date = date.today() - timedelta(days=1)

    pipeline = ReportPipeline(config_mgr)
    pipeline.run(report_key=args.report, target_date=target_date, skip_email=args.no_email)


if __name__ == "__main__":
    main()