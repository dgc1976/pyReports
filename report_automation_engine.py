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
  6. IRReportProcessor: Especialización para IR, Resumen IR Segmentado, CDM por Actividad
                        (Toda la Actividad, AyT Masivo, Migraciones, B2B) y CDM Provincial.
  7. SMTPEmailSender: Relay corporativo SMTP (fsp.si.orange.es:587).
  8. ReportPipeline: Orquestador del ciclo de vida (BQ -> HTML -> SharePoint -> Email compacto).

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
                        logging.warning("Faltan campos obligatorios en config.ini.")
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
    def render_email_body(
        self,
        transformed_data: Dict[str, Any],
        target_date: date,
        sharepoint_web_url: str
    ) -> str:
        pass


# ==============================================================================
# 6. IMPLEMENTACIÓN: REPORTE IR COMPACTO CON TODOS LOS CDMS DE LOOKER
# ==============================================================================
class IRReportProcessor(BaseReportProcessor):
    EXCLUDED_CONTRATAS = {"FOL", "ICONO", "NO_DEFINIDO", "OTROS", "SIN CONTRATA"}

    def __init__(self, table_name: str):
        self.table_name = table_name

    @property
    def report_key(self) -> str:
        return "ir"

    @staticmethod
    def _fmt(val, dec=0):
        if val is None or val == "" or str(val).lower() in ("none", "null", "nan"):
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
            -- Banderas de Segmentación Operativa de Looker
            CASE WHEN service_operation IN ('ALTA', 'TRASLADO', 'MIGRACION', 'SWAP', 'MIGRACION ACOMETIDA') THEN TRUE ELSE FALSE END AS Filtro_NORMALIZADO,
            CASE WHEN service_operation NOT IN ('SWAP') THEN TRUE ELSE FALSE END AS Filtro_NoSwap,
            CASE WHEN service_operation IN ('ALTA', 'TRASLADO') AND activity_group = 'MASIVO' THEN TRUE ELSE FALSE END AS Filtro_AyTMasivo,
            CASE WHEN service_operation IN ('MIGRACION', 'MIGRACION ACOMETIDA') AND activity_group = 'MASIVO' THEN TRUE ELSE FALSE END AS Filtro_MigMasivo,
            CASE WHEN activity_group = 'B2B' THEN TRUE ELSE FALSE END AS Filtro_B2B,

            CASE WHEN tipo = 'Rec' THEN total_3_dias ELSE 0 END AS MediaRecibidas,
            CASE WHEN tipo = 'Real' THEN total_3_dias ELSE 0 END AS MediaRealizadas,
            CASE WHEN tipo = 'UltDiaReal' THEN total_3_dias ELSE 0 END AS RealUltDiaHabil,

            -- Serie de Recibidas (D-8 a D-1)
            CASE WHEN dias_atras = 1 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_1,
            CASE WHEN dias_atras = 2 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_2,
            CASE WHEN dias_atras = 3 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_3,
            CASE WHEN dias_atras = 4 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_4,
            CASE WHEN dias_atras = 5 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_5,
            CASE WHEN dias_atras = 6 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_6,
            CASE WHEN dias_atras = 7 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_7,
            CASE WHEN dias_atras = 8 THEN total_is_ot_created ELSE 0 END AS Recibidas_D_8,

            -- Serie de Realizadas (D-8 a D-1)
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
            AND UPPER(TRIM(normalized_contracts)) NOT IN ('FOL', 'ICONO', 'NO_DEFINIDO', 'OTROS', 'SIN CONTRATA')
            AND normalized_contracts IS NOT NULL
        )
        SELECT 
          COALESCE(brand_group, 'OTROS') AS grupo,
          COALESCE(normalized_contracts, 'SIN CONTRATA') AS contrata,
          COALESCE(normalized_province, 'SIN PROVINCIA') AS provincia,

          -- Totales de flujo y medias generales
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
          SUM(TotalPospuestaMM_KRTV) AS pospuesta_mm,

          -- Segmentaciones para la tabla «Resumen IR»:
          SUM(CASE WHEN Filtro_NORMALIZADO THEN total_bolsa ELSE 0 END) AS bolsa_norm,
          SUM(CASE WHEN Filtro_NORMALIZADO AND tipo = 'Real' THEN total_3_dias ELSE 0 END) AS real_3dias_norm,
          SUM(CASE WHEN Filtro_AyTMasivo THEN total_bolsa ELSE 0 END) AS bolsa_ayt,
          SUM(CASE WHEN Filtro_AyTMasivo AND tipo = 'Real' THEN total_3_dias ELSE 0 END) AS real_3dias_ayt,
          SUM(CASE WHEN Filtro_MigMasivo THEN total_bolsa ELSE 0 END) AS bolsa_mig,
          SUM(CASE WHEN Filtro_MigMasivo AND tipo = 'Real' THEN total_3_dias ELSE 0 END) AS real_3dias_mig,
          SUM(CASE WHEN Filtro_B2B THEN total_bolsa ELSE 0 END) AS bolsa_b2b,
          SUM(CASE WHEN Filtro_B2B AND tipo = 'Real' THEN total_3_dias ELSE 0 END) AS real_3dias_b2b,

          -- 1. CDM: Toda la Actividad (menos SWAP)
          ROUND(SUM(CASE WHEN Filtro_NoSwap THEN MediaRecibidas ELSE 0 END) / 3, 1) AS rec_media_noswap,
          MAX(CASE WHEN Filtro_NoSwap THEN max_ots ELSE 0 END) AS rec_max_noswap,
          MAX(CASE WHEN Filtro_NoSwap THEN Recibidas_D_1 ELSE 0 END) AS rec_d1_noswap,
          ROUND(SUM(CASE WHEN Filtro_NoSwap THEN MediaRealizadas ELSE 0 END) / 3, 1) AS rea_media_noswap,
          MAX(CASE WHEN Filtro_NoSwap THEN max_ots ELSE 0 END) AS rea_max_noswap,
          MAX(CASE WHEN Filtro_NoSwap THEN Realizadas_D_1 ELSE 0 END) AS rea_d1_noswap,
          SUM(CASE WHEN Filtro_NoSwap THEN total_bolsa ELSE 0 END) AS bolsa_noswap,
          SUM(CASE WHEN Filtro_NoSwap THEN total_citaD ELSE 0 END) AS citado_noswap,
          SUM(CASE WHEN Filtro_NoSwap THEN total_citaD1 ELSE 0 END) AS citad1_noswap,
          SUM(CASE WHEN Filtro_NoSwap THEN total_pdtecita ELSE 0 END) AS pdtecita_noswap,

          -- 2. CDM: Altas y Traslados (Masivo)
          ROUND(SUM(CASE WHEN Filtro_AyTMasivo THEN MediaRecibidas ELSE 0 END) / 3, 1) AS rec_media_ayt,
          MAX(CASE WHEN Filtro_AyTMasivo THEN max_ots ELSE 0 END) AS rec_max_ayt,
          MAX(CASE WHEN Filtro_AyTMasivo THEN Recibidas_D_1 ELSE 0 END) AS rec_d1_ayt,
          ROUND(SUM(CASE WHEN Filtro_AyTMasivo THEN MediaRealizadas ELSE 0 END) / 3, 1) AS rea_media_ayt,
          MAX(CASE WHEN Filtro_AyTMasivo THEN max_ots ELSE 0 END) AS rea_max_ayt,
          MAX(CASE WHEN Filtro_AyTMasivo THEN Realizadas_D_1 ELSE 0 END) AS rea_d1_ayt,
          SUM(CASE WHEN Filtro_AyTMasivo THEN total_bolsa ELSE 0 END) AS bolsa_ayt_cdm,
          SUM(CASE WHEN Filtro_AyTMasivo THEN total_citaD ELSE 0 END) AS citado_ayt,
          SUM(CASE WHEN Filtro_AyTMasivo THEN total_pdtecita ELSE 0 END) AS pdtecita_ayt,

          -- 3. CDM: Solo Migraciones (Masivo)
          ROUND(SUM(CASE WHEN Filtro_MigMasivo THEN MediaRecibidas ELSE 0 END) / 3, 1) AS rec_media_mig,
          MAX(CASE WHEN Filtro_MigMasivo THEN max_ots ELSE 0 END) AS rec_max_mig,
          MAX(CASE WHEN Filtro_MigMasivo THEN Recibidas_D_1 ELSE 0 END) AS rec_d1_mig,
          ROUND(SUM(CASE WHEN Filtro_MigMasivo THEN MediaRealizadas ELSE 0 END) / 3, 1) AS rea_media_mig,
          MAX(CASE WHEN Filtro_MigMasivo THEN max_ots ELSE 0 END) AS rea_max_mig,
          MAX(CASE WHEN Filtro_MigMasivo THEN Realizadas_D_1 ELSE 0 END) AS rea_d1_mig,
          SUM(CASE WHEN Filtro_MigMasivo THEN total_bolsa ELSE 0 END) AS bolsa_mig_cdm,
          SUM(CASE WHEN Filtro_MigMasivo THEN total_citaD ELSE 0 END) AS citado_mig,
          SUM(CASE WHEN Filtro_MigMasivo THEN total_pdtecita ELSE 0 END) AS pdtecita_mig,

          -- 4. CDM: Solo B2B
          ROUND(SUM(CASE WHEN Filtro_B2B THEN MediaRecibidas ELSE 0 END) / 3, 1) AS rec_media_b2b,
          MAX(CASE WHEN Filtro_B2B THEN max_ots ELSE 0 END) AS rec_max_b2b,
          MAX(CASE WHEN Filtro_B2B THEN Recibidas_D_1 ELSE 0 END) AS rec_d1_b2b,
          ROUND(SUM(CASE WHEN Filtro_B2B THEN MediaRealizadas ELSE 0 END) / 3, 1) AS rea_media_b2b,
          MAX(CASE WHEN Filtro_B2B THEN max_ots ELSE 0 END) AS rea_max_b2b,
          MAX(CASE WHEN Filtro_B2B THEN Realizadas_D_1 ELSE 0 END) AS rea_d1_b2b,
          SUM(CASE WHEN Filtro_B2B THEN total_bolsa ELSE 0 END) AS bolsa_b2b_cdm,
          SUM(CASE WHEN Filtro_B2B THEN total_citaD ELSE 0 END) AS citado_b2b,
          SUM(CASE WHEN Filtro_B2B THEN total_pdtecita ELSE 0 END) AS pdtecita_b2b

        FROM TOTAL_IR
        GROUP BY grupo, contrata, provincia
        ORDER BY grupo, contrata, provincia;
        """

    def _build_segment_table(self, rows_by_contrata: Dict[str, Dict[str, Any]], key_prefix: str, use_citad1: bool = False) -> Dict[str, Any]:
        """Construye una tabla CDM segmentada idéntica a las capturas de Looker Studio."""
        items = {}
        tot = {
            "rec_media": 0.0, "rec_max": 0, "rec_d1": 0,
            "rea_media": 0.0, "rea_max": 0, "rea_d1": 0,
            "bolsa": 0, "citado": 0, "pdtecita": 0
        }

        for c, d in sorted(rows_by_contrata.items()):
            rec_m = d[f"rec_media_{key_prefix}"]
            rec_max = d[f"rec_max_{key_prefix}"]
            rec_1 = d[f"rec_d1_{key_prefix}"]
            rea_m = d[f"rea_media_{key_prefix}"]
            rea_max = d[f"rea_max_{key_prefix}"]
            rea_1 = d[f"rea_d1_{key_prefix}"]
            b = d[f"bolsa_{key_prefix}"]
            cit = d[f"citad1_{key_prefix}" if use_citad1 else f"citado_{key_prefix}"]
            p_cita = d[f"pdtecita_{key_prefix}"]

            tot["rec_media"] += rec_m
            tot["rec_max"] = max(tot["rec_max"], rec_max)
            tot["rec_d1"] += rec_1
            tot["rea_media"] += rea_m
            tot["rea_max"] = max(tot["rea_max"], rea_max)
            tot["rea_d1"] += rea_1
            tot["bolsa"] += b
            tot["citado"] += cit
            tot["pdtecita"] += p_cita

            ir_3d = (b / rea_m) if rea_m > 0 else None
            ir_ayer = (b / rea_1) if rea_1 > 0 else None
            ir_prev = (b / cit) if cit > 0 else None
            per_pdte = (100 * p_cita / b) if b > 0 else None

            items[c] = {
                "rec_media": self._fmt(round(rec_m)),
                "rec_max": self._fmt(rec_max),
                "rec_d1": self._fmt(rec_1),
                "rea_media": self._fmt(round(rea_m)),
                "rea_max": self._fmt(rea_max),
                "rea_d1": self._fmt(rea_1),
                "bolsa": self._fmt(b),
                "ir_3d": self._fmt(ir_3d, 1),
                "ir_ayer": self._fmt(ir_ayer, 1),
                "ir_previsto": self._fmt(ir_prev, 2),
                "pdtecita": self._fmt(p_cita),
                "per_pdte_cita": f"{self._fmt(per_pdte, 1)}%" if per_pdte is not None else "-",
                "cita_col": self._fmt(cit)
            }

        tot_ir_3d = (tot["bolsa"] / tot["rea_media"]) if tot["rea_media"] > 0 else 0
        tot_ir_ayer = (tot["bolsa"] / tot["rea_d1"]) if tot["rea_d1"] > 0 else 0
        tot_ir_prev = (tot["bolsa"] / tot["citado"]) if tot["citado"] > 0 else 0
        tot_per_pdte = (100 * tot["pdtecita"] / tot["bolsa"]) if tot["bolsa"] > 0 else 0

        total_row = {
            "rec_media": self._fmt(round(tot["rec_media"])),
            "rec_max": self._fmt(tot["rec_max"]),
            "rec_d1": self._fmt(tot["rec_d1"]),
            "rea_media": self._fmt(round(tot["rea_media"])),
            "rea_max": self._fmt(tot["rea_max"]),
            "rea_d1": self._fmt(tot["rea_d1"]),
            "bolsa": self._fmt(tot["bolsa"]),
            "ir_3d": self._fmt(tot_ir_3d, 1),
            "ir_ayer": self._fmt(tot_ir_ayer, 1),
            "ir_previsto": self._fmt(tot_ir_prev, 2),
            "pdtecita": self._fmt(tot["pdtecita"]),
            "per_pdte_cita": f"{self._fmt(tot_per_pdte, 1)}%",
            "cita_col": self._fmt(tot["citado"])
        }

        return {"filas": items, "total": total_row}

    def transform_data(self, raw_rows: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Calcula las estructuras completas de Looker Studio excluyendo contratas inactivas:
          - Elimina FOL, ICONO, NO_DEFINIDO, OTROS.
          - Incorpora los nuevos CDMs segmentados (Toda la Actividad, AyT, Migraciones, B2B).
          - Incorpora la lista detallada CDM por Grupo Marcas, Contrata y Provincia.
        """
        ir_raw = {}
        tot_ir = {
            "rec_media": 0.0, "rec_max": 0, "rec_d2": 0, "rec_d1": 0,
            "rea_media": 0.0, "rea_max": 0, "rea_d2": 0, "rea_d1": 0,
            "bolsa": 0, "citado": 0, "real_ayer": 0
        }

        resumen_ir_raw = {}
        tot_resumen_ir = {
            "bolsa_norm": 0, "real_3dias_norm": 0.0,
            "bolsa_global": 0, "real_3dias_global": 0.0,
            "bolsa_ayt": 0, "real_3dias_ayt": 0.0,
            "bolsa_mig": 0, "real_3dias_mig": 0.0,
            "bolsa_b2b": 0, "real_3dias_b2b": 0.0
        }

        entradas_raw = {}
        tot_entradas = {k: 0.0 if "media" in k else 0 for k in [
            "rec_media", "rec_max", "rec_d8", "rec_d7", "rec_d6", "rec_d5",
            "rec_d4", "rec_d3", "rec_d2", "rec_d1"
        ]}

        cierres_raw = {}
        tot_cierres = {k: 0.0 if "media" in k else 0 for k in [
            "rea_media", "rea_max", "rea_d8", "rea_d7", "rea_d6", "rea_d5",
            "rea_d4", "rea_d3", "rea_d2", "rea_d1"
        ]}

        citas_raw = {}
        tot_citas = {
            "bolsa": 0, "citado": 0, "citad1": 0, "citad2": 0,
            "pdtecita": 0, "pospuesta": 0, "pospuesta_mm": 0
        }

        marcas_agg = {}

        # Colección para el Cuadro de Mando Detallado (CDM)
        cdm_filas = []
        tot_cdm = {
            "rec_media": 0.0, "rec_max": 0, "rec_d1": 0,
            "rea_media": 0.0, "rea_max": 0, "rea_d1": 0,
            "bolsa": 0, "citad1": 0
        }

        # Acumuladores de contratas para los CDMs segmentados
        segment_contratas_raw = {}

        tot_bolsa_g = 0
        tot_rec_media_g = 0.0
        tot_rea_media_g = 0.0
        tot_rea_d1_g = 0
        tot_citado_g = 0
        tot_citad1_g = 0
        tot_pdtecita_g = 0
        tot_pospuesta_g = 0
        tot_pospuesta_mm_g = 0

        for r in raw_rows:
            c = str(r.get("contrata") or "").strip().upper()
            if not c or c in self.EXCLUDED_CONTRATAS:
                continue

            g = str(r.get("grupo") or "OTROS").strip()
            p = str(r.get("provincia") or "SIN PROVINCIA").strip()

            b = int(r.get("bolsa") or 0)
            rec_m = float(r.get("rec_media") or 0.0)
            rec_max = int(r.get("rec_max") or 0)
            rec_2 = int(r.get("rec_d2") or 0)
            rec_1 = int(r.get("rec_d1") or 0)

            rea_m = float(r.get("rea_media") or 0.0)
            rea_max = int(r.get("rea_max") or 0)
            rea_2 = int(r.get("rea_d2") or 0)
            rea_1 = int(r.get("rea_d1") or 0)

            cit = int(r.get("citado") or 0)
            cit1 = int(r.get("citad1") or 0)
            cit2 = int(r.get("citad2") or 0)
            p_cita = int(r.get("pdtecita") or 0)
            posp = int(r.get("pospuesta") or 0)
            posp_mm = int(r.get("pospuesta_mm") or 0)

            b_norm = int(r.get("bolsa_norm") or 0)
            r3_norm = float(r.get("real_3dias_norm") or 0.0)
            b_ayt = int(r.get("bolsa_ayt") or 0)
            r3_ayt = float(r.get("real_3dias_ayt") or 0.0)
            b_mig = int(r.get("bolsa_mig") or 0)
            r3_mig = float(r.get("real_3dias_mig") or 0.0)
            b_b2b = int(r.get("bolsa_b2b") or 0)
            r3_b2b = float(r.get("real_3dias_b2b") or 0.0)

            tot_bolsa_g += b
            tot_rec_media_g += rec_m
            tot_rea_media_g += rea_m
            tot_rea_d1_g += rea_1
            tot_citado_g += cit
            tot_citad1_g += cit1
            tot_pdtecita_g += p_cita
            tot_pospuesta_g += posp
            tot_pospuesta_mm_g += posp_mm

            # Fila individual CDM detallada
            ir_row = (b / rea_m) if rea_m > 0 else None
            ir_ayer_row = (b / rea_1) if rea_1 > 0 else None

            cdm_filas.append({
                "grupo": g,
                "contrata": c,
                "provincia": p,
                "rec_media": self._fmt(round(rec_m)),
                "rec_max": self._fmt(rec_max),
                "rec_d1": self._fmt(rec_1),
                "rea_media": self._fmt(round(rea_m)),
                "rea_max": self._fmt(rea_max),
                "rea_d1": self._fmt(rea_1),
                "bolsa": self._fmt(b),
                "ir_3d": self._fmt(ir_row, 1),
                "ir_ayer": self._fmt(ir_ayer_row, 1),
                "citad1": self._fmt(cit1)
            })

            tot_cdm["rec_media"] += rec_m
            tot_cdm["rec_max"] = max(tot_cdm["rec_max"], rec_max)
            tot_cdm["rec_d1"] += rec_1
            tot_cdm["rea_media"] += rea_m
            tot_cdm["rea_max"] = max(tot_cdm["rea_max"], rea_max)
            tot_cdm["rea_d1"] += rea_1
            tot_cdm["bolsa"] += b
            tot_cdm["citad1"] += cit1

            # Acumulador para CDMs Segmentados por Contrata
            if c not in segment_contratas_raw:
                segment_contratas_raw[c] = {
                    "rec_media_noswap": 0.0, "rec_max_noswap": 0, "rec_d1_noswap": 0,
                    "rea_media_noswap": 0.0, "rea_max_noswap": 0, "rea_d1_noswap": 0,
                    "bolsa_noswap": 0, "citado_noswap": 0, "citad1_noswap": 0, "pdtecita_noswap": 0,

                    "rec_media_ayt": 0.0, "rec_max_ayt": 0, "rec_d1_ayt": 0,
                    "rea_media_ayt": 0.0, "rea_max_ayt": 0, "rea_d1_ayt": 0,
                    "bolsa_ayt": 0, "citado_ayt": 0, "pdtecita_ayt": 0,

                    "rec_media_mig": 0.0, "rec_max_mig": 0, "rec_d1_mig": 0,
                    "rea_media_mig": 0.0, "rea_max_mig": 0, "rea_d1_mig": 0,
                    "bolsa_mig": 0, "citado_mig": 0, "pdtecita_mig": 0,

                    "rec_media_b2b": 0.0, "rec_max_b2b": 0, "rec_d1_b2b": 0,
                    "rea_media_b2b": 0.0, "rea_max_b2b": 0, "rea_d1_b2b": 0,
                    "bolsa_b2b": 0, "citado_b2b": 0, "pdtecita_b2b": 0
                }

            sc = segment_contratas_raw[c]
            sc["rec_media_noswap"] += float(r.get("rec_media_noswap") or 0.0)
            sc["rec_max_noswap"] = max(sc["rec_max_noswap"], int(r.get("rec_max_noswap") or 0))
            sc["rec_d1_noswap"] += int(r.get("rec_d1_noswap") or 0)
            sc["rea_media_noswap"] += float(r.get("rea_media_noswap") or 0.0)
            sc["rea_max_noswap"] = max(sc["rea_max_noswap"], int(r.get("rea_max_noswap") or 0))
            sc["rea_d1_noswap"] += int(r.get("rea_d1_noswap") or 0)
            sc["bolsa_noswap"] += int(r.get("bolsa_noswap") or 0)
            sc["citado_noswap"] += int(r.get("citado_noswap") or 0)
            sc["citad1_noswap"] += int(r.get("citad1_noswap") or 0)
            sc["pdtecita_noswap"] += int(r.get("pdtecita_noswap") or 0)

            sc["rec_media_ayt"] += float(r.get("rec_media_ayt") or 0.0)
            sc["rec_max_ayt"] = max(sc["rec_max_ayt"], int(r.get("rec_max_ayt") or 0))
            sc["rec_d1_ayt"] += int(r.get("rec_d1_ayt") or 0)
            sc["rea_media_ayt"] += float(r.get("rea_media_ayt") or 0.0)
            sc["rea_max_ayt"] = max(sc["rea_max_ayt"], int(r.get("rea_max_ayt") or 0))
            sc["rea_d1_ayt"] += int(r.get("rea_d1_ayt") or 0)
            sc["bolsa_ayt"] += int(r.get("bolsa_ayt_cdm") or 0)
            sc["citado_ayt"] += int(r.get("citado_ayt") or 0)
            sc["pdtecita_ayt"] += int(r.get("pdtecita_ayt") or 0)

            sc["rec_media_mig"] += float(r.get("rec_media_mig") or 0.0)
            sc["rec_max_mig"] = max(sc["rec_max_mig"], int(r.get("rec_max_mig") or 0))
            sc["rec_d1_mig"] += int(r.get("rec_d1_mig") or 0)
            sc["rea_media_mig"] += float(r.get("rea_media_mig") or 0.0)
            sc["rea_max_mig"] = max(sc["rea_max_mig"], int(r.get("rea_max_mig") or 0))
            sc["rea_d1_mig"] += int(r.get("rea_d1_mig") or 0)
            sc["bolsa_mig"] += int(r.get("bolsa_mig_cdm") or 0)
            sc["citado_mig"] += int(r.get("citado_mig") or 0)
            sc["pdtecita_mig"] += int(r.get("pdtecita_mig") or 0)

            sc["rec_media_b2b"] += float(r.get("rec_media_b2b") or 0.0)
            sc["rec_max_b2b"] = max(sc["rec_max_b2b"], int(r.get("rec_max_b2b") or 0))
            sc["rec_d1_b2b"] += int(r.get("rec_d1_b2b") or 0)
            sc["rea_media_b2b"] += float(r.get("rea_media_b2b") or 0.0)
            sc["rea_max_b2b"] = max(sc["rea_max_b2b"], int(r.get("rea_max_b2b") or 0))
            sc["rea_d1_b2b"] += int(r.get("rea_d1_b2b") or 0)
            sc["bolsa_b2b"] += int(r.get("bolsa_b2b_cdm") or 0)
            sc["citado_b2b"] += int(r.get("citado_b2b") or 0)
            sc["pdtecita_b2b"] += int(r.get("pdtecita_b2b") or 0)

            # 1. Agrupación Resumen IR
            if c not in resumen_ir_raw:
                resumen_ir_raw[c] = {
                    "bolsa_norm": 0, "real_3dias_norm": 0.0,
                    "bolsa_global": 0, "real_3dias_global": 0.0,
                    "bolsa_ayt": 0, "real_3dias_ayt": 0.0,
                    "bolsa_mig": 0, "real_3dias_mig": 0.0,
                    "bolsa_b2b": 0, "real_3dias_b2b": 0.0
                }
            resumen_ir_raw[c]["bolsa_norm"] += b_norm
            resumen_ir_raw[c]["real_3dias_norm"] += r3_norm
            resumen_ir_raw[c]["bolsa_global"] += b
            resumen_ir_raw[c]["real_3dias_global"] += (rea_m * 3)
            resumen_ir_raw[c]["bolsa_ayt"] += b_ayt
            resumen_ir_raw[c]["real_3dias_ayt"] += r3_ayt
            resumen_ir_raw[c]["bolsa_mig"] += b_mig
            resumen_ir_raw[c]["real_3dias_mig"] += r3_mig
            resumen_ir_raw[c]["bolsa_b2b"] += b_b2b
            resumen_ir_raw[c]["real_3dias_b2b"] += r3_b2b

            tot_resumen_ir["bolsa_norm"] += b_norm
            tot_resumen_ir["real_3dias_norm"] += r3_norm
            tot_resumen_ir["bolsa_global"] += b
            tot_resumen_ir["real_3dias_global"] += (rea_m * 3)
            tot_resumen_ir["bolsa_ayt"] += b_ayt
            tot_resumen_ir["real_3dias_ayt"] += r3_ayt
            tot_resumen_ir["bolsa_mig"] += b_mig
            tot_resumen_ir["real_3dias_mig"] += r3_mig
            tot_resumen_ir["bolsa_b2b"] += b_b2b
            tot_resumen_ir["real_3dias_b2b"] += r3_b2b

            # 2. Agrupación Looker Tabla IR
            if c not in ir_raw:
                ir_raw[c] = {
                    "rec_media": 0.0, "rec_max": 0, "rec_d2": 0, "rec_d1": 0,
                    "rea_media": 0.0, "rea_max": 0, "rea_d2": 0, "rea_d1": 0,
                    "bolsa": 0, "citado": 0, "real_ayer": 0
                }
            ir_raw[c]["rec_media"] += rec_m
            ir_raw[c]["rec_max"] = max(ir_raw[c]["rec_max"], rec_max)
            ir_raw[c]["rec_d2"] += rec_2
            ir_raw[c]["rec_d1"] += rec_1
            ir_raw[c]["rea_media"] += rea_m
            ir_raw[c]["rea_max"] = max(ir_raw[c]["rea_max"], rea_max)
            ir_raw[c]["rea_d2"] += rea_2
            ir_raw[c]["rea_d1"] += rea_1
            ir_raw[c]["bolsa"] += b
            ir_raw[c]["citado"] += cit
            ir_raw[c]["real_ayer"] += rea_1

            tot_ir["rec_media"] += rec_m
            tot_ir["rec_max"] = max(tot_ir["rec_max"], rec_max)
            tot_ir["rec_d2"] += rec_2
            tot_ir["rec_d1"] += rec_1
            tot_ir["rea_media"] += rea_m
            tot_ir["rea_max"] = max(tot_ir["rea_max"], rea_max)
            tot_ir["rea_d2"] += rea_2
            tot_ir["rea_d1"] += rea_1
            tot_ir["bolsa"] += b
            tot_ir["citado"] += cit
            tot_ir["real_ayer"] += rea_1

            # 3. Agrupación Looker Tabla Entradas
            if c not in entradas_raw:
                entradas_raw[c] = {k: 0.0 if "media" in k else 0 for k in tot_entradas.keys()}
            entradas_raw[c]["rec_media"] += rec_m
            entradas_raw[c]["rec_max"] = max(entradas_raw[c]["rec_max"], rec_max)
            tot_entradas["rec_media"] += rec_m
            tot_entradas["rec_max"] = max(tot_entradas["rec_max"], rec_max)

            for d_idx in range(1, 9):
                k = f"rec_d{d_idx}"
                v = int(r.get(k) or 0)
                entradas_raw[c][k] += v
                tot_entradas[k] += v

            # 4. Agrupación Looker Tabla Cierres
            if c not in cierres_raw:
                cierres_raw[c] = {k: 0.0 if "media" in k else 0 for k in tot_cierres.keys()}
            cierres_raw[c]["rea_media"] += rea_m
            cierres_raw[c]["rea_max"] = max(cierres_raw[c]["rea_max"], rea_max)
            tot_cierres["rea_media"] += rea_m
            tot_cierres["rea_max"] = max(tot_cierres["rea_max"], rea_max)

            for d_idx in range(1, 9):
                k = f"rea_d{d_idx}"
                v = int(r.get(k) or 0)
                cierres_raw[c][k] += v
                tot_cierres[k] += v

            # 5. Agrupación Looker Tabla Citas y Calidad
            if c not in citas_raw:
                citas_raw[c] = {
                    "bolsa": 0, "citado": 0, "citad1": 0, "citad2": 0,
                    "pdtecita": 0, "pospuesta": 0, "pospuesta_mm": 0
                }
            citas_raw[c]["bolsa"] += b
            citas_raw[c]["citado"] += cit
            citas_raw[c]["citad1"] += cit1
            citas_raw[c]["citad2"] += cit2
            citas_raw[c]["pdtecita"] += p_cita
            citas_raw[c]["pospuesta"] += posp
            citas_raw[c]["pospuesta_mm"] += posp_mm

            tot_citas["bolsa"] += b
            tot_citas["citado"] += cit
            tot_citas["citad1"] += cit1
            tot_citas["citad2"] += cit2
            tot_citas["pdtecita"] += p_cita
            tot_citas["pospuesta"] += posp
            tot_citas["pospuesta_mm"] += posp_mm

            # 6. Agrupación por Grupo de Marca
            if g not in marcas_agg:
                marcas_agg[g] = {
                    "bolsa": 0, "rec_media": 0.0, "rea_media": 0.0,
                    "rea_d1": 0, "citado": 0, "citad1": 0,
                    "pdtecita": 0, "pospuesta": 0, "pospuesta_mm": 0
                }
            marcas_agg[g]["bolsa"] += b
            marcas_agg[g]["rec_media"] += rec_m
            marcas_agg[g]["rea_media"] += rea_m
            marcas_agg[g]["rea_d1"] += rea_1
            marcas_agg[g]["citado"] += cit
            marcas_agg[g]["citad1"] += cit1
            marcas_agg[g]["pdtecita"] += p_cita
            marcas_agg[g]["pospuesta"] += posp
            marcas_agg[g]["pospuesta_mm"] += posp_mm

        # Ordenar CDM detallada por Grupo, Contrata y Provincia
        cdm_filas.sort(key=lambda x: (x["grupo"], x["contrata"], x["provincia"]))

        tot_ir_3d_cdm = (tot_cdm["bolsa"] / tot_cdm["rea_media"]) if tot_cdm["rea_media"] > 0 else 0
        tot_ir_ayer_cdm = (tot_cdm["bolsa"] / tot_cdm["rea_d1"]) if tot_cdm["rea_d1"] > 0 else 0

        total_cdm_resumen = {
            "rec_media": self._fmt(round(tot_cdm["rec_media"])),
            "rec_max": self._fmt(tot_cdm["rec_max"]),
            "rec_d1": self._fmt(tot_cdm["rec_d1"]),
            "rea_media": self._fmt(round(tot_cdm["rea_media"])),
            "rea_max": self._fmt(tot_cdm["rea_max"]),
            "rea_d1": self._fmt(tot_cdm["rea_d1"]),
            "bolsa": self._fmt(tot_cdm["bolsa"]),
            "ir_3d": self._fmt(tot_ir_3d_cdm, 1),
            "ir_ayer": self._fmt(tot_ir_ayer_cdm, 1),
            "citad1": self._fmt(tot_cdm["citad1"])
        }

        # Construcción de los 4 CDMs Segmentados de Looker:
        cdm_noswap = self._build_segment_table(segment_contratas_raw, "noswap", use_citad1=False)
        cdm_ayt = self._build_segment_table(segment_contratas_raw, "ayt", use_citad1=False)
        cdm_mig = self._build_segment_table(segment_contratas_raw, "mig", use_citad1=False)
        cdm_b2b = self._build_segment_table(segment_contratas_raw, "b2b", use_citad1=False)

        # Formateo de la Tabla «Resumen IR»
        resumen_ir_contratas = {}
        for c, d in sorted(resumen_ir_raw.items()):
            ir_n = (d["bolsa_norm"] / (d["real_3dias_norm"] / 3)) if d["real_3dias_norm"] > 0 else None
            ir_g = (d["bolsa_global"] / (d["real_3dias_global"] / 3)) if d["real_3dias_global"] > 0 else None
            ir_ayt = (d["bolsa_ayt"] / (d["real_3dias_ayt"] / 3)) if d["real_3dias_ayt"] > 0 else None
            ir_mig = (d["bolsa_mig"] / (d["real_3dias_mig"] / 3)) if d["real_3dias_mig"] > 0 else None
            ir_b2b = (d["bolsa_b2b"] / (d["real_3dias_b2b"] / 3)) if d["real_3dias_b2b"] > 0 else None

            resumen_ir_contratas[c] = {
                "ir_normalizado": self._fmt(ir_n, 1),
                "ir_global": self._fmt(ir_g, 1),
                "ir_ayt": self._fmt(ir_ayt, 1),
                "ir_migraciones": self._fmt(ir_mig, 1),
                "ir_b2b": self._fmt(ir_b2b, 1)
            }

        tot_ir_n = (tot_resumen_ir["bolsa_norm"] / (tot_resumen_ir["real_3dias_norm"] / 3)) if tot_resumen_ir["real_3dias_norm"] > 0 else 0
        tot_ir_g = (tot_resumen_ir["bolsa_global"] / (tot_resumen_ir["real_3dias_global"] / 3)) if tot_resumen_ir["real_3dias_global"] > 0 else 0
        tot_ir_ayt = (tot_resumen_ir["bolsa_ayt"] / (tot_resumen_ir["real_3dias_ayt"] / 3)) if tot_resumen_ir["real_3dias_ayt"] > 0 else 0
        tot_ir_mig = (tot_resumen_ir["bolsa_mig"] / (tot_resumen_ir["real_3dias_mig"] / 3)) if tot_resumen_ir["real_3dias_mig"] > 0 else 0
        tot_ir_b2b = (tot_resumen_ir["bolsa_b2b"] / (tot_resumen_ir["real_3dias_b2b"] / 3)) if tot_resumen_ir["real_3dias_b2b"] > 0 else 0

        total_resumen_ir = {
            "ir_normalizado": self._fmt(tot_ir_n, 1),
            "ir_global": self._fmt(tot_ir_g, 1),
            "ir_ayt": self._fmt(tot_ir_ayt, 1),
            "ir_migraciones": self._fmt(tot_ir_mig, 1),
            "ir_b2b": self._fmt(tot_ir_b2b, 1)
        }

        # Formatear Tabla IR de Looker por Contrata
        looker_ir_contratas = {}
        for c, d in sorted(ir_raw.items()):
            ir_val = (d["bolsa"] / d["rea_media"] ) if d["rea_media"] > 0 else 0
            ir_ayer_val = (d["bolsa"] / d["real_ayer"]) if d["real_ayer"] > 0 else 0
            looker_ir_contratas[c] = {
                "rec_media": self._fmt(round(d["rec_media"])),
                "rec_max": self._fmt(d["rec_max"]),
                "rec_d2": self._fmt(d["rec_d2"]),
                "rec_d1": self._fmt(d["rec_d1"]),
                "rea_media": self._fmt(round(d["rea_media"])),
                "rea_max": self._fmt(d["rea_max"]),
                "rea_d2": self._fmt(d["rea_d2"]),
                "rea_d1": self._fmt(d["rea_d1"]),
                "bolsa": self._fmt(d["bolsa"]),
                "ir": self._fmt(ir_val, 2),
                "ir_ayer": self._fmt(ir_ayer_val, 2),
                "citado": self._fmt(d["citado"])
            }

        total_ir_val = (tot_ir["bolsa"] / tot_ir["rea_media"]) if tot_ir["rea_media"] > 0 else 0
        total_ir_ayer_val = (tot_ir["bolsa"] / tot_ir["real_ayer"]) if tot_ir["real_ayer"] > 0 else 0
        total_looker_ir = {
            "rec_media": self._fmt(round(tot_ir["rec_media"])),
            "rec_max": self._fmt(tot_ir["rec_max"]),
            "rec_d2": self._fmt(tot_ir["rec_d2"]),
            "rec_d1": self._fmt(tot_ir["rec_d1"]),
            "rea_media": self._fmt(round(tot_ir["rea_media"])),
            "rea_max": self._fmt(tot_ir["rea_max"]),
            "rea_d2": self._fmt(tot_ir["rea_d2"]),
            "rea_d1": self._fmt(tot_ir["rea_d1"]),
            "bolsa": self._fmt(tot_ir["bolsa"]),
            "ir": self._fmt(total_ir_val, 2),
            "ir_ayer": self._fmt(total_ir_ayer_val, 2),
            "citado": self._fmt(tot_ir["citado"])
        }

        # Formatear Tabla Entradas de Looker
        timeline_entradas = {
            c: {k: self._fmt(round(v) if "media" in k else v) for k, v in data.items()}
            for c, data in sorted(entradas_raw.items())
        }
        total_entradas = {
            k: self._fmt(round(v) if "media" in k else v)
            for k, v in tot_entradas.items()
        }

        # Formatear Tabla Cierres de Looker (Pág 2)
        timeline_cierres = {
            c: {k: self._fmt(round(v) if "media" in k else v) for k, v in data.items()}
            for c, data in sorted(cierres_raw.items())
        }
        total_cierres = {
            k: self._fmt(round(v) if "media" in k else v)
            for k, v in tot_cierres.items()
        }

        # Formatear Tabla Citas y Calidad de Bolsa de Looker (Pág 2)
        citas_contratas = {}
        for c, d in sorted(citas_raw.items()):
            ir_prev = (d["bolsa"] / d["citado"]) if d["citado"] > 0 else 0
            n_inst = round(d["bolsa"] / 2.5)
            per_pdte = (100 * d["pdtecita"] / d["bolsa"]) if d["bolsa"] > 0 else 0
            den_posp = d["bolsa"] + d["pospuesta_mm"]
            per_posp = (100 * d["pospuesta"] / den_posp) if den_posp > 0 else 0

            citas_contratas[c] = {
                "bolsa": self._fmt(d["bolsa"]),
                "citado": self._fmt(d["citado"]),
                "citad1": self._fmt(d["citad1"]),
                "citad2": self._fmt(d["citad2"]),
                "ir_previsto": self._fmt(ir_prev, 2),
                "n_inst_bjetivo": self._fmt(n_inst),
                "per_pdte_cita": self._fmt(per_pdte, 1),
                "per_pospuesto": self._fmt(per_posp, 1)
            }

        tot_ir_prev = (tot_citas["bolsa"] / tot_citas["citado"]) if tot_citas["citado"] > 0 else 0
        tot_n_inst = round(tot_citas["bolsa"] / 2.5)
        tot_per_pdte = (100 * tot_citas["pdtecita"] / tot_citas["bolsa"]) if tot_citas["bolsa"] > 0 else 0
        tot_den_posp = tot_citas["bolsa"] + tot_citas["pospuesta_mm"]
        tot_per_posp = (100 * tot_citas["pospuesta"] / tot_den_posp) if tot_den_posp > 0 else 0

        total_citas = {
            "bolsa": self._fmt(tot_citas["bolsa"]),
            "citado": self._fmt(tot_citas["citado"]),
            "citad1": self._fmt(tot_citas["citad1"]),
            "citad2": self._fmt(tot_citas["citad2"]),
            "ir_previsto": self._fmt(tot_ir_prev, 2),
            "n_inst_bjetivo": self._fmt(tot_n_inst),
            "per_pdte_cita": self._fmt(tot_per_pdte, 1),
            "per_pospuesto": self._fmt(tot_per_posp, 1)
        }

        # Formatear Resumen por Marca
        kpis_marcas = {}
        for g, vals in marcas_agg.items():
            den_posp = vals["bolsa"] + vals["pospuesta_mm"]
            kpis_marcas[g] = {
                "bolsa": self._fmt(vals["bolsa"]),
                "rec_media": self._fmt(round(vals["rec_media"])),
                "rea_media": self._fmt(round(vals["rea_media"])),
                "rea_d1": self._fmt(vals["rea_d1"]),
                "citado": self._fmt(vals["citado"]),
                "citad1": self._fmt(vals["citad1"]),
                "ir": self._fmt(vals["bolsa"] / vals["rea_media"] if vals["rea_media"] > 0 else 0, 2),
                "ir_ayer": self._fmt(vals["bolsa"] / vals["rea_d1"] if vals["rea_d1"] > 0 else 0, 2),
                "ir_previsto": self._fmt(vals["bolsa"] / vals["citado"] if vals["citado"] > 0 else 0, 2),
                "per_pdte": self._fmt(100 * vals["pdtecita"] / vals["bolsa"] if vals["bolsa"] > 0 else 0, 1),
                "per_posp": self._fmt(100 * vals["pospuesta"] / den_posp if den_posp > 0 else 0, 1)
            }

        den_posp_tot = tot_bolsa_g + tot_pospuesta_mm_g
        total_global = {
            "bolsa": self._fmt(tot_bolsa_g),
            "rec_media": self._fmt(round(tot_rec_media_g)),
            "rec_d1": self._fmt(tot_entradas["rec_d1"]),
            "rea_media": self._fmt(round(tot_rea_media_g)),
            "rea_d1": self._fmt(tot_rea_d1_g),
            "citado": self._fmt(tot_citado_g),
            "citad1": self._fmt(tot_citad1_g),
            "ir": self._fmt(tot_bolsa_g / tot_rea_media_g if tot_rea_media_g > 0 else 0, 2),
            "ir_ayer": self._fmt(tot_bolsa_g / tot_rea_d1_g if tot_rea_d1_g > 0 else 0, 2),
            "ir_previsto": self._fmt(tot_bolsa_g / tot_citado_g if tot_citado_g > 0 else 0, 2),
            "per_pdte": self._fmt(100 * tot_pdtecita_g / tot_bolsa_g if tot_bolsa_g > 0 else 0, 1),
            "per_posp": self._fmt(100 * tot_pospuesta_g / den_posp_tot if den_posp_tot > 0 else 0, 1)
        }

        return {
            "cdm_filas": cdm_filas,
            "total_cdm_resumen": total_cdm_resumen,
            "cdm_noswap": cdm_noswap,
            "cdm_ayt": cdm_ayt,
            "cdm_mig": cdm_mig,
            "cdm_b2b": cdm_b2b,
            "resumen_ir_contratas": resumen_ir_contratas,
            "total_resumen_ir": total_resumen_ir,
            "looker_ir_contratas": looker_ir_contratas,
            "total_looker_ir": total_looker_ir,
            "timeline_entradas": timeline_entradas,
            "total_entradas": total_entradas,
            "timeline_cierres": timeline_cierres,
            "total_cierres": total_cierres,
            "citas_contratas": citas_contratas,
            "total_citas": total_citas,
            "kpis_marcas": kpis_marcas,
            "total_global": total_global
        }

    def generate_html_file(self, raw_rows: List[Dict[str, Any]], target_date: date, output_dir: str) -> str:
        os.makedirs(output_dir, exist_ok=True)
        date_str = target_date.strftime("%Y-%m-%d")
        date_es = target_date.strftime("%d/%m/%Y")
        file_name = f"Dashboard_IR_{date_str}.html"
        output_path = os.path.join(output_dir, file_name)

        filtered_rows = [
            r for r in raw_rows
            if str(r.get("contrata") or "").strip().upper() not in self.EXCLUDED_CONTRATAS
        ]

        json_payload = json.dumps(filtered_rows, default=str, ensure_ascii=False)
        html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Orange — Cuadro de Mando IR y Operaciones ({date_es})</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 0; padding: 10px; background: #f8fafc; font-size: 11px; color: #1e293b; }}
    .header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; border-bottom: 2px solid #ff7900; padding-bottom: 6px; }}
    .brand-container {{ display: flex; align-items: center; gap: 10px; }}
    .orange-logo-box {{
      width: 40px;
      height: 40px;
      background-color: #ff7900;
      border-radius: 8px 0 0 0;
      position: relative;
      display: inline-flex;
      align-items: flex-end;
      padding: 0 0 4px 5px;
      box-sizing: border-box;
      flex-shrink: 0;
    }}
    .orange-logo-text {{
      color: #ffffff;
      font-family: -apple-system, BlinkMacSystemFont, "Helvetica Neue", Arial, sans-serif;
      font-size: 9.5px;
      font-weight: 800;
      line-height: 1;
      letter-spacing: -0.2px;
    }}
    .orange-logo-text sup {{
      font-size: 5px;
      vertical-align: super;
      line-height: 0;
      margin-left: 1px;
    }}
    .header h1 {{ margin: 0; font-size: 15px; color: #0f172a; font-weight: 800; }}
    .header h1 span {{ color: #ff7900; }}
    .filters-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 8px; margin-bottom: 10px; background: #ffffff; padding: 10px; border-radius: 6px; border: 1px solid #e2e8f0; }}
    .filter-item label {{ display: block; font-size: 9.5px; font-weight: 700; color: #64748b; margin-bottom: 2px; text-transform: uppercase; }}
    select, input {{ width: 100%; box-sizing: border-box; padding: 4px 6px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 11px; background: #fff; }}
    .tbl-container {{ overflow-x: auto; background: #fff; border: 1px solid #e2e8f0; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); max-height: 80vh; }}
    table {{ width: 100%; border-collapse: collapse; text-align: right; }}
    th, td {{ padding: 5px 6px; border-bottom: 1px solid #f1f5f9; white-space: nowrap; }}
    thead th {{ background: #fff8f3; color: #475569; font-weight: 700; border-bottom: 2px solid #ffcc99; position: sticky; top: 0; z-index: 10; }}
    .text-left {{ text-align: left; }}
    .row-total {{ background: #fffdf9; font-weight: 800; border-top: 2px solid #ffaa66; border-bottom: 2px solid #ffaa66; color: #0f172a; position: sticky; bottom: 0; z-index: 9; }}
    .pill {{ padding: 2px 5px; border-radius: 6px; font-weight: 700; font-size: 9.5px; display: inline-block; }}
    .pill-green {{ background: #dcfce7; color: #15803d; }}
    .pill-amber {{ background: #fef9c3; color: #a16207; }}
    .pill-red {{ background: #fee2e2; color: #b91c1c; }}
  </style>
</head>
<body>
  <div class="header">
    <div class="brand-container">
      <div class="orange-logo-box">
        <span class="orange-logo-text">orange<sup>TM</sup></span>
      </div>
      <h1 style="border-left: 2px solid #cbd5e1; padding-left: 10px; margin-left: 2px;">
        Cuadro de Mando IR y Operaciones ({date_es})
      </h1>
    </div>
    <span style="font-weight: 700; color: #64748b; font-size: 10.5px;">Orange Operaciones</span>
  </div>

  <div class="filters-grid">
    <div class="filter-item">
      <label>Grupo Marcas</label>
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
          <th class="text-left">Grupo Marcas</th>
          <th class="text-left">Contrata</th>
          <th class="text-left">Provincia</th>
          <th>Rec Media</th><th>Rec Max</th><th>Rec D-1</th>
          <th>Real Media</th><th>Real Max</th><th>Real D-1</th>
          <th>Bolsa</th><th>IR 3D</th><th>IR Ayer</th><th>Cita D+1</th>
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
      if (v <= 3.0) return `<span class="pill pill-green">${{fmt(v, 1)}}</span>`;
      if (v <= 4.0) return `<span class="pill pill-amber">${{fmt(v, 1)}}</span>`;
      return `<span class="pill pill-red">${{fmt(v, 1)}}</span>`;
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

      let sumRecM = 0, sumRecMax = 0, sumRecD1 = 0;
      let sumReaM = 0, sumReaMax = 0, sumReaD1 = 0;
      let sumBolsa = 0, sumCitaD1 = 0;

      rows.forEach(d => {{
        sumRecM += Number(d.rec_media || 0);
        sumRecMax = Math.max(sumRecMax, Number(d.rec_max || 0));
        sumRecD1 += Number(d.rec_d1 || 0);

        sumReaM += Number(d.rea_media || 0);
        sumReaMax = Math.max(sumReaMax, Number(d.rea_max || 0));
        sumReaD1 += Number(d.rea_d1 || 0);

        sumBolsa += Number(d.bolsa || 0);
        sumCitaD1 += Number(d.citad1 || 0);

        const ir3d = Number(d.rea_media) > 0 ? (Number(d.bolsa) / Number(d.rea_media)) : null;
        const irAyer = Number(d.rea_d1) > 0 ? (Number(d.bolsa) / Number(d.rea_d1)) : null;

        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td class="text-left">${{d.grupo}}</td>
          <td class="text-left"><strong>${{d.contrata}}</strong></td>
          <td class="text-left">${{d.provincia}}</td>
          <td>${{fmt(d.rec_media, 0)}}</td><td>${{fmt(d.rec_max)}}</td><td>${{fmt(d.rec_d1)}}</td>
          <td>${{fmt(d.rea_media, 0)}}</td><td>${{fmt(d.rea_max)}}</td><td>${{fmt(d.rea_d1)}}</td>
          <td><strong>${{fmt(d.bolsa)}}</strong></td>
          <td>${{getPill(ir3d)}}</td>
          <td>${{getPill(irAyer)}}</td>
          <td>${{fmt(d.citad1)}}</td>
        `;
        tb.appendChild(tr);
      }});

      const irTot = sumReaM > 0 ? (sumBolsa / sumReaM) : 0;
      const irAyerTot = sumReaD1 > 0 ? (sumBolsa / sumReaD1) : 0;

      document.getElementById('tfoot').innerHTML = `
        <td class="text-left" colspan="3">Total (${{rows.length}} registros)</td>
        <td>${{fmt(sumRecM, 0)}}</td><td>${{fmt(sumRecMax)}}</td><td>${{fmt(sumRecD1)}}</td>
        <td>${{fmt(sumReaM, 0)}}</td><td>${{fmt(sumReaMax)}}</td><td>${{fmt(sumReaD1)}}</td>
        <td><strong>${{fmt(sumBolsa)}}</strong></td>
        <td>${{getPill(irTot)}}</td>
        <td>${{getPill(irAyerTot)}}</td>
        <td>${{fmt(sumCitaD1)}}</td>
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

    def render_email_body(
        self,
        transformed_data: Dict[str, Any],
        target_date: date,
        sharepoint_web_url: str
    ) -> str:
        """
        Renderiza el cuerpo HTML ultracompacto para el correo:
          - Incluye los 4 Cuadros de Mando (CDM) segmentados de Looker:
              1. Toda la Actividad (menos SWAP)
              2. Altas y Traslados (Masivo)
              3. Solo Migraciones (Masivo)
              4. Solo B2B
          - Incluye el Cuadro de Mando Detallado por Provincia.
          - Sin contratas inactivas (FOL, ICONO, NO_DEFINIDO, OTROS eliminadas).
        """
        template_str = """
        <!DOCTYPE html>
        <html>
        <head><meta charset="utf-8"></head>
        <body style="margin: 0; padding: 8px; background-color: #f8fafc; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; color: #1e293b; font-size: 11px;">
          <!-- CABECERA PRINCIPAL CON LOGOTIPO OFICIAL DE ORANGE -->
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 8px;">
            <tr>
              <td valign="middle">
                <table cellpadding="0" cellspacing="0" border="0">
                  <tr>
                    <td valign="middle">
                      <table cellpadding="0" cellspacing="0" border="0" style="width: 36px; height: 36px; background-color: #ff7900; border-radius: 6px 0 0 0;">
                        <tr>
                          <td style="height: 22px; line-height: 0; font-size: 0;">&nbsp;</td>
                        </tr>
                        <tr>
                          <td style="height: 14px; padding-left: 4px; padding-bottom: 3px; vertical-align: bottom;">
                            <span style="color: #ffffff; font-family: Helvetica, Arial, sans-serif; font-size: 8.5px; font-weight: 800; line-height: 1; letter-spacing: -0.2px;">orange<sup style="font-size: 4px; vertical-align: super; line-height: 0;">TM</sup></span>
                          </td>
                        </tr>
                      </table>
                    </td>
                    <td style="padding-left: 10px; vertical-align: middle;">
                      <h2 style="margin: 0; font-size: 16px; font-weight: 800; color: #0f172a; text-transform: uppercase;">
                        Informe Operativo: Entradas, Cierres e IR
                      </h2>
                      <p style="margin: 1px 0 0 0; font-size: 11px; color: #64748b;">
                        Orange Operaciones — Cuadro de Mando Ejecutivo y Seguimiento de Calidad
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
              <td align="right" valign="top">
                <span style="background-color: #0f172a; color: #ffffff; padding: 4px 10px; border-radius: 12px; font-size: 10px; font-weight: 700;">
                  Foto: {{ fecha_ayer_formateada }}
                </span>
              </td>
            </tr>
          </table>

          <!-- BANNER COMPACTO: ACCESO DIRECTO A SHAREPOINT -->
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 10px;">
            <tr>
              <td style="background-color: #eff6ff; border: 1px solid #bfdbfe; border-left: 4px solid #ff7900; border-radius: 4px; padding: 7px 12px;">
                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                  <tr>
                    <td valign="middle">
                      <span style="font-size: 11.5px; font-weight: 800; color: #0f172a;">Cuadro de Mando Interactivo</span>
                      <span style="font-size: 11px; color: #475569; margin-left: 6px;">(Filtros por Grupo, Contrata y Provincia en SharePoint):</span>
                    </td>
                    <td align="right" valign="middle" style="white-space: nowrap; padding-left: 10px;">
                      <a href="{{ sharepoint_web_url }}" target="_blank" style="background-color: #ff7900; color: #ffffff; text-decoration: none; padding: 5px 12px; border-radius: 4px; font-size: 10.5px; font-weight: 700; display: inline-block;">
                        Ver en SharePoint &rarr;
                      </a>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>

          <!-- TARJETAS RESUMEN KPI -->
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 12px;">
            <tr>
              <td width="23.5%" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 3px solid #ff6600; border-radius: 6px; padding: 8px 10px;">
                <div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">GMM (MásMóvil)</div>
                <div style="font-size: 20px; font-weight: 800; color: #0f172a; margin: 2px 0;">{{ kpis_marcas.get('MAS MOVIL', {}).get('ir', '-') }}</div>
                <div style="font-size: 10.5px; color: #475569;">Bolsa: <strong>{{ kpis_marcas.get('MAS MOVIL', {}).get('bolsa', '-') }}</strong> | D-1: <strong>{{ kpis_marcas.get('MAS MOVIL', {}).get('rea_d1', '-') }}</strong></div>
              </td>
              <td width="2%">&nbsp;</td>
              <td width="23.5%" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 3px solid #ff7900; border-radius: 6px; padding: 8px 10px;">
                <div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">OSP (Orange)</div>
                <div style="font-size: 20px; font-weight: 800; color: #0f172a; margin: 2px 0;">{{ kpis_marcas.get('ORANGE', {}).get('ir', '-') }}</div>
                <div style="font-size: 10.5px; color: #475569;">Bolsa: <strong>{{ kpis_marcas.get('ORANGE', {}).get('bolsa', '-') }}</strong> | D-1: <strong>{{ kpis_marcas.get('ORANGE', {}).get('rea_d1', '-') }}</strong></div>
              </td>
              <td width="2%">&nbsp;</td>
              <td width="23.5%" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 3px solid #6366f1; border-radius: 6px; padding: 8px 10px;">
                <div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">KRTV (Euskaltel)</div>
                <div style="font-size: 20px; font-weight: 800; color: #0f172a; margin: 2px 0;">{{ kpis_marcas.get('KRTV', {}).get('ir', '-') }}</div>
                <div style="font-size: 10.5px; color: #475569;">Bolsa: <strong>{{ kpis_marcas.get('KRTV', {}).get('bolsa', '-') }}</strong> | D-1: <strong>{{ kpis_marcas.get('KRTV', {}).get('rea_d1', '-') }}</strong></div>
              </td>
              <td width="2%">&nbsp;</td>
              <td width="23.5%" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 3px solid #0f172a; border-radius: 6px; padding: 8px 10px;">
                <div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Total Proyecto</div>
                <div style="font-size: 20px; font-weight: 800; color: #0f172a; margin: 2px 0;">{{ total_global.get('ir', '-') }}</div>
                <div style="font-size: 10.5px; color: #475569;">Bolsa: <strong>{{ total_global.get('bolsa', '-') }}</strong> | IR Ayer: <strong>{{ total_global.get('ir_ayer', '-') }}</strong></div>
              </td>
            </tr>
          </table>

          <!-- TABLA RESUMEN IR (SEGMENTADA) -->
          <div style="font-size: 13px; font-weight: 800; color: #ff5722; margin-bottom: 4px;">
            Resumen IR
          </div>
          <table width="100%" cellpadding="4" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 4px; font-size: 11px; text-align: right; margin-bottom: 14px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #ff5722; color: #ffffff;">
                <th align="left" style="padding: 6px 8px; font-weight: 700;">Contrata &#9652;</th>
                <th style="padding: 6px 8px; font-weight: 700;">IR Normalizado</th>
                <th style="padding: 6px 8px; font-weight: 700;">IR Global</th>
                <th style="padding: 6px 8px; font-weight: 700;">IR AyT</th>
                <th style="padding: 6px 8px; font-weight: 700;">IR Migraciones</th>
                <th style="padding: 6px 8px; font-weight: 700;">IR B2B</th>
              </tr>
            </thead>
            <tbody>
              {% for contrata, cdata in resumen_ir_contratas.items() %}
              <tr style="border-bottom: 1px solid #f1f5f9;">
                <td align="left" style="padding: 4px 8px; font-weight: 600; color: #1e293b;">{{ contrata }}</td>
                <td style="padding: 4px 8px; font-weight: 600;">{{ cdata.ir_normalizado }}</td>
                <td style="padding: 4px 8px; font-weight: 600;">{{ cdata.ir_global }}</td>
                <td style="padding: 4px 8px;">{{ cdata.ir_ayt }}</td>
                <td style="padding: 4px 8px;">{{ cdata.ir_migraciones }}</td>
                <td style="padding: 4px 8px;">{{ cdata.ir_b2b }}</td>
              </tr>
              {% endfor %}
              <tr style="background-color: #f8fafc; font-weight: 800; border-top: 2px solid #cbd5e1; border-bottom: 2px solid #cbd5e1; color: #0f172a;">
                <td align="left" style="padding: 6px 8px;">Total</td>
                <td style="padding: 6px 8px; font-weight: 900;">{{ total_resumen_ir.ir_normalizado }}</td>
                <td style="padding: 6px 8px; font-weight: 900;">{{ total_resumen_ir.ir_global }}</td>
                <td style="padding: 6px 8px; font-weight: 900;">{{ total_resumen_ir.ir_ayt }}</td>
                <td style="padding: 6px 8px; font-weight: 900;">{{ total_resumen_ir.ir_migraciones }}</td>
                <td style="padding: 6px 8px; font-weight: 900;">{{ total_resumen_ir.ir_b2b }}</td>
              </tr>
            </tbody>
          </table>

          <!-- 1. CDM SEGMENTADO: TODA LA ACTIVIDAD (MENOS SWAP) -->
          <div style="font-size: 13px; font-weight: 800; color: #ff5722; margin-bottom: 4px;">
            IR - Toda la Actividad (menos SWAP)
          </div>
          <table width="100%" cellpadding="3" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 4px; font-size: 10px; text-align: right; margin-bottom: 14px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #ff5722; color: #ffffff;">
                <th align="left" style="padding: 6px 6px; font-weight: 700;">Contrata &#9652;</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec Max</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec D-1</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real Max</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real D-1</th>
                <th style="padding: 6px 5px; font-weight: 700;">Bolsa</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR 3D</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR Ayer</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR Previsto</th>
                <th style="padding: 6px 5px; font-weight: 700;">Pdte Cita</th>
                <th style="padding: 6px 5px; font-weight: 700;">%Pdte Cita</th>
                <th style="padding: 6px 5px; font-weight: 700;">Cita D</th>
              </tr>
            </thead>
            <tbody>
              {% for c, f in cdm_noswap.filas.items() %}
              <tr style="border-bottom: 1px solid #f1f5f9;">
                <td align="left" style="padding: 3px 6px; font-weight: 600; color: #0f172a;">{{ c }}</td>
                <td style="padding: 3px 5px;">{{ f.rec_media }}</td>
                <td style="padding: 3px 5px;">{{ f.rec_max }}</td>
                <td style="padding: 3px 5px; font-weight: 600;">{{ f.rec_d1 }}</td>
                <td style="padding: 3px 5px;">{{ f.rea_media }}</td>
                <td style="padding: 3px 5px;">{{ f.rea_max }}</td>
                <td style="padding: 3px 5px; font-weight: 600;">{{ f.rea_d1 }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #0f172a;">{{ f.bolsa }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #c2410c;">{{ f.ir_3d }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #0369a1;">{{ f.ir_ayer }}</td>
                <td style="padding: 3px 5px;">{{ f.ir_previsto }}</td>
                <td style="padding: 3px 5px;">{{ f.pdtecita }}</td>
                <td style="padding: 3px 5px;">{{ f.per_pdte_cita }}</td>
                <td style="padding: 3px 5px;">{{ f.cita_col }}</td>
              </tr>
              {% endfor %}
              <tr style="background-color: #fff8f3; font-weight: 800; border-top: 2px solid #ff7900; border-bottom: 2px solid #ff7900; color: #0f172a;">
                <td align="left" style="padding: 6px 6px;">Total</td>
                <td style="padding: 6px 5px;">{{ cdm_noswap.total.rec_media }}</td>
                <td style="padding: 6px 5px;">{{ cdm_noswap.total.rec_max }}</td>
                <td style="padding: 6px 5px;">{{ cdm_noswap.total.rec_d1 }}</td>
                <td style="padding: 6px 5px;">{{ cdm_noswap.total.rea_media }}</td>
                <td style="padding: 6px 5px;">{{ cdm_noswap.total.rea_max }}</td>
                <td style="padding: 6px 5px;">{{ cdm_noswap.total.rea_d1 }}</td>
                <td style="padding: 6px 5px; font-weight: 900;">{{ cdm_noswap.total.bolsa }}</td>
                <td style="padding: 6px 5px; font-weight: 900; color: #c2410c;">{{ cdm_noswap.total.ir_3d }}</td>
                <td style="padding: 6px 5px; font-weight: 900; color: #0369a1;">{{ cdm_noswap.total.ir_ayer }}</td>
                <td style="padding: 6px 5px;">{{ cdm_noswap.total.ir_previsto }}</td>
                <td style="padding: 6px 5px;">{{ cdm_noswap.total.pdtecita }}</td>
                <td style="padding: 6px 5px;">{{ cdm_noswap.total.per_pdte_cita }}</td>
                <td style="padding: 6px 5px;">{{ cdm_noswap.total.cita_col }}</td>
              </tr>
            </tbody>
          </table>

          <!-- 2. CDM SEGMENTADO: ALTAS Y TRASLADOS (MASIVO) -->
          <div style="font-size: 13px; font-weight: 800; color: #ff5722; margin-bottom: 4px;">
            IR - Altas y Traslados (Masivo)
          </div>
          <table width="100%" cellpadding="3" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 4px; font-size: 10px; text-align: right; margin-bottom: 14px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #ff5722; color: #ffffff;">
                <th align="left" style="padding: 6px 6px; font-weight: 700;">Contrata &#9652;</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec Max</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec D-1</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real Max</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real D-1</th>
                <th style="padding: 6px 5px; font-weight: 700;">Bolsa</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR 3D</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR Ayer</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR Previsto</th>
                <th style="padding: 6px 5px; font-weight: 700;">Pdte Cita</th>
                <th style="padding: 6px 5px; font-weight: 700;">%Pdte Cita</th>
                <th style="padding: 6px 5px; font-weight: 700;">Cita D</th>
              </tr>
            </thead>
            <tbody>
              {% for c, f in cdm_ayt.filas.items() %}
              <tr style="border-bottom: 1px solid #f1f5f9;">
                <td align="left" style="padding: 3px 6px; font-weight: 600; color: #0f172a;">{{ c }}</td>
                <td style="padding: 3px 5px;">{{ f.rec_media }}</td>
                <td style="padding: 3px 5px;">{{ f.rec_max }}</td>
                <td style="padding: 3px 5px; font-weight: 600;">{{ f.rec_d1 }}</td>
                <td style="padding: 3px 5px;">{{ f.rea_media }}</td>
                <td style="padding: 3px 5px;">{{ f.rea_max }}</td>
                <td style="padding: 3px 5px; font-weight: 600;">{{ f.rea_d1 }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #0f172a;">{{ f.bolsa }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #c2410c;">{{ f.ir_3d }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #0369a1;">{{ f.ir_ayer }}</td>
                <td style="padding: 3px 5px;">{{ f.ir_previsto }}</td>
                <td style="padding: 3px 5px;">{{ f.pdtecita }}</td>
                <td style="padding: 3px 5px;">{{ f.per_pdte_cita }}</td>
                <td style="padding: 3px 5px;">{{ f.cita_col }}</td>
              </tr>
              {% endfor %}
              <tr style="background-color: #fff8f3; font-weight: 800; border-top: 2px solid #ff7900; border-bottom: 2px solid #ff7900; color: #0f172a;">
                <td align="left" style="padding: 6px 6px;">Total</td>
                <td style="padding: 6px 5px;">{{ cdm_ayt.total.rec_media }}</td>
                <td style="padding: 6px 5px;">{{ cdm_ayt.total.rec_max }}</td>
                <td style="padding: 6px 5px;">{{ cdm_ayt.total.rec_d1 }}</td>
                <td style="padding: 6px 5px;">{{ cdm_ayt.total.rea_media }}</td>
                <td style="padding: 6px 5px;">{{ cdm_ayt.total.rea_max }}</td>
                <td style="padding: 6px 5px;">{{ cdm_ayt.total.rea_d1 }}</td>
                <td style="padding: 6px 5px; font-weight: 900;">{{ cdm_ayt.total.bolsa }}</td>
                <td style="padding: 6px 5px; font-weight: 900; color: #c2410c;">{{ cdm_ayt.total.ir_3d }}</td>
                <td style="padding: 6px 5px; font-weight: 900; color: #0369a1;">{{ cdm_ayt.total.ir_ayer }}</td>
                <td style="padding: 6px 5px;">{{ cdm_ayt.total.ir_previsto }}</td>
                <td style="padding: 6px 5px;">{{ cdm_ayt.total.pdtecita }}</td>
                <td style="padding: 6px 5px;">{{ cdm_ayt.total.per_pdte_cita }}</td>
                <td style="padding: 6px 5px;">{{ cdm_ayt.total.cita_col }}</td>
              </tr>
            </tbody>
          </table>

          <!-- 3. CDM SEGMENTADO: SOLO MIGRACIONES (MASIVO) -->
          <div style="font-size: 13px; font-weight: 800; color: #ff5722; margin-bottom: 4px;">
            IR - Solo Migraciones (Masivo)
          </div>
          <table width="100%" cellpadding="3" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 4px; font-size: 10px; text-align: right; margin-bottom: 14px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #ff5722; color: #ffffff;">
                <th align="left" style="padding: 6px 6px; font-weight: 700;">Contrata &#9652;</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec Max</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec D-1</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real Max</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real D-1</th>
                <th style="padding: 6px 5px; font-weight: 700;">Bolsa</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR 3D</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR Ayer</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR Previsto</th>
                <th style="padding: 6px 5px; font-weight: 700;">Pdte Cita</th>
                <th style="padding: 6px 5px; font-weight: 700;">%Pdte Cita</th>
                <th style="padding: 6px 5px; font-weight: 700;">Cita D</th>
              </tr>
            </thead>
            <tbody>
              {% for c, f in cdm_mig.filas.items() %}
              <tr style="border-bottom: 1px solid #f1f5f9;">
                <td align="left" style="padding: 3px 6px; font-weight: 600; color: #0f172a;">{{ c }}</td>
                <td style="padding: 3px 5px;">{{ f.rec_media }}</td>
                <td style="padding: 3px 5px;">{{ f.rec_max }}</td>
                <td style="padding: 3px 5px; font-weight: 600;">{{ f.rec_d1 }}</td>
                <td style="padding: 3px 5px;">{{ f.rea_media }}</td>
                <td style="padding: 3px 5px;">{{ f.rea_max }}</td>
                <td style="padding: 3px 5px; font-weight: 600;">{{ f.rea_d1 }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #0f172a;">{{ f.bolsa }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #c2410c;">{{ f.ir_3d }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #0369a1;">{{ f.ir_ayer }}</td>
                <td style="padding: 3px 5px;">{{ f.ir_previsto }}</td>
                <td style="padding: 3px 5px;">{{ f.pdtecita }}</td>
                <td style="padding: 3px 5px;">{{ f.per_pdte_cita }}</td>
                <td style="padding: 3px 5px;">{{ f.cita_col }}</td>
              </tr>
              {% endfor %}
              <tr style="background-color: #fff8f3; font-weight: 800; border-top: 2px solid #ff7900; border-bottom: 2px solid #ff7900; color: #0f172a;">
                <td align="left" style="padding: 6px 6px;">Total</td>
                <td style="padding: 6px 5px;">{{ cdm_mig.total.rec_media }}</td>
                <td style="padding: 6px 5px;">{{ cdm_mig.total.rec_max }}</td>
                <td style="padding: 6px 5px;">{{ cdm_mig.total.rec_d1 }}</td>
                <td style="padding: 6px 5px;">{{ cdm_mig.total.rea_media }}</td>
                <td style="padding: 6px 5px;">{{ cdm_mig.total.rea_max }}</td>
                <td style="padding: 6px 5px;">{{ cdm_mig.total.rea_d1 }}</td>
                <td style="padding: 6px 5px; font-weight: 900;">{{ cdm_mig.total.bolsa }}</td>
                <td style="padding: 6px 5px; font-weight: 900; color: #c2410c;">{{ cdm_mig.total.ir_3d }}</td>
                <td style="padding: 6px 5px; font-weight: 900; color: #0369a1;">{{ cdm_mig.total.ir_ayer }}</td>
                <td style="padding: 6px 5px;">{{ cdm_mig.total.ir_previsto }}</td>
                <td style="padding: 6px 5px;">{{ cdm_mig.total.pdtecita }}</td>
                <td style="padding: 6px 5px;">{{ cdm_mig.total.per_pdte_cita }}</td>
                <td style="padding: 6px 5px;">{{ cdm_mig.total.cita_col }}</td>
              </tr>
            </tbody>
          </table>

          <!-- 4. CDM SEGMENTADO: SOLO B2B -->
          <div style="font-size: 13px; font-weight: 800; color: #ff5722; margin-bottom: 4px;">
            IR - Solo B2B
          </div>
          <table width="100%" cellpadding="3" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 4px; font-size: 10px; text-align: right; margin-bottom: 14px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #ff5722; color: #ffffff;">
                <th align="left" style="padding: 6px 6px; font-weight: 700;">Contrata &#9652;</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec Max</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec D-1</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real Max</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real D-1</th>
                <th style="padding: 6px 5px; font-weight: 700;">Bolsa</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR 3D</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR Ayer</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR Previsto</th>
                <th style="padding: 6px 5px; font-weight: 700;">Pdte Cita</th>
                <th style="padding: 6px 5px; font-weight: 700;">%Pdte Cita</th>
                <th style="padding: 6px 5px; font-weight: 700;">Cita D</th>
              </tr>
            </thead>
            <tbody>
              {% for c, f in cdm_b2b.filas.items() %}
              <tr style="border-bottom: 1px solid #f1f5f9;">
                <td align="left" style="padding: 3px 6px; font-weight: 600; color: #0f172a;">{{ c }}</td>
                <td style="padding: 3px 5px;">{{ f.rec_media }}</td>
                <td style="padding: 3px 5px;">{{ f.rec_max }}</td>
                <td style="padding: 3px 5px; font-weight: 600;">{{ f.rec_d1 }}</td>
                <td style="padding: 3px 5px;">{{ f.rea_media }}</td>
                <td style="padding: 3px 5px;">{{ f.rea_max }}</td>
                <td style="padding: 3px 5px; font-weight: 600;">{{ f.rea_d1 }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #0f172a;">{{ f.bolsa }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #c2410c;">{{ f.ir_3d }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #0369a1;">{{ f.ir_ayer }}</td>
                <td style="padding: 3px 5px;">{{ f.ir_previsto }}</td>
                <td style="padding: 3px 5px;">{{ f.pdtecita }}</td>
                <td style="padding: 3px 5px;">{{ f.per_pdte_cita }}</td>
                <td style="padding: 3px 5px;">{{ f.cita_col }}</td>
              </tr>
              {% endfor %}
              <tr style="background-color: #fff8f3; font-weight: 800; border-top: 2px solid #ff7900; border-bottom: 2px solid #ff7900; color: #0f172a;">
                <td align="left" style="padding: 6px 6px;">Total</td>
                <td style="padding: 6px 5px;">{{ cdm_b2b.total.rec_media }}</td>
                <td style="padding: 6px 5px;">{{ cdm_b2b.total.rec_max }}</td>
                <td style="padding: 6px 5px;">{{ cdm_b2b.total.rec_d1 }}</td>
                <td style="padding: 6px 5px;">{{ cdm_b2b.total.rea_media }}</td>
                <td style="padding: 6px 5px;">{{ cdm_b2b.total.rea_max }}</td>
                <td style="padding: 6px 5px;">{{ cdm_b2b.total.rea_d1 }}</td>
                <td style="padding: 6px 5px; font-weight: 900;">{{ cdm_b2b.total.bolsa }}</td>
                <td style="padding: 6px 5px; font-weight: 900; color: #c2410c;">{{ cdm_b2b.total.ir_3d }}</td>
                <td style="padding: 6px 5px; font-weight: 900; color: #0369a1;">{{ cdm_b2b.total.ir_ayer }}</td>
                <td style="padding: 6px 5px;">{{ cdm_b2b.total.ir_previsto }}</td>
                <td style="padding: 6px 5px;">{{ cdm_b2b.total.pdtecita }}</td>
                <td style="padding: 6px 5px;">{{ cdm_b2b.total.per_pdte_cita }}</td>
                <td style="padding: 6px 5px;">{{ cdm_b2b.total.cita_col }}</td>
              </tr>
            </tbody>
          </table>

          <!-- NUEVA TABLA: CUADRO DE MANDO DETALLADO (CDM) POR GRUPO, CONTRATA Y PROVINCIA -->
          <div style="font-size: 13px; font-weight: 800; color: #ff5722; margin-bottom: 4px;">
            Cuadro de Mando: Detalle Operativo por Grupo, Contrata y Provincia
          </div>
          <table width="100%" cellpadding="3" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 4px; font-size: 10px; text-align: right; margin-bottom: 14px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #ff5722; color: #ffffff;">
                <th align="left" style="padding: 6px 6px; font-weight: 700;">Grupo Marcas &#9652;</th>
                <th align="left" style="padding: 6px 6px; font-weight: 700;">Contrata &#9652;</th>
                <th align="left" style="padding: 6px 6px; font-weight: 700;">Provincia &#9652;</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec Max</th>
                <th style="padding: 6px 5px; font-weight: 700;">Rec D-1</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real Max</th>
                <th style="padding: 6px 5px; font-weight: 700;">Real D-1</th>
                <th style="padding: 6px 5px; font-weight: 700;">Bolsa</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR 3D</th>
                <th style="padding: 6px 5px; font-weight: 700;">IR Ayer</th>
                <th style="padding: 6px 5px; font-weight: 700;">Cita D+1</th>
              </tr>
            </thead>
            <tbody>
              {% for f in cdm_filas %}
              <tr style="border-bottom: 1px solid #f1f5f9;">
                <td align="left" style="padding: 3px 6px; color: #475569;">{{ f.grupo }}</td>
                <td align="left" style="padding: 3px 6px; font-weight: 600; color: #0f172a;">{{ f.contrata }}</td>
                <td align="left" style="padding: 3px 6px; color: #1e293b;">{{ f.provincia }}</td>
                <td style="padding: 3px 5px;">{{ f.rec_media }}</td>
                <td style="padding: 3px 5px;">{{ f.rec_max }}</td>
                <td style="padding: 3px 5px; font-weight: 600;">{{ f.rec_d1 }}</td>
                <td style="padding: 3px 5px;">{{ f.rea_media }}</td>
                <td style="padding: 3px 5px;">{{ f.rea_max }}</td>
                <td style="padding: 3px 5px; font-weight: 600;">{{ f.rea_d1 }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #0f172a;">{{ f.bolsa }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #c2410c;">{{ f.ir_3d }}</td>
                <td style="padding: 3px 5px; font-weight: 700; color: #0369a1;">{{ f.ir_ayer }}</td>
                <td style="padding: 3px 5px;">{{ f.citad1 }}</td>
              </tr>
              {% endfor %}
              <tr style="background-color: #fff8f3; font-weight: 800; border-top: 2px solid #ff7900; border-bottom: 2px solid #ff7900; color: #0f172a;">
                <td align="left" colspan="3" style="padding: 6px 6px;">Total ({{ cdm_filas|length }} combinaciones)</td>
                <td style="padding: 6px 5px;">{{ total_cdm_resumen.rec_media }}</td>
                <td style="padding: 6px 5px;">{{ total_cdm_resumen.rec_max }}</td>
                <td style="padding: 6px 5px;">{{ total_cdm_resumen.rec_d1 }}</td>
                <td style="padding: 6px 5px;">{{ total_cdm_resumen.rea_media }}</td>
                <td style="padding: 6px 5px;">{{ total_cdm_resumen.rea_max }}</td>
                <td style="padding: 6px 5px;">{{ total_cdm_resumen.rea_d1 }}</td>
                <td style="padding: 6px 5px; font-weight: 900;">{{ total_cdm_resumen.bolsa }}</td>
                <td style="padding: 6px 5px; font-weight: 900; color: #c2410c;">{{ total_cdm_resumen.ir_3d }}</td>
                <td style="padding: 6px 5px; font-weight: 900; color: #0369a1;">{{ total_cdm_resumen.ir_ayer }}</td>
                <td style="padding: 6px 5px;">{{ total_cdm_resumen.citad1 }}</td>
              </tr>
            </tbody>
          </table>

          <!-- TABLA 1: CUADRO DE MANDO IR GENERAL -->
          <div style="font-size: 12px; font-weight: 800; color: #0f172a; margin-bottom: 4px;">
            Cuadro de Mando IR (Índice de Rotación por Contrata)
          </div>
          <table width="100%" cellpadding="4" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #ebdcd0; border-radius: 4px; font-size: 10.5px; text-align: right; margin-bottom: 14px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #fff8f3; color: #475569; border-bottom: 2px solid #ffcc99;">
                <th align="left" style="padding: 6px 8px; color: #0f172a; font-weight: 700;">Contrata</th>
                <th style="padding: 6px 5px;">Rec. Media</th>
                <th style="padding: 6px 5px;">Rec. Max</th>
                <th style="padding: 6px 5px;">Rec. D-2</th>
                <th style="padding: 6px 5px; background-color: #fff3e6; color: #0f172a; font-weight: 700;">Rec. D-1</th>
                <th style="padding: 6px 5px;">Rea. Media</th>
                <th style="padding: 6px 5px;">Rea. Max</th>
                <th style="padding: 6px 5px;">Rea. D-2</th>
                <th style="padding: 6px 5px; background-color: #fff3e6; color: #0f172a; font-weight: 700;">Rea. D-1</th>
                <th style="padding: 6px 6px; font-weight: 800; color: #0f172a;">Bolsa</th>
                <th style="padding: 6px 6px; background-color: #fef08a; color: #0f172a; font-weight: 800;">IR</th>
                <th style="padding: 6px 5px;">IR Ayer</th>
                <th style="padding: 6px 5px;">Citado D</th>
              </tr>
            </thead>
            <tbody>
              {% for contrata, cdata in looker_ir_contratas.items() %}
              <tr style="border-bottom: 1px solid #f8eee7;">
                <td align="left" style="padding: 4px 8px; font-weight: 600; color: #1e293b;">{{ contrata }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rec_media }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rec_max }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rec_d2 }}</td>
                <td style="padding: 4px 5px; background-color: #fffdf5; font-weight: 600;">{{ cdata.rec_d1 }}</td>
                <td style="padding: 4px 5px; font-weight: 600;">{{ cdata.rea_media }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rea_max }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rea_d2 }}</td>
                <td style="padding: 4px 5px; background-color: #fffdf5; font-weight: 600;">{{ cdata.rea_d1 }}</td>
                <td style="padding: 4px 6px; font-weight: 800; color: #0f172a;">{{ cdata.bolsa }}</td>
                <td style="padding: 4px 6px; background-color: #fef9c3; font-weight: 800; color: #0f172a;">{{ cdata.ir }}</td>
                <td style="padding: 4px 5px;">{{ cdata.ir_ayer }}</td>
                <td style="padding: 4px 5px;">{{ cdata.citado }}</td>
              </tr>
              {% endfor %}
              <tr style="background-color: #fffdf9; font-weight: 800; border-top: 2px solid #ff7900; border-bottom: 2px solid #ff7900; color: #0f172a;">
                <td align="left" style="padding: 6px 8px;">Total</td>
                <td style="padding: 6px 5px;">{{ total_looker_ir.rec_media }}</td>
                <td style="padding: 6px 5px;">{{ total_looker_ir.rec_max }}</td>
                <td style="padding: 6px 5px;">{{ total_looker_ir.rec_d2 }}</td>
                <td style="padding: 6px 5px; background-color: #fef08a;">{{ total_looker_ir.rec_d1 }}</td>
                <td style="padding: 6px 5px;">{{ total_looker_ir.rea_media }}</td>
                <td style="padding: 6px 5px;">{{ total_looker_ir.rea_max }}</td>
                <td style="padding: 6px 5px;">{{ total_looker_ir.rea_d2 }}</td>
                <td style="padding: 6px 5px; background-color: #fef08a;">{{ total_looker_ir.rea_d1 }}</td>
                <td style="padding: 6px 6px; color: #0f172a;">{{ total_looker_ir.bolsa }}</td>
                <td style="padding: 6px 6px; background-color: #fef08a; color: #0f172a; font-weight: 900;">{{ total_looker_ir.ir }}</td>
                <td style="padding: 6px 5px;">{{ total_looker_ir.ir_ayer }}</td>
                <td style="padding: 6px 5px;">{{ total_looker_ir.citado }}</td>
              </tr>
            </tbody>
          </table>

          <!-- TABLA 2: ENTRADAS (D-8 A D-1) -->
          <div style="font-size: 12px; font-weight: 800; color: #0f172a; margin-bottom: 4px;">
            Entradas (Evolución Diaria de Recibidas D-8 a D-1 por Contrata)
          </div>
          <table width="100%" cellpadding="4" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #ebdcd0; border-radius: 4px; font-size: 10.5px; text-align: right; margin-bottom: 14px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #fff8f3; border-bottom: 2px solid #ffcc99; color: #475569;">
                <th align="left" style="padding: 6px 8px; color: #0f172a; font-weight: 700;">Contrata</th>
                <th style="padding: 6px 5px; font-weight: 700;">Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Max</th>
                <th style="padding: 6px 5px;">D-8</th>
                <th style="padding: 6px 5px;">D-7</th>
                <th style="padding: 6px 5px;">D-6</th>
                <th style="padding: 6px 5px;">D-5</th>
                <th style="padding: 6px 5px;">D-4</th>
                <th style="padding: 6px 5px;">D-3</th>
                <th style="padding: 6px 5px;">D-2</th>
                <th style="padding: 6px 5px; background-color: #fff3e6; font-weight: 700; color: #0f172a;">D-1</th>
              </tr>
            </thead>
            <tbody>
              {% for contrata, cdata in timeline_entradas.items() %}
              <tr style="border-bottom: 1px solid #f8eee7;">
                <td align="left" style="padding: 4px 8px; font-weight: 600; color: #1e293b;">{{ contrata }}</td>
                <td style="padding: 4px 5px; font-weight: 600;">{{ cdata.rec_media }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rec_max }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rec_d8 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rec_d7 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rec_d6 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rec_d5 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rec_d4 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rec_d3 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rec_d2 }}</td>
                <td style="padding: 4px 5px; background-color: #fffdf5; font-weight: 700; color: #0f172a;">{{ cdata.rec_d1 }}</td>
              </tr>
              {% endfor %}
              <tr style="background-color: #fffdf9; font-weight: 800; border-top: 2px solid #ff7900; border-bottom: 2px solid #ff7900; color: #0f172a;">
                <td align="left" style="padding: 6px 8px;">Total</td>
                <td style="padding: 6px 5px;">{{ total_entradas.rec_media }}</td>
                <td style="padding: 6px 5px;">{{ total_entradas.rec_max }}</td>
                <td style="padding: 6px 5px;">{{ total_entradas.rec_d8 }}</td>
                <td style="padding: 6px 5px;">{{ total_entradas.rec_d7 }}</td>
                <td style="padding: 6px 5px;">{{ total_entradas.rec_d6 }}</td>
                <td style="padding: 6px 5px;">{{ total_entradas.rec_d5 }}</td>
                <td style="padding: 6px 5px;">{{ total_entradas.rec_d4 }}</td>
                <td style="padding: 6px 5px;">{{ total_entradas.rec_d3 }}</td>
                <td style="padding: 6px 5px;">{{ total_entradas.rec_d2 }}</td>
                <td style="padding: 6px 5px; background-color: #fef08a; color: #0f172a;">{{ total_entradas.rec_d1 }}</td>
              </tr>
            </tbody>
          </table>

          <!-- TABLA 3: CIERRES (D-8 A D-1) -->
          <div style="font-size: 12px; font-weight: 800; color: #0f172a; margin-bottom: 4px;">
            Cierres (Evolución Diaria de Realizadas D-8 a D-1 por Contrata)
          </div>
          <table width="100%" cellpadding="4" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 10.5px; text-align: right; margin-bottom: 14px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #f1f5f9; border-bottom: 2px solid #94a3b8; color: #334155;">
                <th align="left" style="padding: 6px 8px; color: #0f172a; font-weight: 700;">Contrata</th>
                <th style="padding: 6px 5px; font-weight: 700;">Media</th>
                <th style="padding: 6px 5px; font-weight: 700;">Max</th>
                <th style="padding: 6px 5px;">D-8</th>
                <th style="padding: 6px 5px;">D-7</th>
                <th style="padding: 6px 5px;">D-6</th>
                <th style="padding: 6px 5px;">D-5</th>
                <th style="padding: 6px 5px;">D-4</th>
                <th style="padding: 6px 5px;">D-3</th>
                <th style="padding: 6px 5px;">D-2</th>
                <th style="padding: 6px 5px; background-color: #e2e8f0; font-weight: 700; color: #0f172a;">D-1</th>
              </tr>
            </thead>
            <tbody>
              {% for contrata, cdata in timeline_cierres.items() %}
              <tr style="border-bottom: 1px solid #f1f5f9;">
                <td align="left" style="padding: 4px 8px; font-weight: 600; color: #1e293b;">{{ contrata }}</td>
                <td style="padding: 4px 5px; font-weight: 600;">{{ cdata.rea_media }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rea_max }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rea_d8 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rea_d7 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rea_d6 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rea_d5 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rea_d4 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rea_d3 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.rea_d2 }}</td>
                <td style="padding: 4px 5px; background-color: #f8fafc; font-weight: 700; color: #0f172a;">{{ cdata.rea_d1 }}</td>
              </tr>
              {% endfor %}
              <tr style="background-color: #0f172a; font-weight: 800; border-top: 2px solid #334155; color: #ffffff;">
                <td align="left" style="padding: 6px 8px; color: #ffffff;">Total</td>
                <td style="padding: 6px 5px; color: #ffffff;">{{ total_cierres.rea_media }}</td>
                <td style="padding: 6px 5px; color: #ffffff;">{{ total_cierres.rea_max }}</td>
                <td style="padding: 6px 5px; color: #ffffff;">{{ total_cierres.rea_d8 }}</td>
                <td style="padding: 6px 5px; color: #ffffff;">{{ total_cierres.rea_d7 }}</td>
                <td style="padding: 6px 5px; color: #ffffff;">{{ total_cierres.rea_d6 }}</td>
                <td style="padding: 6px 5px; color: #ffffff;">{{ total_cierres.rea_d5 }}</td>
                <td style="padding: 6px 5px; color: #ffffff;">{{ total_cierres.rea_d4 }}</td>
                <td style="padding: 6px 5px; color: #ffffff;">{{ total_cierres.rea_d3 }}</td>
                <td style="padding: 6px 5px; color: #ffffff;">{{ total_cierres.rea_d2 }}</td>
                <td style="padding: 6px 5px; background-color: #334155; color: #fef08a;">{{ total_cierres.rea_d1 }}</td>
              </tr>
            </tbody>
          </table>

          <!-- TABLA 4: ESTADO DE CITAS Y CALIDAD -->
          <div style="font-size: 12px; font-weight: 800; color: #0f172a; margin-bottom: 4px;">
            Estado de Citas, Capacidad y Calidad de Bolsa por Contrata
          </div>
          <table width="100%" cellpadding="4" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #ebdcd0; border-radius: 4px; font-size: 10.5px; text-align: right; margin-bottom: 14px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #fff8f3; border-bottom: 2px solid #ffcc99; color: #475569;">
                <th align="left" style="padding: 6px 8px; color: #0f172a; font-weight: 700;">Contrata</th>
                <th style="padding: 6px 6px; font-weight: 800; color: #0f172a;">Bolsa</th>
                <th style="padding: 6px 5px;">Citado D</th>
                <th style="padding: 6px 5px;">Cita D+1</th>
                <th style="padding: 6px 5px;">Cita D+2</th>
                <th style="padding: 6px 5px; background-color: #fff3e6; font-weight: 700; color: #0f172a;">IR Previsto</th>
                <th style="padding: 6px 5px; font-weight: 700;">Obj. Inst</th>
                <th style="padding: 6px 5px;">% Pdte Cita</th>
                <th style="padding: 6px 5px;">% Pospuesto</th>
              </tr>
            </thead>
            <tbody>
              {% for contrata, cdata in citas_contratas.items() %}
              <tr style="border-bottom: 1px solid #f8eee7;">
                <td align="left" style="padding: 4px 8px; font-weight: 600; color: #1e293b;">{{ contrata }}</td>
                <td style="padding: 4px 6px; font-weight: 800; color: #0f172a;">{{ cdata.bolsa }}</td>
                <td style="padding: 4px 5px;">{{ cdata.citado }}</td>
                <td style="padding: 4px 5px;">{{ cdata.citad1 }}</td>
                <td style="padding: 4px 5px;">{{ cdata.citad2 }}</td>
                <td style="padding: 4px 5px; background-color: #fffdf5; font-weight: 700; color: #0f172a;">{{ cdata.ir_previsto }}</td>
                <td style="padding: 4px 5px; font-weight: 700;">{{ cdata.n_inst_bjetivo }}</td>
                <td style="padding: 4px 5px;">{{ cdata.per_pdte_cita }}%</td>
                <td style="padding: 4px 5px;">{{ cdata.per_pospuesto }}%</td>
              </tr>
              {% endfor %}
              <tr style="background-color: #fffdf9; font-weight: 800; border-top: 2px solid #ff7900; border-bottom: 2px solid #ff7900; color: #0f172a;">
                <td align="left" style="padding: 6px 8px;">Total</td>
                <td style="padding: 6px 6px; color: #0f172a;">{{ total_citas.bolsa }}</td>
                <td style="padding: 6px 5px;">{{ total_citas.citado }}</td>
                <td style="padding: 6px 5px;">{{ total_citas.citad1 }}</td>
                <td style="padding: 6px 5px;">{{ total_citas.citad2 }}</td>
                <td style="padding: 6px 5px; background-color: #fef08a; color: #0f172a; font-weight: 900;">{{ total_citas.ir_previsto }}</td>
                <td style="padding: 6px 5px; font-weight: 900;">{{ total_citas.n_inst_bjetivo }}</td>
                <td style="padding: 6px 5px;">{{ total_citas.per_pdte_cita }}%</td>
                <td style="padding: 6px 5px;">{{ total_citas.per_pospuesto }}%</td>
              </tr>
            </tbody>
          </table>

          <!-- TABLA 5: RESUMEN GRUPOS DE MARCAS -->
          <div style="font-size: 12px; font-weight: 800; color: #0f172a; margin-bottom: 4px;">
            Resumen Consolidado por Grupo de Marcas (Stock, Citas y Pospuesto)
          </div>
          <table width="100%" cellpadding="4" cellspacing="0" border="0" style="background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 10.5px; text-align: right; margin-bottom: 12px; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #000000; color: #ffffff;">
                <th align="left" style="padding: 6px 8px;">Grupo</th>
                <th style="padding: 6px 5px;">Rec. Media</th>
                <th style="padding: 6px 5px;">Rea. Media</th>
                <th style="padding: 6px 5px;">Rea. D-1</th>
                <th style="padding: 6px 5px;">Bolsa</th>
                <th style="background-color: #1e293b; color: #fef08a; padding: 6px 6px;">IR</th>
                <th style="padding: 6px 5px;">IR Ayer</th>
                <th style="padding: 6px 5px;">Citado D</th>
                <th style="padding: 6px 5px;">Cita D+1</th>
                <th style="padding: 6px 5px;">IR Prev.</th>
                <th style="padding: 6px 5px;">% Pdte Cita</th>
                <th style="padding: 6px 5px;">% Pospuesto</th>
              </tr>
            </thead>
            <tbody>
              {% for g_nombre in ['MAS MOVIL', 'ORANGE', 'KRTV'] %}
              {% if kpis_marcas.get(g_nombre) %}
              {% set m = kpis_marcas[g_nombre] %}
              <tr style="border-bottom: 1px solid #f1f5f9; background-color: #ffffff;">
                <td align="left" style="padding: 5px 8px; font-weight: bold; color: #0f172a;">{{ g_nombre }}</td>
                <td style="padding: 5px 5px;">{{ m.rec_media }}</td>
                <td style="padding: 5px 5px;">{{ m.rea_media }}</td>
                <td style="padding: 5px 5px;">{{ m.rea_d1 }}</td>
                <td style="padding: 5px 5px; font-weight: bold;">{{ m.bolsa }}</td>
                <td style="background-color: #fff9c3; font-weight: 800; color: #0f172a; padding: 5px 6px;">{{ m.ir }}</td>
                <td style="padding: 5px 5px;">{{ m.ir_ayer }}</td>
                <td style="padding: 5px 5px;">{{ m.citado }}</td>
                <td style="padding: 5px 5px;">{{ m.citad1 }}</td>
                <td style="padding: 5px 5px;">{{ m.ir_previsto }}</td>
                <td style="padding: 5px 5px;">{{ m.per_pdte }}%</td>
                <td style="padding: 5px 5px;">{{ m.per_posp }}%</td>
              </tr>
              {% endif %}
              {% endfor %}
              <tr style="background-color: #0f172a; color: #ffffff; font-weight: 800;">
                <td align="left" style="padding: 7px 8px; color: #ffffff;">TOTAL PROYECTO</td>
                <td style="padding: 7px 5px; color: #ffffff;">{{ total_global.rec_media }}</td>
                <td style="padding: 7px 5px; color: #ffffff;">{{ total_global.rea_media }}</td>
                <td style="padding: 7px 5px; color: #ffffff;">{{ total_global.rea_d1 }}</td>
                <td style="padding: 7px 5px; color: #ffffff;">{{ total_global.bolsa }}</td>
                <td style="background-color: #fef08a; color: #0f172a; font-weight: 900; padding: 7px 6px;">{{ total_global.ir }}</td>
                <td style="padding: 7px 5px; color: #ffffff;">{{ total_global.ir_ayer }}</td>
                <td style="padding: 7px 5px; color: #ffffff;">{{ total_global.citado }}</td>
                <td style="padding: 7px 5px; color: #ffffff;">{{ total_global.citad1 }}</td>
                <td style="padding: 7px 5px; color: #ffffff;">{{ total_global.ir_previsto }}</td>
                <td style="padding: 7px 5px; color: #ffffff;">{{ total_global.per_pdte }}%</td>
                <td style="padding: 7px 5px; color: #ffffff;">{{ total_global.per_posp }}%</td>
              </tr>
            </tbody>
          </table>

          <!-- PIE DE CORREO COMPACTO -->
          <p style="font-size: 10px; color: #64748b; margin-top: 10px; line-height: 1.4;">
            <em>* Métricas extraídas de BigQuery (<code>{{ table_name }}</code>) reproduciendo el modelo de Looker Studio.</em><br>
            Cuadro de mando interactivo alojado en <a href="{{ sharepoint_web_url }}" style="color: #ff7900; text-decoration: underline; font-weight: bold;">SharePoint Orange</a>.
          </p>
        </body>
        </html>
        """
        template = Template(template_str)
        return template.render(
            fecha_ayer_formateada=target_date.strftime("%d/%m/%Y"),
            cdm_filas=transformed_data["cdm_filas"],
            total_cdm_resumen=transformed_data["total_cdm_resumen"],
            cdm_noswap=transformed_data["cdm_noswap"],
            cdm_ayt=transformed_data["cdm_ayt"],
            cdm_mig=transformed_data["cdm_mig"],
            cdm_b2b=transformed_data["cdm_b2b"],
            resumen_ir_contratas=transformed_data["resumen_ir_contratas"],
            total_resumen_ir=transformed_data["total_resumen_ir"],
            looker_ir_contratas=transformed_data["looker_ir_contratas"],
            total_looker_ir=transformed_data["total_looker_ir"],
            timeline_entradas=transformed_data["timeline_entradas"],
            total_entradas=transformed_data["total_entradas"],
            timeline_cierres=transformed_data["timeline_cierres"],
            total_cierres=transformed_data["total_cierres"],
            citas_contratas=transformed_data["citas_contratas"],
            total_citas=transformed_data["total_citas"],
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
                logging.info("¡Correo con Cuadros de Mando (CDM) y tablas de Looker enviado exitosamente!")
        except Exception as e:
            logging.error(f"Error al enviar mensaje por SMTP: {e}")
            raise


# ==============================================================================
# 8. ORQUESTADOR DEL PIPELINE GENERAL
# ==============================================================================
class ReportPipeline:
    """Coordina el flujo completo: BigQuery -> HTML Local -> SharePoint Upload -> Email con Looker embebido."""

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

        # 1. Extracción BigQuery con cálculo simultáneo de todos los CDMs y segmentaciones
        query = processor.build_query(target_date)
        raw_rows = self.extractor.execute_query(query)

        # 2. Transformación de datos para construir CDMs y tablas de Looker
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

        # 5. Envío de correo con los 4 CDMs segmentados y el CDM provincial embebidos en el cuerpo
        if skip_email:
            logging.info("Opción --no-email activada. Se omite el envío por SMTP.")
        else:
            body_html = processor.render_email_body(
                transformed,
                target_date,
                sharepoint_web_url
            )
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
        description="Automatización BigQuery -> Tablas Looker en Email + Subida a SharePoint"
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