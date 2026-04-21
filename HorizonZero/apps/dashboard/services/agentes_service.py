# ═══════════════════════════════════════════════════════════════════
# ARCHIVO 3: apps/dashboard/services/agentes_service.py
# ═══════════════════════════════════════════════════════════════════

import io
import base64
import qrcode
from qrcode.image.svg import SvgImage

from .db_service import ReportesDBService


# ── URL base de encuestas ────────────────────────────────────────
BASE_URL = "https://mercadeo.cajadeande.fi.cr/encuestas/index.php"

# Mapeo encuesta_id → nombre legible
ENCUESTAS_QR = {
    1:  "Satisfacción (Agentes)",
    4:  "Oficina Digital",
    17: "WhatsApp con Agente",
}

# Mapeo unidad_nombre → encuesta_id
UNIDAD_ENCUESTA = {
    "oficina digital": 4,
    "whatsapp":        17,
}

# Nombres legibles por encuesta_id
ENCUESTA_NOMBRE = {
    1:  "Satisfacción (Agentes)",
    4:  "Oficina Digital",
    17: "WhatsApp con Agente",
}


def build_encuesta_url(encuesta_id: int, agente_id: int) -> str:
    return f"{BASE_URL}?id={encuesta_id}&age_id={agente_id}"


def generar_qr_base64(url: str, size: int = 10) -> str:
    """Genera QR como PNG codificado en base64 para incrustar en HTML."""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=size,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#003FB7", back_color="white")
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode()


class AgentesService:

    @staticmethod
    def listar(buscar: str = "", unidad_id: int = None, sucursal_id: int = None):
        conditions = ["1=1"]    # ← quitar "a.activo = 1" de aquí
        if buscar:
            buscar_safe = buscar.replace("'", "''")
            conditions.append(
                f"(a.nombre LIKE '%{buscar_safe}%' OR a.nombre_lower LIKE '%{buscar_safe.lower()}%'"
                f" OR a.login_ad LIKE '%{buscar_safe.lower()}%')"
            )
        if unidad_id:
            conditions.append(f"a.unidad_id = {int(unidad_id)}")
        if sucursal_id:
            conditions.append(f"a.sucursal_id = {int(sucursal_id)}")

        sql = f"""
            SELECT a.agente_id, a.nombre, a.sucursal_id, a.unidad_id,
                ISNULL(a.login_ad, 'pendiente') AS login_ad,
                s.nombre AS sucursal_nombre,
                u.nombre AS unidad_nombre
            FROM (
                SELECT ag.agente_id, ag.nombre, ag.sucursal_id, ag.unidad_id,
                    ag.nombre_lower, ag.login_ad
                FROM dbo.agentes ag
                WHERE ag.activo = 1    -- ← el filtro ya está aquí adentro
            ) a
            JOIN dbo.sucursales s ON s.sucursal_id = a.sucursal_id
            JOIN dbo.unidades   u ON u.unidad_id   = a.unidad_id
            WHERE  {' AND '.join(conditions)}
            ORDER  BY u.nombre, s.nombre, a.nombre
        """
        return ReportesDBService.ejecutar_query(sql)


    @staticmethod
    def obtener(agente_id: int):
        sql = f"""
            SELECT a.agente_id, a.nombre, a.sucursal_id, a.unidad_id,
                ISNULL(a.login_ad, 'pendiente') AS login_ad,
                s.nombre AS sucursal_nombre,
                u.nombre AS unidad_nombre
            FROM (
                SELECT ag.agente_id, ag.nombre, ag.sucursal_id, ag.unidad_id,
                    ag.nombre_lower, ag.login_ad
                FROM dbo.agentes ag
                WHERE ag.activo = 1
            ) a
            JOIN dbo.sucursales s ON s.sucursal_id = a.sucursal_id
            JOIN dbo.unidades   u ON u.unidad_id   = a.unidad_id
            WHERE  a.agente_id = {int(agente_id)}
        """
        rows = ReportesDBService.ejecutar_query(sql)
        return rows[0] if rows else None
    
    @staticmethod
    def buscar_por_login_ad(login_ad: str):
        """Busca si ya existe un agente activo con ese login_ad."""
        login_safe = login_ad.strip().lower().replace("'", "''")
        sql = f"""
            SELECT a.agente_id, a.nombre, a.sucursal_id, a.unidad_id,
                a.login_ad,
                s.nombre AS sucursal_nombre,
                u.nombre AS unidad_nombre
            FROM dbo.agentes a
            JOIN dbo.sucursales s ON s.sucursal_id = a.sucursal_id
            JOIN dbo.unidades   u ON u.unidad_id   = a.unidad_id
            WHERE  a.login_ad = '{login_safe}'
            AND    a.activo   = 1
        """
        rows = ReportesDBService.ejecutar_query(sql)
        return rows[0] if rows else None

    @staticmethod
    def crear(nombre: str, sucursal_id: int, unidad_id: int, login_ad: str = "pendiente") -> int:
        nombre_safe       = nombre.strip().replace("'", "''")
        nombre_lower_safe = nombre.strip().lower().replace("'", "''")
        login_safe        = (login_ad or "pendiente").strip().lower().replace("'", "''")
        insert_sql = f"""
            INSERT INTO dbo.agentes (nombre, sucursal_id, unidad_id, nombre_lower, activo, login_ad)
            VALUES ('{nombre_safe}', {int(sucursal_id)}, {int(unidad_id)},
                    '{nombre_lower_safe}', 1, '{login_safe}')
        """
        ReportesDBService.ejecutar_comando(insert_sql)
        rows = ReportesDBService.ejecutar_query("SELECT CAST(SCOPE_IDENTITY() AS INT) AS agente_id")
        return rows[0]["agente_id"] if rows else None

    @staticmethod
    def actualizar(agente_id: int, nombre: str, sucursal_id: int, unidad_id: int, login_ad: str = None):
        nombre_safe       = nombre.strip().replace("'", "''")
        nombre_lower_safe = nombre.strip().lower().replace("'", "''")
        login_safe        = (login_ad or "pendiente").strip().lower().replace("'", "''")
        sql = f"""
            UPDATE dbo.agentes
            SET    nombre       = '{nombre_safe}',
                sucursal_id  = {int(sucursal_id)},
                unidad_id    = {int(unidad_id)},
                nombre_lower = '{nombre_lower_safe}',
                login_ad     = '{login_safe}'
            WHERE  agente_id    = {int(agente_id)}
        """
        ReportesDBService.ejecutar_comando(sql)

    @staticmethod
    def existe(nombre: str, excluir_id: int = None) -> bool:
        nombre_lower = nombre.strip().lower().replace("'", "''")
        sql = f"SELECT COUNT(*) AS total FROM dbo.agentes WHERE nombre_lower = '{nombre_lower}' AND activo = 1"
        if excluir_id:
            sql += f" AND agente_id != {int(excluir_id)}"
        rows = ReportesDBService.ejecutar_query(sql)
        return rows[0]["total"] > 0

    # Eliminar — soft delete, cambia activo a 0
    @staticmethod
    def eliminar(agente_id: int):
        ReportesDBService.ejecutar_comando(
            f"UPDATE dbo.agentes SET activo = 0 WHERE agente_id = {int(agente_id)}"
        )


    @staticmethod
    def sucursales():
        return ReportesDBService.ejecutar_query(
            "SELECT sucursal_id, nombre FROM dbo.sucursales WHERE activo=1 ORDER BY orden"
        )

    @staticmethod
    def unidades():
        return ReportesDBService.ejecutar_query(
            "SELECT unidad_id, nombre FROM dbo.unidades WHERE activo=1 ORDER BY nombre"
        )

    # KPIs — solo activos
    @staticmethod
    def kpis():
        sql = """
            SELECT
                COUNT(*)                                           AS total,
                COUNT(DISTINCT a.unidad_id)                       AS total_unidades,
                COUNT(DISTINCT a.sucursal_id)                     AS total_sucursales,
                SUM(CASE WHEN u.es_whatsapp = 1 THEN 1 ELSE 0 END) AS total_whatsapp
            FROM dbo.agentes a
            JOIN dbo.unidades u ON u.unidad_id = a.unidad_id
            WHERE a.activo = 1
        """
        rows = ReportesDBService.ejecutar_query(sql)
        return rows[0] if rows else {}

    # Conteo por unidad — solo activos
    @staticmethod
    def conteo_por_unidad():
        sql = """
            SELECT u.nombre AS unidad, COUNT(*) AS total
            FROM   dbo.agentes a
            JOIN   dbo.unidades u ON u.unidad_id = a.unidad_id
            WHERE  a.activo = 1
            GROUP  BY u.nombre
            ORDER  BY total DESC
        """
        return ReportesDBService.ejecutar_query(sql)



    # ── QR helpers ──────────────────────────────────────────────
    @staticmethod
    def qr_data(agente_id: int, unidad_nombre: str):
        """
        Retorna lista con UN solo QR según la unidad del agente.
        - Oficina Digital → encuesta 4
        - Whatsapp        → encuesta 17
        - Resto           → encuesta 1
        """
        unidad_key = (unidad_nombre or "").strip().lower()
        encuesta_id = UNIDAD_ENCUESTA.get(unidad_key, 1)
        enc_nombre  = ENCUESTA_NOMBRE[encuesta_id]

        url    = build_encuesta_url(encuesta_id, agente_id)
        qr_b64 = generar_qr_base64(url)

        return [{
            "encuesta_id": encuesta_id,
            "nombre":      enc_nombre,
            "url":         url,
            "qr_b64":      qr_b64,
        }]
    

    @staticmethod
    def listar_inactivos(buscar: str = "", unidad_id: int = None, sucursal_id: int = None):
        conditions = ["1=1"]

        if buscar:
            buscar_safe = buscar.replace("'", "''")
            conditions.append(
                f"(a.nombre LIKE '%{buscar_safe}%' OR a.nombre_lower LIKE '%{buscar_safe.lower()}%')"
            )
        if unidad_id:
            conditions.append(f"a.unidad_id = {int(unidad_id)}")
        if sucursal_id:
            conditions.append(f"a.sucursal_id = {int(sucursal_id)}")

        sql = f"""
            SELECT a.agente_id, a.nombre, a.sucursal_id, a.unidad_id,
                s.nombre AS sucursal_nombre,
                u.nombre AS unidad_nombre
            FROM (
                SELECT ag.agente_id, ag.nombre, ag.sucursal_id, ag.unidad_id, ag.nombre_lower
                FROM dbo.agentes ag
                WHERE ag.activo = 0
            ) a
            JOIN dbo.sucursales s ON s.sucursal_id = a.sucursal_id
            JOIN dbo.unidades   u ON u.unidad_id   = a.unidad_id
            WHERE  {' AND '.join(conditions)}
            ORDER  BY u.nombre, s.nombre, a.nombre
        """
        return ReportesDBService.ejecutar_query(sql)

    @staticmethod
    def restaurar(agente_id: int):
        ReportesDBService.ejecutar_comando(
            f"UPDATE dbo.agentes SET activo = 1 WHERE agente_id = {int(agente_id)}"
        )