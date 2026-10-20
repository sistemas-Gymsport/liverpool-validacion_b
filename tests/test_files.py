import io
import zipfile

from conftest import configuracion, excel_bytes, pdf_bytes
from factory import create_app
from conftest import Sesion


def _error(respuesta):
    return respuesta.get_json()["error"]


def test_extension_invalida(sesion, documentos):
    excel, pdf = documentos
    r = sesion.verificar(excel, [("picking.pdf", pdf)], nombre_excel="pedidos.csv")
    assert r.status_code == 400 and "solo se aceptan archivos Excel (.xlsx) y PDF" in _error(r)
    r = sesion.verificar(excel, [("picking.txt", pdf)])
    assert r.status_code == 400 and "PDF" in _error(r)


def test_firma_invalida(sesion, documentos):
    excel, pdf = documentos
    r = sesion.verificar(b"no es un zip" * 10, [("picking.pdf", pdf)])
    assert r.status_code == 400
    r = sesion.verificar(excel, [("picking.pdf", b"no es un pdf" * 10)])
    assert r.status_code == 400 and "dañado o no puede leerse" in _error(r)


def test_pdf_corrupto(sesion, documentos):
    excel, _ = documentos
    r = sesion.verificar(excel, [("picking.pdf", b"%PDF-1.4\n" + b"\x00" * 200)])
    assert r.status_code == 422 and "dañado o no puede leerse" in _error(r)


def test_excel_corrupto(sesion, documentos):
    _, pdf = documentos
    r = sesion.verificar(b"PK\x03\x04" + b"\x00" * 300, [("picking.pdf", pdf)])
    assert r.status_code == 400 and "Excel no tiene un formato válido" in _error(r)


def test_excel_sin_estructura_liverpool(sesion, documentos):
    _, pdf = documentos
    r = sesion.verificar(excel_bytes(encabezados=["A", "B"]), [("picking.pdf", pdf)])
    assert r.status_code == 422 and _error(r).startswith("El Excel no parece ser el reporte de pedidos de Liverpool.")


def test_excel_bomba_zip(sesion, documentos):
    _, pdf = documentos
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("xl/workbook.xml", "<workbook/>")
        z.writestr("xl/worksheets/sheet1.xml", "0" * (20 * 1024 * 1024))
    r = sesion.verificar(buffer.getvalue(), [("picking.pdf", pdf)])
    assert r.status_code == 400 and "comprimido" in _error(r)


def test_archivo_vacio(sesion, documentos):
    excel, _ = documentos
    r = sesion.verificar(excel, [("picking.pdf", b"")])
    assert r.status_code == 400 and "vacío" in _error(r)


def test_archivo_demasiado_grande(directorio_temporal, documentos):
    app = create_app(configuracion(max_pdf_mb=1), directorio_temporal=str(directorio_temporal), limpieza_periodica=False)
    s = Sesion(app.test_client())
    s.login()
    excel, pdf = documentos
    r = s.verificar(excel, [("grande.pdf", pdf + b"0" * (1024 * 1024 + 10))])
    assert r.status_code == 413 and "tamaño máximo" in _error(r)


def test_peticion_total_demasiado_grande(directorio_temporal, documentos):
    app = create_app(configuracion(max_total_mb=1), directorio_temporal=str(directorio_temporal), limpieza_periodica=False)
    s = Sesion(app.test_client())
    s.login()
    excel, pdf = documentos
    r = s.verificar(excel, [("grande.pdf", pdf + b"0" * (4 * 1024 * 1024))])
    assert r.status_code == 413 and r.get_json()["error"].startswith("El archivo supera el tamaño máximo")


def test_mas_de_diez_pdf(sesion, documentos):
    excel, pdf = documentos
    r = sesion.verificar(excel, [(f"p{i}.pdf", pdf) for i in range(11)])
    assert r.status_code == 400 and _error(r) == "Puedes subir un máximo de 10 documentos PDF."


def test_sin_csrf_en_verificacion(cliente, sesion, documentos):
    excel, pdf = documentos
    r = sesion.post("/api/verificar", headers={"X-CSRF-Token": ""},
                    data={"excel": (io.BytesIO(excel), "a.xlsx"), "pdfs": [(io.BytesIO(pdf), "b.pdf")]},
                    content_type="multipart/form-data")
    assert r.status_code == 403


def test_nombre_de_archivo_malicioso(sesion, documentos):
    excel, pdf = documentos
    r = sesion.verificar(excel, [("../../etc/<script>alert(1)</script>.pdf", pdf)],
                         nombre_excel="..\\..\\pedidos.xlsx")
    assert r.status_code == 200
    datos = r.get_json()
    assert datos["resumen"]["errores"] == 0


def test_texto_con_posible_xss_se_devuelve_como_dato(sesion):
    import generar_ejemplos as ej
    pedidos = [ej.PedidoFicticio("1000000001", "GS-MANC10", 1, "<img src=x onerror=alert(1)>")]
    renglones = [ej.RenglonFicticio("GS-MANC10", 1, "PAR", "Mancuerna", ["1000000001/<b>alguien</b>"])]
    r = sesion.verificar(excel_bytes(pedidos), [("p.pdf", pdf_bytes(renglones))])
    assert r.status_code == 200
    assert r.headers["Content-Type"].startswith("application/json")
    assert r.headers["X-Content-Type-Options"] == "nosniff"
    detalle = r.get_json()["detalle_pedidos"][0]
    assert detalle["cliente_liverpool"] == "<img src=x onerror=alert(1)>"


def test_documentos_se_eliminan_tras_procesar(sesion, documentos, directorio_temporal):
    excel, pdf = documentos
    r = sesion.verificar(excel, [("picking.pdf", pdf)])
    assert r.status_code == 200
    assert list(directorio_temporal.iterdir()) == []


def test_documentos_se_eliminan_si_falla(sesion, directorio_temporal, documentos):
    excel, _ = documentos
    r = sesion.verificar(excel, [("picking.pdf", b"%PDF-1.4\n" + b"\x00" * (2 * 1024 * 1024))])
    assert r.status_code == 422
    assert list(directorio_temporal.iterdir()) == []
