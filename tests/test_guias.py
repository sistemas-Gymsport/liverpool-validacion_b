import io

import pytest

import generar_ejemplos as ej
import validar_guias as vg
import verificar_pedidos as vp
from conftest import Sesion, excel_bytes, pdf_bytes, pedidos_base, renglones_base


def medidas_bytes(filas=None, **kwargs) -> bytes:
    buffer = io.BytesIO()
    ej.crear_medidas(buffer, filas if filas is not None else ej.MEDIDAS_BASE, **kwargs)
    return buffer.getvalue()


def correr(pedidos=None, renglones=None, medidas=b"usar-base"):
    medidas = medidas_bytes() if medidas == b"usar-base" else medidas
    return vp.verificar(excel_bytes(pedidos), [pdf_bytes(renglones)], medidas)


def revision(resultado, pedido):
    return next(g for g in resultado.guias if g.pedido == pedido)


def test_escenario_correcto_con_medidas():
    r = correr()
    assert (r.errores, r.revisar, r.verificados, r.sin_picking, r.guias_por_revisar) == (0, 0, 4, 1, 0)
    assert r.medidas_cargadas and r.todo_cuadra
    assert all(g.resultado == vp.CORRECTO for g in r.guias)


def test_escenario_correcto_sin_medidas_no_bloquea_todo_cuadra():
    r = correr(medidas=None)
    assert (r.errores, r.revisar) == (0, 0)
    assert not r.medidas_cargadas and r.todo_cuadra


def test_lee_tablas_lado_a_lado_con_espacios_encabezados_y_duplicados():
    filas = ej.MEDIDAS_BASE + [
        (" GS-ESPACIO ", 1, 5, 20, 20, 20),
        ("GS-DOBLE", 1, 10, 50, 20, 20),
        ("GS-DOBLE", 1, 12, 80, 30, 20),
        ("GS-REPETIDO", 1, 3, 10, 20, 30),
        ("GS-REPETIDO", 1, 3, 10, 20, 30),
        ("GS-RACK", 3, 40, 200, 40, 20),
        ("GS-RACK", 3, 35, 120, 40, 20),
        ("GS-ENCABEZADO", "guia", "peso", "largo", "ancho", "alto"),
    ]
    referencias = vg.leer_medidas(medidas_bytes(filas))
    assert {"GS-MANC10", "GS-BARRA20", "GS-ESPACIO"} <= set(referencias)
    assert referencias["GS-DOBLE"].ambigua
    assert not referencias["GS-REPETIDO"].ambigua
    assert referencias["GS-RACK"].guias == 3 and not referencias["GS-RACK"].ambigua
    assert "GS-ENCABEZADO" not in referencias


@pytest.mark.parametrize("paquete", [(10, 10, 10, 1), (0, 0, 0, 0)])
def test_medidas_genericas_de_liverpool(paquete):
    pedidos = pedidos_base()
    pedidos[3].paquetes = [paquete]
    r = correr(pedidos)
    g = revision(r, "1000000004")
    assert g.resultado == vp.REVISAR
    assert any("genéricas o vacías" in o for o in g.observaciones)
    assert any(h.pedido == "1000000004" and h.origen == vp.ORIGEN_EXCEL and h.problema.startswith("Datos de paquetes")
               for h in r.hallazgos)
    assert r.errores == 0


def test_liverpool_sin_paquetes():
    pedidos = pedidos_base()
    pedidos[0].paquetes = []
    g = revision(correr(pedidos), "1000000001")
    assert g.resultado == vp.REVISAR and any("no tiene paquetes" in o for o in g.observaciones)


def test_par_de_pesas_de_17kg_o_mas_lleva_una_guia_por_unidad_sin_referencia():
    pedidos = [ej.PedidoFicticio("2000000001", "GS-HEX17.5", 1, "Ana Prueba", "Set mancuernas de peso fijo 17.5 kg 2 piezas",
                                 paquetes=[(30, 20, 30, 35)])]
    renglones = [ej.RenglonFicticio("GS-HEX17.5", 1, "PAR", "Mancuerna 17.5 kg (PAR)", ["2000000001/Ana Prueba"])]
    g = revision(correr(pedidos, renglones, medidas=None), "2000000001")
    assert g.guias_estimadas == "2"
    assert any("se necesitan al menos 2" in o for o in g.observaciones)


def test_par_de_pesas_ligeras_comparte_guia():
    pedidos = [ej.PedidoFicticio("2000000001", "GS-HEX10", 1, "Ana Prueba", "Set mancuernas de peso fijo 10 kg 2 piezas",
                                 paquetes=[(30, 20, 30, 20)])]
    renglones = [ej.RenglonFicticio("GS-HEX10", 1, "PAR", "Mancuerna 10 kg (PAR)", ["2000000001/Ana Prueba"])]
    g = revision(correr(pedidos, renglones, medidas=None), "2000000001")
    assert g.guias_estimadas == "1" and g.resultado == vp.CORRECTO


def test_metros_de_piso_dos_por_guia():
    pedidos = [ej.PedidoFicticio("2000000001", "GS-PISO-N7", 3, "Ana Prueba", "Piso de hule 1 m",
                                 paquetes=[(50, 50, 50, 10)])]
    renglones = [ej.RenglonFicticio("GS-PISO-N7", 3, "UNIDAD", "Piso de hule", ["2000000001/Ana Prueba"])]
    g = revision(correr(pedidos, renglones, medidas=medidas_bytes([("GS-PISO-N7", 1, 5, 50, 50, 50)])), "2000000001")
    assert g.guias_estimadas == "2 a 3"
    assert any("al menos 2" in o for o in g.observaciones)


def test_productos_similares_hasta_30_kg_comparten_guia():
    medidas = medidas_bytes([("GS-STEP", 1, 12, 90, 15, 35)])
    pedidos = [ej.PedidoFicticio("2000000001", "GS-STEP", 2, "Ana Prueba", "Step", paquetes=[(90, 30, 35, 24)])]
    renglones = [ej.RenglonFicticio("GS-STEP", 2, "UNIDAD", "Step", ["2000000001/Ana Prueba"])]
    assert revision(correr(pedidos, renglones, medidas), "2000000001").resultado == vp.CORRECTO

    pedidos[0].cantidad, renglones[0].cantidad = 5, 5
    pedidos[0].paquetes = [(90, 30, 35, 60)]
    g = revision(correr(pedidos, renglones, medidas), "2000000001")
    assert g.guias_estimadas == "3 a 5"
    assert any("al menos 3" in o for o in g.observaciones)


def test_equipo_que_requiere_varias_guias():
    medidas = medidas_bytes([("GS-RACK", 7, 30, 200, 40, 20)])
    pedidos = [ej.PedidoFicticio("2000000001", "GS-RACK", 1, "Ana Prueba", "Rack", paquetes=[(200, 40, 20, 30)] * 7)]
    renglones = [ej.RenglonFicticio("GS-RACK", 1, "UNIDAD", "Rack", ["2000000001/Ana Prueba"])]
    assert revision(correr(pedidos, renglones, medidas), "2000000001").resultado == vp.CORRECTO
    pedidos[0].paquetes = [(200, 40, 20, 30)] * 2
    g = revision(correr(pedidos, renglones, medidas), "2000000001")
    assert any("al menos 7" in o for o in g.observaciones)


def test_medidas_de_liverpool_distintas_a_la_tabla():
    pedidos = pedidos_base()
    pedidos[4].paquetes = [(15, 150, 15, 20)]
    g = revision(correr(pedidos), "1000000005")
    assert g.resultado == vp.CORRECTO
    pedidos[4].paquetes = [(100, 15, 15, 12)]
    g = revision(correr(pedidos), "1000000005")
    assert g.resultado == vp.REVISAR
    assert any("no coinciden" in o and "100×15×15" in o and "150×15×15" in o for o in g.observaciones)
    assert not g.en_revision
    assert not any(h.pedido == "1000000005" for h in correr(pedidos).hallazgos)


def test_pedido_con_varios_productos_no_es_duplicado():
    paquetes = [(30, 20, 30, 20), (40, 10, 30, 2)]
    pedidos = [
        ej.PedidoFicticio("3000000001", "GS-MANC10", 1, "Ana Prueba", "Mancuerna", paquetes=paquetes, guias="2"),
        ej.PedidoFicticio("3000000001", "GS-BANDALARGA-ROJA", 1, "Ana Prueba", "Banda", paquetes=paquetes, guias="2"),
    ]
    renglones = [
        ej.RenglonFicticio("GS-MANC10", 1, "PAR", "Mancuerna", ["3000000001/Ana Prueba"]),
        ej.RenglonFicticio("GS-BANDALARGA-ROJA", 1, "UNIDAD", "Banda", ["3000000001/Ana Prueba"], partir_sku=True),
    ]
    r = correr(pedidos, renglones)
    assert (r.errores, r.revisar, r.pedidos_liverpool, r.verificados, r.correctos) == (0, 0, 1, 1, 1)
    assert len(r.guias) == 1
    assert r.guias[0].productos == "GS-MANC10×1, GS-BANDALARGA-ROJA×1"
    assert all(d.productos_pedido == 2 for d in r.detalle)
    assert not any("duplicado" in h.problema for h in r.hallazgos)


def test_pedido_con_paquetes_inconsistentes_entre_renglones():
    pedidos = [
        ej.PedidoFicticio("3000000001", "GS-MANC10", 1, "Ana Prueba", "Mancuerna", paquetes=[(30, 20, 30, 20)]),
        ej.PedidoFicticio("3000000001", "GS-TAPETE", 1, "Ana Prueba", "Tapete", paquetes=[(60, 15, 15, 2)]),
    ]
    renglones = [ej.RenglonFicticio("GS-MANC10", 1, "PAR", "Mancuerna", ["3000000001/Ana Prueba"]),
                 ej.RenglonFicticio("GS-TAPETE", 1, "UNIDAD", "Tapete", ["3000000001/Ana Prueba"])]
    g = correr(pedidos, renglones).guias[0]
    assert any("paquetes distintos entre sí" in o for o in g.observaciones)


def test_tabla_de_medidas_invalida():
    with pytest.raises(vp.DocumentoInvalido, match="medidas"):
        correr(medidas=excel_bytes(encabezados=["Sin", "Equipos"]))


def test_api_acepta_tabla_de_medidas(cliente):
    s = Sesion(cliente)
    s.login()
    datos = {
        "excel": (io.BytesIO(excel_bytes()), "reporte.xlsx"),
        "pdfs": [(io.BytesIO(pdf_bytes()), "picking.pdf")],
        "medidas": (io.BytesIO(medidas_bytes()), "MEDIDAS EQUIPOS.xlsx"),
    }
    r = s.post("/api/verificar", data=datos, content_type="multipart/form-data")
    assert r.status_code == 200
    cuerpo = r.get_json()
    assert cuerpo["medidas_cargadas"] is True and cuerpo["resumen"]["guias_por_revisar"] == 0
    datos["medidas"] = (io.BytesIO(b"no es excel"), "medidas.xlsx")
    datos["excel"] = (io.BytesIO(excel_bytes()), "reporte.xlsx")
    datos["pdfs"] = [(io.BytesIO(pdf_bytes()), "picking.pdf")]
    assert s.post("/api/verificar", data=datos, content_type="multipart/form-data").status_code == 400
