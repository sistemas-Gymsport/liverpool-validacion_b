"""Lectura de PDF de guías de envío (Estafeta, FedEx u otras) generadas para pedidos de Liverpool.

De cada página (una guía) se extrae: referencia del pedido (número de pedido de 10 dígitos seguido de
`Q...`), paquetería, número de rastreo, peso, medidas y la nota de contenido que escribe el personal,
por ejemplo `1 PAR IF-V6`, `1 PIEZA IF-HK17.5` o `2 Piezas (1 Par) IF-50/2LB`.
Las páginas escaneadas (imagen) pueden no tener la referencia del pedido; en ese caso el pedido se
deduce después por el SKU de la nota, si hay un único candidato.
"""

from __future__ import annotations

import io
import logging
import re
from dataclasses import dataclass, field

import pdfplumber

logging.getLogger("pdfminer").setLevel(logging.ERROR)

RE_REFERENCIA = re.compile(r"(?<!\d)(\d{10})Q[0-9A-Z]{4,}")
RE_PESO_FEDEX = re.compile(r"ACTWGT:\s*(\d+(?:\.\d+)?)\s*KG", re.I)
RE_PESO = re.compile(r"(?<![\d.])(\d{1,3}(?:\.\d+)?)\s*KG\b", re.I)
RE_DIMS = re.compile(r"DIMS:\s*(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)\s*CM", re.I)
RE_TRK_FEDEX = re.compile(r"\b(\d{4} \d{4} \d{4})\b")
RE_RASTREO = re.compile(r"C[oó]digo de Rastreo:\s*(\d{6,})", re.I)
RE_CONTENIDO = re.compile(
    r"(?<![\w.])(\d+)\s*(pares|par|piezas|pieza|pzas|pza|pz|unidades|unidad)\b\.?\s*"
    r"(?:\(\s*(\d+)\s*(?:pares|par)\s*\)\s*)?"
    r"([A-Z]{1,4}-[A-Z0-9./-]*[A-Z0-9])",
    re.I,
)


@dataclass
class Contenido:
    cantidad: int
    unidad: str
    pares: int | None
    sku: str

    def texto(self) -> str:
        pares = f" ({self.pares} PAR)" if self.pares is not None else ""
        return f"{self.cantidad} {self.unidad}{pares} {self.sku}"


@dataclass
class EtiquetaEnvio:
    archivo: str
    pagina: int
    paqueteria: str
    pedido: str
    rastreo: str
    peso: float | None
    medidas: tuple[float, float, float] | None
    contenido: list[Contenido] = field(default_factory=list)
    legible: bool = True
    pedido_por_sku: bool = False

    @property
    def texto_contenido(self) -> str:
        return ", ".join(c.texto() for c in self.contenido)


def _paqueteria(texto: str) -> str:
    if "ORIGIN ID" in texto or "FEDEX" in texto.upper():
        return "FedEx"
    if "ESTAFETA" in texto.upper() or ("VASCONCELOS" in texto.upper() and RE_RASTREO.search(texto)):
        return "Estafeta"
    if "DHL" in texto.upper():
        return "DHL"
    return "No identificada"


def _unidad(texto: str) -> str:
    texto = texto.lower()
    if texto.startswith("par"):
        return "PAR"
    if texto.startswith(("pieza", "pza", "pz")):
        return "PIEZA"
    return "UNIDAD"


def contenido_de(texto: str) -> list[Contenido]:
    return [Contenido(int(c), _unidad(u), int(p) if p else None, sku.upper())
            for c, u, p, sku in RE_CONTENIDO.findall(texto)]


def parece_guia(texto: str) -> bool:
    return bool(RE_REFERENCIA.search(texto) or "ORIGIN ID" in texto or "TRK#" in texto
                or RE_RASTREO.search(texto) or contenido_de(texto))


def leer_etiquetas(datos: bytes, archivo: str) -> list[EtiquetaEnvio]:
    etiquetas = []
    with pdfplumber.open(io.BytesIO(datos)) as documento:
        for numero, pagina in enumerate(documento.pages, 1):
            texto = pagina.extract_text() or ""
            referencias = list(dict.fromkeys(RE_REFERENCIA.findall(texto)))
            peso = RE_PESO_FEDEX.search(texto) or RE_PESO.search(texto)
            dims = RE_DIMS.search(texto)
            paqueteria = _paqueteria(texto)
            rastreos = RE_TRK_FEDEX.findall(texto) if paqueteria == "FedEx" else RE_RASTREO.findall(texto)
            etiquetas.append(EtiquetaEnvio(
                archivo=archivo,
                pagina=numero,
                paqueteria=paqueteria,
                pedido=referencias[0] if referencias else "",
                rastreo=rastreos[-1] if rastreos else "",
                peso=float(peso.group(1)) if peso else None,
                medidas=tuple(float(v) for v in dims.groups()) if dims else None,
                contenido=contenido_de(texto),
                legible=len(pagina.chars) > 50,
            ))
    return etiquetas
