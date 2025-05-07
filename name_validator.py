#!/usr/bin/env python3
"""
name_validator.py
=================

Valida la coincidencia de personas entre una “Cédula de Determinación de Cuotas” del
SUA‑IMSS (PDF raster) y un listado de control en Excel. Genera un archivo Excel idéntico
al original con una columna extra llamada ``Existe en SUA`` que indica *True* cuando el
trabajador aparece en ambos documentos y *False* en caso contrario.

Novedades de esta versión
-------------------------
* **Variables globales** para directorios (*SUA*, *EXCEL*, *OUTPUT*). Solo introduces los
  nombres de archivo ―¡sin rutas completas!―.
* **Nombre de salida automático**: el Excel resultante mantiene el mismo nombre que el
  original, añadiendo el sufijo ``_REVISADO.xlsx`` en la carpeta *OUTPUT*.
* CLI ultra‑simple: únicamente dos argumentos posicionales (``pdf`` y ``excel``) y
  opcionales ``--debug`` y ``--tesseract-cmd``.
* Se elimina la opción de columna porque siempre es **"Nombre Completo"**.
* Sigue contando con manejo de excepciones, creación de carpetas y logging detallado.

Ejemplo mínimo
~~~~~~~~~~~~~~
```bash
python name_validator.py \
  "10 SUA Z29 IMSS.pdf" \
  "20250116  LISTA DIGITAL SEGURIDAD  CHIVAS VS JUAREZ   FEMENIL.xlsx" \
  --debug        # opcional
```

Dependencias
~~~~~~~~~~~~
```bash
pip install -r requirements.txt
sudo apt-get install tesseract-ocr poppler-utils   # Debian/Ubuntu
```
"""
from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Set

import pandas as pd
import pytesseract
from pdf2image import convert_from_path, exceptions as pdf2image_exceptions
from PIL.Image import Image  # type‑hints
from unidecode import unidecode

__all__ = [
    "normalize_name",
    "PdfNameExtractor",
    "ExcelNameValidator",
    "parse_args",
    "main",
]

# ---------------------------------------------------------------------------
# Variables globales de proyecto
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
PDF_DIR = BASE_DIR / "SUA"
EXCEL_DIR = BASE_DIR / "EXCEL"
OUTPUT_DIR = BASE_DIR / "OUTPUT"
NAME_COLUMN = "Nombre Completo"

# ---------------------------------------------------------------------------
# Configuración de logging
# ---------------------------------------------------------------------------
logger = logging.getLogger("name_validator")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
)

# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def normalize_name(name: str) -> str:
    """Normaliza un nombre para comparación."""
    ascii_txt = unidecode(name)
    clean_txt = re.sub(r"[^A-Za-z\s]", " ", ascii_txt)
    collapsed = re.sub(r"\s+", " ", clean_txt).strip()
    return collapsed.upper()

# ---------------------------------------------------------------------------
# OCR extractor
# ---------------------------------------------------------------------------

class PdfExtractionError(RuntimeError):
    """Error genérico al convertir/leer PDF."""


@dataclass
class PdfNameExtractor:
    pdf_path: Path
    tesseract_cmd: str | None = None
    dpi: int = 300

    def extract_names(self) -> Set[str]:
        if self.tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_cmd

        if not self.pdf_path.is_file():
            raise PdfExtractionError(f"PDF no encontrado: {self.pdf_path}")

        try:
            logger.info("Convirtiendo '%s' a imágenes (dpi=%s)…", self.pdf_path, self.dpi)
            images: List[Image] = convert_from_path(str(self.pdf_path), dpi=self.dpi)
        except pdf2image_exceptions.PDFInfoNotInstalledError as exc:
            raise PdfExtractionError("poppler utils no instalados o 'pdfinfo' no está en PATH") from exc
        except Exception as exc:
            raise PdfExtractionError("Fallo al convertir PDF a imágenes") from exc

        names: Set[str] = set()
        for i, img in enumerate(images, start=1):
            try:
                text = pytesseract.image_to_string(img, lang="spa")
            except pytesseract.TesseractError as exc:
                logger.warning("OCR falló en página %s: %s", i, exc)
                continue

            page_names = {
                normalize_name(line)
                for line in text.splitlines()
                if line.strip() and any(c.isalpha() for c in line)
            }
            names.update({n for n in page_names if len(n.split()) >= 2})
            logger.debug("Página %s: %d nombres detectados", i, len(page_names))

        logger.info("Total de nombres extraídos del PDF: %d", len(names))
        return names

# ---------------------------------------------------------------------------
# Excel validator
# ---------------------------------------------------------------------------

@dataclass
class ExcelNameValidator:
    excel_path: Path
    output_path: Path

    def validate(self, pdf_names: Set[str]) -> None:
        if not self.excel_path.is_file():
            raise FileNotFoundError(f"Excel no encontrado: {self.excel_path}")

        try:
            df = pd.read_excel(self.excel_path)
        except Exception as exc:
            logger.error("Error al leer Excel: %s", exc)
            raise

        if NAME_COLUMN not in df.columns:
            raise KeyError(
                f"La columna '{NAME_COLUMN}' no existe en el Excel. Columnas disponibles: {list(df.columns)}"
            )

        df["_NombreNormalizado"] = df[NAME_COLUMN].astype(str).apply(normalize_name)
        df["Existe en SUA"] = df["_NombreNormalizado"].apply(lambda n: n in pdf_names)
        df.drop(columns="_NombreNormalizado", inplace=True)

        self.output_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            df.to_excel(self.output_path, index=False)
        except Exception as exc:
            logger.error("No se pudo escribir la salida '%s': %s", self.output_path, exc)
            raise

        logger.info("Archivo con validación guardado en '%s'", self.output_path)

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: List[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Valida coincidencia de nombres entre SUA (PDF) y Excel. Solo indica los nombres de archivo.",
        epilog="Ejemplo: python name_validator.py \"10 SUA Z29 IMSS.pdf\" \"20250116  LISTA DIGITAL SEGURIDAD  CHIVAS VS JUAREZ   FEMENIL.xlsx\"",
    )
    parser.add_argument("pdf", help="Nombre del archivo PDF dentro de la carpeta SUA")
    parser.add_argument("excel", help="Nombre del archivo Excel dentro de la carpeta EXCEL")
    parser.add_argument("--debug", action="store_true", help="Activa logging DEBUG")
    parser.add_argument("--tesseract-cmd", default=None, help="Ruta al binario tesseract si no está en PATH")
    parser.add_argument("--dpi", type=int, default=300, help="DPI para OCR (default: 300)")
    parser.add_argument("--pages", type=str, help="Rango de páginas a procesar (e.g., '1-3')")
    return parser.parse_args(argv)

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main(argv: List[str] | None = None) -> None:
    args = parse_args(argv)

    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)

    # Validación de extensiones de archivo
    if not args.pdf.lower().endswith(".pdf"):
        logger.error("El archivo PDF debe tener la extensión '.pdf'")
        sys.exit(1)

    if not args.excel.lower().endswith(".xlsx"):
        logger.error("El archivo Excel debe tener la extensión '.xlsx'")
        sys.exit(1)

    pdf_path = PDF_DIR / args.pdf
    excel_path = EXCEL_DIR / args.excel
    output_path = OUTPUT_DIR / f"{Path(args.excel).stem}_REVISADO.xlsx"

    try:
        pdf_names = PdfNameExtractor(
            pdf_path=pdf_path,
            tesseract_cmd=args.tesseract_cmd,
            dpi=args.dpi,
        ).extract_names()

        ExcelNameValidator(
            excel_path=excel_path,
            output_path=output_path,
        ).validate(pdf_names)

    except (PdfExtractionError, FileNotFoundError, KeyError) as exc:
        logger.error("%s", exc)
        sys.exit(1)
    except Exception as exc:
        logger.exception("Error inesperado: %s", exc)
        sys.exit(2)

    logger.info("Proceso finalizado correctamente ✅")


if __name__ == "__main__":  # pragma: no cover
    main()
