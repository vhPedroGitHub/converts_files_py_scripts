"""Simple CLI to convert Word <-> PDF in bulk.

Usage: run and follow prompts. Requires the following packages:
- docx2pdf (for Word -> PDF). On Windows, this uses installed MS Word; on macOS it uses Preview/Word. LibreOffice is also supported when available.
- pdf2docx (for PDF -> Word).

Install:
    pip install docx2pdf pdf2docx
"""
from pathlib import Path
import sys
from typing import Iterable

try:
    from docx2pdf import convert as docx_to_pdf_convert
except Exception as exc:  # pragma: no cover - handled at runtime
    docx_to_pdf_convert = None
    DOCX2PDF_IMPORT_ERROR = exc
else:
    DOCX2PDF_IMPORT_ERROR = None

try:
    from pdf2docx import Converter as PdfToDocxConverter
except Exception as exc:  # pragma: no cover - handled at runtime
    PdfToDocxConverter = None
    PDF2DOCX_IMPORT_ERROR = exc
else:
    PDF2DOCX_IMPORT_ERROR = None


def ensure_dir(path: Path) -> Path:
    """Create the directory if it does not exist and return it."""
    path.mkdir(parents=True, exist_ok=True)
    return path


def iter_files_with_suffix(directory: Path, suffixes: Iterable[str]):
    """Yield files in directory that end with allowed suffixes (case-insensitive)."""
    lowered = tuple(s.lower() for s in suffixes)
    for item in directory.iterdir():
        if item.is_file() and item.suffix.lower() in lowered:
            yield item


def convert_docx_to_pdf(source_dir: Path) -> None:
    if docx_to_pdf_convert is None:
        raise RuntimeError(
            "docx2pdf is not available. Install with 'pip install docx2pdf' and ensure Word or LibreOffice is installed."
        ) from DOCX2PDF_IMPORT_ERROR

    output_dir = ensure_dir(source_dir / "convert_pdfs")
    files = list(iter_files_with_suffix(source_dir, (".docx",)))
    if not files:
        print("No .docx files found to convert.")
        return

    try:
        # docx2pdf can batch-convert a directory; this avoids only the first file being processed.
        docx_to_pdf_convert(str(source_dir), str(output_dir))
        print(f"Converted {len(files)} files to {output_dir}")
    except Exception as exc:  # pragma: no cover - runtime feedback
        print(f"Failed to convert directory {source_dir}: {exc}")


def convert_pdf_to_docx(source_dir: Path) -> None:
    if PdfToDocxConverter is None:
        raise RuntimeError(
            "pdf2docx is not available. Install with 'pip install pdf2docx'."
        ) from PDF2DOCX_IMPORT_ERROR

    output_dir = ensure_dir(source_dir / "convert_words")
    files = list(iter_files_with_suffix(source_dir, (".pdf",)))
    if not files:
        print("No .pdf files found to convert.")
        return

    for file in files:
        target = output_dir / f"{file.stem}.docx"
        try:
            with PdfToDocxConverter(str(file)) as converter:
                converter.convert(str(target))
            print(f"Converted: {file.name} -> {target.name}")
        except Exception as exc:  # pragma: no cover - runtime feedback
            print(f"Failed to convert {file.name}: {exc}")


def convert_single_docx_to_pdf(file_path: Path) -> None:
    if docx_to_pdf_convert is None:
        raise RuntimeError(
            "docx2pdf is not available. Install with 'pip install docx2pdf' and ensure Word or LibreOffice is installed."
        ) from DOCX2PDF_IMPORT_ERROR

    if not file_path.exists() or file_path.suffix.lower() != ".docx":
        raise FileNotFoundError(f"Archivo .docx no válido: {file_path}")

    output_dir = ensure_dir(file_path.parent / "convert_pdfs")
    target = output_dir / f"{file_path.stem}.pdf"
    try:
        docx_to_pdf_convert(str(file_path), str(target))
        print(f"Converted: {file_path.name} -> {target.name}")
    except Exception as exc:  # pragma: no cover - runtime feedback
        print(f"Failed to convert {file_path.name}: {exc}")


def convert_single_pdf_to_docx(file_path: Path) -> None:
    if PdfToDocxConverter is None:
        raise RuntimeError(
            "pdf2docx is not available. Install with 'pip install pdf2docx'."
        ) from PDF2DOCX_IMPORT_ERROR

    if not file_path.exists() or file_path.suffix.lower() != ".pdf":
        raise FileNotFoundError(f"Archivo .pdf no válido: {file_path}")

    output_dir = ensure_dir(file_path.parent / "convert_words")
    target = output_dir / f"{file_path.stem}.docx"
    try:
        with PdfToDocxConverter(str(file_path)) as converter:
            converter.convert(str(target))
        print(f"Converted: {file_path.name} -> {target.name}")
    except Exception as exc:  # pragma: no cover - runtime feedback
        print(f"Failed to convert {file_path.name}: {exc}")


def prompt_choice() -> int:
    print("Selecciona una opción:")
    print("1) Word (.docx) -> PDF")
    print("2) PDF -> Word (.docx)")
    print("3) Un archivo Word -> PDF")
    print("4) Un archivo PDF -> Word")
    choice = input("Opción [1/2/3/4]: ").strip()
    if choice not in {"1", "2", "3", "4"}:
        raise ValueError("Opción inválida. Elige 1, 2, 3 o 4.")
    return int(choice)


def prompt_directory() -> Path:
    raw = input("Introduce el directorio con los archivos a convertir: ").strip().strip('"')
    path = Path(raw).expanduser()
    if not path.exists() or not path.is_dir():
        raise FileNotFoundError(f"El directorio no existe: {path}")
    return path


def prompt_file(expected_suffix: str) -> Path:
    raw = input(f"Introduce la ruta del archivo {expected_suffix}: ").strip().strip('"')
    path = Path(raw).expanduser()
    if not path.exists() or not path.is_file() or path.suffix.lower() != expected_suffix:
        raise FileNotFoundError(f"Archivo no válido ({expected_suffix}): {path}")
    return path


def main() -> None:
    try:
        choice = prompt_choice()
        if choice in {1, 2}:
            directory = prompt_directory()
        elif choice == 3:
            file_path = prompt_file(".docx")
        else:
            file_path = prompt_file(".pdf")
    except Exception as exc:  # pragma: no cover - runtime feedback
        print(f"Error: {exc}")
        sys.exit(1)

    try:
        if choice == 1:
            convert_docx_to_pdf(directory)
        elif choice == 2:
            convert_pdf_to_docx(directory)
        elif choice == 3:
            convert_single_docx_to_pdf(file_path)
        else:
            convert_single_pdf_to_docx(file_path)
    except Exception as exc:  # pragma: no cover - runtime feedback
        print(f"Error al convertir: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
