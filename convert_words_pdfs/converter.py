"""Simple CLI to convert Word <-> PDF in bulk.

Usage: run and follow prompts. Requires the following packages:
- docx2pdf (for Word -> PDF). On Windows, this uses installed MS Word; on macOS it uses Preview/Word. LibreOffice is also supported when available.
- pdf2docx (for PDF -> Word).

Install:
    pip install docx2pdf pdf2docx
"""
from pathlib import Path
import subprocess
import sys
from typing import Iterable
import shutil

try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover - optional
    win32com = None

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


def convert_doc_to_docx(src: Path, dest_dir: Path) -> Path:
    """Convert legacy .doc to .docx. Prefer LibreOffice (avoids Trust Center), fallback to Word COM."""
    if not src.exists():
        raise FileNotFoundError(src)

    dest = dest_dir / f"{src.stem}.docx"
    last_exc: Exception | None = None

    # Prefer LibreOffice/soffice if available (avoids Trust Center blocks).
    soffice = shutil.which("soffice")
    if soffice:
        try:
            subprocess.check_call([
                soffice,
                "--headless",
                "--convert-to",
                "docx",
                "--outdir",
                str(dest_dir),
                str(src),
            ])
            if dest.exists():
                return dest
            last_exc = RuntimeError("LibreOffice no generó el archivo esperado")
        except Exception as exc:
            last_exc = exc

    # Fallback: Word COM if available.
    if win32com is not None:
        try:
            word = win32com.client.Dispatch("Word.Application")  # type: ignore
            try:
                word.Visible = False
            except Exception:
                pass  # Some environments block setting Visible
            doc = word.Documents.Open(str(src))
            doc.SaveAs(str(dest), FileFormat=16)  # 16 = wdFormatXMLDocument (docx)
            doc.Close(False)
            word.Quit()
            return dest
        except Exception as exc:
            try:
                word.Quit()
            except Exception:
                pass
            last_exc = exc

    raise RuntimeError(
        "No se pudo convertir .doc a .docx. Instala LibreOffice (soffice en PATH) o ajusta Centro de confianza de Word para permitir .doc."
    ) from last_exc


def convert_docx_to_pdf(source_dir: Path) -> None:
    if docx_to_pdf_convert is None:
        raise RuntimeError(
            "docx2pdf is not available. Install with 'pip install docx2pdf' and ensure Word or LibreOffice is installed."
        ) from DOCX2PDF_IMPORT_ERROR

    output_dir = ensure_dir(source_dir / "convert_pdfs")
    files = list(iter_files_with_suffix(source_dir, (".docx", ".doc")))
    if not files:
        print("No .doc/.docx files found to convert.")
        return

    converted = 0
    soffice = shutil.which("soffice")

    for file in files:
        target = output_dir / f"{file.stem}.pdf"
        try:
            source_path = file
            if file.suffix.lower() == ".doc":
                source_path = convert_doc_to_docx(file, output_dir)

            # Prefer LibreOffice if available (more tolerant to mismatched/legacy formats).
            if soffice:
                try:
                    subprocess.check_call([
                        soffice,
                        "--headless",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        str(output_dir),
                        str(source_path),
                    ])
                    if target.exists():
                        converted += 1
                        print(f"Converted via LibreOffice: {file.name} -> {target.name}")
                        continue
                except Exception as soffice_exc:
                    print(f"LibreOffice falló con {file.name}: {soffice_exc}. Intentando Word/docx2pdf...")

            # Fallback to docx2pdf/Word.
            docx_to_pdf_convert(str(source_path.resolve()), str(target.resolve()))
            converted += 1
            print(f"Converted: {file.name} -> {target.name}")
        except Exception as exc:  # pragma: no cover - runtime feedback
            print(f"Failed to convert {file.name}: {exc}")

    print(f"Finalizado: {converted}/{len(files)} archivos convertidos a PDF en {output_dir}")


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

    converted = 0
    for file in files:
        target = output_dir / f"{file.stem}.docx"
        try:
            with PdfToDocxConverter(str(file.resolve())) as converter:
                converter.convert(str(target.resolve()))
            converted += 1
            print(f"Converted: {file.name} -> {target.name}")
        except Exception as exc:  # pragma: no cover - runtime feedback
            print(f"Failed to convert {file.name}: {exc}")

    print(f"Finalizado: {converted}/{len(files)} archivos convertidos a Word en {output_dir}")


def convert_single_docx_to_pdf(file_path: Path) -> None:
    if docx_to_pdf_convert is None:
        raise RuntimeError(
            "docx2pdf is not available. Install with 'pip install docx2pdf' and ensure Word or LibreOffice is installed."
        ) from DOCX2PDF_IMPORT_ERROR

    if not file_path.exists() or file_path.suffix.lower() not in {".docx", ".doc"}:
        raise FileNotFoundError(f"Archivo .doc/.docx no válido: {file_path}")

    output_dir = ensure_dir(file_path.parent / "convert_pdfs")
    target = output_dir / f"{file_path.stem}.pdf"
    try:
        source_path = file_path
        if file_path.suffix.lower() == ".doc":
            source_path = convert_doc_to_docx(file_path, output_dir)

        soffice = shutil.which("soffice")
        if soffice:
            try:
                subprocess.check_call([
                    soffice,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(output_dir),
                    str(source_path),
                ])
                if target.exists():
                    print(f"Converted via LibreOffice: {file_path.name} -> {target.name}")
                    return
            except Exception as soffice_exc:
                print(f"LibreOffice falló con {file_path.name}: {soffice_exc}. Intentando Word/docx2pdf...")

        docx_to_pdf_convert(str(source_path.resolve()), str(target.resolve()))
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
    print("1) Word (.doc/.docx) -> PDF")
    print("2) PDF -> Word (.docx)")
    print("3) Un archivo Word (.doc/.docx) -> PDF")
    print("4) Un archivo PDF -> Word (.docx)")
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


def prompt_file(expected_suffix: str, alternate_suffix: str | None = None) -> Path:
    suffix_text = expected_suffix if alternate_suffix is None else f"{expected_suffix} o {alternate_suffix}"
    raw = input(f"Introduce la ruta del archivo {suffix_text}: ").strip().strip('"')
    path = Path(raw).expanduser()
    valid = {expected_suffix}
    if alternate_suffix:
        valid.add(alternate_suffix)
    if not path.exists() or not path.is_file() or path.suffix.lower() not in valid:
        raise FileNotFoundError(f"Archivo no válido ({suffix_text}): {path}")
    return path


def main() -> None:
    try:
        choice = prompt_choice()
        if choice in {1, 2}:
            directory = prompt_directory()
        elif choice == 3:
            file_path = prompt_file(".docx", ".doc")
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
