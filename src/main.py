"""Script principal del randomizador de pruebas.

Uso:
    python -m src.main <ruta_en_drive> <n_formas>

Ejemplo:
    python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 3
"""

import os
import shutil
import sys
import tempfile

from src.doc_builder import build_docx
from src.drive_client import download_docx, upload_docx
from src.parser import parse_docx
from src.randomizer import generate_forms


def get_form_suffix(index: int) -> str:
    """Genera el sufijo de la forma: A, B, C, ..., Z, AA, AB, ..."""
    if index < 26:
        return chr(ord("A") + index)
    else:
        first = chr(ord("A") + (index // 26) - 1)
        second = chr(ord("A") + (index % 26))
        return first + second


def main():
    """Punto de entrada principal del programa."""
    if len(sys.argv) < 3:
        print("Uso: python -m src.main <ruta_en_drive> <n_formas>")
        print('Ejemplo: python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 3')
        sys.exit(1)

    drive_path = sys.argv[1]
    n_forms = int(sys.argv[2])

    if n_forms < 1:
        print("❌ El número de formas debe ser al menos 1.")
        sys.exit(1)

    path_parts = drive_path.strip("/").split("/")
    filename = path_parts[-1]
    drive_folder = "/".join(path_parts[:-1]) if len(path_parts) > 1 else ""
    base_name = filename.rsplit(".docx", 1)[0]

    print(f"📄 Archivo: {drive_path}")
    print(f"📁 Carpeta destino: {drive_folder or 'raíz de Drive'}")
    print(f"🔄 Generando {n_forms} forma(s)...\n")

    # 1. Descargar archivo desde Google Drive
    print("⬇️  Descargando archivo de Google Drive...")
    local_path = download_docx(drive_path)

    try:
        # 2. Parsear el documento
        print("📖 Parseando documento...")
        exam = parse_docx(local_path)
        print(
            f"   Encontradas {len(exam.questions)} preguntas, "
            f"header con {len(exam.header_paragraphs)} línea(s).\n"
        )

        # 3. Generar N formas aleatorias
        print("🎲 Generando formas aleatorias...")
        forms = generate_forms(exam, n_forms)

        # 4. Construir y subir cada forma
        temp_dir = tempfile.mkdtemp()
        for i, form in enumerate(forms):
            suffix = get_form_suffix(i)
            form_filename = f"{base_name} Forma {suffix}.docx"
            form_path = os.path.join(temp_dir, form_filename)

            print(f"\n📝 Generando: {form_filename}")
            build_docx(form, form_path)

            print(f"⬆️  Subiendo a Drive...")
            upload_docx(form_path, drive_folder, form_filename)

            os.unlink(form_path)

        shutil.rmtree(temp_dir, ignore_errors=True)

    finally:
        if os.path.exists(local_path):
            os.unlink(local_path)

    print(f"\n✅ ¡Listo! Se generaron {n_forms} forma(s) exitosamente.")
    print("   Revisá tu carpeta en Google Drive.")


if __name__ == "__main__":
    main()
