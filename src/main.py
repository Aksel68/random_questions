"""Script principal del randomizador de pruebas.

Uso:
    python -m src.main <ruta_en_drive> <n_formas>

Ejemplo:
    python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 4
"""

import os
import shutil
import sys
import tempfile

from src.answer_key import (
    build_answer_key_xlsx,
    generate_answer_key,
    parse_answer_key_csv,
)
from src.doc_builder import build_docx
from src.drive_client import (
    download_docx,
    download_sheet_as_csv,
    upload_docx,
    upload_xlsx,
)
from src.parser import parse_docx
from src.randomizer import generate_forms


def extract_form_number(filename: str) -> tuple[str, int]:
    """Extrae el número de forma y el prefijo base del nombre de archivo.

    Busca el patrón 'F<número>' al final del nombre base (antes de .docx).

    Args:
        filename: Nombre del archivo (e.g. "Ondas 2026. Ev. 1 F1.docx").

    Returns:
        Tupla (prefixo_base, número_de_forma).
        Ejemplo: ("Ondas 2026. Ev. 1 F", 1)

    Raises:
        ValueError: Si no se encuentra el patrón F<número>.
    """
    import re

    base_name = filename.rsplit(".docx", 1)[0]
    match = re.search(r"^(.*\bF)(\d+)$", base_name)
    if not match:
        raise ValueError(
            f"No se encontró el patrón 'F<número>' en el nombre del archivo: {filename}. "
            f"El archivo debe terminar en F1, F2, etc. antes de la extensión .docx"
        )
    prefix = match.group(1)  # "Ondas 2026. Ev. 1 F"
    form_number = int(match.group(2))  # 1
    return prefix, form_number


def derive_answer_key_path(drive_path: str) -> str:
    """Deriva la ruta del archivo de claves a partir de la ruta de la prueba.

    Convención: en la misma carpeta, con prefijo "Claves " y sin extensión
    (porque es un Google Sheet nativo).

    Ejemplo:
        "Guías/Ondas 2026. Ev. 1 F1.docx" → "Guías/Claves Ondas 2026. Ev. 1 F1"

    Args:
        drive_path: Ruta completa al .docx de la prueba en Drive.

    Returns:
        Ruta al Google Sheet de claves en Drive (sin extensión).
    """
    path_parts = drive_path.strip("/").split("/")
    filename = path_parts[-1]
    folder = "/".join(path_parts[:-1])

    base_name = filename.rsplit(".docx", 1)[0]
    claves_name = f"Claves {base_name}"

    if folder:
        return f"{folder}/{claves_name}"
    return claves_name


def main():
    """Punto de entrada principal del programa."""
    if len(sys.argv) < 3:
        print("Uso: python -m src.main <ruta_en_drive> <n_formas>")
        print('Ejemplo: python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 4')
        sys.exit(1)

    drive_path = sys.argv[1]
    n_forms = int(sys.argv[2])

    if n_forms < 1:
        print("❌ El número de formas debe ser al menos 1.")
        sys.exit(1)

    path_parts = drive_path.strip("/").split("/")
    filename = path_parts[-1]
    drive_folder = "/".join(path_parts[:-1]) if len(path_parts) > 1 else ""

    prefix, original_form_number = extract_form_number(filename)

    print(f"📄 Archivo: {drive_path}")
    print(f"📁 Carpeta destino: {drive_folder or 'raíz de Drive'}")
    print(f"🔢 Forma original: F{original_form_number}")
    print(f"🔄 Generando {n_forms} forma(s)...\n")

    # 1. Descargar archivo desde Google Drive
    print("⬇️  Descargando archivo de Google Drive...")
    local_path = download_docx(drive_path)

    try:
        # 2. Parsear el documento
        print("📖 Parseando documento...")
        exam = parse_docx(local_path)
        print(
            f"   Encontradas {len(exam.questions)} preguntas en tabla[{exam.questions_table_index}] "
            f"({exam.questions_per_row} columnas).\n"
        )

        if len(exam.questions) == 0:
            print("❌ No se detectaron preguntas en el documento.")
            print(
                "   Verificá que las preguntas estén en una tabla y usen formato 'N-' o 'N.'"
            )
            sys.exit(1)

        # 2.5. Descargar pauta original
        answer_key_path = derive_answer_key_path(drive_path)
        print(f"📋 Descargando pauta original: {answer_key_path}")
        try:
            csv_content = download_sheet_as_csv(answer_key_path)
            original_key = parse_answer_key_csv(csv_content)
            print(f"   Encontradas {len(original_key)} respuestas.\n")
        except FileNotFoundError:
            print(
                f"⚠️  No se encontró pauta en '{answer_key_path}'. "
                f"Se generarán las formas sin pautas.\n"
            )
            original_key = None

        # 3. Generar N formas aleatorias
        print("🎲 Generando formas aleatorias...")
        forms = generate_forms(exam, n_forms)

        # 4. Construir y subir cada forma
        temp_dir = tempfile.mkdtemp()
        for i, form_questions in enumerate(forms):
            new_form_number = original_form_number + 1 + i
            form_filename = f"{prefix}{new_form_number}.docx"
            form_path = os.path.join(temp_dir, form_filename)

            print(f"\n📝 Generando: {form_filename}")
            build_docx(
                source_path=local_path,
                questions=form_questions,
                output_path=form_path,
                form_number=new_form_number,
                questions_table_index=exam.questions_table_index,
                questions_per_row=exam.questions_per_row,
            )

            print(f"⬆️  Subiendo a Drive...")
            upload_docx(form_path, drive_folder, form_filename)

            os.unlink(form_path)

            # Generar y subir pauta de esta forma
            if original_key is not None:
                new_key = generate_answer_key(form_questions, original_key)

                key_filename = f"Claves {prefix}{new_form_number}.xlsx"
                key_path = os.path.join(temp_dir, key_filename)
                build_answer_key_xlsx(new_key, key_path)

                print(f"📋 Subiendo pauta: {key_filename}")
                upload_xlsx(key_path, drive_folder, key_filename)
                os.unlink(key_path)

        shutil.rmtree(temp_dir, ignore_errors=True)

    finally:
        if os.path.exists(local_path):
            os.unlink(local_path)

    print(f"\n✅ ¡Listo! Se generaron {n_forms} forma(s) exitosamente.")
    print("   Revisá tu carpeta en Google Drive.")


if __name__ == "__main__":
    main()
