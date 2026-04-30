"""Script principal del randomizador de pruebas.

Uso:
    python -m src.main <ruta_o_path> <n_formas>

La ruta puede ser:
  - Una ruta local a un archivo .docx:       /ruta/al/archivo.docx
  - Una ruta en Google Drive (.docx o Google Doc):
        "Carpeta/Subcarpeta/archivo.docx"
        "Carpeta/Subcarpeta/Mi Google Doc"

La pauta de respuestas se busca automáticamente:
  - Local: mismo directorio, archivo "Claves <nombre>.xlsx"
  - Drive: misma carpeta, Google Sheet o .xlsx llamado "Claves <nombre>"

Ejemplos:
    python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 4
    python -m src.main /tmp/Ondas_F1.docx 4
"""

import os
import re
import shutil
import sys
import tempfile

# ---------------------------------------------------------------------------
# Helpers de nombre de archivo
# ---------------------------------------------------------------------------


def extract_form_number(filename: str) -> tuple[str, int]:
    """Extrae el número de forma y el prefijo base del nombre de archivo.

    Busca el patrón 'F<número>' al final del nombre base (antes de la extensión).

    Args:
        filename: Nombre del archivo (e.g. "Ondas 2026. Ev. 1 F1.docx").

    Returns:
        Tupla (prefijo_base, número_de_forma).
        Ejemplo: ("Ondas 2026. Ev. 1 F", 1)

    Raises:
        ValueError: Si no se encuentra el patrón F<número>.
    """
    base_name = filename
    for ext in (".docx", ".xlsx"):
        if base_name.lower().endswith(ext):
            base_name = base_name[: -len(ext)]
            break

    match = re.search(r"^(.*\bF)(\d+)$", base_name)
    if not match:
        raise ValueError(
            f"No se encontró el patrón 'F<número>' en el nombre del archivo: {filename}. "
            f"El archivo debe terminar en F1, F2, etc. antes de la extensión."
        )
    prefix = match.group(1)  # "Ondas 2026. Ev. 1 F"
    form_number = int(match.group(2))  # 1
    return prefix, form_number


def _strip_extension(filename: str) -> str:
    """Elimina la extensión de un nombre de archivo si es .docx o .xlsx."""
    for ext in (".docx", ".xlsx"):
        if filename.lower().endswith(ext):
            return filename[: -len(ext)]
    return filename


# ---------------------------------------------------------------------------
# Resolución de entrada (local vs Drive)
# ---------------------------------------------------------------------------


def _is_local_path(path: str) -> bool:
    """Determina si el argumento apunta a un archivo local existente."""
    return os.path.isfile(path)


def _resolve_local_answer_key(docx_path: str) -> str | None:
    """Busca la pauta local como 'Claves <nombre_base>.xlsx' en la misma carpeta.

    Args:
        docx_path: Ruta al .docx local.

    Returns:
        Ruta al .xlsx de claves si existe, None en caso contrario.
    """
    folder = os.path.dirname(os.path.abspath(docx_path))
    base_name = _strip_extension(os.path.basename(docx_path))
    candidate = os.path.join(folder, f"Claves {base_name}.xlsx")
    if os.path.isfile(candidate):
        return candidate
    return None


def _derive_drive_answer_key_path(drive_path: str) -> str:
    """Deriva la ruta Drive de la pauta a partir de la ruta del .docx.

    Convención: misma carpeta, prefijo "Claves ", sin extensión.
    Ejemplo: "Guías/Ondas F1.docx" → "Guías/Claves Ondas F1"

    Args:
        drive_path: Ruta completa al archivo de prueba en Drive.

    Returns:
        Ruta al archivo de claves en Drive (sin extensión).
    """
    parts = drive_path.strip("/").split("/")
    filename = parts[-1]
    folder = "/".join(parts[:-1])
    base_name = _strip_extension(filename)
    claves_name = f"Claves {base_name}"
    return f"{folder}/{claves_name}" if folder else claves_name


# ---------------------------------------------------------------------------
# Funciones de pauta local (.xlsx CSV)
# ---------------------------------------------------------------------------


def _parse_local_xlsx_answer_key(xlsx_path: str) -> dict[int, str]:
    """Lee una pauta de respuestas desde un archivo .xlsx local.

    Formato esperado:
        Columna A: "Pregunta" (encabezado) + números
        Columna B: "Respuesta" (encabezado) + letras

    Args:
        xlsx_path: Ruta al archivo .xlsx.

    Returns:
        Dict {número_pregunta: letra_correcta}.
    """
    try:
        import openpyxl
    except ImportError:
        raise ImportError(
            "openpyxl es necesario para leer pautas .xlsx locales. Ejecutá: uv add openpyxl"
        )

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    if ws is None:
        raise ValueError("El archivo Excel no contiene hojas activas.")
    answer_key: dict[int, str] = {}

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return answer_key

    # Detectar si la primera fila es encabezado
    start_row = 1 if str(rows[0][0]).lower() in ("pregunta", "question") else 0

    for row in rows[start_row:]:
        if len(row) < 2:
            continue
        q_raw = row[0]
        a_raw = row[1]
        if q_raw is None or a_raw is None:
            continue
        try:
            answer_key[int(str(q_raw))] = str(a_raw).strip().lower()
        except (ValueError, TypeError):
            continue

    return answer_key


# ---------------------------------------------------------------------------
# Flujo principal
# ---------------------------------------------------------------------------


def main():
    """Punto de entrada principal del programa."""
    if len(sys.argv) < 3:
        print("Uso: python -m src.main <ruta_o_path_drive> <n_formas>")
        print('Ejemplo (Drive): python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 4')
        print("Ejemplo (local): python -m src.main /tmp/Ondas_F1.docx 4")
        sys.exit(1)

    input_path = sys.argv[1]
    n_forms = int(sys.argv[2])

    if n_forms < 1:
        print("❌ El número de formas debe ser al menos 1.")
        sys.exit(1)

    # Importaciones diferidas para no romper si Drive no está configurado
    from src.answer_key import (
        build_answer_key_xlsx,
        generate_answer_key,
        parse_answer_key_csv,
    )
    from src.doc_builder import build_docx
    from src.parser import parse_docx
    from src.randomizer import generate_forms

    is_local = _is_local_path(input_path)

    if is_local:
        # ------------------------------------------------------------------ #
        # Modo local                                                           #
        # ------------------------------------------------------------------ #
        local_docx_path = input_path
        filename = os.path.basename(local_docx_path)
        drive_folder = ""

        print(f"📄 Archivo (local): {local_docx_path}")

        prefix, original_form_number = extract_form_number(filename)
        print(f"🔢 Forma original: F{original_form_number}")
        print(f"🔄 Generando {n_forms} forma(s)...\n")

        # Buscar pauta local
        local_key_path = _resolve_local_answer_key(local_docx_path)
        original_key: dict[int, str] | None = None
        if local_key_path:
            print(f"📋 Pauta encontrada localmente: {local_key_path}")
            try:
                original_key = _parse_local_xlsx_answer_key(local_key_path)
                print(f"   Encontradas {len(original_key)} respuestas.\n")
            except Exception as e:
                print(f"⚠️  No se pudo leer la pauta local: {e}\n")
        else:
            print(
                f"⚠️  No se encontró pauta local 'Claves {_strip_extension(filename)}.xlsx'. "
                f"Se generarán las formas sin pautas.\n"
            )

        # Parsear
        print("📖 Parseando documento...")
        exam = parse_docx(local_docx_path)
        print(
            f"   Encontradas {len(exam.questions)} preguntas en tabla[{exam.questions_table_index}] "
            f"({exam.questions_per_row} columnas).\n"
        )

        if not exam.questions:
            print("❌ No se detectaron preguntas en el documento.")
            sys.exit(1)

        # Generar formas
        print("🎲 Generando formas aleatorias...")
        forms = generate_forms(exam, n_forms)

        output_dir = os.path.dirname(os.path.abspath(local_docx_path))
        temp_dir = tempfile.mkdtemp()

        try:
            for i, form_questions in enumerate(forms):
                new_form_number = original_form_number + 1 + i
                form_filename = f"{prefix}{new_form_number}.docx"
                form_path = os.path.join(output_dir, form_filename)

                print(f"\n📝 Generando: {form_filename}")
                build_docx(
                    source_path=local_docx_path,
                    questions=form_questions,
                    output_path=form_path,
                    form_number=new_form_number,
                    questions_table_index=exam.questions_table_index,
                    questions_per_row=exam.questions_per_row,
                )
                print(f"   Guardado en: {form_path}")

                if original_key is not None:
                    new_key = generate_answer_key(form_questions, original_key)
                    key_filename = f"Claves {prefix}{new_form_number}.xlsx"
                    key_path = os.path.join(output_dir, key_filename)
                    build_answer_key_xlsx(new_key, key_path)
                    print(f"📋 Pauta guardada: {key_path}")
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    else:
        # ------------------------------------------------------------------ #
        # Modo Google Drive                                                    #
        # ------------------------------------------------------------------ #
        from src.drive_client import (
            download_answer_key_as_csv,
            download_file_as_docx,
            upload_docx,
            upload_xlsx,
        )

        drive_path = input_path
        parts = drive_path.strip("/").split("/")
        filename = parts[-1]
        drive_folder = "/".join(parts[:-1]) if len(parts) > 1 else ""

        prefix, original_form_number = extract_form_number(filename)

        print(f"📄 Archivo (Drive): {drive_path}")
        print(f"📁 Carpeta destino: {drive_folder or 'raíz de Drive'}")
        print(f"🔢 Forma original: F{original_form_number}")
        print(f"🔄 Generando {n_forms} forma(s)...\n")

        # 1. Descargar archivo (admite .docx nativo y Google Doc)
        print("⬇️  Descargando archivo de Google Drive...")
        local_path = download_file_as_docx(drive_path)

        try:
            # 2. Parsear
            print("📖 Parseando documento...")
            exam = parse_docx(local_path)
            print(
                f"   Encontradas {len(exam.questions)} preguntas en tabla[{exam.questions_table_index}] "
                f"({exam.questions_per_row} columnas).\n"
            )

            if not exam.questions:
                print("❌ No se detectaron preguntas en el documento.")
                sys.exit(1)

            # 3. Descargar pauta (admite Google Sheet y .xlsx)
            answer_key_path = _derive_drive_answer_key_path(drive_path)
            print(f"📋 Buscando pauta: {answer_key_path}")
            original_key = None
            try:
                csv_content = download_answer_key_as_csv(answer_key_path)
                original_key = parse_answer_key_csv(csv_content)
                print(f"   Encontradas {len(original_key)} respuestas.\n")
            except FileNotFoundError:
                print(
                    f"⚠️  No se encontró pauta en '{answer_key_path}'. "
                    f"Se generarán las formas sin pautas.\n"
                )
            except ValueError as e:
                print(f"⚠️  Tipo de archivo de pauta no soportado: {e}\n")

            # 4. Generar formas aleatorias
            print("🎲 Generando formas aleatorias...")
            forms = generate_forms(exam, n_forms)

            # 5. Construir y subir cada forma
            temp_dir = tempfile.mkdtemp()
            try:
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

                    if original_key is not None:
                        new_key = generate_answer_key(form_questions, original_key)
                        key_filename = f"Claves {prefix}{new_form_number}.xlsx"
                        key_path = os.path.join(temp_dir, key_filename)
                        build_answer_key_xlsx(new_key, key_path)
                        print(f"📋 Subiendo pauta: {key_filename}")
                        upload_xlsx(key_path, drive_folder, key_filename)
                        os.unlink(key_path)
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        finally:
            if os.path.exists(local_path):
                os.unlink(local_path)

    print(f"\n✅ ¡Listo! Se generaron {n_forms} forma(s) exitosamente.")


if __name__ == "__main__":
    main()
