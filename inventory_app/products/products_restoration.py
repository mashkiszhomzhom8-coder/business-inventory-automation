from pathlib import Path

import xlwings as xw

from inventory_app.core.exceptions import ProductAlreadyActiveError, ProductIdNotFoundError, ProductRestorationInputError, ProductRestorationSystemError
from inventory_app.core.database import connect_db
from inventory_app.excel.excel_reader import get_worksheet, price_sheet
from inventory_app.core.logging_config import setup_logging
from inventory_app.core.sync_utils import create_backup, restore_excel_file

logger = setup_logging()


def get_inactive_products(conn):
    """Возвращает список неактивных товаров для отображения пользователю."""
    return conn.execute(
        """
        SELECT id, name, size
        FROM products
        WHERE active = 0
        ORDER BY id;
        """
    ).fetchall()


def format_inactive_products(products):
    """Формирует текстовый список неактивных товаров."""
    return "\n".join(
        f'{product["id"]} — {product["name"]} — {product["size"] or "без размера"}'
        for product in products
    )


def validate_product_ids(conn, raw_ids):
    """Проверяет введенные ID и возможность восстановления товаров."""
    if raw_ids is None:
        return None

    raw_ids = raw_ids.strip()

    if not raw_ids:
        raise ProductRestorationInputError("Не указано ни одного ID товара")

    parts = [part.strip() for part in raw_ids.split(",")]

    if any(not part for part in parts):
        raise ProductRestorationInputError("ID товаров должны быть указаны через запятую")
    if any(not part.isdigit() for part in parts):
        raise ProductRestorationInputError("ID товара должен быть целым положительным числом")

    product_ids = [int(part) for part in parts]

    if any(product_id <= 0 for product_id in product_ids):
        raise ProductRestorationInputError("ID товара должен быть целым положительным числом")
    if len(product_ids) != len(set(product_ids)):
        raise ProductRestorationInputError("Один и тот же ID указан несколько раз")

    placeholders = ",".join("?" for _ in product_ids)

    rows = conn.execute(
        f"""
        SELECT id, active
        FROM products
        WHERE id IN ({placeholders});
        """,
        product_ids
    ).fetchall()

    db_products = {row["id"]: row["active"] for row in rows}

    missing_ids = [product_id for product_id in product_ids if product_id not in db_products]

    if missing_ids:
        raise ProductIdNotFoundError(f"Товары с ID {', '.join(map(str, missing_ids))} отсутствуют в базе данных")

    active_ids = [product_id for product_id in product_ids if db_products[product_id] == 1]

    if active_ids:
        raise ProductAlreadyActiveError(f"Товары с ID {', '.join(map(str, active_ids))} уже активны")

    return product_ids


def activate_products(conn, product_ids):
    """Помечает выбранные товары как активные."""
    placeholders = ",".join("?" for _ in product_ids)

    cursor = conn.execute(
        f"""
        UPDATE products
        SET active = 1
        WHERE id IN ({placeholders})
          AND active = 0;
        """,
        product_ids
    )

    if cursor.rowcount != len(product_ids):
        raise ProductRestorationSystemError("Не удалось восстановить все выбранные товары")


def find_product_rows_in_excel(worksheet, product_ids):
    """Находит строки товаров в прайс-листе по ID."""
    rows_by_id = {}

    last_row = worksheet.range(f"A{worksheet.cells.last_cell.row}").end("up").row

    for row_num in range(2, last_row + 1):
        raw_product_id = worksheet.range(f"A{row_num}").value

        if raw_product_id is None:
            continue

        try:
            product_id = int(raw_product_id)
        except (TypeError, ValueError):
            continue

        if product_id in product_ids: rows_by_id[product_id] = row_num

    missing_ids = [product_id for product_id in product_ids if product_id not in rows_by_id]

    if missing_ids:
        raise ProductRestorationSystemError(f"В Excel не найдены товары с ID: {', '.join(map(str, missing_ids))}")

    return rows_by_id


def restore_product_rows_in_excel(worksheet, product_ids):
    """Возвращает белую заливку строк восстановленных товаров."""
    rows_by_id = find_product_rows_in_excel(worksheet, product_ids)

    for row_num in rows_by_id.values():
        worksheet.range(f"{row_num}:{row_num}").color = None


def show_restoration_success(app, product_ids):
    """Показывает подтверждение успешного восстановления."""
    ids_text = ", ".join(map(str, product_ids))
    app.alert(f"Товары успешно восстановлены.\n\nВосстановленные ID: {ids_text}", title="Восстановление товаров", buttons="ok", mode="info")


def show_restoration_error(app, error, excel_restored):
    """Показывает системную ошибку восстановления."""
    if excel_restored:
        message = f"Восстановление товаров не выполнено.\n\nИзменения в базе данных отменены, Excel-файл восстановлен.\n\nОшибка: {error}"
    else:
        message = f"Восстановление товаров не выполнено.\n\nИзменения в базе данных отменены, но автоматически восстановить Excel-файл не удалось.\n\nОшибка: {error}"

    app.alert(message, title="Ошибка восстановления товаров", buttons="ok", mode="critical")

def prepare_inactive_products():
    """Подготавливает список неактивных товаров для VBA."""
    workbook = xw.Book.caller()
    conn = connect_db()

    try:
        service_sheet = workbook.sheets["_service"]

        products = get_inactive_products(conn)

        if products:
            products_text = format_inactive_products(products)
        else:
            products_text = "Неактивных товаров нет."

        service_sheet.range("A1").value = products_text

    finally:
        conn.close()
# def prepare_inactive_products():
#     """Сохраняет список неактивных товаров во временный файл."""
#     workbook = xw.Book.caller()
#     conn = connect_db()
#
#     try:
#         products = get_inactive_products(conn)
#
#         if products:
#             products_text = format_inactive_products(products)
#         else:
#             products_text = "Неактивных товаров нет."
#
#         excel_path = Path(workbook.fullname)
#         temp_file = excel_path.parent / "inactive_products.txt"
#
#         temp_file.write_text(
#             products_text,
#             encoding="utf-8"
#         )
#
#     finally:
#         conn.close()


def restore_products(workbook, raw_ids):
    """Восстанавливает товары в БД и Excel."""
    excel_path = Path(workbook.fullname)
    backup_path = excel_path.with_name(f"{excel_path.stem}_backup{excel_path.suffix}")
    app = workbook.app
    conn = connect_db()
    backup_created = False

    logger.info("Запуск восстановления товаров: excel_file=%s", excel_path)

    try:
        product_ids = validate_product_ids(conn, raw_ids)
        logger.info("Получены ID товаров для восстановления: %s", product_ids)

        workbook.save()
        create_backup(excel_path, backup_path)
        backup_created = True

        worksheet = get_worksheet(workbook, price_sheet)

        activate_products(conn, product_ids)
        restore_product_rows_in_excel(worksheet, product_ids)

        workbook.save()
        conn.commit()

    except (ProductRestorationInputError, ProductIdNotFoundError, ProductAlreadyActiveError) as error:
        conn.rollback()
        logger.warning("Ошибка пользовательского ввода при восстановлении товаров: %s", error)
        app.alert(str(error), title="Ошибка восстановления товаров", buttons="ok", mode="critical")
        return

    except Exception as error:
        conn.rollback()
        logger.exception("Системная ошибка восстановления товаров")

        excel_restored = False

        if backup_created and backup_path.exists():
            try:
                restore_excel_file(workbook, app, excel_path, backup_path)
                excel_restored = True
            except Exception:
                logger.exception("Не удалось автоматически восстановить Excel-файл")

        show_restoration_error(app, error, excel_restored)
        return

    else:
        if backup_path.exists(): backup_path.unlink()

        logger.info("Восстановление товаров успешно завершено: product_ids=%s", product_ids)
        show_restoration_success(app, product_ids)

    finally:
        conn.close()


def main(raw_ids):
    """Получает книгу от xlwings и запускает восстановление товаров."""
    workbook = xw.Book.caller()
    restore_products(workbook, raw_ids)