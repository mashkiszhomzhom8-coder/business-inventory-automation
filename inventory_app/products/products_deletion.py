from pathlib import Path

import xlwings as xw

from inventory_app.core.exceptions import ProductAlreadyDeletedError, ProductDeletionInputError, ProductIdNotFoundError, ProductDeletionSystemError
from inventory_app.core.database import connect_db
from inventory_app.excel.excel_reader import get_worksheet, price_sheet
from inventory_app.core.logging_config import setup_logging
from inventory_app.core.sync_utils import create_backup, restore_excel_file

logger = setup_logging()




def validate_product_ids(conn, raw_ids):
    """Проверяет строку ID товаров и их доступность для деактивации.

    Returns:
        Список ID товаров или ``None`` при отмене ввода.
    """
    if raw_ids is None:
        return None

    raw_ids = str(raw_ids).strip()

    if not raw_ids:
        raise ProductDeletionInputError("Не указано ни одного ID товара")

    parts = [part.strip() for part in raw_ids.split(",")]

    if any(not part for part in parts):
        raise ProductDeletionInputError("ID товаров должны быть указаны через запятую")

    if any(not part.isdigit() for part in parts):
        raise ProductDeletionInputError("ID товара должен быть целым положительным числом")

    product_ids = [int(part) for part in parts]

    if any(product_id <= 0 for product_id in product_ids):
        raise ProductDeletionInputError("ID товара должен быть целым положительным числом")

    if len(product_ids) != len(set(product_ids)):
        raise ProductDeletionInputError("Один и тот же ID указан несколько раз")

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

    deleted_ids = [product_id for product_id in product_ids if db_products[product_id] == 0]

    if deleted_ids:
        raise ProductAlreadyDeletedError(f"Товары с ID {', '.join(map(str, deleted_ids))} уже удалены")

    return product_ids


def deactivate_products(conn, product_ids):
    """Помечает выбранные товары как неактивные."""
    placeholders = ",".join("?" for _ in product_ids)

    cursor = conn.execute(
        f"""
        UPDATE products
        SET active = 0
        WHERE id IN ({placeholders})
          AND active = 1;
        """,
        product_ids
    )

    if cursor.rowcount != len(product_ids):
        raise ProductDeletionSystemError("Не удалось деактивировать все выбранные товары")

def find_product_rows_in_excel(worksheet, product_ids):
    """Находит строки прайса для указанных ID товаров.

    Returns:
        Словарь номеров строк по ID товара.
    """
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

        if product_id in product_ids:
            rows_by_id[product_id] = row_num

    missing_ids = [product_id for product_id in product_ids if product_id not in rows_by_id]

    if missing_ids:
        raise ProductDeletionSystemError(f"В Excel не найдены товары с ID: {', '.join(map(str, missing_ids))}")

    return rows_by_id

def delete_product_rows_from_excel(worksheet, product_ids):
    """Помечает выбранные товары в Excel серой заливкой."""
    rows_by_id = find_product_rows_in_excel(worksheet, product_ids)
    for row_num in rows_by_id.values():
        worksheet.range(f"{row_num}:{row_num}").color = (166, 166, 166)




def show_deletion_success(app, product_ids):
    """Показывает пользователю подтверждение удаления товаров."""
    ids_text = ", ".join(map(str, product_ids))
    app.alert(f"Товары успешно деактивированы.\n\nНеактивные ID: {ids_text}", title="Деактивация товаров", buttons="ok", mode="info")


def show_deletion_error(app, error, excel_restored):
    logger.info("Вход в show_deletion_error")
    if excel_restored:
        message = f"Деактивация товаров не выполнено.\n\nИзменения в базе данных отменены, Excel-файл восстановлен.\n\nОшибка: {error}"
    else:
        message = f"Деактивация товаров не выполнено.\n\nИзменения в базе данных отменены, но автоматически восстановить Excel-файл не удалось.\n\nОшибка: {error}"
    # logger.info("Перед app.alert")
    app.alert(message, title="Ошибка деактивации товаров", buttons="ok", mode="critical")
    # logger.info("После app.alert")



def delete_products(workbook, raw_ids):
    """Удаляет товары из БД и прайс-листа с резервным копированием."""
    excel_path = Path(workbook.fullname)
    backup_path = excel_path.with_name(f"{excel_path.stem}_backup{excel_path.suffix}")
    app = workbook.app
    conn = connect_db()
    backup_created = False

    logger.info("Запуск деактивации товаров: excel_file=%s", excel_path)

    try:
        product_ids = validate_product_ids(conn, raw_ids)
        logger.info("Получены ID товаров для деактивации: %s", product_ids)

        workbook.save()
        create_backup(excel_path, backup_path)
        backup_created = True

        worksheet = get_worksheet(workbook, price_sheet)

        deactivate_products(conn, product_ids)
        delete_product_rows_from_excel(worksheet, product_ids)

        workbook.save()
        conn.commit()
        logger.info("Все товары успешно деактивированы")

    except (ProductDeletionInputError, ProductIdNotFoundError, ProductAlreadyDeletedError) as error:
        conn.rollback()
        logger.warning("Ошибка пользовательского ввода при деактивации товаров: %s", error)
        app.alert(str(error), title="Ошибка деактивации товаров", buttons="ok", mode="critical")
        return

    except Exception as error:
        conn.rollback()
        logger.exception("Системная ошибка деактивации товаров")

        excel_restored = False

        show_deletion_error(app, error, excel_restored)
        return

    else:
        if backup_path.exists():
            backup_path.unlink()

        show_deletion_success(app, product_ids)

    finally:
        conn.close()

def main(raw_ids):
    """Получает книгу от xlwings и запускает деактивацию товаров."""
    logger.info(type(raw_ids))
    workbook = xw.Book.caller()
    delete_products(workbook, raw_ids)
