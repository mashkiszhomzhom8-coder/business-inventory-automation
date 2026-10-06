import sys
from pathlib import Path
import xlwings as xw
from datetime import date, datetime

from inventory_app.core.database import connect_db
from inventory_app.excel.excel_reader import products_reader, excel_read_supply, get_worksheet, price_sheet
from inventory_app.manual_changes.products_service import update_product_attributes
from inventory_app.manual_changes.remain_service import update_product_remains
from inventory_app.manual_changes.supply_service import process_existing_supplies, process_new_supplies
from inventory_app.core.logging_config import setup_logging
from inventory_app.core.sync_utils import create_backup, restore_excel_file, get_active_season_id, log_sync_error
from inventory_app.core.exceptions import ExcelSynchronizationError
from inventory_app.orders.order_service_excel import get_product_rows, update_excel_current_stock_col

logger = setup_logging()


def process_manual_changes(conn, workbook):
    """Синхронизирует ручные изменения прайса с базой данных.

    Returns:
        Словарь результатов по каждому виду изменений.
    """
    season_id = get_active_season_id(conn)

    excel_products = products_reader(workbook)
    supplies = excel_read_supply(workbook)

    logger.info("Прочитано товаров из Excel: %s", len(excel_products))
    logger.info("Прочитано столбцов поступлений из Excel: %s", len(supplies))

    product_changes = update_product_attributes(conn, excel_products)
    logger.info("Обработано изменений атрибутов товаров: %s", len(product_changes))

    existing_supply_changes = process_existing_supplies(conn, season_id, supplies)
    logger.info("Обработано изменённых поступлений: %s", len(existing_supply_changes))

    new_supplies = process_new_supplies(conn, season_id, supplies)
    logger.info("Обработано новых поступлений: %s", len(new_supplies))

    remain_changes = update_product_remains(conn, excel_products, season_id)
    logger.info("Обработано изменений остатков: %s", len(remain_changes))

    return {
        "product_changes": product_changes,
        "existing_supply_changes": existing_supply_changes,
        "new_supplies": new_supplies,
        "remain_changes": remain_changes
    }


def count_existing_supply_item_changes(existing_supply_changes):
    """Подсчитывает изменённые позиции существующих поступлений."""
    return sum(len(supply["changes"]) for supply in existing_supply_changes)


def has_manual_changes(sync_result):
    """Проверяет, содержит ли результат синхронизации изменения."""
    return any([
        sync_result["product_changes"],
        sync_result["existing_supply_changes"],
        sync_result["new_supplies"],
        sync_result["remain_changes"]
    ])

def format_supply_date(supply_date):
    """Преобразует дату поступления в формат заголовка Excel."""
    if isinstance(supply_date, (date, datetime)):
        return supply_date.strftime("%d.%m.%Y")

    return str(supply_date)


def update_new_supply_headers(worksheet, processed_supplies):
    """Заменяет заголовки новых поступлений на заголовки с ID партии."""
    if not processed_supplies:
        return

    last_column = worksheet.used_range.last_cell.column
    headers = worksheet.range((1, 1), (1, last_column)).value

    if not isinstance(headers, list):
        headers = [headers]

    for processed_supply in processed_supplies:
        old_header = processed_supply["old_header"]
        matching_columns = [
            column_number
            for column_number, header in enumerate(headers, start=1)
            if header == old_header
        ]

        if not matching_columns:
            raise ExcelSynchronizationError(f"Не найден столбец нового поступления с заголовком {old_header!r}")

        if len(matching_columns) > 1:
            raise ExcelSynchronizationError(f"В Excel найдено несколько столбцов с заголовком поступления {old_header!r}")

        column_number = matching_columns[0]
        supply_date = format_supply_date(processed_supply["date"])
        new_header = f"Поступление_#{processed_supply['supply_batch_id']}_{supply_date}"

        worksheet.cells(1, column_number).value = new_header
        headers[column_number - 1] = new_header


def show_error_message(app, error, excel_restored):
    """Показывает пользователю сообщение об ошибке синхронизации."""
    if excel_restored:
        message = (
            "Ручная синхронизация не выполнена.\n\n"
            "Изменения в базе данных отменены.\n"
            "Excel-файл восстановлен из резервной копии и открыт повторно.\n\n"
            f"Причина ошибки:\n{error}"
        )
    else:
        message = (
            "Ручная синхронизация не выполнена.\n\n"
            "Изменения в базе данных отменены.\n"
            "Автоматически восстановить Excel-файл не удалось.\n"
            "Резервная копия сохранена в папке с прайсом.\n\n"
            f"Причина ошибки:\n{error}"
        )

    app.alert(message, title="Ошибка ручной синхронизации", buttons="ok", mode="critical")


def show_success_message(app, sync_result):
    """Показывает пользователю результат ручной синхронизации."""
    product_changes_count = len(sync_result["product_changes"])
    changed_supplies_count = len(sync_result["existing_supply_changes"])
    changed_supply_items_count = count_existing_supply_item_changes(sync_result["existing_supply_changes"])
    new_supplies_count = len(sync_result["new_supplies"])
    remain_changes_count = len(sync_result["remain_changes"])

    if not has_manual_changes(sync_result):
        message = "Ручная синхронизация завершена.\n\nНовых изменений не обнаружено."
    else:
        message = (
            "Все ручные изменения успешно синхронизированы.\n\n"
            f"Изменений названий, размеров и цен: {product_changes_count}\n"
            f"Новых поступлений: {new_supplies_count}\n"
            f"Изменённых поступлений: {changed_supplies_count}\n"
            f"Изменённых позиций в поступлениях: {changed_supply_items_count}\n"
            f"Изменений остатков: {remain_changes_count}"
        )

    app.alert(message, title="Синхронизация завершена", buttons="ok", mode="info")



def proceed_manual_changes(workbook):
    """Запускает ручную синхронизацию с резервным копированием и транзакцией."""
    excel_path = Path(workbook.fullname)
    backup_path = excel_path.with_name(f"{excel_path.stem}_backup{excel_path.suffix}")
    app = workbook.app
    conn = connect_db()
    backup_created = False

    logger.info("Запуск ручной синхронизации: excel_file=%s", excel_path)

    try:
        workbook.save()
        logger.info("Текущие изменения Excel сохранены")

        create_backup(excel_path, backup_path)
        backup_created = True
        logger.info("Создана резервная копия Excel-файла: %s", backup_path.name)

        sync_result = process_manual_changes(conn, workbook)

        worksheet = get_worksheet(workbook, price_sheet)

        update_new_supply_headers(worksheet, sync_result["new_supplies"])
        logger.info("Обновлены заголовки новых поступлений: %s", len(sync_result["new_supplies"]))

        product_rows = get_product_rows(worksheet)
        update_excel_current_stock_col(product_rows, conn, worksheet)
        logger.info("Столбец «Наличие» обновлён из products.current_stock")

        workbook.save()
        logger.info("Excel-файл успешно обновлён и сохранён")

        conn.commit()
        logger.info("Изменения базы данных подтверждены: commit выполнен")

    except Exception as error:
        conn.rollback()
        logger.info("Изменения базы данных отменены: rollback выполнен")
        log_sync_error(error)

        excel_restored = False

        if backup_created and backup_path.exists():
            try:
                restore_excel_file(workbook, app, excel_path, backup_path)
                excel_restored = True
                logger.info("Excel-файл восстановлен из резервной копии и открыт повторно")
            except Exception:
                logger.exception("Не удалось автоматически восстановить Excel-файл")

        show_error_message(app, error, excel_restored)
        raise

    else:
        if backup_path.exists():
            backup_path.unlink()
            logger.info("Резервная копия удалена")

        logger.info("Ручная синхронизация успешно завершена")
        show_success_message(app, sync_result)

    finally:
        conn.close()



def main():
    """Получает книгу от xlwings и запускает ручную синхронизацию."""
    workbook = xw.Book.caller()
    proceed_manual_changes(workbook)

#
# def run_from_python(excel_file):
#     excel_path = Path(excel_file).resolve()
#     workbook = xw.Book(str(excel_path))
#     workbook.set_mock_caller()
#     main()
#
#
# if __name__ == "__main__":
#     if len(sys.argv) != 2:
#         raise RuntimeError("Укажите путь к тестовому Excel-файлу")
#
#     run_from_python(sys.argv[1])
