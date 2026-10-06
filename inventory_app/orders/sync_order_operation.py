import sys

from pathlib import Path

import xlwings as xw

from inventory_app.core.database import connect_db
from inventory_app.excel.excel_reader import excel_order_reader, get_worksheet, price_sheet
from inventory_app.orders.order_service_db import process_orders
from inventory_app.orders.order_service_excel import update_excel_after_orders_processing
from inventory_app.core.exceptions import ExcelDataError, ExcelSynchronizationError
from inventory_app.core.sync_utils import create_backup, restore_excel_file, log_sync_error
from inventory_app.core.logging_config import setup_logging

logger = setup_logging()



def delete_processed_order_sheets(workbook, orders):
    """Удаляет листы заказов, успешно проведённых в БД."""
    sheet_names = [sheet.name for sheet in workbook.sheets]

    for order in orders:
        sheet_name = order["sheet_name"]

        if sheet_name not in sheet_names:
            raise ExcelSynchronizationError(f"Лист заказа «{sheet_name}» отсутствует в Excel-файле")

        workbook.sheets[sheet_name].delete()
        sheet_names.remove(sheet_name)


def show_error_message(app, error, excel_restored):
    """Показывает пользователю сообщение об ошибке проведения заказов."""
    if excel_restored:
        message = f"Синхронизация не выполнена.\n\nИзменения в базе данных отменены.\nExcel-файл восстановлен и открыт повторно.\n\nПричина ошибки:\n{error}"
    else:
        message = f"Синхронизация не выполнена.\n\nИзменения в базе данных отменены.\nАвтоматически восстановить Excel-файл не удалось.\nРезервная копия сохранена в папке с прайсом.\n\nПричина ошибки:\n{error}"

    app.alert(message, title="Ошибка синхронизации", buttons="ok", mode="critical")

def show_success_message(app, processed_orders_count):
    """Показывает пользователю число проведённых заказов."""
    message = f"Все заказы успешно проведены.\n\nОбработано заказов: {processed_orders_count}"
    app.alert(message, title="Синхронизация завершена", buttons="ok", mode="info")



def proceed_orders(workbook):
    """Проводит заказы Excel с резервным копированием и транзакцией."""
    excel_path = Path(workbook.fullname)
    backup_path = excel_path.with_name(f"{excel_path.stem}_backup{excel_path.suffix}")
    app = workbook.app
    conn = connect_db()
    backup_created = False

    logger.info("Запуск синхронизации заказов: excel_file=%s", excel_path)

    try:
        workbook.save()
        logger.info("Текущие изменения Excel сохранены")

        create_backup(excel_path, backup_path)
        backup_created = True
        logger.info("Создана резервная копия Excel-файла: %s", backup_path.name)

        orders = excel_order_reader(workbook)
        logger.info("Найдено листов заказов: %s", len(orders))

        if not orders:
            raise ExcelDataError("В Excel не найдено листов заказов для обработки. Проверьте, чтоб имя листа начиналось со слова 'Заказ'. ")

        order_ids, total_ord_info = process_orders(conn, orders)
        logger.info("Успешно обработано заказов: %s; order_ids=%s", len(order_ids), order_ids)

        worksheet = get_worksheet(workbook, price_sheet)
        update_excel_after_orders_processing(order_ids, conn, worksheet)
        delete_processed_order_sheets(workbook, orders)

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
        logger.info("Синхронизация заказов успешно завершена")
        show_success_message(app, len(order_ids))

    finally:
        conn.close()


def main():
    """Получает книгу от xlwings и запускает проведение заказов."""
    workbook = xw.Book.caller()
    proceed_orders(workbook)


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
