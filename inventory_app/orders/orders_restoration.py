from pathlib import Path

import xlwings as xw

from inventory_app.core.exceptions import (
    OrderAlreadyActiveError,
    OrderIdNotFoundError,
    OrderRestorationInputError,
    OrderRestorationSystemError
)
from inventory_app.core.database import connect_db
from inventory_app.excel.excel_reader import get_worksheet, price_sheet, col_fst_ord
from inventory_app.orders.order_service_excel import (
    get_product_rows,
    update_excel_current_stock_col
)
from inventory_app.core.logging_config import setup_logging
from inventory_app.core.sync_utils import create_backup, restore_excel_file


logger = setup_logging()


def get_inactive_orders(conn):
    """Возвращает список неактивных заказов для отображения пользователю."""
    return conn.execute(
        """
        SELECT id, client_fio, tg_user, date
        FROM orders
        WHERE active = 0
        ORDER BY id;
        """
    ).fetchall()


def format_inactive_orders(orders):
    """Формирует текстовый список неактивных заказов."""
    return "\n".join(
        f'{order["id"]} — {order["client_fio"]} — {order["tg_user"] or "без Telegram"} — {order["date"]}'
        for order in orders
    )


def validate_order_ids(conn, raw_ids):
    """Проверяет введенные ID и возможность восстановления заказов."""
    if raw_ids is None:
        return None

    raw_ids = raw_ids.strip()

    if not raw_ids:
        raise OrderRestorationInputError("Не указано ни одного ID заказа")

    parts = [part.strip() for part in raw_ids.split(",")]

    if any(not part for part in parts):
        raise OrderRestorationInputError("ID заказов должны быть указаны через запятую")

    if any(not part.isdigit() for part in parts):
        raise OrderRestorationInputError("ID заказа должен быть целым положительным числом")

    order_ids = [int(part) for part in parts]

    if any(order_id <= 0 for order_id in order_ids):
        raise OrderRestorationInputError("ID заказа должен быть целым положительным числом")

    if len(order_ids) != len(set(order_ids)):
        raise OrderRestorationInputError("Один и тот же ID указан несколько раз")

    placeholders = ",".join("?" for _ in order_ids)

    rows = conn.execute(
        f"""
        SELECT id, active
        FROM orders
        WHERE id IN ({placeholders});
        """,
        order_ids
    ).fetchall()

    db_orders = {row["id"]: row["active"] for row in rows}

    missing_ids = [
        order_id
        for order_id in order_ids
        if order_id not in db_orders
    ]

    if missing_ids:
        raise OrderIdNotFoundError(
            f"Заказы с ID {', '.join(map(str, missing_ids))} отсутствуют в базе данных"
        )

    active_ids = [
        order_id
        for order_id in order_ids
        if db_orders[order_id] == 1
    ]

    if active_ids:
        raise OrderAlreadyActiveError(
            f"Заказы с ID {', '.join(map(str, active_ids))} уже активны"
        )

    return order_ids


def activate_orders(conn, order_ids):
    """Помечает выбранные заказы как активные."""
    placeholders = ",".join("?" for _ in order_ids)

    cursor = conn.execute(
        f"""
        UPDATE orders
        SET active = 1
        WHERE id IN ({placeholders})
          AND active = 0;
        """,
        order_ids
    )

    if cursor.rowcount != len(order_ids):
        raise OrderRestorationSystemError(
            "Не удалось восстановить все выбранные заказы"
        )


def restore_products_stock(conn, order_ids):
    """Повторно уменьшает наличие товаров по восстановленным заказам."""
    for order_id in order_ids:
        order_items = conn.execute(
            """
            SELECT product_id, quantity
            FROM order_items
            WHERE order_id = ?;
            """,
            (order_id,)
        ).fetchall()

        if not order_items:
            raise OrderRestorationSystemError(
                f"Для заказа с ID {order_id} не найдены товары"
            )

        for item in order_items:
            cursor = conn.execute(
                """
                UPDATE products
                SET current_stock = current_stock - ?
                WHERE id = ?;
                """,
                (item["quantity"], item["product_id"])
            )

            if cursor.rowcount != 1:
                raise OrderRestorationSystemError(
                    f"Не удалось изменить наличие товара с ID {item['product_id']}"
                )


def restore_product_movements(conn, order_ids):
    """Записывает движения товаров по восстановленным заказам."""
    for order_id in order_ids:
        order = conn.execute(
            """
            SELECT season_id, date
            FROM orders
            WHERE id = ?;
            """,
            (order_id,)
        ).fetchone()

        if order is None:
            raise OrderRestorationSystemError(
                f"Заказ с ID {order_id} не найден в базе данных"
            )

        order_items = conn.execute(
            """
            SELECT product_id, quantity
            FROM order_items
            WHERE order_id = ?;
            """,
            (order_id,)
        ).fetchall()

        if not order_items:
            raise OrderRestorationSystemError(
                f"Для заказа с ID {order_id} не найдены товары"
            )

        for item in order_items:
            conn.execute(
                """
                INSERT INTO prod_movement (
                    season_id,
                    product_id,
                    date,
                    type,
                    quantity,
                    source,
                    source_id
                )
                VALUES (?, ?, ?, ?, ?, ?, ?);
                """,
                (
                    order["season_id"],
                    item["product_id"],
                    order["date"],
                    "продажа",
                    -item["quantity"],
                    "order",
                    order_id
                )
            )


def find_order_columns_in_excel(worksheet, order_ids):
    """Находит столбцы заказов в прайс-листе по ID."""
    columns_by_id = {}

    last_column = worksheet.used_range.last_cell.column

    for column_num in range(col_fst_ord, last_column + 1):
        raw_header = worksheet.range((1, column_num)).value

        if raw_header is None:
            continue

        header = str(raw_header).strip()

        for order_id in order_ids:
            expected_start = f"Заказ #{order_id}."

            if header.startswith(expected_start):
                columns_by_id[order_id] = column_num
                break

    missing_ids = [
        order_id
        for order_id in order_ids
        if order_id not in columns_by_id
    ]

    if missing_ids:
        raise OrderRestorationSystemError(
            f"В Excel не найдены заказы с ID: {', '.join(map(str, missing_ids))}"
        )

    return columns_by_id


def restore_order_columns_in_excel(worksheet, order_ids):
    """Возвращает белую заливку столбцов восстановленных заказов."""
    columns_by_id = find_order_columns_in_excel(worksheet, order_ids)

    last_row = worksheet.used_range.last_cell.row

    for column_num in columns_by_id.values():
        worksheet.range(
            (1, column_num),
            (last_row, column_num)
        ).color = None


def update_current_stock_in_excel(worksheet, conn):
    """Обновляет столбец наличия товаров после восстановления заказов."""
    product_rows = get_product_rows(worksheet)
    update_excel_current_stock_col(product_rows, conn, worksheet)


def show_restoration_success(app, order_ids):
    """Показывает подтверждение успешного восстановления."""
    ids_text = ", ".join(map(str, order_ids))

    app.alert(
        f"Заказы успешно восстановлены.\n\nВосстановленные ID: {ids_text}",
        title="Восстановление заказов",
        buttons="ok",
        mode="info"
    )


def show_restoration_error(app, error, excel_restored):
    """Показывает системную ошибку восстановления."""
    if excel_restored:
        message = (
            "Восстановление заказов не выполнено.\n\n"
            "Изменения в базе данных отменены, Excel-файл восстановлен.\n\n"
            f"Ошибка: {error}"
        )
    else:
        message = (
            "Восстановление заказов не выполнено.\n\n"
            "Изменения в базе данных отменены, но автоматически "
            "восстановить Excel-файл не удалось.\n\n"
            f"Ошибка: {error}"
        )

    app.alert(
        message,
        title="Ошибка восстановления заказов",
        buttons="ok",
        mode="critical"
    )


def prepare_inactive_orders():
    """Подготавливает список неактивных заказов для VBA."""
    workbook = xw.Book.caller()
    conn = connect_db()

    try:
        service_sheet = workbook.sheets["_service"]

        orders = get_inactive_orders(conn)

        if orders:
            orders_text = format_inactive_orders(orders)
        else:
            orders_text = "Неактивных заказов нет."

        service_sheet.range("A2").value = orders_text

    finally:
        conn.close()



def restore_orders(workbook, raw_ids):
    """Восстанавливает заказы в БД и Excel."""
    excel_path = Path(workbook.fullname)
    backup_path = excel_path.with_name(
        f"{excel_path.stem}_backup{excel_path.suffix}"
    )

    app = workbook.app
    conn = connect_db()
    backup_created = False

    logger.info(
        "Запуск восстановления заказов: excel_file=%s",
        excel_path
    )

    try:
        order_ids = validate_order_ids(conn, raw_ids)

        logger.info(
            "Получены ID заказов для восстановления: %s",
            order_ids
        )

        workbook.save()
        create_backup(excel_path, backup_path)
        backup_created = True

        worksheet = get_worksheet(workbook, price_sheet)

        activate_orders(conn, order_ids)

        restore_products_stock(conn, order_ids)
        restore_product_movements(conn, order_ids)

        restore_order_columns_in_excel(worksheet, order_ids)
        update_current_stock_in_excel(worksheet, conn)

        workbook.save()
        conn.commit()

    except (
        OrderRestorationInputError,
        OrderIdNotFoundError,
        OrderAlreadyActiveError
    ) as error:
        conn.rollback()

        logger.warning(
            "Ошибка пользовательского ввода при восстановлении заказов: %s",
            error
        )

        app.alert(
            str(error),
            title="Ошибка восстановления заказов",
            buttons="ok",
            mode="critical"
        )

        return

    except Exception as error:
        conn.rollback()

        logger.exception(
            "Системная ошибка восстановления заказов"
        )

        excel_restored = False

        if backup_created and backup_path.exists():
            try:
                restore_excel_file(
                    workbook,
                    app,
                    excel_path,
                    backup_path
                )
                excel_restored = True

            except Exception:
                logger.exception(
                    "Не удалось автоматически восстановить Excel-файл"
                )

        show_restoration_error(
            app,
            error,
            excel_restored
        )

        return

    else:
        if backup_path.exists():
            backup_path.unlink()

        logger.info(
            "Восстановление заказов успешно завершено: order_ids=%s",
            order_ids
        )

        show_restoration_success(
            app,
            order_ids
        )

    finally:
        conn.close()


def main(raw_ids):
    """Получает книгу от xlwings и запускает восстановление заказов."""
    workbook = xw.Book.caller()
    restore_orders(workbook, raw_ids)