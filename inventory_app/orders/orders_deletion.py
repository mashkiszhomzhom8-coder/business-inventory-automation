from datetime import date
from pathlib import Path

import xlwings as xw

from inventory_app.core.exceptions import (
    OrderAlreadyDeletedError,
    OrderDeletionInputError,
    OrderDeletionSystemError,
    OrderIdNotFoundError,
)
from inventory_app.core.database import connect_db
from inventory_app.excel.excel_reader import get_worksheet, price_sheet
from inventory_app.orders.order_service_excel import (
    get_product_rows,
    update_excel_current_stock_col,
)
from inventory_app.core.logging_config import setup_logging
from inventory_app.core.sync_utils import create_backup, restore_excel_file

logger = setup_logging()


def validate_order_ids(conn, raw_ids):
    """Проверяет строку ID заказов и их доступность для деактивации.

    Returns:
        Список ID заказов или ``None`` при отмене ввода.
    """
    if raw_ids is None:
        return None

    raw_ids = str(raw_ids).strip()

    if not raw_ids:
        raise OrderDeletionInputError("Не указано ни одного ID заказа")

    parts = [part.strip() for part in raw_ids.split(",")]

    if any(not part for part in parts):
        raise OrderDeletionInputError(
            "ID заказов должны быть указаны через запятую"
        )

    if any(not part.isdigit() for part in parts):
        raise OrderDeletionInputError(
            "ID заказа должен быть целым положительным числом"
        )

    order_ids = [int(part) for part in parts]

    if any(order_id <= 0 for order_id in order_ids):
        raise OrderDeletionInputError(
            "ID заказа должен быть целым положительным числом"
        )

    if len(order_ids) != len(set(order_ids)):
        raise OrderDeletionInputError(
            "Один и тот же ID заказа указан несколько раз"
        )

    placeholders = ",".join("?" for _ in order_ids)

    rows = conn.execute(
        f"""
        SELECT id, active
        FROM orders
        WHERE id IN ({placeholders});
        """,
        order_ids
    ).fetchall()

    db_orders = {
        row["id"]: row["active"]
        for row in rows
    }

    missing_ids = [
        order_id
        for order_id in order_ids
        if order_id not in db_orders
    ]

    if missing_ids:
        raise OrderIdNotFoundError(
            f"Заказы с ID {', '.join(map(str, missing_ids))} "
            f"отсутствуют в базе данных"
        )

    deleted_ids = [
        order_id
        for order_id in order_ids
        if db_orders[order_id] == 0
    ]

    if deleted_ids:
        raise OrderAlreadyDeletedError(
            f"Заказы с ID {', '.join(map(str, deleted_ids))} уже удалены"
        )

    return order_ids


def deactivate_orders(conn, order_ids):
    """Помечает выбранные заказы как неактивные."""
    placeholders = ",".join("?" for _ in order_ids)

    cursor = conn.execute(
        f"""
        UPDATE orders
        SET active = 0
        WHERE id IN ({placeholders})
          AND active = 1;
        """,
        order_ids
    )

    if cursor.rowcount != len(order_ids):
        raise OrderDeletionSystemError(
            "Не удалось деактивировать все выбранные заказы"
        )


def restore_stock_after_order_deletion(conn, order_ids):
    """Возвращает товары удалённых заказов в наличие и пишет change_log."""
    for order_id in order_ids:
        items = conn.execute(
            """
            SELECT product_id, SUM(quantity) AS quantity
            FROM order_items
            WHERE order_id = ?
            GROUP BY product_id;
            """,
            (order_id,)
        ).fetchall()

        if not items:
            raise OrderDeletionSystemError(
                f"Для заказа с ID {order_id} не найдено ни одной позиции"
            )

        for item in items:
            product_id = item["product_id"]
            quantity = item["quantity"]

            product = conn.execute(
                """
                SELECT current_stock
                FROM products
                WHERE id = ?;
                """,
                (product_id,)
            ).fetchone()

            if product is None:
                raise OrderDeletionSystemError(
                    f"Для заказа с ID {order_id} "
                    f"не найден товар с ID {product_id}"
                )

            old_stock = product["current_stock"]
            new_stock = old_stock + quantity

            cursor = conn.execute(
                """
                UPDATE products
                SET current_stock = ?
                WHERE id = ?;
                """,
                (new_stock, product_id)
            )

            if cursor.rowcount != 1:
                raise OrderDeletionSystemError(
                    f"Не удалось пересчитать наличие товара "
                    f"с ID {product_id}"
                )

            conn.execute(
                """
                INSERT INTO change_log (
                    date,
                    product_id,
                    field,
                    old_value,
                    new_value
                )
                VALUES (?, ?, ?, ?, ?);
                """,
                (
                    date.today().isoformat(),
                    product_id,
                    "current_stock",
                    str(old_stock),
                    str(new_stock),
                )
            )


def find_order_columns_in_excel(worksheet, order_ids):
    """Находит столбцы прайса для указанных ID заказов.

    Returns:
        Словарь номеров столбцов по ID заказа.
    """
    columns_by_id = {}
    last_column = worksheet.used_range.last_cell.column

    for column_num in range(1, last_column + 1):
        raw_header = worksheet.range((1, column_num)).value

        if raw_header is None:
            continue

        header = str(raw_header).strip()

        for order_id in order_ids:
            expected_prefix = f"Заказ #{order_id}."

            if header.startswith(expected_prefix):
                if order_id in columns_by_id:
                    raise OrderDeletionSystemError(
                        f"В Excel найдено несколько столбцов "
                        f"заказа с ID {order_id}"
                    )

                columns_by_id[order_id] = column_num

    missing_ids = [
        order_id
        for order_id in order_ids
        if order_id not in columns_by_id
    ]

    if missing_ids:
        raise OrderDeletionSystemError(
            f"В Excel не найдены заказы с ID: "
            f"{', '.join(map(str, missing_ids))}"
        )

    return columns_by_id


def delete_order_columns_from_excel(worksheet, order_ids):
    """Помечает выбранные заказы в Excel серой заливкой."""
    columns_by_id = find_order_columns_in_excel(
        worksheet,
        order_ids
    )

    last_row = worksheet.range(
        f"A{worksheet.cells.last_cell.row}"
    ).end("up").row

    for column_num in columns_by_id.values():
        worksheet.range(
            (1, column_num),
            (last_row, column_num)
        ).color = (166, 166, 166)


def update_current_stock_in_excel(worksheet, conn):
    """Обновляет столбец «Наличие» после отмены заказов."""
    product_rows = get_product_rows(worksheet)
    update_excel_current_stock_col(
        product_rows,
        conn,
        worksheet
    )


def show_deletion_success(app, order_ids):
    """Показывает пользователю подтверждение удаления заказов."""
    ids_text = ", ".join(map(str, order_ids))

    app.alert(
        f"Заказы успешно деактивированы.\n\n"
        f"Неактивные ID: {ids_text}",
        title="Деактивация заказов",
        buttons="ok",
        mode="info"
    )


def show_deletion_error(app, error, excel_restored):
    """Показывает системную ошибку деактивации заказов."""
    if excel_restored:
        message = (
            "Деактивация заказов не выполнена.\n\n"
            "Изменения в базе данных отменены, "
            "Excel-файл восстановлен.\n\n"
            f"Ошибка: {error}"
        )
    else:
        message = (
            "Деактивация заказов не выполнена.\n\n"
            "Изменения в базе данных отменены, "
            "но автоматически восстановить Excel-файл "
            "не удалось.\n\n"
            f"Ошибка: {error}"
        )

    app.alert(
        message,
        title="Ошибка деактивации заказов",
        buttons="ok",
        mode="critical"
    )


def delete_orders(workbook, raw_ids):
    """Деактивирует заказы, возвращает товары в наличие и обновляет Excel."""
    excel_path = Path(workbook.fullname)
    backup_path = excel_path.with_name(
        f"{excel_path.stem}_backup{excel_path.suffix}"
    )

    app = workbook.app
    conn = connect_db()
    backup_created = False

    logger.info(
        "Запуск деактивации заказов: excel_file=%s",
        excel_path
    )

    try:
        order_ids = validate_order_ids(conn, raw_ids)

        if order_ids is None:
            return

        logger.info(
            "Получены ID заказов для деактивации: %s",
            order_ids
        )

        workbook.save()
        create_backup(excel_path, backup_path)
        backup_created = True

        worksheet = get_worksheet(
            workbook,
            price_sheet
        )

        deactivate_orders(conn, order_ids)
        restore_stock_after_order_deletion(
            conn,
            order_ids
        )
        delete_order_columns_from_excel(
            worksheet,
            order_ids
        )
        update_current_stock_in_excel(
            worksheet,
            conn
        )

        workbook.save()
        conn.commit()

        logger.info("Успешное удаление заказов")

    except (
        OrderDeletionInputError,
        OrderIdNotFoundError,
        OrderAlreadyDeletedError,
    ) as error:
        conn.rollback()

        logger.warning(
            "Ошибка пользовательского ввода "
            "при деактивации заказов: %s",
            error
        )

        app.alert(
            str(error),
            title="Ошибка деактивации заказов",
            buttons="ok",
            mode="critical"
        )
        return

    except Exception as error:
        conn.rollback()
        logger.exception(
            "Системная ошибка деактивации заказов"
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
                    "Не удалось автоматически "
                    "восстановить Excel-файл"
                )

        show_deletion_error(
            app,
            error,
            excel_restored
        )
        return

    else:
        if backup_path.exists():
            backup_path.unlink()

        show_deletion_success(
            app,
            order_ids
        )

    finally:
        conn.close()


def main(raw_ids):
    """Получает книгу от xlwings и запускает деактивацию заказов."""
    logger.info(type(raw_ids))
    workbook = xw.Book.caller()
    delete_orders(workbook, raw_ids)
