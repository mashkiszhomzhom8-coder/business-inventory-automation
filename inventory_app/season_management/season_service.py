from datetime import date
from pathlib import Path

import xlwings as xw

from inventory_app.core.database import connect_db, make_seasons_end_date_nullable
from inventory_app.core.exceptions import SeasonValidationError
from inventory_app.core.logging_config import setup_logging
from inventory_app.excel.excel_reader import col_curr_st
from inventory_app.orders.order_service_excel import get_product_rows, update_excel_current_stock_col
from inventory_app.season_management.season_files import (
    create_archive_copy,
    create_new_season_copy,
    delete_new_season_copy,
)
from inventory_app.season_management.season_prep_valid import (
    FIRST_PRODUCT_ROW,
    PRICE_SHEET_NAME,
    get_active_season,
    get_excel_order_columns,
    get_excel_supply_columns,
    get_last_product_row,
    get_sheet,
    normalize_int,
    read_excel_products,
    validate_excel_products_against_database,
    validate_excel_structure,
    validate_season_can_be_closed,
)


logger = setup_logging()


def get_closing_stock(conn):
    """Возвращает наличие всех товаров на момент закрытия сезона."""

    rows = conn.execute(
        """
        SELECT id, current_stock
        FROM products
        ORDER BY id
        """
    ).fetchall()

    return {
        row["id"]: row["current_stock"]
        for row in rows
    }


def fill_opening_stock(conn, season_id, closing_stock):
    """Фиксирует стартовые остатки нового сезона."""

    for product_id, opening_stock in closing_stock.items():
        conn.execute(
            """
            INSERT INTO season_opening_stock (
                season_id,
                product_id,
                opening_stock
            )
            VALUES (?, ?, ?)
            """,
            (season_id, product_id, opening_stock)
        )


def create_opening_inventory_check(conn, season_id, closing_stock, close_date):
    """Создаёт начальную инвентаризацию нового сезона."""

    cursor = conn.execute(
        """
        INSERT INTO inventory_check (
            season_id,
            date
        )
        VALUES (?, ?)
        """,
        (season_id, close_date)
    )

    inventory_check_id = cursor.lastrowid

    for product_id, opening_stock in closing_stock.items():
        conn.execute(
            """
            INSERT INTO inventory_items (
                inventory_check_id,
                product_id,
                expected_remain,
                actual_remain,
                difference
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                inventory_check_id,
                product_id,
                opening_stock,
                opening_stock,
                0,
            )
        )


def remove_season_columns(sheet):
    """Удаляет из нового прайса все операционные столбцы после H."""

    last_column = sheet.range((1, sheet.cells.last_cell.column)).end("left").column

    if last_column <= col_curr_st:
        return

    first_column = sheet.range((1, col_curr_st + 1)).address.split("$")[1]
    last_column_letter = sheet.range((1, last_column)).address.split("$")[1]

    sheet.range(f"{first_column}:{last_column_letter}").delete()


def prepare_new_season_excel(book, closing_stock):
    """Очищает копию прайса и заполняет стартовые остатки нового сезона."""

    sheet = get_sheet(book, PRICE_SHEET_NAME)
    last_row = get_last_product_row(sheet)

    remove_season_columns(sheet)

    opening_values = []
    empty_supply_values = []

    for row in range(FIRST_PRODUCT_ROW, last_row + 1):
        raw_product_id = sheet.range((row, 1)).value

        if raw_product_id is None or str(raw_product_id).strip() == "":
            opening_values.append([None])
            empty_supply_values.append([None])
            continue

        product_id = normalize_int(raw_product_id, f"ID, строка {row}")

        if product_id not in closing_stock:
            raise SeasonValidationError(
                f"Не найден стартовый остаток для товара с ID {product_id}"
            )

        opening_values.append([closing_stock[product_id]])
        empty_supply_values.append([0])

    sheet.range((FIRST_PRODUCT_ROW, 6), (last_row, 6)).value = empty_supply_values
    sheet.range((FIRST_PRODUCT_ROW, 7), (last_row, 7)).value = opening_values


def close_current_season(conn, active_season, close_date):
    """Закрывает текущий сезон и создаёт запись нового active-сезона."""

    conn.execute(
        """
        UPDATE seasons
        SET status = 'closed', end_date = ?
        WHERE id = ?
          AND status = 'active'
        """,
        (close_date, active_season["id"])
    )


def open_new_season(conn, close_date, new_excel_path):
    """Создаёт новый активный сезон."""

    cursor = conn.execute(
        """
        INSERT INTO seasons (
            start_date,
            end_date,
            excel_file,
            status
        )
        VALUES (?, ?, ?, 'active')
        """,
        (close_date, None, str(Path(new_excel_path).resolve()))
    )

    return cursor.lastrowid


def validate_new_season_excel(conn, book, season_id):
    """Проверяет готовую книгу нового сезона до подтверждения транзакции."""

    validate_excel_structure(book)

    products = read_excel_products(book)
    validate_excel_products_against_database(conn, products)

    sheet = get_sheet(book, PRICE_SHEET_NAME)

    if get_excel_order_columns(book):
        raise SeasonValidationError(
            "В новом сезоне остались столбцы заказов."
        )

    if get_excel_supply_columns(book) != [6]:
        raise SeasonValidationError(
            "В новом сезоне остались дополнительные столбцы поступлений."
        )

    opening_stock = {
        row["product_id"]: row["opening_stock"]
        for row in conn.execute(
            """
            SELECT product_id, opening_stock
            FROM season_opening_stock
            WHERE season_id = ?
            """,
            (season_id,)
        ).fetchall()
    }

    for product_id, product in products.items():
        row = product["row"]
        supply_value = normalize_int(
            sheet.range((row, 6)).value,
            f"Поступление исход, строка {row}"
        )
        remain_value = normalize_int(
            sheet.range((row, 7)).value,
            f"Остатки исход, строка {row}"
        )

        if supply_value != 0:
            raise SeasonValidationError(
                f"В новом сезоне не очищено поступление для товара ID {product_id}"
            )

        if product_id not in opening_stock or remain_value != opening_stock[product_id]:
            raise SeasonValidationError(
                f"Некорректный стартовый остаток товара ID {product_id}"
            )


def show_success_message(app, active_season, close_date):
    app.alert(
        "Сезон успешно закрыт.\n\n"
        f"Закрытый сезон:\n{active_season['start_date']} — {close_date}\n\n"
        f"Создан новый сезон:\nс {close_date}\n\n"
        "Новый рабочий файл открыт.",
        title="Закрытие сезона",
        buttons="ok",
        mode="info"
    )


def show_error_message(app, error):
    app.alert(
        "Закрытие сезона не выполнено.\n\n"
        "Изменения отменены.\n"
        "Рабочим остаётся файл текущего сезона.\n\n"
        f"Причина:\n{error}",
        title="Ошибка закрытия сезона",
        buttons="ok",
        mode="critical"
    )


def close_season(workbook):
    """Закрывает текущий сезон и открывает новый рабочий Excel-файл."""

    app = workbook.app
    conn = connect_db()
    new_book = None
    new_excel_path = None

    try:
        logger.info("Запуск закрытия сезона: excel_file=%s", workbook.fullname)

        active_season, _ = validate_season_can_be_closed(conn, workbook)
        closing_stock = get_closing_stock(conn)
        close_date = date.today().isoformat()

        create_archive_copy(
            workbook.fullname,
            active_season["start_date"],
            close_date
        )

        new_excel_path = create_new_season_copy(
            workbook.fullname,
            close_date
        )

        new_book = app.books.open(str(new_excel_path))
        prepare_new_season_excel(new_book, closing_stock)
        new_book.save()

        make_seasons_end_date_nullable(conn)
        conn.execute("BEGIN")

        current_active_season = get_active_season(conn)

        if current_active_season["id"] != active_season["id"]:
            raise SeasonValidationError(
                "Активный сезон изменился во время закрытия. Повторите операцию."
            )

        close_current_season(conn, active_season, close_date)
        new_season_id = open_new_season(conn, close_date, new_excel_path)
        fill_opening_stock(conn, new_season_id, closing_stock)
        create_opening_inventory_check(
            conn,
            new_season_id,
            closing_stock,
            close_date
        )

        worksheet = get_sheet(new_book, PRICE_SHEET_NAME)
        product_rows = get_product_rows(worksheet)
        update_excel_current_stock_col(product_rows, conn, worksheet)
        new_book.save()

        validate_new_season_excel(conn, new_book, new_season_id)
        conn.commit()

    except Exception as error:
        conn.rollback()
        logger.exception("Ошибка закрытия сезона")

        if new_book is not None:
            try:
                new_book.close()
            except Exception:
                logger.exception("Не удалось закрыть новый Excel-файл")

        if new_excel_path:
            try:
                delete_new_season_copy(new_excel_path)
            except Exception:
                logger.exception("Не удалось удалить незавершённый Excel-файл")

        show_error_message(app, error)
        raise

    else:
        workbook.close()
        new_book.activate()
        show_success_message(app, active_season, close_date)
        logger.info("Сезон успешно закрыт: old_season_id=%s", active_season["id"])

    finally:
        conn.close()


def main():
    workbook = xw.Book.caller()
    close_season(workbook)
