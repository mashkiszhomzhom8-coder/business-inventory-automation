from inventory_app.excel.excel_reader import col_curr_st, col_fst_ord
from inventory_app.core.exceptions import DatabaseSystemError, DeletedActiveError, ExcelDataError


def find_next_order_column(worksheet):
    """Находит первый свободный столбец для нового заказа."""
    last_used_col = worksheet.used_range.last_cell.column

    if last_used_col < col_fst_ord:
        return col_fst_ord

    headers = worksheet.range((1, col_fst_ord), (1, last_used_col)).options(ndim=1).value
    last_filled_col = col_fst_ord - 1

    for offset, header in enumerate(headers):
        if header is not None and str(header).strip() != "":
            last_filled_col = col_fst_ord + offset

    return max(col_fst_ord, last_filled_col + 1)


def get_product_rows(worksheet):
    """Возвращает соответствие ID товаров и строк прайс-листа."""
    last_row = worksheet.used_range.last_cell.row

    if last_row < 2:
        return {}

    product_ids = worksheet.range((2, 1), (last_row, 1)).options(ndim=1).value
    product_rows = {}

    for row, raw_product_id in enumerate(product_ids, start=2):
        if raw_product_id is None or str(raw_product_id).strip() == "":
            continue

        try:
            product_id = int(raw_product_id)
        except (TypeError, ValueError) as exc:
            raise ExcelDataError(f"Строка {row}: ID товара должен быть целым числом") from exc

        if product_id in product_rows:
            raise ExcelDataError(f"Строка {row}: ID товара {product_id} повторяется в Excel")

        product_rows[product_id] = row

    return product_rows


def update_excel_orders_col(order_ids, product_rows, conn, worksheet):
    """Добавляет в Excel столбцы с позициями проведённых заказов."""
    if not product_rows:
        raise ExcelDataError("На листе «Шаблон прайса» не найдено товаров с заполненным ID")

    column_number = find_next_order_column(worksheet)
    last_product_row = max(product_rows.values())

    for order_id in order_ids:
        order = conn.execute(
            """
            SELECT client_fio, tg_user, date
            FROM orders
            WHERE id = ?;
            """,
            (order_id,)
        ).fetchone()

        if order is None:
            raise DatabaseSystemError(f"Созданный заказ с ID {order_id} не найден в БД")

        header = f"Заказ #{order_id}. {order['client_fio']}, {order['tg_user']}, {order['date']}"
        worksheet.range((1, column_number)).value = header

        order_items = conn.execute(
            """
            SELECT product_id, quantity
            FROM order_items
            WHERE order_id = ?;
            """,
            (order_id,)
        ).fetchall()

        column_values = [[None] for _ in range(last_product_row - 1)]

        for item in order_items:
            product_id = item["product_id"]

            if product_id not in product_rows:
                raise DeletedActiveError(f"Активный товар с ID {product_id} отсутствует в прайс-листе")

            excel_row = product_rows[product_id]
            column_values_index = excel_row - 2
            column_values[column_values_index][0] = item["quantity"]

        worksheet.range((2, column_number), (last_product_row, column_number)).value = column_values
        column_number += 1


def update_excel_current_stock_col(product_rows, conn, worksheet):
    """Обновляет столбец текущих остатков активных товаров."""
    if not product_rows:
        raise ExcelDataError("На листе «Шаблон прайса» не найдено товаров с заполненным ID")

    products = conn.execute(
        """
        SELECT id, current_stock
        FROM products
        WHERE active = 1;
        """
    ).fetchall()

    last_product_row = max(product_rows.values())
    stock_values = [[None] for _ in range(last_product_row - 1)]

    for product in products:
        product_id = product["id"]

        if product_id not in product_rows:
            raise DeletedActiveError(f"Активный товар с ID {product_id} отсутствует в прайс-листе")

        excel_row = product_rows[product_id]
        stock_values_index = excel_row - 2
        stock_values[stock_values_index][0] = product["current_stock"]

    worksheet.range((2, col_curr_st), (last_product_row, col_curr_st)).value = stock_values


def update_excel_after_orders_processing(order_ids, conn, worksheet):
    """Обновляет заказы и текущие остатки в прайс-листе."""
    product_rows = get_product_rows(worksheet)
    update_excel_orders_col(order_ids, product_rows, conn, worksheet)
    update_excel_current_stock_col(product_rows, conn, worksheet)
