# from excel_reader import excel_order_reader
from inventory_app.core.database import get_active_season_id
from inventory_app.core.exceptions import DatabaseSystemError, InactiveProductInOrderError, OrderDataError


def find_prod_info(conn, orders):
    """Находит товары заказа и рассчитывает суммы.

    Returns:
        Данные позиций и итогов для каждого заказа.
    """
    total_ord_info = []
    for order in orders:
        found_prods = []
        total_summ = 0

        items = order['items']
        for item in items:
            name = item["prod_name"]
            size = item["size"]
            quantity = item["quantity"]
            product = conn.execute(
                """
                SELECT id, price, active
                FROM products
                WHERE name = ? AND size = ?
                ORDER BY active DESC
                LIMIT 1
                """,
                (name, size)
            ).fetchone()

            if product is None:
                raise OrderDataError(f"Товар из заказа не найден: название={name!r}, размер={size!r}")

            if product["active"] == 0:
                raise InactiveProductInOrderError(
                    f"Товар из заказа был ранее удален: название={name!r}, размер={size!r}"
                )


            total_summ += quantity * product["price"]
            found_prods.append({
                "product_id": product["id"],
                "quantity": quantity,
                "price_at_moment": product["price"]
            })

        total_ord_info.append({"found_prods": found_prods, 'total_summ': total_summ})
    return total_ord_info

def fill_table_orders(conn, orders, total_ord_info, season_id):
    """Добавляет заказы в БД.

    Returns:
        Список ID созданных заказов.
    """
    order_ids = []

    for order, order_info in zip(orders, total_ord_info):
        cursor = conn.execute(
            """
            INSERT INTO orders(
                season_id,
                client_fio,
                tg_user,
                date,
                total_sum
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                season_id,
                order["client_fio"],
                order["tg_user"],
                order["date"],
                order_info["total_summ"]
            )
        )

        order_ids.append(cursor.lastrowid)

    return order_ids



def fill_table_order_items(conn, order_ids, total_ord_info):
    """Добавляет позиции созданных заказов в БД."""
    for order_id, order_info in zip(order_ids, total_ord_info):
        for item in order_info["found_prods"]:
            conn.execute(
                """
                INSERT INTO order_items(
                    order_id,
                    product_id,
                    quantity,
                    price_at_moment
                )
                VALUES (?, ?, ?, ?)
                """,
                (order_id, item["product_id"], item["quantity"], item["price_at_moment"])
            )

def fill_table_prod_movement(conn, orders, order_ids, total_ord_info, season_id):
    """Записывает движения товаров по проведённым заказам."""
    for order, order_id, order_info in zip(orders, order_ids, total_ord_info):
        for item in order_info["found_prods"]:
            conn.execute(
                """
                INSERT INTO prod_movement(
                    season_id,
                    product_id,
                    date,
                    type,
                    quantity,
                    source,
                    source_id
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (season_id, item["product_id"], order["date"], 'продажа', -item["quantity"], 'order', order_id)
            )

def update_table_products(conn, total_ord_info):
    """Уменьшает текущие остатки товаров по позициям заказов."""
    for order_info in total_ord_info:
        for item in order_info["found_prods"]:
            cursor = conn.execute(
                """
                UPDATE products
                SET
                current_stock = current_stock - ?
                WHERE
                id = ?;
            """,
                (item['quantity'], item['product_id'])
            )

            if cursor.rowcount == 0:
                raise DatabaseSystemError(f"Не удалось обновить остаток товара с ID {item['product_id']}")


def process_orders(conn, orders):
    """Проводит заказы в базе данных.

    Returns:
        Кортеж из ID заказов и рассчитанных данных их позиций.
    """
    season_id = get_active_season_id(conn)
    total_ord_info = find_prod_info(conn, orders)
    order_ids = fill_table_orders(conn, orders, total_ord_info, season_id)
    fill_table_order_items(conn, order_ids, total_ord_info)
    fill_table_prod_movement(conn, orders, order_ids, total_ord_info, season_id)
    update_table_products(conn, total_ord_info)

    return order_ids, total_ord_info #возможно стоит удалить total_ord_info

# if __name__ == "__main__":
#     orders = excel_order_reader("Шаблон_прайса.xlsx")
#     process_orders(orders) # версии с ботом/без бота будут иметь разные источники считавания заказа, так что передаем его в параметр функции

