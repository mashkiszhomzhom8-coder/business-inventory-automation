from inventory_app.core.exceptions import DatabaseSystemError
from datetime import date

def get_db_products(conn, excel_products):
    """Получает из БД товары, указанные в Excel.

    Returns:
        Словарь товаров по ID.
    """
    product_ids = [product["id"] for product in excel_products]
    if not product_ids:
        return {}

    placeholders = ", ".join("?" for _ in product_ids)

    rows = conn.execute(
        f"""
        SELECT id, name, size, price, active
        FROM products
        WHERE id IN ({placeholders});
        """,
        product_ids
    ).fetchall()

    db_products = {
        row["id"]: {
            "name": row["name"],
            "size": row["size"],
            "price": row["price"],
            "active": row["active"]
        }
        for row in rows
    }

    missing_ids = set(product_ids) - set(db_products)
    if missing_ids:
        raise DatabaseSystemError(f"Товары с ID {sorted(missing_ids)} присутствуют в Excel, но отсутствуют в БД")

    return db_products


def validate_product_uniqueness(excel_products):
    """Проверяет уникальность пары названия и размера товара в Excel."""
    product_keys = set()

    for product in excel_products:
        product_key = (product["name"], product["size"])

        if product_key in product_keys:
            raise ValueError(f"В Excel обнаружен дубликат товара: название «{product['name']}», размер «{product['size']}»")

        product_keys.add(product_key)

def update_product_attributes(conn, excel_products):
    """Синхронизирует названия, размеры и цены товаров из Excel.

    Returns:
        Список обнаруженных изменений.
    """
    validate_product_uniqueness(excel_products)
    db_products = get_db_products(conn, excel_products)
    changes = []

    for excel_product in excel_products:
        product_id = excel_product["id"]
        db_product = db_products[product_id]

        for field in ("name", "size", "price"):
            old_value = db_product[field]
            new_value = excel_product[field]

            if new_value != old_value:
                changes.append({
                    "product_id": product_id,
                    "field": field,
                    "old_value": old_value,
                    "new_value": new_value
                })

    changed_name_or_size_ids = {
        change["product_id"]
        for change in changes
        if change["field"] in ("name", "size")
    }

    changed_product_ids = {change["product_id"] for change in changes}

    for product_id in changed_name_or_size_ids:
        conn.execute(
            """
            UPDATE products
            SET name = ?
            WHERE id = ?;
            """,
            (f"__temporary_product_{product_id}__", product_id)
        )

    for excel_product in excel_products:
        product_id = excel_product["id"]

        if product_id not in changed_product_ids:
            continue

        conn.execute(
            """
            UPDATE products
            SET name = ?, size = ?, price = ?
            WHERE id = ?;
            """,
            (
                excel_product["name"],
                excel_product["size"],
                excel_product["price"],
                product_id
            )
        )

    for change in changes:
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
                change["product_id"],
                change["field"],
                change["old_value"],
                change["new_value"]
            )
        )

    return changes



