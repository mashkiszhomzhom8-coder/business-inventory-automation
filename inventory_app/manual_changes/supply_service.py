from inventory_app.core.exceptions import DatabaseSystemError, ExcelDataError
from datetime import date


INITIAL_SUPPLY_HEADER = "поступление исход"


def find_new_supplies(supplies):
    """Возвращает поступления без присвоенного ID партии."""
    new_supplies = []
    for supply in supplies:
        if supply["header"].strip().lower() == INITIAL_SUPPLY_HEADER:
            continue
        if supply["supply_batch_id"] is None:
            new_supplies.append(supply)

    return new_supplies

def validate_new_supply(supply):
    """Проверяет дату, позиции и количества нового поступления."""
    if not supply["date"]:
        raise ExcelDataError(f"Для поступления {supply['header']!r} не указана дата")

    if not supply["items"]:
        raise ExcelDataError(f"Поступление {supply['header']!r} не содержит товаров")

    product_ids = set()

    for item in supply["items"]:
        product_id = item["product_id"]
        quantity = item["quantity"]

        if product_id in product_ids:
            raise ExcelDataError(f"В поступлении {supply['header']!r} товар с ID {product_id} указан несколько раз")
        if quantity < 0:
            raise ExcelDataError(f"В поступлении {supply['header']!r} количество товара с ID {product_id} должно быть неотрицательным")

        product_ids.add(product_id)


def validate_supply_products(conn, supply):
    """Проверяет товары поступления и возвращает неактивные ID."""
    inactive_product_ids = []

    for item in supply["items"]:
        product = conn.execute(
            """
            SELECT active
            FROM products
            WHERE id = ?;
            """,
            (item["product_id"],)
        ).fetchone()

        if product is None:
            raise DatabaseSystemError(f"Товар с ID {item['product_id']} есть в поступлении Excel, но отсутствует в БД")

        if product["active"] == 0:
            inactive_product_ids.append(item["product_id"])

    return inactive_product_ids

def insert_supply_batch(conn, season_id, supply):
    """Создаёт партию поступления и возвращает её ID."""
    cursor = conn.execute(
        """
        INSERT INTO supply_batch (
            season_id,
            date
        )
        VALUES (?, ?);
        """,
        (season_id, supply["date"])
    )

    return cursor.lastrowid


def insert_supply_items(conn, supply_batch_id, supply):
    """Добавляет позиции в партию поступления."""
    for item in supply["items"]:
        product = conn.execute(
            """
            SELECT price
            FROM products
            WHERE id = ?;
            """,
            (item["product_id"],)
        ).fetchone()

        if product is None:
            raise DatabaseSystemError(f"Не удалось получить цену товара с ID {item['product_id']}")

        conn.execute(
            """
            INSERT INTO supply_items (
                supply_batch_id,
                product_id,
                quantity,
                price_at_moment
            )
            VALUES (?, ?, ?, ?);
            """,
            (
                supply_batch_id,
                item["product_id"],
                item["quantity"],
                product["price"]
            )
        )


def insert_supply_movements(conn, season_id, supply_batch_id, supply):
    """Записывает движения товаров по новому поступлению."""
    for item in supply["items"]:
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
                season_id,
                item["product_id"],
                supply["date"],
                "поступление",
                item["quantity"],
                "supply_batch",
                supply_batch_id
            )
        )


def update_products(conn, supply, inactive_product_ids):
    """Восстанавливает товары и увеличивает их текущий остаток."""
    for product_id in inactive_product_ids:
        cursor = conn.execute(
            """
            UPDATE products
            SET active = 1
            WHERE id = ?;
            """,
            (product_id,)
        )
        if cursor.rowcount != 1:
            raise DatabaseSystemError(f"Не удалось восстановить товар с ID {product_id}")

    for item in supply["items"]:
        cursor = conn.execute(
            """
            UPDATE products
            SET current_stock = current_stock + ?
            WHERE id = ?;
            """,
            (item["quantity"], item["product_id"])
        )

        if cursor.rowcount != 1:
            raise DatabaseSystemError(f"Не удалось обновить наличие товара с ID {item['product_id']}")


def process_one_new_supply(conn, season_id, supply):
    """Проводит одно новое поступление.

    Returns:
        Данные для обновления заголовка Excel.
    """
    validate_new_supply(supply)
    inactive_product_ids = validate_supply_products(conn, supply)

    supply_batch_id = insert_supply_batch(conn, season_id, supply)
    insert_supply_items(conn, supply_batch_id, supply)
    insert_supply_movements(conn, season_id, supply_batch_id, supply)
    update_products(conn, supply, inactive_product_ids)

    return {
        "old_header": supply["header"],
        "supply_batch_id": supply_batch_id,
        "date": supply["date"]
    }


def process_new_supplies(conn, season_id, supplies):
    """Проводит все новые поступления.

    Returns:
        Список обработанных поступлений.
    """
    new_supplies = find_new_supplies(supplies)
    processed_supplies = []

    for supply in new_supplies:
        processed_supplies.append(process_one_new_supply(conn, season_id, supply))

    return processed_supplies



"""Обрабатывает изменение существующих поступлений."""


def find_existing_supplies(supplies):
    """Возвращает поступления с существующим ID партии."""
    return [supply for supply in supplies if supply["supply_batch_id"] is not None]


def get_supply_items_from_db(conn, supply_batch_id):
    """Получает позиции партии поступления из БД.

    Returns:
        Словарь количеств по ID товара.
    """
    rows = conn.execute(
        """
        SELECT product_id, quantity
        FROM supply_items
        WHERE supply_batch_id = ?;
        """,
        (supply_batch_id,)
    ).fetchall()

    return {row["product_id"]: row["quantity"] for row in rows}


def get_supply_items_from_excel(supply):
    """Преобразует позиции поступления Excel в словарь по ID товара."""
    return {item["product_id"]: item["quantity"] for item in supply["items"]}


def find_supply_item_changes(conn, supply):
    """Находит отличия позиций поступления от данных БД.

    Returns:
        Список изменений количеств.
    """
    db_items = get_supply_items_from_db(conn, supply["supply_batch_id"])
    excel_items = get_supply_items_from_excel(supply)
    product_ids = set(db_items) | set(excel_items)
    changes = []

    for product_id in product_ids:
        old_quantity = db_items.get(product_id, 0)
        new_quantity = excel_items.get(product_id, 0)

        if old_quantity == new_quantity:
            continue

        changes.append({
            "product_id": product_id,
            "old_quantity": old_quantity,
            "new_quantity": new_quantity,
            "difference": new_quantity - old_quantity
        })

    return changes


def find_inactive_changed_products(conn, changes):
    """Возвращает неактивные товары, затронутые изменением поступления."""
    changed_product_ids = [
        change["product_id"] for change in changes
        if not (change["old_quantity"] > change["new_quantity"] and change["new_quantity"] == 0)
    ]

    if not changed_product_ids:
        return set()

    placeholders = ", ".join("?" for _ in changed_product_ids)

    rows = conn.execute(
        f"""
        SELECT id
        FROM products
        WHERE id IN ({placeholders}) AND active != 1;
        """,
        changed_product_ids
    ).fetchall()

    return {row["id"] for row in rows}


def insert_new_supply_item(conn, supply_batch_id, change):
    """Добавляет новую позицию в существующую партию поступления."""
    cursor = conn.execute(
        """
        INSERT INTO supply_items (
            supply_batch_id,
            product_id,
            quantity,
            price_at_moment
        )
        SELECT ?, id, ?, price
        FROM products
        WHERE id = ?;
        """,
        (supply_batch_id, change["new_quantity"], change["product_id"])
    )

    if cursor.rowcount != 1:
        raise DatabaseSystemError(f"Не удалось добавить товар с ID {change['product_id']} в поступление #{supply_batch_id}")


def update_existing_supply_item(conn, supply_batch_id, change):
    """Изменяет количество существующей позиции поступления."""
    cursor = conn.execute(
        """
        UPDATE supply_items
        SET quantity = ?
        WHERE supply_batch_id = ? AND product_id = ?;
        """,
        (change["new_quantity"], supply_batch_id, change["product_id"])
    )

    if cursor.rowcount != 1:
        raise DatabaseSystemError(f"Не удалось изменить товар с ID {change['product_id']} в поступлении #{supply_batch_id}")


def update_supply_items(conn, supply_batch_id, changes):
    """Добавляет или обновляет позиции партии по списку изменений."""
    for change in changes:
        existing_item = conn.execute(
            """
            SELECT id
            FROM supply_items
            WHERE supply_batch_id = ? AND product_id = ?;
            """,
            (supply_batch_id, change["product_id"])
        ).fetchone()

        if existing_item is None:
            insert_new_supply_item(conn, supply_batch_id, change)
        else:
            update_existing_supply_item(conn, supply_batch_id, change)


def log_product_restore(conn, product_id):
    """Записывает в журнал восстановление активности товара."""
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
        (date.today().isoformat(), product_id, "active", "0", "1")
    )


def update_products_after_supply_change(conn, changes, inactive_product_ids):
    """Восстанавливает товары и корректирует их остатки после изменения поставки."""
    for product_id in inactive_product_ids:
        cursor = conn.execute(
            """
            UPDATE products
            SET active = 1
            WHERE id = ?;
            """,
            (product_id,)
        )

        if cursor.rowcount != 1:
            raise DatabaseSystemError(f"Не удалось восстановить товар с ID {product_id}")

        log_product_restore(conn, product_id)

    for change in changes:
        cursor = conn.execute(
            """
            UPDATE products
            SET current_stock = current_stock + ?
            WHERE id = ?;
            """,
            (change["difference"], change["product_id"])
        )

        if cursor.rowcount != 1:
            raise DatabaseSystemError(f"Не удалось обновить наличие товара с ID {change['product_id']}")


def insert_supply_change_movements(conn, season_id, supply, changes):
    """Записывает движения товаров по изменению поступления."""
    for change in changes:
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
                season_id,
                change["product_id"],
                date.today().isoformat(),
                "изменение_поступления",
                change["difference"],
                "supply_batch",
                supply["supply_batch_id"]
            )
        )


def process_one_existing_supply(conn, season_id, supply):
    """Обрабатывает изменения одной существующей поставки.

    Returns:
        Данные обработанного изменения или ``None``.
    """
    changes = find_supply_item_changes(conn, supply)

    if not changes:
        return None

    inactive_product_ids = find_inactive_changed_products(conn, changes)

    update_supply_items(conn, supply["supply_batch_id"], changes)
    update_products_after_supply_change(conn, changes, inactive_product_ids)
    insert_supply_change_movements(conn, season_id, supply, changes)

    return {
        "supply_batch_id": supply["supply_batch_id"],
        "changes": changes,
        "restored_product_ids": list(inactive_product_ids)
    }


def process_existing_supplies(conn, season_id, supplies):
    """Обрабатывает изменения всех существующих поступлений.

    Returns:
        Список обработанных поступлений.
    """
    processed_supplies = []

    for supply in find_existing_supplies(supplies):
        result = process_one_existing_supply(conn, season_id, supply)

        if result is not None:
            processed_supplies.append(result)

    return processed_supplies
