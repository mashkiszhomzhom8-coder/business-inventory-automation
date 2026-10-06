from inventory_app.core.exceptions import DatabaseSystemError
from datetime import date

def get_db_remains(conn, excel_products, season_id):
    """Получает последние зафиксированные остатки товаров сезона.

    Returns:
        Словарь остатков по ID товара.
    """
    product_ids = [product["id"] for product in excel_products]
    if not product_ids:
        return {}

    placeholders = ", ".join("?" for _ in product_ids)

    rows = conn.execute(
        f"""
        SELECT inventory_items.product_id, inventory_items.actual_remain
        FROM inventory_items
        JOIN inventory_check
            ON inventory_check.id = inventory_items.inventory_check_id
        WHERE inventory_check.season_id = ?
          AND inventory_items.product_id IN ({placeholders})
        ORDER BY inventory_check.id;
        """,
        (season_id, *product_ids)
    ).fetchall()

    db_remains = {}

    for row in rows:
        db_remains[row["product_id"]] = row["actual_remain"]

    missing_ids = set(product_ids) - set(db_remains)
    if missing_ids:
        raise DatabaseSystemError(f"Для товаров с ID {sorted(missing_ids)} отсутствует остаток текущего сезона")

    return db_remains

def find_remain_changes(conn, excel_products, season_id):
    """Сравнивает остатки Excel с остатками в БД.

    Returns:
        Список изменений остатков.
    """
    db_remains = get_db_remains(conn, excel_products, season_id)
    changes = []

    for excel_product in excel_products:
        product_id = excel_product["id"]
        expected_remain = db_remains[product_id]
        actual_remain = excel_product["remain"]

        if actual_remain == expected_remain:
            continue

        changes.append({
            "product_id": product_id,
            "expected_remain": expected_remain,
            "actual_remain": actual_remain,
            "difference": actual_remain - expected_remain
        })

    return changes

def create_inventory_check(conn, season_id):
    """Создаёт запись инвентаризации и возвращает её ID."""
    cursor = conn.execute(
        """
        INSERT INTO inventory_check (
            season_id,
            date
        )
        VALUES (?, ?);
        """,
        (season_id, date.today().isoformat())
    )

    return cursor.lastrowid

def insert_inventory_items(conn, inventory_check_id, changes):
    """Добавляет позиции изменений в инвентаризацию."""
    for change in changes:
        conn.execute(
            """
            INSERT INTO inventory_items (
                inventory_check_id,
                product_id,
                expected_remain,
                actual_remain,
                difference
            )
            VALUES (?, ?, ?, ?, ?);
            """,
            (
                inventory_check_id,
                change["product_id"],
                change["expected_remain"],
                change["actual_remain"],
                change["difference"]
            )
        )

def update_prod_curr_st(conn, changes):
    """Изменяет текущий складской остаток товаров."""
    for change in changes:
        cursor = conn.execute(
            """
            UPDATE products
            SET current_stock = current_stock + ?
            WHERE id = ?;
            """,
            (change["difference"], change["product_id"])
        )

        if cursor.rowcount == 0:
            raise DatabaseSystemError(f"Не удалось изменить наличие товара с ID {change['product_id']}: товар отсутствует в БД")

def insert_prod_movements(conn, changes, season_id, inventory_check_id):
    """Записывает движения товаров по инвентаризации."""
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
                "инвентаризация",
                change["difference"],
                "inventory_check",
                inventory_check_id
            )
        )

def update_product_remains(conn, excel_products, season_id):
    """Обрабатывает изменения остатков товаров из Excel.

    Returns:
        Список изменений или пустой список.
    """
    changes = find_remain_changes(conn, excel_products, season_id)

    if not changes:
        return []

    inventory_check_id = create_inventory_check(conn, season_id)
    insert_inventory_items(conn, inventory_check_id, changes)
    update_prod_curr_st(conn, changes)
    insert_prod_movements(conn, changes, season_id, inventory_check_id)

    return changes
