import os
from inventory_app.core.exceptions import SeasonValidationError, WorksheetNameError, ExcelDataError, NoActiveSeasonError, ExtraActiveSeasons, ProductIdNotFoundError, DeletedActiveError

PRICE_SHEET_NAME = "Шаблон прайса"
ADD_PRODUCTS_SHEET_NAME = "Добавление товаров"
SERVICE_SHEET_NAME = "_service"

HEADER_ROW = 1
FIRST_PRODUCT_ROW = 2


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def normalize_path(path):
    """
    Приводит путь к нормальному виду для сравнения.
    """

    if path is None:
        return None

    path = os.path.abspath(os.path.expanduser(str(path)))
    path = os.path.normpath(path)
    path = os.path.normcase(path)

    return path


def normalize_text(value):
    """
    Приводит текстовое значение Excel / БД
    к удобному виду для сравнения.
    """

    if value is None:
        return ""

    return str(value).strip()


def normalize_size(value):
    """
    Пустой размер считаем пустой строкой.
    """

    if value is None:
        return ""

    return str(value).strip()


def normalize_int(value, field_name):
    if value is None or value == "":
        return 0

    try:
        number = float(value)
    except (ValueError, TypeError):
        raise ExcelDataError(
            f"Некорректное числовое значение "
            f"в поле «{field_name}»: {value}"
        )

    if not number.is_integer():
        raise ExcelDataError(
            f"Поле «{field_name}» должно содержать целое число: {value}"
        )

    return int(number)


def get_sheet(book, sheet_name):
    """
    Получает лист Excel по имени.
    """

    try:
        return book.sheets[sheet_name]
    except Exception:
        raise WorksheetNameError(f"В Excel отсутствует обязательный лист «{sheet_name}».")


def get_last_product_row(sheet):
    """
    Определяет последнюю строку товара по столбцу A (ID).
    """

    last_row = sheet.range(f"A{sheet.cells.last_cell.row}").end("up").row

    if last_row < FIRST_PRODUCT_ROW:
        raise ExcelDataError("В прайсе не найдено ни одного товара.")

    return last_row


def get_last_used_column(sheet):
    """
    Определяет последний используемый столбец по первой строке.
    """

    return sheet.range((HEADER_ROW, sheet.cells.last_cell.column)).end("left").column


# ============================================================
# АКТИВНЫЙ СЕЗОН
# ============================================================

def get_active_season(conn):
    """
    Получает активный сезон.

    В БД должен существовать ровно один сезон status = 'active'.
    """

    cursor = conn.execute(
        """
        SELECT
            id,
            start_date,
            end_date,
            excel_file,
            status
        FROM seasons
        WHERE status = 'active'
        """
    )

    rows = cursor.fetchall()

    if len(rows) == 0:
        raise NoActiveSeasonError(
            "Закрытие сезона невозможно.\n\n"
            "В базе данных отсутствует активный сезон."
        )

    if len(rows) > 1:
        raise ExtraActiveSeasons(
            "Закрытие сезона невозможно.\n\n"
            "В базе данных найдено несколько активных сезонов."
        )

    row = rows[0]

    return {
        "id": row[0],
        "start_date": row[1],
        "end_date": row[2],
        "excel_file": row[3],
        "status": row[4],
    }


# ============================================================
# ПРОВЕРКА ТЕКУЩЕГО EXCEL
# ============================================================

def validate_current_season_file(book, active_season):
    """
    Проверяет, что закрытие запускается именно из Excel-файла
    текущего активного сезона.
    """

    current_excel = normalize_path(book.fullname)
    active_excel = normalize_path(active_season["excel_file"])

    if not current_excel:
        raise SeasonValidationError(
            "Текущая Excel-книга не сохранена на диске."
        )

    if not active_excel:
        raise SeasonValidationError(
            "У активного сезона в БД не указан путь к Excel-файлу."
        )

    if current_excel != active_excel:
        raise SeasonValidationError(
            "Закрытие сезона невозможно.\n\n"
            "Текущий Excel не является рабочим файлом активного сезона.\n\n"
            f"Открытый файл:\n{current_excel}\n\n"
            f"Файл активного сезона:\n{active_excel}"
        )


# ============================================================
# ПРОВЕРКА СТРУКТУРЫ EXCEL
# ============================================================

def validate_excel_structure(book):
    """
    Проверяет наличие основных листов и обязательных заголовков.
    """

    price_sheet = get_sheet(book, PRICE_SHEET_NAME)

    get_sheet(book, ADD_PRODUCTS_SHEET_NAME)
    service_sheet = get_sheet(book, SERVICE_SHEET_NAME)

    expected_headers = {
        1: ["ID"],
        2: ["Описание", 'Название', "Название товара"],
        3: ["Опт. цена", 'Цена'],
        4: ["Фото"],
        5: ["Размер", "Размеры"],
        6: ["Поступление исход"],
        7: ["Остатки исход", "Остатки", "Остаток"],
        8: ["Наличие"],
    }

    for column, allowed_headers in expected_headers.items():

        value = price_sheet.range((HEADER_ROW, column)).value

        value = normalize_text(value)

        if value not in allowed_headers:
            expected = " / ".join(allowed_headers)

            raise ExcelDataError(
                "Закрытие сезона невозможно.\n\n"
                "Нарушена структура основного прайса.\n\n"
                f"Столбец {column}: "
                f"ожидался заголовок «{expected}», "
                f"получено «{value}»."
            )

    # _service должен оставаться скрытым
    try:
        if service_sheet.visible:
            raise SeasonValidationError(
                "Закрытие сезона невозможно.\n\n"
                "Служебный лист «_service» не скрыт."
            )
    except AttributeError:
        # Если конкретная версия xlwings не поддерживает .visible,
        # само наличие листа всё равно уже проверено.
        pass


# ============================================================
# ЧТЕНИЕ ТОВАРОВ ИЗ EXCEL
# ============================================================

def read_excel_products(book):
    """
    Читает товары из основного прайса.

    Учитывает объединённые группы размеров:
    название и цена могут быть заполнены только
    в первой строке группы.

    Возвращает:

    {
        product_id: {
            "row": 2,
            "name": "...",
            "price": Decimal(...),
            "size": "...",
            "remain": 5,
            "current_stock": 8
        }
    }
    """

    sheet = get_sheet(book, PRICE_SHEET_NAME)

    last_row = get_last_product_row(sheet)

    products = {}

    last_name = None
    last_price = None

    for row in range(FIRST_PRODUCT_ROW, last_row + 1):

        raw_id = sheet.range((row, 1)).value

        if raw_id is None or raw_id == "":
            continue

        product_id = normalize_int(
            raw_id,
            f"ID, строка {row}"
        )

        if product_id <= 0:
            raise ExcelDataError(
                f"Некорректный ID товара в строке {row}: {product_id}"
            )

        if product_id in products:
            raise ExcelDataError(
                "Закрытие сезона невозможно.\n\n"
                f"ID {product_id} встречается в прайсе несколько раз."
            )

        raw_name = sheet.range((row, 2)).value

        if raw_name is not None and normalize_text(raw_name) != "":
            last_name = normalize_text(raw_name)

        if not last_name:
            raise ExcelDataError(
                f"Не удалось определить название товара "
                f"для строки {row}, ID {product_id}."
            )

        raw_price = sheet.range((row, 3)).value

        if raw_price is not None and raw_price != "":
            last_price = normalize_int(
                raw_price,
                f"Опт. цена, строка {row}"
            )

        if last_price is None:
            raise ExcelDataError(
                f"Не удалось определить цену товара "
                f"для строки {row}, ID {product_id}."
            )

        size = normalize_size(
            sheet.range((row, 5)).value
        )

        remain = normalize_int(
            sheet.range((row, 7)).value,
            f"Остатки, строка {row}"
        )

        current_stock = normalize_int(
            sheet.range((row, 8)).value,
            f"Наличие, строка {row}"
        )

        products[product_id] = {
            "row": row,
            "name": last_name,
            "price": last_price,
            "size": size,
            "remain": remain,
            "current_stock": current_stock,
        }

    return products


# ============================================================
# ТОВАРЫ EXCEL <-> DATABASE
# ============================================================

def get_database_products(conn):
    """
    Получает товары из БД.
    """

    cursor = conn.execute(
        """
        SELECT
            id,
            name,
            price,
            size,
            current_stock,
            active
        FROM products
        """
    )

    products = {}

    for row in cursor.fetchall():

        product_id = row[0]

        products[product_id] = {
            "id": product_id,
            "name": normalize_text(row[1]),
            "price": normalize_int(
                row[2],
                f"products.price, ID {product_id}"
            ),
            "size": normalize_size(row[3]),
            "current_stock": normalize_int(
                row[4],
                f"products.current_stock, ID {product_id}"
            ),
            "active": row[5],
        }

    return products


def validate_excel_products_against_database(conn, excel_products):
    """
    Проверяет:

    - активный товар БД существует в Excel;
    - неактивный товар не находится в основном прайсе;
    - товар из Excel существует в БД;
    - ID совпадает;
    - name совпадает;
    - price совпадает;
    - size совпадает;
    - Excel.Наличие == products.current_stock.
    """

    db_products = get_database_products(conn)

    # ------------------------------------------
    # Сначала проверяем товары из Excel
    # ------------------------------------------

    for product_id, excel_product in excel_products.items():

        if product_id not in db_products:
            raise ProductIdNotFoundError(
                "Закрытие сезона невозможно.\n\n"
                "В прайсе найден товар, которого нет в БД.\n\n"
                f"ID: {product_id}\n"
                f"Строка Excel: {excel_product['row']}"
            )

        db_product = db_products[product_id]

        if not db_product["active"]:
            raise ExcelDataError(
                "Закрытие сезона невозможно.\n\n"
                "В основном прайсе находится неактивный товар.\n\n"
                f"ID: {product_id}"
            )

        if excel_product["name"] != db_product["name"]:
            raise SeasonValidationError(
                "Закрытие сезона невозможно.\n\n"
                "Обнаружено несинхронизированное название товара.\n\n"
                f"ID: {product_id}\n"
                f"Excel: {excel_product['name']}\n"
                f"БД: {db_product['name']}"
            )

        if excel_product["size"] != db_product["size"]:
            raise SeasonValidationError(
                "Закрытие сезона невозможно.\n\n"
                "Обнаружено несинхронизированное значение размера.\n\n"
                f"ID: {product_id}\n"
                f"Excel: {excel_product['size']}\n"
                f"БД: {db_product['size']}"
            )

        if excel_product["price"] != db_product["price"]:
            raise SeasonValidationError(
                "Закрытие сезона невозможно.\n\n"
                "Обнаружена несинхронизированная цена.\n\n"
                f"ID: {product_id}\n"
                f"Excel: {excel_product['price']}\n"
                f"БД: {db_product['price']}"
            )

        if (
            excel_product["current_stock"]
            != db_product["current_stock"]
        ):
            raise SeasonValidationError(
                "Закрытие сезона невозможно.\n\n"
                "Обнаружено расхождение наличия.\n\n"
                f"ID: {product_id}\n"
                f"Наличие в Excel: "
                f"{excel_product['current_stock']}\n"
                f"Наличие в БД: "
                f"{db_product['current_stock']}\n\n"
                "Проведите синхронизацию "
                "и повторите закрытие сезона."
            )

    # ------------------------------------------
    # Теперь проверяем, что все активные товары
    # из БД присутствуют в Excel
    # ------------------------------------------

    for product_id, db_product in db_products.items():

        if not db_product["active"]:
            continue

        if product_id not in excel_products:
            raise DeletedActiveError(
                "Закрытие сезона невозможно.\n\n"
                "Активный товар существует в БД, "
                "но отсутствует в прайсе.\n\n"
                f"ID: {product_id}\n"
                f"Название: {db_product['name']}"
            )


# ============================================================
# ПРОВЕРКА ОСТАТКОВ / ИНВЕНТАРИЗАЦИИ
# ============================================================

def get_last_inventory_remains(conn, season_id):
    """
    Для каждого товара получает последнее actual_remain
    из inventory_items текущего сезона.
    """

    cursor = conn.execute(
        """
        SELECT
            ii.product_id,
            ii.actual_remain
        FROM inventory_items ii
        JOIN inventory_check ic
            ON ic.id = ii.inventory_check_id
        WHERE ic.season_id = ?
          AND ii.id = (
              SELECT MAX(ii2.id)
              FROM inventory_items ii2
              JOIN inventory_check ic2
                  ON ic2.id = ii2.inventory_check_id
              WHERE ii2.product_id = ii.product_id
                AND ic2.season_id = ?
          )
        """,
        (season_id, season_id)
    )

    remains = {}

    for row in cursor.fetchall():

        product_id = row[0]

        remains[product_id] = normalize_int(
            row[1],
            f"inventory_items.actual_remain, ID {product_id}"
        )

    return remains


def validate_inventory_remains(
    conn,
    season_id,
    excel_products
):
    """
    Проверяет, что столбец G «Остатки»
    соответствует последней проведённой инвентаризации.
    """

    db_remains = get_last_inventory_remains(
        conn,
        season_id
    )

    for product_id, product in excel_products.items():

        if product_id not in db_remains:
            raise SeasonValidationError(
                "Закрытие сезона невозможно.\n\n"
                "Для товара отсутствует запись "
                "inventory_items текущего сезона.\n\n"
                f"ID: {product_id}"
            )

        excel_remain = product["remain"]
        db_remain = db_remains[product_id]

        if excel_remain != db_remain:
            raise SeasonValidationError(
                "Закрытие сезона невозможно.\n\n"
                "Обнаружено несинхронизированное "
                "изменение остатков.\n\n"
                f"ID: {product_id}\n"
                f"Остаток в Excel: {excel_remain}\n"
                f"Последний остаток в БД: {db_remain}\n\n"
                "Проведите синхронизацию "
                "и повторите закрытие сезона."
            )


# ============================================================
# ПРОВЕРКА ЛИСТА «ДОБАВЛЕНИЕ ТОВАРОВ»
# ============================================================

def validate_add_products_sheet(book):
    """
    Проверяет, что на листе «Добавление товаров»
    нет незавершённых строк.

    Заголовки находятся в первой строке.
    Проверяем данные начиная со второй строки.
    """

    sheet = get_sheet(
        book,
        ADD_PRODUCTS_SHEET_NAME
    )

    used_range = sheet.used_range

    last_row = (
        used_range.row
        + used_range.rows.count
        - 1
    )

    last_column = (
        used_range.column
        + used_range.columns.count
        - 1
    )

    if last_row < 2:
        return

    for row in range(2, last_row + 1):

        for column in range(1, last_column + 1):

            value = sheet.range(
                (row, column)
            ).value

            if value is not None and str(value).strip() != "":
                raise SeasonValidationError(
                    "Закрытие сезона невозможно.\n\n"
                    "На листе «Добавление товаров» "
                    "обнаружены незавершённые данные.\n\n"
                    f"Строка: {row}\n\n"
                    "Сначала завершите добавление товаров "
                    "или очистите незавершённые данные."
                )


# ============================================================
# НЕОБРАБОТАННЫЕ ЛИСТЫ ЗАКАЗОВ
# ============================================================

def validate_no_order_sheets(book):
    """
    После успешной обработки заказа его отдельный лист
    не должен оставаться в книге.

    Поэтому любой лист, начинающийся с «Заказ»,
    считаем необработанным заказом.
    """

    order_sheets = []

    for sheet in book.sheets:

        sheet_name = normalize_text(sheet.name)

        if sheet_name.lower().startswith("заказ"):
            order_sheets.append(sheet_name)

    if order_sheets:
        names = "\n".join(
            f"• {name}"
            for name in order_sheets
        )

        raise SeasonValidationError(
            "Закрытие сезона невозможно.\n\n"
            "В Excel остались необработанные "
            "листы заказов:\n\n"
            f"{names}\n\n"
            "Проведите синхронизацию "
            "и повторите закрытие сезона."
        )


# ============================================================
# ПОСТУПЛЕНИЯ
# ============================================================

def get_excel_supply_columns(book):
    """
    Получает столбцы поступлений.

    F — первое поступление сезона.

    После H поступления определяются по заголовку,
    начинающемуся со слова «Поступление».
    """

    sheet = get_sheet(book, PRICE_SHEET_NAME)

    columns = []

    # F — базовый столбец первого поступления
    columns.append(6)

    last_column = get_last_used_column(sheet)

    for column in range(9, last_column + 1):

        header = normalize_text(
            sheet.range(
                (HEADER_ROW, column)
            ).value
        )

        if header.lower().startswith("поступление"):
            columns.append(column)

    return columns


def get_database_supply_batches(conn, season_id):
    """
    Получает все поступления текущего сезона
    вместе с товарами внутри каждого поступления.
    """

    cursor = conn.execute(
        """
        SELECT id, date
        FROM supply_batch
        WHERE season_id = ?
        ORDER BY id
        """,
        (season_id,)
    )

    batches = []

    for batch_row in cursor.fetchall():

        batch_id = batch_row[0]

        items_cursor = conn.execute(
            """
            SELECT
                product_id,
                quantity
            FROM supply_items
            WHERE supply_batch_id = ?
            """,
            (batch_id,)
        )

        items = {}

        for item_row in items_cursor.fetchall():

            items[item_row[0]] = normalize_int(
                item_row[1],
                (
                    "supply_items.quantity, "
                    f"supply_batch_id {batch_id}"
                )
            )

        batches.append(
            {
                "id": batch_id,
                "date": batch_row[1],
                "items": items,
            }
        )

    return batches


def validate_supplies(
    conn,
    book,
    season_id,
    excel_products
):
    """
    Проверяет поступления Excel ↔ БД.

    F соответствует первому supply_batch сезона.
    Последующие столбцы «Поступление...»
    соответствуют следующим supply_batch по id.
    """

    sheet = get_sheet(book, PRICE_SHEET_NAME)

    excel_columns = get_excel_supply_columns(book)

    db_batches = get_database_supply_batches(
        conn,
        season_id
    )

    # Если поступлений в БД ещё не было,
    # F должен быть пустым.
    if len(db_batches) == 0:

        for product in excel_products.values():

            row = product["row"]

            value = normalize_int(
                sheet.range((row, 6)).value,
                f"Поступление исход, строка {row}"
            )

            if value != 0:
                raise SeasonValidationError(
                    "Закрытие сезона невозможно.\n\n"
                    "В Excel найдено поступление, "
                    "которого нет в БД.\n\n"
                    f"Строка: {row}"
                )

        # Дополнительных колонок поступлений
        # также быть не должно.
        if len(excel_columns) > 1:
            raise SeasonValidationError(
                "Закрытие сезона невозможно.\n\n"
                "В Excel найдены дополнительные "
                "столбцы поступлений, которых нет в БД."
            )

        return

    if len(excel_columns) != len(db_batches):
        raise SeasonValidationError(
            "Закрытие сезона невозможно.\n\n"
            "Количество поступлений в Excel "
            "не совпадает с количеством поступлений в БД.\n\n"
            f"Excel: {len(excel_columns)}\n"
            f"БД: {len(db_batches)}"
        )

    for column, batch in zip(
        excel_columns,
        db_batches
    ):

        for product_id, product in excel_products.items():

            excel_quantity = normalize_int(
                sheet.range(
                    (product["row"], column)
                ).value,
                (
                    "Поступление, "
                    f"строка {product['row']}"
                )
            )

            db_quantity = batch["items"].get(
                product_id,
                0
            )

            if excel_quantity != db_quantity:
                raise SeasonValidationError(
                    "Закрытие сезона невозможно.\n\n"
                    "Обнаружено несинхронизированное "
                    "поступление.\n\n"
                    f"ID товара: {product_id}\n"
                    f"Поступление ID: {batch['id']}\n"
                    f"Excel: {excel_quantity}\n"
                    f"БД: {db_quantity}"
                )


# ============================================================
# ЗАКАЗЫ В ОСНОВНОМ ПРАЙСЕ
# ============================================================

def get_excel_order_columns(book):
    """
    Все сезонные столбцы после H,
    которые НЕ являются поступлениями,
    считаются столбцами заказов.

    Пустые столбцы игнорируются.
    """

    sheet = get_sheet(book, PRICE_SHEET_NAME)

    last_column = get_last_used_column(sheet)

    columns = []

    for column in range(9, last_column + 1):

        header = normalize_text(
            sheet.range(
                (HEADER_ROW, column)
            ).value
        )

        if header == "":
            continue

        if header.lower().startswith("поступление"):
            continue

        columns.append(column)

    return columns


def get_database_orders(conn, season_id):
    """
    Получает заказы текущего сезона и их позиции.
    """

    cursor = conn.execute(
        """
        SELECT
            id,
            client_fio,
            tg_user,
            date
        FROM orders
        WHERE season_id = ?
        ORDER BY id
        """,
        (season_id,)
    )

    orders = []

    for order_row in cursor.fetchall():

        order_id = order_row[0]

        items_cursor = conn.execute(
            """
            SELECT
                product_id,
                quantity
            FROM order_items
            WHERE order_id = ?
            """,
            (order_id,)
        )

        items = {}

        for item_row in items_cursor.fetchall():

            items[item_row[0]] = normalize_int(
                item_row[1],
                (
                    "order_items.quantity, "
                    f"order_id {order_id}"
                )
            )

        orders.append(
            {
                "id": order_id,
                "client_fio": order_row[1],
                "tg_user": order_row[2],
                "date": order_row[3],
                "items": items,
            }
        )

    return orders


def validate_orders_in_price(
    conn,
    book,
    season_id,
    excel_products
):
    """
    Проверяет столбцы уже проведённых заказов
    основного прайса против order_items.
    """

    sheet = get_sheet(book, PRICE_SHEET_NAME)

    excel_columns = get_excel_order_columns(book)

    db_orders = get_database_orders(
        conn,
        season_id
    )

    if len(excel_columns) != len(db_orders):
        raise SeasonValidationError(
            "Закрытие сезона невозможно.\n\n"
            "Количество проведённых заказов "
            "в Excel не совпадает с БД.\n\n"
            f"Excel: {len(excel_columns)}\n"
            f"БД: {len(db_orders)}"
        )

    for column, order in zip(
        excel_columns,
        db_orders
    ):

        for product_id, product in excel_products.items():

            excel_quantity = normalize_int(
                sheet.range(
                    (product["row"], column)
                ).value,
                (
                    "Заказ, "
                    f"строка {product['row']}"
                )
            )

            db_quantity = order["items"].get(
                product_id,
                0
            )

            if excel_quantity != db_quantity:
                raise SeasonValidationError(
                    "Закрытие сезона невозможно.\n\n"
                    "Обнаружено расхождение "
                    "в проведённом заказе.\n\n"
                    f"Заказ ID: {order['id']}\n"
                    f"ID товара: {product_id}\n"
                    f"Excel: {excel_quantity}\n"
                    f"БД: {db_quantity}"
                )


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ PRE-CHECK
# ============================================================

def validate_season_can_be_closed(conn, book):
    """
    Полная предварительная проверка сезона.

    ВАЖНО:
    функция ничего не изменяет ни в Excel,
    ни в базе данных.

    Если всё хорошо, возвращает:
    active_season, excel_products

    Если обнаружена проблема,
    выбрасывает SeasonValidationError.
    """

    # 1. Ровно один активный сезон
    active_season = get_active_season(conn)

    # 2. Закрытие запущено из правильного Excel
    validate_current_season_file(
        book,
        active_season
    )

    # 3. Структура Excel не повреждена
    validate_excel_structure(book)

    # 4. Не осталось отдельных листов заказов
    validate_no_order_sheets(book)

    # 5. Нет незавершённого добавления товаров
    validate_add_products_sheet(book)

    # 6. Читаем состояние основного прайса
    excel_products = read_excel_products(book)

    # 7. Проверяем товары и главный инвариант:
    # Excel.Наличие == products.current_stock
    validate_excel_products_against_database(
        conn,
        excel_products
    )

    season_id = active_season["id"]

    # 8. Проверяем Остатки
    validate_inventory_remains(
        conn,
        season_id,
        excel_products
    )

    # 9. Проверяем поступления
    validate_supplies(
        conn,
        book,
        season_id,
        excel_products
    )

    # 10. Проверяем уже проведённые заказы
    validate_orders_in_price(
        conn,
        book,
        season_id,
        excel_products
    )

    return active_season, excel_products