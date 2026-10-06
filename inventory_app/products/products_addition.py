import sys

import xlwings as xw

if sys.platform.startswith("darwin"):
    from appscript import k as kw

from inventory_app.core.logging_config import setup_logging
from inventory_app.core.database import connect_db
from datetime import date

from inventory_app.core.sync_utils import get_active_season_id

logger = setup_logging()


PRICE_SHEET_NAME = "Шаблон прайса"
ADD_SHEET_NAME = "Добавление товаров"

FIRST_PRODUCT_ROW = 2
FIRST_INPUT_ROW = 2
REFERENCE_PRODUCT_ID = 778

INPUT_HEADERS = [
    "Название",
    "Размер",
    "Цена",
    "Вставить после товара с ID",
]

LETTER_SIZE_ORDER = {
    "XXS": 1,
    "XS": 2,
    "S": 3,
    "M": 4,
    "L": 5,
    "XL": 6,
    "XXL": 7,
    "XXXL": 8,
}

FONT_NAME = "Times New Roman"
FONT_SIZE = 12
PRICE_FONT_COLOR = (176, 36, 24)

def initialize_new_products_in_current_season(conn, groups):
    """
    Создаёт для новых товаров нулевой остаток текущего сезона.
    """

    season_id = get_active_season_id(conn)

    new_product_ids = []

    for group in groups:
        new_product_ids.extend(group["new_ids"])

    if not new_product_ids:
        return

    # Для новых товаров фиксируем нулевой стартовый остаток сезона.
    for product_id in new_product_ids:
        conn.execute(
            """
            INSERT INTO season_opening_stock (
                season_id,
                product_id,
                opening_stock
            )
            VALUES (?, ?, ?);
            """,
            (
                season_id,
                product_id,
                0,
            ),
        )

    # Одна инвентаризация на всю пачку добавленных товаров.
    cursor = conn.execute(
        """
        INSERT INTO inventory_check (
            season_id,
            date
        )
        VALUES (?, ?);
        """,
        (
            season_id,
            date.today().isoformat(),
        ),
    )

    inventory_check_id = cursor.lastrowid

    for product_id in new_product_ids:
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
                product_id,
                0,
                0,
                0,
            ),
        )


def is_empty(value):
    return value is None or str(value).strip() == ""


def normalize_name(value, row_number):
    if is_empty(value):
        raise ValueError(f"Строка {row_number}: не указано название товара.")
    return str(value).strip()


def normalize_size(value):
    if is_empty(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    if isinstance(value, int):
        return str(value)

    return str(value).strip()


def normalize_price(value, row_number):
    if is_empty(value):
        raise ValueError(f"Строка {row_number}: не указана цена товара.")

    if isinstance(value, (int, float)):
        price = float(value)
    else:
        text = str(value).strip().replace(" ", "").replace(",", ".")
        try:
            price = float(text)
        except ValueError:
            raise ValueError(f"Строка {row_number}: цена должна быть числом.")

    if price < 0:
        raise ValueError( f"Строка {row_number}: цена не может быть отрицательной.")

    return price


def normalize_after_id(value, row_number):
    if is_empty(value):
        return None

    try:
        number = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"Строка {row_number}: в поле 'Вставить после товара с ID' должен быть целый ID.")

    if not number.is_integer() or number <= 0:
        raise ValueError(f"Строка {row_number}: ID должен быть положительным целым числом.")

    return int(number)


def size_sort_key(size):
    text = str(size).strip()

    if text == "":
        return (0, 0)

    try:
        number = float(text.replace(",", "."))
        return (1, number)
    except ValueError:
        pass

    upper_size = text.upper()

    if upper_size in LETTER_SIZE_ORDER:
        return (2, LETTER_SIZE_ORDER[upper_size])

    return (3, upper_size)


# def validate_input_headers(sheet):
#     actual_headers = []
#
#     for column in range(1, 5):
#         value = sheet.range((1, column)).value
#         actual_headers.append("" if value is None else str(value).strip())
#
#     if actual_headers != INPUT_HEADERS:
#         raise ValueError(
#             "На листе 'Добавление товаров' неверные заголовки.\n\n"
#             "Должно быть:\n"
#             "A1 — Название\n"
#             "B1 — Размер\n"
#             "C1 — Цена\n"
#             "D1 — Вставить после товара с ID"
#         )


def get_last_input_row(sheet):
    last_excel_row = sheet.cells.last_cell.row
    last_rows = []

    for column_letter in ("A", "B", "C", "D"):
        row = sheet.range(f"{column_letter}{last_excel_row}").end("up").row
        last_rows.append(row)

    return max(last_rows)


def read_input_rows(sheet):
    # validate_input_headers(sheet)

    last_row = get_last_input_row(sheet)

    if last_row < FIRST_INPUT_ROW:
        raise ValueError("На листе 'Добавление товаров' нет товаров для добавления.")

    products = []

    for row in range(FIRST_INPUT_ROW, last_row + 1):
        name_value = sheet.range(f"A{row}").value
        size_value = sheet.range(f"B{row}").value
        price_value = sheet.range(f"C{row}").value
        after_id_value = sheet.range(f"D{row}").value

        if (is_empty(name_value) and is_empty(size_value) and is_empty(price_value) and is_empty(after_id_value)):
            continue

        products.append(
            {
                "name": normalize_name(name_value, row),
                "size": normalize_size(size_value),
                "price": normalize_price(price_value, row),
                "after_id": normalize_after_id(after_id_value, row),
            }
        )

    if not products:
        raise ValueError("На листе 'Добавление товаров' нет товаров для добавления.")

    return products, last_row


def group_products(products):
    groups = []
    groups_by_key = {}

    for product in products:
        key = (product["name"], product["price"])

        if key not in groups_by_key:
            group = {
                "name": product["name"],
                "price": product["price"],
                "products": [],
                "after_ids": set(),
                "after_id": None,
                "anchor_id": None,
                "new_ids": [],
            }
            groups_by_key[key] = group
            groups.append(group)

        group = groups_by_key[key]
        group["products"].append(product)

        if product["after_id"] is not None:
            group["after_ids"].add(product["after_id"])

    for group in groups:
        if len(group["after_ids"]) > 1:
            ids = sorted(group["after_ids"])
            raise ValueError(
                f"Для группы '{group['name']}' указаны разные места вставки: "
                + ", ".join(str(item) for item in ids)
                + ".\nУ всей группы должно быть одно место вставки."
            )

        if group["after_ids"]:
            group["after_id"] = next(iter(group["after_ids"]))

        seen_sizes = set()

        for product in group["products"]:
            size = product["size"]

            if size in seen_sizes:
                printable_size = size if size != "" else "без размера"
                raise ValueError(f"В группе '{group['name']}' размер '{printable_size}' указан несколько раз.")

            seen_sizes.add(size)

        group["products"].sort(key=lambda product: size_sort_key(product["size"]))

    return groups


def validate_new_products(conn, groups):
    for group in groups:
        for product in group["products"]:
            row = conn.execute(
                """
                SELECT id, active
                FROM products
                WHERE name = ? AND size = ?
                """,
                (product["name"], product["size"]),
            ).fetchone()

            if row is None:
                continue

            product_id, active = row
            printable_size = (product["size"] if product["size"] != "" else "без размера")

            if active == 1:
                raise ValueError(f"Товар уже существует:\n {product_id}. {product['name']} {printable_size}")

            raise ValueError(f"Товар уже есть в БД, но он неактивен:\n"
                f"{product_id}. {product['name']} {printable_size}\n\n"
                f"Его нужно восстановить, а не создавать заново."
            )


def validate_target_ids(conn, groups):
    checked_ids = set()

    for group in groups:
        after_id = group["after_id"]

        if after_id is None or after_id in checked_ids:
            continue

        row = conn.execute(
            "SELECT active FROM products WHERE id = ?",
            (after_id,),
        ).fetchone()

        if row is None:
            raise ValueError(f"Товара с ID {after_id} не существует.")

        if row[0] != 1:
            raise ValueError(f"Товар с ID {after_id} неактивен и отсутствует в прайсе.\n После него нельзя вставить новую группу.")
        checked_ids.add(after_id)


def get_last_product_row(sheet):
    last_excel_row = sheet.cells.last_cell.row
    last_row = sheet.range(f"A{last_excel_row}").end("up").row

    if last_row < FIRST_PRODUCT_ROW:
        return FIRST_PRODUCT_ROW - 1

    return last_row


def find_product_row(sheet, product_id):
    last_row = get_last_product_row(sheet)

    for row in range(FIRST_PRODUCT_ROW, last_row + 1):
        value = sheet.range(f"A{row}").value

        if is_empty(value):
            continue

        try:
            excel_id = int(float(value))
        except (TypeError, ValueError):
            continue

        if excel_id == product_id:
            return row

    return None

def find_group_start_row(sheet, product_row):
    row = product_row

    while row > FIRST_PRODUCT_ROW:
        current_name = sheet.range(f"B{row}").value

        if not is_empty(current_name):
            break

        row -= 1

    return row

def find_group_end_row(sheet, product_row):
    last_row = get_last_product_row(sheet)
    row = product_row

    while row < last_row:
        next_id = sheet.range(f"A{row + 1}").value
        next_name = sheet.range(f"B{row + 1}").value

        if is_empty(next_id):
            break

        if not is_empty(next_name):
            break

        row += 1

    return row


def get_id_from_row(sheet, row):
    value = sheet.range(f"A{row}").value

    if is_empty(value):
        raise ValueError(f"В строке {row} прайса отсутствует ID товара.")

    try:
        return int(float(value))
    except (TypeError, ValueError):
        raise ValueError( f"В строке {row} прайса некорректный ID товара.")


def resolve_anchor_ids(sheet, groups):
    """
    Превращает ID, указанный владельцем, в ID последнего товара
    соответствующей уже существующей группы.

    Поэтому если владелец указал ID середины группы, новая группа
    всё равно будет вставлена после всей существующей группы.
    """
    for group in groups:
        after_id = group["after_id"]

        if after_id is None:
            group["anchor_id"] = None
            continue

        product_row = find_product_row(sheet, after_id)

        if product_row is None:
            raise ValueError(f"Товар с ID {after_id} активен в БД, но отсутствует на листе '{PRICE_SHEET_NAME}'.")

        end_row = find_group_end_row(sheet, product_row)
        group["anchor_id"] = get_id_from_row(sheet, end_row)


def insert_products_into_db(conn, groups):
    for group in groups:
        for product in group["products"]:
            cursor = conn.execute(
                """
                INSERT INTO products
                    (name, size, price, current_stock, active)
                VALUES
                    (?, ?, ?, 0, 1)
                """,
                (
                    product["name"],
                    product["size"],
                    product["price"],
                ),
            )

            group["new_ids"].append(cursor.lastrowid)


def get_insert_row(sheet, anchor_id):
    if anchor_id is None:
        return get_last_product_row(sheet) + 1

    product_row = find_product_row(sheet, anchor_id)

    if product_row is None:
        raise ValueError(f"Не удалось найти строку товара с ID {anchor_id} на листе '{PRICE_SHEET_NAME}'.")

    group_end_row = find_group_end_row(sheet, product_row)
    return group_end_row + 1


def insert_empty_rows(sheet, start_row, count):
    end_row = start_row + count - 1
    sheet.range(f"{start_row}:{end_row}").insert(shift="down")


def set_center_alignment(cell_range):
    if sys.platform.startswith("darwin"):
        cell_range.api.horizontal_alignment.set(kw.horizontal_align_center)
        cell_range.api.vertical_alignment.set(kw.vertical_alignment_center)
    else:
        cell_range.api.HorizontalAlignment = -4108
        cell_range.api.VerticalAlignment = -4108


def copy_border(source_range, target_range, border_name):
    if sys.platform.startswith("darwin"):
        border_names = {
            "left": kw.edge_left,
            "top": kw.edge_top,
            "bottom": kw.edge_bottom,
            "right": kw.edge_right,
            "inside_horizontal": kw.inside_horizontal,
        }

        source_border = source_range.api.get_border(
            which_border=border_names[border_name]
        )

        target_border = target_range.api.get_border(
            which_border=border_names[border_name]
        )

        target_border.line_style.set(source_border.line_style.get())
        target_border.weight.set(source_border.weight.get())

    else:
        border_numbers = {
            "left": 7,
            "top": 8,
            "bottom": 9,
            "right": 10,
            "inside_horizontal": 12,
        }

        border_number = border_numbers[border_name]

        source_border = source_range.api.Borders(border_number)
        target_border = target_range.api.Borders(border_number)

        target_border.LineStyle = source_border.LineStyle
        target_border.Weight = source_border.Weight

def format_group_borders_from_reference(sheet, start_row, end_row):
    reference_row = find_product_row(sheet, REFERENCE_PRODUCT_ID)

    if reference_row is None:
        raise ValueError(
            f"Не удалось найти эталонный товар с ID {REFERENCE_PRODUCT_ID} "
            f"для оформления границ."
        )

    reference_start_row = find_group_start_row(sheet, reference_row)
    reference_end_row = find_group_end_row(sheet, reference_row)

    # Верхняя граница группы.
    copy_border(
        sheet.range(f"A{reference_start_row}:F{reference_start_row}"),
        sheet.range(f"A{start_row}:F{start_row}"),
        "top",
    )

    # Нижняя граница строки над группой.
    # Вместе с верхней границей новой группы она создаёт
    # такую же визуальную линию, как у существующих товаров.
    if start_row > FIRST_PRODUCT_ROW and reference_start_row > FIRST_PRODUCT_ROW:
        copy_border(
            sheet.range(
                f"A{reference_start_row - 1}:F{reference_start_row - 1}"
            ),
            sheet.range(
                f"A{start_row - 1}:F{start_row - 1}"
            ),
            "bottom",
        )

    # Нижняя граница группы.
    copy_border(
        sheet.range(f"A{reference_end_row}:F{reference_end_row}"),
        sheet.range(f"A{end_row}:F{end_row}"),
        "bottom",
    )

    # Левая внешняя граница.
    # copy_border(
    #     sheet.range(f"A{reference_start_row}:A{reference_end_row}"),
    #     sheet.range(f"A{start_row}:A{end_row}"),
    #     "left",
    # )

    # Вертикальные границы между столбцами.
    # Берём правую границу каждого столбца эталона.
    for column in ("B", "C", "D", "E", "F"):
        copy_border(
            sheet.range(
                f"{column}{reference_start_row}:{column}{reference_end_row}"
            ),
            sheet.range(
                f"{column}{start_row}:{column}{end_row}"
            ),
            "right",
        )

    if end_row > start_row:
        # Внутренние горизонтальные линии у ID.
        copy_border(
            sheet.range(
                f"A{reference_start_row}:A{reference_end_row}"
            ),
            sheet.range(
                f"A{start_row}:A{end_row}"
            ),
            "inside_horizontal",
        )

        # Внутренние горизонтальные линии у размера.
        copy_border(
            sheet.range(
                f"E{reference_start_row}:E{reference_end_row}"
            ),
            sheet.range(
                f"E{start_row}:E{end_row}"
            ),
            "inside_horizontal",
        )

        # Внутренние горизонтальные линии у исходного поступления.
        copy_border(
            sheet.range(
                f"F{reference_start_row}:F{reference_end_row}"
            ),
            sheet.range(
                f"F{start_row}:F{end_row}"
            ),
            "inside_horizontal",
        )


def format_group_in_excel(sheet, start_row, end_row):
    group_range = sheet.range(f"A{start_row}:F{end_row}")
    name_range = sheet.range(f"B{start_row}:B{end_row}")
    price_range = sheet.range(f"C{start_row}:C{end_row}")

    # Общий шрифт для всех ячеек новой группы.
    group_range.font.name = FONT_NAME
    group_range.font.bold = False
    group_range.font.color = (0, 0, 0)

    sheet.range(f"A{start_row}:A{end_row}").font.size = 10
    sheet.range(f"E{start_row}:E{end_row}").font.size = 10

    # Название и цена — жирные.
    price_range.font.bold = True
    name_range.font.bold = True

    # ID и цена — бордовые.
    sheet.range(f"A{start_row}:A{end_row}").font.color = PRICE_FONT_COLOR
    price_range.font.color = PRICE_FONT_COLOR

    # Центрирование содержимого во всех ячейках группы.
    set_center_alignment(group_range)

    # Для длинного названия разрешаем перенос текста.
    name_range.wrap_text = True

    # Высоту новых строк берём у предыдущей товарной строки.
    if start_row > FIRST_PRODUCT_ROW:
        previous_row_height = sheet.range(f"{start_row - 1}:{start_row - 1}").row_height

        if previous_row_height is not None:
            sheet.range(f"{start_row}:{end_row}").row_height = previous_row_height

    # Borders
    format_group_borders_from_reference(sheet, start_row, end_row)



def fill_group_in_excel(sheet, start_row, group):
    count = len(group["products"])
    end_row = start_row + count - 1

    if count > 1:
        name_range = sheet.range(f"B{start_row}:B{end_row}")
        price_range = sheet.range(f"C{start_row}:C{end_row}")
        photo_range = sheet.range(f"D{start_row}:D{end_row}")

        name_range.unmerge()
        price_range.unmerge()
        photo_range.unmerge()

        name_range.merge()
        price_range.merge()
        photo_range.merge()

    sheet.range(f"B{start_row}").value = group["name"]
    sheet.range(f"C{start_row}").value = group["price"]

    for index, product in enumerate(group["products"]):
        row = start_row + index
        new_id = group["new_ids"][index]

        sheet.range(f"A{row}").value = new_id
        sheet.range(f"E{row}").value = product["size"]

        # D — фото: не трогаем.
        # F — поступление: новый товар пока не имеет поступления.
        sheet.range(f"F{row}").value = None

        # Новый товар создаётся с нулевым current_stock.
        sheet.range(f"G{row}").value = 0
        sheet.range(f"H{row}").value = 0

    format_group_in_excel(sheet, start_row, end_row)


def insert_groups_into_excel(sheet, groups):
    inserted_blocks = []

    # Ключ — ID конца исходной группы.
    # Значение — последний ID уже вставленной после неё новой группы.
    last_inserted_for_anchor = {}

    for group in groups:
        anchor_id = group["anchor_id"]

        if anchor_id in last_inserted_for_anchor:
            actual_anchor_id = last_inserted_for_anchor[anchor_id]
        else:
            actual_anchor_id = anchor_id

        start_row = get_insert_row(sheet, actual_anchor_id)
        count = len(group["products"])

        insert_empty_rows(sheet, start_row, count)

        inserted_blocks.append(
            {
                "start_row": start_row,
                "count": count,
            }
        )

        fill_group_in_excel(sheet, start_row, group)

        if anchor_id is not None:
            last_inserted_for_anchor[anchor_id] = group["new_ids"][-1]

    return inserted_blocks


def remove_inserted_blocks(sheet, inserted_blocks):
    for block in reversed(inserted_blocks):
        start_row = block["start_row"]
        end_row = start_row + block["count"] - 1
        sheet.range(f"{start_row}:{end_row}").delete(shift="up")


def add_products():
    book = xw.Book.caller()
    sheet_names = [sheet.name for sheet in book.sheets]

    if ADD_SHEET_NAME not in sheet_names:
        raise ValueError(f"В книге нет листа '{ADD_SHEET_NAME}'.")

    if PRICE_SHEET_NAME not in sheet_names:
        raise ValueError(f"В книге нет листа '{PRICE_SHEET_NAME}'.")

    input_sheet = book.sheets[ADD_SHEET_NAME]
    price_sheet = book.sheets[PRICE_SHEET_NAME]

    products, last_input_row = read_input_rows(input_sheet)
    groups = group_products(products)

    input_snapshot = input_sheet.range(
        f"A{FIRST_INPUT_ROW}:D{last_input_row}"
    ).options(ndim=2).value

    conn = connect_db()
    inserted_blocks = []

    try:
        # Сначала только проверки. БД и Excel ещё не меняем.
        validate_new_products(conn, groups)
        validate_target_ids(conn, groups)
        resolve_anchor_ids(price_sheet, groups)

        conn.execute("BEGIN")

        # Создаём записи products и получаем новые ID.
        insert_products_into_db(conn, groups)

        # Создаём для новых товаров нулевой остаток текущего сезона.
        initialize_new_products_in_current_season(conn, groups)

        # Вставляем строки в Excel.
        inserted_blocks = insert_groups_into_excel(price_sheet, groups)

        # После успешной вставки очищаем лист ввода.
        input_sheet.range(
            f"A{FIRST_INPUT_ROW}:D{last_input_row}"
        ).clear_contents()

        # Сначала сохраняем Excel.
        # Пока сохранение не прошло успешно, БД не коммитим.
        book.save()

        conn.commit()

    except Exception:
        conn.rollback()

        # Возвращаем Excel в исходное состояние.
        if inserted_blocks:
            try:
                remove_inserted_blocks(price_sheet, inserted_blocks)
            except Exception:
                pass

        try:
            input_sheet.range(
                f"A{FIRST_INPUT_ROW}:D{last_input_row}"
            ).value = input_snapshot
        except Exception:
            pass

        try:
            book.save()
        except Exception:
            pass

        raise

    finally:
        conn.close()

    total_products = sum(len(group["products"])for group in groups)

    book.app.alert(
        f"Готово.\n"
        f"Добавлено товаров: {total_products}\n"
        f"Добавлено групп: {len(groups)}",
        "Добавление товаров")


def main():
    book = xw.Book.caller()

    try:
        add_products()

    except Exception as error:
        logger.exception("Ошибка при добавлении товаров")

        book.app.alert(
            str(error),
            "Ошибка добавления товаров",
        )