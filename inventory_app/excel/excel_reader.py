"""Читает данные из открытой Excel-книги через xlwings."""

from datetime import date, datetime
import xlwings as xw

from inventory_app.core.exceptions import ExcelDataError, OrderDataError, SupplyHeaderError, WorksheetNameError

price_sheet = "Шаблон прайса"

col_id = 1
col_name = 2
col_price = 3
col_size = 5
col_supply1 = 6
col_remain = 7
col_curr_st = 8
col_fst_ord = 21


def convert_to_date(value, error_cls=ExcelDataError, field="дата"):
    """Преобразует значение Excel в строку даты ISO.

    Args:
        value: Исходное значение даты.
        error_cls: Класс ошибки для некорректного значения.
        field: Название проверяемого поля.

    Returns:
        Дата в формате ``YYYY-MM-DD``.
    """
    if value is None or str(value).strip() == "":
        return date.today().isoformat()

    if isinstance(value, datetime):
        return value.date().isoformat()

    if isinstance(value, date):
        return value.isoformat()

    text = str(value).strip()

    full_formats = ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y")
    for date_format in full_formats:
        try:
            return datetime.strptime(text, date_format).date().isoformat()
        except ValueError:
            continue

    short_formats = ("%d.%m", "%d/%m", "%d-%m")
    for date_format in short_formats:
        try:
            if "." in text:
                text_with_year = f"{text}.{date.today().year}"
                full_format = f"{date_format}.%Y"
            elif "/" in text:
                text_with_year = f"{text}/{date.today().year}"
                full_format = f"{date_format}/%Y"
            else:
                text_with_year = f"{text}-{date.today().year}"
                full_format = f"{date_format}-%Y"

            return datetime.strptime(text_with_year, full_format).date().isoformat()
        except ValueError:
            continue

    raise error_cls(f"В поле «{field}» не удалось распознать дату: {value!r}")


def convert_to_int(value, field, error_cls=ExcelDataError):
    """Преобразует значение Excel в целое число.

    Args:
        value: Исходное значение.
        field: Название проверяемого поля.
        error_cls: Класс ошибки для некорректного значения.

    Returns:
        Целое число или ноль для пустой ячейки.
    """
    if value is None or str(value).strip() == "":
        return 0

    try:
        number = float(value)

        if not number.is_integer():
            raise ValueError

        return int(number)

    except (ValueError, TypeError) as exc:
        raise error_cls(f"В поле {field} ожидалось целое число, получено: {value!r}") from exc


def delete_spaces(value):
    """Возвращает строковое значение без начальных и конечных пробелов."""
    if value is None:
        return ""

    return str(value).strip()


def get_worksheet(workbook, sheet_name):
    """Возвращает лист книги по имени.

    Raises:
        WorksheetNameError: Если лист отсутствует.
    """
    sheet_names = [sheet.name for sheet in workbook.sheets]

    if sheet_name not in sheet_names:
        raise WorksheetNameError(f"Отсутствует обязательный лист «{sheet_name}»")

    return workbook.sheets[sheet_name]


def get_last_used_row(worksheet):
    """Возвращает номер последней используемой строки листа."""
    return worksheet.used_range.last_cell.row


def get_last_used_column(worksheet):
    """Возвращает номер последнего используемого столбца листа."""
    return worksheet.used_range.last_cell.column


def is_empty_row(row_values):
    """Проверяет, что все ячейки строки пусты."""
    return all(value is None or str(value).strip() == "" for value in row_values)


def products_reader(workbook):
    """Читает товары и остатки с листа прайса.

    Args:
        workbook: Открытая Excel-книга.

    Returns:
        Список словарей с данными товаров.
    """
    worksheet = get_worksheet(workbook, price_sheet)
    last_row = get_last_used_row(worksheet)

    if last_row < 2:
        return []

    rows = worksheet.range((2, col_id), (last_row, col_curr_st)).options(ndim=2).value

    all_products = []
    memory_name = None
    memory_price = None
    seen_ids = set()

    for row_num, row_values in enumerate(rows, start=2):
        raw_id = row_values[col_id - 1]
        if raw_id is None or str(raw_id).strip() == "":
            continue
        prod_id = convert_to_int(raw_id, f"Проблема с конвертацией ID, строка {row_num}")

        if prod_id in seen_ids:
            raise ExcelDataError(f"Строка {row_num}: ID товара {prod_id} повторяется в прайс-листе")
        seen_ids.add(prod_id)

        if prod_id <= 0:
            raise ExcelDataError(f"Строка {row_num}: ID товара должен быть больше нуля")


        raw_name = row_values[col_name - 1]
        if raw_name is not None and str(raw_name).strip() != "":
            memory_name = delete_spaces(raw_name)

        if not memory_name:
            raise ExcelDataError(f"Строка {row_num}: для товара с ID {prod_id} не указано название")

        raw_price = row_values[col_price - 1]
        if raw_price is not None and str(raw_price).strip() != "":
            memory_price = convert_to_int(raw_price, f"Проблема с конвертацией цены, строка {row_num}")

        if memory_price is None:
            raise ExcelDataError(f"Строка {row_num}: для товара с ID {prod_id} не указана цена")

        prod_size = delete_spaces(row_values[col_size - 1])
        prod_remain = convert_to_int(row_values[col_remain - 1], f"Остаток, строка {row_num}")
        prod_curr_stock = convert_to_int(row_values[col_curr_st - 1], f"Наличие, строка {row_num}")

        all_products.append(
            {
                "id": prod_id,
                "name": memory_name,
                "price": memory_price,
                "size": prod_size,
                "remain": prod_remain,
                "current_stock": prod_curr_stock
            }
        )

        if not all_products:
            raise ExcelDataError(f"Лист «{price_sheet}» не содержит товаров с заполненным ID")

    return all_products


def read_one_ord(workbook, sheet_name):
    """Читает один заказ с указанного листа Excel.

    Args:
        workbook: Открытая Excel-книга.
        sheet_name: Имя листа заказа.

    Returns:
        Словарь с данными заказа и его позициями.
    """
    worksheet = workbook.sheets[sheet_name]

    raw_client_name, raw_tg, raw_date = worksheet.range("A1:C1").value

    client_name = delete_spaces(raw_client_name)
    tg = delete_spaces(raw_tg)
    order_date = convert_to_date(raw_date, OrderDataError,f"дата заказа на листе «{sheet_name}»")

    if not client_name:
        raise OrderDataError(f"Лист «{sheet_name}»: не указано ФИО клиента")

    order = {
        "sheet_name": sheet_name,
        "client_fio": client_name,
        "tg_user": tg,
        "date": order_date
    }



    last_row = get_last_used_row(worksheet)
    items = []

    if last_row >= 2:
        rows = worksheet.range((2, 1), (last_row, 3)).options(ndim=2).value

        for row_num, row_values in enumerate(rows, start=2):
            raw_prod_name, raw_prod_size, raw_quant = row_values

            if is_empty_row(row_values):
                continue

            prod_name = delete_spaces(raw_prod_name)
            prod_size = delete_spaces(raw_prod_size)
            quant = convert_to_int(raw_quant,f"«количество», лист «{sheet_name}», строка {row_num}", OrderDataError)

            if not prod_name:
                raise OrderDataError(f"Лист «{sheet_name}», строка {row_num}: не указано название товара")
            if quant <= 0:
                raise OrderDataError(f"Лист «{sheet_name}», строка {row_num}: количество должно быть больше нуля")

            items.append(
                {
                    "prod_name": prod_name,
                    "size": prod_size,
                    "quantity": quant
                }
            )

    if not items:
        raise OrderDataError(f"Лист «{sheet_name}»: заказ не содержит товаров")

    order["items"] = items
    return order


def excel_order_reader(workbook):
    """Читает все листы Excel, имена которых начинаются с «Заказ».

    Returns:
        Список заказов.
    """
    actual_orders = []
    for sheet_name in workbook.sheet_names:
        if sheet_name.strip().lower().startswith("заказ"):
            actual_orders.append(read_one_ord(workbook, sheet_name))
    return actual_orders


def get_column_type(header):
    """Определяет тип столбца по его заголовку.

    Returns:
        ``supply``, ``order`` или ``None``.
    """
    normalized_header = delete_spaces(header).lower()

    if normalized_header.startswith("поступление"):
        return "supply"

    if normalized_header.startswith("заказ"):
        return "order"

    return None


def get_supply_header(header):
    """Разбирает заголовок столбца поступления.

    Returns:
        Словарь с заголовком, ID поступления и датой либо ``None``.
    """
    normalized_header = delete_spaces(header)

    if not normalized_header.lower().startswith("поступление"):
        return None

    if normalized_header.lower() == "поступление исход":
        return {
            "header": normalized_header,
            "supply_batch_id": None,
            "date": None
        }

    parts = [part.strip() for part in normalized_header.split("_")]

    if len(parts) == 2:
        return {
            "header": normalized_header,
            "supply_batch_id": None,
            "date": convert_to_date(parts[1], SupplyHeaderError, "дата в заголовке поступления")
        }

    if len(parts) == 3 and parts[1].startswith("#"):
        raw_id = parts[1].replace("#", "").strip()
        batch_id = convert_to_int(raw_id, "ID поступления в заголовке", SupplyHeaderError)

        return {
            "header": normalized_header,
            "supply_batch_id": batch_id,
            "date": convert_to_date(parts[2], SupplyHeaderError, "дата в заголовке поступления")
        }

    raise SupplyHeaderError(
        f"Лист «{price_sheet}»: неверный заголовок поступления {normalized_header!r}; "
        f"ожидается «Поступление исход», «Поступление_дата» или «Поступление_#ID_дата»"
    )


def read_one_supply(column_number, worksheet, rows):
    """Читает одно поступление из столбца прайс-листа.

    Args:
        column_number: Номер столбца поступления.
        worksheet: Лист прайса.
        rows: Значения используемого диапазона.

    Returns:
        Словарь с данными поступления.
    """
    header = rows[0][column_number - 1]

    if header is None or str(header).strip() == "":
        return {}

    one_supply = get_supply_header(header)
    one_supply_items = []

    for row_num, row_values in enumerate(rows[1:], start=2):
        raw_product_id = row_values[col_id - 1]

        if raw_product_id is None or str(raw_product_id).strip() == "":
            continue

        prod_id = convert_to_int(raw_product_id,f"Строка {row_num}, не удалось конвертировать ID")

        cell_value = row_values[column_number - 1]
        quantity = convert_to_int(cell_value,f"Не удалось определить поступление для товара с ID {prod_id}")
        if quantity == 0:
            continue
        if quantity < 0:
            raise ExcelDataError(f'Для товара с id {prod_id} указано отрицательное количество в поставке.')

        one_supply_items.append(
            {
                "product_id": prod_id,
                "quantity": quantity
            }
        )

    one_supply["items"] = one_supply_items
    return one_supply


def excel_read_supply(workbook):
    """Читает все поступления с листа прайса.

    Args:
        workbook: Открытая Excel-книга.

    Returns:
        Список поступлений.
    """
    worksheet = get_worksheet(workbook, price_sheet)

    last_row = get_last_used_row(worksheet)
    last_column = max(get_last_used_column(worksheet), col_fst_ord - 1)

    if last_row < 1:
        return []

    rows = worksheet.range((1, 1), (last_row, last_column)).options(ndim=2).value
    all_supplies = []

    start_supply = read_one_supply(col_supply1, worksheet, rows)
    if not start_supply:
        raise ExcelDataError(f"Лист «{price_sheet}»: отсутствует обязательный заголовок «Поступление исход» в ячейке F1")
    all_supplies.append(start_supply)

    for column_number in range(9, col_fst_ord):
        header = rows[0][column_number - 1]

        if get_column_type(header) != "supply":
            continue

        one_supply = read_one_supply(column_number, worksheet, rows)

        if one_supply:
            all_supplies.append(one_supply)

    return all_supplies


if __name__ == "__main__":
    workbook = xw.Book("Шаблон_прайса.xlsx")

    products = products_reader(workbook)
    print(f"Всего прочитано товаров: {len(products)}")
    for product in products:
        print(product)

    all_orders = excel_order_reader(workbook)
    print(f"Найдено заказов: {len(all_orders)}")
    for order in all_orders:
        print(order)

    supplies = excel_read_supply(workbook)
    print(f"Найдено поставок: {len(supplies)}")
    for supply in supplies:
        print(supply)
