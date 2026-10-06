import os
import shutil
from datetime import date, datetime

from inventory_app.core.config import SEASONS_FOLDER


FILE_PREFIX = "Учет"
EXCEL_EXTENSION = ".xlsm"


def get_seasons_folder():
    """
    Возвращает путь к единой папке, в которой хранятся Excel-файлы всех сезонов.
    """

    folder = os.path.abspath(os.path.expanduser(str(SEASONS_FOLDER)))

    if not os.path.exists(folder):
        raise FileNotFoundError(f"Папка с файлами сезонов не найдена:\n{folder}")

    if not os.path.isdir(folder):
        raise NotADirectoryError(f"Путь к сезонам должен указывать на папку:\n{folder}")

    return folder


def format_date_for_filename(value):
    """
    Приводит дату к формату YYYY-MM-DD.

    Поддерживает:
    - datetime.date
    - datetime.datetime
    - строку YYYY-MM-DD
    - строку DD.MM.YYYY
    - строку, начинающуюся с YYYY-MM-DD
    """

    if isinstance(value, datetime):
        return value.date().isoformat()

    if isinstance(value, date):
        return value.isoformat()

    if value is None:
        raise ValueError("Дата сезона не указана.")

    value = str(value).strip()

    if len(value) < 10:
        raise ValueError(f"Некорректная дата сезона: {value}")

    date_part = value[:10]

    try:
        parsed_date = date.fromisoformat(date_part)
    except ValueError:
        try:
            parsed_date = datetime.strptime(date_part, "%d.%m.%Y").date()
        except ValueError:
            raise ValueError(f"Некорректная дата сезона: {value}")

    return parsed_date.isoformat()


def build_archive_file_path(start_date, close_date):
    """
    Формирует путь архивной копии завершённого сезона.

    Пример:
    Учет_2026-05-01_2026-08-27.xlsm
    """

    folder = get_seasons_folder()

    start_date = format_date_for_filename(start_date)
    close_date = format_date_for_filename(close_date)

    filename = f"{FILE_PREFIX}_{start_date}_{close_date}{EXCEL_EXTENSION}"

    return os.path.join(folder, filename)


def build_new_season_file_path(close_date):
    """
    Формирует путь нового рабочего Excel.

    Пример:
    Учет_2026-08-27_активный.xlsm
    """

    folder = get_seasons_folder()

    close_date = format_date_for_filename(close_date)

    filename = f"{FILE_PREFIX}_{close_date}_активный{EXCEL_EXTENSION}"

    return os.path.join(folder, filename)


def validate_source_excel(source_excel_path):
    """
    Проверяет исходный рабочий Excel перед копированием.
    """

    if not source_excel_path:
        raise ValueError("Не указан путь к текущему Excel-файлу.")

    source_excel_path = os.path.abspath(os.path.expanduser(str(source_excel_path)))

    if not os.path.exists(source_excel_path):
        raise FileNotFoundError(f"Текущий Excel-файл не найден:\n{source_excel_path}")

    if not os.path.isfile(source_excel_path):
        raise ValueError(f"Путь текущего Excel не является файлом:\n{source_excel_path}")

    extension = os.path.splitext(source_excel_path)[1].lower()

    if extension != EXCEL_EXTENSION:
        raise ValueError(f"Рабочий Excel должен иметь формат .xlsm.\nПолучен файл: {source_excel_path}")

    return source_excel_path


def validate_target_inside_seasons_folder(target_excel_path):
    """
    Проверяет, что создаваемый или удаляемый Excel находится именно
    в папке сезонов.
    """

    seasons_folder = os.path.normcase(os.path.normpath(get_seasons_folder()))
    target_folder = os.path.normcase(os.path.normpath(os.path.dirname(target_excel_path)))

    if target_folder != seasons_folder:
        raise ValueError(f"Excel-файл должен находиться в папке сезонов:\n{get_seasons_folder()}")


def copy_excel_file(source_excel_path, target_excel_path):
    """
    Полностью копирует .xlsm через shutil.copy2().

    Если целевой файл уже существует, он не перезаписывается.
    При ошибке частично созданная копия удаляется.
    """

    source_excel_path = validate_source_excel(source_excel_path)
    target_excel_path = os.path.abspath(os.path.expanduser(str(target_excel_path)))

    validate_target_inside_seasons_folder(target_excel_path)

    source_normalized = os.path.normcase(os.path.normpath(source_excel_path))
    target_normalized = os.path.normcase(os.path.normpath(target_excel_path))

    if source_normalized == target_normalized:
        raise ValueError("Исходный и целевой Excel-файл не могут иметь одинаковый путь.")

    if os.path.exists(target_excel_path):
        raise FileExistsError(f"Файл с таким именем уже существует:\n{target_excel_path}")

    try:
        shutil.copy2(source_excel_path, target_excel_path)

        if not os.path.isfile(target_excel_path):
            raise IOError("После копирования новый файл не был найден на диске.")

        source_size = os.path.getsize(source_excel_path)
        target_size = os.path.getsize(target_excel_path)

        if source_size != target_size:
            raise IOError("Размер созданной копии Excel не совпадает с исходным файлом.")

    except Exception:
        if os.path.exists(target_excel_path):
            try:
                os.remove(target_excel_path)
            except OSError:
                pass

        raise

    return target_excel_path


def create_archive_copy(source_excel_path, start_date, close_date):
    """
    Создаёт полную архивную копию завершённого сезона.

    Архив после создания больше не изменяется.
    """

    archive_path = build_archive_file_path(start_date, close_date)

    return copy_excel_file(source_excel_path, archive_path)


def create_new_season_copy(source_excel_path, close_date):
    """
    Создаёт полную копию текущего .xlsm для нового активного сезона.

    На этом этапе файл ещё не очищается и не изменяется.
    """

    new_excel_path = build_new_season_file_path(close_date)

    return copy_excel_file(source_excel_path, new_excel_path)


def delete_new_season_copy(new_excel_path):
    """
    Удаляет незавершённую копию Excel нового сезона при ошибке.

    Архив эта функция удалять не должна.
    """

    if not new_excel_path:
        return

    new_excel_path = os.path.abspath(os.path.expanduser(str(new_excel_path)))

    if not os.path.exists(new_excel_path):
        return

    validate_target_inside_seasons_folder(new_excel_path)

    if not os.path.isfile(new_excel_path):
        raise ValueError(f"Указанный путь не является файлом:\n{new_excel_path}")

    extension = os.path.splitext(new_excel_path)[1].lower()

    if extension != EXCEL_EXTENSION:
        raise ValueError(f"Отказ от удаления: файл не является .xlsm.\n{new_excel_path}")

    os.remove(new_excel_path)
