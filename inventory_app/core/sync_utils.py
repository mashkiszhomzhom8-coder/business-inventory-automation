from shutil import copy2
from .exceptions import DatabaseSystemError, UserError
from .logging_config import setup_logging

logger = setup_logging()


def create_backup(excel_path, backup_path):
    """Создаёт резервную копию Excel-файла.

    Args:
        excel_path: Путь к исходному Excel-файлу.
        backup_path: Путь для резервной копии.
    """
    if not excel_path.exists():
        raise FileNotFoundError(f"Файл {excel_path} не найден")

    copy2(excel_path, backup_path)


def restore_excel_file(workbook, app, excel_path, backup_path):
    """Восстанавливает Excel-файл из копии и открывает его заново.

    Args:
        workbook: Открытая книга, которую нужно закрыть.
        app: Экземпляр Excel xlwings.
        excel_path: Путь к восстанавливаемому файлу.
        backup_path: Путь к резервной копии.

    Returns:
        Повторно открытая книга Excel.
    """
    workbook.close()
    copy2(backup_path, excel_path)
    restored_workbook = app.books.open(str(excel_path))
    backup_path.unlink()
    return restored_workbook


def get_active_season_id(conn):
    """Возвращает ID активного сезона.

    Args:
        conn: Активное соединение с SQLite.

    Raises:
        DatabaseSystemError: Если активный сезон отсутствует.
    """
    season = conn.execute(
        """
        SELECT id
        FROM seasons
        WHERE status = 'active';
        """
    ).fetchone()

    if season is None:
        raise DatabaseSystemError("В базе данных отсутствует активный сезон")

    return season["id"]

def log_sync_error(error):
    """Записывает ошибку синхронизации в журнал с подходящим уровнем.

    Args:
        error: Обрабатываемое исключение.
    """
    if isinstance(error, UserError):
        logger.warning("%s: %s", type(error).__name__, error)
    elif isinstance(error, SystemError):
        logger.exception("Системная ошибка ручной синхронизации")
    else:
        logger.exception("Непредусмотренная ошибка ручной синхронизации")
