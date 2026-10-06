import sqlite3
from pathlib import Path
from typing import Union

from .exceptions import DatabaseSystemError

BASE_DIR = Path(__file__).resolve().parents[2]
DEFAULT_DB_PATH = BASE_DIR / "inventory.db"

DbPath = Union[str, Path]

def connect_db(db_path: DbPath = DEFAULT_DB_PATH):
    """Открывает соединение с SQLite и включает внешние ключи.

    Args:
        db_path: Путь к файлу базы данных.

    Returns:
        Настроенное соединение SQLite.
    """
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def create_tables(conn: sqlite3.Connection) -> None:
    """Создаёт таблицы и индексы приложения, если они отсутствуют.

    Args:
        conn: Активное соединение с SQLite.
    """
    cursor = conn.cursor()

    cursor.executescript(
        """
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            size TEXT DEFAULT '',
            price INTEGER NOT NULL DEFAULT 0,
            current_stock INTEGER NOT NULL DEFAULT 0,
            active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0, 1))
        );
        
        CREATE TABLE IF NOT EXISTS prod_movement (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            season_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            type TEXT NOT NULL CHECK (type IN ('продажа', 'отмена_заказа', 'поступление', 'инвентаризация', 'изменение_поступления')),
            quantity INTEGER NOT NULL,
            source TEXT NOT NULL,
            source_id INTEGER,
            FOREIGN KEY (season_id) REFERENCES seasons(id),
            FOREIGN KEY (product_id) REFERENCES products(id)
        );

        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            season_id INTEGER NOT NULL,
            client_fio TEXT NOT NULL,
            tg_user TEXT,
            date TEXT NOT NULL,
            total_sum INTEGER NOT NULL DEFAULT 0,
            active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0, 1)),
            FOREIGN KEY (season_id) REFERENCES seasons(id)
        );

        CREATE TABLE IF NOT EXISTS order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL,
            quantity INTEGER NOT NULL,
            price_at_moment INTEGER NOT NULL DEFAULT 0,
            FOREIGN KEY (order_id) REFERENCES orders(id),
            FOREIGN KEY (product_id) REFERENCES products(id)
        );
        
        CREATE TABLE IF NOT EXISTS change_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            product_id INTEGER NOT NULL,
            field TEXT NOT NULL,
            old_value TEXT,
            new_value TEXT,
            FOREIGN KEY (product_id) REFERENCES products(id)
        );

        CREATE TABLE IF NOT EXISTS supply_batch (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            season_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            FOREIGN KEY (season_id) REFERENCES seasons(id)
        );

        CREATE TABLE IF NOT EXISTS supply_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supply_batch_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL,
            quantity INTEGER NOT NULL DEFAULT 0,
            price_at_moment INTEGER NOT NULL DEFAULT 0,
            FOREIGN KEY (supply_batch_id) REFERENCES supply_batch(id),
            FOREIGN KEY (product_id) REFERENCES products(id),
            UNIQUE (supply_batch_id, product_id)
        );

        CREATE TABLE IF NOT EXISTS inventory_check (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            season_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            FOREIGN KEY (season_id) REFERENCES seasons(id)
        );

        CREATE TABLE IF NOT EXISTS inventory_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            inventory_check_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL,
            expected_remain INTEGER NOT NULL DEFAULT 0,
            actual_remain INTEGER NOT NULL DEFAULT 0,
            difference INTEGER NOT NULL DEFAULT 0,
            FOREIGN KEY (inventory_check_id) REFERENCES inventory_check(id),
            FOREIGN KEY (product_id) REFERENCES products(id),
            UNIQUE (inventory_check_id, product_id)
        );

        CREATE TABLE IF NOT EXISTS seasons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            start_date TEXT NOT NULL,
            end_date TEXT,
            excel_file TEXT,
            status TEXT NOT NULL CHECK (status IN ('active', 'closed'))
        );
        
        CREATE TABLE IF NOT EXISTS season_opening_stock (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            season_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL,
            opening_stock INTEGER NOT NULL DEFAULT 0,
            FOREIGN KEY (season_id) REFERENCES seasons(id),
            FOREIGN KEY (product_id) REFERENCES products(id),
            UNIQUE (season_id, product_id)
        );

        CREATE UNIQUE INDEX IF NOT EXISTS idx_only_one_active_season
        ON seasons(status)
        WHERE status = 'active';

        CREATE UNIQUE INDEX IF NOT EXISTS idx_active_product_name_size
        ON products(name, size)
        WHERE active = 1;

        CREATE INDEX IF NOT EXISTS idx_products_name_size
        ON products(name, size);

        CREATE INDEX IF NOT EXISTS idx_prod_movement_product_season
        ON prod_movement(product_id, season_id);

        CREATE INDEX IF NOT EXISTS idx_order_items_order
        ON order_items(order_id);

        CREATE INDEX IF NOT EXISTS idx_supply_items_batch
        ON supply_items(supply_batch_id);

        CREATE INDEX IF NOT EXISTS idx_inventory_items_product
        ON inventory_items(product_id);
        """
    )

    conn.commit()


def make_seasons_end_date_nullable(conn):
    """Обновляет старую таблицу seasons, где end_date не мог быть NULL."""

    columns = conn.execute("PRAGMA table_info(seasons)").fetchall()

    end_date_column = next(
        (column for column in columns if column["name"] == "end_date"),
        None
    )

    if end_date_column is None:
        raise DatabaseSystemError("В таблице seasons отсутствует поле end_date")

    if not end_date_column["notnull"]:
        return

    conn.commit()
    conn.execute("PRAGMA foreign_keys = OFF")

    try:
        conn.execute("BEGIN")

        conn.execute(
            """
            CREATE TABLE seasons_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                start_date TEXT NOT NULL,
                end_date TEXT,
                excel_file TEXT,
                status TEXT NOT NULL CHECK (status IN ('active', 'closed'))
            )
            """
        )

        conn.execute(
            """
            INSERT INTO seasons_new (
                id,
                start_date,
                end_date,
                excel_file,
                status
            )
            SELECT
                id,
                start_date,
                end_date,
                excel_file,
                status
            FROM seasons
            """
        )

        conn.execute("DROP INDEX IF EXISTS idx_only_one_active_season")
        conn.execute("DROP TABLE seasons")
        conn.execute("ALTER TABLE seasons_new RENAME TO seasons")
        conn.execute(
            """
            CREATE UNIQUE INDEX idx_only_one_active_season
            ON seasons(status)
            WHERE status = 'active'
            """
        )

        conn.commit()

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.execute("PRAGMA foreign_keys = ON")

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
        WHERE status = 'active'
        """
    ).fetchone()

    if season is None:
        raise DatabaseSystemError("В базе данных нет активного сезона")

    return season["id"]


if __name__ == "__main__":
    connection = connect_db()
    create_tables(connection)
    connection.close()
    print(f"База данных создана: {DEFAULT_DB_PATH}")
