# from datetime import date

# from inventory_app.core.database import connect_db
# conn = connect_db()

#     # 1. Находим активный сезон.
# active_seasons = conn.execute(
#     """
#     SELECT id
#     FROM orders
#     WHERE active = 0;
#     """
# ).fetchall()
# print(active_seasons)