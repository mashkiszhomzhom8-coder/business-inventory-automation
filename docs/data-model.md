# Модель данных

Ниже отражена физическая SQLite-схема, используемая приложением. Поля приведены в сокращённом виде; источник истины — `inventory_app/core/database.py`.

```mermaid
erDiagram
    PRODUCTS {
        int id PK
        text name
        text size
        int price
        int current_stock
        int active
    }
    SEASONS {
        int id PK
        text start_date
        text end_date
        text excel_file
        text status
    }
    ORDERS {
        int id PK
        int season_id FK
        text client_fio
        text tg_user
        text date
        int total_sum
        int active
    }
    ORDER_ITEMS {
        int id PK
        int order_id FK
        int product_id FK
        int quantity
        int price_at_moment
    }
    SUPPLY_BATCH {
        int id PK
        int season_id FK
        text date
    }
    SUPPLY_ITEMS {
        int id PK
        int supply_batch_id FK
        int product_id FK
        int quantity
        int price_at_moment
    }
    INVENTORY_CHECK {
        int id PK
        int season_id FK
        text date
    }
    INVENTORY_ITEMS {
        int id PK
        int inventory_check_id FK
        int product_id FK
        int expected_remain
        int actual_remain
        int difference
    }
    PROD_MOVEMENT {
        int id PK
        int season_id FK
        int product_id FK
        text date
        text type
        int quantity
        text source
        int source_id
    }
    CHANGE_LOG {
        int id PK
        text date
        int product_id FK
        text field
        text old_value
        text new_value
    }
    SEASON_OPENING_STOCK {
        int id PK
        int season_id FK
        int product_id FK
        int opening_stock
    }

    SEASONS ||--o{ ORDERS : contains
    ORDERS ||--|{ ORDER_ITEMS : has
    PRODUCTS ||--o{ ORDER_ITEMS : ordered_as
    SEASONS ||--o{ SUPPLY_BATCH : contains
    SUPPLY_BATCH ||--|{ SUPPLY_ITEMS : has
    PRODUCTS ||--o{ SUPPLY_ITEMS : supplied_as
    SEASONS ||--o{ INVENTORY_CHECK : contains
    INVENTORY_CHECK ||--|{ INVENTORY_ITEMS : has
    PRODUCTS ||--o{ INVENTORY_ITEMS : counted_as
    SEASONS ||--o{ PROD_MOVEMENT : records
    PRODUCTS ||--o{ PROD_MOVEMENT : moves
    PRODUCTS ||--o{ CHANGE_LOG : changes
    SEASONS ||--o{ SEASON_OPENING_STOCK : starts_with
    PRODUCTS ||--o{ SEASON_OPENING_STOCK : opens_with
```

## Ключевые ограничения

- `season_opening_stock` уникален по паре `(season_id, product_id)`.
- `supply_items` уникальны по паре `(supply_batch_id, product_id)`.
- `inventory_items` уникальны по паре `(inventory_check_id, product_id)`.
- Частичный уникальный индекс не допускает более одного `seasons.status = 'active'`.
- Активные товары уникальны по сочетанию названия и размера.
