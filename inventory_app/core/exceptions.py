class AppError(Exception):
    """Базовое исключение приложения."""


class UserError(AppError):
    """Ошибка данных или действий владельца."""


class ExcelDataError(UserError):
    """Некорректные данные или структура Excel-файла."""


class OrderDataError(UserError):
    """Некорректно заполненный заказ."""


class InactiveProductInOrderError(OrderDataError):
    """В заказе указан товар с products.active = 0."""


class SupplyHeaderError(ExcelDataError):
    """Некорректно заполнена шапка поступления."""


class WorksheetNameError(ExcelDataError):
    """Отсутствует обязательный лист или изменено его название."""


class ProductDeletionInputError(UserError):
    """Некорректно указаны товары для деактивации."""


class ProductIdNotFoundError(ProductDeletionInputError):
    """Указанный ID товара отсутствует в базе данных."""


class ProductAlreadyDeletedError(ProductDeletionInputError):
    """Указанный товар уже неактивен."""


class DeletedActiveError(ProductDeletionInputError):
    """Активный товар отсутствует в прайс-листе"""


class SystemError(AppError):
    """Внутренняя ошибка системы."""


class DatabaseSystemError(SystemError):
    """Внутренняя ошибка при работе с базой данных."""


class PriceProductMissingFromDatabaseError(SystemError):
    """Товар присутствует в прайсе, но отсутствует в базе данных."""


class ProductDeletionSystemError(SystemError):
    """Системная ошибка деактивации товара."""


class ExcelSynchronizationError(SystemError):
    """Системная ошибка обновления или синхронизации Excel."""


class ProductRestoreSynchronizationError(ExcelSynchronizationError):
    """После восстановления активный товар не был добавлен в прайс."""

class ProductRestorationInputError(UserError):
    """Некорректный ввод товаров к восстановлению"""
    pass


class ProductAlreadyActiveError(UserError):
    """Пользователь ввел активные id для восстановления"""
    pass

class ProductRestorationSystemError(SystemError):
    """Ошибка восстановления указанного товара"""
    pass

class OrderDeletionInputError(UserError):
    """Некорректно указаны заказы для деактивации."""


class OrderIdNotFoundError(OrderDeletionInputError):
    """Указанный ID заказа отсутствует в базе данных."""


class OrderAlreadyDeletedError(OrderDeletionInputError):
    """Указанный заказ уже неактивен."""


class OrderDeletionSystemError(SystemError):
    """Системная ошибка деактивации заказа."""

class OrderRestorationInputError(UserError):
    """Некорректный ввод заказов для восстановления."""
    pass


class OrderAlreadyActiveError(OrderRestorationInputError):
    """Указанный заказ уже активен."""
    pass


class OrderRestorationSystemError(SystemError):
    """Системная ошибка восстановления заказа."""
    pass

class SeasonValidationError(AppError):
    """
    Ошибка, из-за которой текущий сезон нельзя закрыть.
    Сообщение этого исключения можно показывать владельцу.
    """
    pass

class NoActiveSeasonError(SeasonValidationError):
    """Нет активного сезона в БД"""
    pass

class ExtraActiveSeasons(SeasonValidationError):
    """Найдено несколько активных сезонов"""
    pass