"""
Конфигурация проектов и доступов
"""
import os

BOT_TOKEN = os.getenv("BOT_TOKEN", "")
OWNER_ID = int(os.getenv("OWNER_ID", "1296668664"))
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

# Проекты системы
PROJECTS = {
    "filter_kg": {
        "name": "Filter.kg",
        "emoji": "💧",
        "amo_domain": "filter1kg.amocrm.ru",
        "amo_token": os.getenv("AMO_TOKEN_FILTER", ""),
        "pipeline_id": 10581262,
        "status_map": {
            83450574: "Неразобранное",
            83450786: "Неразобранное",
            83450578: "Физ лица",
            83450582: "Первый контакт",
            83450586: "Вышли на ЛПР",
            83450590: "Назначено демо",
            83450790: "Демо проведено",
            83832670: "Партнёры",
            83450798: "Деньги в кассе",
            83450802: "Принимают решение",
            83450806: "Мотивированный отказ",
            142: "Успешно реализовано",
            143: "Закрыто / не реализовано",
        },
        "won_statuses": [83450798, 142],
        "lost_statuses": [83450806, 143],
    },
    "a_farm": {
        "name": "А-Фарм",
        "emoji": "🏥",
        "amo_domain": os.getenv("AMO_DOMAIN_AFARM", ""),
        "amo_token": os.getenv("AMO_TOKEN_AFARM", ""),
        "pipeline_id": None,
        "status_map": {},
        "won_statuses": [],
        "lost_statuses": [],
    },
    "dkauto": {
        "name": "DKAuto",
        "emoji": "🚗",
        "amo_domain": os.getenv("AMO_DOMAIN_DKAUTO", ""),
        "amo_token": os.getenv("AMO_TOKEN_DKAUTO", ""),
        "pipeline_id": None,
        "status_map": {},
        "won_statuses": [],
        "lost_statuses": [],
    },
    "logistics": {
        "name": "Логистика",
        "emoji": "🚚",
        "amo_domain": os.getenv("AMO_DOMAIN_LOGISTICS", ""),
        "amo_token": os.getenv("AMO_TOKEN_LOGISTICS", ""),
        "pipeline_id": None,
        "status_map": {},
        "won_statuses": [],
        "lost_statuses": [],
    },
}

REPORT_HOUR = 9   # час ежедневного отчёта (UTC+6 Бишкек = UTC+0 + 6)
REPORT_HOUR_UTC = 3  # 9:00 Бишкек = 03:00 UTC
