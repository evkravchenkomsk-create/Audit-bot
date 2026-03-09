"""
Клиент AmoCRM API
"""
import asyncio
import aiohttp
from datetime import datetime, timedelta
from collections import defaultdict
from typing import Optional


class AmoClient:
    def __init__(self, domain: str, token: str):
        self.domain = domain
        self.token = token
        self.base = f"https://{domain}/api/v4"
        self.headers = {"Authorization": f"Bearer {token}"}

    async def _get(self, path: str, params: dict = None) -> dict:
        timeout = aiohttp.ClientTimeout(total=15)
        async with aiohttp.ClientSession(timeout=timeout) as session:
            async with session.get(
                f"{self.base}{path}",
                headers=self.headers,
                params=params or {}
            ) as r:
                if r.status == 204:
                    return {}
                return await r.json()

    async def get_leads_page(self, page: int = 1, limit: int = 250, params: dict = None) -> list:
        p = {"limit": limit, "page": page, "with": "tags", **(params or {})}
        data = await self._get("/leads", p)
        return data.get("_embedded", {}).get("leads", [])

    async def get_all_leads(self, created_from: Optional[int] = None) -> list:
        leads = []
        page = 1
        params = {}
        if created_from:
            params["filter[created_at][from]"] = created_from
        while True:
            batch = await self.get_leads_page(page, params=params)
            if not batch:
                break
            leads.extend(batch)
            if len(batch) < 250:
                break
            page += 1
            await asyncio.sleep(0.3)
        return leads


def get_manager(tags: list) -> str:
    tag_names = [t["name"] for t in tags]
    if "Олеся" in tag_names:
        return "Олеся"
    return "Каниет"


def build_report(leads: list, project_cfg: dict, period_label: str) -> str:
    status_map = project_cfg["status_map"]
    won = project_cfg["won_statuses"]
    lost = project_cfg["lost_statuses"]
    meetings = project_cfg.get("meeting_statuses", [])
    emoji = project_cfg["emoji"]
    name = project_cfg["name"]

    total = len(leads)
    if total == 0:
        return f"{emoji} <b>{name}</b>\n📭 Нет заявок за период"

    won_deals = [l for l in leads if l["status_id"] in won]
    lost_deals = [l for l in leads if l["status_id"] in lost]
    meeting_deals = [l for l in leads if l["status_id"] in meetings]
    revenue = sum(l.get("price") or 0 for l in won_deals)

    conv = round(len(won_deals) / total * 100, 1) if total else 0

    lines = [
        f"{emoji} <b>{name}</b> — {period_label}",
        f"{'─' * 28}",
        f"📥 Создано заявок: <b>{total}</b>",
        f"✅ Продажи (Деньги в кассе): <b>{len(won_deals)}</b> ({conv}%)",
        f"📅 Назначено встреч: <b>{len(meeting_deals)}</b>",
        f"❌ Отказы: <b>{len(lost_deals)}</b> ({round(len(lost_deals)/total*100,1)}%)",
    ]
    if revenue:
        lines.append(f"💰 Выручка: <b>{revenue:,} сом</b>")

    # Воронка — распределение по статусам
    pipeline = defaultdict(int)
    for l in leads:
        status = status_map.get(l["status_id"], f"#{l['status_id']}")
        pipeline[status] += 1

    FUNNEL_ORDER = [
        "Неразобранное", "Физ лица", "Первый контакт",
        "Вышли на ЛПР", "Назначено демо", "Демо проведено",
        "Принимают решение", "Деньги в кассе",
        "Партнёры", "Мотивированный отказ", "Успешно реализовано",
        "Закрыто / не реализовано",
    ]
    funnel_lines = []
    for status in FUNNEL_ORDER:
        count = pipeline.get(status, 0)
        if count > 0:
            funnel_lines.append(f"  {status}: {count}")
    # Неизвестные статусы
    for status, count in pipeline.items():
        if status not in FUNNEL_ORDER:
            funnel_lines.append(f"  {status}: {count}")

    if funnel_lines:
        lines.append(f"\n<b>По статусам:</b>")
        lines.extend(funnel_lines)

    return "\n".join(lines)


async def get_project_report(project_cfg: dict, days: int = 1) -> str:
    domain = project_cfg.get("amo_domain", "")
    token = project_cfg.get("amo_token", "")

    if not domain or not token:
        return f"{project_cfg['emoji']} <b>{project_cfg['name']}</b>\n⚙️ AmoCRM не настроен"

    from_ts = int((datetime.now() - timedelta(days=days)).timestamp())
    label = "сегодня" if days == 1 else f"за {days} дней"

    try:
        client = AmoClient(domain, token)
        leads = await client.get_all_leads(created_from=from_ts)
        return build_report(leads, project_cfg, label)
    except Exception as e:
        return f"{project_cfg['emoji']} <b>{project_cfg['name']}</b>\n❌ Ошибка: {e}"
