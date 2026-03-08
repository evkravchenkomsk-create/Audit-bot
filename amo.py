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
    emoji = project_cfg["emoji"]
    name = project_cfg["name"]

    total = len(leads)
    if total == 0:
        return f"{emoji} <b>{name}</b>\n📭 Нет заявок за период"

    # Воронка
    pipeline = defaultdict(int)
    for l in leads:
        status = status_map.get(l["status_id"], f"Статус {l['status_id']}")
        pipeline[status] += 1

    won_deals = [l for l in leads if l["status_id"] in won]
    lost_deals = [l for l in leads if l["status_id"] in lost]
    revenue = sum(l.get("price") or 0 for l in won_deals)

    # Менеджеры
    mgr_stats = defaultdict(lambda: {"total": 0, "won": 0, "lost": 0})
    for l in leads:
        tags = l.get("_embedded", {}).get("tags", [])
        mgr = get_manager(tags)
        mgr_stats[mgr]["total"] += 1
        if l["status_id"] in won:
            mgr_stats[mgr]["won"] += 1
        if l["status_id"] in lost:
            mgr_stats[mgr]["lost"] += 1

    conv = round(len(won_deals) / total * 100, 1) if total else 0

    lines = [
        f"{emoji} <b>{name}</b> — {period_label}",
        f"{'─' * 30}",
        f"📥 Всего заявок: <b>{total}</b>",
        f"✅ Продажи: <b>{len(won_deals)}</b> ({conv}%)",
        f"❌ Отказы: <b>{len(lost_deals)}</b> ({round(len(lost_deals)/total*100,1)}%)",
    ]
    if revenue:
        lines.append(f"💰 Выручка: <b>{revenue:,} сом</b>")

    # Воронка (топ статусы)
    lines.append(f"\n<b>Воронка:</b>")
    FUNNEL_ORDER = [
        "Неразобранное", "Физ лица", "Первый контакт",
        "Вышли на ЛПР", "Назначено демо", "Демо проведено",
        "Принимают решение", "Деньги в кассе",
        "Мотивированный отказ", "Успешно реализовано",
    ]
    for status in FUNNEL_ORDER:
        count = pipeline.get(status, 0)
        if count > 0:
            pct = round(count / total * 100)
            lines.append(f"  {status}: {count} ({pct}%)")

    # Менеджеры
    if mgr_stats:
        lines.append(f"\n<b>Менеджеры:</b>")
        for mgr, s in sorted(mgr_stats.items()):
            c = round(s["won"] / s["total"] * 100, 1) if s["total"] else 0
            lines.append(f"  👤 {mgr}: {s['total']} заявок, {s['won']} продаж ({c}%)")

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
