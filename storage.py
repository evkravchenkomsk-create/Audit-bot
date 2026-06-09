"""
Хранилище данных трекера (JSON-файлы)
Таблицы: diary, habits, habit_done, tasks, finance, moods
"""
import json
import os
import uuid
from datetime import datetime
from typing import Optional

class TrackerStorage:
    def __init__(self, data_dir: str = "tracker_data"):
        self.data_dir = data_dir
        os.makedirs(data_dir, exist_ok=True)

    def _path(self, name: str) -> str:
        return os.path.join(self.data_dir, f"{name}.json")

    def _load(self, name: str, default=None):
        path = self._path(name)
        if not os.path.exists(path):
            return default if default is not None else {}
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default if default is not None else {}

    def _save(self, name: str, data):
        with open(self._path(name), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _uid(self) -> str:
        return str(uuid.uuid4())[:8]

    # ══ DIARY ══
    def get_diary(self, user_id: int) -> list:
        data = self._load(f"diary_{user_id}", [])
        return data if isinstance(data, list) else []

    def add_diary(self, user_id: int, text: str, mood: int = 3):
        entries = self.get_diary(user_id)
        entries.append({
            "id": self._uid(),
            "text": text,
            "mood": mood,
            "created_at": datetime.now().isoformat()
        })
        self._save(f"diary_{user_id}", entries)

    def delete_diary(self, user_id: int, entry_id: str):
        entries = [e for e in self.get_diary(user_id) if e["id"] != entry_id]
        self._save(f"diary_{user_id}", entries)

    # ══ MOODS ══
    def get_moods(self, user_id: int) -> dict:
        return self._load(f"moods_{user_id}", {})

    def set_mood(self, user_id: int, date_str: str, mood: int):
        moods = self.get_moods(user_id)
        moods[date_str] = mood
        self._save(f"moods_{user_id}", moods)

    # ══ HABITS ══
    def get_habits(self, user_id: int) -> list:
        data = self._load(f"habits_{user_id}", [])
        return data if isinstance(data, list) else []

    def add_habit(self, user_id: int, name: str, icon: str, goal: int, unit: str):
        habits = self.get_habits(user_id)
        habits.append({
            "id": self._uid(),
            "name": name,
            "icon": icon,
            "goal": goal,
            "unit": unit,
            "created_at": datetime.now().isoformat()
        })
        self._save(f"habits_{user_id}", habits)

    def delete_habit(self, user_id: int, habit_id: str):
        habits = [h for h in self.get_habits(user_id) if h["id"] != habit_id]
        self._save(f"habits_{user_id}", habits)
        # Clean up done records
        done = self._load(f"habit_done_{user_id}", {})
        done.pop(habit_id, None)
        self._save(f"habit_done_{user_id}", done)

    def is_habit_done(self, habit_id: str, date_str: str) -> bool:
        # We need user_id but habit_id is enough for cross-user lookup
        # Store done per habit_id globally
        done = self._load(f"hdone_{habit_id}", [])
        return date_str in done

    def toggle_habit_done(self, habit_id: str, user_id: int, date_str: str):
        done = self._load(f"hdone_{habit_id}", [])
        if date_str in done:
            done.remove(date_str)
        else:
            done.append(date_str)
        self._save(f"hdone_{habit_id}", done)

    # ══ TASKS ══
    def get_tasks(self, user_id: int) -> list:
        data = self._load(f"tasks_{user_id}", [])
        return data if isinstance(data, list) else []

    def add_task(self, user_id: int, text: str, task_type: str, due: str = ""):
        tasks = self.get_tasks(user_id)
        tasks.append({
            "id": self._uid(),
            "text": text,
            "type": task_type,
            "due": due,
            "done": False,
            "created_at": datetime.now().isoformat()
        })
        self._save(f"tasks_{user_id}", tasks)

    def toggle_task_done(self, user_id: int, task_id: str):
        tasks = self.get_tasks(user_id)
        for t in tasks:
            if t["id"] == task_id:
                t["done"] = not t.get("done", False)
                break
        self._save(f"tasks_{user_id}", tasks)

    def delete_task(self, user_id: int, task_id: str):
        tasks = [t for t in self.get_tasks(user_id) if t["id"] != task_id]
        self._save(f"tasks_{user_id}", tasks)

    # ══ FINANCE ══
    def get_finance(self, user_id: int) -> list:
        data = self._load(f"finance_{user_id}", [])
        return data if isinstance(data, list) else []

    def add_finance(self, user_id: int, fin_type: str, amount: float,
                    currency: str, category: str, description: str):
        records = self.get_finance(user_id)
        records.append({
            "id": self._uid(),
            "type": fin_type,
            "amount": amount,
            "currency": currency,
            "category": category,
            "description": description,
            "created_at": datetime.now().isoformat()
        })
        self._save(f"finance_{user_id}", records)

    def delete_finance(self, user_id: int, record_id: str):
        records = [r for r in self.get_finance(user_id) if r["id"] != record_id]
        self._save(f"finance_{user_id}", records)
