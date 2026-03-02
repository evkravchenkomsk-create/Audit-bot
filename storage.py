"""
Хранилище данных аудита (JSON-файлы)
В продакшне можно заменить на SQLite или PostgreSQL
"""

import json
import os
import uuid
from datetime import datetime
from typing import Optional

from data import AUDIT_BLOCKS


class AuditStorage:
    def __init__(self, data_dir: str = "audit_data"):
        self.data_dir = data_dir
        os.makedirs(data_dir, exist_ok=True)

    def _audit_path(self, audit_id: str) -> str:
        return os.path.join(self.data_dir, f"{audit_id}.json")

    def _user_path(self, user_id: int) -> str:
        return os.path.join(self.data_dir, f"user_{user_id}.json")

    def create_audit(self, user_id: int, company: str) -> str:
        audit_id = str(uuid.uuid4())[:8]
        audit_data = {
            "id": audit_id,
            "user_id": user_id,
            "company": company,
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat(),
            "answers": {},  # block_idx -> q_idx -> {score, comment}
        }

        with open(self._audit_path(audit_id), 'w', encoding='utf-8') as f:
            json.dump(audit_data, f, ensure_ascii=False, indent=2)

        # Register in user's audit list
        user_audits = self._load_user_audits(user_id)
        user_audits.append({
            "id": audit_id,
            "company": company,
            "created_at": audit_data["created_at"]
        })
        with open(self._user_path(user_id), 'w', encoding='utf-8') as f:
            json.dump(user_audits, f, ensure_ascii=False, indent=2)

        return audit_id

    def get_audit(self, audit_id: str) -> Optional[dict]:
        path = self._audit_path(audit_id)
        if not os.path.exists(path):
            return None
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def _save_audit(self, audit: dict):
        audit["updated_at"] = datetime.now().isoformat()
        with open(self._audit_path(audit["id"]), 'w', encoding='utf-8') as f:
            json.dump(audit, f, ensure_ascii=False, indent=2)

    def save_answer(self, audit_id: str, block_idx: int, q_idx: int, score: int):
        audit = self.get_audit(audit_id)
        if not audit:
            return

        block_key = str(block_idx)
        q_key = str(q_idx)

        if block_key not in audit["answers"]:
            audit["answers"][block_key] = {}

        if q_key not in audit["answers"][block_key]:
            audit["answers"][block_key][q_key] = {}

        audit["answers"][block_key][q_key]["score"] = score
        self._save_audit(audit)

    def save_comment(self, audit_id: str, block_idx: int, q_idx: int, comment: str):
        audit = self.get_audit(audit_id)
        if not audit:
            return

        block_key = str(block_idx)
        q_key = str(q_idx)

        if block_key not in audit["answers"]:
            audit["answers"][block_key] = {}
        if q_key not in audit["answers"][block_key]:
            audit["answers"][block_key][q_key] = {}

        audit["answers"][block_key][q_key]["comment"] = comment
        self._save_audit(audit)

    def get_block_answers(self, audit_id: str, block_idx: int) -> dict:
        audit = self.get_audit(audit_id)
        if not audit:
            return {}
        return audit.get("answers", {}).get(str(block_idx), {})

    def get_completed_blocks(self, audit_id: str) -> list:
        """Блок считается завершённым, если отвечены все вопросы"""
        audit = self.get_audit(audit_id)
        if not audit:
            return []
        completed = []
        for i, block in enumerate(AUDIT_BLOCKS):
            answers = audit.get("answers", {}).get(str(i), {})
            if len(answers) >= len(block["questions"]):
                completed.append(i)
        return completed

    def get_total_score(self, audit_id: str) -> tuple[int, int]:
        audit = self.get_audit(audit_id)
        if not audit:
            return 0, 0

        total = 0
        for block_key, block_answers in audit.get("answers", {}).items():
            for q_key, answer in block_answers.items():
                total += answer.get("score", 0)

        max_total = sum(b["max"] for b in AUDIT_BLOCKS)
        return total, max_total

    def get_block_score(self, audit_id: str, block_idx: int) -> int:
        answers = self.get_block_answers(audit_id, block_idx)
        return sum(a.get("score", 0) for a in answers.values())

    def get_stop_factors(self, audit_id: str) -> list:
        """Возвращает список отмеченных стоп-факторов"""
        audit = self.get_audit(audit_id)
        if not audit:
            return []
        return audit.get("stop_factors", [])

    def _load_user_audits(self, user_id: int) -> list:
        path = self._user_path(user_id)
        if not os.path.exists(path):
            return []
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def get_user_audits(self, user_id: int) -> list:
        return self._load_user_audits(user_id)

    def get_all_answers(self, audit_id: str) -> dict:
        audit = self.get_audit(audit_id)
        if not audit:
            return {}
        return audit.get("answers", {})
