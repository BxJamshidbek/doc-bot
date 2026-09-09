"""User session and product storage management."""

import json
import shutil
import uuid
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Optional, Dict
from src.config import SESSION_BASE_DIR


@dataclass
class ProductItem:
    name: str
    barcode: str
    image_path: str


class UserSession:
    def __init__(self, user_id: int, session_dir: Path):
        self.user_id = user_id
        self.session_dir = session_dir
        self.products: List[ProductItem] = []
        self.current_draft: Dict[str, Optional[str]] = {
            "photo_path": None,
            "name": None,
            "barcode": None,
        }
        self.state_file = self.session_dir / "session.json"
        self._load_state()

    def _load_state(self) -> None:
        """Loads saved session from JSON if available."""
        if not self.state_file.is_file():
            return
        try:
            with open(self.state_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.products = [
                ProductItem(
                    name=item["name"],
                    barcode=item["barcode"],
                    image_path=item["image_path"],
                )
                for item in data.get("products", [])
                if Path(item.get("image_path", "")).is_file()
            ]
        except Exception:
            self.products = []

    def _save_state(self) -> None:
        """Persists session state to JSON."""
        self.session_dir.mkdir(parents=True, exist_ok=True)
        data = {
            "user_id": self.user_id,
            "products": [asdict(p) for p in self.products],
        }
        with open(self.state_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def add_product(self, name: str, barcode: str, image_path: str) -> int:
        """Adds a product to the ordered list and persists state."""
        item = ProductItem(name=name.strip(), barcode=barcode.strip(), image_path=str(image_path))
        self.products.append(item)
        self._save_state()
        return len(self.products)

    def remove_product(self, index_1_based: int) -> Optional[ProductItem]:
        """Removes a product by 1-based index."""
        if 1 <= index_1_based <= len(self.products):
            removed = self.products.pop(index_1_based - 1)
            # Try removing image file if it exists in session dir
            try:
                p = Path(removed.image_path)
                if p.is_file() and self.session_dir in p.parents:
                    p.unlink(missing_ok=True)
            except Exception:
                pass
            self._save_state()
            return removed
        return None

    def clear(self) -> None:
        """
        Clears active products and draft.
        Preserves generated DOCX/PDF export files (retained for 30 days).
        Removes active product images.
        """
        for prod in self.products:
            try:
                p = Path(prod.image_path)
                if p.is_file() and self.session_dir in p.parents:
                    p.unlink(missing_ok=True)
            except Exception:
                pass

        self.products.clear()
        self.reset_draft()
        self._save_state()
        self.cleanup_expired_files(max_age_days=30)

    def cleanup_expired_files(self, max_age_days: int = 30) -> None:
        """Cleans up export documents older than max_age_days."""
        if not self.session_dir.is_dir():
            return
        import time
        now = time.time()
        max_age_sec = max_age_days * 86400
        for f in self.session_dir.glob("*.*"):
            if f.suffix.lower() in [".docx", ".pdf"]:
                try:
                    if now - f.stat().st_mtime > max_age_sec:
                        f.unlink(missing_ok=True)
                except Exception:
                    pass

    def reset_draft(self) -> None:
        """Resets the in-progress draft item and removes uncommitted draft photos."""
        draft_photo = self.current_draft.get("photo_path")
        if draft_photo:
            try:
                p = Path(draft_photo)
                # Check if it's not referenced in any committed product
                if not any(prod.image_path == str(p) for prod in self.products):
                    p.unlink(missing_ok=True)
            except Exception:
                pass
        self.current_draft = {"photo_path": None, "name": None, "barcode": None}


class SessionManager:
    """Manages sessions for all users."""

    def __init__(self, base_dir: Path = SESSION_BASE_DIR):
        self.base_dir = base_dir
        self.base_dir.mkdir(parents=True, exist_ok=True)
        self._sessions: Dict[int, UserSession] = {}

    def get_session(self, user_id: int) -> UserSession:
        """Gets or initializes the user's session."""
        if user_id not in self._sessions:
            user_dir = self.base_dir / str(user_id)
            user_dir.mkdir(parents=True, exist_ok=True)
            self._sessions[user_id] = UserSession(user_id, user_dir)
        return self._sessions[user_id]

    def save_draft_photo(self, user_id: int, file_bytes: bytes, ext: str = ".jpg") -> str:
        """Saves an incoming photo to the user's directory."""
        session = self.get_session(user_id)
        filename = f"draft_{uuid.uuid4().hex[:8]}{ext}"
        target_path = session.session_dir / filename
        with open(target_path, "wb") as f:
            f.write(file_bytes)
        return str(target_path)
