"""User session and product storage management."""

import json
import re
import shutil
import time
import uuid
from datetime import datetime
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Optional, Dict
from src.config import SESSION_BASE_DIR, GENERATED_BASE_DIR, GENERATED_RETENTION_DAYS


@dataclass
class ProductItem:
    name: str
    barcode: str
    image_path: str


@dataclass
class ExportRecord:
    export_id: str
    title: str
    kind: str
    docx_path: str
    pdf_path: Optional[str]
    product_count: int
    page_count: int
    created_at: str


def normalize_document_name(value: str) -> str:
    """Normalizes user-provided document names for display."""
    return re.sub(r"\s+", " ", str(value or "").strip())[:80].strip()


def slugify_document_name(value: str) -> str:
    """Creates a safe filename stem from a document name."""
    normalized = normalize_document_name(value).lower()
    slug = re.sub(r"[^\w.-]+", "-", normalized, flags=re.UNICODE)
    slug = re.sub(r"-+", "-", slug).strip("-.")
    return slug[:64] or "hujjat"


def document_title_from_export_filename(path: Path) -> str:
    """Infers a readable title from an existing generated DOCX/PDF filename."""
    stem = path.stem
    title_stem = re.sub(
        r"[-_](product_labels|labels_preview)_\d{8}_\d{6}_[a-f0-9]{6}$",
        "",
        stem,
    )
    if title_stem == stem:
        title_stem = re.sub(
            r"^(product_labels|labels_preview)_\d{8}_\d{6}_[a-f0-9]{6}$",
            "",
            stem,
        )
    title = normalize_document_name(title_stem.replace("-", " ").replace("_", " "))
    return title or "Mahsulot yorliqlari"


class UserSession:
    def __init__(self, user_id: int, session_dir: Path, generated_dir: Optional[Path] = None):
        self.user_id = user_id
        self.session_dir = session_dir.resolve()
        self.generated_dir = (generated_dir or (self.session_dir / "generated")).resolve()
        self.document_name: Optional[str] = None
        self.exports: List[ExportRecord] = []
        self.products: List[ProductItem] = []
        self.current_draft: Dict[str, Optional[str]] = {
            "photo_path": None,
            "name": None,
            "barcode": None,
        }
        self.state_file = self.session_dir / "session.json"
        self.session_dir.mkdir(parents=True, exist_ok=True)
        self.generated_dir.mkdir(parents=True, exist_ok=True)
        self._load_state()

    @staticmethod
    def _is_within(path: Path, base_dir: Path) -> bool:
        """Returns True only when path resolves inside base_dir."""
        try:
            path.resolve().relative_to(base_dir.resolve())
            return True
        except ValueError:
            return False

    @staticmethod
    def _safe_unlink(path: Path, allowed_base_dirs: List[Path]) -> None:
        """Deletes one file only when it is inside an allowed bot-owned directory."""
        resolved = path.resolve()
        if not resolved.is_file():
            return
        if not any(UserSession._is_within(resolved, base_dir) for base_dir in allowed_base_dirs):
            return
        resolved.unlink(missing_ok=True)

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
            self.document_name = normalize_document_name(data.get("document_name", "")) or None
            self.exports = [
                ExportRecord(
                    export_id=str(item.get("export_id", "")).strip(),
                    title=normalize_document_name(item.get("title", "")) or "Hujjat",
                    kind="final" if item.get("kind") == "final" else "preview",
                    docx_path=str(item.get("docx_path", "")).strip(),
                    pdf_path=str(item.get("pdf_path", "")).strip() or None,
                    product_count=int(item.get("product_count") or 0),
                    page_count=int(item.get("page_count") or 0),
                    created_at=str(item.get("created_at", "")).strip(),
                )
                for item in data.get("exports", [])
                if str(item.get("export_id", "")).strip()
            ]
            self._sync_export_records(save=False)
        except Exception:
            self.products = []
            self.document_name = None
            self.exports = []

    def _save_state(self) -> None:
        """Persists session state to JSON."""
        self.session_dir.mkdir(parents=True, exist_ok=True)
        data = {
            "user_id": self.user_id,
            "document_name": self.document_name,
            "products": [asdict(p) for p in self.products],
            "exports": [asdict(record) for record in self.exports],
        }
        with open(self.state_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def set_document_name(self, name: str) -> str:
        """Sets the active document name and persists it."""
        clean_name = normalize_document_name(name)
        if not clean_name:
            raise ValueError("Hujjat nomi bo‘sh bo‘lishi mumkin emas.")
        self.document_name = clean_name
        self._save_state()
        return clean_name

    def get_document_title(self) -> str:
        """Returns the active document title with a safe fallback."""
        return self.document_name or "Mahsulot yorliqlari"

    def add_product(self, name: str, barcode: str, image_path: str) -> int:
        """Adds a product to the ordered list and persists state."""
        item = ProductItem(name=name.strip(), barcode=barcode.strip(), image_path=str(image_path))
        self.products.append(item)
        self._save_state()
        return len(self.products)

    def get_product(self, index_1_based: int) -> Optional[ProductItem]:
        """Returns a product by 1-based index."""
        if 1 <= index_1_based <= len(self.products):
            return self.products[index_1_based - 1]
        return None

    def update_product_name(self, index_1_based: int, name: str) -> Optional[ProductItem]:
        """Updates a product name by 1-based index."""
        item = self.get_product(index_1_based)
        if not item:
            return None
        item.name = name.strip()
        self._save_state()
        return item

    def update_product_barcode(self, index_1_based: int, barcode: str) -> Optional[ProductItem]:
        """Updates a product barcode by 1-based index."""
        item = self.get_product(index_1_based)
        if not item:
            return None
        item.barcode = barcode.strip()
        self._save_state()
        return item

    def update_product_image(self, index_1_based: int, image_path: str) -> Optional[ProductItem]:
        """Updates a product image and removes the replaced bot-owned file when safe."""
        item = self.get_product(index_1_based)
        if not item:
            return None

        old_path = Path(item.image_path)
        item.image_path = str(image_path)
        self._save_state()

        try:
            resolved_old = old_path.resolve()
            still_used = any(
                Path(prod.image_path).resolve() == resolved_old
                for prod in self.products
                if prod.image_path
            )
            if not still_used:
                self._safe_unlink(resolved_old, [self.session_dir])
        except Exception:
            pass

        return item

    def swap_products(self, first_index_1_based: int, second_index_1_based: int) -> Optional[tuple[ProductItem, ProductItem]]:
        """Swaps two products by their 1-based positions and persists state."""
        if not (
            1 <= first_index_1_based <= len(self.products)
            and 1 <= second_index_1_based <= len(self.products)
        ):
            return None

        first_idx = first_index_1_based - 1
        second_idx = second_index_1_based - 1
        first_item = self.products[first_idx]
        second_item = self.products[second_idx]
        if first_idx != second_idx:
            self.products[first_idx], self.products[second_idx] = second_item, first_item
            self._save_state()
        return first_item, second_item

    def move_product_to_position(
        self,
        source_index_1_based: int,
        target_index_1_based: int,
    ) -> Optional[tuple[ProductItem, int, int]]:
        """
        Moves one product to an exact 1-based position and shifts the others.

        Example: moving #7 to #1 makes old #1 become #2, old #2 become #3, etc.
        """
        total = len(self.products)
        if not (
            1 <= source_index_1_based <= total
            and 1 <= target_index_1_based <= total
        ):
            return None

        source_idx = source_index_1_based - 1
        target_idx = target_index_1_based - 1
        item = self.products[source_idx]
        if source_idx != target_idx:
            item = self.products.pop(source_idx)
            self.products.insert(target_idx, item)
            self._save_state()
        return item, source_index_1_based, target_index_1_based

    def remove_product(self, index_1_based: int) -> Optional[ProductItem]:
        """Removes a product by 1-based index."""
        if 1 <= index_1_based <= len(self.products):
            removed = self.products.pop(index_1_based - 1)
            # Try removing image file if it exists in session dir
            try:
                p = Path(removed.image_path)
                if p.is_file() and self.session_dir in p.parents:
                    self._safe_unlink(p, [self.session_dir])
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
                    self._safe_unlink(p, [self.session_dir])
            except Exception:
                pass

        self.products.clear()
        self.document_name = None
        self.reset_draft()
        self._save_state()
        self.cleanup_expired_files()

    def new_export_docx_path(self, prefix: str) -> Path:
        """Returns a unique persistent generated DOCX path for preview/final exports."""
        safe_prefix = "product_labels" if prefix == "product_labels" else "labels_preview"
        title_prefix = f"{slugify_document_name(self.document_name)}_" if self.document_name else ""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        short_id = uuid.uuid4().hex[:6]
        self.generated_dir.mkdir(parents=True, exist_ok=True)
        return self.generated_dir / f"{title_prefix}{safe_prefix}_{timestamp}_{short_id}.docx"

    def _export_record_has_existing_file(self, record: ExportRecord) -> bool:
        """Returns True if at least one file in an export record still exists in generated storage."""
        allowed_dirs = [self.generated_dir, self.session_dir]
        for raw_path in [record.docx_path, record.pdf_path]:
            if not raw_path:
                continue
            path = Path(raw_path)
            if path.is_file() and any(self._is_within(path, base_dir) for base_dir in allowed_dirs):
                return True
        return False

    def _prune_export_records(self, save: bool = True) -> None:
        """Drops export metadata whose files no longer exist."""
        before = len(self.exports)
        self.exports = [
            record
            for record in self.exports
            if self._export_record_has_existing_file(record)
        ]
        if save and len(self.exports) != before:
            self._save_state()

    def _import_existing_export_files(self) -> bool:
        """Adds generated DOCX files that pre-date export metadata to history."""
        search_dirs = [self.generated_dir, self.session_dir]
        existing_dirs = [directory for directory in search_dirs if directory.is_dir()]
        if not existing_dirs:
            return False

        known_docx_paths = {
            Path(record.docx_path).resolve()
            for record in self.exports
            if record.docx_path
        }
        imported: List[ExportRecord] = []
        candidates = []
        for directory in existing_dirs:
            candidates.extend(directory.glob("*.docx"))

        for docx_path in sorted(candidates, key=lambda p: p.stat().st_mtime if p.is_file() else 0, reverse=True):
            try:
                resolved_docx = docx_path.resolve()
                if resolved_docx in known_docx_paths:
                    continue
                if not any(self._is_within(resolved_docx, base_dir) for base_dir in search_dirs):
                    continue

                pdf_path = resolved_docx.with_suffix(".pdf")
                pdf_value = str(pdf_path) if pdf_path.is_file() else None
                kind = "final" if "product_labels" in resolved_docx.stem else "preview"
                imported.append(
                    ExportRecord(
                        export_id=uuid.uuid5(uuid.NAMESPACE_URL, str(resolved_docx)).hex[:12],
                        title=document_title_from_export_filename(resolved_docx),
                        kind=kind,
                        docx_path=str(resolved_docx),
                        pdf_path=pdf_value,
                        product_count=0,
                        page_count=0,
                        created_at=datetime.fromtimestamp(resolved_docx.stat().st_mtime).isoformat(timespec="seconds"),
                    )
                )
                known_docx_paths.add(resolved_docx)
            except Exception:
                continue

        if imported:
            self.exports.extend(imported)
            self.exports.sort(key=lambda record: record.created_at, reverse=True)
            self.exports = self.exports[:100]
            return True
        return False

    def _sync_export_records(self, save: bool = True) -> None:
        """Keeps export history aligned with generated files on disk."""
        before = [asdict(record) for record in self.exports]
        self._prune_export_records(save=False)
        self._import_existing_export_files()
        if save and [asdict(record) for record in self.exports] != before:
            self._save_state()

    def record_export(
        self,
        *,
        kind: str,
        docx_path: str | Path,
        pdf_path: Optional[str | Path],
        product_count: int,
        page_count: int,
        title: Optional[str] = None,
    ) -> ExportRecord:
        """Stores metadata for a generated preview/final document."""
        record = ExportRecord(
            export_id=uuid.uuid4().hex[:12],
            title=normalize_document_name(title) or self.get_document_title(),
            kind="final" if kind == "final" else "preview",
            docx_path=str(Path(docx_path).resolve()),
            pdf_path=str(Path(pdf_path).resolve()) if pdf_path else None,
            product_count=int(product_count),
            page_count=int(page_count),
            created_at=datetime.now().isoformat(timespec="seconds"),
        )
        self.exports.insert(0, record)
        self.exports = self.exports[:100]
        self._save_state()
        return record

    def list_exports(self, limit: int = 10) -> List[ExportRecord]:
        """Returns recent generated documents whose files still exist."""
        self._sync_export_records()
        return self.exports[:limit]

    def get_export(self, export_id: str) -> Optional[ExportRecord]:
        """Finds an export record by id if its files still exist."""
        self._sync_export_records()
        normalized_id = str(export_id or "").strip()
        for record in self.exports:
            if record.export_id == normalized_id:
                return record
        return None

    def cleanup_expired_files(
        self,
        max_age_days: int = GENERATED_RETENTION_DAYS,
        now_ts: Optional[float] = None,
    ) -> None:
        """
        Cleans old generated documents and stale draft/temp files.
        Active product images and the active draft image are preserved.
        """
        now = time.time() if now_ts is None else now_ts
        max_age_sec = max_age_days * 86400
        active_paths = {
            Path(prod.image_path).resolve()
            for prod in self.products
            if prod.image_path
        }
        draft_photo = self.current_draft.get("photo_path")
        if draft_photo:
            active_paths.add(Path(draft_photo).resolve())

        cleanup_targets = []
        if self.generated_dir.is_dir():
            cleanup_targets.extend(self.generated_dir.glob("*.*"))
        if self.session_dir.is_dir():
            cleanup_targets.extend(self.session_dir.glob("*.*"))

        for f in cleanup_targets:
            try:
                if not f.is_file() or f.resolve() in active_paths:
                    continue
                if now - f.stat().st_mtime <= max_age_sec:
                    continue
                suffix = f.suffix.lower()
                is_export = suffix in {".docx", ".pdf"}
                is_stale_session_file = f.name.startswith("draft_") or f.name == "temp_check.png"
                if is_export or is_stale_session_file:
                    self._safe_unlink(f, [self.session_dir, self.generated_dir])
            except Exception:
                pass
        self._prune_export_records()

    def reset_draft(self) -> None:
        """Resets the in-progress draft item and removes uncommitted draft photos."""
        draft_photo = self.current_draft.get("photo_path")
        if draft_photo:
            try:
                p = Path(draft_photo)
                # Check if it's not referenced in any committed product
                if not any(prod.image_path == str(p) for prod in self.products):
                    self._safe_unlink(p, [self.session_dir])
            except Exception:
                pass
        self.current_draft = {"photo_path": None, "name": None, "barcode": None}


class SessionManager:
    """Manages sessions for all users."""

    def __init__(
        self,
        base_dir: Path = SESSION_BASE_DIR,
        generated_base_dir: Path = GENERATED_BASE_DIR,
    ):
        self.base_dir = base_dir.resolve()
        self.generated_base_dir = generated_base_dir.resolve()
        self.base_dir.mkdir(parents=True, exist_ok=True)
        self.generated_base_dir.mkdir(parents=True, exist_ok=True)
        self._sessions: Dict[int, UserSession] = {}

    def get_session(self, user_id: int) -> UserSession:
        """Gets or initializes the user's session."""
        if user_id not in self._sessions:
            user_dir = self.base_dir / str(user_id)
            generated_dir = self.generated_base_dir / str(user_id)
            user_dir.mkdir(parents=True, exist_ok=True)
            generated_dir.mkdir(parents=True, exist_ok=True)
            self._sessions[user_id] = UserSession(user_id, user_dir, generated_dir)
        return self._sessions[user_id]

    def save_draft_photo(self, user_id: int, file_bytes: bytes, ext: str = ".jpg") -> str:
        """Saves an incoming photo to the user's directory."""
        session = self.get_session(user_id)
        filename = f"draft_{uuid.uuid4().hex[:8]}{ext}"
        target_path = session.session_dir / filename
        with open(target_path, "wb") as f:
            f.write(file_bytes)
        return str(target_path)

    def cleanup_all_expired_files(self, max_age_days: int = GENERATED_RETENTION_DAYS) -> None:
        """Runs retention cleanup for all known session/generated user directories."""
        user_ids = {
            int(p.name)
            for p in self.base_dir.iterdir()
            if p.is_dir() and p.name.isdigit()
        }
        user_ids.update(
            int(p.name)
            for p in self.generated_base_dir.iterdir()
            if p.is_dir() and p.name.isdigit()
        )

        for user_id in user_ids:
            session = self.get_session(user_id)
            session.cleanup_expired_files(max_age_days=max_age_days)
