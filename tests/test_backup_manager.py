# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Tests for core/backup_manager.py."""
from __future__ import annotations

import os
import time

from core.backup_manager import _prune_backups, create_backup

# ── create_backup ─────────────────────────────────────────────────────────────

class TestCreateBackup:
    def test_backup_file_created(self, tmp_path):
        src = tmp_path / "workbook.xlsx"
        src.write_bytes(b"fake excel data")
        backup = create_backup(str(src))
        assert os.path.exists(backup)

    def test_backup_in_backups_subdir(self, tmp_path):
        src = tmp_path / "workbook.xlsx"
        src.write_bytes(b"data")
        backup = create_backup(str(src))
        assert os.path.basename(os.path.dirname(backup)) == "backups"

    def test_backup_name_has_timestamp(self, tmp_path):
        src = tmp_path / "myfile.xlsx"
        src.write_bytes(b"data")
        backup = create_backup(str(src))
        name = os.path.basename(backup)
        # Name format: myfile_YYYYMMDD_HHMMSS.xlsx
        assert name.startswith("myfile_")
        assert name.endswith(".xlsx")
        parts = name[len("myfile_"):-len(".xlsx")]
        assert len(parts) == len("YYYYMMDD_HHMMSS")

    def test_backup_contents_match_original(self, tmp_path):
        src = tmp_path / "book.xlsm"
        content = b"binary excel content 12345"
        src.write_bytes(content)
        backup = create_backup(str(src))
        with open(backup, "rb") as f:
            assert f.read() == content

    def test_original_file_unchanged(self, tmp_path):
        src = tmp_path / "book.xlsx"
        content = b"original content"
        src.write_bytes(content)
        create_backup(str(src))
        assert src.read_bytes() == content

    def test_backup_path_returned(self, tmp_path):
        src = tmp_path / "book.xlsx"
        src.write_bytes(b"x")
        result = create_backup(str(src))
        assert isinstance(result, str)
        assert os.path.isabs(result)

    def test_backups_dir_created_if_missing(self, tmp_path):
        src = tmp_path / "book.xlsx"
        src.write_bytes(b"x")
        backup_dir = tmp_path / "backups"
        assert not backup_dir.exists()
        create_backup(str(src))
        assert backup_dir.exists()

    def test_multiple_backups_created(self, tmp_path):
        src = tmp_path / "book.xlsx"
        src.write_bytes(b"x")
        b1 = create_backup(str(src))
        time.sleep(1.1)  # ensure distinct timestamps
        b2 = create_backup(str(src))
        assert b1 != b2
        assert os.path.exists(b1)
        assert os.path.exists(b2)

    def test_xlsm_extension_preserved(self, tmp_path):
        src = tmp_path / "macro_book.xlsm"
        src.write_bytes(b"x")
        backup = create_backup(str(src))
        assert backup.endswith(".xlsm")


# ── _prune_backups ─────────────────────────────────────────────────────────────

class TestPruneBackups:
    def _make_backups(self, backup_dir: str, base: str, ext: str, count: int) -> list[str]:
        """Create `count` fake backup files with sequential timestamps."""
        created = []
        for i in range(count):
            name = f"{base}_2026010{i:01d}_120000{ext}"
            path = os.path.join(backup_dir, name)
            with open(path, "w") as f:
                f.write(f"backup {i}")
            created.append(name)
        return created

    def test_prune_removes_oldest(self, tmp_path):
        bdir = str(tmp_path)
        files = self._make_backups(bdir, "book", ".xlsx", 5)
        _prune_backups(bdir, "book", ".xlsx", max_keep=3)
        remaining = os.listdir(bdir)
        # Oldest two (files[0], files[1]) should be removed
        assert files[0] not in remaining
        assert files[1] not in remaining
        assert files[4] in remaining

    def test_prune_keeps_exact_count(self, tmp_path):
        bdir = str(tmp_path)
        self._make_backups(bdir, "book", ".xlsx", 5)
        _prune_backups(bdir, "book", ".xlsx", max_keep=3)
        remaining = [f for f in os.listdir(bdir)
                     if f.startswith("book_") and f.endswith(".xlsx")]
        assert len(remaining) == 3

    def test_prune_no_op_when_under_limit(self, tmp_path):
        bdir = str(tmp_path)
        self._make_backups(bdir, "book", ".xlsx", 2)
        _prune_backups(bdir, "book", ".xlsx", max_keep=5)
        remaining = [f for f in os.listdir(bdir)
                     if f.startswith("book_") and f.endswith(".xlsx")]
        assert len(remaining) == 2

    def test_prune_exact_limit_no_deletion(self, tmp_path):
        bdir = str(tmp_path)
        self._make_backups(bdir, "book", ".xlsx", 3)
        _prune_backups(bdir, "book", ".xlsx", max_keep=3)
        remaining = [f for f in os.listdir(bdir)
                     if f.startswith("book_") and f.endswith(".xlsx")]
        assert len(remaining) == 3

    def test_prune_ignores_other_files(self, tmp_path):
        bdir = str(tmp_path)
        self._make_backups(bdir, "book", ".xlsx", 5)
        # Add an unrelated file
        unrelated = os.path.join(bdir, "readme.txt")
        with open(unrelated, "w") as f:
            f.write("unrelated")
        _prune_backups(bdir, "book", ".xlsx", max_keep=3)
        assert os.path.exists(unrelated)

    def test_prune_silently_handles_missing_dir(self, tmp_path):
        missing_dir = str(tmp_path / "nonexistent")
        # Should not raise
        _prune_backups(missing_dir, "book", ".xlsx", max_keep=3)

    def test_prune_with_max_keep_zero(self, tmp_path):
        """max_keep=0 should remove all backups."""
        bdir = str(tmp_path)
        self._make_backups(bdir, "book", ".xlsx", 3)
        _prune_backups(bdir, "book", ".xlsx", max_keep=0)
        remaining = [f for f in os.listdir(bdir)
                     if f.startswith("book_") and f.endswith(".xlsx")]
        assert len(remaining) == 0


# ── Integration: create_backup + prune ───────────────────────────────────────

class TestBackupManagerIntegration:
    def test_max_backups_respected(self, tmp_path):
        src = tmp_path / "book.xlsx"
        src.write_bytes(b"x")
        for _ in range(5):
            time.sleep(1.1)
            create_backup(str(src), max_backups=3)
        backup_dir = tmp_path / "backups"
        files = [f for f in os.listdir(backup_dir)
                 if f.startswith("book_") and f.endswith(".xlsx")]
        assert len(files) <= 3
