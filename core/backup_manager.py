# Copyright (c) 2026 William E. Smith (williamsmithe@icloud.com). All rights reserved.
# Proprietary and confidential. Unauthorized use, copying, or distribution
# of this file, via any medium, is strictly prohibited. See LICENSE for details.
"""Versioned backup creation for Excel files."""
from __future__ import annotations

import os
import shutil
from datetime import datetime


def create_backup(file_path: str, max_backups: int = 20) -> str:
    """
    Copy *file_path* to a timestamped backup in a ./backups/ sub-folder
    next to the original file.  Returns the backup path.

    If there are more than *max_backups* files in the folder, the oldest
    ones are deleted.
    """
    file_path = os.path.abspath(file_path)
    parent_dir = os.path.dirname(file_path)
    base, ext = os.path.splitext(os.path.basename(file_path))

    backup_dir = os.path.join(parent_dir, "backups")
    os.makedirs(backup_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{base}_{timestamp}{ext}"
    backup_path = os.path.join(backup_dir, backup_name)

    shutil.copy2(file_path, backup_path)

    # Prune old backups (keep most recent *max_backups*)
    _prune_backups(backup_dir, base, ext, max_backups)

    return backup_path


def _prune_backups(backup_dir: str, base: str, ext: str, max_keep: int) -> None:
    try:
        entries = [
            e for e in os.listdir(backup_dir)
            if e.startswith(base + "_") and e.endswith(ext)
        ]
        entries.sort()  # lexicographic = chronological for YYYYMMDD_HHMMSS
        while len(entries) > max_keep:
            oldest = entries.pop(0)
            try:
                os.remove(os.path.join(backup_dir, oldest))
            except OSError:
                pass
    except OSError:
        pass
