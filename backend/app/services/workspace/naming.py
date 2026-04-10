from __future__ import annotations

import re
from datetime import UTC, datetime


def timestamp() -> datetime:
    return datetime.now(UTC)


def copy_name(name: str, existing_names: set[str]) -> str:
    suffix_index = 1
    while True:
        candidate = f"{name} copy" if suffix_index == 1 else f"{name} copy {suffix_index}"
        if candidate.lower() not in existing_names:
            return candidate
        suffix_index += 1


def slugify(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", value.strip().lower()).strip("-")
    return slug or "item"


def unique_slug(value: str, existing_slugs: set[str]) -> str:
    base_slug = slugify(value)
    candidate = base_slug
    suffix = 2

    while candidate in existing_slugs:
        candidate = f"{base_slug}-{suffix}"
        suffix += 1

    return candidate


def unique_key(value: str, existing_keys: set[str]) -> str:
    base_key = slugify(value).replace("-", "_")
    candidate = base_key
    suffix = 2

    while candidate in existing_keys:
        candidate = f"{base_key}_{suffix}"
        suffix += 1

    return candidate