from __future__ import annotations

INVALID_FORMULA_ALIAS_MESSAGE = (
    "Formula names cannot contain ] because spreadsheet references use bracket syntax."
)


def normalize_formula_alias(value: object) -> str | None:
    normalized = str(value).strip() if value is not None else ""
    return normalized or None


def validate_formula_alias(value: str) -> str | None:
    if "]" in value:
        return INVALID_FORMULA_ALIAS_MESSAGE

    return None


def normalize_and_validate_formula_alias(value: object) -> str | None:
    normalized = normalize_formula_alias(value)
    if normalized is None:
        return None

    validation_error = validate_formula_alias(normalized)
    if validation_error:
        raise ValueError(validation_error)

    return normalized


def sanitize_formula_alias(value: object) -> str | None:
    normalized = normalize_formula_alias(value)
    if normalized is None:
        return None

    return normalized if validate_formula_alias(normalized) is None else None