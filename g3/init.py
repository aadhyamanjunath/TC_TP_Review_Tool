"""Checkpoint review tool package exports."""

try:
    # package import
    from .logic import run
except Exception:  # pragma: no cover
    # script/folder import
    from logic import run

__all__ = ["run"]

