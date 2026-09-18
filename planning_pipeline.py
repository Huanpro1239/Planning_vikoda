"""Backward-compatible proxy to the canonical planning.pipeline module."""

from planning import pipeline as _impl

SOURCE_KEYS = _impl.SOURCE_KEYS
prepare_pipeline_output = _impl.prepare_pipeline_output

__all__ = ["SOURCE_KEYS", "prepare_pipeline_output"]


def __getattr__(name):
    return getattr(_impl, name)


def __dir__():
    return sorted(set(globals()) | set(dir(_impl)))
