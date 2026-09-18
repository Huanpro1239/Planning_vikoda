"""Backward-compatible proxy to the canonical planning.khsx_ki module."""

from planning import khsx_ki as _impl


def __getattr__(name):
    return getattr(_impl, name)


def __dir__():
    return sorted(set(globals()) | set(dir(_impl)))
