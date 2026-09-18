"""Backward-compatible proxy to the canonical planning.metrics module."""

from planning import metrics as _impl


def __getattr__(name):
    return getattr(_impl, name)


def __dir__():
    return sorted(set(globals()) | set(dir(_impl)))


if __name__ == "__main__":
    _impl.main()
