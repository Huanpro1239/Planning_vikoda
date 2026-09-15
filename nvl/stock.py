"""Transitional facade cho đồng bộ tồn NVL.

API cũ vẫn nằm ở ``sync_nvl_stock.py`` trong giai đoạn migrate để workflow và test
hiện tại không bị phá. Code mới nên import từ ``nvl.stock``.
"""

from sync_nvl_stock import *  # noqa: F401,F403
