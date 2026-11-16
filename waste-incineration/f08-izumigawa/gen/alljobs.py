# -*- coding: utf-8 -*-
"""全工事（メーカー 16 件＋地元業者 10 件）を 1 回だけ作って使い回す。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import local  # noqa: E402
from model2 import all_maker_jobs  # noqa: E402

_J = []


def jobs():
    if not _J:
        _J.extend(all_maker_jobs())
        _J.extend(local.build())
        _J.sort(key=lambda j: (j["fy"], j["cid"]))
    return _J
