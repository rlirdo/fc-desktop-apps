# -*- coding: utf-8 -*-
import os
import sys


def app_dir():
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def resource_path(rel):
    return os.path.join(app_dir(), rel)
