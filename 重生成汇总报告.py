# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import importlib
mod = importlib.import_module('股票交易分析系统')
mod.generate_summary_html()
print("OK")
