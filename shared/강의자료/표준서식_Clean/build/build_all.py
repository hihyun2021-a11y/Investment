# -*- coding: utf-8 -*-
"""강의용 표준서식(Clean) 5종 일괄 생성: python3 build_all.py"""
import importlib
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from common import render  # noqa: E402

DATE = "20261011"
SPECS = [
    ("f1_spa", "부동산매매계약서", "서식_01부동산매매계약서_강의용Clean_v01.docx"),
    ("f2_loan", "담보대출약정서", "서식_02담보대출약정서_강의용Clean_v01.docx"),
    ("f3_loc_equity", "인수확약서(증권사)", "서식_03인수확약서_증권사_종류주식_강의용Clean_v01.docx"),
    ("f4_loc_loan", "인수확약서(대출)", "서식_04인수확약서_대출_금융확약서_강의용Clean_v01.docx"),
    ("f5_lease", "임대차계약서(가상)", "서식_05임대차계약서_가상_강의용Clean_v01.docx"),
]

for mod, title, fname in SPECS:
    m = importlib.import_module(mod)
    path = os.path.join(OUT, f"{DATE}_{fname}")
    render(m.B, path, title)
    print("saved", path)
