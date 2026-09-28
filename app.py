# -*- coding: utf-8 -*-
"""
LMS SMP Negeri 1 Cijambe
Sistem pembelajaran online terpadu dan portal akademik sekolah.
Peran: guru, siswa, kepala sekolah, orang tua.
Stack: Streamlit, SQLite, Google Gemini.
"""
import hashlib
import hmac
import html as _html
import inspect
import io
import json
import os
import random
import re
import sqlite3
from contextlib import closing
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from PIL import Image

try:
    import google.generativeai as genai
except Exception:  # paket belum terpasang
    genai = None
try:
    import streamlit.components.v1 as components
    from streamlit_geolocation import streamlit_geolocation
except Exception:  # paket lokasi belum terpasang
    streamlit_geolocation = None
    import streamlit.components.v1 as components

st.set_page_config(page_title="LMS SMP Negeri 1 Cijambe", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")
try:  # cadangan bila .streamlit/config.toml belum dipasang
    for _k, _v in {"base": "light", "primaryColor": "#1B4F9C", "backgroundColor": "#F3F5F9",
                   "secondaryBackgroundColor": "#FFFFFF", "textColor": "#1B2A41"}.items():
        st._config.set_option(f"theme.{_k}", _v)
