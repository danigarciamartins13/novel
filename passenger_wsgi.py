import sys
import os

# Garante que a pasta da app está no path
sys.path.insert(0, os.path.dirname(__file__))

from app import app as application, init_db, init_admin

init_db()
init_admin()
