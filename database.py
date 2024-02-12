import os
import sqlite3

from config import base_dir


conn = sqlite3.connect(os.path.join(base_dir, "xlsx.db"))

