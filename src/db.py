"""
src/db.py - Database connection management for SQLite
Ensures safe read-only access and proper URI handling.
"""

import os
import sqlite3

DB_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "data", "app.db"))

def get_db_connection(read_only: bool = True) -> sqlite3.Connection:
    """
    Returns a configured SQLite connection.
    If read_only is True, opens with ?mode=ro to prevent lock contention.
    """
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(f"Database file not found at {DB_PATH}. Run src/generate_data.py first.")
    
    if read_only:
        uri = f"file:{DB_PATH}?mode=ro"
        conn = sqlite3.connect(uri, uri=True)
    else:
        conn = sqlite3.connect(DB_PATH)
        conn.execute("PRAGMA foreign_keys = ON;")
        
    return conn
