"""
Run this ONCE to hash all existing plaintext passwords in the DB.
After running, delete this file.

Usage:
    python migrate_passwords.py
"""

import pymysql
from werkzeug.security import generate_password_hash

conn = pymysql.connect(
    host="localhost",
    user="amaresh",
    password="1234",          # your actual DB password
    database="fee_management",
    cursorclass=pymysql.cursors.DictCursor
)

try:
    with conn.cursor() as cursor:
        # Fetch all users
        cursor.execute("SELECT id, password FROM users")
        users = cursor.fetchall()

        updated = 0
        skipped = 0

        for user in users:
            raw = user['password']

            # Skip if already hashed (werkzeug hashes start with "scrypt:" or "pbkdf2:")
            if raw and (raw.startswith('scrypt:') or raw.startswith('pbkdf2:')):
                skipped += 1
                continue

            hashed = generate_password_hash(raw)
            cursor.execute(
                "UPDATE users SET password=%s WHERE id=%s",
                (hashed, user['id'])
            )
            updated += 1

    conn.commit()
    print(f"✅ Done. {updated} passwords hashed, {skipped} already hashed (skipped).")

finally:
    conn.close()