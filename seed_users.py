"""
Run this once (or whenever you add new voters) to populate the users table
with valid voter IDs and starting passwords.

Edit the `sample_users` list below with your real student/teacher list
before running it for the actual election. Each voter's starting password
defaults to their own Voter ID -- tell them to change it after first login
at the "Change password" link on the ballot page.

Usage:
    python seed_users.py
"""

import sqlite3
from werkzeug.security import generate_password_hash

conn = sqlite3.connect('database.db')
cursor = conn.cursor()

cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        voter_id TEXT PRIMARY KEY,
        full_name TEXT NOT NULL,
        role TEXT CHECK(role IN ('Student', 'Teacher', 'Admin')),
        has_voted INTEGER DEFAULT 0,
        password_hash TEXT
    )
''')

# Sample users: (Voter ID, Name, Role, Starting Password)
# Replace this list with your actual GREAT AUNTY AYO COLLEGE, MOWE
# student and staff register for the 2026 election.
sample_users = [
    ('STU101', 'John Doe', 'Student', 'STU101'),
    ('STU102', 'Mary Jane', 'Student', 'STU102'),
    ('TCH001', 'Mr. Smith', 'Teacher', 'TCH001'),
    ('ADM001', 'Admin Staff', 'Admin', 'ADM001'),
]

for voter_id, full_name, role, starting_password in sample_users:
    cursor.execute(
        '''INSERT INTO users (voter_id, full_name, role, password_hash)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(voter_id) DO UPDATE SET
             full_name=excluded.full_name,
             role=excluded.role
        ''',
        (voter_id, full_name, role, generate_password_hash(starting_password))
    )

conn.commit()
conn.close()
print(f"{len(sample_users)} voter records added/verified successfully!")
print("Starting password for each voter is their own Voter ID (they should change it after logging in).")
