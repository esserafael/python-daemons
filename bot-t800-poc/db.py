import sqlite3 as sl 

con = sl.connect('t-800.db')

with con:
    con.execute("""
        CREATE TABLE Calls (
            id	INTEGER NOT NULL,
            id_call	TEXT NOT NULL,
            id_callchain	TEXT,
            state	TEXT,
            incoming_date	TEXT,
            terminated_date	TEXT,
            join_token	TEXT,
            join_weburl	TEXT,
            PRIMARY KEY(id AUTOINCREMENT)
        );
    """)

    con.execute("""
        CREATE TABLE Meetings (
            id	INTEGER NOT NULL,
            Nome	TEXT,
            join_weburl	TEXT,
            PRIMARY KEY(id AUTOINCREMENT)
        );
    """)