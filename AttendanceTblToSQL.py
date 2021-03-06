# -*- coding: utf-8 -*-
import win32com.client
import os.path
import sqlite3
import mysql.connector
import datetime

#For testing, reading and writing ExcelFile
def LinkExcel(xslFile):
    #Set targetFile
    if not os.path.exists(xlsFile):
        exit()
        #Create Link
    xlsApp = win32com.client.Dispatch('Excel.Application')
    xlsWorkbook = xlsApp.Workbooks.Open(xlsFile)
    xlsSheets = xlsWorkbook.Sheets
    xlsSheet = xlsSheets[0]
    #Used Zone
    SourceZone = xlsSheet.UsedRange
    #Extract Data of target zone
    for xlsRow in SourceZone.Rows:
        #print(xlsRow.Value)
        (listID, atTime, doorID, atEvent, cardID, staffName, *otherInfo), = xlsRow.Value
        if (xlsRow.Row == 1) or (staffName is None):
            continue
        print(atTime, staffName)


    #Close File & Process
    xlsWorkbook.Close()
    xlsApp.Quit()
    xlsApp = None

def IntoSQLite(xlsFile, dbFile=':memory:'):
    #AttList To DB, Just for testing
    #Set targetFile
    if not os.path.exists(xlsFile):
        exit()
        #Create Link
    #Create a connection to sqlite
    con = sqlite3.connect(dbFile)
    cur = con.cursor()
    #check if staff table is exist
    cur.execute('SELECT count(*) FROM sqlite_master WHERE type="table" and name="staff"')
    if cur.fetchone()[0] == 0:
        #create staff table
        cur.execute('''CREATE TABLE staff(
        cardID      INTEGER PRIMARY KEY,
        name    TEXT)''')
        #check if attendance table is exist
    cur.execute('SELECT count(*) FROM sqlite_master WHERE type="table" and name="attendance"')
    if cur.fetchone()[0] == 0:
        #create attendance table
        cur.execute('''CREATE TABLE attendance(
        time    TEXT,
        cardID   INTEGER)
        ''')
    con.commit()
    #cur.execute('SELECT count(*) FROM sqlite_master WHERE type="table"')
    #cur.execute('SELECT 2*3')
    #print(cur)
    #print(cur.fetchone()[0])
    #con.close()
    xlsApp = win32com.client.Dispatch('Excel.Application')
    xlsWorkbook = xlsApp.Workbooks.Open(xlsFile)
    xlsSheets = xlsWorkbook.Sheets
    xlsSheet = xlsSheets[0]
    #Used Zone
    SourceZone = xlsSheet.UsedRange
    #Extract Data of target zone
    for xlsRow in SourceZone.Rows:
        #print(xlsRow.Value)
        (listID, atTime, doorID, atEvent, cardID, staffName, *otherInfo), = xlsRow.Value
        if (xlsRow.Row == 1) or (staffName is None):
            continue
        cur.execute('SELECT cardID FROM staff WHERE name=?', (staffName,))
        if cur.fetchone() == None:
            cur.execute('INSERT INTO staff(cardID, name) VALUES (?, ?)', (cardID, staffName))
        cur.execute('INSERT INTO attendance(time, cardID) VALUES (?, ?)', (atTime, cardID))
        #print(atTime, staffName)
    con.commit()
    #Close File & Process
    xlsWorkbook.Close()
    xlsApp.Quit()
    xlsApp = None
    cur.execute('SELECT date(time), name FROM attendance NATURAL JOIN staff')
    someday = datetime.date(2013, 4, 8)
    someone = '张小哲'
    cur.execute('SELECT time(time, "+8 hour"), name FROM attendance NATURAL JOIN staff WHERE date(time) = ? AND name = ?',
        (someday, someone))
    for line in cur.fetchall():
        print(line)
    con.close()

def IntoMySQL(xlsFile):
    #AttList To DB, Just for testing
    #Set targetFile
    if not os.path.exists(xlsFile):
        exit()
        #Create Link
    #Create a connection to mysql
    config = {
        'user': 'tempuser',
        'password': 'temp',
        'host': '192.168.1.52',
        'database': 'test'
    }
    con = mysql.connector.connect(**config)
    cur = con.cursor()
    #Set CharacterSet to UTF-8
    #cur.execute('SET NAMES utf8')
    #check if staff table is exist
    cur.execute('DROP TABLE IF EXISTS staff')
    #create staff table
    cur.execute('''CREATE TABLE IF NOT EXISTS staff(
    cardID   INT(8) PRIMARY KEY,
    name     CHAR(4))
    CHARACTER SET utf8
    ''')
    #check if attendance table is exist
    cur.execute('DROP TABLE IF EXISTS attendance')
    #create attendance table
    cur.execute('''CREATE TABLE IF NOT EXISTS attendance(
    time     DATETIME,
    cardID   INT(8))
    CHARACTER SET utf8
    ''')
    con.commit()
    
    xlsApp = win32com.client.Dispatch('Excel.Application')
    xlsWorkbook = xlsApp.Workbooks.Open(xlsFile)
    xlsSheets = xlsWorkbook.Sheets
    xlsSheet = xlsSheets[0]
    #Used Zone
    SourceZone = xlsSheet.UsedRange
    #Extract Data of target zone
    for xlsRow in SourceZone.Rows:
        #print(xlsRow.Value)
        (listID, atTime, doorID, atEvent, cardID, staffName, *otherInfo), = xlsRow.Value
        if (xlsRow.Row == 1) or (staffName is None):
            continue
        cur.execute('SELECT cardID FROM staff WHERE name=%s', (staffName,))
        if cur.fetchone() == None:
            cur.execute('INSERT INTO staff(cardID, name) VALUES (%s, %s)', (cardID, staffName))
        cur.execute('INSERT INTO attendance(time, cardID) VALUES (%s, %s)', (atTime, cardID))
        #print(atTime, staffName)
    con.commit()
    #Close File & Process
    xlsWorkbook.Close()
    xlsApp.Quit()
    xlsApp = None
    # cur.execute('SELECT date(time), name FROM attendance NATURAL JOIN staff')
    # someday = datetime.date(2013, 4, 8)
    # someone = '张小哲'
    # cur.execute('SELECT time(time, "+8 hour"), name FROM attendance NATURAL JOIN staff WHERE date(time) = ? AND name = ?',
    #     (someday, someone))
    # for line in cur.fetchall():
    #     print(line)
    con.close()

if __name__ == '__main__':
    xlsFile = 'e:\\My.Work\\app\\4月份门禁卡通行情况.xls'
    dbFile = 'e:\\My.Work\\app\\test.db'
    #LinkDB('')

    #LinkExcel(xlsFile)
    #Test(xlsFile, dbFile)
    #IntoSQLite(xlsFile)
    IntoMySQL(xlsFile)