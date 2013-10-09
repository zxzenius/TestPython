# -*- coding: utf-8 -*-
import win32com.client
import os.path
import mysql.connector
import datetime


def IntoMySQL(xlsFile):
    #AttList To DB, Just for testing
    #Set targetFile
    if not os.path.exists(xlsFile):
        exit()
        #Create Link
    #Create a connection to mysql
    config = {
        'user': 'xydbadmin',
        'password': 'x1nyuan1',
        'host': '10.4.8.106',
        'database': 'test'
    }
    con = mysql.connector.connect(**config)
    cur = con.cursor()
    #Set CharacterSet to UTF-8
    #cur.execute('SET NAMES utf8')
    #check if staff table is exist
    #cur.execute('DROP TABLE IF EXISTS staff')
    #create staff table
    cur.execute('''CREATE TABLE IF NOT EXISTS staff(
    card_id   INT(8) unsigned,
    name     VARCHAR(4),
    PRIMARY KEY (card_id)
    )
    ENGINE = MyISAM
    CHARACTER SET utf8
    ''')
    #check if doorevent table is exist
    #cur.execute('DROP TABLE IF EXISTS door_event')
    #create attendance table
    cur.execute('''CREATE TABLE IF NOT EXISTS door_event(
    event_code INT(10) unsigned AUTO_INCREMENT,
    time     DATETIME,
    card_id   INT(8) unsigned,
    PRIMARY KEY (event_code)
    )
    ENGINE = MyISAM
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
        cur.execute('SELECT card_id FROM staff WHERE name=%s', (staffName,))
        if cur.fetchone() == None:
            cur.execute('INSERT INTO staff(card_id, name) VALUES (%s, %s)', (cardID, staffName))
        cur.execute('INSERT INTO door_event(time, card_id) VALUES (%s, %s)', (atTime, cardID))
        #print(atTime, staffName)
    #con.commit()
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
    xlsFile = 'd:\\Work\\dev\\EXCEL\\9月份门禁卡通行情况.xls'


    #LinkExcel(xlsFile)
    #Test(xlsFile, dbFile)
    #IntoSQLite(xlsFile)
    IntoMySQL(xlsFile)