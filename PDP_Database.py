# -*- coding: utf-8 -*-
import os.path
import sqlite3
import win32com.client


def connect_autocad(file):
    acad_progid = 'AutoCAD.Application.16.2'  # For AutoCAD2006
    acad_app = win32com.client.gencache.EnsureDispatch(acad_progid)
    acad_app.Visible = False
    acad_docs = acad_app.Documents
    try:
        acad_doc = acad_docs.Open(file)
        acad_entities = acad_doc.ModelSpace
        print(acad_entities.Count)
        counter = 0
        for acad_entity in acad_entities:
            if acad_entity.ObjectName == 'AcDbBlockReference':
                blockref = win32com.client.CastTo(acad_entity, 'IAcadBlockReference2')
                if blockref.EffectiveName.lower() == 'tag_number':
                    attrs = blockref.GetAttributes()
                    for attr in attrs:
                        if attr.TagString == 'TAG':
                            print(attr.TextString)

        acad_doc.Close()
    finally:
        acad_app.Quit()


if __name__ == '__main__':
    file = 'E:\\My.Work\\Project\\ShunCheng.SNG.Liq\\PnID\\ShunCheng.SNG.Liq.PnID_2014.0112A.dwg'
    connect_autocad(file)
