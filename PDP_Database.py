# -*- coding: utf-8 -*-
import os.path
import sqlite3
import win32com.client
import re


def connect_autocad(file):
    acad_progid = 'AutoCAD.Application.16.2'  # For AutoCAD2006
    acad_app = win32com.client.gencache.EnsureDispatch(acad_progid)
    acad_app.Visible = False
    acad_docs = acad_app.Documents
    pipe_list = set()
    strainer_list = set()
    equip_list = set()
    inst_list = set()
    try:
        print('Connecting...')
        acad_doc = acad_docs.Open(file)
        acad_entities = acad_doc.ModelSpace
        print(acad_entities.Count)
        counter = 0
        for acad_entity in acad_entities:
            if acad_entity.ObjectName == 'AcDbBlockReference':
                blockref = win32com.client.CastTo(acad_entity, 'IAcadBlockReference2')
                block_name = blockref.EffectiveName.lower()
                # Pipe Tag
                if block_name == 'tag_number':
                    attrs = blockref.GetAttributes()
                    for attr in attrs:
                        if attr.TagString == 'TAG' and attr.TextString.strip():
                            pipe_list.add(attr.TextString.strip())
                            counter += 1
                            break
                    continue
                    # Strainer
                if block_name.startswith('strainer'):
                    for attr in blockref.GetAttributes():
                        if attr.TagString == 'TAG' and attr.TextString.strip():
                            strainer_list.add(attr.TextString.strip())
                            break
                    continue
                    # Equip
                if block_name.startswith('equiptag'):
                    for attr in blockref.GetAttributes():
                        if attr.TagString == 'TAG' and attr.TextString.strip():
                            equip_list.add(attr.TextString.strip())
                            break
                    continue
                    # Inst
                if block_name in ('di_local', 'sh_pri_front', 'interlock', 'sc_local'):
                    inst_func = ''
                    inst_loop = ''
                    for attr in blockref.GetAttributes():
                        if attr.TagString == 'FUNCTION' and attr.TextString.strip():
                            inst_func = attr.TextString.strip()
                            continue
                        if attr.TagString == 'TAG' and attr.TextString.strip():
                            inst_loop = attr.TextString.strip()
                            continue
                        if inst_func and inst_loop:
                            inst_list.add('-'.join((inst_func, inst_loop)))
                            break
                    continue
        acad_doc.Close()
    finally:
        acad_app.Quit()
        print(strainer_list)
        process_pipedata(pipe_list)
        process_strainer(strainer_list)
        print(equip_list)
        print(inst_list)


def unit_number(loop_number):
    unit_code = int(str(loop_number)[:1])
    #Area 10000~20000 => Unit 4
    if unit_code < 30000:
        return 400
        #Area 30000~50000 => Unit 5
    if unit_code < 60000:
        return 500
        #Area 60000 => Unit 6
    if unit_code < 70000:
        return 600
        #Area 70000 => Unit 7
    if unit_code < 80000:
        return 700
        #Area 90000 => Unit 9
    if unit_code < 100000:
        return 900


def process_pipedata(pipe_list):
    print('Processing Pipe...')
    pipe_dict = dict()
    for pipe in pipe_list:
        #print(pipe)
        # Pipe "Service Loop - Diameter - Class [- InsulType]"
        pipe_tag = re.match('([A-Z]+)(\d+)-(\w+)-(\w+)-*(\w*)', pipe).groups()
        pipe_service, pipe_loop, pipe_dn, pipe_cls, pipe_insul = pipe_tag
        pipe_loop = int(pipe_loop)
        if pipe_service in pipe_dict:
            if not pipe_loop in pipe_dict[pipe_service]:
                pipe_dict[pipe_service].append(pipe_loop)
        else:
            pipe_dict[pipe_service] = [pipe_loop]
            #print(pipe_tag)
    print(pipe_dict)
    for loop_list in pipe_dict.values():
        loop_list.sort()
        area_prev = 0
        counter = 1
        for loop_index in range(len(loop_list)):
            area_curr = unit_number(loop_list[loop_index])
            if area_curr != area_prev:
                counter = 1
            loop_list[loop_index] = area_curr + counter
            counter += 1
            area_prev = area_curr
    print(pipe_dict)


def process_strainer(strainer_list):
    print('Processing Strainer...')
    for strainer in strainer_list:
        # Strainer "STR - Loop Suffix"
        strainer_code, strainer_loop, strainer_suffix = re.match('(\w+)-(\d+)(\w*)', strainer).groups()
        strainer_loop = int(strainer_loop)
        print(strainer, strainer_loop)


def process_equip(equip_list):
    print('Processing Equip...')
    for equip in equip_list:
        #Equip "EquipCode - Tag Suffix"
        equip_code, equip_tag, equip_suffix = re.match('(\w+)-(\d+)(\w*)', equip).groups()
        equip_tag = int(equip_tag)
        


if __name__ == '__main__':
    file = 'd:\\Work\\Project\\ShunCheng.SNG.Liq\\PnID\\ShunCheng.SNG.Liq.PnID_2014.0112A.dwg'
    connect_autocad(file)
