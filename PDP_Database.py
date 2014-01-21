# -*- coding: utf-8 -*-
import os.path
import mysql.connector
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
                            if attr.TextString.strip()[:1].isalpha():  # Except: PSV-MC..
                                break
                            inst_loop = attr.TextString.strip()
                            continue
                        if inst_func and inst_loop:
                            inst_list.add('-'.join((inst_func, inst_loop)))
                            break
                    continue
        acad_doc.Close()
    finally:
        acad_app.Quit()
        process_pipedata(pipe_list)
        process_strainer(strainer_list)
        process_equip(equip_list)
        process_inst(inst_list)


def unit_number(loop_number):
    unit_code = int(str(loop_number)[:1])
    #Area 10000~20000 => Unit 4
    if unit_code < 3:
        return 400
        #Area 30000~50000 => Unit 5
    if unit_code < 6:
        return 500
        #Area 60000 => Unit 6
    if unit_code < 7:
        return 600
        #Area 70000 => Unit 7
    if unit_code < 8:
        return 700
        #Area 90000 => Unit 9
    if unit_code < 10:
        return 900


def process_pipedata(pipe_list):
    print('Processing Pipe...')
    pipe_dict = dict()
    for pipe in pipe_list:
        # Pipe "Service Loop - Diameter - Class [- InsulType]"
        pipe_tag = re.match('([A-Z]+)(\d+)-(\w+)-(\w+)-*(\w*)', pipe).groups()
        pipe_service, pipe_loop, pipe_dn, pipe_cls, pipe_insul = pipe_tag
        update_dict(pipe_service, pipe_loop, pipe_dict)
    sort_dict(pipe_dict)
    for key, loop_list in pipe_dict.items():
        renumber(loop_list)
    print(pipe_dict)


def process_strainer(strainer_list):
    print('Processing Strainer...')
    strainer_dict = dict()
    for strainer in strainer_list:
        # Strainer "STR - Loop Suffix"
        strainer_code, strainer_loop, strainer_suffix = re.match('(\w+)-(\d+)(\w*)', strainer).groups()
        update_dict(strainer_code, strainer_loop, strainer_dict)
    sort_dict(strainer_dict)
    for key, loop_list in strainer_dict.items():
        renumber(loop_list)
    print(strainer_dict)


def process_equip(equip_list):
    print('Processing Equip...')
    equip_dict = dict()
    for equip in equip_list:
        #Equip "EquipCode - Tag Suffix"
        equip_code, equip_loop, equip_suffix = re.match('(\w+)-(\d+)(\w*)', equip).groups()
        update_dict(equip_code, equip_loop, equip_dict)
    sort_dict(equip_dict)
    for key, loop_list in equip_dict.items():
        renumber(loop_list)
    print(equip_dict)


def process_inst(inst_list):
    print('Processing Inst...')
    inst_dict = dict()
    for inst in inst_list:
        #Inst "Function - Loop Suffix"
        inst_code, inst_loop, inst_suffix = re.search('(\w+)-(\d+)(\w*)', inst).groups()
        update_dict(inst_code, inst_loop, inst_dict)
    sort_dict(inst_dict)
    normal_keylist = {'SC', 'FO', 'PSV', 'PRV', 'TG', 'PG', 'FG', 'LG', 'PDG', 'YL'}
    temp_keylist = {'TE', 'TT', 'TIT', 'TI', 'TIC', 'TCV'}  # Temperature Loop
    press_keylist = {'PIT', 'PT', 'PI', 'PIC', 'PCV'}  # Pressure Loop
    flow_keylist = {'FIT', 'FIQ', 'FI', 'FIC', 'FCV'}  # Flow Loop
    level_keylist = {'LIT', 'LI', 'LIC', 'LCV'}
    pd_keylist = {'PDIT', 'PDT', 'PDI'}  # Pressure Diff Loop
    ai_keylist = {'AE', 'AIT', 'AT', 'AI'}  # Analyse Loop
    kcv_keylist = {'KS', 'KCV'}  # KCV Loop
    hcv_keylist = {'HIC', 'HCV'}
    xcv_keylist = {'XCV', 'XY', 'HS', 'ZI', 'I'}
    cv_keylist = level_keylist | hcv_keylist | xcv_keylist
    for key, loop_list in inst_dict.items():
        if key in normal_keylist:
            renumber(loop_list)
    for keylist in (temp_keylist, press_keylist, flow_keylist, pd_keylist, ai_keylist, kcv_keylist, cv_keylist):
        ex_loop_list = set()
        for key in keylist:
            if key in inst_dict:
                ex_loop_list.update(inst_dict[key])
        ex_loop_list = list(ex_loop_list)
        ex_loop_list.sort()
        renumber(ex_loop_list)
        loop_dict = dict(ex_loop_list)
        #print(loop_dict)
        for key in keylist:
            if key in inst_dict:
                loop_list = inst_dict[key]
                for index in range(len(loop_list)):
                    loop_list[index] = (loop_list[index], loop_dict[loop_list[index]])
    print(inst_dict)


def update_dict(key, val, tag_dict):
    val = int(val)
    if key in tag_dict:
        if not val in tag_dict[key]:
            tag_dict[key].append(val)
    else:
        tag_dict[key] = [val]


def sort_dict(tag_dict):
    for val_list in tag_dict.values():
        val_list.sort()


def renumber(loop_list):
    area_prev = 0
    counter = 1
    for index in range(len(loop_list)):
        area_curr = unit_number(loop_list[index])
        if area_curr != area_prev:
            counter = 1
        loop_list[index] = (loop_list[index], area_curr + counter)
        counter += 1
        area_prev = area_curr


if __name__ == '__main__':
    file = 'd:\\Work\\Project\\ShunCheng.SNG.Liq\\PnID\\ShunCheng.SNG.Liq.PnID_2014.0121.dwg'
    connect_autocad(file)
