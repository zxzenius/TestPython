# -*- coding: utf-8 -*-
import os.path
import mysql.connector
import win32com.client
import re
import datetime
import os


def extract_info_autocad(dwg_file):
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
        acad_doc = acad_docs.Open(dwg_file)
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
                        if attr.TagString == 'TAG' and attr.TextString.strip():
                            if attr.TextString.strip()[:1].isalpha():  # Except: PSV-MC..
                                break
                            inst_loop = attr.TextString.strip()
                        if attr.TagString == 'FUNCTION' and attr.TextString.strip():
                            inst_func = attr.TextString.strip()
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
    into_db(pipe_dict, 'pipe')


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
    into_db(strainer_dict, 'strainer')


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
    into_db(equip_dict, 'equip')


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
    kcv_keylist = {'KS', 'KCV', 'KVI'}  # KCV Loop
    hcv_keylist = {'HIC', 'HCV'}
    xcv_keylist = {'XCV', 'XY', 'HS', 'XVI', 'I'}
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
    into_db(inst_dict, 'inst')


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


def into_db(data_dict, table):
    print(' '.join(('Into Table', table)))
    con = mysql.connector.connect(**db_config())
    cur = con.cursor()
    cur.execute('DROP TABLE IF EXISTS {tb}'.format(tb=table))
    cur.execute('''CREATE TABLE IF NOT EXISTS {tb}(
    event_code INT(10) unsigned AUTO_INCREMENT,
    tag_code  VARCHAR(6),
    old_loop  INT(5) unsigned,
    new_loop  INT(4) unsigned,
    PRIMARY KEY (event_code)
    )
    ENGINE = InnoDB
    CHARACTER SET utf8
    '''.format(tb=table))
    con.commit()
    for key, val_list in data_dict.items():
        for old_loop, new_loop in val_list:
            cur.execute('INSERT INTO {tb}(tag_code, old_loop, new_loop) VALUES (%s, %s, %s)'.format(tb=table),
                        (key, old_loop, new_loop))
    con.commit()
    con.close()


def db_config():
    config = {
        'user': 'xydbadmin',
        'password': 'x1nyuan1',
        'host': '10.4.8.106',
        'database': 'test'
    }
    return config


def read_db():
    print('Extracting Data From Database...')
    tables = ('pipe', 'equip', 'strainer', 'inst')
    con = mysql.connector.connect(**db_config())
    cur = con.cursor()
    result_list = []
    for table in tables:
        cur.execute('SELECT tag_code, old_loop, new_loop FROM {tb} ORDER BY tag_code, old_loop'.format(tb=table))
        result_list.append([fetch_row for fetch_row in cur])
    return result_list


def convert_list(loop_list, sep=''):
    result_dict = dict()
    for tag_code, old_loop, new_loop in loop_list:
        result_dict[sep.join((tag_code, str(old_loop)))] = sep.join((tag_code, '{0:#04d}'.format(new_loop)))
    print(result_dict)
    return result_dict


def reform_inst(loop_list):
    result_dict = dict()
    for tag_code, old_loop, new_loop in loop_list:
        if tag_code not in result_dict.keys():
            result_dict[tag_code] = dict()
        result_dict[tag_code][str(old_loop)] = '{0:#04d}'.format(new_loop)
    print(result_dict)
    return result_dict


def update_dwg(dwg_file, pipe_list=[], equip_list=[], str_list=[], inst_list=[]):
    if pipe_list and equip_list and str_list and inst_list:
        acad_progid = 'AutoCAD.Application.16.2'  # For AutoCAD2006
        acad_app = win32com.client.gencache.EnsureDispatch(acad_progid)
        acad_app.Visible = False
        acad_docs = acad_app.Documents
        pipe_dict = convert_list(pipe_list)
        equip_dict = convert_list(equip_list, '-')
        str_dict = convert_list(str_list, '-')
        inst_dict = reform_inst(inst_list)
        insttag_dict = convert_list(inst_list, '-')
        try:
            print(''.join(('Updating...', dwg_file)))
            acad_doc = acad_docs.Open(dwg_file)
            acad_entities = acad_doc.ModelSpace
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
                                pipe_tag, sep, other = attr.TextString.strip().partition('-')
                                if pipe_tag in pipe_dict.keys():
                                    attr.TextString = sep.join((pipe_dict[pipe_tag], other))
                                    counter += 1
                                break
                        continue
                        # Strainer
                    if block_name.startswith('strainer'):
                        for attr in blockref.GetAttributes():
                            if attr.TagString == 'TAG' and attr.TextString.strip():
                                str_tag, str_suffix = re.match('(\w+-\d+)(\S*)', attr.TextString.strip()).groups()
                                if str_tag in str_dict.keys():
                                    attr.TextString = ''.join((str_dict[str_tag], str_suffix))
                                    counter += 1
                                break
                        continue
                        # Equip
                    if block_name.startswith('equip'):
                        for attr in blockref.GetAttributes():
                            if attr.TagString == 'TAG' and attr.TextString.strip():
                                equip_tag, equip_suffix = re.match('(\w+-\d+)(\S*)', attr.TextString.strip()).groups()
                                if equip_tag in equip_dict.keys():
                                    attr.TextString = ''.join((equip_dict[equip_tag], equip_suffix))
                                    counter += 1
                                break
                        continue
                        # Inst
                    if block_name in ('di_local', 'sh_pri_front', 'interlock', 'sc_local'):
                        inst_func = ''
                        inst_loop = ''
                        flag = False
                        for attr in blockref.GetAttributes():
                            if attr.TagString == 'TAG' and attr.TextString.strip():
                                if attr.TextString.strip()[:1].isalpha():  # Except: PSV-MC..
                                    break
                                inst_loop, inst_suffix = re.match('(\d+)(\S*)', attr.TextString.strip()).groups()
                            if attr.TagString == 'FUNCTION' and attr.TextString.strip():
                                inst_func = attr.TextString.strip()
                            if inst_func and inst_loop:
                                flag = True
                                break
                        if flag:
                            if inst_func in inst_dict.keys() and inst_loop in inst_dict[inst_func].keys():
                                for attr in blockref.GetAttributes():
                                    if attr.TagString == 'TAG':
                                        attr.TextString = ''.join((inst_dict[inst_func][inst_loop], inst_suffix))
                                        counter += 1
                                        break
                        continue
                    if block_name.startswith('connector'):
                        for attr in blockref.GetAttributes():
                            if attr.TagString == 'OriginOrDestination':
                                found_flag = False
                                for change_dict in (insttag_dict, equip_dict, pipe_dict):
                                    for old_tag, new_tag in change_dict.items():
                                        if old_tag in attr.TextString:
                                            found_flag = True
                                            attr.TextString = attr.TextString.replace(old_tag, new_tag, 1)
                                            counter += 1
                                            break
                                    if found_flag:
                                        break
                                break
                        continue

            acad_doc.SaveAs(new_filename(dwg_file))
            print(str(counter), ' ok')
            acad_doc.Close()
        finally:
            acad_app.Quit()


def update_doc(doc_file, tbl_list):
    word_progid = 'Word.Application'
    word_app = win32com.client.gencache.EnsureDispatch(word_progid)
    word_app.Visible = False
    word_docs = word_app.Documents
    counter = 0
    try:
        print(' '.join(('Updating...', doc_file)))
        word_doc = word_docs.Open(doc_file)
        word_range = word_doc.Content
        for tag_dict in tbl_list:
            for old_tag, new_tag in tag_dict.items():
                if word_range.Find.Execute(FindText=old_tag, ReplaceWith=new_tag, MatchCase=True,
                                           Replace=win32com.client.constants.wdReplaceAll):
                    counter += 1
        if counter > 0:
            word_doc.SaveAs(new_filename(doc_file))
    finally:
        print(''.join((str(counter), 'ok')))
        word_app.Quit()


def update_xls(xls_file, tbl_list):
    excel_progid = 'Excel.Application'
    excel_app = win32com.client.gencache.EnsureDispatch(excel_progid)
    excel_app.Visible = False
    excel_workbooks = excel_app.Workbooks
    counter = 0
    try:
        print(' '.join(('Updating...', xls_file)))
        excel_workbook = excel_workbooks.Open(xls_file)
        for excel_sheet in excel_workbook.Sheets:
            for cell in excel_sheet.UsedRange:
                found = False
                #for cell in row:
                #cell = cell.Cells(1,1)
                if (not cell.HasFormula) and cell.Value:
                    for tag_dict in tbl_list:
                        for old_tag, new_tag in tag_dict.items():
                            if old_tag in cell.Value:
                                cell.Value = cell.Value.replace(old_tag, new_tag)
                                counter += 1
                                found = True
                                break
                        if found:
                            break
        if counter > 0:
            excel_workbook.SaveAs(new_filename(xls_file))
    finally:
        print(''.join((str(counter), 'ok')))
        excel_app.Quit()


def new_filename(old_filename):
    root_name, ext_name = os.path.splitext(old_filename)
    main_name, sep, date_suffix = root_name.partition('_')
    today = datetime.date.today()
    date_suffix = today.strftime('%Y.%m%d')
    result = ''.join((main_name, sep, date_suffix, ext_name))
    return result


if __name__ == '__main__':
    base_file = 'd:\\Work\\Project\\ShunCheng.SNG.Liq\\PnID\\ShunCheng.SNG.Liq.PnID_2014.0123B.dwg'
    target_file = 'd:\\Work\\Project\\ShunCheng.SNG.Liq\\PnID\\ShunCheng.SNG.Liq.PnID_2014.0121t.dwg'
    target_path = 'd:\\Work\\Project\\ShunCheng.SNG.Liq\\test'
    doc_file = 'd:\\Work\\Project\\ShunCheng.SNG.Liq\\doc\\过滤器明细表_2014.0115.doc'
    xls_file = 'd:\\Work\\Project\\ShunCheng.SNG.Liq\\doc\\2.仪表索引表_2013.1104.xlsx'
    extract_info_autocad(base_file)
    pipes, equips, strs, insts = read_db()
    for root, dirs, files in os.walk(target_path):
        for name in files:
            pass
            #update_file()
            #update_dwg(target_file, pipes, equips, strs, insts)
            #pipe_dict = convert_list(pipes)
            #equip_dict = convert_list(equips, '-')
            #str_dict = convert_list(strs, '-')
            #inst_dict = convert_list(insts, '-')
            #tbl_list = [inst_dict, str_dict, pipe_dict, equip_dict]
            #tbl_list = [equip_dict]
            #update_doc(doc_file, tbl_list)
            #update_xls(xls_file, tbl_list)
            #for tl in read_db():
    #    convert_list(tl)
    #    reform_inst(tl)
