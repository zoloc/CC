def generate_implment_table(implmenttable, hatable, impacttable):
    import db
    import cfg
    db.create_implement_table(implmenttable)
    iptlst = db.get_impact_harnessnumber_lst(impacttable)
    for harnessnumber in iptlst:
        orgHA = db.get_ha_orglst(hatable, harnessnumber)
        impactHA = db.get_iptha(impacttable, harnessnumber)

        for each_HA in orgHA:
            orgcfg = set(cfg.cfgstr_to_cfglst(each_HA.effectivity))
            iptcfg = set(cfg.cfgstr_to_cfglst(impactHA.effectivity))
            if set.isdisjoint(orgcfg, iptcfg):  # orgHAcfg^impactHAcfg = 0
                pass
            elif not set.isdisjoint(orgcfg, iptcfg):
                if orgcfg == iptcfg:
                    # orgHAcfg == impactHAcfg 由于后续包含判断包括相等所以先考虑相等情况
                    HAbefore = each_HA
                    HAafter = each_HA.revise()
                    db.input_data_to_implement_table(implmenttable, HAbefore, HAafter)
                elif set.issuperset(orgcfg, iptcfg):
                    # orgHAcfg >= impactHAcfg
                    HAbefore = each_HA
                    HAafter = each_HA.cfg_after_cut(impactHA.effectivity)
                    HAnew = each_HA.new_cfg_after_cut(hatable, implmenttable, impactHA.effectivity)
                    db.input_data_to_implement_table(implmenttable, HAbefore, HAafter)
                    db.input_data_to_implement_table(implmenttable, HAbefore, HAnew)
                elif set.issuperset(iptcfg, orgcfg):
                    # impactHAcfg >= orgHAcfg
                    HAbefore = each_HA
                    HAafter = each_HA.revise()
                    db.input_data_to_implement_table(implmenttable, HAbefore, HAafter)
                else:
                    # 仅相交的无包含的情况
                    HAbefore = each_HA
                    HAafter = each_HA.cfg_after_cut(impactHA.effectivity)
                    HAnew = each_HA.new_cfg_after_cut(hatable, implmenttable, cfg.cfgset_to_cfgstr(orgcfg & iptcfg))
                    db.input_data_to_implement_table(implmenttable, HAbefore, HAafter)
                    db.input_data_to_implement_table(implmenttable, HAbefore, HAnew)

def effectivity_filter(cfgfilter, IDEALtable, wbname):
    import pymysql
    import cfg
    from openpyxl import Workbook
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    wb = Workbook()
    ws = wb.active
    title = ('序号', '主部段', '部段', 'ATA章', 'ATA系统', 'DCI编号', 'DCI名称', '架次有效性', 'DDM编号', 'DDM名称', 'DDM版本', '当前状态', '对称件', '类型')
    ws.append(title)
    cur.execute('select * from ' + IDEALtable + ' where 架次有效性<>"NULL"')
    extract = cur.fetchall()
    for each_row in extract:
        cfgset = set(cfg.cfgstr_to_cfglst(each_row[7]))
        if set.issuperset(cfgset, set(cfg.cfgstr_to_cfglst(cfgfilter))):
            print(each_row)
            ws.append(each_row)
    wb.save(filename = wbname + '.xlsx')

def IDEAL_DDM_compare(wbname,cfgfilter=''):
    import pymysql
    import cfg
    from openpyxl import Workbook
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    wb = Workbook()
    ws = wb.active
    title = ('序号', '主部段', '部段', 'ATA章', 'ATA系统', 'DCI编号', 'DCI名称', '架次有效性', 'DDM编号', 'DDM名称', 'DDM版本', '当前状态', '对称件', '类型', 'from')
    ws.append(title)
    cur.execute(
        'SELECT *,"from_old"\
        FROM ddmlist_old\
        where not exists (\
        SELECT 1 FROM ddmlist_new where ddmlist_old.DDM编号=ddmlist_new.DDM编号 and ddmlist_old.DDM版本=ddmlist_new.DDM版本)\
        union all\
        SELECT *,"from_new"\
        FROM ddmlist_new\
        where not exists (\
        SELECT 1 FROM ddmlist_old where ddmlist_old.DDM编号=ddmlist_new.DDM编号 and ddmlist_old.DDM版本=ddmlist_new.DDM版本)\
        ORDER BY 序号'
        )
    extract = cur.fetchall()
    if cfgfilter == '':
        for each_row in extract:
            ws.append(each_row)
    else:
        for each_row in extract:
            cfgset = set(cfg.cfgstr_to_cfglst(each_row[7]))
            if set.issuperset(cfgset, set(cfg.cfgstr_to_cfglst(cfgfilter))):
                print(each_row)
                ws.append(each_row)
    wb.save(filename = wbname + '.xlsx')