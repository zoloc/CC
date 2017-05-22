# coding: utf-8
def input_HA(wbname):
    import xlrd
    import pymysql
    book = xlrd.open_workbook(wbname + '.xls')
    sheet = book.sheets()[0]
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    tbdrop = 'drop table if exists ' + wbname
    cur.execute(tbdrop)
    conn.commit()
    tbcreate = 'create table ' + wbname + '(id INT(11) NOT NULL AUTO_INCREMENT,\
    `Valid_or_not` VARCHAR(50) NULL DEFAULT NULL,\
    `DCI_Number` VARCHAR(30) NULL DEFAULT NULL,\
    `HA_DM_Number` VARCHAR(30) NULL DEFAULT NULL,\
    `Harness_Number` VARCHAR(30) NULL DEFAULT NULL,\
    `HA_DM_Name` VARCHAR(30) NULL DEFAULT NULL,\
    `Basic_Number` VARCHAR(20) NULL DEFAULT NULL,\
    `Configuration_No` VARCHAR(10) NULL DEFAULT NULL,\
    `HA_Version` VARCHAR(10) NULL DEFAULT NULL,\
    `Effectivity` VARCHAR(50) NULL DEFAULT NULL,\
    `GH_compared` VARCHAR(10) NULL DEFAULT NULL,\
    `Delivery_Date` VARCHAR(50) NULL DEFAULT NULL,\
    `Released_Date` VARCHAR(50) NULL DEFAULT NULL,\
    `Rejected_Date` VARCHAR(50) NULL DEFAULT NULL,\
    `Comments` VARCHAR(1000) NULL DEFAULT NULL,\
    `ECP_Number` VARCHAR(50) NULL DEFAULT NULL,\
    `IDEAL_Status` VARCHAR(50) NULL DEFAULT NULL,\
    `Item_Type` VARCHAR(50) NULL DEFAULT NULL,\
    `Path` VARCHAR(50) NULL DEFAULT NULL,\
    PRIMARY KEY (`id`))\
    COLLATE="utf8_general_ci"'
    cur.execute(tbcreate)
    conn.commit()
    query = 'insert into ' + wbname + '(Valid_or_not, DCI_Number, HA_DM_Number, Harness_Number, HA_DM_Name, Basic_Number, Configuration_No, HA_Version, Effectivity, GH_compared, Delivery_Date, Released_Date, Rejected_Date, Comments, ECP_Number, IDEAL_Status, Item_Type, Path) values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'
    for r in range(1, sheet.nrows):
        Valid_or_not = sheet.cell(r, 0).value.strip()
        DCI_Number = sheet.cell(r, 1).value.strip()
        HA_DM_Number = sheet.cell(r, 2).value.strip()
        Harness_Number = sheet.cell(r, 3).value.strip()
        HA_DM_Name = sheet.cell(r, 4).value.strip()
        Basic_Number = sheet.cell(r, 5).value.strip()
        Configuration_No = sheet.cell(r, 6).value.strip()
        HA_Version = sheet.cell(r, 7).value.strip()
        Effectivity = sheet.cell(r, 8).value.strip()
        GH_compared = sheet.cell(r, 9).value
        Delivery_Date = sheet.cell(r, 10).value
        Released_Date = sheet.cell(r, 11).value
        Rejected_Date = sheet.cell(r, 12).value
        Comments = sheet.cell(r, 13).value.strip()
        ECP_Number = sheet.cell(r, 14).value.strip()
        IDEAL_Status = sheet.cell(r, 15).value.strip()
        Item_Type = sheet.cell(r, 16).value.strip()
        Path = sheet.cell(r, 17).value.strip()

        values = (
        Valid_or_not, DCI_Number, HA_DM_Number, Harness_Number[-4:], HA_DM_Name, Basic_Number, Configuration_No,
        HA_Version, Effectivity, GH_compared, Delivery_Date, Released_Date, Rejected_Date, Comments, ECP_Number,
        IDEAL_Status, Item_Type, Path)
        values = ['NULL' if x == '' else x for x in values]
        values = ['NULL' if x == 0 else x for x in values]
        print(tuple(values))
        cur.execute(query, tuple(values))
        conn.commit()
    cur.close()
    conn.close()
    pass

def input_HI(wbname):
    import xlrd
    import pymysql
    book = xlrd.open_workbook(wbname + '.xls')
    sheet = book.sheets()[0]
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    tbdrop = 'drop table if exists ' + wbname
    cur.execute(tbdrop)
    conn.commit()
    tbcreate = 'create table ' + wbname + '(id INT(11) NOT NULL AUTO_INCREMENT,\
    `Valid_or_not` VARCHAR(50) NULL DEFAULT NULL,\
    `DCI_Number` VARCHAR(30) NULL DEFAULT NULL,\
    `HI_DM_Number` VARCHAR(30) NULL DEFAULT NULL,\
    `EDZ_Number` VARCHAR(30) NULL DEFAULT NULL,\
    `HI_DM_Name` VARCHAR(30) NULL DEFAULT NULL,\
    `Basic_Number` VARCHAR(20) NULL DEFAULT NULL,\
    `Configuration_No` VARCHAR(10) NULL DEFAULT NULL,\
    `HI_Version` VARCHAR(10) NULL DEFAULT NULL,\
    `Effectivity` VARCHAR(50) NULL DEFAULT NULL,\
    `Delivery_Date` VARCHAR(50) NULL DEFAULT NULL,\
    `Released_Date` VARCHAR(50) NULL DEFAULT NULL,\
    `Rejected_Date` VARCHAR(50) NULL DEFAULT NULL,\
    `Comments` VARCHAR(1000) NULL DEFAULT NULL,\
    `ECP_Number` VARCHAR(50) NULL DEFAULT NULL,\
    `IDEAL_Status` VARCHAR(50) NULL DEFAULT NULL,\
    `Item_Type` VARCHAR(50) NULL DEFAULT NULL,\
    `Path` VARCHAR(50) NULL DEFAULT NULL,\
    PRIMARY KEY (`id`))\
    COLLATE="utf8_general_ci"'
    cur.execute(tbcreate)
    conn.commit()
    query = 'insert into ' + wbname + '(Valid_or_not, DCI_Number, HI_DM_Number, EDZ_Number, HI_DM_Name, Basic_Number, Configuration_No, HI_Version, Effectivity, Delivery_Date, Released_Date, Rejected_Date, Comments, ECP_Number, IDEAL_Status, Item_Type, Path) values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'
    for r in range(1, sheet.nrows):
        Valid_or_not = sheet.cell(r, 0).value.strip()
        DCI_Number = sheet.cell(r, 1).value.strip()
        HI_DM_Number = sheet.cell(r, 2).value.strip()
        HI_DM_Name = sheet.cell(r, 3).value.strip()
        Basic_Number = sheet.cell(r, 4).value.strip()
        Configuration_No = sheet.cell(r, 5).value.strip()
        HI_Version = sheet.cell(r, 6).value.strip()
        Effectivity = sheet.cell(r, 7).value
        Delivery_Date = sheet.cell(r, 8).value
        Released_Date = sheet.cell(r, 9).value
        Rejected_Date = sheet.cell(r, 10).value
        Comments = sheet.cell(r, 11).value.strip()
        ECP_Number = sheet.cell(r, 12).value.strip()
        IDEAL_Status = sheet.cell(r, 13).value.strip()
        Item_Type = sheet.cell(r, 14).value
        Path = sheet.cell(r, 15).value.strip()

        values = (
        Valid_or_not, DCI_Number, HI_DM_Number, HI_DM_Number[2:4] + HI_DM_Number[5], HI_DM_Name, Basic_Number, Configuration_No,
        HI_Version, Effectivity, Delivery_Date, Released_Date, Rejected_Date, Comments, ECP_Number,
        IDEAL_Status, Item_Type, Path)
        values = ['NULL' if x == '' else x for x in values]
        values = ['NULL' if x == 0 else x for x in values]
        print(tuple(values))
        cur.execute(query, tuple(values))
        conn.commit()
    cur.close()
    conn.close()
    pass

def input_ImpactList(wbname):
    import xlrd
    import pymysql
    book = xlrd.open_workbook(wbname + '.xls')
    sheet = book.sheets()[0]
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    tbdrop = 'drop table if exists ' + wbname
    cur.execute(tbdrop)
    conn.commit()
    tbcreate = 'create table ' + wbname + '(id INT(11) NOT NULL AUTO_INCREMENT,\
    `Change_Number` VARCHAR(50) NULL DEFAULT NULL,\
    `ImpactedItem` VARCHAR(50) NULL DEFAULT NULL,\
    `Change_Action` VARCHAR(50) NULL DEFAULT NULL,\
    `EDZ` VARCHAR(50) NULL DEFAULT NULL,\
    `Type` VARCHAR(50) NULL DEFAULT NULL,\
    `Harness` VARCHAR(50) NULL DEFAULT NULL,\
    `ECP_Num` VARCHAR(50) NULL DEFAULT NULL,\
    PRIMARY KEY (`id`))\
    COLLATE="utf8_general_ci"'
    cur.execute(tbcreate)
    conn.commit()
    query = 'insert into ' + wbname + '(Change_Number, ImpactedItem, Change_Action, EDZ, Type, Harness, ECP_Num) values (%s, %s, %s, %s, %s, %s, %s)'
    for r in range(1, sheet.nrows):
        Change_Number = sheet.cell(r, 0).value.strip()
        ImpactedItem = sheet.cell(r, 1).value.strip()
        Change_Action = sheet.cell(r, 2).value.strip()
        EDZ = sheet.cell(r, 3).value
        Type = sheet.cell(r, 4).value.strip()
        Harness = sheet.cell(r, 5).value
        ECP_Num = sheet.cell(r, 6).value.strip()

        values = (Change_Number, ImpactedItem, Change_Action, EDZ, Type, Harness, ECP_Num)
        values = ['NULL' if x == '' else x for x in values]
        values = ['NULL' if x == 0 else x for x in values]
        print(tuple(values))
        cur.execute(query, tuple(values))
        conn.commit()
    cur.close()
    conn.close()
    pass

def input_IDEAL(wbname):
    import xlrd
    import pymysql
    book = xlrd.open_workbook(wbname + '.xls')
    sheet = book.sheet_by_name('1')
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    tbdrop = 'drop table if exists ' + wbname
    cur.execute(tbdrop)
    conn.commit()
    tbcreate = 'create table ' + wbname + '(序号 INT(11) NOT NULL,\
    `主部段` VARCHAR(100) NULL DEFAULT NULL,\
    `部段` VARCHAR(50) NULL DEFAULT NULL,\
    `ATA章` VARCHAR(50) NULL DEFAULT NULL,\
    `ATA系统` VARCHAR(100) NULL DEFAULT NULL,\
    `DCI编号` VARCHAR(100) NULL DEFAULT NULL,\
    `DCI名称` VARCHAR(100) NULL DEFAULT NULL,\
    `架次有效性` VARCHAR(50) NULL DEFAULT NULL,\
    `DDM编号` VARCHAR(50) NULL DEFAULT NULL,\
    `DDM名称` VARCHAR(100) NULL DEFAULT NULL,\
    `DDM版本` VARCHAR(50) NULL DEFAULT NULL,\
    `当前状态` VARCHAR(50) NULL DEFAULT NULL,\
    `对称件` VARCHAR(50) NULL DEFAULT NULL,\
    `类型` VARCHAR(50) NULL DEFAULT NULL,\
    PRIMARY KEY (`序号`))\
    COLLATE="utf8_general_ci"'
    cur.execute(tbcreate)
    conn.commit()
    query = 'insert into ' + wbname + '(序号, 主部段, 部段, ATA章, ATA系统, DCI编号, DCI名称, 架次有效性, DDM编号, DDM名称, DDM版本, 当前状态, 对称件, 类型) values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'
    for r in range(1, sheet.nrows):
        序号 = sheet.cell(r, 0).value
        主部段 = sheet.cell(r, 1).value
        部段 = sheet.cell(r, 2).value
        ATA章 = sheet.cell(r, 3).value
        ATA系统 = sheet.cell(r, 4).value
        DCI编号 = sheet.cell(r, 5).value
        DCI名称 = sheet.cell(r, 6).value
        架次有效性 = sheet.cell(r, 7).value
        DDM编号 = sheet.cell(r, 8).value
        DDM名称 = sheet.cell(r, 9).value
        DDM版本 = sheet.cell(r, 10).value
        当前状态 = sheet.cell(r, 11).value
        对称件 = sheet.cell(r, 12).value
        类型 = sheet.cell(r, 13).value

        values = (
            序号, 主部段, 部段, ATA章, ATA系统, DCI编号, DCI名称,
            架次有效性, DDM编号, DDM名称, DDM版本, 当前状态, 对称件, 类型)
        values = ['NULL' if x == '' else x for x in values]
        values = ['NULL' if x == 0 else x for x in values]
        # values2 = ['NULL' if x == '' else x for x in values]
        print(tuple(values))
        cur.execute(query, tuple(values))
        conn.commit()
    cur.close()
    conn.close()
    pass

def ha_cfg_extract(tablename, harnessnumber):
#extract provided harness configuration_No
    import pymysql
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    cur.execute('select Configuration_No from ' + tablename + ' where Valid_or_not="YES" and harness_number=' + '"' + harnessnumber + '"')
    return cur.fetchall()

def ha_cfg_extract_implement(tablename, harnessnumber):
#extract provided harness configuration_No which current implment table has used
    import pymysql
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    basicnumber = '88' + harnessnumber[0] + '0C' + harnessnumber[-3:] + '00G'
    query = 'select partnumber_after from ' + tablename + ' where partnumber_after like ' + '"%' + basicnumber + '%"'
    cur.execute(query)
    return cur.fetchall()

def get_ha_orglst(tablename, harnessnumber):
# extract a list of HA Class according to the provided orginal tablename and harnessnumber
    from ha import HA
    import pymysql
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    cur.execute('select * from ' + tablename + ' where Valid_or_not="YES" and harness_number=' + '"' + harnessnumber + '"')
    return [HA(row[3], row[4], row[8], row[9]) for row in cur.fetchall()]

def get_iptha(tablename, harnessnumber):
# extract a list of HA Class according to the provided impact tablename and harnessnumber
    from ha import HA
    import pymysql
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    cur.execute('select * from ' + tablename + ' where harness=' + '"' + harnessnumber + '"')
    f = cur.fetchone()
    return HA(f[2], f[6], '', f[8])

def get_impact_harnessnumber_lst(impacttable):
# extract a list[] of harnessnumber with provided impact list table name
    import pymysql
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    cur.execute('select * from ' + impacttable + ' where type="HA"')
    return [row[6] for row in cur.fetchall()]

def create_implement_table(tablename):
    import pymysql
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    tbdrop = 'drop table if exists ' + tablename
    cur.execute(tbdrop)
    tbcreate = 'create table ' + tablename + '(id INT(11) NOT NULL AUTO_INCREMENT,\
    `Partnumber_before` VARCHAR(50) NULL DEFAULT NULL,\
    `Version_before` VARCHAR(50) NULL DEFAULT NULL,\
    `Effectivity_before` VARCHAR(50) NULL DEFAULT NULL,\
    `Partnumber_after` VARCHAR(50) NULL DEFAULT NULL,\
    `Version_after` VARCHAR(50) NULL DEFAULT NULL,\
    `Effectivity_after` VARCHAR(50) NULL DEFAULT NULL,\
    PRIMARY KEY (`id`))\
    COLLATE="utf8_general_ci"'
    cur.execute(tbcreate)
    cur.close
    conn.close
    return

def input_data_to_implement_table(implmenttable, HAbefore, HAafter):
    import pymysql
    conn = pymysql.connect(host='localhost', user='root', passwd='init#201605', db='cc', charset='utf8')
    cur = conn.cursor()
    rowinsert = 'insert into ' + implmenttable + '(Partnumber_before, Version_before, Effectivity_before, Partnumber_after, Version_after, Effectivity_after) values (%s,%s,%s,%s,%s,%s)'
    values = (HAbefore.partnumber, HAbefore.version, HAbefore.effectivity, HAafter.partnumber, HAafter.version, HAafter.effectivity)
    cur.execute(rowinsert, tuple(values))
    conn.commit()

