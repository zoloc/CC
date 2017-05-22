# coding: utf-8
def cfgpiecelst_to_cfgpiecestr(cfgpiecelst):
# convert a continuous configuration list like:[10101,10102,10103] to configuration scope as:'10101-10103'
    string = list(map(str, cfgpiecelst))
    if len(string) > 2:
        string = string[0], '-', string[-1]
        return ''.join(string)
    else:
        return ','.join(string)


def cfglst_to_cfgpiecelst(cfglst):
# convert a discontinuous configuration list like:[10112, 10101,10102,10103,10105,10106,10108,10109]
# to configuration list with different small continuous configuration list as [[10101,10102,10103],[10105,10106],[10108,10109],[10112]]
    from itertools import groupby
    cfgpiecelst = []
    for k, g in groupby(enumerate(sorted(cfglst)), lambda x: x[1] - x[0]):
        cfgpiece = [v for i, v in g]
        cfgpiecelst.append(cfgpiece)
    return cfgpiecelst


def cfgset_to_cfgstr(cfgset):
# convert a configuration set like:{10101,10102,10103,10105,10106,10107,10108,10109}
# to configuration scope as:'10101-10103,10105-10109'
    import cfg
    cfglist = sorted(cfg.cfglst_to_cfgpiecelst(list(cfgset)))
    return ','.join(map(cfg.cfgpiecelst_to_cfgpiecestr, cfglist))


def cfgstr_to_cfglst(cfgstr):
# convert a configuration scope like: '10101-10103,10106-10109' with space is acceptable
# to configuration set as: {10101,10102,10103,10105,10106,10107,10108,10109}
    try:
        removecomma = cfgstr.strip().split(',')
        cfglst = []
        for each_item in removecomma:
            cfglst.extend(list(range(int(each_item.split('-')[0]), int(each_item.split('-')[-1]) + 1)))
        return cfglst
    except ValueError as e:
        print('configuratin input error', e)

def cfgold_to_cfgnew(cfgold, cfgchg):
    import cfg
#get result once new configuration implemented
#***Notice: The result of revise and no impact configuration is same, maybe need modify***
    if set(cfg.cfgstr_to_cfglst(cfgold)) ^ set(cfg.cfgstr_to_cfglst(cfgchg)) == set():
        return cfg.cfgset_to_cfgstr(cfgold)
    else:
        return cfg.cfgset_to_cfgstr(set(cfg.cfgstr_to_cfglst(cfgold)) - set(cfg.cfgstr_to_cfglst(cfgchg)))

def get_newcfg_number(hatable, implementtable, harnessnumber):
#get a new number based on current exist GXX for provided harness
    import db
    a = db.ha_cfg_extract(hatable, harnessnumber)
    currentcfgnumberlst = []
    for (row,) in a:
        currentcfgnumberlst.append(int(row)) #extract current used cfgnumber in ha table with a list
    b = db.ha_cfg_extract_implement(implementtable, harnessnumber)
    for (row,) in b:
        currentcfgnumberlst.append(int(row[-2:]))  #extract current used cfgnumber in implment table with a list
    cfgnumberlst = list(range(20,100)) #need update in future
    return str(sorted(list(set(cfgnumberlst) - set(currentcfgnumberlst)))[0])

