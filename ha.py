class HA(object):
    def __init__(self, partnumber, harnessnumber, version, effectivity):
        self.partnumber = partnumber
        self.harnessnumber = harnessnumber
        self.basicnumber = partnumber[:10]
        self.cfgnumber = partnumber[-2:]
        self.version = version
        self.effectivity = effectivity

    def revise(self):
        version = chr(ord(self.version) + 1)
        return HA(self.partnumber, self.harnessnumber, version, self.effectivity)

    def cfg_after_cut(self, effectivitychg):
        import cfg
        effectivity = cfg.cfgold_to_cfgnew(self.effectivity, effectivitychg)
        return HA(self.partnumber, self.harnessnumber, self.version, effectivity)

    def new_cfg_after_cut(self, fromtable, implementtable, effectivitychg):
        import cfg
        effectivity = effectivitychg
        partnumber = self.basicnumber + 'G' + cfg.get_newcfg_number(fromtable, implementtable, self.harnessnumber)
        return HA(partnumber, self.harnessnumber, 'A', effectivity)

    def print_as_list(self):
        lst = [self.partnumber, self.harnessnumber, self.version, self.effectivity]
        return lst

