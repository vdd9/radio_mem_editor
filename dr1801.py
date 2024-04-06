#!/usr/bin/python

import re
from common import *

# note:
# stangre values after 0x0001C6C0
# encryption 0x0001D7E0
# VFO sttings 0x0001DD00

class Zone:
    def __init__(self, id, block, offset, _):
        block.ushort[offset+0x24] = id
        self.block = block
        self.offset = offset
        self.val_of_ = ByteValueHandler(block, offset, {
            'namesize':    0x20,
            'channels_nbr': 0x22
        })
            
    @property
    def channels_IDs(self):
        return [ self.block.ushort[self.offset+0x28+2*i] for i in range(0,self.val_of_['channels_nbr']) ]
    @channels_IDs.setter
    def channels_IDs(self, new_list):
        self.val_of_["channels_nbr"] = len(new_list)
        for i in range(0,len(new_list)):
            self.block.ushort[self.offset+0x28+2*i] = new_list[i]

    @property
    def name(self):
        return estract_string_til_zero(self.block[self.offset:], 0x20)
    @name.setter
    def name(self, new_name):
        self.val_of_['namesize'] = len(new_name[:0x20])
        self.block[self.offset:self.offset+0x20] = bytearray(new_name[:0x20].ljust(0x20,'\x00'), 'utf-8')
    
    def __str__(self):
        return f'{self.name}: {self.channels_IDs}'

class ScanList:
    def __init__(self, id, block, offset, _):
        block[offset]= id+1
        self.block = block
        self.offset = offset
        self.val_of_ = ByteValueHandler(block, offset, {
            'channels_nbr': 0x01,
            'priority1': 0x02, # 0 node, 1 fixed, 2 selected
            'priority2': 0x03, # 0 node, 1 fixed, 2 selected
            'txreply': 0x08, # 0 node, 1 fixed, 2 selected
            'residence': 0x0C, # (Byte*0.5 == DisplayValue) 0.0 to 10.0
            'txresidence': 0x0D, # (Byte*0.5 == DisplayValue) 0.0 to 10.0
        })
        
    ## ID ##
        
    @property
    def id(self):
        return self.block[self.offset]-1
    
    ## Channels ##

    @property
    def channels_IDs(self):
        return [ self.block.ushort[self.offset+0x30+2*i] for i in range(0,self.val_of_['channels_nbr']) ]
    @channels_IDs.setter
    def channels_IDs(self, new_list):
        self.val_of_["channels_nbr"] = len(new_list)
        for i in range(0,len(new_list)):
            self.block.ushort[self.offset+0x30+2*i] = new_list[i]

    ## Priorities ##

    @property
    def priority1_channel(self):
        return self.block.ushort[self.offset+0x04]
    @priority1_channel.setter
    def priority1_channel(self, value):
        self.block.ushort[self.offset+0x04] = value
    @property
    def priority2_channel(self):
        return self.block.ushort[self.offset+0x06]
    @priority2_channel.setter
    def priority2_channel(self, value):
        self.block.ushort[self.offset+0x06] = value
    @property
    def txreply_channel(self):
        return self.block.ushort[self.offset+0x0A]
    @txreply_channel.setter
    def txreply_channel(self, value):
        self.block.ushort[self.offset+0x0A] = value

    ## Name ##

    @property
    def name(self):
        return estract_string_til_zero(self.block[self.offset+0x10:], 0x20)
    @name.setter
    def name(self, new_name):
        self.block[self.offset+0x10:self.offset+0x30] = bytearray(new_name[:0x20].ljust(0x20,'\x00'), 'utf-8')
    
    def __str__(self):
        return f'{self.name}: {self.channels_IDs}'

class Channel:
    channels_size = 0x34
    def __init__(self, id, block, offset, offset_name):
        block[offset]= id+1
        self.block = block
        self.offset = offset
        self.offset_name = offset_name
        self.is_ = BoolHandler(block, offset, {
            'high':     0x03,
            'DTMF':     0x16,
            'wide':     0x19,
            'autoscan': 0x1A,
            'lonework': 0x2F
        },{
            'dmr': (0x02,0x01,0x03)
        },{
            'talkaround':     (0x15,0x80),
            'directTDMA':     (0x15,0x02),
            'privateconfirm': (0x15,0x01)
        })
        self.val_of_ = ByteValueHandler(block, offset, {
            'talkgroup':    0x0C,
            'PTTstate':     0x0E,
            'colorcode':    0x10,
            'slot':         0x11,
            'encryption':   0x14,
            'alarm':        0x18,
            'autoscanlist': 0x1B,
            'PTTID':        0x28,
            'receivegroup': 0x2A
        })

    ## clear ##

    def clear(self):
        for k in self.is_.keys():
            self.is_[k] = False
        for k in self.val_of_.keys():
            self.val_of_[k] = 0
        self.txDcs = ''
        self.rxDcs = ''

    ## ID ##
        
    @property
    def id(self):
        return self.block[self.offset]

    ## NAME ##

    @property
    def name(self):
        return estract_string_til_zero(self.block[self.offset_name:], 0x14)
    @name.setter
    def name(self, new_name):
        self.block[self.offset_name:self.offset_name+0x14] = bytearray(new_name[:0x14].ljust(0x14,'\x00'), 'utf-8')

    ## DSIPLAY ##

    def __str__(self):
        return f'{self.name}  '+('~','')[self.is_['dmr']]+'('+('low','high')[self.is_['high']]+('-narow','-wide')[self.is_['wide']]+f')\n RX: {(self.rxFreq):0>9.5f}  {self.rxDcs}\n TX: {(self.txFreq):0>9.5f}  {self.txDcs}'
    def __repr__(self):
        return self.name.ljust(12," ")+" ".join([str(b).zfill(3) for b in self.block[self.offset:self.offset+self.channels_size]])\
        + "\n            ID "+str(self.id).ljust(4," ")+"|"+("AMA","DMR")[self.is_['dmr']]+"|"+("Low","Hi ")[self.is_['high']]+"| RX: "+f'{(self.rxFreq):0>9.5f} | TX: {(self.txFreq):0>9.5f}'+" |TG |   |PTT|   |Col|Slt|       |Enc| & |DTM|   |Sys|"\
        +("NFM","FM ")[self.is_['wide']]+"|autosca|R: "+self.rxDcs.ljust(12)+"|T: "+self.txDcs.ljust(12)+"|               |PID|   |RG |               |lon|"
    def dump(self):
        print("========")
        print(self.name)
        print("RX:", self.rxFreq, self.rxDcs)
        print("TX:", self.txFreq, self.txDcs)
        for k in self.is_.keys():
            print(f'{k}:', self.is_[k])
        for k in self.val_of_.keys():
            print(f'{k}:', str(self.val_of_[str(k)]))

    ## FREQENCY ##

    @property
    def txFreq(self):
        return self._freq(8)
    @txFreq.setter
    def txFreq(self, new_freq):
        self._set_freq(8, new_freq)

    @property
    def rxFreq(self):
        return self._freq(4)
    @rxFreq.setter
    def rxFreq(self, new_freq):
        self._set_freq(4, new_freq)

    def _freq(self, idx):
        return self.block.uint[self.offset+idx]/1000000
    def _set_freq(self, idx, new_freq):
        self.block.uint[self.offset+idx] = int(new_freq*1000000)

    ## DCS ##

    @property
    def txDcs(self):
        return self._dcs(32)
    @txDcs.setter
    def txDcs(self, new_dcs):
        self._set_dcs(32, new_dcs)

    @property
    def rxDcs(self):
        return self._dcs(28)
    @rxDcs.setter
    def rxDcs(self, new_dcs):
        self._set_dcs(28, new_dcs)
    
    def _dcs(self, idx):
        idx += self.offset
        if self.block[idx+2] == 0:
            return ''
        elif self.block[idx+2] == 1:
            return "CTCSS " + str(self.block.ushort[idx]/10) + "Hz"
        elif self.block[idx+2] == 2:
            return "DCS " + str(self.block.ushort[idx]) + ("N","I")[self.block[idx+3]]
        return "WTF"
    def _set_dcs(self, idx, new_dcs):
        idx += self.offset
        if new_dcs and len(new_dcs)>0:
            if new_dcs.startswith("D"):
                self.block[idx+2] = 2
                self.block.ushort[idx] = int(re.findall(r"(\d+)", new_dcs)[0])
                self.block[idx+3] = new_dcs.endswith("I")
            else:
                self.block[idx+2] = 1
                self.block.ushort[idx] = int(float(re.findall(r"(?:\d*\.*\d+)", new_dcs)[0])*10)
        else:
            write_int_little_indian(self.block, idx, 0)

class DR1801():
    def __init__(self,file):
        self.file = file
        self.ba = bytearray(self.file.read())
        self.channels = binaryList(Channel,self.ba,0xA660,0x34,0xA65C,0x00017660,0x14)
        self.scanlists = binaryList(ScanList,self.ba,0xA33C,0x50,0xA338)
        self.zones = binaryList(Zone,self.ba,0x0420,0x68,0x0418)

    def importxlsx(self, path='freqs.xlsx'):
        book = openpyxl.load_workbook(path)
        # Parse Channels
        freqs_sheet = book.worksheets[0]
        chan_ix = 0
        previous_freq=0.0
        self.channels.length = freqs_sheet.max_row
        for row in freqs_sheet.iter_rows():
            current_chan = self.channels[chan_ix]
            current_chan.clear()
            current_chan.name = row[0].value
            freq = str(row[1].value)
            if freq.startswith('+') or freq.startswith('-'):
                current_chan.txFreq = previous_freq + float(freq)
            else:
                current_chan.txFreq = float(freq)
            freq = row[2].value
            if freq:
                current_chan.rxFreq = current_chan.txFreq + float(freq)
            else:
                current_chan.rxFreq = current_chan.txFreq
            if row[3].value:
                current_chan.txDcs = str(row[3].value)
            if row[4].value:
                current_chan.rxDcs = str(row[4].value)
            if row[5].value:
                for boolname in row[5].value.split('/'):
                    current_chan.is_[boolname] = True
            if row[6].value:
                for val in row[6].value.split('/'):
                    k, v = val.split(":")
                    current_chan.val_of_[k] = int(v)
            previous_freq=current_chan.txFreq
            chan_ix +=1
        # Parse Zones
        zones_sheet = book.worksheets[1]
        self.zones.length = zones_sheet.max_column
        for cx in range(1,zones_sheet.max_column+1):
            current_zone = self.zones[cx-1]
            current_zone.name = zones_sheet[openpyxl.utils.get_column_letter(cx)+'1'].value
            current_zone.channels_IDs = [int(c[0].value[8:])-1 for c in zones_sheet[openpyxl.utils.get_column_letter(cx)+'2:'+openpyxl.utils.get_column_letter(cx)+'33'] if c[0].value]
        # Parse Scanlists
        scanlists_sheet = book.worksheets[2]
        self.scanlists.length = scanlists_sheet.max_column
        for cx in range(1,scanlists_sheet.max_column+1):
            current_scanlist = self.scanlists[cx-1]
            current_scanlist.name = scanlists_sheet[openpyxl.utils.get_column_letter(cx)+'1'].value
            scan_properties = scanlists_sheet[openpyxl.utils.get_column_letter(cx)+'2:'+openpyxl.utils.get_column_letter(cx)+'9']
            current_scanlist.val_of_['priority1'] = ('fixed','channel','selected').index(scan_properties[0][0].value)
            current_scanlist.priority1_channel = int(scan_properties[1][0].value[8:])-1
            current_scanlist.val_of_['priority2'] = ('fixed','channel','selected').index(scan_properties[2][0].value)
            current_scanlist.priority2_channel = int(scan_properties[3][0].value[8:])-1
            current_scanlist.val_of_['txreply'] = ('fixed','channel','selected').index(scan_properties[4][0].value)
            current_scanlist.txreply_channel = int(scan_properties[5][0].value[8:])-1
            current_scanlist.val_of_["residence"] = int(float(scan_properties[6][0].value)/0.5)
            current_scanlist.val_of_["txresidence"] = int(float(scan_properties[7][0].value)/0.5)
            current_scanlist.channels_IDs = [int(c[0].value[8:])-1 for c in scanlists_sheet[openpyxl.utils.get_column_letter(cx)+'10:'+openpyxl.utils.get_column_letter(cx)+'25'] if c[0].value]
            current_scanlist.channels_IDs = current_scanlist.channels_IDs

    def save(self):
        self.file.seek(0)
        self.file.write(self.ba)

    def writexlsx(self, path='freqs.xlsx'):
        book = openpyxl.Workbook()
        # Export Channels
        freqs_sheet = book.active
        freqs_sheet.title = "chan"
        previous_freq = 0.0
        row_id = 1
        for chan in self.channels:
            freqs_sheet.column_dimensions[openpyxl.utils.get_column_letter(row_id)].width = 15
            freqs_sheet.cellInput(row_id, 1, chan.name)
            if chan.txFreq - previous_freq > 0 and chan.txFreq - previous_freq < 0.06:
                freqs_sheet.cellInput(row_id, 2, f'+{(chan.txFreq-previous_freq):.5f}', 'Could be the shift from previous Channel.')
            else:
                freqs_sheet.cellInput(row_id, 2, f'{chan.txFreq:0>9.5f}', 'Could be the shift from previous Channel.')
            if chan.rxFreq != chan.txFreq:
                freqs_sheet.cellInput(row_id, 3, f'{(chan.rxFreq-chan.txFreq):+.5f}', 'TX shift.\nlet empty if same than RX.')
            freqs_sheet.cellInput(row_id, 4, f'{chan.txDcs}', 'Starts with \'D\'=> DCS\n + ends with \'I\'=> Inverted.\nElse => Ctcss.')
            freqs_sheet.cellInput(row_id, 5, f'{chan.rxDcs}', 'Examples:\nCtcss 67.0 Hz / DCS 446 I / Tone 67.0 / T67.0 / D023 / D023I')
            freqs_sheet.cellInput(row_id, 6, '/'.join([k for k,v in chan.is_.items() if v == True]), '/'.join(chan.is_.keys()))
            freqs_sheet.cellInput(row_id, 7, '/'.join([f'{k}:{v}' for k,v in chan.val_of_.items() if v != 0]), '/'.join(chan.val_of_.keys()))
            previous_freq=chan.txFreq
            row_id += 1
        # Export Zones
        zones_sheet = book.create_sheet("zone")
        zone_idx = 1
        for zone in self.zones:
            zones_sheet.column_dimensions[openpyxl.utils.get_column_letter(zone_idx)].width = 15
            zones_sheet.cellInput(1, zone_idx, zone.name)
            row_idx = 2
            for idx in zone.channels_IDs:
                zones_sheet.cellInput(row_idx, zone_idx, "=chan!$A"+str(idx+1), 'Must refers to the cell like:\n=chan!$A1')
                row_idx += 1
            zone_idx += 1
        # Export ScanLists
        scanlists_sheet = book.create_sheet("scan")
        scanlists_idx = 1
        prio = ('fixed','channel','selected')
        for scanlist in self.scanlists:
            scanlists_sheet.column_dimensions[openpyxl.utils.get_column_letter(scanlists_idx)].width = 15
            scanlists_sheet.cellInput(1, scanlists_idx, scanlist.name)
            scanlists_sheet.cellInput(2, scanlists_idx, prio[scanlist.val_of_['priority1']],'priority 1:\nfixed\nchannel\nselected')
            scanlists_sheet.cellInput(3, scanlists_idx, "=chan!$A"+str(scanlist.priority1_channel+1))
            scanlists_sheet.cellInput(4, scanlists_idx, prio[scanlist.val_of_['priority2']],'priority 2:\nfixed\nchannel\nselected')
            scanlists_sheet.cellInput(5, scanlists_idx, "=chan!$A"+str(scanlist.priority2_channel+1))
            scanlists_sheet.cellInput(6, scanlists_idx, prio[scanlist.val_of_['txreply']],'TX Reply:\nfixed\nchannel\nselected')
            scanlists_sheet.cellInput(7, scanlists_idx, "=chan!$A"+str(scanlist.txreply_channel+1))
            scanlists_sheet.cellInput(8, scanlists_idx, f'{scanlist.val_of_["residence"]*0.5:.1f}', 'Rx Residence Time')
            scanlists_sheet.cellInput(9, scanlists_idx, f'{scanlist.val_of_["txresidence"]*0.5:.1f}', 'Tx Residence Time')
            row_idx = 10
            for idx in scanlist.channels_IDs:
                scanlists_sheet.cellInput(row_idx, scanlists_idx, "=chan!$A"+str(idx+1), 'Must refers to the cell like:\n=chan!$A1')
                row_idx += 1
            scanlists_idx += 1
        book.save(path)

    def __del__(self):
        self.file.close()
