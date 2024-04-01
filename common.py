#!/usr/bin/python

import struct
import openpyxl

def estract_string_til_zero(block, limit):
    i = 0
    while i < limit and block[i] != 0:
        i+=1
    string = str(block[:i], "utf-8")
    return string

def read_short_little_indian(block, offset):
    return struct.unpack('<H',block[offset:offset+2])[0]

def write_short_little_indian(block, offset, val):
    struct.pack_into('<H',block,offset,val)

def read_int_little_indian(block, offset):
    return struct.unpack('<I',block[offset:offset+4])[0]

def write_int_little_indian(block, offset, val):
    struct.pack_into('<I',block,offset,val)


## Add new methode to excel worksheets
def myCellInput(self,row,column,value,comment=None):
    self.cell(row=row,column=column).value = value
    if comment:
        self.cell(row=row,column=column).comment = openpyxl.comments.Comment(comment,'')
openpyxl.worksheet.worksheet.Worksheet.cellInput = myCellInput
##

class ByteValueHandler:
    '''
    Usage example:
        val_of_ = ByteValueHandler(bytearray, offset, {
            'name':         byte_location,
            'talkgroup':    0x0C,
            'PTTID':        0x0E,
            'colorcode':    0x10,
            'slot':         0x11,
            'encryption':   0x14,
        })

        val_of_['colorcode'] = 3
        if val_of_['colorcode'] == 3:
            print('color code is 3')

    '''
    def __init__(self, block, base_offset, access_map):
        self.block = block
        self.offset = base_offset
        self.access_map = access_map
    def __getitem__(self, key):
        if isinstance(key, str):
            if key in self.access_map:
                return self.block[self.offset+self.access_map[key]]
            else:
                raise IndexError(f'Key ({key}) not present.')
        else:
            raise TypeError('Invalid argument type.')
    def __setitem__(self, key, new_val):
        if new_val < 0 or new_val > 0xFF :
            raise ValueError(f'Given number ({new_val}) don\'t fit in a single byte.')
        else:
            if key in self.access_map:
                self.block[self.offset+self.access_map[key]] = new_val
            else:
                raise IndexError(f'Key ({key}) not present.')
    def keys(self):
        return list(self.access_map.keys())
    def items(self):
        for k in self.keys():
            yield (k, self[k])

class BoolHandler:
    '''
    Usage example:
        is_ = BoolHandler(bytearray, offset, {
            'name':    byte_location,
            'high':    0x03,
            'DTMF':    0x16,
            'wide':    0x19
        },{
            'name': ( byte_location, value_if_False, value_if_True ),
            'dmr':  ( 0x02,          0x01,           0x03          )
        },{
            'name':          ( byte_location, bit_location )
            'talkaround':    ( 0x15,          0x80         ),
            'directTDMA':    ( 0x15,          0x02         )
        })

        is_['dmr'] = True
        if is_['dmr']:
            print('it is dmr')

    '''
    def __init__(self, block, base_offset, access_map, access_spec_val_map, access_spec_bit_map):
        self.block = block
        self.offset = base_offset
        self.access_map = access_map
        self.access_spec_val_map = access_spec_val_map
        self.access_spec_bit_map = access_spec_bit_map
    def __getitem__(self, key):
        if isinstance(key, str):
            if key in self.access_map:
                return bool(self.block[self.offset+self.access_map[key]])
            elif key in self.access_spec_val_map:
                return bool(self.block[self.offset+self.access_spec_val_map[key][0]] == self.access_spec_val_map[key][2])
            elif key in self.access_spec_bit_map:
                return bool(self.block[self.offset+self.access_spec_bit_map[key][0]]&self.access_spec_bit_map[key][1])
            else:
                raise IndexError(f'Key({key}) not present.')
        else:
            raise TypeError(f'Invalid argument type.')
    def __setitem__(self, key, new_bool):
        if isinstance(key, str):
            if key in self.access_map:
                self.block[self.offset+self.access_map[key]] = (0x00,0x01)[bool(new_bool)]
            elif key in self.access_spec_val_map:
                self.block[self.offset+self.access_spec_val_map[key][0]] = self.access_spec_val_map[key][1:][bool(new_bool)]
            elif key in self.access_spec_bit_map:
                if new_bool:
                    self.block[self.offset+self.access_spec_bit_map[key][0]] |= self.access_spec_bit_map[key][1]
                else:
                    self.block[self.offset+self.access_spec_bit_map[key][0]] &= ~(self.access_spec_bit_map[key][1])
            else:
                raise IndexError(f'Key({key}) not present.')
        else:
            raise TypeError(f'Invalid argument type.')
    def keys(self):
        return { **self.access_map, **self.access_spec_val_map, **self.access_spec_bit_map }.keys()
    def items(self):
        for k in self.keys():
            yield (k, self[k])

class binaryList:
    def __init__(self, ConstructorCallback, bytearray, address, size, length_address, names_address=0, names_size=0):
        self.ConstructorCallback = ConstructorCallback
        self.ba = bytearray
        self.address = address
        self.object_size = size
        self.names_address = names_address
        self.names_size = names_size
        self.length_address = length_address

    @property
    def length(self):
        return struct.unpack('<H',self.ba[self.length_address:self.length_address+2])[0]
    @length.setter
    def length(self, new_length):
        struct.pack_into('<H',self.ba,self.length_address,new_length)
    def __len__(self):
        return self.length

    def __getitem__(self, key):
        if isinstance(key, slice):
            # Get the start, stop, and step from the slice
            return [self[ii] for ii in xrange(*key.indices(len(self)))]
        elif isinstance(key, int):
            if key < 0: # Handle negative indices
                key += len(self)
            if key < 0 or key >= len(self):
                raise IndexError(f'list index ({key}) out of range.')
            start_idx = self.address + self.object_size * key
            start_name_idx = self.names_address + self.names_size * key
            return self.ConstructorCallback(key, self.ba, start_idx, start_name_idx)
        else:
            raise TypeError(f'Invalid argument type.')
    def __setitem__(self, key, chanel):
        return
    def __delitem__(self, key):
        if isinstance(key, slice):
            # Get the start, stop, and step from the slice
            for ii in xrange(*key.indices(len(self))):
                del self[ii]
        elif isinstance(key, int):
            if key < 0: # Handle negative indices
                key += len(self)
            if key < 0 or key >= len(self):
                raise IndexError(f'list index ({key}) out of range.')
            print(f'TODO zero everywhere for channel {key}')
        else:
            raise TypeError(f'Invalid argument type.')
