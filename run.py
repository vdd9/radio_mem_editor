#!/usr/bin/python
from dr1801 import *

def main():
    radio = DR1801(open('DM-1801A6 Factory setting.accps', 'r+b'))
    radio.importxls('freqs.xls')
    for chan in radio.channels:
        print(chan.id, chan.name, "=")
        print(repr(chan))
    radio.save()
    radio.writexls('freqs.xls')
    
if __name__ == "__main__":
        main()