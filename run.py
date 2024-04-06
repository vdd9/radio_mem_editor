#!/usr/bin/python
from dr1801 import *

def main():
    radio = DR1801(open('DM-1801A6_Factory_setting.accps', 'r+b'))
    radio.importxlsx()
    radio.save()
    # radio.writexlsx()
    
if __name__ == "__main__":
        main()