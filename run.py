from get.scrap import Scrap
from get.constants import *

inst = Scrap()

inst.land_first_page()

loop = True
while loop:
    command = input('Enter ok for save data, x to close: ')
    if command == 'ok':
        file_name = input('Enter File Name: ')
        inst.collect_data(file_name)
    if command == 'x':
        loop = False



