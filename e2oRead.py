#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
from subprocess import check_output

#--------------------------------------------------------------------
def part_of_speech(pos):
    if pos == 'v':
        return 'କ୍ରିୟା'
    if pos == 'adv':
        return 'କ୍ରିୟା-ବିଶେଷଣ'
    if pos == 'n':
        return 'ବିଶେଷ୍ୟ'
    if pos == 'a':
        return 'ବିଶେଷଣ'
    if pos == '':
        return 'ଶବ୍ଦାର୍ଥ'

#----------------------------------------------------------------------
def write_file(content):
    content = content.strip() + '\n\n'
    file_name = 'out.txt'
    text_file = open(file_name, "a")
    text_file.write(content)
    text_file.close()


#----------------------------------------------------------------------
def open_file(path, starting_num, upto_lim):
    book = xlrd.open_workbook(path)
 
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)

    r = starting_num #row index
    c = 0 # column index
    while r != upto_lim+1:
        # read a row
        #print r,'th row'
        #print first_sheet.row_values(r)

        # read a cell
        # cell = first_sheet.cell(r,c)
        #print first_sheet.cell(r,c+1).value, first_sheet.cell(r,c+3).value

        # pronounciation
        us_pron = check_output(["espeak", "-q", "--ipa",
                                    '-v', 'en-us',
                                    first_sheet.cell(r,c+1).value], shell=True)
        uk_pron = check_output(["espeak", "-q", "--ipa",
                                    '-v', 'gd',
                                    first_sheet.cell(r,c+1).value], shell=True)
    
        if first_sheet.cell(r,c+4).value.encode('utf-8').strip() != '':
            mean2 = '\n# ' + first_sheet.cell(r,c+4).value.encode('utf-8')
        else:
            mean2 = ''
        if first_sheet.cell(r,c+5).value.encode('utf-8') != '':
            mean3 = '\n# ' + first_sheet.cell(r,c+5).value.encode('utf-8')
        else:
            mean3 = ''
            
        # wikt format
        data = "'''" + first_sheet.cell(r,c+1).value.encode('ascii') + "'''\n" + '== ଇଂରାଜୀ ==\n' + '=== ଉଚ୍ଚାରଣ ===\n{{ଇଂରାଜୀ ଉଚ୍ଚାରଣ|' + us_pron.strip() + '|' + uk_pron.strip() + '}}\n=== ' + str(part_of_speech(first_sheet.cell(r,c+2).value)) + ' ===\n{{ଇଂରାଜୀ ' + str(part_of_speech(first_sheet.cell(r,c+2).value)) + '|' + first_sheet.cell(r,c+1).value.encode('ascii') + '}}\n# ' + first_sheet.cell(r,c+3).value.encode('utf-8') + mean2 + mean3 + '\n\n[[ଶ୍ରେଣୀ:ଇଂରାଜୀ ' + str(part_of_speech(first_sheet.cell(r,c+2).value)) + ']]\n[[en:' + first_sheet.cell(r,c+1).value.encode('ascii') + ']]'
        #data = data.decode('utf-8')

        write_file(data)

        r+=1

#----------------------------------------------------------------------
if __name__ == "__main__":
    path = 'e2o.xlsx'
    start_index = input('enter the starting value: ')
    end_index = input('enter the end index: ')
    if start_index - end_index <=50:
        limit = end_index - start_index
    else:
        limit = 50 # set the limit
    print 'writting the data...'
    open_file(path, start_index, limit)


