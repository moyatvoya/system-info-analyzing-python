import pandas as pd
import numpy as np
import glob
import re
import xlsxwriter
import codecs

def return_hostname():
    substr1,substr2,substr3 = 'Host Name:','Имя узла:','€¬п г§« :'   # Substring to search for.
    hostname = ' '                              # Hostname return variable.
    for line in mylines:                        # string to be searched
        index1, index2, index3 = 0,0,0              # current index: character being compared
        while index1 < len(line) and index2 < len(line) and index3 < len(line): # While index has not exceed string lenght,
            index1 = line.find(substr1, index1)    # set index to first occurence of 'Host Name:'
            index2 = line.find(substr2, index2)    # set index to first occurence of 'Имя узла:'
            index3 = line.find(substr3, index3)    # set index to first occurence of 'Имя узла:'
            if index1 == -1 and index2 == -1 and index3 == -1:# If nothing was found,
                break                           # exit the while loop.
            hostname = line[15:].replace(' ','')   # Append host name to hostname return variable. Replace spaces.
            index1 += len(substr1)                # Increment the index by the length of substring 'Host Name:'.
            index2 += len(substr2)                # Increment the index by the length of substring 'Имя узла:'.
            index3 += len(substr3)                # Increment the index by the length of substring 'Имя узла:'.
    return hostname

# Function is not complete
def return_ipv4():
    substr = 'Host Name:'                       # Substring to search for.
    ipv4 = ' '                                  # IPv4 return variable.
    for line in mylines:                        # string to be searched
        index = 0                               # current index: character being compared
        prev = 0                                # previous index: last character compared
        while index < len(line):                # While index has not exceed string lenght,
            index = line.find(substr, index)    # set index to first occurence of 'Host Name:'
            if index == -1:                     # If nothing was found,
                break                           # exit the while loop.
            ipv4 = line[27:]                    # Append IPv4 address to ipv4 return variable.
            index += len(substr)                # Increment the index by the length of substring 'IPv4'.
    return ipv4                                 # return IPv4

def return_osname():
    substr1,substr2,substr3 = 'OS Name:', 'Название ОС:', 'Ќ §ў ­ЁҐ Ћ‘:' # Substring to search for.
    osname = ' '                                  # OS name return variable.
    for line in mylines:                        # string to be searched
        index1, index2, index3 = 0,0,0               # current index: character being compared
        while index1 < len(line) and index2 < len(line) and index3 < len(line): # While index has not exceed string lenght,
            index1 = line.find(substr1, index1)    # set index to first occurence of 'OS Name:'
            index2 = line.find(substr2, index2)     # set index to first occurence of 'Название ОС:'
            index3 = line.find(substr3, index3)    # set index to first occurence of 'Название ОС:'
            if index1 == -1 and index2 == -1 and index3 == -1:# If nothing was found,
                break                           # exit the while loop.
            osname = line[15:].replace('  ',',')   # Append host name to hostname return variable. Remove spaces.
            osname = osname.replace(', ','')
            osname = osname.replace(',','')
            index1 += len(substr1)                # Increment the index by the length of substring 'OS Name:'.
            index2 += len(substr2)                # Increment the index by the length of substring 'OS Name:'.
            index3 += len(substr3)                # Increment the index by the length of substring 'OS Name:'.
    # New loop to find Service Pack
    for line in mylines:
        index = 0                               # current index: character being compared
        if line.find('=== Product ===') != -1: #Check if we are at the beginning of the file.
            break
        substr = "Service Pack "
        service_pack = ' '
        while index < len(line):
            index = line.find(substr, index)    # set index to first occurence of 'Service Pack'
            if index == -1:
                break
            service_pack = ' SP' + line[index + 13]
            osname = osname + service_pack
            break
    return osname                                 # return OS name

def return_totalram():
    substr1,substr2,substr3,substr4 = 'Total Physical Memory:', 'Полный объем физической памяти:', 'Џ®«­л© ®ЎкҐ¬ дЁ§ЁзҐбЄ®© Ї ¬пвЁ:', 'Total Physical Memory:'  # Substring to search for.
    total_ram = ' '                                 # total ram return variable.
    for line in mylines:                        # string to be searched                           # current index: character being compared
        if line.find('=== Product ===') != -1:
            break
        index1, index2, index3, index4 = 0,0,0,0          # current index: character being compared
        #Loop for getting RAM
        while index1 < len(line) and index2 < len(line) and index3 < len(line) and index4 < len(line): # While index has not exceed string lenght,
            index1 = line.find(substr1, index1)    # set index to first occurence of 'Total Physical Memory:'
            index2 = line.find(substr2, index2)     # set index to first occurence of 'Полный объем физической памяти:'
            index3 = line.find(substr3, index3)    # set index to first occurence of 'Џ®«­л© ®ЎкҐ¬ дЁ§ЁзҐбЄ®© Ї ¬пвЁ:'
            index4 = line.find(substr4, index4)
            if index1 == -1 and index2 == -1 and index3 == -1 and index4 == -1:# If nothing was found,
                break                           # exit the while loop.
            total_ram = line[-9:]
            new_str = ''
            numbers_arr = []
            numbers_arr = re.findall(r'\d+',total_ram)
            for d in numbers_arr:
                new_str = new_str + '%s' % (d)
            total_ram = new_str.replace(' ','')
            total_ram = total_ram.replace(',','')
            total_ram = total_ram.replace('я','')
            total_ram = total_ram.replace('џ','')
            total_ram = str(round((float(total_ram))/1000,2))
            index1 += len(substr1)
            index2 += len(substr2)
            index3 += len(substr3)
            index4 += len(substr4)
    return '%s' % (total_ram) + ' GB'

def return_freeram():
    substr1,substr2,substr3 = 'Available Physical Memory:', 'Доступная физическая память:', '„®бвгЇ­ п дЁ§ЁзҐбЄ п Ї ¬пвм:' # Substring to search for.
    free_ram = ' '                                 # total ram return variable.
    for line in mylines:                        # string to be searched                           # current index: character being compared
        if line.find('=== Product ===') != -1:
            break
        index1, index2, index3 = 0,0,0          # current index: character being compared
        #Loop for getting free ram
        while index1 < len(line) and index2 < len(line) and index3 < len(line): # While index has not exceed string lenght,
            index1 = line.find(substr1, index1)    # set index to first occurence of 'Total Physical Memory:'
            index2 = line.find(substr2, index2)     # set index to first occurence of 'Полный объем физической памяти:'
            index3 = line.find(substr3, index3)    # set index to first occurence of 'Џ®«­л© ®ЎкҐ¬ дЁ§ЁзҐбЄ®© Ї ¬пвЁ:'
            if index1 == -1 and index2 == -1 and index3 == -1:# If nothing was found,
                break                           # exit the while loop.
            free_ram = line[-9:].replace(' ','')
            new_str = ''
            numbers_arr = []
            numbers_arr = re.findall(r'\d+',free_ram)
            for d in numbers_arr:
                new_str = new_str + '%s' % (d)
            free_ram = new_str.replace(' ','')
            free_ram = free_ram.replace(',','')
            free_ram = free_ram.replace('я','')
            free_ram = str(round((float(free_ram))/1000,2))
            index1 += len(substr1)
            index2 += len(substr2)
            index3 += len(substr3)
    return '%s' % (free_ram) + ' GB'

def return_diskinfo():
    disk_caption = ''                           # disk info return variable.
    substr_caption = 'Caption='                     # substring disk volume name
    substr_size = 'Size='                           # substring disk size name
    size_in_gb = 0
    #Loop to get disks info
    for x in range(900):                        # Information should be in the first 900 strings of file
        index = 0                               # current index: character being compared
        if mylines[x].find('=== Disk_Usage ===') != -1: # quit the loop if the string has found
            break
        while mylines[x].find('=== Logical Disks ===') != -1:#loop will execute only one time.
            disk_caption = 'Объем накопителя информации - '
            for i in range(x,x+45):                         # If numbers ov virtual disks less than 5
                index = mylines[i].find(substr_caption)    # set index to occurence of 'Caption='
                if index != -1:                           # If substring has found
                    line = mylines[i]
                    disk_caption = disk_caption + 'диск ' + line[-2:] + ' '
                index = mylines[i].find(substr_size)    # set index to occurence of 'Size='
                if index != -1 and 6 < len(mylines[i]):
                    line = mylines[i]
                    size_in_gb = round((float(line[5:]))/1000000000,2) # Set size in GB
                    if 17 < size_in_gb:
                        disk_caption = disk_caption + str(size_in_gb) + ' GB; '
            x += 1                                          # To quit the loop
    return disk_caption

def return_hardware():
    substr_ghz = 'GHz'                              # substring cpu name
    substr_cores = 'berOfCor'                  # substring number of cpu
    substr_cpuspeed = 'CurrentClockSpeed'           # substring of cpu speed
    substr_motherboard = 'Name='                    # substring of motherboard name
    substr_motherboardversion = 'Version='                 # substring of motherboard version
    hardware = ' '                                  # hardware return variable.
    cpu_name = ' '                                  # cpu name return variable.
    cpu_cores = []                                  # cpu name return variable.
    cpu_speed = ' '                                  # cpu speed return variable.
    motherboard_name = ''                          # motherboard name return variable.
    #Loop to get cpu name
    for line in mylines:                        # string to be searched
        index = 0                               # current index: character being compared
        index1, index2, index3 = 0,0,0          # current index: character being compared
        #Loop to get cpu name
        while index < len(line):
            index = line.find(substr_ghz, index)
            if index == -1:                     # If nothing was found,
                break                           # exit the while loop.
            cpu_name = line[5:]
            index += len(substr_ghz)
        index = 0
        #loop for getting cpu speed
        while index < len(line):
            index = line.find(substr_cpuspeed, index)
            if index == -1:                     # If nothing was found,
                break                           # exit the while loop.
            cpu_speed = str(round((float(line[-4:]))/1000,2)) + ' GHz' # read 4 symbols from the end of line and calculate the value in GHz
            index += len(substr_cpuspeed)
        index = 0
        #Loop to get numbers of cpu
        while index < len(line):
            index = line.find(substr_cores, index)
            if index == -1:                     # If nothing was found,
                break                           # exit the while loop.
            cpu_cores = re.findall(r'\d+',line) # Get numbers from the strings
            index += len(substr_cores)
        index = 0
    # if there is no 'berOfCor' in the file
    if cpu_cores == []:
        cpu_cores = re.findall(r'\d+', mylines[16])
    #Loop for getting a model of the motherboard
    for line in mylines:                        # string to be searched
        index = 0                               # current index: character being compared
        if line.find('=== CPUnameOnly ===') != -1: #for using only first occurence of 'Name='.
            break
        while index < len(line):
            index = line.find(substr_motherboard, index)    # set index to first occurence of 'Name='
            if index == -1:                     # If nothing was found,
                break                           # exit the while loop.
            motherboard_name = line[5:]
            index += len(substr_motherboard)
            break
        index=0
        #Loop for getting a version of the motherboard
        while index < len(line):
            index = line.find(substr_motherboardversion, index)    # set index to first occurence of 'Name='
            if index == -1:                     # If nothing was found,
                break                           # exit the while loop.
            motherboard_name = motherboard_name + ' ' + line[8:]
            index += len(substr_motherboardversion)

    motherboard_name = motherboard_name.replace('  ',' ')
    if cpu_name != ' ':
        hardware = 'CPU - ' + cpu_name
        hardware = hardware + '; Number of CPU - ' + '%s' % (cpu_cores[0]) + ' \n'
        hardware = hardware + 'ОЗУ - ' + return_totalram() + ' \n'
        hardware = hardware + return_diskinfo() + ' \n' + motherboard_name
    else:
        hardware = 'CPU - ' + cpu_speed
        hardware = hardware + '; Number of CPU - ' + '%s' % (cpu_cores[0]) + ' \n'
        hardware = hardware + 'ОЗУ - ' + return_totalram() + '\n'
        hardware = hardware + return_diskinfo() + ' \n' + motherboard_name
    return hardware                                 # return hardware info

def return_biosversion():
#    substr = 'BIOSVersion={'                    # Substring to search for.
#    biosversion = ' '                           # bios version return variable.
#    for line in mylines:                        # string to be searched
#        index = 0                               # current index: character being compared
#        while index < len(line):                # While index has not exceed string lenght,
#            index = line.find(substr, index)    # set index to first occurence of 'Host Name:'
#            if index == -1:                     # If nothing was found,
#                break                           # exit the while loop.
#            biosversion = line[13:-1].replace('\"','')# Append biosversion address to biosversion return variable.
#            biosversion = biosversion.replace('  ',' ')
#            biosversion = biosversion.replace(',',';')
#            index += len(substr)                # Increment the index by the length of substring.
    substr1,substr2,substr3 = 'Версия BIOS:', 'BIOS Version:', '‚ҐабЁп BIOS:' # Substring to search for.
    biosversion = ' '                                  # OS name return variable.
    for line in mylines:                        # string to be searched
        index1, index2, index3 = 0,0,0               # current index: character being compared
        while index1 < len(line) and index2 < len(line) and index3 < len(line): # While index has not exceed string lenght,
            index1 = line.find(substr1, index1)    # set index to first occurence of 'OS Name:'
            index2 = line.find(substr2, index2)     # set index to first occurence of 'Название ОС:'
            index3 = line.find(substr3, index3)    # set index to first occurence of 'Название ОС:'
            if index1 == -1 and index2 == -1 and index3 == -1:# If nothing was found,
                break                           # exit the while loop.
            biosversion = line[24:].replace('  ',',')   # Append host name to hostname return variable. Remove spaces.
            biosversion = biosversion.replace(', ','')
            biosversion = biosversion.replace(',','')
            index1 += len(substr1)                # Increment the index by the length of substring 'OS Name:'.
            index2 += len(substr2)                # Increment the index by the length of substring 'OS Name:'.
            index3 += len(substr3)
    return biosversion                          # return biosversion

def return_hotfix ():
    substr = '['                             # Substring to search for.
    hotfix = []                                 # List of installed hotfixes
    hotfix_str = ''                             # hotfix return variable.
    index_line = 0                              # Number of the line where substring is located
    index_character = 0                         # Number of the first character of substring
    for x in range (30,100):                        # string to be searched
        index_line = x
        index = 0                               # current index: character being compared
        index = mylines[x].find(substr, index)    # set index to first occurence of '['
        index_character = index                 # Save index of the first character of substring.
        if index != -1:                         # If we parsed unless the string we need - quit the loop.
            break
    while mylines[index_line].find(substr) != -1:# Search hotfixes in the next lines in a row
        line = mylines[index_line]
        hotfix.append(line[index_character+6:])
        index_line += 1
    hotfix.sort()
    for line in hotfix:                         # Hotfixes to the string
        hotfix_str = hotfix_str + line + ', \n'
    hotfix_str = hotfix_str[:-2]               # Remove ', ' at the end of the string
    return hotfix_str

def return_disksize ():
    disksize = 0                            # size of the disk
    disksize_str = ''                       # disksize return variable.
    disks_array = []
    str = return_diskinfo()                 # string with disk info
    disks_array = re.findall(r'[-+]?[.]?[\d]+',str)  # Return numbers from the string
    for d in disks_array:
        if d[0] == '.':
             disks_array.remove(d)
    for d in disks_array:
        disksize = disksize + int(d)
    disksize_str = '%s' % (disksize) + ' GB'
    return disksize_str

# Function is not complete
def return_cpuusage ():
    substr = 'LoadPercentage='                             # Substring to search for.
    cpu_u = ''                                # Cpu usage variable to return
    for line in mylines:                        # string to be searched
        index = 0                               # current index: character being compared
        index = line.find(substr, index)    # set index to first occurence of 'LoadPercentage='
        if index != -1:                 # set index to first occurence of 'LoadPercentage='
            cpu_u = line[15:]
    return 'Процессор: ' + cpu_u + '%; '

def return_diskusage_gb():
    disk_diskusage_gb = 'ПЗУ: занято '                           # disk info return variable.
    disk_freespace = 0
    size_in_gb = 0
    substr_size = 'Capacity='                           # substring disk size name
    substr_freespace = 'FreeSpace='                           # substring disk free space name
    disk_size_array = []
    disk_freespace_array= []
    hdd_free_sp = 0
    str1, str2 = '', ''
    #Loop to get disks usage info in gb
    for x in range(500):                        # Information should be in the first 500 strings of file
        index = 0                               # current index: character being compared
        if mylines[x].find('=== Disk_Usage ===') != -1: # quit the loop if the string has found
            break
        while mylines[x].find('=== Volume ===') != -1:#loop will execute only one time.
            for i in range(x,x+45):                         # If numbers ov virtual disks less than 5
                index = mylines[i].find(substr_size)    # set index to occurence of 'Capacity='
                if index != -1 and 15 < len(mylines[i]): # Condition for disk size
                    str1 = mylines[i]                   # String with substring 'Capacity='
                    size_in_gb = round((float(str1[9:]))/1000000,0)  # Set total size in GB
                    if 17 < size_in_gb:
                        disk_size_array.append('%s' % (size_in_gb))
                index = mylines[i].find(substr_freespace)    # set index to occurence of 'FreeSpace='
                disks_array0 = []
                disks_array0 = re.findall(r'[-+]?[.]?[\d]+', mylines[i])  # Return numbers from the string
                for d in disks_array0:
                    hdd_free_sp = int(disks_array0[0])
                if index != -1 and 15 < len(mylines[i]): # Condition for free space
                    str2 = mylines[i]                 # String with substring 'FreeSpace='
                    disk_freespace = round((float(str2[10:]))/1000000,0)
                    if 2 < disk_freespace:
                        disk_freespace_array.append('%s' % (disk_freespace))
            x = x + 1                                          # To quit the loop
    # Calculating disk size - free size
    disk_used_space_in_gb = 0
    for i in range(0,len(disk_size_array)):
        disk_size_array[i] = str(float('%s' % (disk_size_array[i])) - float('%s' % (disk_freespace_array[i])))
    for i in range(0,len(disk_size_array)):
        disk_used_space_in_gb = disk_used_space_in_gb + float('%s' % (disk_size_array[i]))
    disk_diskusage_gb = disk_diskusage_gb + ' %s' % (round(disk_used_space_in_gb/1000,1))
    return disk_diskusage_gb + ' GB; '

def return_ramusage():
    ramusage = 0
    totalram = 0
    freeram = 0
    ramusage_str = ''
    str1 = return_totalram() # get total ram
    str2 = return_freeram() # get free ram
    str1 = str1[:-4]
    str2 = str2[:-4]
    totalram = round(float('%s' % (str1)),3)
    freeram = round(float('%s' % (str2)),3)
    ramusage = round (totalram - freeram, 3)
    #ramusage_str = return_ramusage + '%s' % (ramusage) + ' GB,'
    return 'ОЗУ: занято ' '%s' % (ramusage) + ' GB;'

def return_parsed_data():
    list_one_parsed_file = []
    list_one_parsed_file.append(return_hostname())      # Append host name to result array.
    list_one_parsed_file.append(return_osname())        # Append OS name to result array.
    list_one_parsed_file.append(return_hardware())      # Append hardware info to result array.
    list_one_parsed_file.append(return_biosversion())   # Append bios version to result array.
    list_one_parsed_file.append(return_hotfix())        # Append hotfix info to result array.
    list_one_parsed_file.append(return_ramusage() + return_cpuusage())   # Append usage info to result array.
    list_one_parsed_file.append(return_diskusage_gb())
    return list_one_parsed_file

mylines = []                                # Declare an empty list named mylines.
txt_list = []                               # Declare an empty list for file names.
info_list = []                              # Declare an empty dynamic double array. Result array.

# Get number or txt files
for infile in glob.glob("*.txt"):           # Parse all txt files.
    while infile[:4] != 'test':             # Starts with not 'test'
        txt_list.append(infile)             # Appending for '...txt' to txt_list.
        break                               # Loop works only one time.

# Open file after file and parse data from the files
for indx in range(0,len(txt_list)):
    # Txt file must have encoding utf-8
    with codecs.open("%s" % (txt_list[indx]),"r","utf-8") as myfile:   # Open lorem.txt for reading text data.
    #with open("%s" % (txt_list[indx]),"r") as myfile:   # Open lorem.txt for reading text data.
        for line in myfile:                     # For each line, stored as myline,
            mylines.append(line.rstrip('\n'))   # strip newline and add to list.
    info_list.append(return_parsed_data())
    mylines = []

info_sheet = xlsxwriter.Workbook('info_sheet_test.xlsx')
worksheet = info_sheet.add_worksheet()

col = 0
for row, data in enumerate(info_list):  # Export data from list. Write the data to a sequence of cells
    worksheet.write_row(row,col,data)
info_sheet.close()

print()
print(info_list)
print()
