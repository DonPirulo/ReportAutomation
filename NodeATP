import openpyxl
import datetime
import re
fecha = datetime.datetime.now()
counter = 0
wb = openpyxl.load_workbook('Template Node.xlsx')
sheet = wb['ATP']
hostname = input('Insert Hostname: ')
ring= input('Insert Ring: ')
#hostname = 'SGS_KruisStreet_SRA4-1'
#region = input('Insert ring: ')
area = 'SGS_JDO'
loopbackcounter=0
loopbackPTPcounter=0
model=''
textfile = hostname + '.txt'
regex = re.compile(r"(?<!\d)\d{5}(?!\d)")


eventcounter = 0
with open(textfile, 'r') as f:
    ###Hostname find###
    for line in f:
        counter = counter + 1

        ######### VERSION #################
        if line.find('TiMOS') is not -1:
            sheet['K25'] = line[2:17]
            sheet['K26'] = line[10:17]

        ######        ISIS       #######
        if line.find('area-id') is not -1:
            isis = line[20:30]
            sheet['K33'] = isis[-1:] + '(' + isis + ')'
        if line.find('level-capability level-1') is not -1:
            sheet['Q33'] = 'L1'
        if line.find('level-capability level-2') is not -1:
            sheet['Q33'] = 'L2'

        ######        ASN       #######
        if line.find('autonomous') is not -1:
            sheet['K23'] = str(regex.search(line).group())

        if line.find('Network_Queue_7705') is not -1:
            sheet['K51'] = 'Vodacom_CTN_Network_Queue_7705'

        ######        IPs       #######
        if line.find('interface "system"') is not -1:
            loopbackcounter =counter+1
        if counter == loopbackcounter:
            if line.find('address') is not -1:
                sheet['K21'] = line[20:-4]
                sheet['K22'] = line[20:-4]
                sheet['K35'] = line[20:-4]
                sheet['Q42'] = line[20:-4]

        if line.find('interface "PTP_loopback"') is not -1:
            loopbackPTPcounter = counter+1
        if counter == loopbackPTPcounter:
            if line.find('address') is not -1:
                sheet['K36'] = line[20:-4]

        if line.find('event') is not -1:
            eventcounter = eventcounter+1
        if eventcounter > 11:
            sheet['I57'] = 'Y'
            sheet['J57'] = ''
        if line.find('g8275dot1') is not -1:
            sheet['K47'] = 'ITU-T G.8275.1'



######        Model       #######
if hostname.find('SARA') is not -1:
    sheet['J4'] = 'X'
    sheet['K29'] = 'Singleton'
    sheet['J30'] = 'N'
    sheet['K30'] = ''
if hostname.find('SRA4') is not -1:
    sheet['P4'] = 'X'
    sheet['I30'] = 'Y'

if hostname.find('SAR8') is not -1:
    sheet['M4'] = 'X'
    sheet['I30'] = 'Y'

sheet['K19'] = area[0:3]
sheet['K20'] = area
sheet['F7'] = hostname
sheet['F9'] = hostname
sheet['K14'] = hostname
sheet['K42'] = hostname
sheet['W33'] = ring
sheet['F68'] = fecha.strftime('%d %B %Y')

sheet = wb['Current Config']

counter = 1
with open(textfile, 'r') as f:
    ###Hostname find###
    for line in f:
        counter = counter + 1
        cell = 'C' + str(counter)
        sheet[cell] = line

wb.save("%s.xlsx" % hostname)
