import pyperclip

def convertMac():
    '''Program to convert from non hyphenated MAC Addresses to fully Hyphenated Addresses
    These MAC Addresses are then copied into the clipboard
    EX: 1234567891234567 -> 1234-5678-9123-4567'''

    mac = input('Type your MAC address here: ')
    count = 0
    convertedMac = ''
    for num in mac:
        if count%2 == 1:
            convertedMac += num + '-'
            count+=1 
        elif count%2 == 0:
            convertedMac += num 
            count+=1 
    print(convertedMac[0:17])
    pyperclip.copy(convertedMac[0:17])

while True:
    convertMac()