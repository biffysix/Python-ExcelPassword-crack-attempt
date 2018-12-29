#need this to use raw_imput:NameError: name 'raw_input' is not defined
#python 2 used raw_input Python 3 just uses "input"
# path: C:\Users\LennyT61\Desktop\Youtube\ExcelPassword\Excelprotect.xlsx
# need to do in command line
## not working properly as it takes first line in wordlist and uses that as the
## password crack

from win32com.client import Dispatch
import sys

file = input('Path: ')
wordlist = 'wordlist.txt'

word = open(wordlist, 'r')
allpass = word.readlines()
word.close()

for password in allpass:
    password = password.strip()
    print ("Testing password: "+password)
    instance = Dispatch('Excel.Application')

    try:
        instance.Workbooks.Open(file, False, True, None, password)
        print ("Password Cracked: "+password)
        break

    except:
        pass
