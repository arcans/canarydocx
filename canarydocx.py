import os
import re
import time
import shutil
import zipfile
import tempfile
import base64
import os.path
from shutil import copy

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def getdefaultpath():
    def_path0 = '776f7264'
    def_path1 = '5f72656c73'
    def_path2 = '666f6f746572322e786d6c2e72656c73'
    def_sep = '2f'
    mime_path = def_path0+def_sep+def_path1+def_sep+def_path2
    defaultpath = bytes.fromhex(mime_path).decode('ascii')
    return defaultpath
 
def getfullpayload(paywhat):
    pay_a = '3c3f786d6c2076657273696f6e3d22312e302220656e636f64696e673d225554462d3822207374616e64616c6f6e653d22796573223f3e0a3c52656c6174696f6e736869707320786d6c6e733d22687474703a2f2f736368656d61732e6f70656e786d6c666f726d6174732e6f72672f7061636b6167652f323030362f72656c6174696f6e7368697073223e3c52656c6174696f6e736869702049643d22724964312220547970653d22687474703a2f2f736368656d61732e6f70656e786d6c666f726d6174732e6f72672f6f6666696365446f63756d656e742f323030362f72656c6174696f6e73686970732f696d61676522205461726765743d22'
    pay_b = '22205461726765744d6f64653d2245787465726e616c22202f3e3c2f52656c6174696f6e73686970733e'
    pay_aa = bytes.fromhex(pay_a).decode('ascii')
    pay_bb = bytes.fromhex(pay_b).decode('ascii')
    return pay_aa+paywhat+pay_bb
    
def generatepayload1(zipname, filename, data):
    # generate a temp file
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(zipname))
    os.close(tmpfd)

    # create a temp copy of the archive without filename            
    with zipfile.ZipFile(zipname, 'r') as zin:
        with zipfile.ZipFile(tmpname, 'w') as zout:
            zout.comment = zin.comment # preserve the comment
            for item in zin.infolist():
                if item.filename != filename:
                    zout.writestr(item, zin.read(item.filename))

    # replace with the temp archive
    os.remove(zipname)
    os.rename(tmpname, zipname)

    # now add filename with its new data
    with zipfile.ZipFile(zipname, mode='a', compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(filename, data)

regex = re.compile(
        r'^(?:http|ftp)s?://' # http:// or https://
        r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+(?:[A-Z]{2,6}\.?|[A-Z0-9-]{2,}\.?)|' #domain...
        r'localhost|' #localhost...
        r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})' # ...or ip
        r'(?::\d+)?' # optional port
        r'(?:/?|[/?]\S+)$', re.IGNORECASE)

print(r"""
	 .---.   .--.  .-. .-.  .--.  .----..-.  .-.
	/  ___} / {} \ |  `| | / {} \ | {}  }\ \/ / 
	\     }/  /\  \| |\  |/  /\  \| .-. \ }  {  
	 `---' `-'  `-'`-' `-'`-'  `-'`-' `-' `--'  
	.----.  .----.  .---..-.  .-.               
	| {}  \/  {}  \/  ___}\ \/ /                
	|     /\      /\     }/ /\ \                
	`----'  `----'  `---'`-'  `-'               
""")


print(f"{bcolors.HEADER}\n 	CANARY TOKEN GENERATOR FOR MS WORD\n 	DOCUMENTS\n		By: Arnice Candoza \n			Edition: [May 29, 2021] \n			v1.0 \n {bcolors.ENDC}")
lol = True
lel = True
lul = True
time.sleep(5)
while lol:
    file_name = input(f"{bcolors.OKCYAN} Enter the path to MS-Word file : \n  >>> {bcolors.ENDC}")
    if os.path.isfile(file_name) and zipfile.is_zipfile(file_name):
        lol=False
        print('\n * Target File ==> '+file_name+'\n')
        while lel:
            token_name = input(f"{bcolors.OKCYAN} Enter your canary token : \n  >>> {bcolors.ENDC}")
            if (re.match(regex, token_name) is not None):
                lel=False
                while lul:
                   print(f'{bcolors.OKBLUE}\n Injecting {bcolors.ENDC}\n ['+token_name+']'+ f'{bcolors.FAIL}'+' ==> '+f'{bcolors.ENDC}'+'['+file_name+']')
                   time.sleep(5)
                   try:
                      generatepayload1(file_name, getdefaultpath(), getfullpayload(token_name))
                      now = str(int(round(time.time() * 1000)))
                      newfile = file_name+"_"+now+"_payload.docx"
                      copy(file_name, newfile)
                      print(f"{bcolors.OKGREEN}\n  STATUS: Payload injected succesfully{bcolors.ENDC}\n  Location: "+newfile)
                   except Exception as e:
	                   print(f'{bcolors.FAIL}\n[ ERROR ] : \n' + str(e))             
                   lul=False               
            else:
                print(f"{bcolors.WARNING}\n ERROR : Invalid canary token !\n")
                time.sleep(2)
    else:
       print(f"{bcolors.WARNING}\n ERROR : Invalid file !\n")
       time.sleep(2)