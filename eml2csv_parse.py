#coding: cp1252
from email.header import decode_header
import email
from base64 import b64decode
import sys
from email.parser import Parser as EmailParser
from email.utils import parseaddr
from io import StringIO
import os
import re
import html

class Requst:    
    def __init__(self):
        self.FirstName = str()
        self.LastName = str()
        self.Address = str()
        self.Phone = str()
        self.Email = str()
        self.MainPackage = str()
        self.Coupon = str()
        self.PackageTotal = str()
        self.Timestamp = str()
        self.PackageSetup = str()
        self.TaxDue = str()
        self.Term_of_lease = str()
        self.Connection = str()
        self.Control_Panel = str()
        self.KVM = str()
        self.Management_Service = str()
        self.OS = str()
        self.cPanel_WHM = str()
        self.PLESK_8 = str()
        self.Power_Pack_for_Plesk = str()
        self.GameServer_for_Plesk = str()
        self.App_Pack_for_PLESK = str()
        self.Antivirus_for_PLESK = str()
        self.SpamAssassin_for_PLESK = str()
##
def unescape(s):  ##замена html спец. символов в тексте.
    t = s
    d = {'&amp;' : '&',
         '&reg;' : "",
         '&lt;' : '<',
         '&gt;' : '>',
         '&cent;' : "",
         '&pound;' : "",
         '&yen;' :  "",
         '&euro;' : "",
         '&sect;' : '',
         '&copy;' : "",
         '&apos;' : "'",
         '&quot;' : '"',
         '&Atilde;': 'A',
         '&curren;': '',
         '&para;' : "",
         '&nbsp;' : ' '
     }
    for ket in d:
        t = t.replace(ket, d[ket])
    return t
def get_body_html(content):
    p = EmailParser()
    try:
        msgobj = p.parse(content)
    except Exception:
         return ''
    html = None
    for part in msgobj.walk():
        if part.get_content_type() == "text/html":
            if html is None:
                html = ''
            uhtml = (part.get_payload(decode=True))#.decode('utf-8', 'replace')
            html += uhtml.decode('utf8', 'replace')## str(uhtml.encode('utf-8', 'replace'))
    return html
def get_all_eml_files_in_dir(dir):  ##Функция возвращает список всех .eml в выбранной директории
    farray = os.listdir(dir)
    farray = [x.split('.')  for x in farray]
    f_array_parsed = []
    for i in range(len(farray)):
        f = farray[i]
        if (len(f) >= 1) and (f[-1] == 'eml'):
            tmp_str = dir+'\\'
            for j in range(len(f)):
                tmp_str += f[j]
                if j < (len(f)-1):
                    tmp_str += '.'
            f_array_parsed.append(tmp_str)
    return (f_array_parsed)
def get_str_to_format(list_args): ##Функця получает список аргументов, и возвращает строку для csv форматирования по их количеству
    s = ''
    l = len(list_args)
    for i in range(l):
        s += '{'+str(i)+'};'
    s = s[:-1]
    return s
def replace_euro_char(s):  ##Функция замены сивола евро (UTF8) в строке на пробел.
    res = s
    if '\\\\u0432\\\\u201a' in s:
        idx =s.index('\\\\u0432\\\\u201a')
        res = s[0: idx]+' '+s[idx+len('\\\\u0432\\\\u201a')+1:]
    return res
def str_sub_array(array, i, j):
    res = ''
    for el in array[i:j]:
        res += replace_euro_char(el)+' '
    res = res[:-1]
    return res
    
 ##Функции разбора данных
def get_firstname(array, te, save, i):      ##First Name
    if (te == "First"):
        return array[i+2]
    return ''
def get_lastname(array, te, save, i):       ##Last Name
    if (te == "Last"):
        return array[i+2]
    return ''
def get_adressposition(array, te, save, i):        ##Adressposition
    if (te == "Address:"):
        return  i + 1
    return -1
def get_phone_and_adress(array, te, save, i, adressposition):          ##Phone && Adress
    result = ["", '']
    adress = ''
    if (te == "Phone:"):
        for m in range(adressposition, i):
            adress += " " + array[m]
        Phone = array[i+1]
        result = [Phone, adress]
    return result
def get_timestamp(array, te, save, i):      ##Timestamp
    if (te == "Timestamp:"):
        return array[i+1].replace(':', '')
    return ''
def get_mainpackage_and_term_of_lease(array, te, save, i):    ##Main Package && Term of lease
    MainPackage = ''
    Term_of_lease = ''
    if (te == "Main") and ('Package'.upper() in array[i+1].upper() and ('@' in array[i+1: i+10])):
        MainPackage = array[i+2] + " " + array[i+3] + array[i+4] + array[i+5]
        if ':' in MainPackage:
            j = MainPackage.index(':')
            MainPackage = MainPackage[: j]
        j = -1
        for k in range(i+2, i+8):
            if 'Month' in array[k]:
                j = k
                break
        if j != -1:
            if ':' in array[j]:
                k = array[j].index(':')
                Term_of_lease =  array[j][k+1: ]
            else:
                Term_of_lease = array[j]
        return [MainPackage, Term_of_lease]
    return ["", '']
def get_package_total(array, te, save, i):   ##PackageTotal
    if (save == "Package(s)" and te == "Total"):
        return replace_euro_char(array[i+1])
    return ''
def get_package_setup(array, te, save, i):   ##Package Setup
    if (save == "Package(s)" and te == "Setup"):
        return replace_euro_char(array[i+1])
    return ''
def get_tax_due(array, te, save, i):         ##Tax Due
    if (save == "Tax" and te == "Due"):
        return replace_euro_char(array[i+1])
    return ''
def get_coupon(array, te, save, i):         ##Coupon
    Coupon = ''
    if (te == "Coupon"):
        if '(' in array[i+1]:
            j = -1
            for e_idx in range(i+1, i+7):
                if ')' in array[e_idx]:
                    j = e_idx
        if j != -1:
            Coupon = str_sub_array(array, i+1, j+2)
        else:
            Coupon = str_sub_array(array, i+1, i+3)
        return Coupon
    return ''
def get_connection(array, te, save, i):     ##Connection
    Connection = ''
    if ('Connection' == te) or ('Connection:' == te):
        if 'BIT' in array[i-1].upper():
            Connection = array[i-1]+': '+array[i+1]
        else:
            if ('---' in array[i+1]) or ('NO' in (array[i+1]).upper()):
                Connection = array[i+1]
            else:
                if 'Setup:' in array[i+1: i+10]:
                    j = array[i+1: i+10].index('Setup:')+1
                    Connection = str_sub_array(array, i+1, i+j+2)
                else:
                    Connection = str_sub_array(array, i+1, i+4)
        return Connection
    return ''
def get_control_panel(array, te, save, i):          ##Control Panel
    Control_Panel = ''
    if (save == "Control" and ("Panel:" in te)):
        if 'NO' in (array[i+1]).upper():
            Control_Panel = array[i+1]
        else:
            Control_Panel =str_sub_array(array, i+1, i+5)
        return Control_Panel
    return ''
def get_kvm(array, te, save, i):                ##(Extra port to manage server, i)
    if (('KVM' in save) and ('Extra' in te)):
        return array[i+5]
    return ''
def get_management_service(array, te, save, i): ##Manage Service
    Management_Service = ''
    if (('Management' in save) and  ('Service' in te)):
        if 'NO' in array[i+1].upper():
            Management_Service = array[i+1]
        else:
            if 'Setup:' in array[i+1: i+10]:
                j = array[i+1: i+10].index('Setup:')+2
                Management_Service = str_sub_array(array, i+1, i+j+1)
            else:
                Management_Service = str_sub_array(array, i+1, i+5)
        return Management_Service
    return ''
def get_os(array, te, save, i):             ##OS
    if 'OS:' in te:
        if 'FREE' in array[i: i+15]:
                j = array[i: i+15].index('FREE')
        else:
            if 'edition' in array[i: i+15]:
                j = array[i: i+15].index('edition')
            else:
                j = 4
        return str_sub_array(array, i+1, i+j+1)
    return ''
def get_cpanel_whm(array, te, save, i):     ##cPanel/WHM
    cPanel_WHM = ''
    if 'cPanel/WHM' in te:
        if ('No' in array[i+1]):
            cPanel_WHM = array[i+1]
        else:
            if 'Setup:' in array[i+1: i+15]:
                j = array[i+1: i+20].index('Setup:')+3
            else:
                if ('(Self' in array[i+1: i+10]):
                    j =  array[i+1: i+10].index('(Self')+2
                elif ('Self' in array[i+1: i+10]):
                    j =  array[i+1: i+10].index('Self')+2
                else:
                    tmp_array = [x.upper() for x in array[i+1: i+20]]
                    if 'BILLING' in tmp_array:
                        j = tmp_array.index('Billing'.upper())+1
            cPanel_WHM = str_sub_array(array, i+1, i+j)
        return cPanel_WHM
    return ''
def get_PLESK_8(array, te, save, i):                    ##PLESK 8
    if 'PLESK 8:' in te:
        if ('No' in array[i+1]):
            return array[i+1]
        else:
            return str_sub_array(array, i+1, i+6)
    return ''
def get_power_pack_for_plesk(array, te, save, i):       ##PowerPack for PLESK
    if (('Pack' in te) and ('Power' in save)):
        if array[i-2].upper() == 'PLESK':
            j = i+1
        else:
            j = i+3
        if ('NO' in array[j].upper()):
            return replace_euro_char(array[j])
        elif (array[i+3] == 'Additional'):
              return replace_euro_char(array[i+2])
        else:
            return str_sub_array(array, j, j+3)
    return ''
def get_gameserver_for_plesk(array, te, save, i):       ##GameServer for PLESK
    if ('GameServer' in save) and ('for' in te):
        if 'NO' in array[i+2].upper():
            return array[i+2]
        else:
            return replace_euro_char(array[i+2])
    return ''
def get_app_pack_for_plesk(array, te, save, i):         ##AppPack for PLESK
    if ('App' in save) and ('Pack' in te):
        return replace_euro_char(array[i+3])
    return ''
def get_antivirus_for_plesk(array, te, save, i):        ##AV for PLESK
    if ('Antivirus' in save) and ('for' in te):
        if 'NO' in array[i+2].upper():
            return array[i+2]
        else:
            return str_sub_array(array, i+2, i+8).split(' ')[-1]
    return ''
def get_spamassassin_for_plesk(array, te, save, i):     ##SpamAssasin for PLESK
    if ('SpamAssassin' in save) and ('for' in te):
        if 'NO' in array[i+2].upper():
            return array[i+2]
        else:
            if 'Notes:' in array[i+2: i+5]:
                j = array[i+2: i+5].index('Notes:')+1
                return str_sub_array(array, i+2, i+j)
            else:
                return str_sub_array(array, i+2, i+5)
    return ''

def main():
    print("Enter the path:")
    dir = input().replace("'\'", "'\\\'"*2)
    f_array = get_all_eml_files_in_dir(dir)
    csv = []
    
    ##CSV заголовки
    headers = (                             
        'First Name',
        'Last Name',
        'Adress',
        'Phone',
        'Timestamp',
        'Main Package',
        'PackageTotal',
        'Package Setup',
        'Tax Due',
      ##  'Coupon',
        'Term of lease',
        'Connection',
        'Control Panel',
        'KVM',
        'Manage Service',
        'OS      ',
        'cPanel/WHM',
        'PLESK 8',
        'PowerPack for PLESK',
        'GameServer for PLESK',
        'AppPack for PLESK',
        'AV for PLESK',
        'SpamAssassin for PLESK'
    )
    csv.append( get_str_to_format(headers).format(*headers))
    error_count = 0
    fout = open(dir+'\oput.csv', 'w')
    f_log_out = open(dir+'\log.txt', 'w')
    print("Files: {0}\n".format(len(f_array)))
    print("Current file:")
    for path in f_array:
        if '087.eml' in path:
            print('!')
        print('\r'*10, end="")
        print(f_array.index(path)+1, end="")
        fin = open(path, 'r')
        t = str()
        req = Requst()
        t = get_body_html(fin)
        if t:
            t = re.sub(r'\<[^>]*\>', '', t)
            t = t.replace('\n', ' ').replace("&nbsp;", '')
            t = repr(t)
            t = unescape(t)
            array = t.split(" ")
            while '' in array:
                i = array.index("")
                array.pop(i)
            i = 0
            adressposition = 0
            save = ''
            form_str = '';
            
            ##Проверки: если параметр еще не заполнен, вызвать функицю проверки-заполнения
            for te in array:
                if (req.FirstName == ''):              
                    req.FirstName = get_firstname(array, te, save, i)
                if (req.LastName == ''):   
                    req.LastName = get_lastname(array, te, save, i)
                if (adressposition <= 0):
                    adressposition = get_adressposition(abs, te, save, i)
                if ((req.Phone  == '') or (req.Address  == '')):
                    tmp_list = get_phone_and_adress(array, te, save, i, adressposition);
                    req.Phone =tmp_list[0]
                    req.Address = tmp_list[1]
                if (req.MainPackage == '') or (req.Term_of_lease == ''):
                    tmp_list = get_mainpackage_and_term_of_lease(array, te, save, i)
                    req.MainPackage = tmp_list[0]
                    req.Term_of_lease = tmp_list[1]
                if (req.Timestamp == ''):
                    req.Timestamp = get_timestamp(array, te, save, i)
                if (req.PackageTotal == ''):
                    req.PackageTotal = get_package_total(array, te, save, i)
                if (req.PackageSetup == ''):
                    req.PackageSetup = get_package_setup(array, te, save, i)
                if (req.TaxDue == ''):
                    req.TaxDue = get_tax_due(array, te, save, i)
                if (req.Coupon == ''):
                   req.Coupon = get_coupon(array, te, save, i)
                if (req.Control_Panel == ''):
                   req.Control_Panel = get_control_panel(array, te, save, i)
                if (req.Connection == ''):
                    req.Connection = get_connection(array, te, save, i)
                if (req.KVM == ''):
                    req.KVM= get_kvm(array, te, save, i)
                if (req.Management_Service == ''):
                   req.Management_Service = get_management_service(array, te, save, i)
                if (req.OS == ''):
                    req.OS = get_os(array, te, save, i)
                if  req.cPanel_WHM == '':
                    req.cPanel_WHM = get_cpanel_whm(array, te, save, i)
                if req.PLESK_8 == '':
                    req.PLESK_8 = get_PLESK_8(array, te, save, i)
                if req.Power_Pack_for_Plesk == '':
                    req.Power_Pack_for_Plesk = get_power_pack_for_plesk(array, te, save, i)
                if req.GameServer_for_Plesk  == '':
                    req.GameServer_for_Plesk = get_gameserver_for_plesk(array, te, save, i)
                if req.App_Pack_for_PLESK == '':
                    req.App_Pack_for_PLESK = get_app_pack_for_plesk(array, te, save, i)
                if req.Antivirus_for_PLESK == '':
                   req.Antivirus_for_PLESK = get_antivirus_for_plesk(array, te, save, i)
                if (req.SpamAssassin_for_PLESK == ''):
                    req.SpamAssassin_for_PLESK = get_spamassassin_for_plesk(array, te, save, i)
                i += 1
                save = te
            fin.close()
            if (req.FirstName != ''):
                ##Kортеж всех элементов req. для возможности быстрого комментирования.
                ##Что бы пропустить колонку- просто закомментируйте одну из строк ниже:
                csv_tuple = (                                              
                            req.FirstName,                      ##First Name
                            req.LastName,                       ##Last Name
                            req.Address,                        ##Adress
                            req.Phone,                          ##Phone
                            req.Timestamp,                      ##Timestamp
                            req.MainPackage,                    ##Main Package
                            req.PackageTotal,                   ##PackageTotal
                            req.PackageSetup,                   ##Package Setup
                            req.TaxDue,                         ##Tax Due
###                        req.Coupon,                          ##Coupon
                            req.Term_of_lease,                  ##Term of lease
                            req.Connection,                     ##Connection
                            req.Control_Panel,                  ##Control Panel
                            req.KVM,                            ##(Extra port to manage server)
                            req.Management_Service,             ##Manage Service
                            req.OS,                             ##OS
                            req.cPanel_WHM,                     ##cPanel/WHM
                            req.PLESK_8,                        ##PLESK 8
                            req.Power_Pack_for_Plesk,           ##PowerPack for PLESK
                            req.GameServer_for_Plesk,           ##GameServer for PLESK
                            req.App_Pack_for_PLESK,             ##AppPack for PLESK
                            req.Antivirus_for_PLESK,            ##AV for PLESK
                            req.SpamAssassin_for_PLESK
                )

                if form_str == '': 
                    form_str = get_str_to_format(csv_tuple)             ##Получение строки для форматирования
                csv.append(form_str.format(*csv_tuple))                 ##Форматирование csv строки для текущих параметров 

            else:               ##Ошибка
                f_log_out.writelines("Error of data recognize: "+path+'\n')
                error_count += 1
        else:                   ##Ошибка
            f_log_out.writelines("Error of getting html: "+path+'\n')
            error_count += 1
    print('\n')
    for n in csv:                                                       ##Запись всех элементов массива csv в выходной файл
        while '\\'+'\\' in n:
            i = n.index('\\\\')
            n = n[0: i]+n[i+7: ]
        try:
            d = str(n.encode('ascii', errors='ignore')) 
        except Exception:
            fout.writelines(n+'\n')
        else:
            fout.writelines(d[2: -1]+'\n')
    fout.close()
    f_log_out.close()
    print("Errors found in: {0} files. log.txt\n".format(error_count))
    
main()