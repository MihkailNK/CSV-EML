#coding: UTF8
import csv
import os
import random
import string
from email import generator
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, date, time


class Parser()
def get_purchase_information_tax_row(dict_names, row_dict, position): ##возвращает строку нижней таблицы раздела Purchase Information (html-код)
    size = 1
    param_name = dict_names[position]
    if param_name == "":
        return ""
    param_value = row_dict[position]

    try:
        if float(param_value) == 0.0:
            return ''
    except ValueError:
            return ''
    ##исключение - размер шрифта для позиции total_paid = 2, для остальных =1
    if position == 'total_paid':            
        size = 2
    ##Форматирование html строки исходя из текущего заголовка и значения
    result = """
        <tr>
            <td align = right>
                <font face="Verdana, Arial, Helvetica, sans-serif" size="1">
                    <b>{param_name}</b>
                </font>
            </td>
            <td align = right>
                <font face="Verdana, Arial, Helvetica, sans-serif" size="{size}">
                    <nobr>
                        <b>&euro;{param_value} </b>
                    </nobr>
                </font>
            </td>
        </tr>""".format(param_name=param_name, param_value=param_value, size=size)
    return result 
def get_purchase_information_pcakage_row(dict_names, row_dict, position):##возвращает строку верхней таблицы раздела Purchase Information (html-код)
    if position not in row_dict:
        return ""
    result = ""
    param_name = dict_names[position]   ##Получить заголовок из словаря заголовков
    param_value = row_dict[position]    ##Получить значение из словаря значений
    if param_value ==  "":              ##Если значение пустое, вернуть пустую строку
        return ""
    row_html = """<tr>
				    <td width=50% align = right>
					    <font face="Verdana, Arial, Helvetica, sans-serif" size="1">{param_name}:</font>
				    </td>
				    <td width=50%>
					    <font face="Verdana, Arial, Helvetica, sans-serif" size="1">
						    <table width=100% cellpadding=0 cellspacing=0>
							    <tr>
								    <td>
									    <input type="hidden" value="1">
										    <font face="Verdana, Arial, Helvetica, sans-serif" size="1">
											    <b>{param_value}</b>
										    </font>
									    </td>
									    <td align=right>
										    <font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;</font>
									    </td>
								    </tr>
							    </table>
						    </font>
					    </td>
				    </tr>""".format(param_name=param_name, param_value=param_value) ##Форматирует html исходя из значения и заголовка текущей позиции
    return row_html 
def get_purchase_information(row_dict):   ##возвращает html-код сформированного раздела Purchase Information
    purchase_information_names = {         #Словарь заголовков нижней таблицы раздела
        'package_price_paid' : 'Package(s) Total',
        'setup_price' : 'Package(s) Setup',
        'sub_total' : 'Sub Total',
        'vat_paid' : 'Tax Due',
        'total_paid' : 'Total Due'
    }
    purcharse_package_names = {
        'control_panel' : 'Control Panel',
        'additional_IP' : 'Additional IP',
        'domain_to_be_config' : 'Domain to be configured on setup',
        'management_service' : 'Management Service:',
        'os' : 'Operation system',
        'ssl' : 'SSL Certificate',
        'billing_system' : 'Billing System',
        'connection' : 'Connection',
        'cPanel_WHM' : 'cPanel/WHM',
        'plesk_8' : 'PLESK 8',
        'power_pack' : 'Power Pack for Plesk',
        'gameserver' : 'GameServer for Plesk',
        'app_pack' : 'App Pack for PLESK',
        'hdd_raid' : 'HDD Raid',
        'antivirus' : 'Antivirus for PLESK',
        'spamAssasin' : 'SpamAssassin for PLESK'
    }  ##Словарь заголовков верхней части раздела
    ##Разделительная полоса
    tag_hr = """<tr>
				    <td colspan=2>
					    <hr size=1>
					</td>
				</tr>"""
    ##Генерирует строки нижней части раздела
    generated_rows0 = tag_hr
    generated_rows0 += (get_purchase_information_tax_row(purchase_information_names, row_dict, 'package_price_paid'
    generated_rows0 += (get_purchase_information_tax_row(purchase_information_names, row_dict, 'setup_price'
    generated_rows0 += tag_hr
    generated_rows0 += (get_purchase_information_tax_row(purchase_information_names, row_dict, 'sub_total'
    generated_rows0 += (get_purchase_information_tax_row(purchase_information_names, row_dict, 'vat_paid'
    generated_rows0 += (get_purchase_information_tax_row(purchase_information_names, row_dict, 'total_paid'
    generated_rows0 += tag_hr
    ##Генерирует строки верхней части раздела
    generated_rows1 = ""
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'control_panel'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'additional_IP'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'domain_to_be_config'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'management_service'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'os'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'ssl'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'billing_system'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'connection'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'cPanel_WHM'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'plesk_8'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'power_pack'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'gameserver'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'app_pack'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'hdd_raid'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'antivirus'
    generated_rows1 += (get_purchase_information_pcakage_row(purcharse_package_names, row_dict, 'spamAssasin'

    ##Разбор нетипичного случая 
    package_name = row_dict['package_name']
    term_of_lease = ''
    if ' : ' in row_dict['package_name']:
        tmp_idx = row_dict['package_name'].index(' : ')
        package_name = row_dict['package_name'][:tmp_idx]
        if 'Month' in row_dict['package_name']:
            tmp_idx2 = row_dict['package_name'].index('Month')
            term_of_lease = row_dict['package_name'][tmp_idx+1:tmp_idx2+len('Month')]

    #Формирование значения Setup Price
    setup_price = row_dict['setup_price']
    if setup_price == '':
        setup_price = 'FREE'
    else:
        try:
            if float(setup_price) == 0.0:
                setup_price = 'FREE'
            else:
                setup_price = '&euro' + setup_price;
        except ValueError:
            print('Error: ' + str(ValueError))
            setup_price = ''


    ##Генерирует html раздела исходя из сформированных строк верхней и нижней части
    result = """<table width=450>
	<tr>
		<td>
			<pre>
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="2">
					<tr>
						<td align = left width=35%>
							<font face="Verdana, Arial, Helvetica, sans-serif" size="1">
								<nobr>
									<b>Main Package</b>
								</nobr>
							</font>
						</td>
						<td align = right>
							<nobr>
								<font face="Verdana, Arial, Helvetica, sans-serif" size="1">
									<b>{package_name}</b>&nbsp;&nbsp;<br>&nbsp;&nbsp;{term_of_lease} @ &euro;{package_price_paid} + Setup: {setup_price}</font>
								</nobr>
							</td>
						</tr>"""
                        
    result = result.format(package_name=package_name, term_of_lease=term_of_lease, package_price_paid=row_dict['package_price_paid'], setup_price=setup_price)
    result += generated_rows1 + generated_rows0 
    result += """
            <tr>
                <td align = center colspan=2>
                    <font face="Verdana, Arial, Helvetica, sans-serif" size="1">
                        <b>*</b> &euro;{package_price_paid} will be due on the next renewal date: {vat_free_customer}<br>(Domain renewals billed separately.)</font>
                </td>
            </tr>
                </table>
            </td>
            </tr>"""
    result.format(**row_dict)
    result += """
                        </table>
                    </pre>
                </td>
            </tr>
            </table>
            <br clear = left>
            <br /> 
    """

    return result  
def get_contact_information_row(dict_names, row_dict, position):
    if position not in row_dict:
        param_value = ''
    if position not in dict_names:
        return ''
    try:
        param_name = dict_names[position]
        param_value = row_dict[position]
        if param_name == 'Address':
            if ', ' in param_value:
                tmp_list = param_value.split(' ')
                idx = 2
                for i in range(len(tmp_list)):
                    if ', ' in tmp_list[i]:
                        idx = i
                        break
                param_value = ''
                tmp_list = tmp_list[:idx] + ['<br /> '] +tmp_list[idx:]
                for el in tmp_list:
                    param_value += el+' '
            return """{param_name}:<br /> 
                {param_value}<br /> 
                <br /> """.format(param_name=param_name, param_value=param_value)
        elif param_name == 'Email':
            return  '''{param_name}: <a href=mailto:{param_value}>{param_value}</a><br /> '''.format(param_name=param_name, param_value=param_value)
        else:
            return '{param_name}: {param_value} <br /> '.format(param_name=param_name, param_value=param_value)
    except Exception:
        print("Error in contact information (column: " + position + ' is not recognized\n')

def get_contact_information_html(row_dict): ##Возвращает сформированный раздел Contact Information(html-код)
    dict_names = {
        'first_name' : 'First Name',
        'last_name' : 'Last Name',
        'address' : 'Address',
        'phone_number' : 'Phone',
        'email' : 'Email'
    }

    result = ""
    result += (get_contact_information_row(dict_names, row_dict, 'first_name'))
    result += (get_contact_information_row(dict_names, row_dict, 'last_name'
    result += ("Company or Domain Name: <br /> "
    result += (get_contact_information_row(dict_names, row_dict, 'address'
    result += (get_contact_information_row(dict_names, row_dict, 'phone_number'
    result += (get_contact_information_row(dict_names, row_dict, 'email'
    result += ("""<br /> 
    <br /> """)
    return result

def get_account_or_billing_information_row(dict_names, row_dict, position): #Возвращает строку, по формату подходящую для разделов Account Information и Billing Information
    if position not in row_dict:            ##Если текущей позиции (названия столбца) нет в полученной из csv таблице
        return ''                           ##вернуть пустую стрку
    if row_dict[position] == '':            ##Если строка в таблице все таки есть, но ее значение не заполнено для текущей строки
        return ''                           ##вернуть пустую строку
    position_name = dict_names[position]    ##Получаем заголовок из словаря dict_names
    position_value = row_dict[position]     ##Получаем значение из словаря row_dict
    if position == 'account_url':           ##исключение для position == account_url - другой формат возвращаемой html строки
        return '{position_name}: <A href={position_value}>{position_value}</a><br /> '.format(position_name=position_name, position_value=position_value)
    else:
        return '{position_name}: {position_value}<br /> '.format(position_name=position_name, position_value=position_value) ##Форматирование строти исходя из заголовка и значения

def account_information_randomizer(row_dict): ##Рандомизирует данные "account information"
    a_z = string.ascii_lowercase
    a_z0_9 = a_z + '0123456789'
    refferal = ''
    password = ''
    try:
        len = random.randint(5, 15)
        rand_idx = [random.randrange(26) for x in range(len)]
        for i in range(len):
            refferal += a_z[rand_idx[i]]
        refferal = refferal[0].upper() + refferal[1:]

        len = random.randint(5, 15)
        rand_idx = [random.randrange(36) for x in range(len)]
        for i in range(len):
            password += a_z0_9[rand_idx[i]]

        login = row_dict['email']
        row_dict.update({
            'referral' : refferal,
            'password' : password,
            'login' : login
        })
    except Exception:
        print("Error in function: account_information_randomizer\n" + str(Exception))
        
def get_account_information(row_dict):  ##возвращает сформированный раздел Account Information (html-код)
    dict_names = {                      ##Словарь заголовков строк html таблицы
        'referral' : 'Referred by',
        'password' : 'Password',
        'login' : ' Login/Email'
    }
    account_information_randomizer(row_dict)
    result = ''      
    result += (get_account_or_billing_information_row(dict_names, row_dict, 'referral'
    result += (get_account_or_billing_information_row(dict_names, row_dict, 'password'
    result += (''' -------------------<br /> '''                                            ##Разделитель
    result += (get_account_or_billing_information_row(dict_names, row_dict, 'login'
    result += (get_account_or_billing_information_row(dict_names, row_dict, 'password'
    return result

def billing_info_randomizer(row_dict):##Рандомизирует данные по кредитной карте
        if 'payment_type' in row_dict:
            if row_dict['payment_type'] =='CreditCard':
                try:
                    credit_card_number = 'xxxx-xxxx-xxxx-{0:0>4}'.format(str(random.randrange(100, 9999)))
                    cardholder_name = row_dict['first_name'] + ' ' + row_dict['last_name']
                    cvv2 = '{0:0<3}'.format(random.randint(10, 999))
                    card_exp_date = ''
                    if row_dict['time_stamp'] != '':
                        tmp_list = row_dict['time_stamp'].split('.')
                        y_random_inc = random.randint(0, 2)
                        yy = str(int(tmp_list[0]) + y_random_inc)
                        if y_random_inc:
                            mm = str(random.randint(1, 12))
                        else:
                            mm_t = int(tmp_list[1])
                            mm = str(random.randint(mm_t, 12))
                        card_exp_date = mm + '//' + yy

                    row_dict.update({'cardholder_name':cardholder_name,
                                     'card_exp_date':card_exp_date,
                                     'CVV2':cvv2,
                                     'credit_card_number':credit_card_number})
                except Exception:
                    print("Error in function: billing_info_randomizer\n" + str(Exception))

def get_billing_information(row_dict):  ##возвращает сформированный раздел Billing Information (html-код)
    result = ''
    dict_names = {        ##Словарь заголовков раздела
        'payment_type' : 'Payment Method',
        'cardholder_name' : 'Cardholder Name',
        'credit_card_number' : 'Credit Card Number',
        'card_exp_date' : 'Exp Date',
        'CVV2' : 'CVV2',
        'ip_from_purchase' : 'Order From IP',
        'time_stamp' : 'Timestamp'
    }
    billing_info_randomizer(row_dict)
    ##Запрос на формирования html кода строк, по возможным заголовкам раздела)
    result += (get_account_or_billing_information_row(dict_names, row_dict, 'payment_type'
    result += ("<br /> "*2
    result += (get_account_or_billing_information_row(dict_names, row_dict, 'cardholder_name'
    result += (get_account_or_billing_information_row(dict_names, row_dict, 'credit_card_number'
    result += (get_account_or_billing_information_row(dict_names, row_dict, 'card_exp_date'
    result += (get_account_or_billing_information_row(dict_names, row_dict, 'CVV2'
    result += ("<br /> "
    result += (get_account_or_billing_information_row(dict_names, row_dict, 'ip_from_purchase'
    result += ("<br /> "
    result += (get_account_or_billing_information_row(dict_names, row_dict, 'time_stamp'
    return result

def eml_write(row_dict, row_index, out_path): ##Формирует eml исходя из полученной из csv строки
    sender = 'Billing@hosting-ie.com'       ##Отправитель
    recepiant = 'Billing@hosting-ie.com'    ##Получатель
    subject = 'Signup Invoice/Receipt for ' + row_dict['first_name'] + ' ' + row_dict['last_name'] ##тема письма

    ##Формирование MIME объекта (eml)
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = recepiant
    contact_information = get_contact_information_html(row_dict) ##Получить сформированный html раздела Contact Information
    purchase_information = get_purchase_information(row_dict) ##Получить сформированный html раздела Purchase Information
    account_information = get_account_information(row_dict) ##Получить сформированный html раздела Account Information
    billing_information = get_billing_information(row_dict) ##Получить сформированный html раздела Billing Information
   
   ##Сформировать html часть письма
    html = """\
        <html>
            <head></head>
            <body>
               <b>Contact Information</b><br /> 
                -------------------<br /> """ + contact_information + """
               <b>Purchase Information</b><br /> 
                --------------------<br /> 
                <br /> """ + purchase_information + """
                <b>Account Information</b><br /> 
                -------------------<br /> """ + account_information + """
                <br /> 
                <br /> 
                <b>Billing Information</b><br /> 
                -------------------<br /> """ + billing_information + """
            </body>
        </html>
        """
    part = MIMEText(html, 'html')
    msg.attach(part)  ##Прикрепление сформированного html к пимьму
    outfile = open(out_path+"\\{0:0>8}.eml".format(row_index), 'w') ##Открытие файла на запись - имя файла зависит от текущей строки
    gen = generator.Generator(outfile)
    gen.flatten(msg)   ##Запись пиьма в файл


def main():  ##Основное "тело" функции
    in_path = str(input("Enter the path to .csv:"))
    ifile = open(in_path, 'r') ##Открытие input на чтение

    out_path = ".\\output-folder" ##Путь к папке с сформированными eml

    if not os.path.isdir(out_path):
        os.mkdir(out_path) 
    lines_in_file = len(ifile.readlines())-1
    ifile.seek(0)
    csv_reader = csv.reader(ifile, delimiter=', ')  ## получение CSV Reader для текущего файла с разделителем ', '
    header = []
    is_first = True 
    Second = True
    i = 0
    print("Process: ")
    for row in csv_reader: ##Построчно читаем файл как csv
        procent = float(i)/lines_in_file*100.0
        print('\r' * 10, end='')
        print(str('{0}').format(str(procent)[:4])+"%", end='')
        if is_first:  ##Если строка - первая
            header = row            ##Сформировать список заголовков
            is_first = False
        else: 
            i += 1
            d = dict(zip(header, row))  ##ассоциация списка заголовков, со списком заголовков
            
            ##Расчет Sub Total
            d.update({'sub_total':d['package_price_paid']})
            try:
                if d['setup_price'] != '': ##Если SetupPrice != ''
                    sub_total = float(d['package_price_paid'])+float(d['setup_price']) ##SubTotal = setup_price+package_price_paid
                    d.update({'sub_total':str(sub_total)})
            except Exception:
                print('Error ' + str(Exception))
            eml_write(d, i, out_path)##Сформировать файл .eml - исходя из данных текущей строки csv
    pass


main()