import os #Modul za OS
import time #Modul za ispis vremena
import datetime #Modul za ispis datuma
import shutil #Modul za upravljanje folderima
import win32com.client #Modul za Outlook
import xlsxwriter #Modul za upis u excel
import pdfReaderImport as pdfR #Moj modul za čitanje pdf-a i provjeru korisnika

#Prebacuje račune u zadani folder (radi)
def send_inv(inv_folder, file):
    global main_folder
    if file[0:3] == 'Rač':#Mislim da je ponovna provjera nepotrebna
        time.sleep(5)
        sent = pdfR.oib_isolation(main_folder, file)
        if sent == True:
            data[5][1].append('True')
        else:
            data[5][1].append('False')
        sent = False
        from_sorce = main_folder + '/' + file
        to_sorce = inv_folder + '/' + file
        shutil.move(from_sorce, to_sorce)

#Prebacuje ostale dokumente u zadani folder
#win32com.client sada radi, to me za sada kosta outlook mailova i pisanje podataka u excel tablicu
def send_dn(dn_folder, file):
    global main_folder
    global dn_list
    if file[0:3] != 'Rač':#Mislim da je ponovna provjera nepotrebna
        time.sleep(5)
        from_sorce = main_folder + '/' + file
        to_sorce = dn_folder + '/' + file
        dn_list.append(file)
        sent = mail(from_sorce)
        if sent == True:
            data[5][1].append('True')
        else:
            data[5][1].append('False')
        sent = False   
        shutil.move(from_sorce, to_sorce)

#Slanje otpremnice u Phoenix
def mail(path):
    global text
    win32com.client.Dispatch('outlook.application')
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    #mail.To = 'z.brzak@phoenix-farmacija.hr; s.persic@phoenix-farmacija.hr'
    #mail.Cc = 'petar.huljek@biomax.hr; ordering@biomax.hr; a.pintaric@phoenix-farmacija.hr'

    mail.To = 'petar.huljek@biomax.hr'
    mail.subject = 'Biomax otpremnica'
    mail.Body = text
    mail.Attachments.Add(path)
    mail.Send()

    return True

#Funkcija za upis u excel tablicu i poziv u funkciju za premještanje tablice u poseban folder
def excel_data_entry(data):
    excel_move_path = 'C:/Users/Petar/OneDrive - BIOMAX d.o.o/biomax sharing/IRA 2022/Evidencija'
    now = datetime.datetime.now()
    sufix = 0
    excel_name = now.strftime('%d%m%Y')
    excel_name = excel_name + '.xlsx' #Nazivanje excel datoteke po datumu pisanja 
    while True:
        if excel_name in os.listdir(excel_move_path):
            sufix -= 1
            excel_name = now.strftime('%d%m%Y') + str(sufix) + '.xlsx' #Dodavanje sufixa sada funkcionira. 
        else:
            break
    #print(excel_name)
    workbook = xlsxwriter.Workbook(excel_name)
    worksheet = workbook.add_worksheet()
    col = 0
    row = 0
    for collumn, queue in data:
        worksheet.write(row, col, collumn)
        for j in range(len(queue)):
            #print(collumn, queue)
            worksheet.write(row + 1 + j, col, queue[j])
        col += 1
    workbook.close()
    excel_move(excel_name)

#Funkcija prebacuje gotovu xlsx tablicu u određeni folder
def excel_move(excel_name): 
    excel_save_path = 'C:/Petar/Java/Python'
    excel_move_path = 'C:/Users/Petar/OneDrive - BIOMAX d.o.o/biomax sharing/IRA 2022/Evidencija'
    excel_move_path = excel_move_path + '/' + excel_name
    excel_save_path = excel_save_path + '/' + excel_name
    shutil.move(excel_save_path, excel_move_path)
    return

#Kreiranje ugnježđene liste s potrebnim podatcima
def data_input(file, size, number):
    global data
    data[0][1].append(number)
    data[1][1].append(file)
    if file[:3] == 'Rač':
        data[2][1].append('Račun')
    else:
        data[2][1].append('Otpremnica')
    data[3][1].append(size)
    catch = datetime.datetime.now()
    catch = catch.strftime('%H:%M:%S')
    data[4][1].append(catch)
    #if file[0:3] == 'Rač':
    #    data[5][1].append('False')
    #else:
    #    data[5][1].append('True')
    return data

#Kreirana funkcija za upis teksta koji se dodaje u novi mail na kraju dana
#sa popisom svih poslanih otpremnica taj dan
def dn_mail_tekst(dn_list):
    mail_part = ''
    for dn in dn_list:
        mail_part = mail_part + dn + '\n'
    return mail_part

def dn_mail_check():
    global dn_list
    mail_part = dn_mail_tekst(dn_list) #Kreiranje teksta iz liste otpremnica
    win32com.client.Dispatch('outlook.application')
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    #mail.To = 'z.brzak@phoenix-farmacija.hr; s.persic@phoenix-farmacija.hr'
    #mail.Cc = 'ordering@biomax.hr; josip.sabljic@biomax.hr; domagoj.nikolic@biomax.hr'

    mail.To = 'petar.huljek@biomax.hr'
    mail.subject = 'Biomax otpremnice za danas'
    mail.Body = header + mail_part + footer
    mail.Send()

main_folder = 'C:/Petar/Java/Python/main_folder'
inv_folder = 'C:/Petar/Java/Python/inv_folder'
dn_folder = 'C:/Petar/Java/Python/dn_folder'
text = '''Poštovana,

Molim isporuku po dokumentu iz privitka.

Unaprijed hvala.

Srdačan pozdrav/Best regards,

*****************
Petar Huljek
Biomax d.o.o.
Perjavička putina 5
10090 Zagreb
Tel:  +385-1-3470173
Fax: +385-1-3470195
Email: petar.huljek@biomax.hr
WWW: http://www.biomax.hr/'''
header = '''Poštovana,

Današnje otpremnice su:

'''
footer = '''
Unaprijed hvala.

Srdačan pozdrav/Best regards,

*****************
Petar Huljek
Biomax d.o.o.
Perjavička putina 5
10090 Zagreb
Tel:  +385-1-3470173
Fax: +385-1-3470195
Email: petar.huljek@biomax.hr
WWW: http://www.biomax.hr/'''
data = [
    ['Br.', list()],
    ['Naziv dokumenta', list()],
    ['Tip dokumenta', list()],
    ['Veličina dokumenta', list()],
    ['Vrijeme slanja', list()],
    ['Poslan mail', list()]
]
dn_list = []

def main():   
    number = 0
    condition = True
    while condition:
        for file in os.listdir(main_folder):
            path = os.path.join(main_folder, file)
            if os.path.isfile(path) and os.path.getsize(path) > 0:#Provjerava postoji li datoteka i zauzima li neku memoriju
                number += 1
                size = os.path.getsize(path)
                data = data_input(file, size, number)
                if file[0:3] == 'Rač': #Određuje sadržaj datoteke
                    #print('Premjesti u račun!')
                    send_inv(inv_folder, file)
                else: #Sve što nije račun se šalje mailom
                    #print('Premjesti u poslano')
                    send_dn(dn_folder, file)
        now = datetime.datetime.now()
        set_time = now.replace(hour= 15, minute= 10, second=0)
        if now > set_time:#provjera vremena rada programa, pomaknuti kasnije na željeno vrijeme npr. 15:30
            #print('Execute excel function!!')
            excel_data_entry(data) #Bolja pozicija za unos u tablicu
            #print(data)
            dn_mail_check()
            condition = False
        else:
            time.sleep(10)

if __name__ == "__main__":
   try:
      main()
   except KeyboardInterrupt:
      excel_data_entry(data)
      dn_mail_check()
      pass
