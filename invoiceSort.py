import PyPDF2
import win32com

buyer_dict = {
    'VAT': 'email@email.com'
}
text = '''Poštovani/a,

U privitku je račun za robu koja je poslana Phoenixom.

Srdačan pozdrav,

******************
Petar Huljek
Biomax d.o.o.
Perjavička putina 5
10090 Zagreb
Tel:  +385-1-3470173
Fax: +385-1-3470195
Email: petar.huljek@biomax.hr
WWW: http://www.biomax.hr/

'''

#Funkcija izolira OIB kupca iz računa i šalje OIB u funkciju check_vat koja vrača binarnu vrijednost ,
#ukoliko je oib u riječniku buyer_dict. Nakon toga šalje direktorij do računa i mail iz rječnika u funkciju slanja računa kupcu. 
def oib_isolation (location_folder, file_name):
    full_fName = location_folder + '/' + file_name
    file = open(full_fName, 'rb')
    fileReader = PyPDF2.PdfFileReader(file)
    filePage = fileReader.getPage(0)
    tekst = filePage.extractText()
    tekst2 = tekst[480:550]
    #print(tekst)
    #print('_____________')
    #print(tekst2)
    vat = ''
    counter = 0
    for i in range(len(tekst2)):
        if tekst2[i].isdigit() and counter < 11:
            vat = vat + tekst2[i]
            counter += 1

    print(vat)
    file.close()

    check = check_vat(vat)

    if check:
        inv_send(full_fName, buyer_dict[vat])
    
    return check

    #print(tekst[173:250]) #Indeks za kompletnu adresu kupca, odstupanje po par indeksa(mislim da ovisi o formatu datuma, dužini narudžbenice)
    #print(tekst2)

#Funkcija prima OIB kupca kao ulazni parametar i provjerava popis kupaca u riječniku buyer_dict, vrača True ako je oib u rječniku i False ako nije.
def check_vat(vat):
    global buyer_dict
    if vat in buyer_dict.keys():
        print(vat, buyer_dict[vat]) #Provjera
        return True
    else: return False


#Funkcija prima direktorij do računa i mail kupca i šalje račun na mail koristeći Outlook.
def inv_send(path, buyer_mail):
    global text
    win32com.client.Dispatch('outlook.application')
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    #mail.To = buyer_mail #Tu bi trebao ići mail kupca
    #mail.Cc = 'petar.huljek@biomax.hr; ordering@biomax.hr;'

    mail.To = 'petar.huljek@biomax.hr'
    mail.subject = 'Biomax račun'
    mail.Body = text + buyer_mail
    mail.Attachments.Add(path)
    mail.Send()
