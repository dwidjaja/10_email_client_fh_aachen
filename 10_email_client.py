'''
FH Aachen
Abschlussprojekt Informationstechnologie 1 WS18/19
Fachbereich Wirtschaftwissenschaften
'''

from bs4 import BeautifulSoup
import urllib.request
import tkinter as tk
from tkinter import *
from tkinter import messagebox
import smtplib

class Dictionary:

    def getData(self, source, id):

        self.source = source
        self.id = id

        Dictionary.employee_dictionary = {}

        for paragraph in BeautifulSoup(urllib.request.urlopen(self.source).read(), 'lxml').find_all('div', id=self.id):
            name = paragraph.find('h2', style='padding-top:0px').string
            email = paragraph.select('a')[1].text.replace('(at)', '@')
            try:
                subject = paragraph.find('span', style='border: 1px solid transparent;').string
            except Exception:
                subject = ''
            Dictionary.employee_dictionary.update({name: [email, subject]})
            
        return Dictionary.employee_dictionary

class GUI:

    # get the dictionary from every 'Fachbereich'
    fb1 = Dictionary().getData('https://www.fh-aachen.de/fachbereiche/architektur/menschen/', 'mensch_356')
    fb2 = Dictionary().getData('https://www.fh-aachen.de/fachbereiche/bauingenieurwesen/menschen/', 'mensch_1939')
    fb3 = Dictionary().getData('https://www.fh-aachen.de/fachbereiche/chemieundbiotechnologie/menschen/', 'mensch_209')
    fb4 = Dictionary().getData('https://www.fh-aachen.de/fachbereiche/gestaltung/menschen/', 'mensch_138')
    fb5 = Dictionary().getData('https://www.fh-aachen.de/fachbereiche/elektrotechnik-und-informationstechnik/menschen/', 'mensch_358')
    fb6 = Dictionary().getData('https://www.fh-aachen.de/fachbereiche/luft-und-raumfahrttechnik/menschen/', 'mensch_359')
    fb7 = Dictionary().getData('https://www.fh-aachen.de/fachbereiche/wirtschaft/menschen/', 'mensch_139')
    fb8 = Dictionary().getData('https://www.fh-aachen.de/fachbereiche/maschinenbau-und-mechatronik/menschen/', 'mensch_78')
    fb9 = Dictionary().getData('https://www.fh-aachen.de/fachbereiche/medizintechnik-und-technomathematik/menschen/', 'mensch_360')
    fb10 = Dictionary().getData('https://www.fh-aachen.de/fachbereiche/energietechnik/menschen/', 'mensch_361')
    fb_collection = [fb1, fb2, fb3, fb4, fb5, fb6, fb7, fb8, fb9, fb10]

    def __init__(self):

        self.root = tk.Tk()
        self.root.title('E-Mail Versender')
        self.root.geometry('805x670')
        self.root.configure(bg='light grey')

        self.label1 = tk.Label(self.root, text='', width = 2, height = 4, bg='#00b2a9')
        self.label2 = tk.Label(self.root, text='', width = 2, height = 4, bg='black')
        self.label3 = tk.Label(self.root, font= ('Calibri',14), bg='light grey', text='FH-Kennung')
        self.label4 = tk.Label(self.root, font= ('Calibri',14), bg='light grey', text='FH-Email')
        self.label5 = tk.Label(self.root, font= ('Calibri',14), bg='light grey', text='Passwort')
        self.label6 = tk.Label(self.root, font= ('Calibri Bold',18), bg='light grey', text='{0:50s}{1:1s}'.format('E-Mail an',':'))
        self.label7 = tk.Label(self.root, text='', bg='light grey')                 # Label for selected_employee
        self.label8 = tk.Label(self.root, text='', bg='light grey')                 # Label for Lehrgebiet
        self.label9 = tk.Label(self.root, font= ('Calibri',14), bg='light grey', text='Betreff')

        self.button1 = tk.Button(self.root, width = 9, height = 5, text = ' E-Mail \n senden', font = ('Calibri Bold',18), relief=RAISED, command = self.send_email)

        self.entry1 = tk.Entry(self.root, width = 45, borderwidth=.1)               # Entry for FH-Kennung
        self.entry2 = tk.Entry(self.root, width = 45, borderwidth=.1)               # Entry for FH-Email
        self.entry3 = tk.Entry(self.root, width = 45, borderwidth=.1, show = '*')   # Entry for Password
        self.entry4 = tk.Entry(self.root, width = 45, borderwidth=.1)               # Entry for Betreff

        self.text1 = tk.Text(self.root, width = 111, height = 15)                   # Textbox for Meldung

        self.labelframe1 = tk.LabelFrame(self.root, bg='light grey', font = ('Calibri Italic',20), padx=4, pady=4, text='Fachbereich')
        self.labelframe2 = tk.LabelFrame(self.root, bg='light grey', font = ('Calibri Italic',20), padx=4, pady=4, text='Name des Mitarbeiters')

        self.listbox1 = tk.Listbox(self.labelframe1, width = 35, exportselection = 0)
        list_fb = ['01 - Architektur', \
                   '02 - Bauingenieurwesen', \
                   '03 - Chemie und Biotechnologie', \
                   '04 - Gestaltung', \
                   '05 - Elektrotechnik und Informationstechnik', \
                   '06 - Luft und Raumfahrttechnik', \
                   '07 - Wirtschaftwissenschaften', \
                   '08 - Maschinenbau und Mechatronik', \
                   '09 - Medizintechnik und Technomathematik', \
                   '10 - Energietechnik']
        for fb in list_fb:
            self.listbox1.insert(END, fb)

        for i in range(0, len(list_fb), 2):
            self.listbox1.itemconfigure(i, background='#f0f0ff')

        # scrollbar for the listbox 2
        self.scrollbar = tk.Scrollbar(self.labelframe2, orient=VERTICAL, bd=1)
        self.listbox2 = tk.Listbox(self.labelframe2, yscrollcommand=self.scrollbar.set, width = 43, exportselection=0)
        self.scrollbar.config(command=self.listbox2.yview)

        # placing the widgets
        self.label1.place(x=770, y=33)          # Logo Mint
        self.label2.place(x=770, y=156)         # Logo Black
        self.label3.place(x=10, y=285)          # Label for 'FH-Kennung'
        self.label4.place(x=10, y=315)          # Label for 'FH-Email'
        self.label5.place(x=10, y=345)          # Label for 'Password'
        self.label6.place(x=10, y=228)          # Label for 'E-Mail an :'
        self.label7.place(x=265, y=228)         # Label for selected_employee
        self.label8.place(x=265, y=254)         # Label for Lehrgebiet
        self.label9.place(x=10, y=375)          # Label for 'Betreff'

        self.entry1.place(x=265, y=285)         # Entry for FH-Kenunng
        self.entry2.place(x=265, y=315)         # Entry for FH-Email
        self.entry3.place(x=265, y=345)         # Entry for Password
        self.entry4.place(x=265, y=375)         # Entry for Betreff

        self.button1.place(x=694, y=284)        # Button for SEND

        self.text1.place(x=10, y=410)           # Entry for Meldung

        self.labelframe1.place(x=10, y=20)      # Listbox for Fachbereich
        self.labelframe2.place(x=345, y=20)     # Listbox for Name

        self.listbox1.pack()
        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.listbox2.pack()

        # binding
        self.listbox1.bind('<<ListboxSelect>>', self.create_listbox_employee)

        self.root.mainloop()

    def create_listbox_employee(self, *args):

        self.listbox2.delete(0, END)

        selected_fb = self.listbox1.get(self.listbox1.curselection())

        for k,v in self.fb_collection[int(selected_fb[1])-1].items():
            self.listbox2.insert(END,k)
            
        self.listbox2.bind('<<ListboxSelect>>', self.create_labelname)

        for i in range(0, len(self.fb_collection[int(selected_fb[1])-1]), 2):
            self.listbox2.itemconfigure(i, background='#f0f0ff')

        self.selected_fb = selected_fb

    def create_labelname(self, *args):

        selected_employee = self.listbox2.get(self.listbox2.curselection())
        self.label7.config(text=selected_employee, font = ('Calibri Bold', 20))

        subject = self.fb_collection[int(self.selected_fb[1])-1][selected_employee][1]
        self.label8.config(text=subject, font = ('Calibri Italic', 14))

        self.selected_employee = selected_employee

    def send_email(self, *args):

        try:
            to_email = self.fb_collection[int(self.selected_fb[1])-1][self.selected_employee][0]
            fh_kennung = self.entry1.get() + '@ad.fh-aachen.de'
            fh_email = self.entry2.get()
            email_password = self.entry3.get()
            betreff = self.entry4.get()
            meldung = self.text1.get('1.0', END)
            message = 'Subject: {}\n\n{}'.format(betreff, meldung).encode('utf8')

            server = smtplib.SMTP('mail.fh-aachen.de', 587)
            server.starttls()
            server.login(fh_kennung, email_password)
            server.sendmail(fh_email, to_email , message)
            server.quit()

            self.entry1.delete(0, END)
            self.entry2.delete(0, END)
            self.entry3.delete(0, END)
            self.entry4.delete(0, END)
            self.text1.delete('1.0', END)

            messagebox.showinfo('Info', 'E-Mail erfolgreich versendet.')

        except (Exception, smtplib.SMTPAuthenticationError):

            messagebox.showinfo('Error!', 'Es sind Fehler aufgetreten. Bitte überprüfen Sie Ihre Angaben.')

if __name__ == '__main__':
    GUI()
