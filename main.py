import tkinter as tk
import smtplib
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path
from typing import List
from tkinter import ttk, filedialog


global email
global password
global subject    # E-Postanın Konusu
global server
global listMails
global listNames
global certificatePath
global message


root = tk.Tk();
root.geometry("700x700")
root.title("Mail Automation")

errorLabel = tk.Label(root)
errorLabel.place(x = 40, y=1)

user_name_label = tk.Label(root,
                  text = "Host Email").place(x = 40,
                                           y = 20)

user_password_label = tk.Label(root,
                      text = "Password").place(x = 40,
                                               y = 60)
user_subject_label = tk.Label(root,
                      text = "Subject").place(x = 40,
                                               y = 100)


user_email_input = tk.Entry(root,width= 40)
user_email_input.place(x=110,y=20)


user_password_entry = tk.Entry(root,width=40)
user_password_entry.place(x=110,y=60)

user_subject_entry = tk.Entry(root,width =60)
user_subject_entry.place(x=110,y=100)






certificate_path = tk.Label(root)
emails_path = tk.Label(root)
names_path = tk.Label(root)


def sertifikalar():
    global certificatePath
    folder_selected = filedialog.askdirectory()
    certificate_path.place(x=230,y=130)
    certificate_path.config(text= folder_selected)
    certificatePath = folder_selected

def mailler():
    global listMails
    # filename = filedialog.askopenfilename()
    tf = filedialog.askopenfilename(
        initialdir="C:/Users/MainFrame/Desktop/",
        title="Open Text file",
        filetypes=(("Text Files", "*.txt"),)
    )
    tf = open(tf)
    listMails = []
    #listMails = tf.read().splitlines()
    for line in tf:
        temp= line.strip()
        if not temp == '': #Burada dosyadaki olası boşluklar kontrol edildi
            listMails.append(temp)
    emails_path.place(x=160, y=180)
    emails_path.config(text = tf.name)
    tf.close()



def isimler():
    global listNames
    tf = filedialog.askopenfilename(
        initialdir="C:/Users/MainFrame/Desktop/",
        title="Open Text file",
        filetypes=(("Text Files", "*.txt"),)
    )
    tf = open(tf)
    listNames =[]

    for line in tf:
        temp= line.strip()
        if not temp == '': #Burada dosyadaki olası boşluklar kontrol edildi
            listNames.append(temp)
    """listNames= tf.readlines()
    listNames = [x.strip() for x in listNames]"""
    names_path.place(x=160,y=230)
    names_path.config(text= tf.name)
    tf.close()


choose_certificate_label = tk.Label(root,
                      text = "Select image FOLDER :").place(x = 40,
                                               y = 130)

choose_certificate_button = tk.Button(root,
                       text="Choose", command= sertifikalar).place(x=40,
                                            y=150)


choose_emails_label = tk.Label(root,
                      text = "Email txt file :").place(x = 40,
                                               y = 180)

choose_emails_button = tk.Button(root,
                       text="Choose", command= mailler).place(x=40,
                                            y=200)


choose_names_label = tk.Label(root,
                      text = "Name text file :").place(x = 40,
                                               y = 230)

choose_names_button = tk.Button(root,
                       text="Choose", command= isimler).place(x=40,
                                            y=250)

subject_screen_label = tk.Label(root,
                            text ="Email Mesajı Girin:").place(x=40, y=290)

message_Text = tk.Text(root, height=8, width=60)
message_Text.place(x=40, y=320)


log_screen_label = tk.Label(root,
                            text ="Log Screen :").place(x=40, y=470)

log_Text = tk.Text(root, height=8, width=60)
log_Text.place(x=40, y=500)

def appendLog(myStr):
    log_Text.insert(tk.END,"\n"+myStr)

def startSending():
    global email
    global password
    global server
    global certificatePath
    global listNames
    global listMails
    global message
    global subject
    #Boş bırakılan alan var mı o kontrol ediliyor. İstenilirse arttırılabilir, değiştirilebilir

    if(user_email_input.get() =="" or user_password_entry.get() =="" or user_subject_entry.get()==""):
        errorLabel.config(text="Lütfen boş bırakılan yerleri doldurunuz !")
        return

    message = message_Text.get("1.0","end-1c")
    email = user_email_input.get()
    password = user_password_entry.get()
    subject = user_subject_entry.get()
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email, password)

    for i in range(len(listMails)):
        sendTo = listMails[i]
        name = listNames[i]

        nameENG = name.replace("İ", "I")
        nameENG = nameENG.replace("ı", "i")
        nameENG = nameENG.replace("ö", "o")
        nameENG = nameENG.replace("Ö", "O")
        nameENG = nameENG.replace("ü", "u")
        nameENG = nameENG.replace("Ü", "U")
        nameENG = nameENG.replace("ğ", "g")
        nameENG = nameENG.replace("Ğ", "G")
        nameENG = nameENG.replace("ş", "s")
        nameENG = nameENG.replace("Ş", "S")
        nameENG = nameENG.replace("ç", "c")
        nameENG = nameENG.replace("Ç", "C")
        inviteLink = "linktr.ee/ieeege"

        fileLocation = certificatePath +"/" + nameENG + ".jpg"
        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = sendTo
        msg['Subject'] = subject
        fileName = os.path.basename(fileLocation)
        msg.attach(MIMEText(message, 'plain'))
        attachment = open(fileLocation,'rb')
        fName = nameENG+".jpeg"
        part = MIMEBase('image', 'png', filename=fName)
        part.add_header('Content-Disposition', 'attachment', filename=fName)
        part.add_header('X-Attachment-Id', '0')
        part.add_header('Content-ID', '<0>'.format(nameENG))
        part.set_payload(attachment.read())
        encoders.encode_base64(part)

        msg.attach(part)

        text = msg.as_string()
        server.sendmail(email, sendTo, text)
        textForLog = listMails[i]+" adresine mail gönderildi."
        appendLog(textForLog)
        attachment.close()


submit_button = tk.Button(root,
                       text="Start Sending !", command= startSending).place(x=570,
                                            y=650)

def saveInstance():
    tempE = user_email_input.get()
    tempP = user_password_entry.get()
    tempS = user_subject_entry.get()
    tempM = message_Text.get("1.0",'end-1c')

    dictionary_data = {"hostEmail": tempE, "hostPassword": tempP, "hostSubject":tempS, "message":tempM}

    f = open("SavedInstances.txt", "w")
    f.write(str(dictionary_data))
    f.close()
    errorLabel.config(text="Email bilgileri kaydedildi. Bir sonraki açılışta otomatik doldurulacaktır")

save_button = tk.Button(root,
                       text="Save Instances !", command= saveInstance).place(x=460,
                                            y=650)

def set_text(e,text): #Burası Entry widget'ı editleme metodu
    e.delete(0,'end')
    e.insert(0,text)
    return

def setTextInput(textWidget,text): #Burası text widget'ı editleme metodu
    textWidget.delete(1.0,"end")
    textWidget.insert(1.0, text)

def loadInstance():
    t = open("SavedInstances.txt", "r") #boş olup olmadığını test etmek için farklı adla açıldı. diğer işlemleri etkilemesi engellendi
    f = open("SavedInstances.txt", "r")
    emptyTest = t.read(1)
    t.close()
    #boş değilse if kısmına giriyor.
    if emptyTest:
        temp = f.read()
        dictionary = eval(temp)
        set_text(user_email_input,str(dictionary['hostEmail']))
        set_text(user_password_entry,str(dictionary['hostPassword']))
        set_text(user_subject_entry,str(dictionary['hostSubject']))
        setTextInput(message_Text,str(dictionary['message']))

    f.close()




loadInstance()


root.mainloop()


