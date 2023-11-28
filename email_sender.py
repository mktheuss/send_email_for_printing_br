import smtplib
import tkinter.messagebox
from datetime import datetime
from tkinter import *
import ssl
from os.path import exists
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# IMPORTANT: Most of the strings here are in Brazillian Portuguese.

def checkingTheHour():
    # Gets the current date and time
    currentTime = datetime.now()
    currentHourString = currentTime.strftime("%H")
    currentHour = int(currentHourString)
    # Checks what kind of greetings should be used based on the current system hour
    if 0 <= currentHour <= 12:
        greetingsText = 'Bom dia'
    elif 12 < currentHour <= 18:
        greetingsText = 'Boa tarde'
    else:
        greetingsText = 'Boa noite'
    return str(greetingsText)

# This funcion creates the whole e-mail. Should be refactored to reduce its size.
def sendingEmail():
    copies = (number_of_copies.get(1.0, END)).strip('\n')
    subject = "Impressão de documento!"
    body = (greetingsString + '! Preciso imprimir ' + str(copies) + ' cópias do arquivo em anexo!\n\n' +
                  'Poderiam me encaminhar o valor da impressão?\n\n'
                  'Passo em breve para retirar. Obrigada!')
    sender_email = "insert email here!"
    receiver_email = (email_receiver.get(1.0, END)).strip('\n')
    password = 'insert password here!'

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  # Recommended for mass emails

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    # Defines the word file to attach
    filename = (my_file.get(1.0, END)).strip('\n') + '.docx'  # In same directory as script

    # Open file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)

# Grabs the texts from the textboxes and checks if they are correct and if the attachment will work.
# There are more verifications that could be done, but this is a simple, straight forward version.
def grabText():
    copies = (number_of_copies.get(1.0, END)).strip('\n')
    if copies.isalpha():
        mensagem = tkinter.messagebox.Message(root, message='Número de cópias não está correto!', title='Aviso!')
        mensagem.show()
    text = (my_file.get(1.0, END)).strip('\n') + '.docx'
    bool = exists(text)
    if bool == True:
        mensagem = tkinter.messagebox.Message(root, message='O arquivo existe! Pode clicar em "Enviar e-mail logo abaixo!"',
                                   title='Aviso!')
        mensagem.show()
    else:
        mensagem = tkinter.messagebox.Message(root, message='O arquivo não existe! Favor rever o nome pra ver se digitou certo!"',
                                   title='Aviso!')
        mensagem.show()

# Starting tkinter
root = Tk()

# Title of the window
root.title('Enviar e-mail')

# Size of the window
root.geometry('500x600')

# Text on top of the window
w = Label(root, text="1. Digite o nome do arquivo que desejas mandar no campo!")
w.pack(pady=5)

# Creating the text box for the 
my_file = Text(root, width=50, height=1)
my_file.insert('end', 'Exemplo: APARTAMENTO 02 CONTRATO')
my_file.pack(pady=20)

# Creating a label to explain the textbox below (e-mail receiver)
my_label = Label(root, text='2. Digite abaixo o endereço de e-mail de quem vai imprimir!\n'
                            'Por padrão, já deixei o e-mail da Qualicopy!')
my_label.pack(pady=5)

# Creating the textbox to inform who will receive the e-mail
email_receiver = Text(root, width=50, height=1)
email_receiver.insert('end', 'qualicopy.pomerode@gmail.com')
email_receiver.pack(pady=20)

# Creating a label to explain the second textbox (number of copies)
my_label2 = Label(root, text='3. Digite abaixo o número de cópias que deseja, somente números!')
my_label2.pack(pady=5)

# Creating the text box for the number of copies
number_of_copies = Text(root, width=10, height=1)
number_of_copies.insert('end', '02')
number_of_copies.pack(pady=20)

# Creating a label to explain the check button (checks if the file exists on the directory of this script)
my_label3 = Label(root, text="4. Clique no botão 'Confirmar nome do arquivo'.\n"
                             "Leia as instruções da mensagem e corrija o que precisar!\n"
                             "Ele vai verificar se o arquivo realmente existe!")
my_label3.pack(pady=5)

# Creating the button frame
button_frame = Frame(root)
button_frame.pack()

# Creates button to grab text from the text box
grab_text = Button(button_frame, text='Confirmar nome do arquivo!', command=grabText)
grab_text.grid(row=0, column=0, pady=20)

# Creating a label to explain the Send E-mail button
my_label3 = Label(root, text="5. Clique no botão 'Enviar e-mail!'.\n")
my_label3.pack(pady=5)

# Creating a second button frame
button_frame2 = Frame(root)
button_frame2.pack()

# Creates the button to send e-mail
greetingsString = checkingTheHour()
send_email = Button(button_frame2, text='Enviar e-mail!', command=sendingEmail)
send_email.grid(row=1, column=0, pady=20)

# Runs the loop
root.mainloop()
