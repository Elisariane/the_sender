import smtplib
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox


# Configurações do servidor SMTP
smtp_server = "smtp.gmail.com"
smtp_port = 587

# Função para enviar um único e-mail com anexo
def send_email(cliente, coordenador_email, gerente_email, superintendente_email, email_user, password_user):
    try:
        message = MIMEMultipart()
        message["From"] = email_user
        message["To"] = coordenador_email
        message["Cc"] = f"{gerente_email},{superintendente_email}"
        message["Subject"] = cliente['nome']
        message.attach(MIMEText(cliente['texto'], "plain"))

        # Anexa o arquivo personalizado
        attachment_path = cliente['anexo']
        with open(attachment_path, 'rb') as attachment:
            att = MIMEBase('application', 'octet-stream')
            att.set_payload(attachment.read())
            encoders.encode_base64(att)
            att.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
            message.attach(att)
        
        # Conecta ao servidor SMTP
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(email_user, password_user)

        # Envia o email
        server.sendmail(email_user, [coordenador_email] + gerente_email.split(",") + superintendente_email.split(","), message.as_string())
        server.quit()

        print(f'E-mail do(a) cliente {cliente["nome"]} enviado para coordenador')
    except Exception as e:
        print(f'Houve um erro ao tentar enviar email para {coordenador_email}: {e}')

# Lê a lista de clientes do arquivo CSV
def ler_clientes(caminho_csv):
    clientes = []
    with open(caminho_csv, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            clientes.append(row)
    return clientes

# Função principal para enviar e-mails simultaneamente
def send_emails(caminho_csv, coordenador_email, gerente_email, superintendente_email, email_user, password_user):
    threads = []
    clientes = ler_clientes(caminho_csv)

    for cliente in clientes:
        thread = threading.Thread(target=send_email, args=(cliente, coordenador_email, gerente_email, superintendente_email, email_user, password_user))
        threads.append(thread)
        thread.start()
    
    # Aguarda todas as threads terminarem
    for thread in threads:
        thread.join()

    print("Todos os e-mails foram enviados.")

def selecionar_csv():
    caminho_csv = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    entry_csv.delete(0, tk.END)
    entry_csv.insert(0, caminho_csv)

# Função para iniciar o envio de e-mails
def iniciar_envio():
    caminho_csv = entry_csv.get()
    coordenador_email = entry_coordenador_email.get()
    gerente_email = entry_gerente_email.get()
    superintendente_email = entry_superintendente_email.get()
    email_user = entry_email_user.get()
    password_user = entry_password_user.get()
    if not caminho_csv or not coordenador_email or not gerente_email or not superintendente_email or not email_user or not password_user:
        messagebox.showwarning("Atenção", "Por favor, preencha todos os campos.")
        return
    send_emails(caminho_csv, coordenador_email.replace(" ", ""), gerente_email.replace(" ", ""), superintendente_email.replace(" ", ""), email_user.replace(" ", ""), password_user)
    messagebox.showinfo("Sucesso", "Todos os e-mails foram enviados.")

# Configuração da Janela
root = tk.Tk()
root.title("Envio de E-mails em Massa")
# Elementos da Interface
label_csv = tk.Label(root, text="Arquivo CSV:")
label_csv.grid(row=0, column=0, padx=10, pady=10)

entry_csv = tk.Entry(root, width=50)
entry_csv.grid(row=0, column=1, padx=10, pady=10)

btn_csv = tk.Button(root, text="Selecionar", command=selecionar_csv)
btn_csv.grid(row=0, column=2, padx=10, pady=10)

label_email_user = tk.Label(root, text="Seu e-mail:")
label_email_user.grid(row=1, column=0, padx=10, pady=10)

entry_email_user = tk.Entry(root, width=50)
entry_email_user.grid(row=1, column=1, padx=10, pady=10)

label_password_user = tk.Label(root, text="Digite sua senha gerada para Apps no Gmail:")
label_password_user.grid(row=2, column=0, padx=10, pady=10)

entry_password_user = tk.Entry(root, width=50, show="*")
entry_password_user.grid(row=2, column=1, padx=10, pady=10)

label_coordenador_email = tk.Label(root, text="E-mail do Coordenador:")
label_coordenador_email.grid(row=3, column=0, padx=10, pady=10)

entry_coordenador_email = tk.Entry(root, width=50)
entry_coordenador_email.grid(row=3, column=1, padx=10, pady=10)

label_gerente_email = tk.Label(root, text="E-mail do Gerente:")
label_gerente_email.grid(row=4, column=0, padx=10, pady=10)

entry_gerente_email = tk.Entry(root, width=50)
entry_gerente_email.grid(row=4, column=1, padx=10, pady=10)

label_superintendente_email = tk.Label(root, text="E-mail do Superintendente:")
label_superintendente_email.grid(row=5, column=0, padx=10, pady=10)

entry_superintendente_email = tk.Entry(root, width=50)
entry_superintendente_email.grid(row=5, column=1, padx=10, pady=10)

btn_enviar = tk.Button(root, text="Enviar E-mails", command=iniciar_envio)
btn_enviar.grid(row=7, column=1, padx=10, pady=20)
# Iniciar a Interface
root.mainloop()