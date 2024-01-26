import mysql.connector
from mysql.connector import Error
import time
import tkinter as tk
from tkinter import ttk, messagebox
import schedule
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pystray import MenuItem as item, Icon as TrayIcon, Menu as TrayMenu
from PIL import Image, ImageTk
import tempfile
import os

tray_icon = None

def on_exit(icon, item):
    icon.stop()
    encerrar_programa()

def on_show(icon, item):
    root.deiconify()
    tray_icon.stop()

def minimize_to_tray():
    icon_path = "iconfinder-document03-1622833_121957.ico"
    image = Image.open(icon_path)
    photo = ImageTk.PhotoImage(image)

    # Esconder a janela principal
    root.iconify()

    # Criar janela de bandeja de sistema
    tray_window = tk.Toplevel(root)
    tray_window.overrideredirect(True)
    tray_window.geometry("0x0")
    tray_window.iconphoto(True, photo)

    # Adicionar menu à janela de bandeja
    tray_menu = tk.Menu(tray_window, tearoff=0)
    tray_menu.add_command(label="Abrir", command=on_show)
    tray_menu.add_command(label="Sair", command=on_exit)
    tray_window.config(menu=tray_menu)

# Tamanho inicial da janela
largura_inicial = 800
altura_inicial = 600

# Definir as variáveis no escopo global
label_status = None
label_rodape = None
conexaoLevarCloud = None
cursor = None
execucoes = []

def conectar_banco():
    global conexaoLevarCloud
    global cursor

    tentativas_maximas = 3
    tentativas = 0

    while tentativas < tentativas_maximas:
        try:
            conexaoLevarCloud = mysql.connector.connect(
                host='34.70.96.46',
                user='levar_bi',
                password='@Qpzm1708',
                database='db_levar',
                connect_timeout=60
            )
            cursor = conexaoLevarCloud.cursor()
            print("Conexão bem-sucedida!")
            return conexaoLevarCloud

        except Error as err:
            tentativas += 1
            print(f"Erro de Conexão: {err}")
            print(f"Tentativa de reconexão {tentativas}/{tentativas_maximas}")
            time.sleep(5)  # Aguardar um pouco antes de tentar reconectar

    messagebox.showerror("Erro de Conexão", "Não foi possível reconectar ao banco de dados.")
    encerrar_programa()

def reconectar_e_executar(query):
    tentativas_maximas = 3
    tentativas = 0

    while tentativas < tentativas_maximas:
        try:
            conexao = conectar_banco()
            if conexao:
                cursor = conexao.cursor(dictionary=True)
                cursor.execute(query)
                resultados = cursor.fetchall()
                cursor.close()
                conexao.close()

                # Transformar a lista de tuplas em um dicionário
                resultados_dict = {key: value for key, value in resultados}

                return resultados_dict
            else:
                return {}

        except Error as err:
            print(f"Erro ao executar a query: {err}")
            # Verificar se o erro é relacionado à desconexão
            if "Lost connection" in str(err):
                tentativas += 1
                print(f"Tentativa de reconexão {tentativas}/{tentativas_maximas}")
                time.sleep(5)  # Aguardar um pouco antes de tentar reconectar
                continue
            else:
                break

    print("Não foi possível reconectar ao banco de dados.")
    return {}

def contar_lidos_e_executar_comando():
    global cursor

    if not cursor:
        conectar_banco()

    label_rodape.config(text="Aguarde em execução...")

    select = "SELECT lido, COUNT(*) FROM lt_tbenvio3 GROUP BY lido"
    cursor.execute(select)
    resultados = cursor.fetchall()

    # Inicializar a quantidade de não lidos como 0
    quantidade_nao_lidos = 0

    # Iterar sobre os resultados e verificar se há 'NAO'
    for resultado in resultados:
        lido, quantidade = resultado
        if lido == 'NAO':
            quantidade_nao_lidos = quantidade

    if quantidade_nao_lidos == 0:
        messagebox.showinfo("Informação", "Não há transportes não lidos. Não é necessário executar o comando.")
    elif quantidade_nao_lidos < 50:
        messagebox.showinfo("Informação", "Menos de 50 transportes não lidos. Não é necessário executar o comando.")
    else:
        enviar_comando()

    # Enviar o resultado por e-mail
    enviar_email_resultados(resultados)

    # Remover mensagem "Aguarde em execução"
    label_rodape.config(text="")


def verificar_contagem():
    global cursor

    label_rodape.config(text="Aguarde em execução...")

    try:
        if not cursor or not is_connection_alive(cursor):
            conectar_banco()

        # Ajuste na lógica da query para contar apenas os registros em que a coluna 'lido' é 'NAO'
        select = "SELECT lido, COUNT(*) FROM lt_tbenvio3 GROUP BY lido"
        cursor.execute(select)
        resultados = cursor.fetchall()

        # Remover tabela anterior
        for widget in root.winfo_children():
            if isinstance(widget, ttk.Treeview):
                widget.destroy()

        # Criar nova tabela para o contador de notas
        tree = ttk.Treeview(root, columns=('Lido', 'Quantidade'), show='headings')
        tree.heading('#1', text='Lido')
        tree.heading('#2', text='Quantidade')

        for i, resultado in enumerate(resultados, start=1):
            tree.insert('', i, values=resultado)

        tree.pack(pady=10)

    except mysql.connector.Error as err:
        # Tratar erro de desconexão aqui, se necessário
        if "Lost connection" in str(err):
            messagebox.showerror("Erro na Query", "Erro de conexão perdida. Tentando reconectar...")
            conectar_banco()  # Tentar reconectar
            verificar_contagem()  # Chamar a função novamente após a reconexão
        else:
            messagebox.showerror("Erro na Query", f"Erro ao executar a query: {err}")

    # Remover mensagem "Aguarde em execução"
    label_rodape.config(text="")

def is_connection_alive(cursor):
    try:
        cursor.execute("SELECT 1")  # Execute uma consulta simples para verificar a conexão
        return True
    except mysql.connector.Error as err:
        return False



def enviar_comando():
    global cursor

    if not cursor:
        verificar_contagem()

    label_rodape.config(text="Aguarde em execução...")

    url = "https://api.levartransportes.com.br/command.php"
    payload = {'token': 'seu_token'}

    try:
        response = requests.get(url, params=payload)
        response.raise_for_status()

        data, hora = formatar_data_hora()
        status = f"Status code: {response.status_code}"

        # Remover tabela anterior
        for widget in root.winfo_children():
            if isinstance(widget, ttk.Treeview):
                widget.destroy()

        # Criar nova tabela para o resultado do comando
        tree = ttk.Treeview(root, columns=('Data/Hora', 'Status'), show='headings')
        tree.heading('#1', text='Data/Hora')
        tree.heading('#2', text='Status')

        resultado = f"{data} {hora} - {status}"
        tree.insert('', 0, values=(resultado,))
        tree.pack(pady=10)

        # Registrar a execução na lista
        registrar_execucao(resultado)

    except requests.exceptions.RequestException as e:
        data, hora = formatar_data_hora()
        status = f"Erro: {e}"

        # Criar o corpo do email
        email_body = f"{data} {hora} - {status}"

        # Enviar email
        enviar_email(email_body)

        # Remover tabela anterior
        for widget in root.winfo_children():
            if isinstance(widget, ttk.Treeview):
                widget.destroy()

        # Criar nova tabela para o resultado do comando (erro)
        tree = ttk.Treeview(root, columns=('Data/Hora', 'Status'), show='headings')
        tree.heading('#1', text='Data/Hora')
        tree.heading('#2', text='Status')

        resultado = f"{data} {hora} - {status}"
        tree.insert('', 0, values=(resultado,))
        tree.pack(pady=10)

        # Registrar a execução na lista
        registrar_execucao(resultado)

    # Remover mensagem "Aguarde em execução"
    label_rodape.config(text="")

def formatar_data_hora():
    agora = time.strftime("%Y-%m-%d %H:%M:%S")
    data, hora = agora.split()
    return data, hora

def enviar_email(corpo_email):
    # Configurar informações do servidor SMTP do Gmail
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587  # Porta padrão para TLS
    smtp_user = 'hudsonralves@gmail.com'
    smtp_password = 'pnqg ilap wfcv mnjf'

    # Configurar informações do email
    sender_email = 'ti@levartransportes.com.br'
    receiver_email = 'hudsonralves@gmail.com'
    subject = 'Resultado do Comando'

    # Criar mensagem MIME
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    # Adicionar corpo do email
    message.attach(MIMEText(corpo_email, 'plain'))

    # Configurar conexão SMTP com TLS
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(sender_email, receiver_email, message.as_string())

def enviar_email_resultados(resultados):
    # Configurações do servidor SMTP
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587  # Porta padrão para TLS
    smtp_user = 'hudsonralves@gmail.com'
    smtp_password = 'pnqg ilap wfcv mnjf'

    # Configurar informações do email
    sender_email = 'ti@levartransportes.com.br'
    receiver_email = 'hudsonralves@gmail.com'
    subject = 'Resultado do Comando'

    lidos = 0
    nao_lidos = 0

    # Iterar sobre os resultados e verificar se há 'NAO'
    for resultado in resultados:
        lido, quantidade = resultado
        if lido == 'NAO':
            nao_lidos = quantidade
        if lido == 'SIM':
            lidos = quantidade

    # Construir o corpo do e-mail
    corpo_email = f"""
    Olá,

    Aqui estão os resultados da consulta:

    Quantidade de Lidos: {lidos}
    Quantidade de Não Lidos: {nao_lidos}

    Atenciosamente,
    Jaspion
    """

    # Configurar o e-mail MIME
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.attach(MIMEText(corpo_email, "plain"))

    # Conectar-se ao servidor SMTP
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)

        # Enviar e-mail
        server.sendmail(smtp_user, receiver_email, message.as_string())

    print("E-mail enviado com sucesso!")

def registrar_execucao(resultado):
    global execucoes
    execucoes.append(resultado)

def exibir_execucoes():
    if not execucoes:
        messagebox.showinfo("Lista de Execuções", "Agendador não executado")
    else:
        # Exibir a lista de execuções
        messagebox.showinfo("Lista de Execuções", "\n".join(execucoes))

def iniciar_agendador():
    contar_lidos_e_executar_comando()
    root.after(2 * 60 * 60 * 1000, iniciar_agendador)  # Verificar a cada 2 horas (2 * 60 * 60 * 1000 ms)

    if label_status:
        label_status.config(text="Status: Ativo")

def encerrar_programa():
    global cursor
    if cursor:
        cursor.close()

    global conexaoLevarCloud
    if conexaoLevarCloud:
        conexaoLevarCloud.close()

    root.destroy()

root = tk.Tk()
root.title("Agendador de reset could")

# Configurar o tamanho inicial da janela
largura_inicial = 800
altura_inicial = 600

# Configurar a geometria da janela
root.geometry(f"{largura_inicial}x{altura_inicial}")

label_status = tk.Label(root, text="Status: Inativo", fg="green")
label_status.pack(pady=10)

label_rodape = tk.Label(root, text="", fg="blue")
label_rodape.pack(side=tk.BOTTOM, pady=10)

menu_principal = tk.Menu(root)
root.config(menu=menu_principal)

opcoes_menu = tk.Menu(menu_principal)
menu_principal.add_cascade(label="Opções", menu=opcoes_menu)
opcoes_menu.add_command(label="Contador de Notas e Executar Comando", command=contar_lidos_e_executar_comando)
opcoes_menu.add_command(label="Contador de Notas", command=verificar_contagem)
opcoes_menu.add_command(label="Exibir Execuções", command=exibir_execucoes)
opcoes_menu.add_separator()
opcoes_menu.add_command(label="Sair", command=encerrar_programa)

iniciar_agendador()

minimize_to_tray()

root.protocol("WM_DELETE_WINDOW", lambda: root.iconify())

root.mainloop()