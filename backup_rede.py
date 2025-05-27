import os
import shutil
import zipfile
import smtplib
import datetime
import logging
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from rich.console import Console

# -------------------- SETUP --------------------

console = Console()

# Carrega variáveis do .env
load_dotenv()
email_remetente = os.getenv("EMAIL_REMETENTE")
senha_email = os.getenv("SENHA_APP")

# Lista de (caminho, nome identificador)
origens = [
    (r"\\NTB-TIC002\Users\User\Desktop\BACKUP", "pops-ti"),
    #(r"\\NTB-FIN-01\Users\Financeiro\Desktop\BACKUP", "Financeiro"),
    #(r"\\NTB-SUP-REC\Users\Rafaela\Desktop\BACKUP", "Rafa"),
    #(r"\\NTB-RH-01\Users\Norma\Desktop\BACKUP", "Norma"),
    #(r"\\NTB-ALM-01\Users\Almoxarifado\Desktop\BACKUP", "Julinha"),
    #(r"\\NTB-FAR-01\Users\farmacia\Desktop\BACKUP", "Farmacia")
    # Adicione mais assim:
    # (r"\\OutroPC\Compartilhamento\RH", "rh"),
]

destino_base = r"\\NTB-BKP-01\Users\BACKUPDG\Desktop\BACKUP"
log_path = os.path.join(destino_base, "log_backup.txt")

emails_destinatarios = [
    "tic@hospitalubarana.com.br",
    "tic.trainee@hospitalubarana.com.br"
]

logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s'
)

# -------------------- FUNÇÕES --------------------

def compactar_pasta(origem, destino_zip):
    with zipfile.ZipFile(destino_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for pasta_raiz, _, arquivos in os.walk(origem):
            for arquivo in arquivos:
                caminho_completo = os.path.join(pasta_raiz, arquivo)
                arcname = os.path.relpath(caminho_completo, origem)
                zipf.write(caminho_completo, arcname)

def enviar_email(sucesso, mensagem):
    assunto = "✅ Backup Concluído com Sucesso" if sucesso else "❌ Falha no Backup"
    msg = MIMEMultipart()
    msg['From'] = email_remetente
    msg['To'] = ", ".join(emails_destinatarios)
    msg['Subject'] = assunto
    msg.attach(MIMEText(mensagem, 'plain'))

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(email_remetente, senha_email)
            smtp.send_message(msg)
        console.print("[bold green]✔ E-mail de notificação enviado com sucesso.[/bold green]")
    except Exception as e:
        console.print(f"[bold red]Erro ao enviar e-mail:[/bold red] {e}")
        logging.error(f"Erro ao enviar e-mail: {e}")

# -------------------- EXECUÇÃO --------------------

def executar_backup():
    data_atual = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    logs = []

    for origem, nome in origens:
        zip_destino = os.path.join(destino_base, f"{nome}_backup_{data_atual}.zip")

        try:
            if not os.path.exists(origem):
                raise FileNotFoundError(f"Pasta não encontrada: {origem}")

            compactar_pasta(origem, zip_destino)
            msg = f"[{nome.upper()}] Backup realizado com sucesso: {zip_destino}"
            console.print(f"[green]{msg}[/green]")
            logging.info(msg)
            logs.append(msg)
        except Exception as e:
            erro = f"[{nome.upper()}] ERRO: {str(e)}"
            console.print(f"[red]{erro}[/red]")
            logging.error(erro)
            logs.append(erro)

    sucesso_geral = all("sucesso" in log.lower() for log in logs)
    enviar_email(sucesso_geral, "\n".join(logs))

if __name__ == "__main__":
    executar_backup()
