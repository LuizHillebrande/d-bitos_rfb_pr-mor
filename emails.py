import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configuração do remetente (coloque seu e-mail e senha de app do Gmail)
EMAIL_REMETENTE = "luizhill.dev@gmail.com"
SENHA_APP = "nqlf fgch thrs kpht"

# Leitura do arquivo Excel
arquivo_excel = "mensagens3.xlsx"
df = pd.read_excel(arquivo_excel)

# Iterar sobre as linhas do Excel e enviar e-mails
for index, row in df.iterrows():
    mensagem = str(row.iloc[1]).strip()  # Coluna 2 (Índice 1) - Mensagem
    email_destinatario = str(row.iloc[2]).strip()  # Coluna 3 (Índice 2) - E-mail

    if pd.isna(mensagem) or pd.isna(email_destinatario):
        print(f"⚠️ Linha {index + 2} ignorada: Mensagem ou e-mail ausente.")
        continue

    # Criando o e-mail
    msg = MIMEMultipart()
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = email_destinatario
    msg["Subject"] = "Aviso sobre irregularidade fiscal"
    msg.attach(MIMEText(mensagem, "plain"))

    try:
        # Conectar ao servidor SMTP do Gmail e enviar o e-mail
        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()  # Habilita segurança
        servidor.login(EMAIL_REMETENTE, SENHA_APP)
        servidor.sendmail(EMAIL_REMETENTE, email_destinatario, msg.as_string())
        servidor.quit()
        print(f"✅ E-mail enviado para {email_destinatario}")

    except Exception as e:
        print(f"❌ Erro ao enviar e-mail para {email_destinatario}: {e}")

print("📩 Processo concluído.")
