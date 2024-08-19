import pandas as pd
import pywhatkit
from datetime import datetime
import re
import time

df = pd.read_excel("C:/Users/USER/Documents/sendHappyBirthdayMessage/Alunos.xlsx")
if "DatadeAniversario" not in df.columns:
    print(
        "A coluna 'DatadeAniversario' não foi encontrada. Colunas disponíveis:",
        df.columns,
    )


def limpar_numero(numero):
    # Remove tudo que não é dígito
    numero = re.sub(r"\D", "", numero)
    # Se o número tiver 11 dígitos, remova o terceiro caractere
    if len(numero) == 11:
        numero = numero[:2] + numero[3:]
    # Caso tenha um código de país e DDD, remova o '+55' se presente
    if numero.startswith("55"):
        numero = numero[2:]
    if numero.startswith("+55"):
        numero = numero[3:]
    return numero


def enviar_mensagem(numero, nome):
    mensagem = f"Olá {nome}, a Service Of Well Control gostaria de lhe desejar um feliz aniversário!🎉🎂!! Que seu dia seja repleto de alegria e momentos especiais!!!"
    print(f"Tentando enviar mensagem para {numero}")
    try:
        pywhatkit.sendwhatmsg_instantly(
            numero,
            mensagem,
            wait_time=10,
            tab_close=True,  # Fechar a aba após o envio da mensagem
            close_time=10,  # Tempo para a aba permanecer aberta após o envio da mensagem
        )
        print(f"Mensagem enviada para {nome}, de número {numero}")
    except Exception as e:
        print(
            f"Falha ao enviar mensagem para o {nome} de número: {numero} | código do erro: {e}"
        )


# Exemplo de uso


hoje = datetime.now()
hoje_str = hoje.strftime("%m-%d")
aniversariantes_encontrados = False

for index, row in df.iterrows():
    aniversario = row["DatadeAniversario"]
    if isinstance(aniversario, datetime):
        aniversario_str = aniversario.strftime("%m-%d")
        if aniversario_str == hoje_str:
            aniversariantes_encontrados = True
            numero = str(row["numero"]).strip()
            numero = limpar_numero(numero)
            if not numero.startswith("+"):
                numero = "+55" + numero
            enviar_mensagem(numero, row["nome"])
            time.sleep(2)
if not aniversariantes_encontrados:
    print("Não há aniversariantes hoje.")
    input("Pressione Enter para concluir o processo...")
else:
    print("Processo concluído!")

    input("Pressione Enter para concluir o processo...")
