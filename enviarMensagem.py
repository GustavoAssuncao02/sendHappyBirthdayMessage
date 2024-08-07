import pandas as pd
import pywhatkit
from datetime import datetime

df = pd.read_excel(
    "C:/Users/Gustavo/Documents/TI/Projetos/sendHappyBirthdayMessage/Alunos.xlsx"
)
if "DatadeAniversario" not in df.columns:
    print(
        "A coluna 'DatadeAniversario' não foi encontrada. Colunas disponíveis:",
        df.columns,
    )
else:

    def enviar_mensagem(numero, nome):
        mensagem = f"Olá {nome}, a Service Of Well Control gostaria de lhe desejar um feliz aniversário! 🎉🎂 Que o seu dia seja repleto de alegria e momentos especiais!"
        print(f"Tentando enviar mensagem para {numero}")
        try:
            pywhatkit.sendwhatmsg_instantly(
                numero, mensagem, wait_time=8, tab_close=True, close_time=8
            )
            print(f"Mensagem enviada para {nome}, de número {numero}")
        except Exception as e:
            print(
                f"Falha ao enviar mensagem para o {nome}| de número: {numero} | código do erro: {e}"
            )

    hoje = datetime.now()
    hoje_str = hoje.strftime("%m-%d")
    for index, row in df.iterrows():
        aniversario = row["DatadeAniversario"]
        if isinstance(aniversario, datetime):
            aniversario_str = aniversario.strftime("%m-%d")
            if aniversario_str == hoje_str:
                numero = str(row["numero"])
                if not numero.startswith("+"):
                    numero = "+" + numero.strip()
                enviar_mensagem(numero, row["nome"])

    print("Processo concluído!")

    input("Pressione Enter para concluir o processo...")
