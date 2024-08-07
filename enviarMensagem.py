import pandas as pd
import pywhatkit
from datetime import datetime

df = pd.read_excel(
    "C:/Users/Gustavo/Documents/TI/Projetos/sendHappyBirthdayMessage/Alunos.xlsx"
)
if "DatadeAniversario" not in df.columns:
    print(
        "The 'Birthday' column was not found. Available columns:",
        df.columns,
    )
else:

    def send_message(numero, nome):
        message = f"Hello {nome}, Service Of Well Control would like to wish you a happy birthday! 🎉🎂 May your day be filled with joy and special moments!"
        print(f"Trying to send message to {numero}")
        try:
            pywhatkit.sendwhatmsg_instantly(
                numero, message, wait_time=8, tab_close=True, close_time=8
            )
            print(f"Message sent to {nome}, numero {numero}")
        except Exception as e:
            print(
                f"Failed to send message to {nome} | numero: {numero} | error code: {e}"
            )

    today = datetime.now()
    today_str = today.strftime("%m-%d")
    for index, row in df.iterrows():
        birthday = row["DatadeAniversario"]
        if isinstance(birthday, datetime):
            birthday_str = birthday.strftime("%m-%d")
            if birthday_str == today_str:
                numero = str(row["numero"])
                if not numero.startswith("+"):
                    numero = "+" + numero.strip()
                send_message(numero, row["nome"])

    print("Process completed!")

    input("Press Enter to complete the process...")
