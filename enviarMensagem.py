import pandas as pd
import pywhatkit
from datetime import datetime

df = pd.read_excel(
    "C:/Users/Gustavo/Documents/TI/Projetos/sendHappyBirthdayMessage/Alunos.xlsx"
)
if "Birthday" not in df.columns:
    print(
        "The 'Birthday' column was not found. Available columns:",
        df.columns,
    )
else:

    def send_message(number, name):
        message = f"Hello {name}, Service Of Well Control would like to wish you a happy birthday! 🎉🎂 May your day be filled with joy and special moments!"
        print(f"Trying to send message to {number}")
        try:
            pywhatkit.sendwhatmsg_instantly(
                number, message, wait_time=8, tab_close=True, close_time=8
            )
            print(f"Message sent to {name}, number {number}")
        except Exception as e:
            print(
                f"Failed to send message to {name} | number: {number} | error code: {e}"
            )

    today = datetime.now()
    today_str = today.strftime("%m-%d")
    for index, row in df.iterrows():
        birthday = row["Birthday"]
        if isinstance(birthday, datetime):
            birthday_str = birthday.strftime("%m-%d")
            if birthday_str == today_str:
                number = str(row["number"])
                if not number.startswith("+"):
                    number = "+" + number.strip()
                send_message(number, row["name"])

    print("Process completed!")

    input("Press Enter to complete the process...")
