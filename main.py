import pandas as pd
from openpyxl import load_workbook
import smtplib
from email.mime.text import MIMEText
from datetime import datetime

def calculate_age(birthday):
    today = datetime.now()
    age = today.year - birthday.year - ((today.month, today.day) < (birthday.month, birthday.day))
    return age

def send_birthday_greetings(email, name, job, interests, age):
    # Izveidojiet pielāgotu dzimšanas dienas ziņojumu
    greeting_message = f"Dear {name},\n\nHappy Birthday! 🎉🎂 Wishing you an amazing day filled with joy and celebration.\n\n"\
                       f"I also know how passionate you are about {interests}. Your enthusiasm adds a special touch to your personality.\n\n"\
                       f"Congratulations on turning {age}! 🎈🥳 May this new chapter be filled with exciting opportunities and wonderful experiences.\n\n"\
                       f"Best regards,\nKristers"
    try:
        # izveidojiet drōsu savienojumu ar SMTP serveri
        server = smtplib.SMTP_SSL('smtp.epasts.com', 465)
        server.login('TavsEpasts@gmail.com', 'aplikacijas_parole')#pieejama tava epasta iestatījumos

        # izveido un nosūti epastu
        msg = MIMEText(greeting_message)
        msg['Subject'] = 'Birthday Greetings'
        msg['From'] = 'TavsEpasts@gmail.com'
        msg['To'] = email

        server.sendmail('TavsEpasts@gmail.com', [email], msg.as_string())
        server.quit()
        print(f"Birthday greetings sent to {name} ({email}) with updated age {age}")
    except Exception as e:
        print(f"Error sending email to {name} ({email}): {e}")

def main():
    try:
        # nolasa datus no excel failu
        excel_file = 'datubaze1.xlsx'#ievieto šeit path uz savu excel datubāzi
        df = pd.read_excel(excel_file)

        today = datetime.now().strftime('%m-%d')

        # Pārbaudiet katru rindu un, ja nepieciešams, atjauniniet vecumu, pēc tam nosūtiet sveicienus, ja tā ir viņu dzimšanas diena
        for index, row in df.iterrows():
            birthday = row['Dzimšanas datums']
            calculated_age = calculate_age(birthday)

            # Atjauniniet vecumu, ja tas neatbilst aprēķinātajam vecumam
            if row['Vecums'] != calculated_age:
                df.at[index, 'Vecums'] = calculated_age

            # Pārbaudiet, vai tā ir viņu dzimšanas diena
            if today == birthday.strftime('%m-%d'):
                send_birthday_greetings(row['E-pasta adrese'], row['Vārds'], row['Darba amats'], row['Intereses'], df.at[index, 'Vecums'])

        # Saglabājiet atjaunināto DataFrame atpakaļ Excel failā
        df.to_excel(excel_file, index=False, engine='openpyxl')
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
