# Dzimšanas dienu Apsveikumu Sūtītājs

## Projekta Uzdevums
Šis projekts ir izveidots, lai automatizētu dzimšanas dienu apsveikumu sūtīšanu, izmantojot datus no Excel datubāzes. Programma nolasīs dzimšanas dienu datus, aprēķinās jauno vecumu, atjaunos datus Excel failā un nosūtīs personalizētus apsveikumus pa e-pastu tiem, kuriem ir dzimšanas diena.

## Python Bibliotēkas
- **pandas**: Tiek izmantota datu apstrādei un tabulu manipulācijai.
- **openpyxl**: Nodrošina iespēju lasīt un rakstīt Excel failus, nepieciešamo datu nolasīšanai un saglabāšanai.
- **smtplib** un **email.mime.text**: Izmantojot šīs bibliotēkas, tiek veidots un nosūtīts e-pasts ar dzimšanas dienas apsveikumu.
- **datetime**: Tiek izmantots laika apstrādei, it īpaši, lai aprēķinātu vecumu un salīdzinātu dzimšanas dienas.

## Programmatūras Izmantošana
1. Ievietojiet savu Excel datubāzi ar dzimšanas dienu datiem un citiem nepieciešamajiem laukiem.(## Excel Faila Kollonu Nosaukumi)
2. Konfigurējiet programmu, iestatot ceļu uz Excel failu un e-pasta serveri (SMTP serveri) ar saviem piekļuves datiem.
3. Palaidiet programmu, un tā automātiski atjaunos vecumu un nosūtīs apsveikumus tiem, kuriem ir dzimšanas diena.

## Veiktās Izmaiņas Kodā
- Izveidota jauna funkcija `calculate_age`, lai aprēķinātu vecumu.
- Izveidota funkcija `send_birthday_greetings`, lai izveidotu un nosūtītu dzimšanas dienas apsveikumus.
- Pārbaudīts katrs ieraksts Excel failā, atjauninot vecumu un nosūtot apsveikumus, ja nepieciešams.
- Izveidota funkcija `main`, kas palaista, kad programma tiek izsaukta, lai nodrošinātu galveno darbību plūsmu.

## Excel Faila Kollonu Nosaukumi
1. **Vārds**: Persona vārds.
2. **Dzimšanas datums**: Dzimšanas dienas datums.(excel kollonas tips jānomaina uz "Short Date")
3. **E-pasta adrese**: Persona e-pasta adrese.
4. **Darba amats**: Persona darba amats.
5. **Intereses**: Persona intereses.
6. **Vecums**: Persona vecums, kas tiek automātiski atjaunināts.


## Koda Setup Process
1. **Iegūstiet kodu**: Nokopējiet vai lejupielādējiet projektu uz savu datoru.
2. **Excel faila aizpildīšana**: Izveidojiet savu vai projekta sniegto excel failu un aizpildiet to ar vajadzīgo informāciju.
3. **Instalējiet nepieciešamās bibliotēkas**: Pārliecinieties, ka jums ir uzstādītas nepieciešamās Python bibliotēkas. To var izdarīt, izpildot šo komandu terminālī vai komandrindā:

    ```bash
    pip install pandas openpyxl
    ```

4. **Konfigurējiet programmu**:
    - Atveriet failu `main.py` un iestatiet ceļu uz savu Excel failu, rediģējot šo rindiņu:

        ```python
        excel_file = 'ceļš/uz/jūsu/datubāzi.xlsx'
        ```

    - Turpinot šajā pašā failā, iestatiet savas e-pasta servera datus un piekļuves informāciju:

        ```python
        server = smtplib.SMTP_SSL('smtp.epasts.com', 465)
        server.login('TavsEpasts@gmail.com', 'aplikacijas_parole')#pieejama tava epasta iestatījumos
        ```

	```python
        server.sendmail('TavsEpasts@gmail.com', [email], msg.as_string())
        ```

5. **Palaidiet programmu**:
    - Pirms koda palaišanas pārliecinaties, ka excel fails ir aizvērts.
    - Atveriet termināli vai komandrindu un pārliecinieties, ka jums ir nokļuvis(-usi) projekta direktorijā.
    - Izpildiet programmu, izmantojot šo komandu:

        ```bash
        python main.py
        ```

6. **Pārbaudiet rezultātus**:
    - Pēc programmas pabeigšanas pārbaudiet Excel failu, lai redzētu atjauninātos datus un nosūtītos dzimšanas dienas apsveikumus.

**Piezīme**: Lai programma strādātu, pārliecinieties, ka ir iestatīti pareizie e-pasta servera dati un ceļš uz Excel failu ir pareizs.
