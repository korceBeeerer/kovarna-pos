# Kovárenský systém - Pokladna

Tabletová aplikace pro evidenci objednávek v kovárně. Zaměstnanec nakliká produkty na tabletu a data se automaticky ukládají do existujícího Excel souboru na PC podle aktuálního dne v měsíci.

## Funkce

- 📱 **Kompaktní textové rozhraní** - Mnoho produktů na jedné obrazovce, rychlé klikání
- 📊 **Načítání produktů z Excelu** - Produkty se automaticky načtou z vašeho Excel souboru
- 🛍️ **Košík** - Přidávání produktů, úprava množství, zobrazení celkové ceny
- 💱 **Dual měna** - Zobrazení cen v CZK i EUR s nastavitelným kurzem
- 📅 **Automatický zápis podle dne** - Objednávky se ukládají do správného dne v měsíci
- ➕ **Zakázkové produkty** - Možnost přidat speciální zakázku přímo do košíku (helma, nestandardní práce)
- ↩️ **Storno** - Odečtení chybně zadaných nebo vrácených položek z Excelu
- 💾 **Automatické zálohování** - Denní zálohy Excelu + tlačítko pro manuální zálohu
- ⚙️ **Nastavení z prohlížeče** - Excel soubor nahrajete a stáhnete přímo v aplikaci
- 🌐 **Lokální síť** - Komunikace přes WiFi, **bez potřeby internetu**
- ✅ **Real-time status** - Kontrola připojení mezi tabletem a PC

## Jak to funguje

### Struktura Excel souboru

Váš Excel soubor musí mít strukturu, kde:
- List **"Tržby"** obsahuje data
- Každý den má **50 řádků** (blok produktů)
- **Den 1** začíná na řádku 1, **Den 2** na řádku 51, **Den 3** na řádku 101, atd.
- **DŮLEŽITÉ:** Každý měsíc má jiný počet dní, proto musíte připravit Excel s odpovídajícím počtem bloků:
  - **Leden** (31 dní): 31 bloků = 1550 řádků
  - **Únor** (28/29 dní): 28/29 bloků = 1400/1450 řádků
  - **Březen** (31 dní): 31 bloků = 1550 řádků
  - **Duben** (30 dní): 30 bloků = 1500 řádků
  - atd.

**Sloupce:**
- **A**: Číslo
- **B**: Produkt (název)
- **C**: Cena
- **D**: Kusy (sem se zapisuje počet prodaných kusů)
- **E**: Celk. (vzorec =D×C)

**Příklad:**
```
Řádek 1:  | Produkt      | Cena  | Kusy | Celk.
Řádek 2:  | Mince Kl     | 25    | 0    | =D2*C2
Řádek 3:  | Podkova      | 125   | 0    | =D3*C3
...
Řádek 50: (prázdný nebo celkem)
Řádek 51: | Produkt      | Cena  | Kusy | Celk.  <- Den 2
Řádek 52: | Mince Kl     | 25    | 0    | =D52*C52
...
```

Když dnes (např. 15. ledna) někdo koupí **2× Mince Kl**, aplikace:
1. Najde blok pro den 15 (řádek 701 = 1 + 14×50)
2. Najde řádek s produktem "Mince Kl" v tomto bloku
3. Navýší hodnotu ve sloupci **Kusy** o 2

## Instalace a spuštění

### 1. Nainstalujte Python

Stáhněte a nainstalujte Python 3.8 nebo novější z [python.org](https://www.python.org/downloads/)

Při instalaci zaškrtněte **"Add Python to PATH"**

### 2. Nahrajte váš Excel soubor

Excel se vybírá přímo v aplikaci, soubory Pythonu není potřeba upravovat:

1. Spusťte server a otevřete aplikaci v prohlížeči (na PC i na tabletu)
2. Klikněte na **⚙️ Nastavení** v hlavičce (když soubor nebyl nalezen, okno se otevře samo)
3. Vyberte Excel soubor (.xlsx) a klikněte na **Nahrát a použít**
4. Server ověří, že soubor obsahuje list **"Tržby"**, uloží ho do složky `data/` v aplikaci a znovu načte produkty

**Důležité:** aplikace pak zapisuje do **nahrané kopie** ve složce `data/`, ne do původního souboru,
který jste vybrali. Aktuální data otevřete přímo z `data/` na PC se serverem, nebo je stáhněte
tlačítkem **⬇ Stáhnout aktuální Excel** v Nastavení.

- Když nahrajete soubor se stejným názvem, původní se nejdřív zazálohuje do `data/backup/` (`..._PREPSANO_...`)
- Nastavení se uloží do `backend/config.json` a zůstane platné i po restartu serveru
- Prohlížeč neumí serveru sdělit cestu k souboru, proto se soubor nahrává

### 3. Nainstalujte závislosti

Otevřete příkazový řádek (CMD) ve složce `backend/` a spusťte:

```bash
pip install -r requirements.txt
```

### 4. Spusťte server

**Rychlý způsob** - dvakrát klikněte na:
```
start_server.bat
```

**Nebo ručně:**
```bash
cd backend
python server.py
```

Server se spustí a vypíše IP adresu:
```
==================================================
KOVÁRENSKÝ SYSTÉM - SERVER
==================================================

Server běží na:
  - Lokálně: http://localhost:5000
  - V síti: http://192.168.0.189:5000

Pro zastavení serveru stiskněte Ctrl+C
==================================================
```

### 5. Připojte tablet

1. Ujistěte se, že tablet i PC jsou na **stejné WiFi síti**
2. Na tabletu otevřete webový prohlížeč
3. Zadejte IP adresu z konzole, například:
   ```
   http://192.168.0.189:5000
   ```

## Použití aplikace

### Nastavení kurzu EUR

V hlavičce aplikace najdete pole **"1 EUR ="**:
- Nastavte aktuální kurz (výchozí: 25 Kč)
- Všechny ceny v košíku se automaticky přepočítají
- Zobrazuje se **CZK i EUR** pro každou položku i celkovou částku

### Produkty z Excelu

- Načtou se automaticky při spuštění ze souboru
- Zobrazují se jako textová okýnka (název + cena v CZK)
- **Nelze je mazat** (jsou pevná součást Excelu)
- Kliknutím přidáte 1 ks do košíku

### Zakázkové produkty (speciální objednávky)

Pro nestandardní zakázky (např. "Zakázková helma"):
1. Klikněte na **"+ Přidat zakázku"**
2. Zadejte název, cenu (Kč) a množství
3. Položka se přidá **přímo do košíku**
4. **NEPŘIDÁVÁ** se do seznamu produktů (je to jednorázová zakázka)

### Vytvoření objednávky

1. **Naklikejte produkty** - Každé kliknutí přidá 1 ks do košíku
2. **Nebo přidejte zakázku** - Tlačítko "+ Přidat zakázku" pro speciální položky
3. **Upravte množství** - Tlačítka **+** / **-** v košíku
4. **Zkontrolujte cenu** - Zobrazuje se v CZK i EUR
5. **Uložte** - Tlačítko **"Uložit objednávku"**

Data se automaticky zapíší do Excelu:
- Do správného dne (podle aktuálního data)
- Navýší počet kusů u existujících produktů
- Přidá nové řádky pro zakázkové produkty

### Storno (odečtení položek)

Když je potřeba vrátit už uloženou objednávku nebo opravit překlep:
1. Klikněte na **Storno** pod produkty. Košík zčervená a přepne se do režimu storna
2. Naklikejte produkty, které se mají odečíst (stejně jako při prodeji, včetně **+ Přidat zakázku** pro zakázkové položky)
3. Klikněte na **Stornovat položky** a potvrďte
4. Množství se odečte ze sloupce **Kusy** v dnešním dni
5. Tlačítkem **Zpět na prodej** se vrátíte do běžného režimu

Storno je bezpečné:
- Stornovat jde jen produkty, které jsou dnes v Excelu zapsané
- Počet kusů nemůže klesnout pod nulu
- Když nejde stornovat některá položka, neprovede se nic a aplikace řekne proč

### Zálohování

**Automatická záloha:**
- Vytvoří se automaticky při prvním uložení každého dne
- Zálohy se ukládají do složky `backup/` vedle Excel souboru
- Název: `Tabulka_tržeb_ZALOHA_2025-12-13_14-30-00.xlsx`
- Pokud dnes již byla záloha vytvořena, další se neprovede (aby nedošlo ke zbytečným kopiím)

**Manuální záloha:**
- Klikněte na tlačítko **💾 Zálohovat** v hlavičce aplikace
- Záloha se vytvoří okamžitě s aktuálním časovým razítkem
- Můžete ji použít kdykoliv před důležitými změnami

**Kde najdu zálohy:**
- Ve stejné složce jako váš Excel soubor → podsložka `backup/`
- Např: `C:\Users\martin.korcak\Downloads\backup\`

### Prohlížení dat

Otevřete váš Excel soubor a podívejte se na list **"Tržby"**:
- Najděte blok pro dnešní den
- Ve sloupci **Kusy** uvidíte prodané množství
- Sloupec **Celk.** se přepočítá automaticky (vzorec)

## Řešení problémů

### Firewall blokuje připojení

Pokud tablet nemůže přistoupit k PC:

1. **Povolte Python v Windows Firewall:**
   - Otevřete `wf.msc`
   - Příchozí pravidla → Nové pravidlo
   - Port → TCP → 5000 → Povolit

2. **Nebo vypněte firewall dočasně** (jen pro test):
   - Windows Defender Firewall → Zapnout/vypnout → Vypnout pro soukromé sítě

### Excel soubor nenalezen

```
CHYBA: Excel soubor nenalezen
```

**Řešení:**
1. Otevřete **⚙️ Nastavení** v aplikaci a nahrajte Excel soubor znovu
2. Zkontrolujte, že nahraný soubor nikdo nesmazal ze složky `data/`

### Produkty se nenačítají

**Zkontrolujte:**
1. Je v Nastavení nahraný správný Excel soubor?
2. Je list pojmenován **"Tržby"**?
3. Je struktura správná? (Sloupec B = Produkt, C = Cena)

### Data se neukládají

**Možné příčiny:**
1. **Excel je otevřený** - Zavřete ho před ukládáním
2. **Soubor je jen pro čtení** - Zkontrolujte oprávnění
3. **Nesprávná struktura** - Ujistěte se, že bloky dnů jsou správně

## Technické informace

- **Frontend**: HTML, CSS, Vanilla JavaScript
- **Backend**: Python 3.8+, Flask
- **Excel**: openpyxl
- **Komunikace**: REST API přes HTTP (lokální síť)
- **Offline režim**: Aplikace funguje **bez internetu**

### API Endpointy

- `GET /api/products` - Načte produkty z Excelu
- `POST /api/order` - Uloží objednávku do Excelu
- `POST /api/storno` - Stornuje položky (odečte je z dnešního dne)
- `GET /api/config` - Vrátí aktuální Excel soubor
- `POST /api/config/upload` - Nahraje nový Excel soubor (multipart, pole `file`)
- `GET /api/excel/download` - Stáhne aktuální Excel soubor
- `POST /api/config` - Nastaví cestu k souboru přímo na PC se serverem (`{"excel_file": "..."}`)
- `POST /api/backup` - Vytvoří manuální zálohu Excelu
- `GET /api/backups` - Seznam existujících záloh
- `GET /api/health` - Kontrola stavu serveru

## Důležité poznámky

💾 **Automatické zálohování** - Systém vytváří denní zálohy automaticky, ale můžete kdykoli použít tlačítko "Zálohovat" pro okamžitou zálohu

⚠️ **Nezapomeňte uzavřít Excel** před ukládáním objednávek

⚠️ **Struktura souboru musí zůstat stejná** - 50 řádků na den

⚠️ **Zakázkové produkty** se přidávají přímo do košíku a při uložení objednávky se zapíší do Excelu jako nový řádek v příslušném dni. Zapisují se jen do prázdných řádků nad řádkem „Celkem“. Když v daném dni dojdou, objednávka se neuloží a aplikace ohlásí chybu (přidejte v Excelu prázdné řádky nad „Celkem“)

## Podpora

Pokud narazíte na problém:
1. Zkontrolujte konzoli serveru (CMD okno)
2. Zkontrolujte konzoli prohlížeče (F12 → Console)
3. Zkontrolujte nahraný Excel soubor v **⚙️ Nastavení**
