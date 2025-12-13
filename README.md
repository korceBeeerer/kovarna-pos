# Kovárenský systém - Pokladna

Tabletová aplikace pro evidenci objednávek v kovárně. Zaměstnanec nakliká produkty na tabletu a data se automaticky ukládají do existujícího Excel souboru na PC podle aktuálního dne v měsíci.

## Funkce

- 📱 **Kompaktní textové rozhraní** - Mnoho produktů na jedné obrazovce, rychlé klikání
- 📊 **Načítání produktů z Excelu** - Produkty se automaticky načtou z vašeho Excel souboru
- 🛍️ **Košík** - Přidávání produktů, úprava množství, zobrazení celkové ceny
- 📅 **Automatický zápis podle dne** - Objednávky se ukládají do správného dne v měsíci
- ➕ **Vlastní produkty** - Možnost přidat zakázkové produkty (helma, speciální objednávky)
- 💾 **Automatické zálohování** - Denní zálohy Excelu + tlačítko pro manuální zálohu
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

### 2. Nastavte cestu k vašemu Excel souboru

Otevřete soubor `backend/excel_handler.py` a na **řádku 6** upravte cestu:

```python
EXCEL_FILE = r"C:\Users\martin.korcak\Downloads\Tabulka tržeb.xlsx"
```

Změňte na cestu k **VAŠEMU** Excel souboru, například:
```python
EXCEL_FILE = r"C:\Kovarna\Data\Tržby_2025.xlsx"
```

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

### Práce s produkty

**Produkty z Excelu:**
- Načtou se automaticky při spuštění
- Zobrazují se jako textová okýnka (název + cena)
- **Nelze je smazat** (jsou součástí Excelu)

**Vlastní produkty (zakázky):**
- Klikněte na **"Spravovat produkty"**
- Zadejte název a cenu (např. "Zakázková helma", 2500 Kč)
- Označeny zeleným kolečkem **"V"** (vlastní)
- Můžete je smazat (hover nad kartou → × tlačítko)
- Ukládají se do localStorage tabletu

### Vytvoření objednávky

1. **Naklikejte produkty** - Každé kliknutí přidá 1 ks do košíku
2. **Upravte množství** - Tlačítka **+** / **-** v košíku
3. **Zkontrolujte cenu** - Zobrazuje se celková částka
4. **Uložte** - Tlačítko **"Uložit objednávku"**

Data se automaticky zapíší do Excelu:
- Do správného dne (podle aktuálního data)
- Navýší počet kusů u existujících produktů
- Přidá nové řádky pro vlastní produkty

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
1. Zkontrolujte cestu v `backend/excel_handler.py` na řádku 6
2. Ujistěte se, že soubor existuje
3. Použijte **absolutní cestu** s `r"..."` prefixem

### Produkty se nenačítají

**Zkontrolujte:**
1. Je správná cesta k Excel souboru?
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
- `POST /api/backup` - Vytvoří manuální zálohu Excelu
- `GET /api/backups` - Seznam existujících záloh
- `GET /api/health` - Kontrola stavu serveru

## Důležité poznámky

💾 **Automatické zálohování** - Systém vytváří denní zálohy automaticky, ale můžete kdykoli použít tlačítko "Zálohovat" pro okamžitou zálohu

⚠️ **Nezapomeňte uzavřít Excel** před ukládáním objednávek

⚠️ **Struktura souboru musí zůstat stejná** - 50 řádků na den

⚠️ **Vlastní produkty** se ukládají pouze v tabletu (localStorage), ne v Excelu - pokud se ale prodají, zapíší se do Excelu

## Podpora

Pokud narazíte na problém:
1. Zkontrolujte konzoli serveru (CMD okno)
2. Zkontrolujte konzoli prohlížeče (F12 → Console)
3. Zkontrolujte cestu k Excel souboru v `excel_handler.py`
