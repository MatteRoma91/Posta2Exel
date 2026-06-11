# Posta2Excel

**Posta2Excel** è un'applicazione desktop cross-platform (Windows e macOS) che estrae automaticamente i contatti da archivi email (`.pst`, `.ost`, `.mbox`), consolidandoli in un file Excel professionale e leggibile.

L'elaborazione è **100% offline**: nessun dato viene inviato a server esterni.

---

## Indice

- [Funzionalità](#funzionalità)
- [Requisiti](#requisiti)
- [Installazione](#installazione)
- [Avvio](#avvio)
- [Utilizzo](#utilizzo)
- [Output Excel](#output-excel)
- [Come funziona](#come-funziona)
- [Struttura del progetto](#struttura-del-progetto)
- [Gestione errori](#gestione-errori)
- [Limitazioni note](#limitazioni-note)
- [Risoluzione problemi](#risoluzione-problemi)
- [Licenza](#licenza)

---

## Funzionalità

- Estrazione contatti da **tutte le cartelle** PST/OST (Inbox, Inviata, Archivio, ecc.) e da archivi **MBOX** (Gmail Takeout, Thunderbird, Apple Mail)
- Raccolta di **mittenti** e **destinatari** (To, Cc, Bcc) dai metadati ufficiali delle email
- Arricchimento automatico con **telefono** e **società** dalle firme email (solo per persone reali)
- **Deduplicazione intelligente**: una riga per persona (nome normalizzato) o per azienda (dominio), con email multiple in celle multilinea
- **Esclusione automatica** di newsletter, notifiche transazionali e mittenti robot (Amazon, Google, ecc.)
- Conservazione di **contatti business utili** (`info@`, `supporto@`, `ufficio@`, …) su domini aziendali
- Interfaccia grafica moderna con barra di avanzamento (nessun blocco dell'UI)
- Export Excel formattato con header professionale e testo a capo

---

## Requisiti

| Requisito | Versione consigliata |
|---|---|
| Python | 3.10 o superiore |
| Sistema operativo | Windows 10/11 o macOS |
| Spazio disco | Dipende dalla dimensione del PST/OST |

### Dipendenze Python

| Pacchetto | Uso |
|---|---|
| `customtkinter` | Interfaccia grafica |
| `pandas` | Consolidamento dati e export |
| `openpyxl` | Formattazione Excel |
| `libpff-python` | Lettura file PST/OST (`import pypff`) — **non richiesto** per file `.mbox` |

---

## Installazione

### 1. Clona il repository

```bash
git clone https://github.com/MatteRoma91/Posta2Exel.git
cd Posta2Exel
```

### 2. (Consigliato) Crea un ambiente virtuale

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

### 3. Installa le dipendenze

```bash
pip install customtkinter pandas openpyxl libpff-python
```

#### Windows — fallback se `libpff-python` non compila

Su alcuni sistemi Windows la compilazione di `libpff-python` fallisce per mancanza di Visual C++. In quel caso:

```bash
pip install libpff-python-windows
```

L'import nel codice resta `import pypff`.

#### macOS

La versione `libpff-python` 20231205 su PyPI include wheel precompilati per macOS.

---

## Avvio

```bash
python main.py
```

Si aprirà la finestra **Posta2Excel - Estrattore Contatti Professionale**.

---

## Utilizzo

1. Clicca **Seleziona file .PST / .OST / .MBOX** e scegli l'archivio email
2. Clicca **Seleziona Cartella di Destinazione** per indicare dove salvare l'Excel
3. Clicca **AVVIA ESTRAZIONE**
4. Attendi il completamento: la barra di avanzamento e la label di stato mostrano l'attività corrente
5. Al termine troverai `rubrica_estratta.xlsx` nella cartella scelta

### Suggerimenti

- **Chiudi Outlook** prima di analizzare un file `.ost` ancora in uso
- Per PST di grandi dimensioni l'elaborazione può richiedere diversi minuti: l'interfaccia resta reattiva
- Se un messaggio è corrotto, l'elaborazione continua; i dettagli finiscono in `posta2excel_errors.log`

---

## Output Excel

Il file generato si chiama **`rubrica_estratta.xlsx`** e contiene le colonne:

| Colonna | Contenuto |
|---|---|
| Nome | Nome del contatto |
| Cognome | Cognome del contatto |
| Email | Uno o più indirizzi email (separati da a capo) |
| Telefono | Uno o più numeri (separati da a capo) |
| Società | Ragione sociale o dominio aziendale |

### Formattazione

- Header con sfondo grigio antracite (`#2F3542`), testo bianco in grassetto
- Griglia visibile
- Celle dati con testo a capo e allineamento verticale in alto
- Larghezza colonne adattata automaticamente al contenuto

### Esempio di cella con valori multipli

```
| Email                          |
|--------------------------------|
| mario.rossi@azienda.it         |
| mario@gmail.com                |
```

---

## Come funziona

### Flusso generale

```
PST/OST/MBOX → Scansione messaggi → Classificazione contatti → Parsing firme → Consolidamento → Excel
```

### Da dove arrivano i dati

| Campo | Fonte |
|---|---|
| Email, Nome, Cognome | Metadati del messaggio (mittente e destinatari To/Cc/Bcc) |
| Telefono | Firma email del mittente (ultime 15 righe del corpo, solo persone reali) |
| Società | Firma del mittente (persone) oppure dominio aziendale (solo contatti `info@` / `supporto@`) |

### Regole di consolidamento

1. **Durante la scansione** ogni contatto è classificato come **persona**, **ruolo business** (`info@`, `supporto@`, …) oppure **escluso** (automatico/spam)
2. **Persone**: unite per **token di nome normalizzati** (gestisce varianti come `Matteo Di Benedetto` e `DI BENEDETTO MATTEO`), fondendo tutte le email associate
3. **Ruoli business**: uniti per **dominio aziendale**; in Excel solo la colonna **Società** è valorizzata (Nome e Cognome vuoti)
4. Email, telefoni e società distinti vengono uniti nella stessa cella, uno per riga

### Cosa NON fa

- **Non** estrae email dal testo del corpo dei messaggi (anti-spam / anti-harvesting)
- **Non** legge la rubrica nativa di Outlook come entità separata
- **Non** unisce contatti con nomi diversi (es. "M. Rossi" e "Mario Rossi" restano righe distinte)

### Filtri automatici

**Esclusi** (non compaiono in rubrica):

- Mittenti transazionali e newsletter: `noreply`, `conferma-ordine@`, `promotion@`, `shipment-tracking@`, notifiche Google/Amazon/GitHub, ecc.
- Display name chiaramente automatici (brand, servizi, oggetti email)
- Spam su domini consumer con nomi non plausibili

**Conservati**:

- **Persone reali** con nome plausibile nei metadati (mittenti e destinatari To/Cc)
- **Ruoli business** su domini aziendali: `info@`, `supporto@`, `ufficio@`, `commerciale@`, `vendite@`, `amministrazione@`

Per i ruoli business (es. `info@bperbanca.it`): **Società** = `Bperbanca`, **Nome** e **Cognome** vuoti.

### Parsing nomi (persone)

- Formati supportati: `Nome Cognome`, `COGNOME NOME` (es. `DI BENEDETTO MATTEO` → Matteo / Di Benedetto)
- Particelle italiane: `Di`, `De`, `Della`, `Van`, … restano nel cognome
- **Hint dall'email**: da `matteo.di.benedetto1991@gmail.com` si ricava Matteo / Di Benedetto per disambiguare varianti
- Decodifica header MIME (`=?UTF-8?Q?...?=`) per Gmail Takeout / MBOX
- Rifiuto di nomi che sembrano brand, newsletter o oggetti email

### Campo Società

| Tipo contatto | Società valorizzata da |
|---|---|
| Persona | Solo firma email (ragioni sociali con Srl, Spa, …) — **non** dal dominio |
| Ruolo business (`info@`, `supporto@`, …) | Dominio aziendale (unica colonna compilata oltre all'email) |

### Parsing firma

Prima di analizzare la firma, il software isola il blocco del messaggio più recente, tagliando le citazioni dei thread Outlook:

- `-----Original Message-----`
- `--- Inizio messaggio inoltrato ---`
- `Da:` / `From:` a inizio riga
- Righe che iniziano con `>`

Le righe contenenti `@` vengono ignorate durante l'estrazione di telefono e società. Il telefono viene estratto solo se sulla **stessa riga** compare una parola chiave (`tel`, `cell`, `telefono`, …), per evitare falsi positivi da codici ordine nelle email automatiche.

---

## Struttura del progetto

```
Posta2Exel/
├── backend.py      # Logica di estrazione, parsing e export Excel
├── main.py         # Interfaccia grafica e gestione thread
├── .gitignore
└── README.md
```

### `backend.py`

| Modulo | Responsabilità |
|---|---|
| `ContactExtractor` | Orchestrazione scansione PST/OST/MBOX |
| `_classify_contact` | Classificazione persona / ruolo / automatico / spam |
| `_extract_sender` / `_extract_recipients` | Estrazione anagrafica dai metadati |
| `_parse_signature_for_enrichment` | Telefono e società dalla firma del mittente |
| `_build_export_dataframe` | Groupby e consolidamento per Excel |
| `export_to_excel` | Scrittura e stile del file `.xlsx` |
| `run_extraction` | Entry point chiamato dal thread della GUI |

### `main.py`

| Componente | Responsabilità |
|---|---|
| `Posta2ExcelApp` | Finestra principale customtkinter |
| `_poll_queue` | Aggiornamento thread-safe di progresso e stato |
| `_start_extraction` | Avvio thread daemon per il backend |

---

## Gestione errori

| Livello | Comportamento |
|---|---|
| Singolo messaggio corrotto | Log in `posta2excel_errors.log`, elaborazione continua |
| Cartella inaccessibile | Log, si passa alle cartelle successive |
| File PST/OST non apribile | Messaggio di errore nella GUI |
| Errore export Excel | Messaggio nella GUI + log |

Formato log:

```
2026-06-11 14:32:01 | Folder=Inbox | Index=1204 | Subject=Oggetto mail | Error=...
```

---

## Limitazioni note

- **Destinatari senza header completi**: alcuni messaggi (calendario, Teams, notifiche) potrebbero non avere destinatari nei metadati
- **API pypff**: le binding Python non espongono tutti i metodi C; il software usa un fallback su `transport_headers`
- **Firma con email sulla stessa riga del telefono**: la riga viene scartata per policy anti-harvesting
- **OST in uso**: Outlook deve essere chiuso durante l'analisi di file `.ost` attivi
- **Nomi varianti**: "Mario Rossi" e "M. Rossi" non vengono fusi automaticamente

---

## Risoluzione problemi

### `ModuleNotFoundError: No module named 'pypff'`

```bash
pip install libpff-python
# oppure su Windows:
pip install libpff-python-windows
```

### Errore apertura file PST/OST

- Verifica che il file non sia aperto da Outlook
- Controlla i permessi di lettura sul file
- Per file `.ost`, chiudi completamente Outlook prima dell'estrazione

### Excel vuoto o con pochi contatti

- L'archivio potrebbe contenere principalmente messaggi automatici (esclusi di proposito)
- Con Gmail Takeout, assicurati di selezionare il file `.mbox` corretto
- Verifica che le email abbiano metadati mittente/destinatario leggibili
- Controlla `posta2excel_errors.log` per messaggi saltati

### L'interfaccia non risponde

Non dovrebbe accadere: l'estrazione gira in un thread separato. Se si verifica, riavvia l'applicazione e segnala il problema aprendo una issue su GitHub.

---

## Licenza

Progetto open source. Consultare il repository per eventuali aggiornamenti alla licenza.

---

## Contributi

Pull request e segnalazioni bug sono benvenuti su [GitHub](https://github.com/MatteRoma91/Posta2Exel).
