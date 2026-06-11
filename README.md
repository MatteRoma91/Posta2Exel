# Posta2Excel

**Posta2Excel** è un'applicazione desktop cross-platform (Windows e macOS) che estrae automaticamente i contatti da file di backup Outlook (`.pst` e `.ost`), consolidandoli in un file Excel professionale e leggibile.

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

- Estrazione contatti da **tutte le cartelle** del file PST/OST (Inbox, Inviata, Archivio, ecc.)
- Raccolta di **mittenti** e **destinatari** (To, Cc, Bcc) dai metadati ufficiali delle email
- Arricchimento automatico con **telefono** e **società** dalle firme email
- **Deduplicazione intelligente**: una riga per persona, con email e telefoni multipli in celle multilinea
- Filtro automatico degli indirizzi automatici (`noreply@`, `newsletter@`, ecc.)
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
| `libpff-python` | Lettura file PST/OST (`import pypff`) |

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

1. Clicca **Seleziona file .PST / .OST** e scegli il backup Outlook
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
PST/OST → Scansione cartelle → Estrazione metadati → Parsing firme → Consolidamento → Excel
```

### Da dove arrivano i dati

| Campo | Fonte |
|---|---|
| Email, Nome, Cognome | Metadati del messaggio (mittente e destinatari To/Cc/Bcc) |
| Telefono | Firma email del mittente (ultime 15 righe del corpo) |
| Società | Firma del mittente, oppure dominio email aziendale |

### Regole di consolidamento

1. **Durante la scansione** ogni contatto è identificato dalla propria **email** (chiave univoca)
2. **In export** i contatti con stesso Nome e Cognome (entrambi valorizzati) vengono uniti in una riga
3. I contatti senza nome completo restano su **righe separate** (es. email anonime)
4. Email, telefoni e società distinti vengono uniti nella stessa cella, uno per riga

### Cosa NON fa

- **Non** estrae email dal testo del corpo dei messaggi (anti-spam / anti-harvesting)
- **Non** legge la rubrica nativa di Outlook come entità separata
- **Non** unisce contatti con nomi diversi (es. "M. Rossi" e "Mario Rossi" restano righe distinte)

### Filtri automatici

**Indirizzi bloccati** (prefissi robot):

`noreply`, `no-reply`, `newsletter`, `notifiche`, `notifications`, `bounce`, `delivery`, `mailer-daemon`

**Indirizzi conservati** (account funzionali):

`info@`, `commerciale@`, `vendite@`, `amministrazione@`, `ufficio@`, `ufficiotecnico@`

Per account generici come `info@azienda.it`, se il nome visualizzato è vuoto o generico, il software usa il dominio come nome provvisorio (es. `Rossi-idraulica` da `info@rossi-idraulica.it`).

### Parsing firma

Prima di analizzare la firma, il software isola il blocco del messaggio più recente, tagliando le citazioni dei thread Outlook:

- `-----Original Message-----`
- `--- Inizio messaggio inoltrato ---`
- `Da:` / `From:` a inizio riga
- Righe che iniziano con `>`

Le righe contenenti `@` vengono ignorate durante l'estrazione di telefono e società.

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
| `ContactExtractor` | Orchestrazione scansione PST/OST |
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

- Il PST potrebbe contenere principalmente messaggi automatici (filtrati)
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
