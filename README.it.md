# outlook-classic-mcp

Connettore MCP locale che collega Claude (Claude desktop e Cowork) a **Outlook Classic per Windows**: posta, calendario, rubrica, attività, note e cartelle, attraverso il modello a oggetti COM di Outlook. Usa la sessione di Outlook già aperta sul PC: nessuna password, nessun server esterno, nessuna porta di rete.

*English version: [README.md](README.md).*

## Che cosa fa

Una volta installato, in Claude desktop o Cowork puoi scrivere cose come "ci sono mail non lette del cliente Rossi?", "metti in calendario una telefonata martedì alle 10", "prepara una bozza di risposta a quest'ultima mail", "salva gli allegati di quella fattura", "cosa ho in agenda questa settimana", e Claude lo fa nel tuo Outlook, con i tuoi account, senza uscire dal PC.

I nomi degli strumenti e dei parametri sono in inglese (il progetto è pensato per essere usato ovunque); Claude li usa da solo, tu gli parli nella tua lingua.

Strumenti esposti (32):

| Area | Strumenti |
|---|---|
| Stato e cartelle | `outlook_status`, `list_folders`, `create_folder` |
| Posta | `list_mail`, `search_mail`, `read_mail` (anche con le intestazioni di trasporto per SPF, DKIM, DMARC), `save_attachments`, `save_mail` (.msg con impronta SHA-256), `mark_mail`, `move_mail` |
| Bozze | `create_draft`, `create_reply_draft` |
| Calendario | `list_events`, `read_event`, `create_event` (anche serie ricorrenti), `update_event` (intera serie o singola occorrenza) |
| Rubrica | `search_contacts`, `list_contacts`, `read_contact`, `create_contact`, `update_contact` |
| Attività | `list_tasks`, `read_task`, `create_task`, `update_task` |
| Note | `list_notes`, `read_note`, `create_note`, `update_note` |
| Avvisi | `send_alert` (unica funzione di invio, destinatario fisso) |
| Cancellazione | `delete_mail`, `delete_event` (solo nel cestino, con conferma dell'oggetto) |

## Regole di sicurezza, scritte nel codice

Il connettore è pensato per essere usato da un'intelligenza artificiale, quindi i limiti non sono raccomandazioni in un prompt ma controlli nel codice, che Claude non può aggirare:

- **Non invia posta.** Le mail nascono come bozze nella cartella Bozze: le rileggi, le firmi se vuoi, le spedisci tu. L'unica eccezione è `send_alert`, che scrive solo dal mittente al destinatario fissati in `config.json` (tipicamente da un tuo indirizzo a un altro tuo indirizzo): Claude sceglie oggetto e testo, mai a chi scrivere.
- **Non cancella in modo definitivo.** `delete_mail` e `delete_event` spostano nel cestino della casella, solo se l'oggetto passato coincide con quello reale, mai dal cestino stesso, mai riunioni con invitati. Si spengono con `"deletion_enabled": false`. Contatti, attività, note e cartelle non si cancellano affatto. Unica eccezione tecnica: la cancellazione di una singola occorrenza di una serie, che Outlook non fa passare dal cestino (lo strumento lo dichiara nella sua descrizione, così Claude lo dice all'utente prima di chiedere conferma).
- **Le mail non lette restano non lette**, anche dopo che Claude le ha lette, ne ha salvato gli allegati, ha preparato una risposta o le ha spostate, e anche quando una di queste operazioni si interrompe a metà.
- **Scrive file solo nelle cartelle che hai elencato** in `save_folders`, senza mai sovrascrivere, e accanto a ogni file lascia un'impronta SHA-256 con data e ora (catena di custodia). Gli allegati salvati hanno nomi ripuliti (niente caratteri invisibili che invertono il testo, niente flussi alternativi) e il contrassegno "scaricato da Internet", così Visualizzazione protetta di Office e SmartScreen continuano a valere.
- **I percorsi si controllano come testo prima di toccare il disco.** Un percorso di rete (`\\host\condivisione`) passato da Claude viene rifiutato senza che Windows si colleghi a quel server. Gli allegati salvati di un tipo che farebbe collegare Explorer a un server appena apri la cartella (`.scf`, `.url`, `.lnk`, `desktop.ini`...) o che partirebbero con un doppio clic ricevono `.txt` in coda al nome; le bozze HTML che puntano a un percorso di rete o locale (`<img src="\\host\...">`) vengono rifiutate. Così si chiudono le vie di fuga dell'hash NTLM che conosciamo; cosa resta a te è nel modello di sicurezza qui sotto.
- **Allegati delle bozze solo da `attach_folders`**, che devi impostare esplicitamente (assente = nessun allegato) e che dovrebbe essere diversa da `save_folders`; le mail archiviate (`.msg`) e le impronte non sono mai allegabili. Così una mail non può far allegare un file privato, o una conversazione archiviata, a una bozza per uno sconosciuto.
- **Per il connettore esistono solo le caselle che elenchi.** Caselle condivise o delegate, archivi e cartelle pubbliche non presenti in `mailboxes` non si raggiungono, nemmeno con il nome esatto, e i loro nomi non vengono mostrati.
- **Nessuna credenziale.** Si aggancia all'Outlook già aperto; se Outlook è chiuso risponde con un errore, non lo avvia e non chiede password.
- **Non tocca regole, impostazioni, account.**
- **Tutte le scritture finiscono in un registro** (`%LOCALAPPDATA%\outlook-classic-mcp\audit.log`: strumento, casella, id, oggetto, e valori vecchi e nuovi quando cambiano indirizzo o telefono di un contatto; mai il testo). Gli avvisi si registrano prima dell'invio.
- **Un Outlook bloccato non fa sorprese dopo.** Se Outlook non risponde (finestra di password o PGP aperta), la richiesta non ancora partita viene annullata e le nuove vengono rifiutate finché Outlook non è libero, invece di accumularsi e partire tutte insieme più tardi.
- **Fasi progressive.** Con `"phase": 1` il connettore legge soltanto e non scrive niente, nemmeno file; 2 aggiunge file su disco, letto/non letto, bozze, contatti, attività, note, cartelle, spostamenti e cancellazione delle mail nel cestino; 3 il calendario; 4 gli avvisi (al massimo `alert.max_per_hour`, predefinito 10, contati su disco così un riavvio non azzera). Si parte da 1 e si sale quando ci si fida.

## Modello di sicurezza: cosa protegge il codice, cosa resta a te

Claude legge mail scritte da sconosciuti, e una mail può contenere testo scritto apposta per pilotare un'intelligenza artificiale ("prompt injection"). Le regole qui sopra valgono anche se Claude venisse ingannato: il connettore non può scrivere a nessuno tranne il destinatario configurato, non può cancellare in modo definitivo, non può scrivere fuori dalle cartelle consentite, non può allegare file presi altrove, non può raggiungere caselle non elencate.

Quello che il codice non può giudicare resta a te:

- **Bozze.** Una bozza può essere indirizzata a chiunque e si sincronizza sul server di posta (Exchange, IMAP) anche se non la invii. Controlla i destinatari prima di premere Invia.
- **Modifiche a contatti, attività, note ed eventi** sono permesse dalle fasi 2 e 3. Sono segnalate al client come distruttive (Claude desktop può chiederti conferma) e i cambi di indirizzo dei contatti si registrano con valore vecchio e nuovo; resta il fatto che "aggiorna l'email di questo contatto" arrivato da una mail è uno schema classico di frode.
- **La cancellazione** è spenta in partenza. Se la attivi, conferma in chat solo elementi che riconosci.
- **Gli allegati salvati restano file scritti da sconosciuti.** I tipi pericolosi vengono neutralizzati e ogni file porta il contrassegno di Internet, ma un documento Office con macro o un PDF con script viene salvato così com'è: aprili come apriresti qualunque allegato.

Hai trovato una vulnerabilità? Segnalala in privato (vedi `SECURITY.md`), non in una issue pubblica.

## Requisiti

- Windows 10 o 11 con **Outlook Classic** (l'Outlook di Microsoft 365 o Office 2016 e successivi, processo `OUTLOOK.EXE`). Il "nuovo Outlook" (`olk.exe`) non espone COM e non funziona.
- Outlook aperto come utente normale (non "come amministratore": i processi con privilegi diversi non si vedono fra loro).
- **Python 3.11 o più recente** da [python.org](https://www.python.org/downloads/windows/), oppure [uv](https://docs.astral.sh/uv/).
- **Claude desktop** per Windows, con Cowork se vuoi usarlo da lì.

Funziona con qualunque account configurato in Outlook: Exchange e Microsoft 365, IMAP, POP, file di dati .pst, e anche server come Kerio Connect tramite il suo connettore KOFF (è l'ambiente in cui è nato e in cui è verificato ogni giorno). La lingua di Outlook non conta: le cartelle predefinite vengono raggiunte per tipo interno, con ricerca per nome in inglese, italiano, tedesco, francese e spagnolo come ripiego.

## Installazione

1. Scarica il repository (Code > Download ZIP, oppure `git clone`) in una cartella stabile, per esempio `C:\Users\<tuo utente>\Claude\outlook-classic-mcp`. Non spostarla dopo l'installazione: il percorso finisce nella configurazione di Claude desktop.

2. Copia `config.example.json` in `config.json` e compila i valori fra `< >` (sotto trovi la spiegazione di ogni chiave). Serve soprattutto il nome esatto dei tuoi archivi di Outlook: lo leggi nel riquadro delle cartelle di Outlook (il nome in grassetto in cima a ogni account), oppure lo stampa lo script di installazione al passo successivo.

3. Apri Outlook. Poi, in **Windows PowerShell** (utente normale), dalla cartella del connettore:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\install.ps1"
   ```

   Lo script crea un ambiente Python isolato (`.venv`), installa `mcp` e `pywin32`, prova il collegamento con Outlook, prova ad avviare il server e registra la voce `outlook` nel file `claude_desktop_config.json` di Claude desktop, conservando gli altri server MCP che hai già.

4. Chiudi **completamente** Claude desktop (icona nell'area di notifica > Esci) e riaprilo. Il connettore compare fra i connettori locali. Prova con: "controlla lo stato di Outlook".

5. Consigliato: lancia la prova automatica, che esercita tutte le funzioni di scrittura con lettere accentate e ripulisce dopo di sé:

   ```powershell
   .\.venv\Scripts\python.exe .\test_accents.py
   ```

   Deve finire con `RESULT: ALL OK`. (Richiede `"phase": 3` o `4` e `attach_folders` impostato; con `"deletion_enabled": false` i controlli di cancellazione vengono saltati.)

   I test unitari non hanno bisogno di Outlook e girano ovunque: `.\.venv\Scripts\python.exe -m unittest -v test_unit`.

Per aggiornare: sostituisci i file, rilancia `install.ps1` se sono cambiate le dipendenze, e riavvia completamente Claude desktop. Il codice del server viene caricato all'avvio dell'app.

## Configurazione (`config.json`)

| Chiave | Obbligatoria | Significato |
|---|---|---|
| `phase` | no (1) | 1 sola lettura (non scrive niente, nemmeno file), 2 anche file, letto/non letto, bozze, contatti, attività, note, cartelle, spostamenti e cancellazione mail, 3 anche calendario, 4 anche `send_alert` |
| `timeout_seconds` | no (40) | attesa massima per una chiamata a Outlook; oltre, errore spiegato ("probabile finestra aperta in Outlook") |
| `max_body_chars` | no (20000) | lunghezza massima dei testi restituiti in una lettura |
| `mailboxes` | **sì** | alias breve → nome esatto dell'archivio in Outlook; gli archivi non elencati non si raggiungono, nemmeno con il nome esatto |
| `default_mailbox` | **sì** | alias usato quando Claude non ne indica uno |
| `save_folders` | **sì** | cartelle (e sottocartelle) in cui il connettore può scrivere file |
| `attach_folders` | no (nessuna) | cartelle (e sottocartelle) da cui `create_draft` può prendere allegati; assente = nessun allegato; tienile separate da `save_folders` |
| `deletion_enabled` | no (false) | attiva `delete_mail` e `delete_event` |
| `alert.sender`, `alert.recipient`, `alert.subject_prefix`, `alert.max_per_hour` | no | mittente (un account di Outlook) e unico destinatario di `send_alert`, prefisso dell'oggetto, limite orario (10); senza la sezione, l'invio è sempre rifiutato |
| `courtesy_pause_ms`, `pause_every_items` | no (10, 3) | pause durante le letture lunghe per lasciare respirare la finestra di Outlook; rilette al volo senza riavvio |

`config.json` è personale ed è escluso dal repository tramite `.gitignore`: non pubblicarlo.

## Come lavora con Claude

- Claude vede ogni strumento con la sua descrizione (la docstring in `server.py`) e gli alias delle caselle nelle istruzioni del server. Gli elenchi restituiscono solo metadati (id, oggetto, mittente, data); il contenuto si legge a parte, così le conversazioni restano leggere.
- Gli id sono gli `EntryID` di Outlook: Claude li riceve dagli elenchi e li passa alle funzioni di lettura e modifica. Un elemento spostato cambia id.
- Le date si scrivono `AAAA-MM-GG` o `AAAA-MM-GG HH:MM`, ora locale.
- Per le serie ricorrenti, `list_events` restituisce una riga per occorrenza con l'id della serie; `update_event` e `delete_event` agiscono sull'intera serie, oppure su una sola occorrenza se ricevono `occurrence_date` (l'inizio esatto riportato dall'elenco).
- Le etichette colorate degli appuntamenti (`phone_call`, `business`, `important`...) sono la vecchia proprietà "Label" di Outlook: alcuni server (per esempio Kerio Connect) le mostrano come colori, altri le ignorano senza danno.

Suggerimento per chi usa Cowork con attività programmate: tieni Outlook aperto, perché il connettore non lo avvia; se un'attività deve avvisarti di qualcosa, `send_alert` è fatta per quello.

## Lezioni tecniche (Outlook COM da Python)

Raccolte durante lo sviluppo, utili a chi vuole modificare il codice.

- **Un solo thread COM.** Gli oggetti di Outlook vanno usati dal thread che li ha creati (STA). Il server è asincrono, quindi tutte le chiamate passano da un thread dedicato con una coda e un timeout (`ComWorker`). Un blocco di Outlook (finestra modale) produce un errore spiegato, non un server appeso.
- **Filtri per data in DASL con date ISO in UTC** (`@SQL="urn:schemas:httpmail:datereceived" >= '2026-11-09 15:00'`). Il formato Jet classico (`[ReceivedTime] >= 'MM/dd/yyyy'`) viene interpretato con le impostazioni regionali di Windows e su un sistema italiano scambia giorno e mese.
- **Date in scrittura marcate come UTC.** pywin32 converte una `datetime` senza fuso da ora locale a UTC spostandola; passandola con `tzinfo=UTC` arriva a Outlook invariata e viene letta come ora locale (`com_date`).
- **Ricorrenze:** `IncludeRecurrences = True`, `Items.Sort("[Start]")`, poi `Restrict` con il filtro DASL, e si itera con l'enumeratore (mai `Count` o indice). Una serie si costruisce con `GetRecurrencePattern()` dopo aver impostato `Start`, `End` e `Duration`, nell'ordine RecurrenceType, DayOfWeekMask o DayOfMonth/MonthOfYear, Interval, PatternStartDate, StartTime, EndTime, Duration, e per ultima la fine. La singola occorrenza si raggiunge con `GetOccurrence(data_inizio)`.
- **Accenti nelle bozze:** impostare `InternetCodepage = 65001` prima di scrivere `HTMLBody`, e scrivere sempre HTML (anche per il "testo semplice"): con `Body` la bozza resta nella pagina di codice del profilo e gli accenti possono corrompersi. Nelle risposte, leggere l'HTML dell'originale **prima** di cambiare la pagina di codice, togliere i `<meta charset>` e inserire il testo nuovo subito dopo `<body>`.
- **Cancellare un evento con `Delete()`, non con `Move()` nel cestino.** Posta eliminata è una cartella di posta: alcuni provider convertono l'appuntamento spostato in un messaggio con allegato .msg (e accenti rotti). `Delete()` è ciò che fa il tasto Canc e lascia un evento ripristinabile.
- **Outlook nasconde gli allegati di "livello 1"** (`.exe`, `.scf`, `.lnk`, `.url`...) al modello a oggetti, anche sulle bozze rilette per id: `save_attachments` non li vede mai. La neutralizzazione con `.txt` copre i tipi che Outlook lascia passare (`desktop.ini`, `.library-ms`, `.iso`...).
- **Dopo `Send()` l'oggetto COM non è più leggibile:** i dati da restituire si prendono prima.
- **Pause di cortesia.** Ogni chiamata gira sul filo principale di Outlook: durante una lettura di centinaia di elementi l'interfaccia resta ferma per l'utente. Una pausa di pochi millisecondi ogni pochi elementi riduce molto il fastidio; non lo elimina, perché due istanze di Outlook sullo stesso profilo non esistono.
- **Errori parlanti.** Ogni eccezione viene convertita in `ToolError` con il messaggio di Outlook (`excepinfo[2]`): Claude riceve "Mailbox 'x' not found" invece di "Error executing tool", e sa spiegarlo all'utente.

## Versioni

Le versioni sono etichettate (`v1.0.0`...) e descritte in `CHANGELOG.md`; `outlook_status` riporta `server_version`. Solo l'ultima versione riceve correzioni.

## Limiti noti

- Serve il PC acceso, l'utente collegato a Windows e Outlook Classic aperto, non elevato.
- Le mail cifrate (PGP, S/MIME) arrivano cifrate: le decifra solo Outlook nella sua finestra.
- La ricerca nel corpo dei messaggi (`search_mail` con `in_body=true`) è lenta sulle cartelle con decine di migliaia di mail: per questo richiede un intervallo di date.
- La cancellazione di una singola occorrenza di una serie non passa dal cestino (comportamento di Outlook).
- Outlook e Claude desktop devono girare con gli stessi privilegi (entrambi utente normale).
- Una richiesta già entrata in Outlook quando scade il tempo non si può interrompere e può completarsi quando Outlook si libera: il messaggio di errore lo dice.
- Due archivi con lo stesso nome visualizzato (tipico di più file di dati chiamati "File di dati di Outlook") vengono rifiutati: rinominali in Outlook.
- `install.ps1` con `uv` può scaricare un Python 3.11 se non ce n'è uno installato.
- Provato ogni giorno su Windows 11 con Outlook 2024 in italiano, Kerio Connect tramite KOFF e Microsoft 365. Altre lingue e altri provider si affidano ai ripieghi descritti sopra: le segnalazioni sono benvenute.

## Struttura del repository

| File | Cosa fa |
|---|---|
| `server.py` | il server MCP: tutte le funzioni e tutti i vincoli, commentato per intero |
| `config.example.json` | modello di configurazione con la spiegazione di ogni chiave |
| `install.ps1` | crea `.venv`, installa le dipendenze, prova il COM, registra il server in Claude desktop |
| `test_accents.py` | prova completa su un Outlook vero: funzioni di scrittura con lettere accentate, più i controlli di sicurezza |
| `test_unit.py` | test unitari che girano ovunque (senza Outlook): percorsi consentiti, nomi dei file, filtro HTML, thread COM, limite degli avvisi, inventario degli strumenti |
| `.github/workflows/tests.yml` | esegue i test unitari su Linux e Windows a ogni push |
| `CHANGELOG.md` | cosa è cambiato in ogni versione |
| `requirements.txt` | le versioni esatte delle librerie con cui il connettore è provato |
| `SECURITY.md` | come segnalare una vulnerabilità in privato |
| `README.md` | questo documento in inglese |
| `.gitignore` | esclude `config.json`, `.venv`, registri e copie di sicurezza |

## Licenza

MIT: vedi `LICENSE`. Il software è fornito così com'è; prima di dare a un'intelligenza artificiale accesso alla propria posta, si consiglia di partire dalla fase 1 e di leggere il registro `audit.log` nei primi giorni.
