# outlook-classic-mcp

Connettore MCP locale che collega Claude (Claude desktop e Cowork) a **Outlook Classic per Windows**: posta, calendario, rubrica, attività, note e cartelle, attraverso il modello a oggetti COM di Outlook. Usa la sessione di Outlook già aperta sul PC: nessuna password, nessun server esterno, nessuna porta di rete.

> **English summary.** A local MCP server (Python, pywin32) that lets Claude desktop and Cowork work with classic Outlook for Windows through COM automation: read and search mail (unread mail stays unread), save attachments and .msg files with SHA-256 fingerprints, create drafts and reply drafts, read/create/modify calendar events including recurring series and single occurrences, manage contacts, tasks, notes and folders. Safety rules are enforced in code, not in prompts: no sending (except one alert tool with a fixed, configured recipient), no permanent deletion (trash only, with subject confirmation), no credentials. Everything personal lives in `config.json`. The code, tool names and documentation are in Italian; Claude handles that transparently.

## Che cosa fa

Una volta installato, in Claude desktop o Cowork puoi scrivere cose come "ci sono mail non lette del cliente Rossi?", "metti in calendario una telefonata martedì alle 10", "prepara una bozza di risposta a quest'ultima mail", "salva gli allegati di quella fattura", "cosa ho in agenda questa settimana", e Claude lo fa nel tuo Outlook, con i tuoi account, senza uscire dal PC.

Strumenti esposti (32):

| Area | Strumenti |
|---|---|
| Stato e cartelle | `stato_outlook`, `elenca_cartelle`, `crea_cartella` |
| Posta | `elenca_mail`, `cerca_mail`, `leggi_mail` (anche intestazioni di trasporto per SPF, DKIM, DMARC), `salva_allegati`, `salva_mail` (.msg con impronta SHA-256), `segna_mail`, `sposta_mail` |
| Bozze | `crea_bozza`, `crea_bozza_risposta` |
| Calendario | `elenca_eventi`, `leggi_evento`, `crea_evento` (anche serie ricorrenti), `modifica_evento` (intera serie o singola occorrenza) |
| Rubrica | `cerca_contatti`, `elenca_contatti`, `leggi_contatto`, `crea_contatto`, `modifica_contatto` |
| Attività | `elenca_attivita`, `leggi_attivita`, `crea_attivita`, `modifica_attivita` |
| Note | `elenca_note`, `leggi_nota`, `crea_nota`, `modifica_nota` |
| Avvisi | `invia_avviso` (unica funzione di invio, destinatario fisso) |
| Cancellazione | `elimina_mail`, `elimina_evento` (solo nel cestino, con conferma dell'oggetto) |

## Regole di sicurezza, scritte nel codice

Il connettore è pensato per essere usato da un'intelligenza artificiale, quindi i limiti non sono raccomandazioni in un prompt ma controlli nel codice, che Claude non può aggirare:

- **Non invia posta.** Le mail nascono come bozze nella cartella Bozze: le rileggi, le firmi se vuoi, le spedisci tu. L'unica eccezione è `invia_avviso`, che scrive solo dal mittente al destinatario fissati in `config.json` (tipicamente da un tuo indirizzo a un altro tuo indirizzo): Claude sceglie oggetto e testo, mai a chi scrivere.
- **Non cancella in modo definitivo.** `elimina_mail` ed `elimina_evento` spostano nel cestino della casella, solo se l'oggetto passato coincide con quello reale, mai dal cestino stesso, mai riunioni con invitati. Si spengono con `"cancellazione_abilitata": false`. Contatti, attività, note e cartelle non si cancellano affatto. Unica eccezione tecnica: la cancellazione di una singola occorrenza di una serie, che Outlook non fa passare dal cestino (lo strumento lo dichiara nella sua descrizione, così Claude lo dice all'utente prima di chiedere conferma).
- **Le mail non lette restano non lette**, anche dopo che Claude le ha lette, ne ha salvato gli allegati, ha preparato una risposta o le ha spostate.
- **Scrive file solo nelle cartelle che hai elencato** in `cartelle_salvataggio`, senza mai sovrascrivere, e accanto a ogni file lascia un'impronta SHA-256 con data e ora (catena di custodia).
- **Nessuna credenziale.** Si aggancia all'Outlook già aperto; se Outlook è chiuso risponde con un errore, non lo avvia e non chiede password.
- **Non tocca regole, impostazioni, account.**
- **Tutte le scritture finiscono in un registro** (`%LOCALAPPDATA%\outlook-classic-mcp\audit.log`: strumento, casella, id, oggetto; mai il contenuto).
- **Fasi progressive.** Con `"fase": 1` il connettore legge soltanto; 2 aggiunge bozze e scritture non distruttive; 3 il calendario; 4 gli avvisi. Si parte da 1 e si sale quando ci si fida.

## Requisiti

- Windows 10 o 11 con **Outlook Classic** (l'Outlook di Microsoft 365 o Office 2016 e successivi, processo `OUTLOOK.EXE`). Il "nuovo Outlook" (`olk.exe`) non espone COM e non funziona.
- Outlook aperto come utente normale (non "come amministratore": i processi con privilegi diversi non si vedono fra loro).
- **Python 3.11 o più recente** da [python.org](https://www.python.org/downloads/windows/), oppure [uv](https://docs.astral.sh/uv/).
- **Claude desktop** per Windows, con Cowork se vuoi usarlo da lì.

Funziona con qualunque account configurato in Outlook: Exchange e Microsoft 365, IMAP, POP, file di dati .pst, e anche server come Kerio Connect tramite il suo connettore KOFF (è l'ambiente in cui è nato e in cui è verificato ogni giorno).

## Installazione

1. Scarica il repository (Code > Download ZIP, oppure `git clone`) in una cartella stabile, per esempio `C:\Users\<tuo utente>\Claude\outlook-classic-mcp`. Non spostarla dopo l'installazione: il percorso finisce nella configurazione di Claude desktop.

2. Copia `config.example.json` in `config.json` e compila i valori fra `< >` (sotto trovi la spiegazione di ogni chiave). Serve soprattutto il nome esatto dei tuoi archivi di Outlook: lo leggi nel riquadro delle cartelle di Outlook (il nome in grassetto in cima a ogni account), oppure lo stampa lo script di installazione al passo successivo.

3. Apri Outlook. Poi, in **Windows PowerShell** (utente normale), dalla cartella del connettore:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\installa.ps1"
   ```

   Lo script crea un ambiente Python isolato (`.venv`), installa `mcp` e `pywin32`, prova il collegamento con Outlook, prova ad avviare il server e registra la voce `outlook` nel file `claude_desktop_config.json` di Claude desktop, conservando gli altri server MCP che hai già.

4. Chiudi **completamente** Claude desktop (icona nell'area di notifica > Esci) e riaprilo. Il connettore compare fra i connettori locali. Prova con: "controlla lo stato di Outlook".

5. Consigliato: lancia la prova automatica, che esercita tutte le funzioni di scrittura con lettere accentate e ripulisce dopo di sé:

   ```powershell
   .\.venv\Scripts\python.exe .\prova_accenti.py
   ```

   Deve finire con `ESITO: TUTTO OK`. (Richiede `"fase": 3` o `4`; con `"cancellazione_abilitata": false` i controlli di cancellazione vengono saltati.)

Per aggiornare: sostituisci i file, rilancia `installa.ps1` se sono cambiate le dipendenze, e riavvia completamente Claude desktop. Il codice del server viene caricato all'avvio dell'app.

## Configurazione (`config.json`)

| Chiave | Obbligatoria | Significato |
|---|---|---|
| `fase` | no (1) | 1 sola lettura, 2 anche bozze, contatti, attività, note, cartelle, spostamenti e cancellazione mail, 3 anche calendario, 4 anche `invia_avviso` |
| `timeout_secondi` | no (40) | attesa massima per una chiamata a Outlook; oltre, errore spiegato ("probabile finestra aperta in Outlook") |
| `max_caratteri_corpo` | no (20000) | lunghezza massima dei testi restituiti in una lettura |
| `caselle` | **sì** | alias breve → nome esatto dell'archivio in Outlook; gli archivi non elencati restano invisibili |
| `casella_predefinita` | **sì** | alias usato quando Claude non ne indica uno |
| `cartelle_salvataggio` | **sì** | cartelle (e sottocartelle) in cui il connettore può scrivere file |
| `cancellazione_abilitata` | no (false) | attiva `elimina_mail` ed `elimina_evento` |
| `avviso.mittente`, `avviso.destinatario`, `avviso.prefisso_oggetto` | no | mittente (un account di Outlook) e unico destinatario di `invia_avviso`; senza la sezione, l'invio è sempre rifiutato |
| `pausa_cortesia_ms`, `pausa_ogni_elementi` | no (10, 3) | pause durante le letture lunghe per lasciare respirare la finestra di Outlook; rilette al volo senza riavvio |

`config.json` è personale ed è escluso dal repository tramite `.gitignore`: non pubblicarlo.

## Come lavora con Claude

- Claude vede ogni strumento con la sua descrizione (la docstring in `server.py`) e gli alias delle caselle nelle istruzioni del server. Gli elenchi restituiscono solo metadati (id, oggetto, mittente, data); il contenuto si legge a parte, così le conversazioni restano leggere.
- Gli id sono gli `EntryID` di Outlook: Claude li riceve dagli elenchi e li passa alle funzioni di lettura e modifica. Un elemento spostato cambia id.
- Le date si scrivono `AAAA-MM-GG` o `AAAA-MM-GG HH:MM`, ora locale.
- Per le serie ricorrenti, `elenca_eventi` restituisce una riga per occorrenza con l'id della serie; `modifica_evento` ed `elimina_evento` agiscono sull'intera serie, oppure su una sola occorrenza se ricevono `data_occorrenza` (l'inizio esatto riportato dall'elenco).
- Le etichette colorate degli appuntamenti (`telefonata`, `lavoro`, `importante`...) sono la vecchia proprietà "Label" di Outlook: alcuni server (per esempio Kerio Connect) le mostrano come colori, altri le ignorano senza danno.

Suggerimento per chi usa Cowork con attività programmate: tieni Outlook aperto, perché il connettore non lo avvia; se un'attività deve avvisarti di qualcosa, `invia_avviso` è fatta per quello.

## Lezioni tecniche (Outlook COM da Python)

Raccolte durante lo sviluppo, utili a chi vuole modificare il codice.

- **Un solo thread COM.** Gli oggetti di Outlook vanno usati dal thread che li ha creati (STA). Il server è asincrono, quindi tutte le chiamate passano da un thread dedicato con una coda e un timeout (`ComWorker`). Un blocco di Outlook (finestra modale) produce un errore spiegato, non un server appeso.
- **Filtri per data in DASL con date ISO in UTC** (`@SQL="urn:schemas:httpmail:datereceived" >= '2026-11-09 15:00'`). Il formato Jet classico (`[ReceivedTime] >= 'MM/dd/yyyy'`) viene interpretato con le impostazioni regionali di Windows e su un sistema italiano scambia giorno e mese.
- **Date in scrittura marcate come UTC.** pywin32 converte una `datetime` senza fuso da ora locale a UTC spostandola; passandola con `tzinfo=UTC` arriva a Outlook invariata e viene letta come ora locale (`com_data`).
- **Ricorrenze:** `Items.Sort("[Start]")`, `IncludeRecurrences = True`, poi `Restrict` con il filtro DASL, e si itera con l'enumeratore (mai `Count` o indice). Una serie si costruisce con `GetRecurrencePattern()` dopo aver impostato `Start`, `End` e `Duration`, nell'ordine RecurrenceType, DayOfWeekMask o DayOfMonth/MonthOfYear, Interval, PatternStartDate, StartTime, EndTime, Duration, e per ultima la fine. La singola occorrenza si raggiunge con `GetOccurrence(data_inizio)`.
- **Accenti nelle bozze:** impostare `InternetCodepage = 65001` prima di scrivere `HTMLBody`, e scrivere sempre HTML (anche per il "testo semplice"): con `Body` la bozza resta nella pagina di codice del profilo e gli accenti possono corrompersi. Nelle risposte, leggere l'HTML dell'originale **prima** di cambiare la pagina di codice, togliere i `<meta charset>` e inserire il testo nuovo subito dopo `<body>`.
- **Cancellare un evento con `Delete()`, non con `Move()` nel cestino.** Posta eliminata è una cartella di posta: alcuni provider convertono l'appuntamento spostato in un messaggio con allegato .msg (e accenti rotti). `Delete()` è ciò che fa il tasto Canc e lascia un evento ripristinabile.
- **Dopo `Send()` l'oggetto COM non è più leggibile:** i dati da restituire si prendono prima.
- **Pause di cortesia.** Ogni chiamata gira sul filo principale di Outlook: durante una lettura di centinaia di elementi l'interfaccia resta ferma per l'utente. Una pausa di pochi millisecondi ogni pochi elementi riduce molto il fastidio; non lo elimina, perché due istanze di Outlook sullo stesso profilo non esistono.
- **Errori parlanti.** Ogni eccezione viene convertita in `ToolError` con il messaggio di Outlook (`excepinfo[2]`): Claude riceve "Casella 'x' non trovata" invece di "Error executing tool", e sa spiegarlo all'utente.

## Limiti noti

- Serve il PC acceso, l'utente collegato a Windows e Outlook Classic aperto, non elevato.
- Le mail cifrate (PGP, S/MIME) arrivano cifrate: le decifra solo Outlook nella sua finestra.
- La ricerca nel corpo dei messaggi (`cerca_mail` con `in_corpo=true`) è lenta sulle cartelle con decine di migliaia di mail: per questo richiede un intervallo di date.
- La cancellazione di una singola occorrenza di una serie non passa dal cestino (comportamento di Outlook).
- Outlook e Claude desktop devono girare con gli stessi privilegi (entrambi utente normale).

## Struttura del repository

| File | Cosa fa |
|---|---|
| `server.py` | il server MCP: tutte le funzioni e tutti i vincoli, commentato riga per riga |
| `config.example.json` | modello di configurazione con la spiegazione di ogni chiave |
| `installa.ps1` | crea `.venv`, installa le dipendenze, prova il COM, registra il server in Claude desktop |
| `prova_accenti.py` | prova automatica delle funzioni di scrittura con lettere accentate |
| `.gitignore` | esclude `config.json`, `.venv`, registri e copie di sicurezza |

## Licenza

MIT: vedi `LICENSE`. Il software è fornito così com'è; prima di dare a un'intelligenza artificiale accesso alla propria posta, si consiglia di partire dalla fase 1 e di leggere il registro `audit.log` nei primi giorni.
