"""
outlook-classic-mcp: server MCP locale che collega Claude (Claude desktop, Cowork) a Outlook Classic per Windows.

COME FUNZIONA, IN BREVE
- Il server gira sul PC dell'utente, lanciato da Claude desktop (trasporto stdio: nessuna porta di rete aperta).
- Parla con Outlook Classic attraverso il modello a oggetti COM (libreria pywin32), agganciandosi alla sessione di
  Outlook GIA' APERTA dall'utente. Non avvia Outlook, non conosce password: usa gli account che Outlook ha già.
- Ogni funzione esposta a Claude (uno "strumento" MCP) è una coroutine `async def` decorata con `@mcp.tool`.
  Il lavoro vero avviene in una piccola funzione interna `f()` che viene eseguita nel thread COM (vedi ComWorker).
- Tutto ciò che dipende dall'utente (nomi delle caselle, cartelle in cui si può salvare, mittente e destinatario
  degli avvisi) sta in `config.json`, mai nel codice. Il file `config.example.json` spiega ogni chiave.

REGOLE FISSATE NEL CODICE (non in un prompt, quindi non aggirabili da Claude)
- Nessuna funzione invia posta, salvo `invia_avviso`, che ha mittente e destinatario fissi presi dalla
  configurazione: Claude può scegliere solo oggetto e testo, mai a chi scrivere.
- Nessuna funzione cancella in modo definitivo. `elimina_mail` ed `elimina_evento` spostano l'elemento nel cestino
  della sua casella (recuperabile), solo se `oggetto_atteso` coincide con l'oggetto reale, mai dal cestino, mai
  riunioni con invitati. Si disattivano con `"cancellazione_abilitata": false`. Unica eccezione: la cancellazione di
  una singola occorrenza di una serie, che Outlook non fa passare dal cestino (è documentato nello strumento).
- Nessuna funzione cancella contatti, attività, note o cartelle; nessuna tocca regole, impostazioni o account.
- Ogni mail non letta che Claude legge (o salva, o a cui prepara una risposta, o che sposta) torna non letta.
- `salva_mail` e `salva_allegati` scrivono solo nelle cartelle consentite dalla configurazione, mai sovrascrivendo,
  e lasciano accanto a ogni file un'impronta SHA-256 (utile come catena di custodia).
- Tutte le scritture finiscono in un registro (audit.log): strumento, casella, id; mai il contenuto.

FASI (chiave "fase" di config.json): 1 = sola lettura, 2 = anche bozze, contatti, attività, note, cartelle,
3 = anche calendario, 4 = anche invio avvisi vincolato. Si parte da 1 e si sale quando ci si fida.

LINGUA: nomi degli strumenti, parametri e messaggi sono in italiano, perché il connettore è nato per un utente
italiano e Claude li legge senza problemi. Il codice è commentato in italiano.
"""

from __future__ import annotations

import asyncio
import datetime as dt
import hashlib
import html as htmlmod
import json
import logging
import os
import queue
import re
import sys
import threading
import time
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations

# ----------------------------------------------------------------------------
# Configurazione
# ----------------------------------------------------------------------------

# La configurazione sta in config.json, nella stessa cartella di questo file. Senza config.json il server si ferma
# con un messaggio chiaro: il file config.example.json va copiato in config.json e compilato (vedi README).
BASE = Path(__file__).resolve().parent
_cfg_path = BASE / "config.json"
if not _cfg_path.exists():
    sys.stderr.write(
        f"outlook-classic-mcp: manca {_cfg_path}. Copiare config.example.json in config.json e compilarlo.\n"
    )
    sys.exit(1)
CONFIG: dict[str, Any] = json.loads(_cfg_path.read_text("utf-8-sig"))  # utf-8-sig: tollera il BOM di Blocco note


def _richiesta(chiave: str) -> Any:
    """Legge una chiave obbligatoria di config.json; se manca spiega quale, invece di fallire più avanti in modo oscuro."""
    if chiave not in CONFIG:
        sys.stderr.write(f"outlook-classic-mcp: in config.json manca la chiave obbligatoria \"{chiave}\" (vedi config.example.json).\n")
        sys.exit(1)
    return CONFIG[chiave]


# Fase attiva (1-4): limita gli strumenti disponibili, vedi richiede_fase().
FASE = int(CONFIG.get("fase", 1))
# Tempo massimo di attesa per una singola chiamata a Outlook. Se Outlook ha una finestra modale aperta (richiesta di
# password, avviso) non risponde: meglio un errore spiegato che un blocco infinito.
TIMEOUT = int(CONFIG.get("timeout_secondi", 40))
# Lunghezza massima del corpo (mail, evento, nota, attività) restituito a Claude, per non saturare la conversazione.
MAX_CORPO = int(CONFIG.get("max_caratteri_corpo", 20000))
# Pause di cortesia: ogni chiamata COM gira sul filo principale di Outlook, lo stesso che disegna la finestra. Durante
# una lettura lunga (centinaia di elementi) Outlook resterebbe bloccato per l'utente; una breve pausa ogni pochi elementi
# gli lascia il tempo di rispondere ai clic. 0 = nessuna pausa. Valori sensati: 10-25 ms ogni 2-3 elementi.
PAUSA_CORTESIA_MS = float(CONFIG.get("pausa_cortesia_ms", 10))
PAUSA_OGNI = max(1, int(CONFIG.get("pausa_ogni_elementi", 3)))
_pausa_cache = {"mtime": 0.0, "controllo": 0.0}


def aggiorna_pause() -> None:
    """Rilegge pausa_cortesia_ms e pausa_ogni_elementi da config.json se il file è cambiato (controllo al massimo
    ogni 5 secondi): così si tarano senza riavviare il Claude desktop."""
    global PAUSA_CORTESIA_MS, PAUSA_OGNI
    adesso = time.time()
    if adesso - _pausa_cache["controllo"] < 5:
        return
    _pausa_cache["controllo"] = adesso
    try:
        m = _cfg_path.stat().st_mtime
        if m != _pausa_cache["mtime"]:
            _pausa_cache["mtime"] = m
            c = json.loads(_cfg_path.read_text("utf-8"))
            PAUSA_CORTESIA_MS = float(c.get("pausa_cortesia_ms", 10))
            PAUSA_OGNI = max(1, int(c.get("pausa_ogni_elementi", 3)))
    except Exception:  # noqa: BLE001
        pass

# Caselle: alias breve -> nome dell'archivio come compare in Outlook (Store.DisplayName, cioè il nome in grassetto
# nel riquadro delle cartelle). Claude usa l'alias ("lavoro", "pec"...) e il connettore risolve l'archivio.
# Se Claude passa un nome che non è un alias, viene provato come nome esatto di archivio.
CASELLE: dict[str, str] = {str(k).lower(): str(v) for k, v in _richiesta("caselle").items()}
CASELLA_PREDEFINITA = str(_richiesta("casella_predefinita")).lower()
if CASELLA_PREDEFINITA not in CASELLE:
    sys.stderr.write(f"outlook-classic-mcp: \"casella_predefinita\" ({CASELLA_PREDEFINITA}) non è fra gli alias di \"caselle\".\n")
    sys.exit(1)

# Cancellazione (solo spostamento nel cestino della casella). Interruttore in config.json, spento se assente.
CANCELLAZIONE = bool(CONFIG.get("cancellazione_abilitata", False))

# Invio avvisi: UNICA eccezione alla regola "nessun invio". Mittente e destinatario vengono dalla configurazione e
# NON sono parametri dello strumento: Claude non può scegliere a chi scrivere. La sezione è facoltativa: se manca,
# invia_avviso rifiuta con un messaggio chiaro (utile a chi vuole un connettore che non invii proprio nulla).
_avviso = CONFIG.get("avviso") or {}
AVVISO_MITTENTE = str(_avviso.get("mittente", "")).strip().lower()      # indirizzo di un account presente in Outlook
AVVISO_DESTINATARIO = str(_avviso.get("destinatario", "")).strip()      # l'unico destinatario possibile
AVVISO_PREFISSO = str(_avviso.get("prefisso_oggetto", "[Avviso Claude] "))  # anteposto a ogni oggetto, per riconoscerli

# Cartelle in cui il connettore può scrivere file (allegati, .msg): mai altrove, mai sovrascrivendo.
CARTELLE_SALVATAGGIO = [Path(p) for p in _richiesta("cartelle_salvataggio")]
if not CARTELLE_SALVATAGGIO:
    sys.stderr.write("outlook-classic-mcp: \"cartelle_salvataggio\" deve contenere almeno una cartella.\n")
    sys.exit(1)

# Etichette colorate degli appuntamenti: proprietà MAPI "Label" (0x8214) dell'appuntamento, con i valori storici di
# Outlook 2003/2007 (0 = nessuna ... 10 = telefonata). Outlook moderno non le mostra più nella sua interfaccia, ma
# alcuni server le usano ancora: per esempio Kerio Connect (tramite KOFF) le mappa sulle proprie etichette colorate.
# Su Exchange o IMAP vengono scritte e rilette senza effetti visibili: innocue.
ETICHETTE: dict[str, int] = {
    "nessuna": 0, "importante": 1, "lavoro": 2, "personale": 3, "vacanza": 4, "da_non_perdere": 5,
    "viaggio": 6, "preparazione": 7, "compleanno": 8, "anniversario": 9, "telefonata": 10,
}
ETICHETTE_INV = {v: k for k, v in ETICHETTE.items()}
PROP_LABEL = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/82140003"

# Registri: server.log (tecnico, con gli errori) e audit.log (una riga JSON per ogni scrittura: strumento, casella,
# id, oggetto; mai il contenuto). Stanno in %LOCALAPPDATA%\outlook-classic-mcp, fuori dal repository.
LOGDIR = Path(os.environ.get("LOCALAPPDATA", str(BASE))) / "outlook-classic-mcp"
LOGDIR.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.FileHandler(LOGDIR / "server.log", encoding="utf-8"), logging.StreamHandler(sys.stderr)],
)
log = logging.getLogger("outlook-classic-mcp")
AUDIT = LOGDIR / "audit.log"


def audit(strumento: str, **campi: Any) -> None:
    """Aggiunge una riga al registro delle scritture. Lo chiamano tutti gli strumenti che modificano qualcosa."""
    riga = {"quando": dt.datetime.now().isoformat(timespec="seconds"), "strumento": strumento, **campi}
    with AUDIT.open("a", encoding="utf-8") as f:
        f.write(json.dumps(riga, ensure_ascii=False) + "\n")


# ----------------------------------------------------------------------------
# Thread COM (STA): tutte le chiamate a Outlook passano di qui, una alla volta, con timeout.
#
# Perché: gli oggetti COM di Outlook vanno usati dal thread che li ha creati (modello "single-threaded apartment").
# Il server MCP è asincrono (asyncio) e potrebbe servire più richieste; un solo thread dedicato con una coda garantisce
# che Outlook riceva una chiamata alla volta, nell'ordine giusto, e che un blocco di Outlook non blocchi il server.
# ----------------------------------------------------------------------------

class ComWorker:
    def __init__(self) -> None:
        self.q: "queue.Queue[tuple[Any, dict]]" = queue.Queue()  # coda di lavori (funzione, "futuro")
        self.app = None                                           # oggetto Outlook.Application, agganciato alla prima chiamata
        self.t = threading.Thread(target=self._loop, name="outlook-com", daemon=True)
        self.t.start()

    def _loop(self) -> None:
        """Ciclo del thread COM: inizializza COM una volta sola, poi esegue i lavori uno dopo l'altro."""
        pythoncom.CoInitialize()
        while True:
            fn, fut = self.q.get()
            try:
                fut["r"] = fn()
            except BaseException as e:  # noqa: BLE001
                fut["e"] = e
            finally:
                fut["ev"].set()  # sveglia chi aspetta in run()

    def run(self, fn):
        """Mette fn in coda e aspetta il risultato al massimo TIMEOUT secondi. Chiamata da un thread qualsiasi."""
        fut: dict[str, Any] = {"ev": threading.Event()}
        self.q.put((fn, fut))
        if not fut["ev"].wait(TIMEOUT):
            # Il lavoro resta in corso dentro Outlook: alla prossima chiamata la coda riparte da sola quando Outlook si libera.
            raise TimeoutError(
                f"Outlook non risponde entro {TIMEOUT} s. Probabile finestra modale aperta in Outlook "
                "(richiesta password, passphrase PGP, avviso): chiuderla e riprovare."
            )
        if "e" in fut:
            raise fut["e"]
        return fut["r"]

    def outlook(self):
        """Si aggancia alla sessione di Outlook già aperta (GetActiveObject). Non avvia Outlook e non chiede credenziali.
        GetActiveObject fallisce se Outlook è chiuso, se è ancora in avvio con una finestra davanti, oppure se gira con
        privilegi diversi da quelli di questo processo (per esempio Outlook "come amministratore")."""
        if self.app is not None:
            try:
                _ = self.app.Version  # Outlook è ancora vivo? Se è stato chiuso, l'oggetto è morto e si riaggancia.
                return self.app
            except Exception:  # noqa: BLE001
                self.app = None
        try:
            self.app = win32com.client.GetActiveObject("Outlook.Application")
        except pythoncom.com_error as e:
            raise RuntimeError(
                "Outlook Classic non risulta aperto (o è aperto come amministratore). "
                f"Aprire Outlook e riprovare. Dettaglio: {e}"
            ) from e
        return self.app


COM = ComWorker()


async def com(fn):
    """Ponte fra il mondo asincrono del server MCP e il thread COM: esegue fn là e aspetta senza bloccare asyncio.
    Ogni errore arriva a Claude come ToolError con un messaggio comprensibile, mai come 'Error executing tool' anonimo:
    così Claude sa spiegare all'utente cosa è andato storto (casella sbagliata, Outlook chiuso, data non valida...)."""
    try:
        return await asyncio.to_thread(COM.run, fn)
    except ToolError:
        raise
    except pywintypes.com_error as e:
        # excepinfo[2] contiene il testo dell'errore di Outlook, quando c'è
        dettaglio = e.excepinfo[2] if getattr(e, "excepinfo", None) and len(e.excepinfo) > 2 and e.excepinfo[2] else str(e)
        log.exception("Errore COM")
        raise ToolError(f"Outlook ha risposto con un errore: {dettaglio}") from e
    except Exception as e:  # noqa: BLE001
        log.exception("Errore")
        raise ToolError(f"{type(e).__name__}: {e}") from e


# ----------------------------------------------------------------------------
# Utilità: navigazione fra archivi e cartelle, date, formattazione dei risultati
# ----------------------------------------------------------------------------

def richiede_fase(n: int) -> None:
    """Blocca uno strumento se la fase configurata è più bassa di quella richiesta."""
    if FASE < n:
        raise ToolError(f"Funzione non abilitata: richiede la fase {n}, il connettore è in fase {FASE} (config.json).")


def ns():
    """Il NameSpace MAPI: la radice da cui si raggiungono archivi (Stores), account e singoli elementi (GetItemFromID)."""
    return COM.outlook().GetNamespace("MAPI")


def store_da_alias(alias: str | None):
    """Dall'alias di config.json (o dal nome esatto dell'archivio) all'oggetto Store di Outlook. Vuoto = casella predefinita."""
    alias = (alias or CASELLA_PREDEFINITA).strip()
    nome = CASELLE.get(alias.lower(), alias)
    for s in ns().Stores:
        if s.DisplayName == nome:
            return s
    raise ValueError(f"Casella '{alias}' non trovata. Caselle note: {', '.join(CASELLE)} (oppure il nome esatto dell'archivio).")


def cartella_da_percorso(store, percorso: str | None):
    """Percorso relativo alla radice dell'archivio, con '/' (o '\\') fra i livelli, senza distinzione fra maiuscole e
    minuscole. Vuoto = radice dell'archivio. Esempio: 'Posta in arrivo/Clienti/Rossi'."""
    f = store.GetRootFolder()
    for parte in [p for p in re.split(r"[\\/]", percorso or "") if p]:
        trovata = None
        for sub in f.Folders:
            if sub.Name.lower() == parte.lower():
                trovata = sub
                break
        if trovata is None:
            raise ValueError(f"Cartella '{parte}' non trovata in '{f.FolderPath}'.")
        f = trovata
    return f


def posta_in_arrivo(store):
    """La Posta in arrivo dell'archivio: cartella predefinita 6 (olFolderInbox); in mancanza, cerca per nome."""
    try:
        return store.GetDefaultFolder(6)
    except Exception:  # noqa: BLE001
        try:
            return cartella_da_percorso(store, "Posta in arrivo")
        except ValueError:
            return cartella_da_percorso(store, "Inbox")


def calendario(store, nome: str | None):
    """Il Calendario dell'archivio (cartella predefinita 9, olFolderCalendar) oppure, con 'nome', un altro calendario
    raggiunto per percorso. Rifiuta qualunque cartella che non sia di calendario: un evento non deve mai finire in Bozze
    o in una cartella di posta (alcuni server lo trasformerebbero in un messaggio)."""
    if nome:
        cal = cartella_da_percorso(store, nome)
    else:
        try:
            cal = store.GetDefaultFolder(9)
        except Exception:  # noqa: BLE001
            try:
                cal = cartella_da_percorso(store, "Calendario")
            except ValueError:
                cal = cartella_da_percorso(store, "Calendar")
    try:
        if cal.DefaultItemType != 1:
            raise ValueError(f"'{cal.FolderPath}' non è una cartella di calendario: un evento si crea solo nel Calendario, mai in Bozze o in una cartella di posta.")
    except ValueError:
        raise
    except Exception:  # noqa: BLE001
        pass
    return cal


def verifica_evento_puro(e, cal_path: str | None = None) -> None:
    """Regola assoluta: un evento o memo è sempre un appuntamento semplice nel Calendario, mai una riunione con
    invitati (che Outlook tratterebbe come invito da spedire o disdire) e mai un elemento in Bozze o in posta.
    I controlli sono tolleranti: se una proprietà non è leggibile si passa oltre, se è leggibile e sbagliata si rifiuta."""
    try:
        if int(e.MeetingStatus) != 0:
            raise ValueError("L'elemento è una riunione con invitati: il connettore non crea né modifica inviti. Va gestito a mano dall'utente in Outlook.")
    except ValueError:
        raise
    except Exception:  # noqa: BLE001
        pass
    try:
        if (e.RequiredAttendees or "").strip() or (e.OptionalAttendees or "").strip():
            raise ValueError("L'elemento ha dei partecipanti: il connettore non crea né modifica inviti.")
    except ValueError:
        raise
    except Exception:  # noqa: BLE001
        pass
    try:
        padre = e.Parent
        if padre.DefaultItemType != 1:
            raise ValueError(f"L'evento sta in '{padre.FolderPath}', che non è un calendario: operazione rifiutata.")
    except ValueError:
        raise
    except Exception:  # noqa: BLE001
        pass


def bozze(store):
    """La cartella Bozze dell'archivio (cartella predefinita 16, olFolderDrafts); in mancanza, cerca per nome."""
    try:
        return store.GetDefaultFolder(16)
    except Exception:  # noqa: BLE001
        try:
            return cartella_da_percorso(store, "Bozze")
        except ValueError:
            return cartella_da_percorso(store, "Drafts")


def account_per_store(store):
    """L'account di Outlook che consegna in questo archivio: serve per impostare SendUsingAccount sulle bozze, così la
    bozza nasce già con il mittente giusto. Può non esistere (archivi PST, cartelle pubbliche): allora None."""
    for a in ns().Accounts:
        try:
            if a.DeliveryStore is not None and a.DeliveryStore.StoreID == store.StoreID:
                return a
        except Exception:  # noqa: BLE001
            continue
    return None


def elemento(id_: str, casella: str | None):
    """Dall'id (EntryID) restituito dagli elenchi all'oggetto Outlook (mail, evento, contatto...). Restituisce anche
    lo Store, che serve a chi deve trovare Bozze o cestino della stessa casella."""
    store = store_da_alias(casella)
    try:
        return ns().GetItemFromID(id_, store.StoreID), store
    except pythoncom.com_error as e:
        raise ValueError(f"Elemento non trovato nella casella '{casella or CASELLA_PREDEFINITA}' (id non valido o elemento spostato). {e}") from e


def fmt(d) -> str | None:
    """Data COM di Outlook -> testo 'AAAA-MM-GGTHH:MM' (ora locale), il formato usato in tutte le risposte."""
    if d is None:
        return None
    try:
        return dt.datetime(d.year, d.month, d.day, d.hour, d.minute, d.second).isoformat(timespec="minutes")
    except Exception:  # noqa: BLE001
        return str(d)


def parse_dt(s: str | None, fine_giornata: bool = False) -> dt.datetime | None:
    """Testo -> datetime. Accetta 'AAAA-MM-GG', 'AAAA-MM-GG HH:MM', 'AAAA-MM-GGTHH:MM'. Con fine_giornata una data
    senza ora diventa le 23:59:59 (utile come limite superiore di un intervallo)."""
    if not s:
        return None
    s = s.strip().replace("T", " ")
    for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            d = dt.datetime.strptime(s, f)
            if f == "%Y-%m-%d" and fine_giornata:
                d = d.replace(hour=23, minute=59, second=59)
            return d
        except ValueError:
            continue
    raise ValueError(f"Data non riconosciuta: '{s}'. Usare AAAA-MM-GG oppure AAAA-MM-GG HH:MM.")


def dasl_data(d: dt.datetime) -> str:
    """Data per i filtri DASL (Items.Restrict con '@SQL=...'): Outlook la confronta in UTC e nel formato ISO,
    indipendentemente dalla lingua di Windows. Lezione imparata: il formato Jet classico ('[ReceivedTime] >= 'MM/dd/yyyy'')
    viene interpretato con le impostazioni regionali del PC e su un Windows italiano scambia giorno e mese."""
    locale_tz = dt.datetime.now().astimezone().tzinfo
    return d.replace(tzinfo=locale_tz).astimezone(dt.timezone.utc).strftime("%Y-%m-%d %H:%M")


def com_data(d: dt.datetime):
    """Data da scrivere in una proprietà di Outlook (Start, End, DueDate...). Outlook vuole l'ora locale "di parete";
    pywin32 converte una datetime senza fuso da ora locale a UTC, spostandola (di un'ora in Italia). Marcandola come UTC
    passa invariata e Outlook la legge come ora locale: è il comportamento che ci si aspetta scrivendo '2026-11-09 16:00'."""
    return pywintypes.Time(d.replace(tzinfo=dt.timezone.utc))


def dasl_lit(s: str) -> str:
    """Testo dentro un filtro DASL: l'apice si raddoppia, come in SQL."""
    return s.replace("'", "''")


def smtp_mittente(item) -> str:
    """Indirizzo SMTP del mittente. Con Exchange SenderEmailAddress è un indirizzo X.500 illeggibile: in quel caso si
    passa da GetExchangeUser per avere l'indirizzo vero."""
    try:
        addr = item.SenderEmailAddress or ""
        if item.SenderEmailType == "EX":
            try:
                eu = item.Sender.GetExchangeUser()
                if eu is not None and eu.PrimarySmtpAddress:
                    return eu.PrimarySmtpAddress
            except Exception:  # noqa: BLE001
                pass
        return addr
    except Exception:  # noqa: BLE001
        return ""


def meta_mail(item, alias: str) -> dict[str, Any]:
    """I metadati di una mail come li vede Claude negli elenchi: id, oggetto, mittente, data, letta o no, numero di
    allegati, cartella. Mai il corpo: quello si chiede a parte con leggi_mail, così gli elenchi restano leggeri."""
    d: dict[str, Any] = {"id": item.EntryID, "casella": alias}
    for chiave, attr in (("oggetto", "Subject"), ("non_letta", "UnRead"), ("a", "To"), ("cc", "CC")):
        try:
            v = getattr(item, attr)
            d[chiave] = v.strip() if isinstance(v, str) else v
        except Exception:  # noqa: BLE001
            d[chiave] = None
    try:
        d["data"] = fmt(item.ReceivedTime)
    except Exception:  # noqa: BLE001
        d["data"] = None
    try:
        d["da"] = f"{item.SenderName} <{smtp_mittente(item)}>"
    except Exception:  # noqa: BLE001
        d["da"] = None
    try:
        d["allegati"] = item.Attachments.Count
    except Exception:  # noqa: BLE001
        d["allegati"] = 0
    try:
        d["dimensione"] = item.Size
    except Exception:  # noqa: BLE001
        pass
    try:
        d["cartella"] = item.Parent.FolderPath
    except Exception:  # noqa: BLE001
        pass
    return d


def respira(n: int) -> None:
    """Pausa di cortesia: chiamata con il contatore degli elementi lavorati; dorme ogni PAUSA_OGNI elementi."""
    if PAUSA_CORTESIA_MS > 0 and n % PAUSA_OGNI == 0:
        time.sleep(PAUSA_CORTESIA_MS / 1000.0)


def itera(items, massimo: int):
    """Itera una collezione Items con GetFirst/GetNext, al massimo 'massimo' elementi, con pause di cortesia.
    GetFirst/GetNext è più veloce dell'accesso per indice (items[i]) con i provider che non tengono tutto in cache."""
    aggiorna_pause()
    n = 0
    it = items.GetFirst()
    while it is not None and n < massimo:
        yield it
        n += 1
        respira(n)
        it = items.GetNext()


def percorso_consentito(p: Path) -> bool:
    """True se p sta dentro una delle cartelle_salvataggio di config.json. resolve() normalizza '..' e collegamenti,
    così non si esce dalla cartella consentita con trucchi di percorso."""
    p = p.resolve()
    for root in CARTELLE_SALVATAGGIO:
        try:
            p.relative_to(root.resolve())
            return True
        except ValueError:
            continue
    return False


def nome_libero(cartella: Path, nome: str) -> Path:
    """Un nome di file non ancora usato nella cartella: se 'nome' esiste già si aggiunge ' (1)', ' (2)'... Il connettore
    non sovrascrive mai niente. Un nome è occupato anche se esiste soltanto la sua impronta '<nome>.sha256'."""
    base = Path(nome).name or "allegato"
    dest = cartella / base
    i = 1
    while dest.exists() or dest.with_name(dest.name + ".sha256").exists():
        dest = cartella / f"{Path(base).stem} ({i}){Path(base).suffix}"
        i += 1
    return dest


# Proprietà MAPI PR_TRANSPORT_MESSAGE_HEADERS: le intestazioni Internet complete (Received, Authentication-Results,
# DKIM-Signature, Return-Path...), quelle che servono per controllare SPF, DKIM e DMARC di una mail sospetta.
PROP_INTESTAZIONI = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
OL_MSG_UNICODE = 9  # OlSaveAsType.olMSGUnicode: formato .msg con testo Unicode (gli accenti si conservano)
_RE_NOME_VIETATO = re.compile(r'[\\/:*?"<>|\x00-\x1f]+')  # caratteri che Windows non accetta in un nome di file


def nome_da_oggetto(oggetto: str, massimo: int = 80) -> str:
    """Oggetto ripulito per un nome di file Windows: niente caratteri vietati, spazi compattati, lunghezza limitata."""
    s = _RE_NOME_VIETATO.sub(" ", oggetto or "")
    s = re.sub(r"\s+", " ", s).strip()
    s = s[:massimo].rstrip(" .")
    return s or "senza oggetto"


def data_per_nome(item) -> str:
    """AAAA-MM-GG_HHMM dalla data di ricezione; per le bozze (mai ricevute) dalla data di creazione."""
    for attr in ("ReceivedTime", "CreationTime"):
        try:
            d = getattr(item, attr)
            if d is not None and 1990 < d.year < 3000:
                return f"{d.year:04d}-{d.month:02d}-{d.day:02d}_{d.hour:02d}{d.minute:02d}"
        except Exception:  # noqa: BLE001
            continue
    return dt.datetime.now().strftime("%Y-%m-%d_%H%M")


def impronta_sha256(file: Path, **contesto: Any) -> dict[str, str]:
    """Calcola lo SHA-256 del file appena salvato e scrive accanto '<file>.sha256' (catena di custodia: dimostra che
    il file non è cambiato da quando è stato estratto da Outlook).
    Formato: righe di commento '#' con data e ora di acquisizione e contesto, poi la riga standard
    '<impronta>  <nome file>', verificabile con 'sha256sum -c' (che salta i commenti) o confrontando Get-FileHash.
    Il file .sha256 non viene mai sovrascritto (apertura in modo 'x')."""
    h = hashlib.sha256()
    with file.open("rb") as fh:
        for blocco in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(blocco)
    hexd = h.hexdigest()
    acquisito = dt.datetime.now().astimezone().isoformat(timespec="seconds")
    righe = ["# Impronta SHA-256 per la catena di custodia (outlook-classic-mcp)",
             f"# file: {file.name}",
             f"# dimensione_byte: {file.stat().st_size}",
             f"# acquisito: {acquisito}",
             "# postazione: " + os.environ.get("COMPUTERNAME", "")]
    for k, v in contesto.items():
        righe.append(f"# {k}: {str(v).replace(chr(10), ' ').replace(chr(13), ' ')}")
    righe.append(f"{hexd}  {file.name}")
    f_imp = file.with_name(file.name + ".sha256")
    with f_imp.open("x", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(righe) + "\n")
    return {"sha256": hexd, "file_sha256": str(f_imp), "acquisito": acquisito}


def testo_a_html(testo: str) -> str:
    """Testo semplice -> HTML minimo (caratteri speciali protetti, a capo come <br>). Tutte le bozze passano da HTML,
    vedi crea_bozza per il motivo."""
    return "<div style=\"font-family:Calibri,sans-serif;font-size:11pt\">" + htmlmod.escape(testo).replace("\n", "<br>") + "</div>"


UTF8_CPID = 65001  # pagina di codice UTF-8 per MailItem.InternetCodepage
_RE_META_CHARSET = re.compile(r"<meta[^>]*charset[^>]*>", re.I)
_RE_BODY_OPEN = re.compile(r"<body[^>]*>", re.I)


def forza_utf8(item) -> None:
    """Imposta la pagina di codice UTF-8 sull'elemento PRIMA di scrivere HTMLBody.
    Senza questo Outlook può ricodificare il testo con la pagina di codice dell'originale (per esempio Windows-1252)
    e le lettere accentate arrivano come 'Ã ' o simili."""
    try:
        item.InternetCodepage = UTF8_CPID
    except Exception:  # noqa: BLE001
        pass


def html_con_testo_in_testa(html_esistente: str, testo_html: str) -> str:
    """Inserisce testo_html subito dopo <body> dell'HTML esistente (la citazione dell'originale in una risposta),
    togliendo i <meta charset> che Outlook userebbe per ricodificare il contenuto. Se non c'è <body>, lo mette in testa."""
    base = _RE_META_CHARSET.sub("", html_esistente or "")
    m = _RE_BODY_OPEN.search(base)
    if m:
        return base[:m.end()] + testo_html + "<br>" + base[m.end():]
    return testo_html + "<br>" + base


def meta_evento(e, alias: str, cal_path: str | None = None) -> dict[str, Any]:
    """I metadati di un appuntamento per gli elenchi: id, oggetto, luogo, inizio, fine, durata, ricorrente o no,
    etichetta colorata, categorie, promemoria. La descrizione si legge a parte con leggi_evento."""
    d: dict[str, Any] = {"casella": alias}
    try:
        d["id"] = e.EntryID
    except Exception as ex:  # noqa: BLE001
        # Occorrenza di una serie modificata che Outlook non sa più risolvere: si segnala e si va avanti.
        d["id"] = None
        d["avviso"] = "occorrenza non leggibile: " + (ex.excepinfo[2] if getattr(ex, "excepinfo", None) and len(ex.excepinfo) > 2 and ex.excepinfo[2] else str(ex))
    for chiave, attr in (("oggetto", "Subject"), ("luogo", "Location"), ("tutto_il_giorno", "AllDayEvent"),
                         ("durata_minuti", "Duration"), ("categorie", "Categories"), ("ricorrente", "IsRecurring"),
                         ("promemoria", "ReminderSet")):
        try:
            d[chiave] = getattr(e, attr)
        except Exception:  # noqa: BLE001
            d[chiave] = None
    try:
        d["inizio"] = fmt(e.Start)
        d["fine"] = fmt(e.End)
    except Exception:  # noqa: BLE001
        pass
    try:
        lbl = int(e.PropertyAccessor.GetProperty(PROP_LABEL))
        d["etichetta"] = ETICHETTE_INV.get(lbl, str(lbl))
    except Exception:  # noqa: BLE001
        d["etichetta"] = None
    if cal_path:
        d["calendario"] = cal_path
    return d



# ----------------------------------------------------------------------------
# Utilità per rubrica, attività, note, cartelle
# ----------------------------------------------------------------------------

# Codici OlObjectClass degli elementi: servono a verificare che un id indichi davvero il tipo atteso.
CLASSI = {43: "mail", 26: "evento", 40: "contatto", 48: "attivita", 44: "nota", 46: "rapporto", 53: "riunione"}
# Codici OlItemType delle cartelle (Folder.DefaultItemType).
TIPI_CARTELLA = {0: "posta", 1: "calendario", 2: "contatti", 3: "attivita", 4: "diario", 5: "note"}
# Nomi con cui i vari provider chiamano il cestino (italiano e inglese): per riconoscerlo anche su IMAP e PEC.
NOMI_CESTINO = {"posta eliminata", "deleted items", "cestino", "trash", "elementi eliminati"}


def cartella_predefinita(store, codice: int, nome_it: str, nome_en: str):
    """Cartella predefinita di un archivio (GetDefaultFolder); se il provider non la dichiara, la cerca per nome
    italiano e poi inglese. Codici utili: 13 = Attività, 12 = Note, 10 = Contatti, 3 = Posta eliminata."""
    try:
        return store.GetDefaultFolder(codice)
    except Exception:  # noqa: BLE001
        try:
            return cartella_da_percorso(store, nome_it)
        except ValueError:
            return cartella_da_percorso(store, nome_en)


def cartelle_di_tipo(store, tipo: int, profondita: int = 3):
    """Tutte le cartelle di un archivio di un dato DefaultItemType (2 = contatti, 3 = attivita, 5 = note)."""
    out = []

    def walk(fold, liv):
        for sub in fold.Folders:
            try:
                if sub.DefaultItemType == tipo:
                    out.append(sub)
            except Exception:  # noqa: BLE001
                pass
            if liv < profondita:
                walk(sub, liv + 1)

    walk(store.GetRootFolder(), 1)
    return out


def verifica_classe(item, attesa: int, cosa: str) -> None:
    """Rifiuta un id che indica un elemento di tipo diverso da quello atteso (per esempio un contatto passato a elimina_mail)."""
    try:
        c = item.Class
    except Exception:  # noqa: BLE001
        return
    if c != attesa:
        raise ValueError(f"L'id indicato non è {cosa} (classe {CLASSI.get(c, c)}).")


def e_cestino(store, fold) -> bool:
    """True se la cartella è il cestino dell'archivio o sta dentro un cestino (per nome, in italiano o inglese)."""
    try:
        if fold.EntryID == store.GetDefaultFolder(3).EntryID:
            return True
    except Exception:  # noqa: BLE001
        pass
    try:
        parti = [p.lower() for p in fold.FolderPath.split("\\") if p]
        return any(p in NOMI_CESTINO for p in parti)
    except Exception:  # noqa: BLE001
        return False


def meta_contatto(c, alias: str, completo: bool = False) -> dict[str, Any]:
    """Scheda di un contatto: i campi principali negli elenchi, tutti i campi (indirizzi, note...) con completo=True.
    I campi vuoti non vengono restituiti, per tenere corte le risposte."""
    d: dict[str, Any] = {"casella": alias}
    try:
        d["id"] = c.EntryID
    except Exception:  # noqa: BLE001
        d["id"] = None
    campi = [("nome_completo", "FullName"), ("azienda", "CompanyName"), ("ruolo", "JobTitle"),
             ("email", "Email1Address"), ("email2", "Email2Address"), ("tel_lavoro", "BusinessTelephoneNumber"),
             ("cellulare", "MobileTelephoneNumber"), ("tel_casa", "HomeTelephoneNumber"), ("categorie", "Categories")]
    if completo:
        campi += [("nome", "FirstName"), ("cognome", "LastName"), ("email3", "Email3Address"),
                  ("indirizzo_lavoro", "BusinessAddress"), ("indirizzo_casa", "HomeAddress"), ("sito_web", "WebPage"),
                  ("reparto", "Department"), ("note", "Body"), ("archiviato_come", "FileAs")]
    for chiave, attr in campi:
        try:
            v = getattr(c, attr)
            if v not in (None, ""):
                d[chiave] = v.strip() if isinstance(v, str) else v
        except Exception:  # noqa: BLE001
            pass
    try:
        d["cartella"] = c.Parent.FolderPath
    except Exception:  # noqa: BLE001
        pass
    return d


def meta_attivita(t, alias: str, completo: bool = False) -> dict[str, Any]:
    """Dati di un'attività. Le date 'vuote' in Outlook valgono 1.1.4501: vengono restituite come null."""
    d: dict[str, Any] = {"casella": alias}
    try:
        d["id"] = t.EntryID
    except Exception:  # noqa: BLE001
        d["id"] = None
    for chiave, attr in (("oggetto", "Subject"), ("completata", "Complete"), ("percentuale", "PercentComplete"),
                         ("priorita", "Importance"), ("categorie", "Categories"), ("stato", "Status"),
                         ("promemoria", "ReminderSet")):
        try:
            d[chiave] = getattr(t, attr)
        except Exception:  # noqa: BLE001
            d[chiave] = None
    for chiave, attr in (("scadenza", "DueDate"), ("inizio", "StartDate"), ("completata_il", "DateCompleted"),
                         ("promemoria_il", "ReminderTime")):
        try:
            v = getattr(t, attr)
            d[chiave] = None if (v is None or v.year >= 4500) else fmt(v)
        except Exception:  # noqa: BLE001
            d[chiave] = None
    if completo:
        try:
            d["descrizione"] = (t.Body or "")[:MAX_CORPO]
        except Exception:  # noqa: BLE001
            pass
    try:
        d["cartella"] = t.Parent.FolderPath
    except Exception:  # noqa: BLE001
        pass
    return d


def meta_nota(n, alias: str, completo: bool = False) -> dict[str, Any]:
    """Dati di una nota (i 'post-it' di Outlook): titolo (la prima riga), categorie, data di modifica, testo se completo."""
    d: dict[str, Any] = {"casella": alias}
    try:
        d["id"] = n.EntryID
    except Exception:  # noqa: BLE001
        d["id"] = None
    for chiave, attr in (("titolo", "Subject"), ("categorie", "Categories")):
        try:
            d[chiave] = getattr(n, attr)
        except Exception:  # noqa: BLE001
            d[chiave] = None
    try:
        d["modificata"] = fmt(n.LastModificationTime)
    except Exception:  # noqa: BLE001
        pass
    if completo:
        try:
            d["testo"] = (n.Body or "")[:MAX_CORPO]
        except Exception:  # noqa: BLE001
            pass
    return d


def _applica_contatto(c, nome, cognome, azienda, ruolo, email, email2, tel_lavoro, cellulare, tel_casa,
                      indirizzo_lavoro, sito_web, note, categorie) -> None:
    """Scrive sul contatto solo i campi passati (None = non toccare). Condivisa fra crea_contatto e modifica_contatto."""
    for valore, attr in ((nome, "FirstName"), (cognome, "LastName"), (azienda, "CompanyName"), (ruolo, "JobTitle"),
                         (email, "Email1Address"), (email2, "Email2Address"), (tel_lavoro, "BusinessTelephoneNumber"),
                         (cellulare, "MobileTelephoneNumber"), (tel_casa, "HomeTelephoneNumber"),
                         (indirizzo_lavoro, "BusinessAddress"), (sito_web, "WebPage"), (note, "Body"),
                         (categorie, "Categories")):
        if valore is not None:
            setattr(c, attr, valore)


def _applica_attivita(t, oggetto, scadenza, inizio, descrizione, priorita, promemoria, categorie, completata,
                      percentuale) -> None:
    """Scrive sull'attività solo i campi passati (None = non toccare). promemoria='' toglie il promemoria.
    Condivisa fra crea_attivita e modifica_attivita."""
    if oggetto is not None:
        t.Subject = oggetto
    if scadenza is not None:
        t.DueDate = com_data(parse_dt(scadenza))
    if inizio is not None:
        t.StartDate = com_data(parse_dt(inizio))
    if descrizione is not None:
        t.Body = descrizione
    if priorita is not None:
        p = {"bassa": 0, "normale": 1, "alta": 2}.get(str(priorita).lower())
        if p is None:
            raise ValueError("Priorità non valida: bassa, normale, alta.")
        t.Importance = p
    if promemoria is not None:
        if promemoria == "":
            t.ReminderSet = False
        else:
            t.ReminderSet = True
            t.ReminderTime = com_data(parse_dt(promemoria))
    if categorie is not None:
        t.Categories = categorie
    if percentuale is not None:
        t.PercentComplete = int(percentuale)
    if completata is not None:
        if completata:
            t.MarkComplete()
        else:
            t.Complete = False
            t.Status = 0

# ----------------------------------------------------------------------------
# Server MCP
#
# `instructions` è il testo che Claude riceve insieme all'elenco degli strumenti: gli dice cosa c'è e quali sono le
# regole d'uso. Gli strumenti sono le funzioni decorate con @mcp.tool; la docstring di ciascuna diventa la descrizione
# che Claude legge per capire quando e come usarla, quindi è scritta per lui.
# ----------------------------------------------------------------------------

mcp = MCPServer(
    "outlook",
    instructions=(
        "Connettore locale per Outlook Classic dell'utente, sul suo PC. "
        f"Casella predefinita: alias '{CASELLA_PREDEFINITA}' (archivio \"{CASELLE[CASELLA_PREDEFINITA]}\"); "
        "altre caselle per alias: " + ", ".join(CASELLE) + ". "
        "Copre posta, calendario, rubrica (contatti), attività, note e cartelle. Non invia (tranne invia_avviso, con "
        "destinatario fisso) e non tocca regole o impostazioni. Cancella solo con elimina_mail ed elimina_evento, che "
        "spostano nel cestino (recuperabile) e vanno usate SOLO dopo la conferma esplicita dell'utente in chat. "
        "Le mail non lette restano non lette dopo la lettura. "
        f"Fase attiva: {FASE} (1 lettura, 2 bozze e scritture, 3 calendario, 4 avvisi)."
    ),
)

# Annotazioni MCP: dicono al client se uno strumento è di sola lettura o distruttivo (Claude desktop le usa per
# decidere quando chiedere conferma all'utente).
RO = ToolAnnotations(read_only_hint=True, destructive_hint=False, open_world_hint=False)
RW = ToolAnnotations(read_only_hint=False, destructive_hint=False, open_world_hint=False)


# ---------------------------------------------------------------- fase 1: lettura della posta

@mcp.tool(annotations=RO)
async def stato_outlook() -> dict:
    """Verifica che Outlook Classic sia aperto e raggiungibile; riporta versione, profilo, fase attiva e caselle disponibili."""
    def f():
        app = COM.outlook()
        n = app.GetNamespace("MAPI")
        archivi = [s.DisplayName for s in n.Stores]
        return {
            "ok": True, "versione_outlook": app.Version, "profilo": n.CurrentProfileName,
            "fase": FASE, "casella_predefinita": CASELLA_PREDEFINITA,
            "caselle": {a: {"archivio": nome, "presente": nome in archivi} for a, nome in CASELLE.items()},
            "archivi_in_outlook": archivi,
        }
    return await com(f)


@mcp.tool(annotations=RO)
async def elenca_cartelle(casella: str | None = None, percorso: str = "", profondita: int = 2) -> list[dict]:
    """Elenca le cartelle di una casella (alias o nome archivio) a partire da un percorso (vuoto = radice), con conteggi.
    Tipi: posta, calendario, contatti, attivita, note."""
    tipi = TIPI_CARTELLA

    def f():
        store = store_da_alias(casella)
        radice = cartella_da_percorso(store, percorso)
        out: list[dict] = []

        def walk(fold, prefisso: str, liv: int):
            for sub in fold.Folders:
                respira(len(out) + 1)
                p = f"{prefisso}/{sub.Name}" if prefisso else sub.Name
                riga = {"percorso": p, "tipo": tipi.get(sub.DefaultItemType, str(sub.DefaultItemType))}
                try:
                    riga["elementi"] = sub.Items.Count
                    riga["non_letti"] = sub.UnReadItemCount
                except Exception:  # noqa: BLE001
                    pass
                out.append(riga)
                if liv < profondita:
                    walk(sub, p, liv + 1)

        walk(radice, percorso.strip("/\\"), 1)
        return out
    return await com(f)


@mcp.tool(annotations=RO)
async def elenca_mail(casella: str | None = None, cartella: str = "Posta in arrivo", giorni: int = 7,
                      massimo: int = 30, solo_non_lette: bool = False) -> list[dict]:
    """Elenca le mail più recenti di una cartella (predefinita: Posta in arrivo della casella predefinita), dalla più nuova.
    Restituisce solo i metadati (id, data, mittente, oggetto, non letta, numero allegati), mai il corpo."""
    def f():
        store = store_da_alias(casella)
        fold = posta_in_arrivo(store) if cartella.lower() in ("posta in arrivo", "inbox", "") else cartella_da_percorso(store, cartella)
        da = dt.datetime.now() - dt.timedelta(days=max(1, giorni))
        filtro = f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{dasl_data(da)}'"
        if solo_non_lette:
            filtro += " AND \"urn:schemas:httpmail:read\" = 0"
        items = fold.Items.Restrict(filtro)
        items.Sort("[ReceivedTime]", True)
        return [meta_mail(it, casella or CASELLA_PREDEFINITA) for it in itera(items, max(1, min(massimo, 200)))]
    return await com(f)


@mcp.tool(annotations=RO)
async def cerca_mail(testo: str, casella: str | None = None, cartella: str = "Posta in arrivo",
                     in_oggetto: bool = True, in_corpo: bool = False, mittente: str | None = None,
                     da_data: str | None = None, a_data: str | None = None, massimo: int = 30) -> list[dict]:
    """Cerca mail in una cartella per testo nell'oggetto e/o nel corpo, mittente (nome o indirizzo) e intervallo di date (AAAA-MM-GG).
    La ricerca nel corpo (in_corpo=true) è lenta e blocca Outlook: richiede da_data (intervallo di date). Restituisce solo metadati."""
    def f():
        store = store_da_alias(casella)
        fold = posta_in_arrivo(store) if cartella.lower() in ("posta in arrivo", "inbox", "") else cartella_da_percorso(store, cartella)
        cond: list[str] = []
        t = dasl_lit(testo)
        parti = []
        if in_oggetto:
            parti.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{t}%'")
        if in_corpo:
            parti.append(f"\"urn:schemas:httpmail:textdescription\" LIKE '%{t}%'")
        if parti:
            cond.append("(" + " OR ".join(parti) + ")")
        if mittente:
            m = dasl_lit(mittente)
            cond.append(f"(\"urn:schemas:httpmail:fromname\" LIKE '%{m}%' OR \"urn:schemas:httpmail:fromemail\" LIKE '%{m}%')")
        d1, d2 = parse_dt(da_data), parse_dt(a_data, fine_giornata=True)
        if d1:
            cond.append(f"\"urn:schemas:httpmail:datereceived\" >= '{dasl_data(d1)}'")
        if d2:
            cond.append(f"\"urn:schemas:httpmail:datereceived\" <= '{dasl_data(d2)}'")
        if not cond:
            raise ValueError("Nessun criterio di ricerca.")
        if in_corpo and not d1:
            raise ValueError("La ricerca nel corpo richiede da_data (e meglio anche a_data): su cartelle grandi bloccherebbe Outlook per secondi.")
        items = fold.Items.Restrict("@SQL=" + " AND ".join(cond))
        items.Sort("[ReceivedTime]", True)
        return [meta_mail(it, casella or CASELLA_PREDEFINITA) for it in itera(items, max(1, min(massimo, 200)))]
    return await com(f)


@mcp.tool(annotations=RO)
async def leggi_mail(id: str, casella: str | None = None, formato: str = "testo", max_caratteri: int | None = None,
                     intestazioni: bool = False) -> dict:
    """Legge una mail (id da elenca_mail/cerca_mail): intestazioni, corpo (formato 'testo' o 'html') ed elenco allegati.
    Con intestazioni=true aggiunge 'intestazioni_trasporto': le intestazioni Internet complete (Received,
    Authentication-Results, DKIM-Signature, Return-Path...) per controllare SPF, DKIM e DMARC.
    Se la mail era non letta, resta non letta."""
    def f():
        item, _ = elemento(id, casella)
        era_non_letta = bool(item.UnRead)
        d = meta_mail(item, casella or CASELLA_PREDEFINITA)
        lim = max_caratteri or MAX_CORPO
        corpo = item.HTMLBody if formato == "html" else item.Body
        corpo = corpo or ""
        d["corpo"] = corpo[:lim]
        d["corpo_troncato"] = len(corpo) > lim
        alle = []
        for i in range(1, item.Attachments.Count + 1):
            a = item.Attachments.Item(i)
            alle.append({"indice": i, "nome": a.FileName, "dimensione": a.Size})
        d["elenco_allegati"] = alle
        if intestazioni:
            try:
                d["intestazioni_trasporto"] = item.PropertyAccessor.GetProperty(PROP_INTESTAZIONI) or ""
            except Exception:  # noqa: BLE001
                d["intestazioni_trasporto"] = ""
            if not d["intestazioni_trasporto"]:
                d["intestazioni_nota"] = ("Intestazioni di trasporto non disponibili per questo elemento (bozze, elementi "
                                          "creati in Outlook o non conservate dall'archivio). Per la prova forense serve "
                                          "l'.eml originale dal server.")
        if era_non_letta and not item.UnRead:
            item.UnRead = True
            item.Save()
        d["non_letta"] = bool(item.UnRead)
        return d
    return await com(f)


@mcp.tool(annotations=RW)
async def salva_allegati(id: str, cartella_destinazione: str, casella: str | None = None,
                         indici: list[int] | None = None) -> list[dict]:
    """Salva su disco gli allegati di una mail (tutti, o quelli con gli indici dati da leggi_mail) in una cartella
    consentita (una di quelle in "cartelle_salvataggio" di config.json, o una sua sottocartella). Non sovrascrive mai
    file esistenti; accanto a ogni file scrive l'impronta SHA-256. Se la mail era non letta, resta non letta."""
    def f():
        dest = Path(cartella_destinazione)
        if not percorso_consentito(dest):
            raise ValueError("Cartella non consentita. Consentite: " + "; ".join(str(p) for p in CARTELLE_SALVATAGGIO))
        dest.mkdir(parents=True, exist_ok=True)
        item, _ = elemento(id, casella)
        era_non_letta = bool(item.UnRead)
        out = []
        for i in range(1, item.Attachments.Count + 1):
            if indici and i not in indici:
                continue
            a = item.Attachments.Item(i)
            target = nome_libero(dest, a.FileName)
            a.SaveAsFile(str(target))
            imp = impronta_sha256(target, origine="allegato di una mail", casella=casella or CASELLA_PREDEFINITA,
                                  oggetto_mail=item.Subject or "", allegato=a.FileName, id_mail=id)
            out.append({"indice": i, "nome": a.FileName, "salvato_in": str(target), "dimensione": target.stat().st_size,
                        **imp})
        if era_non_letta and not item.UnRead:
            item.UnRead = True
            item.Save()
        audit("salva_allegati", casella=casella or CASELLA_PREDEFINITA, id=id, file=[o["salvato_in"] for o in out],
              sha256=[o["sha256"] for o in out])
        return out
    return await com(f)


@mcp.tool(annotations=RW)
async def salva_mail(id: str, cartella_destinazione: str, casella: str | None = None) -> dict:
    """Salva la mail completa come file .msg (Unicode): intestazioni di trasporto, testo, HTML e allegati.
    Solo in una cartella consentita ("cartelle_salvataggio" di config.json, o una sottocartella), mai sovrascrivendo.
    Nome: AAAA-MM-GG_HHMM_<oggetto>.msg con la data di ricezione. Accanto scrive <file>.sha256 con impronta, data e
    ora di acquisizione. Se la mail era non letta, resta non letta.
    Nota: per la prova forense il .msg NON sostituisce l'.eml originale (la firma DKIM non si può riverificare)."""
    def f():
        dest = Path(cartella_destinazione)
        if not percorso_consentito(dest):
            raise ValueError("Cartella non consentita. Consentite: " + "; ".join(str(p) for p in CARTELLE_SALVATAGGIO))
        dest.mkdir(parents=True, exist_ok=True)
        item, _ = elemento(id, casella)
        if int(getattr(item, "Class", 0)) != 43:
            raise ValueError("L'elemento non è una mail (salva_mail salva solo messaggi di posta).")
        era_non_letta = bool(item.UnRead)
        oggetto = item.Subject or ""
        target = nome_libero(dest, f"{data_per_nome(item)}_{nome_da_oggetto(oggetto)}.msg")
        try:
            item.SaveAs(str(target), OL_MSG_UNICODE)
        finally:
            if era_non_letta and not item.UnRead:
                item.UnRead = True
                item.Save()
        if not target.is_file():
            raise ValueError(f"Outlook non ha scritto il file {target}.")
        imp = impronta_sha256(target, origine="mail completa (.msg Unicode, MailItem.SaveAs)",
                              casella=casella or CASELLA_PREDEFINITA, oggetto_mail=oggetto, id_mail=id,
                              nota="il .msg non sostituisce l'.eml originale per la verifica DKIM")
        audit("salva_mail", casella=casella or CASELLA_PREDEFINITA, id=id, file=str(target), sha256=imp["sha256"])
        return {"id": id, "oggetto": oggetto, "salvato_in": str(target), "dimensione": target.stat().st_size,
                "non_letta": bool(item.UnRead), **imp}
    return await com(f)


@mcp.tool(annotations=RW)
async def segna_mail(id: str, letta: bool, casella: str | None = None) -> dict:
    """Segna una mail come letta (letta=true) o non letta (letta=false)."""
    def f():
        item, _ = elemento(id, casella)
        item.UnRead = not letta
        item.Save()
        audit("segna_mail", casella=casella or CASELLA_PREDEFINITA, id=id, letta=letta)
        return {"id": id, "non_letta": bool(item.UnRead)}
    return await com(f)


# ---------------------------------------------------------------- fase 2: bozze

@mcp.tool(annotations=RW)
async def crea_bozza(oggetto: str, corpo: str, a: str = "", cc: str = "", casella: str | None = None,
                     formato: str = "testo", allegati: list[str] | None = None) -> dict:
    """Crea una bozza di mail nella cartella Bozze della casella (predefinita: la casella predefinita di config.json).
    Non invia mai: la bozza resta in Outlook perché l'utente la riveda, la firmi se vuole e la spedisca lui.
    Destinatari separati da ';'. formato: 'testo' oppure 'html' (corpo già in HTML). 'allegati' = percorsi di file sul PC."""
    richiede_fase(2)

    def f():
        store = store_da_alias(casella)
        fold = bozze(store)
        m = fold.Items.Add(0)  # olMailItem, creato direttamente nella cartella Bozze della casella giusta
        acc = account_per_store(store)
        if acc is not None:
            m.SendUsingAccount = acc  # così la bozza ha già il mittente di quella casella
        m.Subject = oggetto
        m.To = a
        m.CC = cc
        # Anche il formato 'testo' passa dalla via HTML con pagina di codice UTF-8: scrivendo m.Body la bozza resta
        # nella pagina di codice predefinita del profilo e le lettere accentate possono arrivare corrotte.
        forza_utf8(m)
        if formato == "html":
            m.HTMLBody = _RE_META_CHARSET.sub("", corpo)
        else:
            m.HTMLBody = testo_a_html(corpo)
        for p in allegati or []:
            pp = Path(p)
            if not pp.is_file():
                raise ValueError(f"Allegato non trovato: {p}")
            m.Attachments.Add(str(pp))
        m.Save()
        audit("crea_bozza", casella=casella or CASELLA_PREDEFINITA, id=m.EntryID, oggetto=oggetto)
        return {"id": m.EntryID, "cartella": fold.FolderPath, "oggetto": oggetto, "a": a, "cc": cc, "inviata": False}
    return await com(f)


@mcp.tool(annotations=RW)
async def crea_bozza_risposta(id: str, corpo: str, casella: str | None = None, a_tutti: bool = False) -> dict:
    """Crea una bozza di risposta a una mail (id da elenca_mail/cerca_mail), con il testo dato sopra la citazione
    dell'originale. Non invia mai. Se la mail originale era non letta, resta non letta."""
    richiede_fase(2)

    def f():
        item, store = elemento(id, casella)
        era_non_letta = bool(item.UnRead)
        r = item.ReplyAll() if a_tutti else item.Reply()
        # Ordine obbligatorio: PRIMA si legge l'HTML della risposta con la pagina di codice originale (altrimenti i
        # caratteri in Windows-1252 riletti come UTF-8 diventano "\ufffd"), POI si forza UTF-8, POI si scrive il nuovo HTML.
        html_orig = r.HTMLBody or ""
        forza_utf8(r)
        r.HTMLBody = html_con_testo_in_testa(html_orig, testo_a_html(corpo))
        r.Save()
        # Reply() salva nelle Bozze dell'archivio predefinito del profilo: se la mail era in un'altra casella, la bozza
        # viene spostata nelle Bozze di quella casella, dove l'utente se la aspetta.
        try:
            dest = bozze(store)
            if r.Parent.StoreID != store.StoreID:
                r = r.Move(dest)
        except Exception:  # noqa: BLE001
            pass
        if era_non_letta and not item.UnRead:
            item.UnRead = True
            item.Save()
        audit("crea_bozza_risposta", casella=casella or CASELLA_PREDEFINITA, id_originale=id, id=r.EntryID)
        return {"id": r.EntryID, "cartella": r.Parent.FolderPath, "oggetto": r.Subject, "a": r.To, "cc": r.CC, "inviata": False}
    return await com(f)


# ---------------------------------------------------------------- fase 3: calendario

@mcp.tool(annotations=RO)
async def elenca_eventi(da_data: str | None = None, a_data: str | None = None, casella: str | None = None,
                        calendario_nome: str | None = None, massimo: int = 100) -> list[dict]:
    """Elenca gli eventi e i memo di un calendario fra due date (AAAA-MM-GG; predefinito: da oggi a +7 giorni),
    con le ricorrenze espanse (ogni occorrenza di una serie compare come riga, con lo stesso id della serie).
    'calendario_nome' vuoto = Calendario principale; altrimenti il percorso di un altro calendario (vedi elenca_cartelle)."""
    def f():
        store = store_da_alias(casella)
        cal = calendario(store, calendario_nome)
        d1 = parse_dt(da_data) or dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        d2 = parse_dt(a_data, fine_giornata=True) or (d1 + dt.timedelta(days=7))
        # Sequenza richiesta da Outlook per espandere le serie: Sort su [Start], IncludeRecurrences = True, poi Restrict
        # con un filtro DASL sulle date (ISO, UTC). In un altro ordine le ricorrenze non vengono espanse o il filtro
        # viene ignorato. Si itera con l'enumeratore, mai con Count/indice (con le ricorrenze Count non è definito).
        items = cal.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        r = items.Restrict(f"@SQL=\"urn:schemas:calendar:dtstart\" >= '{dasl_data(d1)}' AND \"urn:schemas:calendar:dtstart\" <= '{dasl_data(d2)}'")
        out = []
        lim = max(1, min(massimo, 500))
        aggiorna_pause()
        for e in r:
            m = meta_evento(e, casella or CASELLA_PREDEFINITA, cal.FolderPath)
            out.append(m)
            respira(len(out))
            if len(out) >= lim:
                break
            # guardia: con le ricorrenze espanse l'enumerazione non deve mai superare la data finale
            try:
                if dt.datetime(e.Start.year, e.Start.month, e.Start.day, e.Start.hour, e.Start.minute) > d2:
                    break
            except Exception:  # noqa: BLE001
                pass
        return out
    return await com(f)


@mcp.tool(annotations=RO)
async def leggi_evento(id: str, casella: str | None = None) -> dict:
    """Legge un evento o memo del calendario (id da elenca_eventi), con descrizione completa."""
    def f():
        e, _ = elemento(id, casella)
        d = meta_evento(e, casella or CASELLA_PREDEFINITA)
        d["descrizione"] = (e.Body or "")[:MAX_CORPO]
        try:
            d["promemoria_minuti_prima"] = e.ReminderMinutesBeforeStart
            d["partecipanti"] = e.RequiredAttendees
            if e.IsRecurring:
                rp = e.GetRecurrencePattern()
                d["ricorrenza"] = {"tipo": NOMI_RICORRENZA.get(rp.RecurrenceType, rp.RecurrenceType), "intervallo": rp.Interval,
                                   "inizio": fmt(rp.PatternStartDate), "fine": None if rp.NoEndDate else fmt(rp.PatternEndDate)}
        except Exception:  # noqa: BLE001
            pass
        return d
    return await com(f)


def _applica_evento(e, oggetto, inizio, fine, durata_minuti, descrizione, luogo, tutto_il_giorno,
                    etichetta, categorie, promemoria_minuti) -> None:
    """Scrive sull'appuntamento solo i campi passati (None = non toccare). Condivisa fra crea_evento e modifica_evento.
    promemoria_minuti negativo = nessun promemoria. L'etichetta colorata va nella proprietà MAPI PROP_LABEL."""
    if oggetto is not None:
        e.Subject = oggetto
    if inizio is not None:
        e.Start = com_data(parse_dt(inizio))
    if fine is not None:
        e.End = com_data(parse_dt(fine))
    elif durata_minuti is not None:
        e.Duration = int(durata_minuti)
    if descrizione is not None:
        e.Body = descrizione
    if luogo is not None:
        e.Location = luogo
    if tutto_il_giorno is not None:
        e.AllDayEvent = bool(tutto_il_giorno)
    if categorie is not None:
        e.Categories = categorie
    if promemoria_minuti is not None:
        if int(promemoria_minuti) < 0:
            e.ReminderSet = False
        else:
            e.ReminderSet = True
            e.ReminderMinutesBeforeStart = int(promemoria_minuti)
    if etichetta is not None:
        k = etichetta.strip().lower()
        if k not in ETICHETTE:
            raise ValueError("Etichetta non valida. Valide: " + ", ".join(ETICHETTE))
        e.PropertyAccessor.SetProperty(PROP_LABEL, ETICHETTE[k])


# Codici OlRecurrenceType e maschera OlDaysOfWeek di Outlook.
TIPI_RICORRENZA = {"giornaliera": 0, "settimanale": 1, "mensile": 2, "annuale": 5}
GIORNI_SETTIMANA = {"dom": 1, "lun": 2, "mar": 4, "mer": 8, "gio": 16, "ven": 32, "sab": 64}
NOMI_RICORRENZA = {0: "giornaliera", 1: "settimanale", 2: "mensile", 3: "mensile (n-esimo giorno)", 5: "annuale", 6: "annuale (n-esimo giorno)"}


def _applica_ricorrenza(e, ricorrenza: str, intervallo: int | None, giorni_settimana: str | None,
                        fine_ricorrenza: str | None, occorrenze: int | None) -> dict:
    """Trasforma l'appuntamento in una serie. Va chiamata dopo aver impostato Start, End e Duration, perché il modello
    di ricorrenza (RecurrencePattern) prende da lì orario e durata. Ordine delle proprietà verificato sul campo:
    RecurrenceType, poi DayOfWeekMask o DayOfMonth/MonthOfYear, Interval, PatternStartDate, StartTime, EndTime, Duration,
    e per ultimo la fine (PatternEndDate, Occurrences oppure NoEndDate)."""
    k = (ricorrenza or "").strip().lower()
    if k not in TIPI_RICORRENZA:
        raise ValueError("Ricorrenza non valida: giornaliera, settimanale, mensile, annuale.")
    inizio = dt.datetime(e.Start.year, e.Start.month, e.Start.day, e.Start.hour, e.Start.minute)
    fine = dt.datetime(e.End.year, e.End.month, e.End.day, e.End.hour, e.End.minute)
    rp = e.GetRecurrencePattern()
    rp.RecurrenceType = TIPI_RICORRENZA[k]
    if k == "settimanale":
        mask = 0
        for g in [x.strip().lower()[:3] for x in (giorni_settimana or "").replace(";", ",").split(",") if x.strip()]:
            if g not in GIORNI_SETTIMANA:
                raise ValueError("Giorni della settimana non validi: usare lun, mar, mer, gio, ven, sab, dom separati da virgola.")
            mask |= GIORNI_SETTIMANA[g]
        if mask == 0:
            mask = [2, 4, 8, 16, 32, 64, 1][inizio.weekday()]  # il giorno di inizio
        rp.DayOfWeekMask = mask
    elif k == "mensile":
        rp.DayOfMonth = inizio.day
    elif k == "annuale":
        rp.MonthOfYear = inizio.month
        rp.DayOfMonth = inizio.day
    rp.Interval = max(1, int(intervallo or 1))
    rp.PatternStartDate = com_data(inizio.replace(hour=0, minute=0, second=0))
    rp.StartTime = com_data(inizio)
    rp.EndTime = com_data(fine)
    rp.Duration = int((fine - inizio).total_seconds() // 60)
    if fine_ricorrenza:
        rp.PatternEndDate = com_data(parse_dt(fine_ricorrenza).replace(hour=0, minute=0, second=0))
    elif occorrenze:
        rp.Occurrences = int(occorrenze)
    else:
        rp.NoEndDate = True
    try:
        fine_serie = None if rp.NoEndDate else fmt(rp.PatternEndDate)
    except Exception:  # noqa: BLE001
        fine_serie = None
    return {"tipo": k, "intervallo": rp.Interval, "giorni": giorni_settimana or None,
            "fine": fine_serie, "occorrenze": (int(occorrenze) if occorrenze and not fine_ricorrenza else None)}


@mcp.tool(annotations=RW)
async def crea_evento(oggetto: str, inizio: str, fine: str | None = None, durata_minuti: int | None = 30,
                      descrizione: str = "", luogo: str = "", tutto_il_giorno: bool = False,
                      etichetta: str | None = None, categorie: str | None = None, promemoria_minuti: int | None = 15,
                      casella: str | None = None, calendario_nome: str | None = None,
                      ricorrenza: str | None = None, intervallo: int | None = None, giorni_settimana: str | None = None,
                      fine_ricorrenza: str | None = None, occorrenze: int | None = None) -> dict:
    """Crea un evento o un memo nel calendario (predefinito: Calendario della casella predefinita).
    Date 'AAAA-MM-GG HH:MM'. Etichette (colori, visibili con server che le supportano, per esempio Kerio): telefonata
    (giallo, per i memo), lavoro, importante, personale, vacanza, da_non_perdere, viaggio, preparazione, compleanno,
    anniversario, nessuna. Categorie separate da ';'. promemoria_minuti = -1 per nessun promemoria. Non invia inviti.
    SERIE (opzionale): ricorrenza = giornaliera | settimanale | mensile | annuale; intervallo = ogni quanti giorni,
    settimane, mesi o anni (predefinito 1); giorni_settimana per la settimanale, es. "lun,mer,ven" (predefinito: il
    giorno di inizio); fine_ricorrenza = 'AAAA-MM-GG' oppure occorrenze = numero di volte; senza fine la serie
    non termina. Esempio: "ogni venerdì per 5 volte" = ricorrenza=settimanale, giorni_settimana=ven, occorrenze=5."""
    richiede_fase(3)

    def f():
        store = store_da_alias(casella)
        cal = calendario(store, calendario_nome)
        e = cal.Items.Add(1)  # olAppointmentItem: un appuntamento semplice, mai una riunione
        _applica_evento(e, oggetto, inizio, fine, durata_minuti, descrizione, luogo, tutto_il_giorno,
                        etichetta, categorie, promemoria_minuti)
        serie = None
        if ricorrenza:
            serie = _applica_ricorrenza(e, ricorrenza, intervallo, giorni_settimana, fine_ricorrenza, occorrenze)
        try:
            e.MeetingStatus = 0  # olNonMeeting: nessun invito, niente da spedire
        except Exception:  # noqa: BLE001
            pass
        e.Save()
        verifica_evento_puro(e)
        # Controllo finale: l'evento deve stare proprio nel calendario richiesto (difesa contro provider che spostano).
        try:
            if e.Parent.EntryID != cal.EntryID:
                raise ValueError(f"L'evento è stato salvato in '{e.Parent.FolderPath}' invece che nel calendario '{cal.FolderPath}': segnalarlo all'utente.")
        except ValueError:
            raise
        except Exception:  # noqa: BLE001
            pass
        audit("crea_evento", casella=casella or CASELLA_PREDEFINITA, calendario=cal.FolderPath, id=e.EntryID, oggetto=oggetto, serie=serie)
        out = meta_evento(e, casella or CASELLA_PREDEFINITA, cal.FolderPath)
        if serie:
            out["serie"] = serie
        return out
    return await com(f)


def occorrenza_di(e, data_occorrenza: str):
    """Restituisce la singola occorrenza di una serie che inizia alla data e ora indicate ('AAAA-MM-GG HH:MM', come
    il campo 'inizio' restituito da elenca_eventi), tramite RecurrencePattern.GetOccurrence. Errore chiaro se l'evento
    non è una serie o nessuna occorrenza inizia esattamente a quell'ora."""
    if not bool(e.IsRecurring):
        raise ValueError("L'evento non è una serie: data_occorrenza vale solo per gli eventi ricorrenti.")
    quando = parse_dt(data_occorrenza)
    rp = e.GetRecurrencePattern()
    try:
        return rp.GetOccurrence(com_data(quando))
    except Exception as ex:  # noqa: BLE001
        raise ValueError(f"Nessuna occorrenza della serie inizia il {quando.strftime('%d.%m.%Y alle %H:%M')}: usare l'inizio esatto riportato da elenca_eventi. ({ex})") from ex


@mcp.tool(annotations=RW)
async def modifica_evento(id: str, oggetto: str | None = None, inizio: str | None = None, fine: str | None = None,
                          durata_minuti: int | None = None, descrizione: str | None = None, luogo: str | None = None,
                          tutto_il_giorno: bool | None = None, etichetta: str | None = None, categorie: str | None = None,
                          promemoria_minuti: int | None = None, casella: str | None = None,
                          data_occorrenza: str | None = None) -> dict:
    """Modifica solo i campi indicati di un evento o memo (id da elenca_eventi). Per un evento ricorrente modifica
    l'intera serie, oppure una sola occorrenza se si passa data_occorrenza ('AAAA-MM-GG HH:MM', l'inizio riportato
    da elenca_eventi): quella occorrenza diventa un'eccezione della serie. Non cancella mai."""
    richiede_fase(3)

    def f():
        e, _ = elemento(id, casella)
        verifica_classe(e, 26, "un evento")
        verifica_evento_puro(e)
        bersaglio = occorrenza_di(e, data_occorrenza) if data_occorrenza else e
        _applica_evento(bersaglio, oggetto, inizio, fine, durata_minuti, descrizione, luogo, tutto_il_giorno,
                        etichetta, categorie, promemoria_minuti)
        bersaglio.Save()
        audit("modifica_evento", casella=casella or CASELLA_PREDEFINITA, id=id, occorrenza=data_occorrenza)
        out = meta_evento(bersaglio, casella or CASELLA_PREDEFINITA)
        if data_occorrenza:
            out["id"] = id
            out["occorrenza_modificata"] = data_occorrenza
            out["nota"] = "modificata solo questa occorrenza (eccezione della serie); l'id resta quello della serie"
        return out
    return await com(f)


# ---------------------------------------------------------------- fase 4: avvisi (unica eccezione all'invio)

@mcp.tool(annotations=RW)
async def invia_avviso(oggetto: str, corpo: str) -> dict:
    """Invia un avviso automatico (per esempio da un'attività programmata) all'utente stesso. Mittente e destinatario
    sono fissi nella configurazione del connettore e NON sono parametri: è l'unica funzione di invio e non può scrivere
    a nessun altro. Testo semplice."""
    richiede_fase(4)
    if not AVVISO_MITTENTE or not AVVISO_DESTINATARIO:
        raise ToolError('Avvisi non configurati: in config.json manca la sezione "avviso" con "mittente" e "destinatario".')

    def f():
        # L'account mittente deve esistere in Outlook: si cerca per indirizzo SMTP.
        acc = None
        for a in ns().Accounts:
            if (a.SmtpAddress or "").lower() == AVVISO_MITTENTE:
                acc = a
                break
        if acc is None:
            raise RuntimeError(f"Account mittente {AVVISO_MITTENTE} non trovato in Outlook.")
        m = COM.outlook().CreateItem(0)  # olMailItem
        m.SendUsingAccount = acc
        m.To = AVVISO_DESTINATARIO
        m.Subject = AVVISO_PREFISSO + oggetto
        m.Body = corpo
        if m.To != AVVISO_DESTINATARIO:  # difesa in profondità: nessun altro destinatario possibile
            raise RuntimeError("Destinatario non conforme.")
        oggetto_finale = m.Subject  # dopo Send() l'oggetto COM non è più leggibile: i dati si prendono prima
        m.Send()
        audit("invia_avviso", oggetto=oggetto_finale)
        return {"inviato": True, "da": AVVISO_MITTENTE, "a": AVVISO_DESTINATARIO, "oggetto": oggetto_finale}
    return await com(f)


# ---------------------------------------------------------------- rubrica (contatti)

@mcp.tool(annotations=RO)
async def cerca_contatti(testo: str, casella: str | None = None, cartella: str | None = None, massimo: int = 30) -> list[dict]:
    """Cerca persone e aziende nella rubrica per nome, cognome, azienda, indirizzo email o numero di telefono.
    Senza 'cartella' cerca in tutte le cartelle di contatti della casella (Contatti e sottocartelle, altre rubriche).
    Restituisce i dati principali; per la scheda completa usare leggi_contatto."""
    def f():
        store = store_da_alias(casella)
        cartelle = [cartella_da_percorso(store, cartella)] if cartella else cartelle_di_tipo(store, 2)
        t = dasl_lit(testo)
        # Filtro DASL su tutti i campi utili: un solo passaggio sul server invece di leggere ogni contatto.
        # Se il provider non supporta il filtro (com_error), si ripiega sulla scansione dei campi principali.
        campi = ["urn:schemas:contacts:cn", "urn:schemas:contacts:fileas", "urn:schemas:contacts:givenName",
                 "urn:schemas:contacts:sn", "urn:schemas:contacts:o", "urn:schemas:contacts:email1",
                 "urn:schemas:contacts:email2", "urn:schemas:contacts:telephoneNumber", "urn:schemas:contacts:mobile",
                 "urn:schemas:contacts:homePhone"]
        filtro = "@SQL=" + " OR ".join(f"\"{c}\" LIKE '%{t}%'" for c in campi)
        out: list[dict] = []
        lim = max(1, min(massimo, 200))
        for fold in cartelle:
            try:
                items = fold.Items.Restrict(filtro)
                for c in itera(items, lim - len(out)):
                    if c.Class == 40:
                        out.append(meta_contatto(c, casella or CASELLA_PREDEFINITA))
            except pywintypes.com_error:
                # ripiego: scansione semplice sui campi principali
                tl = testo.lower()
                for c in itera(fold.Items, 2000):
                    try:
                        if c.Class != 40:
                            continue
                        blob = " ".join(str(getattr(c, a) or "") for a in ("FullName", "CompanyName", "Email1Address", "MobileTelephoneNumber", "BusinessTelephoneNumber"))
                        if tl in blob.lower():
                            out.append(meta_contatto(c, casella or CASELLA_PREDEFINITA))
                            if len(out) >= lim:
                                break
                    except Exception:  # noqa: BLE001
                        continue
            if len(out) >= lim:
                break
        return out
    return await com(f)


@mcp.tool(annotations=RO)
async def elenca_contatti(casella: str | None = None, cartella: str = "Contatti", massimo: int = 100) -> list[dict]:
    """Elenca i contatti di una cartella della rubrica (predefinita: Contatti; altre per percorso, per esempio
    'Contatti/<sottocartella>', vedi elenca_cartelle), in ordine alfabetico."""
    def f():
        store = store_da_alias(casella)
        fold = cartella_da_percorso(store, cartella)
        items = fold.Items
        items.Sort("[FileAs]")
        return [meta_contatto(c, casella or CASELLA_PREDEFINITA) for c in itera(items, max(1, min(massimo, 500))) if c.Class == 40]
    return await com(f)


@mcp.tool(annotations=RO)
async def leggi_contatto(id: str, casella: str | None = None) -> dict:
    """Legge la scheda completa di un contatto (id da cerca_contatti/elenca_contatti)."""
    def f():
        c, _ = elemento(id, casella)
        verifica_classe(c, 40, "un contatto")
        return meta_contatto(c, casella or CASELLA_PREDEFINITA, completo=True)
    return await com(f)


@mcp.tool(annotations=RW)
async def crea_contatto(nome: str = "", cognome: str = "", azienda: str = "", ruolo: str = "", email: str = "",
                        email2: str = "", tel_lavoro: str = "", cellulare: str = "", tel_casa: str = "",
                        indirizzo_lavoro: str = "", sito_web: str = "", note: str = "", categorie: str = "",
                        casella: str | None = None, cartella: str = "Contatti") -> dict:
    """Crea un contatto nella rubrica (cartella predefinita: Contatti della casella predefinita). Serve almeno
    nome, cognome o azienda. Non cancella e non modifica contatti esistenti: per quello usare modifica_contatto."""
    richiede_fase(2)

    def f():
        if not (nome or cognome or azienda):
            raise ValueError("Indicare almeno nome, cognome o azienda.")
        store = store_da_alias(casella)
        fold = cartella_da_percorso(store, cartella)
        c = fold.Items.Add(2)
        _applica_contatto(c, nome, cognome, azienda, ruolo, email, email2, tel_lavoro, cellulare, tel_casa,
                          indirizzo_lavoro, sito_web, note, categorie)
        c.Save()
        audit("crea_contatto", casella=casella or CASELLA_PREDEFINITA, cartella=fold.FolderPath, id=c.EntryID)
        return meta_contatto(c, casella or CASELLA_PREDEFINITA, completo=True)
    return await com(f)


@mcp.tool(annotations=RW)
async def modifica_contatto(id: str, nome: str | None = None, cognome: str | None = None, azienda: str | None = None,
                            ruolo: str | None = None, email: str | None = None, email2: str | None = None,
                            tel_lavoro: str | None = None, cellulare: str | None = None, tel_casa: str | None = None,
                            indirizzo_lavoro: str | None = None, sito_web: str | None = None, note: str | None = None,
                            categorie: str | None = None, casella: str | None = None) -> dict:
    """Modifica solo i campi indicati di un contatto esistente (id da cerca_contatti). Non cancella mai."""
    richiede_fase(2)

    def f():
        c, _ = elemento(id, casella)
        verifica_classe(c, 40, "un contatto")
        _applica_contatto(c, nome, cognome, azienda, ruolo, email, email2, tel_lavoro, cellulare, tel_casa,
                          indirizzo_lavoro, sito_web, note, categorie)
        c.Save()
        audit("modifica_contatto", casella=casella or CASELLA_PREDEFINITA, id=id)
        return meta_contatto(c, casella or CASELLA_PREDEFINITA, completo=True)
    return await com(f)


# ---------------------------------------------------------------- attività

@mcp.tool(annotations=RO)
async def elenca_attivita(casella: str | None = None, solo_aperte: bool = True, massimo: int = 100,
                          cartella: str | None = None) -> list[dict]:
    """Elenca le attività (cartella Attività della casella), per data di scadenza. solo_aperte=false include le completate."""
    def f():
        store = store_da_alias(casella)
        fold = cartella_da_percorso(store, cartella) if cartella else cartella_predefinita(store, 13, "Attività", "Tasks")
        items = fold.Items
        items.Sort("[DueDate]")
        if solo_aperte:
            items = items.Restrict("[Complete] = False")
        return [meta_attivita(t, casella or CASELLA_PREDEFINITA) for t in itera(items, max(1, min(massimo, 500))) if t.Class == 48]
    return await com(f)


@mcp.tool(annotations=RO)
async def leggi_attivita(id: str, casella: str | None = None) -> dict:
    """Legge un'attività completa (id da elenca_attivita)."""
    def f():
        t, _ = elemento(id, casella)
        verifica_classe(t, 48, "un'attività")
        return meta_attivita(t, casella or CASELLA_PREDEFINITA, completo=True)
    return await com(f)


@mcp.tool(annotations=RW)
async def crea_attivita(oggetto: str, scadenza: str | None = None, inizio: str | None = None, descrizione: str = "",
                        priorita: str = "normale", promemoria: str | None = None, categorie: str = "",
                        casella: str | None = None, cartella: str | None = None) -> dict:
    """Crea un'attività (cartella Attività della casella predefinita). Date 'AAAA-MM-GG' o 'AAAA-MM-GG HH:MM';
    priorita: bassa, normale, alta; promemoria = data e ora del promemoria."""
    richiede_fase(2)

    def f():
        store = store_da_alias(casella)
        fold = cartella_da_percorso(store, cartella) if cartella else cartella_predefinita(store, 13, "Attività", "Tasks")
        t = fold.Items.Add(3)
        _applica_attivita(t, oggetto, scadenza, inizio, descrizione, priorita, promemoria, categorie, None, None)
        t.Save()
        audit("crea_attivita", casella=casella or CASELLA_PREDEFINITA, cartella=fold.FolderPath, id=t.EntryID, oggetto=oggetto)
        return meta_attivita(t, casella or CASELLA_PREDEFINITA, completo=True)
    return await com(f)


@mcp.tool(annotations=RW)
async def modifica_attivita(id: str, oggetto: str | None = None, scadenza: str | None = None, inizio: str | None = None,
                            descrizione: str | None = None, priorita: str | None = None, promemoria: str | None = None,
                            categorie: str | None = None, completata: bool | None = None, percentuale: int | None = None,
                            casella: str | None = None) -> dict:
    """Modifica solo i campi indicati di un'attività; completata=true la segna come completata, false la riapre.
    promemoria='' toglie il promemoria. Non cancella mai."""
    richiede_fase(2)

    def f():
        t, _ = elemento(id, casella)
        verifica_classe(t, 48, "un'attività")
        _applica_attivita(t, oggetto, scadenza, inizio, descrizione, priorita, promemoria, categorie, completata, percentuale)
        t.Save()
        audit("modifica_attivita", casella=casella or CASELLA_PREDEFINITA, id=id, completata=completata)
        return meta_attivita(t, casella or CASELLA_PREDEFINITA, completo=True)
    return await com(f)


# ---------------------------------------------------------------- note

@mcp.tool(annotations=RO)
async def elenca_note(casella: str | None = None, massimo: int = 100) -> list[dict]:
    """Elenca le note (cartella Note della casella), dalla più recente."""
    def f():
        store = store_da_alias(casella)
        fold = cartella_predefinita(store, 12, "Note", "Notes")
        items = fold.Items
        items.Sort("[LastModificationTime]", True)
        return [meta_nota(n, casella or CASELLA_PREDEFINITA) for n in itera(items, max(1, min(massimo, 500))) if n.Class == 44]
    return await com(f)


@mcp.tool(annotations=RO)
async def leggi_nota(id: str, casella: str | None = None) -> dict:
    """Legge il testo completo di una nota (id da elenca_note)."""
    def f():
        n, _ = elemento(id, casella)
        verifica_classe(n, 44, "una nota")
        return meta_nota(n, casella or CASELLA_PREDEFINITA, completo=True)
    return await com(f)


@mcp.tool(annotations=RW)
async def crea_nota(testo: str, categorie: str = "", casella: str | None = None) -> dict:
    """Crea una nota (la prima riga del testo diventa il titolo)."""
    richiede_fase(2)

    def f():
        store = store_da_alias(casella)
        fold = cartella_predefinita(store, 12, "Note", "Notes")
        n = fold.Items.Add(5)
        n.Body = testo
        if categorie:
            n.Categories = categorie
        n.Save()
        audit("crea_nota", casella=casella or CASELLA_PREDEFINITA, id=n.EntryID)
        return meta_nota(n, casella or CASELLA_PREDEFINITA, completo=True)
    return await com(f)


@mcp.tool(annotations=RW)
async def modifica_nota(id: str, testo: str | None = None, categorie: str | None = None, casella: str | None = None) -> dict:
    """Sostituisce il testo e/o le categorie di una nota esistente. Non cancella mai."""
    richiede_fase(2)

    def f():
        n, _ = elemento(id, casella)
        verifica_classe(n, 44, "una nota")
        if testo is not None:
            n.Body = testo
        if categorie is not None:
            n.Categories = categorie
        n.Save()
        audit("modifica_nota", casella=casella or CASELLA_PREDEFINITA, id=id)
        return meta_nota(n, casella or CASELLA_PREDEFINITA, completo=True)
    return await com(f)


# ---------------------------------------------------------------- cartelle e spostamenti

@mcp.tool(annotations=RW)
async def crea_cartella(nome: str, percorso_padre: str = "", tipo: str = "posta", casella: str | None = None) -> dict:
    """Crea una sottocartella (tipo: posta, calendario, contatti, attivita, note) sotto 'percorso_padre'
    (vuoto = radice della casella). Non rinomina e non cancella cartelle."""
    richiede_fase(2)

    def f():
        codici = {"posta": 6, "calendario": 9, "contatti": 10, "attivita": 13, "note": 12}
        if tipo not in codici:
            raise ValueError("Tipo non valido: posta, calendario, contatti, attivita, note.")
        store = store_da_alias(casella)
        padre = cartella_da_percorso(store, percorso_padre)
        for sub in padre.Folders:
            if sub.Name.lower() == nome.lower():
                raise ValueError(f"Esiste già una cartella '{sub.Name}' in '{padre.FolderPath}'.")
        nuova = padre.Folders.Add(nome, codici[tipo])
        audit("crea_cartella", casella=casella or CASELLA_PREDEFINITA, cartella=nuova.FolderPath)
        return {"percorso": nuova.FolderPath, "tipo": tipo}
    return await com(f)


@mcp.tool(annotations=RW)
async def sposta_mail(id: str, cartella_destinazione: str, casella: str | None = None,
                      casella_destinazione: str | None = None) -> dict:
    """Sposta una mail in un'altra cartella (anche di un'altra casella). Rifiuta qualsiasi destinazione che sia
    un cestino (Posta eliminata, Cestino, Trash): il connettore non cancella. Lo stato letto/non letto resta com'era."""
    richiede_fase(2)

    def f():
        item, store = elemento(id, casella)
        verifica_classe(item, 43, "una mail")
        store_dest = store_da_alias(casella_destinazione) if casella_destinazione else store
        dest = cartella_da_percorso(store_dest, cartella_destinazione)
        if e_cestino(store_dest, dest):
            raise ValueError("Destinazione rifiutata: è un cestino, e il connettore non cancella.")
        if dest.DefaultItemType != 0:
            raise ValueError("La cartella di destinazione non è una cartella di posta.")
        era_non_letta = bool(item.UnRead)
        nuovo = item.Move(dest)
        try:
            if bool(nuovo.UnRead) != era_non_letta:
                nuovo.UnRead = era_non_letta
                nuovo.Save()
        except Exception:  # noqa: BLE001
            pass
        audit("sposta_mail", casella=casella or CASELLA_PREDEFINITA, id=id, destinazione=dest.FolderPath, nuovo_id=nuovo.EntryID)
        return {"id": nuovo.EntryID, "casella": casella_destinazione or casella or CASELLA_PREDEFINITA, "cartella": dest.FolderPath}
    return await com(f)


# ---------------------------------------------------------------- cancellazione (nel cestino, recuperabile)

# Strumenti marcati come distruttivi: Claude desktop può chiedere conferma all'utente prima di eseguirli.
DEL = ToolAnnotations(read_only_hint=False, destructive_hint=True, idempotent_hint=False, open_world_hint=False)


def richiede_cancellazione() -> None:
    """Blocca gli strumenti di cancellazione se l'interruttore in config.json è spento (o assente)."""
    if not CANCELLAZIONE:
        raise ToolError('Cancellazione disattivata: in config.json "cancellazione_abilitata" non è true.')


def cestino_di(store):
    """Il cestino della casella: cartella predefinita 3 (olFolderDeletedItems), oppure, con provider che non la
    dichiarano (IMAP, caselle PEC), una cartella di posta con uno dei nomi in NOMI_CESTINO nei primi due livelli."""
    try:
        return store.GetDefaultFolder(3)
    except Exception:  # noqa: BLE001
        pass

    def walk(fold, liv):
        for sub in fold.Folders:
            try:
                if sub.Name.lower() in NOMI_CESTINO and sub.DefaultItemType == 0:
                    return sub
            except Exception:  # noqa: BLE001
                continue
            if liv < 2:
                r = walk(sub, liv + 1)
                if r is not None:
                    return r
        return None

    f = walk(store.GetRootFolder(), 1)
    if f is None:
        raise ValueError("Questa casella non ha un cestino riconoscibile: cancellazione rifiutata.")
    return f


def controlla_oggetto(item, oggetto_atteso: str) -> str:
    """Prima di cancellare, l'oggetto reale deve coincidere con quello che l'utente ha confermato: protegge da un id
    sbagliato o da un elemento spostato nel frattempo."""
    reale = (item.Subject or "").strip()
    if reale != (oggetto_atteso or "").strip():
        raise ValueError(f"Cancellazione rifiutata: l'oggetto reale è «{reale}», diverso da quello confermato «{oggetto_atteso}».")
    return reale


@mcp.tool(annotations=DEL)
async def elimina_mail(id: str, oggetto_atteso: str, casella: str | None = None) -> dict:
    """Sposta una mail nel cestino della sua casella (recuperabile). USARE SOLO dopo che l'utente ha confermato
    in chat quella mail precisa. 'oggetto_atteso' deve coincidere con l'oggetto reale, altrimenti rifiuta.
    Rifiuta le mail già nel cestino (niente cancellazione definitiva)."""
    richiede_fase(2)
    richiede_cancellazione()

    def f():
        item, store = elemento(id, casella)
        verifica_classe(item, 43, "una mail")
        ogg = controlla_oggetto(item, oggetto_atteso)
        if e_cestino(store, item.Parent):
            raise ValueError("La mail è già nel cestino: la cancellazione definitiva non è permessa.")
        dest = cestino_di(store)
        origine = item.Parent.FolderPath
        nuovo = item.Move(dest)
        audit("elimina_mail", casella=casella or CASELLA_PREDEFINITA, id=id, oggetto=ogg, da=origine, a=dest.FolderPath)
        return {"eliminata": True, "oggetto": ogg, "da": origine, "cestino": dest.FolderPath, "id_nel_cestino": nuovo.EntryID}
    return await com(f)


@mcp.tool(annotations=DEL)
async def elimina_evento(id: str, oggetto_atteso: str, serie_intera: bool = False, casella: str | None = None,
                         data_occorrenza: str | None = None) -> dict:
    """Sposta un evento o memo del calendario nel cestino della sua casella (recuperabile). USARE SOLO dopo che
    l'utente ha confermato in chat quell'evento preciso. 'oggetto_atteso' deve coincidere con l'oggetto reale.
    Per un evento ricorrente: serie_intera=true cancella tutta la serie (nel cestino); data_occorrenza
    ('AAAA-MM-GG HH:MM', l'inizio riportato da elenca_eventi) cancella SOLO quella occorrenza. ATTENZIONE: la singola
    occorrenza non passa dal cestino (diventa un'eccezione della serie) e si recupera solo ricreandola: va detto
    all'utente prima della conferma. Rifiuta le riunioni con invitati (cancellarle manderebbe disdette) e gli eventi
    già nel cestino."""
    richiede_fase(3)
    richiede_cancellazione()

    def f():
        e, store = elemento(id, casella)
        verifica_classe(e, 26, "un evento")
        ogg = controlla_oggetto(e, oggetto_atteso)
        if serie_intera and data_occorrenza:
            raise ValueError("Indicare o serie_intera=true o data_occorrenza, non entrambi.")
        try:
            stato = int(e.MeetingStatus)
        except Exception:  # noqa: BLE001
            stato = 0
        if stato != 0:
            raise ValueError("È una riunione con invitati: cancellarla potrebbe mandare disdette. Va cancellata a mano dall'utente in Outlook.")
        if e_cestino(store, e.Parent):
            raise ValueError("L'evento è già nel cestino: la cancellazione definitiva non è permessa.")
        if data_occorrenza:
            occ = occorrenza_di(e, data_occorrenza)
            inizio_occ = fmt(occ.Start)
            occ.Delete()
            audit("elimina_evento", casella=casella or CASELLA_PREDEFINITA, id=id, oggetto=ogg, occorrenza=inizio_occ,
                  serie_intera=False, nota="singola occorrenza, senza cestino")
            return {"eliminato": True, "oggetto": ogg, "occorrenza": inizio_occ, "serie_intera": False,
                    "nota": "cancellata solo questa occorrenza: la serie continua; non passa dal cestino, si recupera solo ricreandola"}
        if bool(e.IsRecurring) and not serie_intera:
            raise ValueError("Evento ricorrente: serve serie_intera=true (sparisce tutta la serie) oppure data_occorrenza (solo quella occorrenza).")
        inizio = fmt(e.Start)
        dest = cestino_di(store)
        origine = e.Parent.FolderPath
        # Delete(), non Move(): è la stessa operazione del tasto Canc in Outlook. Verificato sul campo che con alcuni
        # provider (Kerio Connect via KOFF) un appuntamento spostato con Move() in Posta eliminata, che è una cartella
        # di posta, viene convertito in un messaggio con allegato .msg e accenti rotti; con Delete() resta un elemento
        # di calendario ripristinabile.
        e.Delete()
        # Si ritrova l'evento nel cestino per restituirne il nuovo id (comodo per ripristinarlo o verificarlo).
        id_cestino = None
        try:
            r = dest.Items.Restrict("@SQL=\"urn:schemas:httpmail:subject\" = '" + dasl_lit(ogg) + "'")
            for it in itera(r, 20):
                if it.Class == 26:
                    id_cestino = it.EntryID
                    break
        except Exception:  # noqa: BLE001
            pass
        audit("elimina_evento", casella=casella or CASELLA_PREDEFINITA, id=id, oggetto=ogg, inizio=inizio,
              serie_intera=bool(serie_intera), da=origine, a=dest.FolderPath, id_nel_cestino=id_cestino)
        return {"eliminato": True, "oggetto": ogg, "inizio": inizio, "serie_intera": bool(serie_intera),
                "da": origine, "cestino": dest.FolderPath, "id_nel_cestino": id_cestino,
                "nota": "spostato nel cestino con Delete, come dal tasto Canc di Outlook: resta un elemento di calendario ripristinabile"}
    return await com(f)


if __name__ == "__main__":
    # Avviato da Claude desktop con il trasporto stdio: il server legge le richieste MCP dallo standard input e
    # risponde sullo standard output. Per questo i log vanno su file e su stderr, mai su stdout.
    log.info("Avvio outlook-classic-mcp, fase %s, timeout %s s, cancellazione %s", FASE, TIMEOUT, CANCELLAZIONE)
    mcp.run(transport="stdio")
