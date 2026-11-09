# prova_accenti.py - prova automatica di outlook-classic-mcp, con lettere accentate.
#
# QUANDO: dopo l'installazione e dopo ogni modifica a server.py, prima di riavviare Claude desktop. Gira sul codice
# su disco (importa server.py direttamente), quindi non serve che Claude desktop sia aperto; serve Outlook aperto.
# COME (dalla cartella del connettore, in PowerShell):
#   .\.venv\Scripts\python.exe .\prova_accenti.py            esegue la prova e ripulisce
#   .\.venv\Scripts\python.exe .\prova_accenti.py --tieni    lascia gli elementi creati, per guardarli in Outlook
#
# COSA FA: crea bozze (testo, HTML, risposta all'ultima mail ricevuta, con allegato), salva una mail come .msg e il suo
# allegato con impronta SHA-256, legge le intestazioni di trasporto, crea un evento e lo cancella (controllando che nel
# cestino resti un evento), crea una serie settimanale di 3 e ne modifica e cancella singole occorrenze, poi la serie
# intera, crea una nota e un contatto. Ogni testo viene riletto con le funzioni del connettore e deve tornare identico,
# senza i segni tipici degli accenti corrotti ("Ã", "â€", "Â", "\ufffd"). Tutti gli elementi hanno "PROVA ACCENTI" nel
# titolo. Alla fine CANCELLA SOLO gli elementi che ha creato lei stessa (riconosciuti dall'id), spostandoli nel cestino.
# Non invia niente: invia_avviso non fa parte della prova.
#
# COSA SERVE: config.json compilato, Outlook Classic aperto come utente normale, almeno una mail ricevuta negli ultimi
# 7 giorni nella casella predefinita, "fase": 3 o 4 (la prova scrive bozze ed eventi) e "cancellazione_abilitata": true
# per i controlli di cancellazione (altrimenti quelli vengono saltati e segnati come tali). I file di prova vanno in una sottocartella
# "prova_accenti_<data>" della PRIMA cartella di "cartelle_salvataggio", creata e poi cancellata dalla prova.
# ESITO: l'ultima riga deve essere "ESITO: TUTTO OK".

import asyncio
import datetime as dt
import hashlib
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import server  # noqa: E402

TESTO = "Prova accenti: già servirà così perché più città è caffè àèéìòù ÀÈÉÌÒÙ € l'apostrofo tipografico ’ e le virgolette “doppie”."
SPIE = ("Ã", "â€", "Â", "\ufffd")
TIENI = "--tieni" in sys.argv
creati: list[tuple[str, str]] = []  # (descrizione, entry_id)
esiti: list[tuple[str, bool, str]] = []
# Destinatario fittizio delle bozze (non vengono mai inviate) e cartella dei file di prova, dentro la prima consentita.
DESTINATARIO_PROVA = "prova.accenti@example.com"
CARTELLA_PROVA = server.CARTELLE_SALVATAGGIO[0] / dt.datetime.now().strftime("prova_accenti_%Y%m%d_%H%M%S")


def sha(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def controlla_impronta(nome: str, file: str, impronta: str, file_imp: str) -> None:
    f, fi = Path(file), Path(file_imp)
    ok = f.is_file() and fi.is_file() and sha(f) == impronta
    if ok:
        righe = fi.read_text("utf-8").splitlines()
        ok = righe[-1] == f"{impronta}  {f.name}" and any(r.startswith("# acquisito: ") for r in righe)
    esiti.append((nome, ok, "" if ok else f"file={file} impronta={impronta} file_sha256={file_imp}"))


async def stato_bozza(eid: str) -> tuple[int, str]:
    """Pagina di codice e HTML che Outlook ha davvero salvato nella bozza (leggi_mail in testo non basta: il 04.11.2026
    la prova era OK e il 06.11.2026 una bozza in formato testo aveva gli accenti corrotti)."""
    def f():
        it = server.ns().GetItemFromID(eid)
        return int(it.InternetCodepage), it.HTMLBody or ""
    return await server.com(f)


async def verifica_bozza_salvata(nome: str, eid: str) -> None:
    cp, html = await stato_bozza(eid)
    ok = cp == server.UTF8_CPID and not any(sp in html for sp in SPIE) and "citt" in html
    esiti.append((nome + ": UTF-8 e HTML senza spie", ok, "" if ok else f"InternetCodepage={cp} html={html[:200]}"))


def verifica(nome: str, letto: str, atteso: str = TESTO) -> None:
    ok = atteso in (letto or "") and not any(sp in (letto or "") for sp in SPIE)
    dettaglio = "" if ok else (letto or "")[:300].replace("\r", " ").replace("\n", " ")
    esiti.append((nome, ok, dettaglio))


async def main() -> None:
    stato = await server.stato_outlook()
    print("Outlook", stato["versione_outlook"], "fase", stato["fase"])

    # 1. bozza di testo
    b = await server.crea_bozza(oggetto="PROVA ACCENTI bozza testo", corpo=TESTO, a=DESTINATARIO_PROVA)
    creati.append(("bozza testo", b["id"]))
    r = await server.leggi_mail(b["id"], formato="testo")
    verifica("crea_bozza (testo)", r["corpo"])
    await verifica_bozza_salvata("crea_bozza (testo)", b["id"])

    # 2. bozza HTML
    b = await server.crea_bozza(oggetto="PROVA ACCENTI bozza html", corpo=server.testo_a_html(TESTO),
                                a=DESTINATARIO_PROVA, formato="html")
    creati.append(("bozza html", b["id"]))
    r = await server.leggi_mail(b["id"], formato="testo")
    verifica("crea_bozza (html)", r["corpo"])
    await verifica_bozza_salvata("crea_bozza (html)", b["id"])

    # 3. bozza di risposta all'ultima mail della Posta in arrivo
    mails = await server.elenca_mail(giorni=7, massimo=1)
    if mails:
        orig = mails[0]
        orig_testo = (await server.leggi_mail(orig["id"], formato="testo"))["corpo"]
        b = await server.crea_bozza_risposta(orig["id"], corpo=TESTO)
        creati.append(("bozza risposta", b["id"]))
        r = await server.leggi_mail(b["id"], formato="testo")
        verifica("crea_bozza_risposta (testo nuovo)", r["corpo"])
        # la citazione dell'originale non deve contenere spie di codifica
        spie_orig = any(sp in orig_testo for sp in SPIE)
        spie_risp = any(sp in (r["corpo"] or "") for sp in SPIE)
        esiti.append(("crea_bozza_risposta (citazione originale)", spie_orig or not spie_risp,
                      "" if (spie_orig or not spie_risp) else (r["corpo"] or "")[:300]))
    else:
        esiti.append(("crea_bozza_risposta", False, "nessuna mail negli ultimi 7 giorni"))

    # 3b. salva_mail e salva_allegati con impronta (07.11.2026)
    CARTELLA_PROVA.mkdir(parents=True, exist_ok=False)
    allegato = CARTELLA_PROVA / "sorgente" / "allegato già più.txt"
    allegato.parent.mkdir()
    allegato.write_text(TESTO, "utf-8")
    ogg = "PROVA ACCENTI salva_mail già più: perché? <ok>"
    b = await server.crea_bozza(oggetto=ogg, corpo=TESTO, a=DESTINATARIO_PROVA,
                                allegati=[str(allegato)])
    creati.append(("bozza con allegato", b["id"]))
    await server.segna_mail(b["id"], letta=False)  # la rende non letta: salva_mail e salva_allegati devono lasciarla così
    dest = CARTELLA_PROVA / "salvate"
    s1 = await server.salva_mail(b["id"], str(dest))
    p1 = Path(s1["salvato_in"])
    esiti.append(("salva_mail (file .msg con nome ripulito e accenti)",
                  p1.suffix == ".msg" and "PROVA ACCENTI salva_mail già più perché ok" in p1.name
                  and p1.stat().st_size > 1000, p1.name))
    controlla_impronta("salva_mail (impronta SHA-256 e file .sha256)", s1["salvato_in"], s1["sha256"], s1["file_sha256"])
    s2 = await server.salva_mail(b["id"], str(dest))
    esiti.append(("salva_mail (mai sovrascrivere: secondo file distinto)",
                  s2["salvato_in"] != s1["salvato_in"] and sha(p1) == s1["sha256"], s2["salvato_in"]))
    r = await server.leggi_mail(b["id"])
    esiti.append(("salva_mail (mail non letta resta non letta)", bool(r["non_letta"]) and bool(s1["non_letta"]),
                  f"non_letta={r['non_letta']}"))
    a1 = await server.salva_allegati(b["id"], str(dest))
    ok_a = len(a1) == 1 and Path(a1[0]["salvato_in"]).read_text("utf-8") == TESTO
    esiti.append(("salva_allegati (allegato identico)", ok_a, str(a1)[:200]))
    if a1:
        controlla_impronta("salva_allegati (impronta SHA-256 e file .sha256)", a1[0]["salvato_in"], a1[0]["sha256"],
                           a1[0]["file_sha256"])
    r = await server.leggi_mail(b["id"])
    esiti.append(("salva_allegati (mail non letta resta non letta)", bool(r["non_letta"]), f"non_letta={r['non_letta']}"))
    # una cartella FUORI da quelle consentite deve essere rifiutata (la cartella temporanea di Windows, se non consentita)
    fuori = Path(tempfile.gettempdir()) / "prova_accenti_fuori"
    if server.percorso_consentito(fuori):
        esiti.append(("salva_mail (cartella non consentita rifiutata)", True, "saltato: la cartella temporanea è fra quelle consentite"))
    else:
        try:
            await server.salva_mail(b["id"], str(fuori))
            esiti.append(("salva_mail (cartella non consentita rifiutata)", False, f"ha salvato in {fuori}"))
        except Exception:  # noqa: BLE001
            esiti.append(("salva_mail (cartella non consentita rifiutata)", True, ""))

    # 3c. intestazioni di trasporto e salva_mail su una mail vera ricevuta (non letta, se ce n'è una)
    vere = await server.elenca_mail(giorni=7, massimo=1, solo_non_lette=True) or await server.elenca_mail(giorni=7, massimo=1)
    if vere:
        v = vere[0]
        r = await server.leggi_mail(v["id"], intestazioni=True, max_caratteri=200)
        hdr = r.get("intestazioni_trasporto") or ""
        esiti.append(("leggi_mail intestazioni=true (Received presenti)", "received:" in hdr.lower(),
                      r.get("intestazioni_nota", "") or hdr[:200]))
        esiti.append(("leggi_mail intestazioni=true (mail non letta resta non letta)", r["non_letta"] == v["non_letta"],
                      f"prima={v['non_letta']} dopo={r['non_letta']}"))
        sv = await server.salva_mail(v["id"], str(dest))
        dopo = await server.leggi_mail(v["id"], max_caratteri=10)
        controlla_impronta("salva_mail su mail ricevuta (impronta)", sv["salvato_in"], sv["sha256"], sv["file_sha256"])
        esiti.append(("salva_mail su mail ricevuta (stato letto/non letto invariato"
                      + (", era NON letta)" if v["non_letta"] else ", era già letta: verifica parziale)"),
                      dopo["non_letta"] == v["non_letta"], f"prima={v['non_letta']} dopo={dopo['non_letta']}"))
    else:
        esiti.append(("intestazioni e salva_mail su mail ricevuta", False, "nessuna mail negli ultimi 7 giorni"))

    # 4. evento
    inizio = (dt.datetime.now() + dt.timedelta(days=1)).replace(hour=15, minute=0, second=0, microsecond=0)
    e = await server.crea_evento(oggetto="PROVA ACCENTI evento già così", inizio=inizio.strftime("%Y-%m-%d %H:%M"),
                                 durata_minuti=15, descrizione=TESTO, etichetta="telefonata", promemoria_minuti=-1)
    ev = await server.leggi_evento(e["id"])
    verifica("crea_evento (descrizione)", ev["descrizione"])
    verifica("crea_evento (oggetto)", ev["oggetto"], "PROVA ACCENTI evento già così")
    # cancellazione con elimina_evento: nel cestino deve restare un EVENTO (classe 26) con gli accenti intatti
    if server.CANCELLAZIONE:
        esito = await server.elimina_evento(e["id"], oggetto_atteso="PROVA ACCENTI evento già così")
        idc = esito.get("id_nel_cestino")
        if idc:
            nel_cestino = await server.leggi_evento(idc)
            verifica("elimina_evento (evento nel cestino, descrizione)", nel_cestino.get("descrizione", ""))
            verifica("elimina_evento (evento nel cestino, oggetto)", nel_cestino.get("oggetto", ""), "PROVA ACCENTI evento già così")
            esiti.append(("elimina_evento (resta un evento, non un messaggio)", True, ""))
        else:
            esiti.append(("elimina_evento (evento ritrovato nel cestino)", False, "nessun elemento di calendario con quell'oggetto nel cestino: forse convertito in messaggio"))
    else:
        creati.append(("evento", e["id"]))
        esiti.append(("elimina_evento", True, "saltato: cancellazione_abilitata=false"))

    # 4b. serie ricorrente: settimanale, 3 volte, poi cancellazione dell'intera serie
    inizio_serie = (dt.datetime.now() + dt.timedelta(days=2)).replace(hour=16, minute=0, second=0, microsecond=0)
    sr = await server.crea_evento(oggetto="PROVA ACCENTI serie già così", inizio=inizio_serie.strftime("%Y-%m-%d %H:%M"),
                                  durata_minuti=15, descrizione=TESTO, etichetta="telefonata", promemoria_minuti=-1,
                                  ricorrenza="settimanale", occorrenze=3)
    esiti.append(("crea_evento (serie: ricorrente)", bool(sr.get("ricorrente")), str(sr)[:200]))
    lista = await server.elenca_eventi(da_data=inizio_serie.strftime("%Y-%m-%d"),
                                       a_data=(inizio_serie + dt.timedelta(days=30)).strftime("%Y-%m-%d"), massimo=300)
    occ = [x for x in lista if (x.get("oggetto") or "") == "PROVA ACCENTI serie già così"]
    esiti.append(("crea_evento (serie: 3 occorrenze in calendario)", len(occ) == 3, f"trovate {len(occ)} occorrenze"))
    if occ:
        verifica("crea_evento (serie: oggetto con accenti)", occ[0]["oggetto"], "PROVA ACCENTI serie già così")
    # 4c. singola occorrenza: modifica della seconda, cancellazione della terza, le altre restano
    if len(occ) == 3:
        occ_ordinate = sorted(occ, key=lambda x: x["inizio"])
        seconda, terza = occ_ordinate[1]["inizio"], occ_ordinate[2]["inizio"]
        try:
            m2 = await server.modifica_evento(sr["id"], luogo="Sala già prenotata", data_occorrenza=seconda)
            esiti.append(("modifica_evento (singola occorrenza)", m2.get("occorrenza_modificata") == seconda, str(m2)[:200]))
        except Exception as ex:  # noqa: BLE001
            esiti.append(("modifica_evento (singola occorrenza)", False, str(ex)[:200]))
        if server.CANCELLAZIONE:
            try:
                d3 = await server.elimina_evento(sr["id"], oggetto_atteso="PROVA ACCENTI serie già così", data_occorrenza=terza)
                esiti.append(("elimina_evento (singola occorrenza)", d3.get("occorrenza") == terza, str(d3)[:200]))
            except Exception as ex:  # noqa: BLE001
                esiti.append(("elimina_evento (singola occorrenza)", False, str(ex)[:200]))
            dopo1 = await server.elenca_eventi(da_data=inizio_serie.strftime("%Y-%m-%d"),
                                               a_data=(inizio_serie + dt.timedelta(days=30)).strftime("%Y-%m-%d"), massimo=300)
            rim1 = sorted([x for x in dopo1 if (x.get("oggetto") or "") == "PROVA ACCENTI serie già così"], key=lambda x: x["inizio"])
            esiti.append(("elimina_evento (singola occorrenza: restano 2 su 3)", len(rim1) == 2, f"rimaste {len(rim1)}"))
            luogo2 = rim1[1].get("luogo") if len(rim1) == 2 else None
            esiti.append(("modifica_evento (occorrenza modificata conservata, luogo con accento)", luogo2 == "Sala già prenotata", f"luogo={luogo2}"))
    else:
        esiti.append(("occorrenze singole", False, "saltate: la serie non aveva 3 occorrenze"))
    if server.CANCELLAZIONE:
        es = await server.elimina_evento(sr["id"], oggetto_atteso="PROVA ACCENTI serie già così", serie_intera=True)
        dopo = await server.elenca_eventi(da_data=inizio_serie.strftime("%Y-%m-%d"),
                                          a_data=(inizio_serie + dt.timedelta(days=30)).strftime("%Y-%m-%d"), massimo=300)
        rimaste = [x for x in dopo if (x.get("oggetto") or "") == "PROVA ACCENTI serie già così"]
        esiti.append(("elimina_evento (serie intera sparita dal calendario)", len(rimaste) == 0, f"rimaste {len(rimaste)}"))
    else:
        creati.append(("serie", sr["id"]))

    # 5. nota
    n = await server.crea_nota(testo="PROVA ACCENTI nota\n" + TESTO)
    creati.append(("nota", n["id"]))
    nt = await server.leggi_nota(n["id"])
    verifica("crea_nota", nt["testo"])

    # 6. contatto
    c = await server.crea_contatto(nome="Prova", cognome="Accenti già così", azienda="Società è più", note=TESTO,
                                   email="prova.accenti@example.com")
    creati.append(("contatto", c["id"]))
    ct = await server.leggi_contatto(c["id"])
    verifica("crea_contatto (note)", ct.get("note", ""))
    verifica("crea_contatto (azienda)", ct.get("azienda", ""), "Società è più")


def pulisci() -> None:
    if TIENI:
        print("Elementi di prova lasciati in Outlook (--tieni):", ", ".join(f"{d}" for d, _ in creati))
        print("File di prova lasciati in:", CARTELLA_PROVA)
        return
    if CARTELLA_PROVA.is_dir() and CARTELLA_PROVA.name.startswith("prova_accenti_"):
        shutil.rmtree(CARTELLA_PROVA)  # solo la cartella creata da questa prova
        print("Cartella dei file di prova cancellata:", CARTELLA_PROVA)
    def f():
        ns = server.ns()
        n = 0
        for desc, eid in creati:
            try:
                it = ns.GetItemFromID(eid)
                sogg = getattr(it, "Subject", "") or getattr(it, "FullName", "")
                if "PROVA ACCENTI" in str(sogg) or "Prova Accenti" in str(sogg):
                    it.Delete()  # va in Posta eliminata, recuperabile
                    n += 1
            except Exception as ex:  # noqa: BLE001
                print(f"  pulizia {desc}: non riuscita ({ex})")
        return n
    print(f"Elementi di prova spostati in Posta eliminata: {server.COM.run(f)}")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    finally:
        pulisci()
    ok = all(e[1] for e in esiti)
    for nome, esito, dettaglio in esiti:
        print(("OK   " if esito else "ERR  ") + nome + ("" if esito else "  -> " + dettaglio))
    print("ESITO:", "TUTTO OK" if ok else "CI SONO ERRORI")
    sys.exit(0 if ok else 1)
