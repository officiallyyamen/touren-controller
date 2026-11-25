import csv
import json
import random
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

random.seed(11)
ORT = Path(__file__).parent.parent
ZIEL = ORT / "daten"
for teil in ("standorte", "wiegescheine"):
    (ZIEL / teil).mkdir(parents=True, exist_ok=True)

MONATE = ["Jän", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt"]
STANDORTE = {"WIE": "Wien", "GRZ": "Graz", "LIN": "Linz", "SBG": "Salzburg", "VIL": "Villach", "STP": "St. Pölten"}
ABFALL = {"Restmüll": 118, "Bioabfall": 62, "Altpapier": -35, "Kunststoff": 48, "Glas": -12, "Gewerbeabfall": 132}
KUNDEN = ["Marktgemeinde", "Stadtgemeinde", "Wohnbau", "Handelskette", "Klinikum", "Industriepark", "Gemeindeverband", "Hotelgruppe"]
NAMEN = ["Auwald", "Birkenfeld", "Donautal", "Eschenbach", "Föhrenau", "Grünberg", "Hainbach", "Lindenhof"]
STUNDENSATZ = 58.0
DIESEL = 0.52

touren = {}
for kurz in STANDORTE:
    for nr in range(1, 9):
        art = list(ABFALL)[(nr + len(kurz) + ord(kurz[0])) % 6]
        kapazitaet = random.choice([8, 10, 12])
        touren[f"{kurz}-{nr:02d}"] = {
            "standort": kurz,
            "kunde": f"{KUNDEN[(nr + ord(kurz[1])) % 8]} {NAMEN[(nr * 3 + ord(kurz[2])) % 8]}",
            "art": art,
            "kap": kapazitaet,
            "fuellung": random.uniform(0.68, 0.9),
            "km": random.randint(55, 150),
            "h": random.uniform(5.2, 8.4),
            "wochentag": nr % 5,
        }
touren["GRZ-03"].update(fuellung=0.37, km=142, h=8.1)
touren["VIL-06"].update(fuellung=0.44, km=151)
touren["STP-02"].update(km=138, h=8.3)
for t in touren.values():
    kosten = t["km"] * DIESEL + t["h"] * STUNDENSATZ + t["kap"] * t["fuellung"] * ABFALL[t["art"]]
    t["preis"] = round(max(kosten, 260) * random.uniform(1.14, 1.34) / 5) * 5
for tour, faktor in (("GRZ-03", 0.86), ("VIL-06", 0.93), ("STP-02", 0.95), ("LIN-05", 1.03)):
    t = touren[tour]
    kosten = t["km"] * DIESEL + t["h"] * STUNDENSATZ + t["kap"] * t["fuellung"] * ABFALL[t["art"]]
    t["preis"] = round(max(kosten, 260) * faktor / 5) * 5

fahrten, scheine, schein_nr = [], [], 70000
for tour, t in touren.items():
    tag = date(2026, 1, 5) + timedelta(days=t["wochentag"])
    while tag <= date(2026, 10, 30):
        if random.random() > 0.015:
            teuer = 1.24 if t["standort"] == "LIN" and tag >= date(2026, 7, 1) else 1.0
            km = round(t["km"] * random.uniform(0.95, 1.07))
            h = round(t["h"] * random.uniform(0.94, 1.1) * 4) / 4
            diesel = round(km * DIESEL * teuer * random.uniform(0.96, 1.05), 2)
            tonnen = round(min(t["kap"] * 1.04, t["kap"] * t["fuellung"] * random.uniform(0.88, 1.12)), 2)
            fahrten.append((tour, tag, f"{t['standort'][0]}-{700 + int(tour[-2:])}", km, h, diesel))
            schein_nr += 1
            scheine.append([f"W{schein_nr}", tag, tour, t["art"], int(tonnen * 1000)])
        tag += timedelta(days=7)

abgabe_fehlt = [f for f in fahrten if f[0].startswith("VIL") and f[1].month == 9]
geliefert = [f for f in fahrten if f not in abgabe_fehlt]


def uhr(h):
    return f"{int(h)}:{round(h % 1 * 60):02d}"


def de(z, stellen=2):
    return f"{z:,.{stellen}f}".replace(",", " ").replace(".", ",").replace(" ", ".")


FORM = {
    "WIE": (["Datum", "Tour", "LKW", "km", "Stunden", "Diesel €"], lambda f: [f[1], f[0], f[2], f[3], f[4], f[5]], 0),
    "GRZ": (["Tag", "Tourennr.", "Fahrzeug", "Kilometer", "Arbeitszeit", "Dieselkosten"], lambda f: [f[1], f"grz {int(f[0][-2:])}", f[2], f[3], uhr(f[4]), f[5]], 2),
    "LIN": (["Einsatzdatum", "Tour-ID", "Kennzeichen", "gefahrene km", "Std.", "Treibstoff EUR"], lambda f: [f[1].strftime("%d.%m.%Y"), f[0], f[2], f[3], de(f[4]), de(f[5])], 0),
    "SBG": (["Tour", "Datum", "Wagen", "Strecke", "Dauer", "Diesel"], lambda f: [int(f[0][-2:]), f[1], f[2], f[3], f[4], f[5]], 3),
    "VIL": (["DATUM", "TOUR", "LKW-NR", "KM", "STUNDEN", "TANKKOSTEN"], lambda f: [f[1], f"V{f[0][-2:]}", f[2], f"{f[3]} km", f[4], f"€ {de(f[5])}"], 1),
    "STP": (["Datum", "Tour Nr", "Auto", "Km-Stand Differenz", "Arbeitsstunden", "Diesel brutto"], lambda f: [f[1], f[0].replace("-", "/"), f[2], f[3], f[4], round(f[5] * 1.2, 2)], 0),
}
for kurz, (kopf, zeile, vorspann) in FORM.items():
    wb = Workbook()
    wb.remove(wb.active)
    for m, name in enumerate(MONATE, 1):
        teil = [f for f in geliefert if f[0].startswith(kurz) and f[1].month == m]
        if not teil:
            continue
        ws = wb.create_sheet(f"{name} 2026")
        for i in range(vorspann):
            ws.append([["Grünkreis Entsorgung", f"Standort {STANDORTE[kurz]}", ""][i]])
        ws.append(kopf)
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
        for f in sorted(teil, key=lambda f: f[1]):
            ws.append(zeile(f))
        if kurz in ("WIE", "SBG"):
            ws.append(["Summe" if kurz == "WIE" else None, None, None, sum(f[3] for f in teil), None, None])
        for s in "ABCDEF":
            ws.column_dimensions[s].width = 18
    wb.save(ZIEL / "standorte" / f"Touren_{STANDORTE[kurz].replace(' ', '_').replace('.', '')}.xlsx")

doppelt = random.sample(scheine, 17)
ohne = set(s[0] for s in random.sample(scheine, 26))
alle = sorted(scheine + [list(s) for s in doppelt], key=lambda s: s[1])
for q in range(1, 5):
    teil = [s for s in alle if (s[1].month - 1) // 3 + 1 == q]
    if teil:
        with open(ZIEL / "wiegescheine" / f"waage_2026_Q{q}.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Schein", "Datum", "Tour", "Abfallart", "Netto_kg"])
            for s in teil:
                w.writerow([s[0], s[1].strftime("%d.%m.%Y"), s[2], s[3], "" if s[0] in ohne else s[4]])

with open(ZIEL / "vertraege.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow(["Tour", "Kunde", "Abfallart", "Kapazitaet_t", "Preis_Abfuhr", "Gueltig_ab"])
    for tour, t in touren.items():
        w.writerow([tour, t["kunde"], t["art"], t["kap"], de(t["preis"]), "01.01.2026"])
        if not tour.startswith("STP") and tour != "GRZ-03":
            t["neu"] = round(t["preis"] * 1.03, 2)
            w.writerow([tour, t["kunde"], t["art"], t["kap"], de(t["neu"]), "01.07.2026"])
with open(ZIEL / "abfallarten.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow(["Abfallart", "Entsorgung_je_t"])
    w.writerows([a, de(p)] for a, p in ABFALL.items())
with open(ZIEL / "standorte.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow(["Kuerzel", "Standort", "Datei", "Kopfzeile", "Diesel_brutto"])
    for k, n in STANDORTE.items():
        w.writerow([k, n, f"Touren_{n.replace(' ', '_').replace('.', '')}.xlsx", FORM[k][2], int(k == "STP")])
with open(ZIEL / "spalten.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow(["Kuerzel", "Alt", "Neu"])
    for k, (kopf, _, _) in FORM.items():
        ordnung = ["Tour", "Datum", "Fahrzeug", "km", "Stunden", "Diesel"] if k == "SBG" else ["Datum", "Tour", "Fahrzeug", "km", "Stunden", "Diesel"]
        w.writerows([k, a, n] for a, n in zip(kopf, ordnung))


def preis(tour, tag):
    t = touren[tour]
    return t.get("neu", t["preis"]) if tag >= date(2026, 7, 1) else t["preis"]


gewicht = {(s[2], s[1]): (0 if s[0] in ohne else s[4] / 1000) for s in scheine}
summe = defaultdict(lambda: defaultdict(float))
for tour, tag, _, km, h, diesel in geliefert:
    for k in (("alle",), ("ort", tour[:3]), ("tour", tour), ("monat", tour[:3], tag.month)):
        z = summe[k]
        z["fahrten"] += 1
        z["erloes"] += preis(tour, tag)
        z["diesel"] += diesel
        z["personal"] += h * STUNDENSATZ
        z["entsorgung"] += gewicht[(tour, tag)] * ABFALL[touren[tour]["art"]]
        z["tonnen"] += gewicht[(tour, tag)]
for z in summe.values():
    z["db"] = z["erloes"] - z["diesel"] - z["personal"] - z["entsorgung"]

wb = Workbook()
ws = wb.active
ws.title = "Plan 2026"
ws.append(["Standort", "Kennzahl", *MONATE])
for k, n in STANDORTE.items():
    for kenn, feld, f in (("Erlös", "erloes", 1.0), ("Kosten", None, {"WIE": 1.012, "GRZ": 0.955, "LIN": 0.965, "SBG": 1.004, "VIL": 0.985, "STP": 1.006}[k])):
        werte = []
        for m in range(1, 11):
            z = summe.get(("monat", k, m)) or summe[("monat", k, 8)]
            basis = z["erloes"] if feld else z["erloes"] - z["db"]
            werte.append(round(basis * f * random.uniform(0.985, 1.012) / 100) * 100)
        ws.append([n, kenn, *werte])
wb.save(ZIEL / "plan_2026.xlsx")

ergebnis = {"gesamt": summe[("alle",)], "orte": {k: summe[("ort", k)] for k in STANDORTE}, "verlust": {t: round(summe[("tour", t)]["db"]) for t in touren if summe[("tour", t)]["db"] < 0}, "doppelte_scheine": 17, "scheine_ohne_gewicht": 26, "fahrten_ohne_gewicht": sum(1 for f in geliefert if gewicht[(f[0], f[1])] == 0)}
json.dump(ergebnis, open(ORT / "werkzeug" / "soll.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(ergebnis, ensure_ascii=False, indent=1))
