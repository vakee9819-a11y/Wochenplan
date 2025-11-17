import pandas as pd
import json
from datetime import time, datetime, timedelta
import pytz

# Parameter für x Tage im Voraus / -x Tage zurück (heutiger Tag = -1, generiert den gestrigen auch mit)
while True:
    try:
        x = int(input("Geben Sie die Anzahl der Tage ein (positiv für zukünftige Tage, negativ für vergangene Tage, -1 für gestern + heute): "))
        break  # gültige Zahl -> Loop verlassen
    except ValueError:
        print("Ungültige Eingabe. Bitte geben Sie eine ganze Zahl ein.")

startdatum = datetime.today() + timedelta(days=1 if x > 0 else 0)
anzahl_tage = abs(x)

df = pd.read_excel("Stammdaten.xlsx", sheet_name="Aufgaben", engine="openpyxl")
df2 = pd.read_excel("Stammdaten.xlsx", sheet_name="Schichten", engine="openpyxl")

aktive_aufgaben = df[df["Aktiv"] == "Ja"]

def get_tasks_for_day_and_shift(day, shift):
    tage_map = {
        "Monday": ["Monday", "Weekdays", "Daily"],
        "Tuesday": ["Tuesday", "Weekdays", "Daily"],
        "Wednesday": ["Wed-Fri", "Wednesday", "Weekdays", "Daily"],
        "Thursday": ["Wed-Fri", "Thursday", "Weekdays", "Daily"],
        "Friday": ["Wed-Fri", "Fri-Son", "Friday", "Weekdays", "Daily"],
        "Saturday": ["Fri-Son", "Weekend", "Daily"],
        "Sunday": ["Fri-Son", "Weekend", "Daily"]
    }
    return aktive_aufgaben[
        (aktive_aufgaben["Periode"].isin(tage_map.get(day, []))) &
        (aktive_aufgaben["Schicht"] == shift)
    ]

cet = pytz.timezone("Europe/Zurich")
utc = pytz.utc
for i in range(anzahl_tage + (1 if x < 0 else 0)):
    offset = i if x > 0 else -i
    aktuelles_datum = startdatum + timedelta(days=offset)
    aktueller_wochentag = aktuelles_datum.strftime("%A")
    aktuelle_kw = aktuelles_datum.isocalendar().week
    datum_cet = cet.localize(datetime.combine(aktuelles_datum.date(), time.min))
    datum_iso = datum_cet.isoformat()
    due_date = (datetime.now() - timedelta(hours=2)).strftime("%Y-%m-%dT%H:%M:%SZ")

    aufgaben_liste = []
    for schicht in ["Frühdienst 1", "Frühdienst 2", "Spätdienst", "Pikettdienst"]:
        aufgaben = get_tasks_for_day_and_shift(aktueller_wochentag, schicht).copy()
        aufgaben["Start"] = aufgaben["Start"].apply(lambda x: x.strftime("%H:%M") if isinstance(x, time) else str(x).strip())
        aufgaben["Ende"] = aufgaben["Ende"].apply(lambda x: x.strftime("%H:%M") if isinstance(x, time) else str(x).strip())
        aufgaben["Eskalation"] = aufgaben["Eskalation"].apply(lambda x: x.strftime("%H:%M") if isinstance(x, time) else str(x).strip())
        aufgaben["Wiki Link"] = aufgaben["Wiki Link"].fillna("")

        for _, row in aufgaben.iterrows():
            start_time = datetime.strptime(row["Start"], "%H:%M").time()
            end_time = datetime.strptime(row["Ende"], "%H:%M").time()
            esc_time = datetime.strptime(row["Eskalation"], "%H:%M").time()
            start_dt = cet.localize(datetime.combine(aktuelles_datum.date(), start_time)).astimezone(utc)
            end_dt = cet.localize(datetime.combine(aktuelles_datum.date(), end_time)).astimezone(utc)
            esc_dt = cet.localize(datetime.combine(aktuelles_datum.date(), esc_time)).astimezone(utc)

            task = {
                "taskDate": datum_iso,
                "taskStart": row["Start"],
                "taskStartTS": start_dt.isoformat(),
                "taskEnd": row["Ende"],
                "taskEndTS": end_dt.isoformat(),
                "taskEscalationTS": esc_dt.isoformat(),
                "taskEscalationInterval": "PT5M",
                "taskName": row["Aufgabenname"],
                "taskDescription": row["Aufgabenbeschreibung"],
                "taskWiki": row["Wiki Link"],
                "taskResponsible": schicht
            }
            aufgaben_liste.append(task)

    ein_tages_json = {
        "variables": {
            "_case.name": f"Wochenplan KW{aktuelle_kw} erzeugen",
            "_case.dueDate": due_date,
            "tasks": aufgaben_liste
        }
    }

    dateiname = f"aufgaben_{aktueller_wochentag}_KW{aktuelle_kw}.json"
    with open(dateiname, "w", encoding="utf-8") as f:
        json.dump(ein_tages_json, f, ensure_ascii=False, indent=4)
    print(f"Datei für {aktueller_wochentag}, {aktuelles_datum.date()} wurde erstellt.")