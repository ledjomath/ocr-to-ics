import pandas as pd
from doctr.io import DocumentFile
from doctr.models import ocr_predictor
from datetime import datetime, timedelta
from ics import Calendar, Event
import re

ocr_model = ocr_predictor(det_arch="db_resnet50", reco_arch="crnn_vgg16_bn", pretrained=True)

def extract_text_blocks(pdf_path):
    doc = DocumentFile.from_pdf(pdf_path)
    result = ocr_model(doc)
    blocks = []
    for page_num, page in enumerate(result.pages):
        for block in page.blocks:
            for line in block.lines:
                text = " ".join([w.value for w in line.words])
                x0 = min(w.geometry[0][0] for w in line.words)
                y0 = min(w.geometry[0][1] for w in line.words)
                blocks.append({"page": page_num + 1, "x": x0, "y": y0, "text": text})
    return blocks

def extract_events(blocks, month="juillet", year=2025):
    mois_num = {
        "janvier": 1, "février": 2, "mars": 3, "avril": 4,
        "mai": 5, "juin": 6, "juillet": 7, "août": 8,
        "septembre": 9, "octobre": 10, "novembre": 11, "décembre": 12
    }[month]

    from collections import defaultdict
    grouped_by_y = defaultdict(list)
    for b in blocks:
        rounded_y = round(b["y"], 1)
        grouped_by_y[rounded_y].append((b["x"], b["text"]))

    lignes = []
    for y in sorted(grouped_by_y):
        ligne = " ".join(text for x, text in sorted(grouped_by_y[y]))
        lignes.append(ligne)

    events = []
    current_date = None
    for line in lignes:
        line = line.strip()
        if not line:
            continue

        date_match = re.search(r"(lun|mar|mer|jeu|ven|sam|dim)\.\s*(\d{1,2})", line.lower())
        if date_match:
            try:
                day = int(date_match.group(2))
                current_date = datetime(year, mois_num, day)
            except:
                continue
        elif any(h in line for h in [":", "-"]) and current_date:
            h_match = re.search(r"(\d{1,2}:\d{2})(?:-(\d{1,2}:\d{2}))?", line)
            if h_match:
                h_debut = h_match.group(1)
                h_fin = h_match.group(2) or None
                dt_start = datetime.strptime(f"{current_date.date()} {h_debut}", "%Y-%m-%d %H:%M")
                dt_end = datetime.strptime(f"{current_date.date()} {h_fin}", "%Y-%m-%d %H:%M") if h_fin else dt_start + timedelta(hours=3)
                titre = line.split(h_match.group(0))[-1].strip()
                desc = ""
                if "(" in titre:
                    titre, desc = titre.split("(", 1)
                    desc = desc.replace(")", "").strip()
                events.append({
                    "Date": current_date.strftime("%Y-%m-%d"),
                    "Début": dt_start.strftime("%H:%M"),
                    "Fin": dt_end.strftime("%H:%M"),
                    "Titre": titre.strip(),
                    "Description": desc
                })
    return pd.DataFrame(events)

def generate_ics(df, filename="planning_juillet25.ics"):
    cal = Calendar()
    for row in df.itertuples():
        e = Event()
        e.name = row.Titre
        dt_start = datetime.strptime(f"{row.Date} {row.Début}", "%Y-%m-%d %H:%M")
        dt_end = datetime.strptime(f"{row.Date} {row.Fin}", "%Y-%m-%d %H:%M")
        if dt_end <= dt_start:
            dt_end += timedelta(days=1)
        e.begin = dt_start
        e.end = dt_end
        e.description = row.Description
        cal.events.add(e)
    with open(filename, "w", encoding="utf-8") as f:
        f.writelines(cal)

def process_pdf(pdf_path):
    blocks = extract_text_blocks(pdf_path)
    df = extract_events(blocks)
    df.to_excel("planning_juillet25.xlsx", index=False)
    generate_ics(df, "planning_juillet25.ics")
    return "planning_juillet25.xlsx", "planning_juillet25.ics"