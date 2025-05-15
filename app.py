import os
import uuid
import time
from datetime import datetime
from flask import Flask, request, render_template, send_file, redirect
import openpyxl
import requests
from io import BytesIO

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def sprawdz_nip_w_vat(nip):
    nip = str(nip).replace('-', '').strip()
    if not nip.isdigit() or len(nip) != 10:
        return nip, "Nieprawidłowy NIP", "Błąd"

    base_url = "https://wl-api.mf.gov.pl/api/search/nip/"
    today = datetime.today().strftime('%Y-%m-%d')
    request_id = str(uuid.uuid4())

    url = f"{base_url}{nip}?date={today}"
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'RequestId': request_id
    }

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            result = response.json()
            subject = result['result']['subject']
            if subject:
                nazwa = subject.get('name', 'Brak danych')
                status_vat = subject.get('statusVat', 'Nieznany')
                return nip, nazwa, status_vat
            else:
                return nip, "Nie znaleziono w rejestrze", "Brak"
        else:
            return nip, "Błąd odpowiedzi", f"Kod {response.status_code}"
    except Exception as e:
        return nip, "Błąd zapytania", str(e)

def wczytaj_nipy_z_excel(plik):
    wb = openpyxl.load_workbook(plik)
    ws = wb.active
    nipy = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        nip = row[0]
        if nip:
            nipy.append(str(nip).strip())
    return nipy

def zapisz_do_excel(wyniki):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Wyniki"
    ws.append(["NIP", "Nazwa podmiotu", "Status VAT"])
    for wiersz in wyniki:
        ws.append(wiersz)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "plik" not in request.files:
            return "Nie przesłano pliku", 400

        plik = request.files["plik"]
        if plik.filename == "":
            return "Nie wybrano pliku", 400

        nipy = wczytaj_nipy_z_excel(plik)
        wyniki = []
        for nip in nipy:
            wynik = sprawdz_nip_w_vat(nip)
            wyniki.append(wynik)
            time.sleep(0.3)  # przerwa dla API

        wynikowy_plik = zapisz_do_excel(wyniki)
        return send_file(wynikowy_plik, download_name="wyniki_nip.xlsx", as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)