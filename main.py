import os
import requests
import pandas as pd
from tkinter import Tk, Button, Label, filedialog, Toplevel, Text, Checkbutton, BooleanVar
from tkinter import ttk
import threading
import zipfile


# Funktion zum Überprüfen, ob eine URL gültig ist
def is_valid_url(url):
    return isinstance(url, str) and (url.startswith("http://") or url.startswith("https://"))
    
# Funktion zum Herunterladen eines Bildes
def download_image(url, destination):
    try:
        response = requests.get(url, stream=True)
        if response.status_code == 200:
            with open(destination, 'wb') as out_file:
                out_file.write(response.content)
            return True
        else:
            print(f"Fehler beim Herunterladen: {url} (Statuscode: {response.status_code})")
            return False
    except Exception as e:
        print(f"Fehler beim Herunterladen: {url} ({e})")
        return False

# Funktion zum Zippen der Bilder in Batches
def zip_files_in_batches(directory, excel_filename, batch_size=1000):
    files = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
    total_files = len(files)
    batch_number = 1

    for i in range(0, total_files, batch_size):
        batch_files = files[i:i + batch_size]
        zip_filename = os.path.join(directory, f"{excel_filename}_batch_{batch_number}.zip")
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for file in batch_files:
                zipf.write(os.path.join(directory, file), file)
        batch_number += 1

    # Lösche die originalen Bilder
    for file in files:
        os.remove(os.path.join(directory, file))

# Fenster für fehlgeschlagene Downloads anzeigen
def show_failed_downloads(failed_downloads):
    if failed_downloads:
        top = Toplevel(root)
        top.title("Fehlgeschlagene Downloads")
        top.geometry("400x300")
        
        text = Text(top)
        text.pack(expand=True, fill='both')
        
        for item in failed_downloads:
            text.insert('end', f"{item}\n")

# Excel-Datei laden
def load_excel_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if filepath:
        global df, excel_filename, excel_folder
        df = pd.read_excel(filepath)
        excel_filename = os.path.splitext(os.path.basename(filepath))[0]  # Dateiname ohne Erweiterung
        excel_folder = os.path.dirname(filepath)  # Ordner der Excel-Datei
        label.config(text=f"Geladene Datei: {os.path.basename(filepath)}")
        start_button.config(state="normal")

# Ausgabeordner wählen
def choose_output_directory():
    global output_dir
    output_dir = filedialog.askdirectory()
    if output_dir:
        label_output.config(text=f"Ausgabeordner: {output_dir}")

# Bilder verarbeiten
def process_images():
    total_images = sum([1 for index, row in df.iterrows() for column in df.columns[3:]
                        if pd.notna(row[column]) and (is_valid_url(row[column]) or use_local_images.get())])
    progress_bar['maximum'] = total_images
    processed_images = 0
    failed_downloads = []

    for index, row in df.iterrows():
        asin = row['ASIN']
        for column in df.columns[3:]:
            image_value = row[column]

            if pd.notna(image_value):
                image_name = f"{asin}.{column}.jpg"
                destination = os.path.join(output_dir, image_name)

                if use_local_images.get():
                    # Bild lokal kopieren
                    local_image_path = os.path.join(excel_folder, image_value)
                    if os.path.isfile(local_image_path):
                        try:
                            with open(local_image_path, 'rb') as fsrc, open(destination, 'wb') as fdst:
                                fdst.write(fsrc.read())
                            processed_images += 1
                        except Exception as e:
                            failed_downloads.append(f"ASIN: {asin}, Datei: {image_value}, Fehler: {e}")
                    else:
                        failed_downloads.append(f"ASIN: {asin}, Datei nicht gefunden: {image_value}")
                else:
                    # Bild per URL laden
                    if is_valid_url(image_value):
                        if download_image(image_value, destination):
                            processed_images += 1
                        else:
                            failed_downloads.append(f"ASIN: {asin}, URL: {image_value}")

                progress_bar['value'] = processed_images
                label_status.config(text=f"Verarbeite Bild {processed_images} von {total_images}")
                root.update_idletasks()

    zip_files_in_batches(output_dir, excel_filename)
    label_status.config(text="Verarbeitung abgeschlossen!")
    show_failed_downloads(failed_downloads)

# Verarbeitung starten (in separatem Thread)
def start_processing():
    threading.Thread(target=process_images).start()

# === GUI ===
root = Tk()
root.title("Excel Bildverarbeitung")

use_local_images = BooleanVar(value=False)  # Variable für lokale Bildverwendung

load_button = Button(root, text="Excel Datei laden", command=load_excel_file)
load_button.pack(pady=10)

label = Label(root, text="Keine Datei geladen")
label.pack(pady=5)

choose_dir_button = Button(root, text="Ausgabeordner wählen", command=choose_output_directory)
choose_dir_button.pack(pady=10)

label_output = Label(root, text="Kein Ausgabeordner gewählt")
label_output.pack(pady=5)

# Checkbutton für lokale Bilder
local_image_checkbox = Checkbutton(root, text="Bilder lokal laden (statt per URL)", variable=use_local_images)
local_image_checkbox.pack(pady=5)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=10)

start_button = Button(root, text="Start", state="disabled", command=start_processing)
start_button.pack(pady=20)

label_status = Label(root, text="")
label_status.pack(pady=10)

root.mainloop()
