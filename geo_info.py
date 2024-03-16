import pandas as pd
import requests
import os

# Funktion zum Überprüfen, ob eine IP bereits in der Ausgabedatei vorhanden ist
def is_ip_already_processed(ip, processed_ips):
    return ip in processed_ips

# Funktion zum Abrufen von Geolokationsdaten für eine IP von ipinfo.io
def get_geo_info(ip):
    try:
        response = requests.get(f"https://ipinfo.io/{ip}/json")
        data = response.json()
        return data
    except Exception as e:
        print(f"Fehler beim Abrufen der Informationen für die IP {ip}: {e}")
        return None

# Lese die IP-Adressen aus der Excel-Datei
try:
    df = pd.read_excel('overview.xlsx', sheet_name='IPs')
except FileNotFoundError:
    print("Die Datei 'overview.xlsx' konnte nicht gefunden werden.")
    exit()

# Überprüfe, ob die Ausgabedatei existiert und Daten enthält
if os.path.exists('geo_info.xlsx'):
    try:
        existing_df = pd.read_excel('geo_info.xlsx')
        processed_ips = set(existing_df['IP'])
    except pd.errors.EmptyDataError:
        existing_df = pd.DataFrame()
        processed_ips = set()
        print("Die Datei 'geo_info.xlsx' ist leer.")
    except KeyError:
        print("Die Spalte 'IP' wurde in 'geo_info.xlsx' nicht gefunden.")
        exit()
else:
    existing_df = pd.DataFrame()
    processed_ips = set()

# Liste zum Speichern der Geolokationsdaten
geo_data = []

# Durchlaufen der IP-Adressen in der Excel-Datei
total_ips = len(df) - len(existing_df)
processed_count = 0

for ip in df['IP']:
    if is_ip_already_processed(ip, processed_ips):
        print(f"IP {ip} wurde bereits verarbeitet, überspringe...")
        continue

    data = get_geo_info(ip)
    if data is not None:
        geo_data.append({'IP': ip, 
                         'Hostname': data.get('hostname', 'N/A'), 
                         'Land': data.get('country', 'N/A'), 
                         'Organisation': data.get('org', 'N/A'),
                         'Stadt': data.get('city', 'N/A'),
                         'Region': data.get('region', 'N/A'),
                         'Lokation': data.get('loc', 'N/A'),
                         'Postleitzahl': data.get('postal', 'N/A'),
                         'Zeitzone': data.get('timezone', 'N/A'),
                         'ASN': data.get('asn', 'N/A')})
        processed_ips.add(ip)  # Hinzufügen der IP-Adresse zu den verarbeiteten IPs
    else:
        geo_data.append({'IP': ip, 
                         'Hostname': 'N/A', 
                         'Land': 'N/A', 
                         'Organisation': 'N/A',
                         'Stadt': 'N/A',
                         'Region': 'N/A',
                         'Lokation': 'N/A',
                         'Postleitzahl': 'N/A',
                         'Zeitzone': 'N/A',
                         'ASN': 'N/A'})

    processed_count += 1
    print(f"Fortschritt: {processed_count}/{total_ips} IPs verarbeitet", end='\r')

# Erstelle ein DataFrame mit den neuen Geolokationsdaten
geo_df = pd.DataFrame(geo_data)

# Hänge die neuen Daten an die vorhandenen Daten an und entferne Duplikate basierend auf der IP-Spalte
updated_df = pd.concat([existing_df, geo_df]).drop_duplicates(subset=['IP'], keep='last')

# Schreibe die aktualisierten Geolokationsdaten in die Excel-Datei
updated_df.to_excel('geo_info.xlsx', index=False)

print("\nGeolokationsdaten wurden erfolgreich in 'geo_info.xlsx' geschrieben.")
input("Aufgabe abgeschlossen, drücke Enter um das Programm zu beenden.")
