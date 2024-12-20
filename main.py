from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import FileResponse
import pandas as pd
import requests
import os
import simplekml
import logging

# Logging aktivieren
logging.basicConfig(level=logging.INFO)

# FastAPI-App initialisieren
app = FastAPI()

# Google Maps API-Key (ersetzen mit deinem tatsächlichen Key)
API_KEY = os.getenv('GOOGLE_MAPS_API_KEY', 'AIzaSyBYUMStwyOUqAO609ooXqULkwLki9w-XRI')

@app.get("/")
async def root():
    """
    Root-Endpunkt für Health-Check.
    """
    return {"message": "Service läuft"}

@app.get("/favicon.ico")
async def favicon():
    """
    Dummy-Endpunkt für Favicon-Anfragen.
    """
    return {"message": "No favicon available"}

@app.post("/upload/")
async def upload_file(file: UploadFile):
    """
    Endpunkt für den Upload einer Excel-Datei und Verarbeitung der Adressen.
    """
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Invalid file format. Please upload an Excel file.")
    
    try:
        # Lese die Datei und bereinige die Spaltennamen
        street_list = pd.read_excel(file.file)
        street_list.columns = street_list.columns.str.strip()  # Entferne Leerzeichen aus Spaltennamen
        
        # Prüfe auf doppelte Spaltennamen
        if street_list.columns.duplicated().any():
            raise HTTPException(status_code=400, detail="The Excel file contains duplicate column names. Please check the file.")
        
        logging.info(f"Columns in the uploaded file: {list(street_list.columns)}")
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to process the file: {e}")
    
    # Erforderliche Spalten prüfen
    required_columns = ['Straße', 'HsNr', 'PLZ', 'Ort']
    for col in required_columns:
        if col not in street_list.columns:
            raise HTTPException(status_code=400, detail=f"The Excel file must contain the '{col}' column.")
    
    latitude_list = []
    longitude_list = []

    # Geokodierung der Adressen
    for idx, row in street_list.iterrows():
        full_address = f"{row['Straße']} {row['HsNr']}, {row['PLZ']} {row['Ort']}"
        url = f'https://maps.googleapis.com/maps/api/geocode/json?address={full_address}&key={API_KEY}'
        
        try:
            response = requests.get(url, timeout=20)
            response.raise_for_status()
            data = response.json()
            
            if data['status'] == 'OK':
                location = data['results'][0]['geometry']['location']
                latitude_list.append(location['lat'])
                longitude_list.append(location['lng'])
            else:
                latitude_list.append(None)
                longitude_list.append(None)
                logging.warning(f"Geocoding failed for address: {full_address} with status: {data['status']}")
        except Exception as e:
            latitude_list.append(None)
            longitude_list.append(None)
            logging.error(f"Error fetching coordinates for address {full_address}: {e}")
    
    street_list['Latitude'] = latitude_list
    street_list['Longitude'] = longitude_list

    # Filtere valide Koordinaten
    valid_coords = street_list.dropna(subset=['Latitude', 'Longitude'])
    if valid_coords.empty:
        raise HTTPException(status_code=500, detail="No valid addresses with coordinates found.")
    
    # Speichere als KML-Datei
    try:
        kml = simplekml.Kml()
        for _, row in valid_coords.iterrows():
            if pd.notna(row['Latitude']) and pd.notna(row['Longitude']):
                kml.newpoint(name=f"{row['Straße']} {row['HsNr']}", coords=[(row['Longitude'], row['Latitude'])])
        kml_output_path = "streets_map_with_house_numbers.kml"
        kml.save(kml_output_path)
        logging.info(f"KML file successfully created: {kml_output_path}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to create KML file: {e}")
    
    return FileResponse(kml_output_path, media_type="application/vnd.google-earth.kml+xml", filename=kml_output_path)


# Dynamischer Port für Railway (nur für lokale Ausführung notwendig)
if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8080))  # Nutze PORT-Umgebungsvariable oder 8080 als Standard
    uvicorn.run(app, host="0.0.0.0", port=port)
