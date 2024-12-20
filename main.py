from io import BytesIO
import pandas as pd
from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import FileResponse
import requests
import os
import simplekml
import logging

# Logging aktivieren
logging.basicConfig(level=logging.INFO)

app = FastAPI()

# Google Maps API-Key
API_KEY = os.getenv('GOOGLE_MAPS_API_KEY', 'AIzaSyBYUMStwyOUqAO609ooXqULkwLki9w-XRI')

@app.post("/upload/")
async def upload_file(file: UploadFile):
    """
    Endpunkt zum Hochladen und Verarbeiten einer Excel-Datei.
    """
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Invalid file format. Please upload an Excel file.")

    try:
        # Konvertiere die Datei in einen BytesIO-Stream für pandas
        file_content = await file.read()
        excel_data = pd.read_excel(BytesIO(file_content))
        excel_data.columns = excel_data.columns.str.strip()  # Entferne Leerzeichen aus den Spaltennamen
        
        # Prüfe auf doppelte Spaltennamen
        if excel_data.columns.duplicated().any():
            raise HTTPException(status_code=400, detail="The Excel file contains duplicate column names. Please check the file.")
        
        logging.info(f"Columns in the uploaded file: {list(excel_data.columns)}")
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to process the file: {e}")

    # Überprüfe, ob die erforderlichen Spalten vorhanden sind
    required_columns = ['Straße', 'HsNr', 'PLZ', 'Ort']
    for col in required_columns:
        if col not in excel_data.columns:
            raise HTTPException(status_code=400, detail=f"The Excel file must contain the '{col}' column.")
    
    latitude_list = []
    longitude_list = []

    # Geokodierung der Adressen
    for idx, row in excel_data.iterrows():
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
    
    excel_data['Latitude'] = latitude_list
    excel_data['Longitude'] = longitude_list

    # Filtere valide Koordinaten
    valid_coords = excel_data.dropna(subset=['Latitude', 'Longitude'])
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
