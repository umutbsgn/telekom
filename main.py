from io import BytesIO
import pandas as pd
from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import requests
import os
import simplekml
import logging

# Logging aktivieren
logging.basicConfig(level=logging.INFO)

# FastAPI-Anwendung initialisieren
app = FastAPI()

# CORS-Middleware hinzufügen
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://id-preview--791b6969-6164-4810-878e-d0ca76188f6f.lovable.app"],  # Erlaubte Frontend-Domain
    allow_credentials=True,
    allow_methods=["*"],  # Erlaube alle HTTP-Methoden
    allow_headers=["*"],  # Erlaube alle Header
)

# Hardcodierte Google Maps API-Key
API_KEY = "AIzaSyBYUMStwyOUqAO609ooXqULkwLki9w-XRI"

@app.get("/")
async def root():
    """
    Root-Endpunkt, um zu prüfen, ob der Service läuft.
    """
    return {"message": "Service läuft"}

@app.post("/upload/")
async def upload_file(file: UploadFile):
    """
    Endpunkt zum Hochladen und Verarbeiten einer Excel-Datei.
    """
    # Überprüfe das Dateiformat
    if not file.filename.endswith(".xlsx"):
        logging.error(f"Invalid file format: {file.filename}")
        return JSONResponse(
            status_code=400,
            content={"status": "error", "message": "Invalid file format. Please upload an Excel file."}
        )

    try:
        # Datei einlesen
        file_content = await file.read()
        excel_data = pd.read_excel(BytesIO(file_content))
        excel_data.columns = excel_data.columns.str.strip()  # Entferne Leerzeichen aus Spaltennamen
        
        # Überprüfe auf doppelte Spaltennamen
        if excel_data.columns.duplicated().any():
            logging.error(f"Duplicate columns in file: {file.filename}")
            return JSONResponse(
                status_code=400,
                content={"status": "error", "message": "The Excel file contains duplicate column names. Please check the file."}
            )
        logging.info(f"Columns in the uploaded file: {list(excel_data.columns)}")
        
    except Exception as e:
        logging.error(f"Error processing file {file.filename}: {e}")
        return JSONResponse(
            status_code=500,
            content={"status": "error", "message": f"Failed to process the file: {e}"}
        )

    # Überprüfe, ob die erforderlichen Spalten vorhanden sind
    required_columns = ['Straße', 'HsNr', 'PLZ', 'Ort']
    for col in required_columns:
        if col not in excel_data.columns:
            logging.error(f"Missing column: {col} in file: {file.filename}")
            return JSONResponse(
                status_code=400,
                content={"status": "error", "message": f"The Excel file must contain the '{col}' column."}
            )
    
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
        logging.error(f"No valid coordinates found in file: {file.filename}")
        return JSONResponse(
            status_code=500,
            content={"status": "error", "message": "No valid addresses with coordinates found."}
        )
    
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
        logging.error(f"Error creating KML file: {e}")
        return JSONResponse(
            status_code=500,
            content={"status": "error", "message": f"Failed to create KML file: {e}"}
        )
    
    # JSON-Antwort mit Erfolgsmeldung
    return {
        "message": "KML file successfully created.",
        "file_url": kml_output_path
    }

@app.get("/upload/")
async def upload_info():
    """
    GET-Endpunkt für /upload/, um benutzerfreundliche Informationen zu liefern.
    """
    return {"message": "This endpoint only accepts POST requests for file uploads. Please use POST to upload a valid Excel file."}
