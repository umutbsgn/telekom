from io import BytesIO
import pandas as pd
from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import requests
import os
import simplekml
import logging
import traceback

# Logging aktivieren
logging.basicConfig(level=logging.INFO)

# FastAPI-Anwendung initialisieren
app = FastAPI()

# CORS-Middleware hinzufügen
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Google Maps API-Key
API_KEY = "AIzaSyBYUMStwyOUqAO609ooXqULkwLki9w-XRI"

@app.get("/")
async def root():
    """
    Root-Endpunkt, um zu prüfen, ob der Service läuft.
    """
    logging.info("Root endpoint called. Service is running.")
    return {"message": "Service läuft"}

@app.post("/upload/")
async def upload_file(file: UploadFile):
    """
    Endpunkt zum Hochladen und Verarbeiten einer Excel-Datei.
    Erwartete Spalten: 'Straße', 'HsNr', 'PLZ', 'Ort', 'Haushalte'
    """
    # Überprüfe das Dateiformat
    if not file.filename.endswith(".xlsx"):
        logging.error(f"Invalid file format: {file.filename}")
        return JSONResponse(
            status_code=400,
            content={"status": "error", "message": "Invalid file format. Please upload an Excel file."}
        )

    try:
        # Datei einlesen und prüfen
        file_content = await file.read()
        if not file_content:
            logging.error("Uploaded file is empty.")
            return JSONResponse(
                status_code=400,
                content={"status": "error", "message": "Uploaded file is empty. Please provide a valid Excel file."}
            )

        excel_data = pd.read_excel(BytesIO(file_content))
        excel_data.columns = excel_data.columns.str.strip()  # Leerzeichen aus Spaltennamen entfernen
        
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
        logging.debug(traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={"status": "error", "message": f"Failed to process the file: {e}"}
        )

    # Überprüfe, ob die erforderlichen Spalten vorhanden sind
    required_columns = ['Straße', 'HsNr', 'PLZ', 'Ort', 'Haushalte']
    missing_columns = [col for col in required_columns if col not in excel_data.columns]
    if missing_columns:
        logging.error(f"Missing columns: {missing_columns} in file: {file.filename}")
        return JSONResponse(
            status_code=400,
            content={"status": "error", "message": f"The Excel file is missing required columns: {', '.join(missing_columns)}"}
        )
    
    latitude_list = []
    longitude_list = []

    # Geokodierung der Adressen
    for idx, row in excel_data.iterrows():
        full_address = f"{row['Straße']} {row['HsNr']}, {row['PLZ']} {row['Ort']}"
        url = f'https://maps.googleapis.com/maps/api/geocode/json?address={full_address}&key={API_KEY}'
        
        try:
            response = requests.get(url, timeout=20)
            logging.info(f"Request URL: {url}")
            logging.info(f"Response Status: {response.status_code}, Response Body: {response.text}")
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
    
    # Speichere als KML-Datei mit benutzerdefiniertem Stil
    try:
        kml = simplekml.Kml()
        
        # Definiere den magentafarbenen Stil für die Punkte
        magenta_style = simplekml.Style()
        magenta_style.iconstyle.color = 'ffff00ff'  # Magenta in KML-Farbformat
        magenta_style.iconstyle.scale = 1.0
        magenta_style.iconstyle.icon.href = 'https://upload.wikimedia.org/wikipedia/commons/thumb/a/ac/Location_dot_magenta.svg/1024px-Location_dot_magenta.svg.png'
        
        for _, row in valid_coords.iterrows():
            if pd.notna(row['Latitude']) and pd.notna(row['Longitude']):
                pnt = kml.newpoint(name=f"{row['Straße']} {row['HsNr']}", coords=[(row['Longitude'], row['Latitude'])])
                pnt.style = magenta_style
                
                # Falls Haushalte vorhanden ist, Beschreibung hinzufügen
                if pd.notna(row['Haushalte']):
                    pnt.description = f"Haushalte ({int(row['Haushalte'])})"
        
        kml_output_path = os.path.join(os.getcwd(), "streets_map_with_house_numbers.kml")
        kml.save(kml_output_path)
        logging.info(f"KML file successfully created and saved at: {kml_output_path}")
    except Exception as e:
        logging.error(f"Error creating KML file: {e}")
        logging.debug(traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={"status": "error", "message": f"Failed to create KML file: {e}"}
        )
    
    # Datei als Antwort zurückgeben
    return FileResponse(
        path=kml_output_path,
        media_type="application/vnd.google-earth.kml+xml",
        filename="streets_map_with_house_numbers.kml"
    )

@app.get("/upload/")
async def upload_info():
    """
    GET-Endpunkt für /upload/, um benutzerfreundliche Informationen zu liefern.
    """
    return {"message": "This endpoint only accepts POST requests for file uploads. Please use POST to upload a valid Excel file."}
