from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import FileResponse
import pandas as pd
import requests
import folium
import os
import simplekml
import time

app = FastAPI()

API_KEY = 'Your_Google_Maps_API_Key_Here'

@app.post("/upload/")
async def upload_file(file: UploadFile):
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Invalid file format. Please upload an Excel file.")
    
    try:
        street_list = pd.read_excel(file.file)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to process the file: {e}")
    
    required_columns = ['Straße', 'HsNr', 'PLZ', 'Ort']
    for col in required_columns:
        if col not in street_list.columns:
            raise HTTPException(status_code=400, detail=f"The Excel file must contain the '{col}' column.")

    latitude_list = []
    longitude_list = []

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
        except Exception as e:
            latitude_list.append(None)
            longitude_list.append(None)
            print(f"Error fetching coordinates for address {full_address}: {e}")
    
    street_list['Latitude'] = latitude_list
    street_list['Longitude'] = longitude_list

    # Filter valid coordinates
    valid_coords = street_list.dropna(subset=['Latitude', 'Longitude'])
    if valid_coords.empty:
        raise HTTPException(status_code=500, detail="No valid addresses with coordinates found.")
    
    # Save as KML
    kml = simplekml.Kml()
    for _, row in valid_coords.iterrows():
        if pd.notna(row['Latitude']) and pd.notna(row['Longitude']):
            kml.newpoint(name=f"{row['Straße']} {row['HsNr']}", coords=[(row['Longitude'], row['Latitude'])])
    kml_output_path = "streets_map_with_house_numbers.kml"
    kml.save(kml_output_path)
    
    return FileResponse(kml_output_path, media_type="application/vnd.google-earth.kml+xml", filename=kml_output_path)
