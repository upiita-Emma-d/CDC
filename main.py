import tkinter as tk
from tkinter import filedialog
import shapefile
from pyproj import Proj, transform
import pandas as pd
from math import atan2, degrees, sqrt
import os

def calculate_distance(point1, point2):
    return sqrt((point1[0] - point2[0])**2 + (point1[1] - point2[1])**2)

def calculate_bearing(point1, point2):
    angle_rad = atan2(point2[1] - point1[1], point2[0] - point1[0])
    angle_deg = degrees(angle_rad)
    bearing = (angle_deg + 360) % 360
    return bearing

def calculate_rumbo(bearing):
    degrees = int(bearing)
    minutes = int((bearing - degrees) * 60)
    seconds = (bearing - degrees - minutes/60) * 3600
    return f"N {degrees}°{minutes}'{seconds:.2f}\" E"

def process_shapefile(shapefile_path, csv_output_path):
    sf = shapefile.Reader(shapefile_path)
    
    in_proj = Proj('epsg:6372')  # Ajusta según el EPSG de tu shapefile
    out_proj = Proj('epsg:4326')  # EPSG para WGS84
    
    records = []

    for shape in sf.shapes():
        points = shape.points
        for i in range(len(points) - 1):
            point1 = points[i]
            point2 = points[i + 1]
            distance = calculate_distance(point1, point2)
            bearing = calculate_bearing(point1, point2)
            rumbo = calculate_rumbo(bearing)
            
            lat1, lon1 = transform(in_proj, out_proj, point1[0], point1[1])

            records.append({
                "Est": i+1,
                "PV": i+2,
                "Rumbo": rumbo,
                "Distancia": distance,
                "Latitud": lat1,
                "Longitud": lon1,
            })
    
    df = pd.DataFrame(records)
    df.to_csv(csv_output_path, index=False)
    print(f"Archivo CSV guardado en {csv_output_path}")

def select_shapefile():
    root = tk.Tk()
    root.withdraw()  # No queremos una ventana completa de Tk, solo la caja de diálogo
    file_path = filedialog.askopenfilename(filetypes=[("Shapefiles", "*.shp")])
    return file_path

if __name__ == "__main__":
    shapefile_path = select_shapefile()
    if shapefile_path:
        print(f"Archivo seleccionado: {shapefile_path}")
        # Define la ruta de salida para el archivo CSV
        csv_output_path = 'salida.csv'
        # Procesa el shapefile seleccionado
        process_shapefile(shapefile_path, csv_output_path)
    else:
        print("No se seleccionó ningún archivo.")
