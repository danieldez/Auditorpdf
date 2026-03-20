import os
import sys


#NOMBRE DE LA CARPETA
APP_NAME = "AuditorValidaciones"


def get_app_path():
    #%LOCALAPPDATA% ES LA RUTA ESTANDAR PARA GUARDAR DATOS DE APLICACIONES EN WINDOWS
    base_path = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
    app_path = os.path.join(base_path, APP_NAME)
    
    # SI LA CARPETA NO EXISTE, LA CREAMOS
    if not os.path.exists(app_path):
        os.makedirs(app_path)
        
        
    return app_path

def get_json_path():
    return os.path.join(get_app_path(), 'plantillas.json')

#variable global para la ruta del JSON y para usar en todo el proyecto
RUTAS = {
    "json": get_json_path(),
    "carpeta": get_app_path()
}