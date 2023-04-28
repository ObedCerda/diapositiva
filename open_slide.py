import socket
import json
import time
import uno
import subprocess
import os
from xmlrpc.server import SimpleXMLRPCServer
import Pyro4
from com.sun.star.connection import NoConnectException
from com.sun.star.beans import PropertyValue

JSON_FILE_PATH = "/home/a-slider/Documentos/scripts/slide_data.json"


def connect_to_libreoffice():
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_ctx)
    ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")

    return ctx, resolver
    if name in json_data:
        return json_data[name]
    else:
        return None
    
# def close_connection(ctx):
#     bridge = ctx.getBridge()
#     if bridge:
#         bridge.dispose()

def release_port(resolver):
    try:
        bridge = resolver.getBridge()
        bridge.terminate()
    except Exception as e:
        print(f"Error al liberar el puerto: {e}")


def open_document_at_slide(ctx, file_path,slide_number):
    try:
        
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        #Open file in LibreOffice
        url = uno.systemPathToFileUrl(file_path)
        document = desktop.loadComponentFromURL(url, "_blank", 0, ())

        #change to specific slide

        if document is not None:
            controller = document.getCurrentController()
            if controller is not None:
                controller.setCurrentPage(document.getDrawPages().getByIndex(slide_number -1))

        else:
            print("No se pudo obtener el documento")
        
    except Exception as e:
        print(f"Erro: {e}")

def load_presentation_data_from_json(json_file, presentation_name):
    with open(json_file, "r") as file:
        data = json.load(file)
    if presentation_name in data:
        return data[presentation_name]
    else:
        return None
    
# def open_presentation_with_name(received_name):

#     json_file_path = "/home/a-slider/Documentos/scripts/slide_data.json"
#     json_data =read_json_file(json_file_path)

# received_name = "H123J123.odp"
def main():
    uri1 = "PYRO:obj_7dd1226089f0437fb35bcd08ac755165@localhost:37459"
    uri2 = "PYRO:obj_4628de88e5b944d1b16f675650413d67@localhost:37459"
    uri3 = "PYRO:obj_09b593d4228d420aa3f8cb96d85d160b@localhost:37459"
    uri4 = "PYRO:obj_09b593d4228d420aa3f8cb96d85d160b@localhost:37459"
    uri5 = "PYRO:obj_9e73c0a9947145ba89fa9b25be81bab6@localhost:37459"

    function1_proxy = Pyro4.Proxy(uri1)
    function2_proxy = Pyro4.Proxy(uri2)
    function3_proxy = Pyro4.Proxy(uri3)
    function4_proxy = Pyro4.Proxy(uri4)
    function5_proxy = Pyro4.Proxy(uri5)

 
    function1_proxy()
    function2_proxy()
    function3_proxy()
    function4_proxy()
    function5_proxy()







if __name__ == "__main__":
    ctx,resolver = connect_to_libreoffice()
    presentation_name = "H123J123.odp"
    slide_number = load_presentation_data_from_json(JSON_FILE_PATH, presentation_name)
    if slide_number is not None:
        file_path = "/home/a-slider/Documentos/Slides" + "/" + presentation_name
        open_document_at_slide(ctx, file_path, slide_number)
        print(slide_number)

    else: 
        print(f"No se encontro el nombre {presentation_name} en el archivo JSON. ")

    release_port(resolver)
    subprocess.Popen(["python3", "slide_idx.py"])


# json_file_path = "/home/a-slider/Documentos/scripts/slide_data.json"
# json_data = read_json_file(json_file_path)

# received_name = "H123J123.odp"

# slide_number = find_slide_number_by_name(received_name, json_data)


#Esta a la escucha de el ID desde la DB
# server = SimpleXMLRPCServer(("192.168.3.28", 8000))
# print("escuchando en el puerto 8000...")
# server.register_function(open_doc, "open_doc")


# stop_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
# stop_socket.bind(("localhost", 8001))
# stop_socket.listen(1)


# while not stop_server:
#     server.handle_request()




    





