import uno
import time
import os
import json
import subprocess
import Pyro4
import threading
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import DisposedException

script_dir = os.path.dirname(os.path.abspath(__file__))


@Pyro4.expose
def command_exe():
    command = "ibreoffice --accept='socket,host=localhost,port=2003;urp;' --norestore --nologo --nodefault --impress"
    subprocess.Popen(command, shell=True)
@Pyro4.expose
def connect_to_libreoffice():
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_ctx)
    ctx = resolver.resolve("uno:socket,host=localhost,port=2003;urp;StarOffice.ComponentContext")
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop


    ###Abrir escucha 
    ###libreoffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo --nodefault --impress
@Pyro4.expose
def get_current_slide_idx(presentation, controller):
    slides = presentation.getDrawPages()
    slide_count = slides.getCount()
    current_page = controller.getCurrentPage()

    if current_page is None:
        return -1

    for i in range(slide_count):
        slide = slides.getByIndex(i)
        if slide == current_page: 
            return i 
         
    return -1
@Pyro4.expose
def save_slide_idx(filename, slide_index, data_file=None):
    if data_file is None:
        #home_dir = os.path.expanduser("~")
        data_file = os.path.join(script_dir, "slide_data.json")
    print(f"Guardando datos en: {data_file}")
    
    data = {}

    if os.path.exists(data_file):
        try: 
            with open(data_file, "r") as f:
                data = json.load(f)
        except Exception as e:
            print(f"error al leer el archivo JSON: {e}")

    data[filename] = slide_index 

    try: 
        with open(data_file, "w") as f:
            json.dump(data, f)
    
    except Exception as e:
        print(f"Error al guardar el archivo JSON: {e}")
@Pyro4.expose
def load_slide_idx(filename, data_file=None):

    if data_file is None:
        #home_dir = os.path.expanduser("~")
        data_file = os.path.join(script_dir, "slide_data.json")
    
    print(f"cargando datos desde: {data_file}")
    
    if not os.path.exists(data_file):
        return -1
    
    try: 
        with open(data_file, "r") as f:
            data = json.load(f)

    except Exception as e: 
        print(f"Error al leer el archivo JSON: {e}")
        return -1
    return data.get(filename, -1)

def infinite_loop():
    desktop = connect_to_libreoffice()
    #last_slide_idx = -1
    while True:
        try:
            presentation = desktop.getCurrentComponent()
            if presentation is not None and presentation.supportsService("com.sun.star.presentation.PresentationDocument"):
                print("presentacion detectada")
                file_url = presentation.getURL()
                filename = os.path.basename(file_url)
                current_slide_idx = load_slide_idx(filename)

                controller = presentation.getCurrentController()

                if controller is not None:
                    slide_index = get_current_slide_idx(presentation,controller)
                    if slide_index != current_slide_idx:
                         current_slide_idx = slide_index
                         save_slide_idx(filename, current_slide_idx)
                         print(f"slide actual: {current_slide_idx + 1}" )

                else:
                    print("controlador no valido o incompatible")
            else:
                print("la presentacion no esta abierta o no compatible")
        except Exception as e:
            print("Error", e)
        time.sleep(1)
        
def main():
    loop_thread = threading.Thread(target=infinite_loop, daemon=True)
    loop_thread.start()
    #last_slide_idx = -1
    daemon = Pyro4.Daemon(port=37459)
    uri1 = daemon.register(command_exe)
    uri2 = daemon.register(connect_to_libreoffice)
    uri3 = daemon.register(get_current_slide_idx)
    uri4 = daemon.register(save_slide_idx)
    uri5 = daemon.register(load_slide_idx)

    print(f"function1 esta disponible en: {uri1}")
    print(f"function2 esta disponible en: {uri2}")
    print(f"function3 esta disponible en: {uri3}")
    print(f"function4 esta disponible en: {uri4}")
    print(f"function5 esta disponible en: {uri5}")

    daemon.requestLoop()
    

    #current_slide_idx = None

    # while True:
    #     try: 
    #         presentation = desktop.getCurrentComponent()
    #         if presentation is not None and presentation.supportsService("com.sun.star.presentation.PresentationDocument"):
    #             print("presentacion detectada")


    #             controller = presentation.getCurrentController()
    #             #slide_index = get_current_slide_idx(presentation,controller)
    #             if controller is not None:
    #                 slide_index = get_current_slide_idx(presentation,controller)

    #                 if slide_index != current_slide_idx:
    #                     current_slide_idx = slide_index
    #                     print(f"slide actual: {current_slide_idx + 1}" )

    #             else:
    #                 print("controlador no valido o incompatible")
    #         else:
    #             print("la presentacion no esta abierta o no compatible")
    #     except Exception as e:
    #         print("Error", e)
    #     time.sleep(1)

if __name__ == "__main__":
    main()

