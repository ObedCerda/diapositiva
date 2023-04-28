import uno
import time
import os
import json
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import DisposedException

script_dir = os.path.dirname(os.path.abspath(__file__))



#Archivo pensado para recibir por rpc valor de id = nombre de presentacion y abrir presentacion // Incompleto

def find_presentation(presentation_name, search_directory):
    for root, _, files in os.walk(search_directory):
        for file in files:
            if file == presentation_name:
                return os.path.join(root, file)
    return None

def open_presentation(desktop, file_path, slide_index):
    input_url = uno.systemPathToFileUrl(file_path)

    hidden_property = PropertyValue()
    hidden_property.Name = "hidden"
    hidden_property.Value = True

    doc = desktop.loadComponentFromURL(input_url, "_blank", 0, (hidden_property,))

    if doc.supportsService("com.sun.star.presentation.PresentationDocument"):
        doc.CurrentController.setCurrentPage(doc.DrawPages.getByIndex(slide_index))

        visible_property = PropertyValue()
        visible_property.Name = "Hidden"
        visible_property.Value = False

        doc.storeToUrl(input_url, (visible_property,))
        doc.dispose()

def main():
    desktop = connect_to_libreoffice()
    #last_slide_idx = -1
    slide_data = load_slide_idx()
    presentation_name = input("ingresar nombre de la presentacion")

    file_path = None
    slide_index = -1

    for key, value in slide_data.items():
        if key.replace("%20", " ") == presentation_name:
            file_path = uno.fileUrlToSystemPath(key)
            slide_index = value
            break
        
    if file_path is None:
        search_directory = "/home/a-slider/Documentos/Slides"
        file_path = find_presentation(presentation_name, search_directory)

    if file_path is not None and slide_index >= 0:
        open_presentation(desktop, file_path, slide_index)
    else:
        print("No se encontro")


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

