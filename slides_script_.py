
from xmlrpc.server import SimpleXMLRPCServer

var_global = None

def process_string(input_string):
    print(f"Procesando string: {type(input_string)}")
    var_global = input_string
    return type(var_global)
    




def main():
    server = SimpleXMLRPCServer(("192.168.3.28", 8000))
    print("Escuchando en el puerto  8000...")
    server.register_function(process_string, "process_string")
    server.serve_forever()
    print(var_global)

if __name__ == "__main__":
    main()