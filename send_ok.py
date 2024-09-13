import socket

def send_message(message):
    client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    client_socket.connect(('localhost', 6400))
    client_socket.send(message.encode())
    client_socket.close()

if __name__ == "__main__":
    send_message("OK")

    