import socket
client_sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
client_sock.sendto(b'trigger', ('127.0.0.1', 47123))
print("Trigger sent!")
