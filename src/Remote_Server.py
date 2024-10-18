import socket
import time
import win32api
import win32gui
import atexit
import msvcrt
import threading
from collections import namedtuple
from ctypes import windll, CFUNCTYPE, POINTER, c_int, c_void_p, byref
import cv2
import win32con
import os
SHEIGHT, SWIDTH = win32api.GetSystemMetrics(1), win32api.GetSystemMetrics(0)

height, width = 0, 0
KeyboardEvent = namedtuple('KeyboardEvent', ['event_type', 'key_code',
                                             'scan_code', 'alt_pressed',
                                             'time'])
handlers = []
SERV_IP = "10.100.102.3"
press_list = []
depress_list = []
remove_list = []
count = 0

VK_CODE = {'backspace': 0x08,
           'tab': 0x09,
           'clear': 0x0C,
           'enter': 0x0D,
           'shift': 0x10,
           'ctrl': 0x11,
           'alt': 0x12,
           'pause': 0x13,
           'caps_lock': 0x14,
           'esc': 0x1B,
           'spacebar': 0x20,
           'page_up': 0x21,
           'page_down': 0x22,
           'end': 0x23,
           'home': 0x24,
           'left_arrow': 0x25,
           'up_arrow': 0x26,
           'right_arrow': 0x27,
           'down_arrow': 0x28,
           'select': 0x29,
           'print': 0x2A,
           'execute': 0x2B,
           'print_screen': 0x2C,
           'ins': 0x2D,
           'del': 0x2E,
           'help': 0x2F,
           '0': 0x30,
           '1': 0x31,
           '2': 0x32,
           '3': 0x33,
           '4': 0x34,
           '5': 0x35,
           '6': 0x36,
           '7': 0x37,
           '8': 0x38,
           '9': 0x39,
           'a': 0x41,
           'b': 0x42,
           'c': 0x43,
           'd': 0x44,
           'e': 0x45,
           'f': 0x46,
           'g': 0x47,
           'h': 0x48,
           'i': 0x49,
           'j': 0x4A,
           'k': 0x4B,
           'l': 0x4C,
           'm': 0x4D,
           'n': 0x4E,
           'o': 0x4F,
           'p': 0x50,
           'q': 0x51,
           'r': 0x52,
           's': 0x53,
           't': 0x54,
           'u': 0x55,
           'v': 0x56,
           'w': 0x57,
           'x': 0x58,
           'y': 0x59,
           'z': 0x5A,
           'numpad_0': 0x60,
           'numpad_1': 0x61,
           'numpad_2': 0x62,
           'numpad_3': 0x63,
           'numpad_4': 0x64,
           'numpad_5': 0x65,
           'numpad_6': 0x66,
           'numpad_7': 0x67,
           'numpad_8': 0x68,
           'numpad_9': 0x69,
           'multiply_key': 0x6A,
           'add_key': 0x6B,
           'separator_key': 0x6C,
           'subtract_key': 0x6D,
           'decimal_key': 0x6E,
           'divide_key': 0x6F,
           'F1': 0x70,
           'F2': 0x71,
           'F3': 0x72,
           'F4': 0x73,
           'F5': 0x74,
           'F6': 0x75,
           'F7': 0x76,
           'F8': 0x77,
           'F9': 0x78,
           'F10': 0x79,
           'F11': 0x7A,
           'F12': 0x7B,
           'F13': 0x7C,
           'F14': 0x7D,
           'F15': 0x7E,
           'F16': 0x7F,
           'F17': 0x80,
           'F18': 0x81,
           'F19': 0x82,
           'F20': 0x83,
           'F21': 0x84,
           'F22': 0x85,
           'F23': 0x86,
           'F24': 0x87,
           'num_lock': 0x90,
           'scroll_lock': 0x91,
           'left_shift': 0xA0,
           'right_shift ': 0xA1,
           'left_control': 0xA2,
           'right_control': 0xA3,
           'left_menu': 0xA4,
           'right_menu': 0xA5,
           'browser_back': 0xA6,
           'browser_forward': 0xA7,
           'browser_refresh': 0xA8,
           'browser_stop': 0xA9,
           'browser_search': 0xAA,
           'browser_favorites': 0xAB,
           'browser_start_and_home': 0xAC,
           'volume_mute': 0xAD,
           'volume_Down': 0xAE,
           'volume_up': 0xAF,
           'next_track': 0xB0,
           'previous_track': 0xB1,
           'stop_media': 0xB2,
           'play/pause_media': 0xB3,
           'start_mail': 0xB4,
           'select_media': 0xB5,
           'start_application_1': 0xB6,
           'start_application_2': 0xB7,
           'attn_key': 0xF6,
           'crsel_key': 0xF7,
           'exsel_key': 0xF8,
           'play_key': 0xFA,
           'zoom_key': 0xFB,
           'clear_key': 0xFE,
           '+': 0xBB,
           ',': 0xBC,
           '-': 0xBD,
           '.': 0xBE,
           '/': 0xBF,
           '`': 0xC0,
           ';': 0xBA,
           '[': 0xDB,
           '\\': 0xDC,
           ']': 0xDD,
           "'": 0xDE,
           '`': 0xC0}

SERV_PORT_SCREEN = 50010
SERV_PORT_MOUSE = 50011
SERV_PORT_KEYBOARD = 50012
SERV_PORT_DISCONNECT = 50013
SERV_PORT_KEYBOARD_REMOVE = 50014

server_socket_screen = socket.socket()  # the server's screen socket
server_socket_screen.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
server_socket_screen.bind((SERV_IP, SERV_PORT_SCREEN))  # binds the server an address
server_socket_screen.listen(1)

server_socket_mouse = socket.socket()  # the server's mouse socket
server_socket_mouse.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
server_socket_mouse.bind((SERV_IP, SERV_PORT_MOUSE))  # binds the server an address
server_socket_mouse.listen(1)

server_socket_keyboard = socket.socket()  # the server's keyboard socket
server_socket_keyboard.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
server_socket_keyboard.bind((SERV_IP, SERV_PORT_KEYBOARD))  # binds the server an address
server_socket_keyboard.listen(1)

server_socket_keyboard_remove = socket.socket()  # the server's key removal socket
server_socket_keyboard_remove.bind((SERV_IP, SERV_PORT_KEYBOARD_REMOVE))  # binds the server an address
server_socket_keyboard_remove.listen(1)

server_socket_disconnect = socket.socket()
server_socket_disconnect.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
server_socket_disconnect.bind((SERV_IP, SERV_PORT_DISCONNECT))
server_socket_disconnect.listen(1)

client_socket_screen = ''
client_socket_mouse = ''
client_socket_keyboard = ''
check_exit = False

save_path=str(raw_input("Please enter an existing path to save the screenshots in, for instance: C:\\User\n"))
while os.path.exists(save_path)== False:
    print 'Error! please try another path!'
    save_path = str(raw_input("Please enter an existing path to save the screenshots in, for instance: C:\\User\n"))
save_path+="\\MySnapshot1.jpg"
#save_path = "C:\\Users\\Aviv\\Desktop\\MySnapshot1.jpg"

print 'starting the VNC!'


def handle_screen():
    '''this function recieves and shows the screen capture from the client'''
    global client_socket_screen
    global save_path
    (client_socket_screen, client_address_screen) = server_socket_screen.accept()  # accepts the screen socket
    client_socket_screen = client_socket_screen
    print 'A new client screen connection established!'
    while True:  # infinite loop
        try:
            s = ""  # the whole img data
            data = ""  # the recieved data through the socket
            while data != "The file transmittion has ended.":  # untill this is sent
                data = client_socket_screen.recv(4096)  # recieves parts of the data
                if data != "The file transmittion has ended.":
                    s += data
            try:
                with open(save_path, "wb") as handle:  # saves the image on the computer
                    handle.write(s)
            except:
                continue
            img = cv2.imread(save_path)  # saves the img in cv2 format
            cv2.imshow("Client_Screen", img)  # displays the image
            cv2.waitKey(1)  # a must for cv2 imshow
        except socket.error:
            client_socket_screen.close()  # closes the client socket
            cv2.destroyWindow("Client_Screen")
            cv2.destroyAllWindows()  # closes cv2 windows
            print "\nclient's screenshare disconnected\n"
            (client_socket_screen, client_address_screen) = server_socket_screen.accept()  # accepts the server socket
            print 'A new client screen connection established!'


def handle_cursor_values(h, w):
    '''
    this function recieves the client's screen dimensions gets the mouse data from the server, recieves the client's
    screen dimensions and adjusts accordingly, returns the finished cursor values, in string
    '''
    x, y = win32api.GetCursorPos()  # the mouse's current pos
    # global SWIDTH, SHEIGHT #the screen's dimensions
    # x = int(x /(SWIDTH/float(int(w)))) #the width of the screen
    # y = int(y / (SHEIGHT/float(int(h)))) #the height of the screen
    state_left = win32api.GetKeyState(0x01)  # the state of left button
    state_right = win32api.GetKeyState(0x02)  # the state of left button
    return str(x) + "," + str(y) + "," + str(state_left) + "," + str(state_right)  # the finished cursor values


def handle_cursor_output():
    global client_socket_mouse
    '''sends the mouse data to the client'''
    (client_socket_mouse, client_address_mouse) = server_socket_mouse.accept()  # accepts the mouse socket
    client_socket_mouse = client_socket_mouse
    print 'A new client mouse connection established!'
    global height, width  # the client's screen dimensions
    height, width = client_socket_mouse.recv(10).split(",")  # recieves the client's screen dimensions
    while True:  # infinite loop
        try:
            data = handle_cursor_values(height, width)  # the cursor's data
            time.sleep(0.001)  # to prevent scrambled data
            client_socket_mouse.send(str(len(data)))  # the length of the data
            time.sleep(0.001)
            client_socket_mouse.send(data)  # send the data to the client
        except socket.error:
            client_socket_mouse.close()  # closes the client socket
            print "\nclient's mouse disconnected\n"
            (client_socket_mouse, client_address_mouse) = server_socket_mouse.accept()  # accepts the cursor socket
            print 'A new client mouse connection established!'
            height, width = client_socket_mouse.recv(10).split(",")  # the new client's screen dimensions


def key_status(key):
    '''this  function checks the recieved key's status
     it recieves a key, and if it returns True, the key is up, otherwise the key is down
     '''
    keystate = win32api.GetKeyState(key)  # the key's state
    if keystate == 0 or keystate == 1:  # the key is unpressed
        return True
    else:
        return False  # the key is pressed


def keyboard_input():
    '''handles the keyboard input, by sending the client clicked/holded keys'''
    (client_socket, client_address) = server_socket_keyboard.accept()  # accepts the socket
    print 'A new client keyboard connection established!'
    while True:  # infinite loop
        try:
            pressed_key = msvcrt.getch()  # tries to recieve keyboard click
            depress_list.append(VK_CODE.get(pressed_key))  # the list of pressed keys, to depress them eventually
            data = "press," + str(VK_CODE.get(pressed_key))  # the value of the pressed key
            time.sleep(0.005)  # to prevent data scramble
            client_socket.send(str(len(data)))  # send the length of data
            time.sleep(0.005)
            client_socket.send(data)  # send the data to the client
        except socket.error:
            client_socket.close()  # close the client socket
            print "client's keyboard disconnected\n"
            (client_socket, client_address) = server_socket_keyboard.accept()  # accepts the keyboard socket
            print 'A new client keyboard connection established!'
            # print press_list
            # win32api.keybd_event()
            # win32api.GetKeyState()


def keys_status():
    '''handles holded keys/keys that were pressed, and sends their current value to the client'''
    (client_socket, client_address) = server_socket_keyboard_remove.accept()  # accepts the socket
    print 'A new client keyboard correction connection established!'
    try:
        while True:  # infinite loop
            if len(depress_list) > 0:  # checks if there are elements
                for i in len(depress_list):  # runs through the whole list
                    keystate = win32api.GetKeyState(depress_list[i])  # checks the keystate
                    if keystate == 0 or keystate == 1:  # means that the keys are up
                        data = "remove," + str(depress_list[i])  # the key to rise on the client
                        time.sleep(0.001)  # prevent data scramble
                        client_socket.send(str(len(data)))  # send the length of data
                        time.sleep(0.001)
                        client_socket.send(data)  # send the data to the client
                        remove_list.append(
                            i)  # the keys to remove from the depress list, they were changed to up already
            if len(remove_list) > 0:  # checks if there are elements
                count = 0
                for k in remove_list:  # runs through the list
                    del depress_list[k - count]  # removes the elements needed to remove
                    count += 1
    except socket.error:
        client_socket.close()  # closes the client socket
        print "\nclient's keyboard correction disconnected\n"
        (client_socket, client_address) = server_socket_keyboard_remove.accept()  # accepts the key removal socket
        print 'A new client keyboard correction connection established!'
    except:
        print "Error!"


def start_keyboard():
    while True:
        try:
            global client_socket_keyboard, check_exit
            (client_socket_keyboard, client_address_keyboard) = server_socket_keyboard.accept()
            # accepts the keyboard's socket connection
            client_socket_keyboard = client_socket_keyboard
            print 'A new client keyboard connection established!'
            handle_keyboard(client_socket_keyboard)
        except socket.error:
            client_socket_keyboard.close()
            print '\nkeyboard disconnected!'
        finally:
            pass


def handle_keyboard(client_address_keyboard):
    """
    handles the keyboard input, by sending the client clicked/holded keys
    Calls handlers for each keyboard event received.(blocking call)
    """
    global check_exit
    event_types = {win32con.WM_KEYDOWN: 'key down',
                   win32con.WM_KEYUP: 'key up',
                   0x104: 'key down',  # Alt key down.
                   0x105: 'key up',  # Alt key up.
                   }

    def low_level_handler(nCode, wParam, lParam):
        """
        Processes a low level Windows keyboard event
        """
        event = KeyboardEvent(event_types[wParam], lParam[0], lParam[1],
                              lParam[2] == 32, lParam[3])  # a keyboard event
        handle_keys(event, client_address_keyboard)  # sends the event to the client
        return windll.user32.CallNextHookEx(hook_id, nCode, wParam, lParam)  # calls the next hook

    CMPFUNC = CFUNCTYPE(c_int, c_int, c_int, POINTER(c_void_p))  # Our low level handler signature.
    pointer = CMPFUNC(low_level_handler)  # Convert the Python handler into C pointer.

    # Hook both press and unpress for common keys.
    hook_id = windll.user32.SetWindowsHookExA(win32con.WH_KEYBOARD_LL, pointer,
                                              win32api.GetModuleHandle(None), 0)

    # Register to remove the hook when the interpreter exits.
    atexit.register(windll.user32.UnhookWindowsHookEx, hook_id)
    while check_exit is False:
        msg = win32gui.GetMessage(None, 0, 0)  # uses the winmsg with the handler
        win32gui.TranslateMessage(byref(msg))
        win32gui.DispatchMessage(byref(msg))
        pass


def handle_keys(event, client_socket_keyboard):
    '''handles holded keys/keys that were pressed, and sends their current value to the client'''
    global check_exit
    lst = list(event)
    data = lst[0] + "," + str(lst[1]) + "," + str(lst[2]) + "," + str(int(lst[3])) + "," + str(lst[4])
    try:
        client_socket_keyboard.send(str(len(data)))  # send the length of data
        client_socket_keyboard.send(data)  # send the data to the client
    except:
        client_socket_keyboard.close()


def close_all():
    '''this function closes all socket connections for an optimal shutdown.'''
    global client_socket_keyboard
    global client_socket_mouse
    global client_socket_screen
    (client_socket, client_address) = server_socket_disconnect.accept()  # accepts the screen socket
    print 'A new client disconnect connection established!'
    while True:  # infinite loop
        try:
            data = client_socket.recv(4)
            if data == 'EXIT':
                global check_exit
                check_exit = True
                client_socket_mouse.close()
                client_socket_keyboard.close()
                client_socket_screen.close()
                client_socket.close()
                cv2.destroyWindow("Client_Screen")
                cv2.destroyAllWindows()  # closes cv2 windows
                print "closing client\n"
                time.sleep(2)
                break
        except socket.error:
            client_socket.close()  # closes the client socket
            cv2.destroyWindow("Client_Screen")
            cv2.destroyAllWindows()  # closes cv2 windows
            print "closing client\n"
            break


def main():
    hs_thread = threading.Thread(target=handle_screen, args=())  # handle screen thread
    hs_thread.start()  # start it
    # hs_thread.join()

    hm_thread = threading.Thread(target=handle_cursor_output, args=())  # handle mouse thread
    hm_thread.start()  # start it

    hk_thread = threading.Thread(target=start_keyboard, args=())  # handle keyboard thread
    hk_thread.start()  # start it

    close_all()
    # os.system("mspaint")#open paint to prevent harming the desktop of the server

    # hs_thread=threading.Thread(target=handle_keyboard, args=(client_socket_keyboard, ))
    # hs_thread.start()
    # hs_thread.join()
    while True:
        time.sleep(1)


if __name__ == '__main__':
    main()
