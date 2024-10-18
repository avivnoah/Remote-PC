import socket
import threading
import win32api
import time
import StringIO
import ctypes
import msvcrt
import sys
import os
from ctypes import wintypes

import cv2
import numpy
import mss
import win32con
from PIL import Image


CLIENT_IP = "10.100.102.3"

CLIENT_PORT_SCREEN = 50010
CLIENT_PORT_MOUSE = 50011
CLIENT_PORT_KEYBOARD = 50012
CLIENT_PORT_DISCONNECT = 50013

client_socket_screen = socket.socket()  #the screen client's socket
client_socket_mouse = socket.socket()  #the mouse client's socket
client_socket_keyboard = socket.socket()  #the keyboard input client's socket
client_socket_disconnect = socket.socket()  #the disconnect clients socket

SHEIGHT, SWIDTH = win32api.GetSystemMetrics(1), win32api.GetSystemMetrics(0)
monitor = {'top': 0, 'left': 0, 'width': 1366, 'height': 768}  #the margin of the monitor and it's resolution
sbuff = []  #a buffer that contains screenshots

user32 = ctypes.WinDLL('user32', use_last_error=True)

check_stop = False

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
INPUT_HARDWARE = 2

KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
KEYEVENTF_SCANCODE = 0x0008

MAPVK_VK_TO_VSC = 0
VK_TAB = 0x09
VK_MENU = 0x12

wintypes.ULONG_PTR = wintypes.WPARAM

#data of mouse is unused, but needed for the program to work correctly
class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx", wintypes.LONG),
                ("dy", wintypes.LONG),
                ("mouseData", wintypes.DWORD),
                ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = (("uMsg", wintypes.DWORD),
                ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD))


class KEYBDINPUT(ctypes.Structure):
    _fields_ = (("wVk", wintypes.WORD),
                ("wScan", wintypes.WORD),
                ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

    def __init__(self, *args, **kwds):
        super(KEYBDINPUT, self).__init__(*args, **kwds)
        # some programs use the scan code even if KEYEVENTF_SCANCODE
        # isn't set in dwFflags, so attempt to map the correct code.
        if not self.dwFlags & KEYEVENTF_UNICODE:
            self.wScan = user32.MapVirtualKeyExW(self.wVk,
                                                 MAPVK_VK_TO_VSC, 0)


class INPUT(ctypes.Structure):  #the input class
    class _INPUT(ctypes.Union):  #a c-like union
        _fields_ = (("ki", KEYBDINPUT),  #the keyboard input
                    ("mi", MOUSEINPUT),  #the mouse input
                    ("hi", HARDWAREINPUT))  #the hardware input(winmsg, for instance)

    _anonymous_ = ("_input",)
    _fields_ = (("type", wintypes.DWORD),  #the type of input and input fields
                ("_input", _INPUT))


LPINPUT = ctypes.POINTER(INPUT)

#def check_count(result, func, args):
#if result == 0:
#raise ctypes.WinError(ctypes.get_last_error())
#return args

def press_key(hexKeyCode):
    key = INPUT(type=INPUT_KEYBOARD,
                ki=KEYBDINPUT(wVk=hexKeyCode))
    user32.SendInput(1, ctypes.byref(key), ctypes.sizeof(key))


def release_key(hexKeyCode):
    key = INPUT(type=INPUT_KEYBOARD,
                ki=KEYBDINPUT(wVk=hexKeyCode,
                              dwFlags=KEYEVENTF_KEYUP))
    user32.SendInput(1, ctypes.byref(key), ctypes.sizeof(key))


def set_mouse_pos(w, h):
    '''
    recieves the screen's height and width, and places the cursor accordingly
    '''
    win32api.SetCursorPos((w, h))  #places the cursor in w,h


def handle_mouse_clicks(state_left, state_right, new_state_left, new_state_right):
    '''
    this function recieves the current and the new mouse left&right buttons' states, and changes them if they have been
    changed, and it returns the new mouse button states, to allow their hold
    '''
    if new_state_left != state_left:  #Button's state changed
        state_left = new_state_left
        #print 'new_state_left'
        if new_state_left < 0:
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)  #left button pressed
            #print 'left button pressed'
        else:
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)  #left button unpressed
            #print 'left button released'

    if new_state_right != state_right:  #Button's state changed
        state_right = new_state_right
        #print 'new_state_right'
        if new_state_right < 0:
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)  #right button pressed
            #print 'right button pressed'
        else:
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)  #right button unpressed
            #print 'right button released'
    return new_state_left, new_state_right  #returns the new buttons' states


def handle_mouse():
    '''
    this function connects to the server with the cursor's socket, and controls the client's cursor by activating the
    cursor's functions. moreover, it changes the cursor state if it was changed on the server side
    '''
    global check_stop
    global client_socket_mouse
    while check_stop is False:
        client_socket_mouse.connect((CLIENT_IP, CLIENT_PORT_MOUSE))  #connects to the server with the mouse socket
        print 'A new mouse client connection established!'
        client_socket_mouse.send(str(SHEIGHT) + "," + str(SWIDTH))
        try:
            state_left, state_right = 0, 0  #the starting states are unpressed
            while check_stop is False:  #an infinite loop
                ln = client_socket_mouse.recv(2)  #the length of the data that is going to be recieved
                if ln.isdigit():  #checks if the length is only digits, and wasn't mixed with letters
                    s = str(client_socket_mouse.recv(int(ln)))  #the cursor's new data
                    s = s.split(",")
                    if s:
                        if len(s) >= 2 and s[0] != "" and s[1] != "":  #checks if the length makes sense
                            set_mouse_pos(int(s[0]), int(s[1]))  #sets the cursor's position to a new one
                        if len(s) >= 4 and s[2] != "" and s[3] != "":  #checks again, if  the length makes sense
                            state_left, state_right = handle_mouse_clicks(state_left, state_right, int(s[2]), int(s[3]))
                            #changes the cursor's buttons
        except socket.error:
            client_socket_mouse.close()
            print "Client's mouse Disconnected!"
            break
    try:
        client_socket_mouse.close()
        client_socket_mouse = socket.socket()
        print "Client's mouse IS Disconnected!\n"
    finally:
        pass


def handle_screen_input():
    '''
    handles the client's screen input, the recieved screenshots, and stores them in a buffer
    '''
    while True:  #an infinite loop
        if len(sbuff) < 10:  #limits the buffer size to 10
            buffer = StringIO.StringIO()  #a buffer to store the image in
            img = numpy.array(mss.mss().grab(monitor))  #the image in numpy array
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)  #changes it to BGR from RGB to support cv2 display format
            image = Image.fromarray(img.astype('uint8'), 'RGB')  #then turns the image to RGB
            image = Image.fromarray(img.astype('uint8'))  #procceses it to the right format to be saved on server side
            image.save(buffer, "JPEG", quality=10)  #saves it as a jpeg image, and lowers it's quality to be sent faster
            sbuff.append(buffer.getvalue())  #adds the image to the buffer


def handle_screen_output():
    '''
    this function connects to the server with the screen's socket, and sends the screenshots to the server, as long
    as the buffer isn't empty
    '''
    global check_stop
    global client_socket_screen
    client_socket_screen.connect((CLIENT_IP, CLIENT_PORT_SCREEN))
    #connects between the screen socket and the following address
    while check_stop is False:  #an infinite loop
        try:
            if sbuff:  #checks if the screenshots' buffer isn't empty, if it isn't
                #, it proceeds to the following lines of code
                data = sbuff[0]  #takes the first screenshot from the buffer
                del sbuff[0]  #and deletes it from it
                while len(data) != 0:  #runs till data var is empty
                    client_socket_screen.send(data[0:4096])  #sends the data in parts
                    data = data[4096:]  #and removes it from data
                client_socket_screen.send("The file transmittion has ended.")
                #and the server recieves untill this message is sent
        except socket.error:
            client_socket_screen.close()
            print "Client's screen output Disconnected!"
            break
    try:
        client_socket_screen.close()
        client_socket_screen = socket.socket()
        print "Client's keyboard IS Disconnected!\n"
    finally:
        pass


def handle_keyboard_input():
    '''
    this function connects to the server with the keyboard's socket, and changes the keyboard's button states if they
    were changed on the server side
    '''
    global check_stop
    global client_socket_keyboard
    while check_stop is False:  #an infinite loop
        try:
            client_socket_keyboard.connect((CLIENT_IP, CLIENT_PORT_KEYBOARD))
            #connects between the keyboard socket and the following address
            print 'A new keyboard client connection established!'
            while check_stop is False:  #an infinite loop
                ln = client_socket_keyboard.recv(2)  #the length of the data that is going to be recieved
                if ln.isdigit():  #checks if the length is only digits, and wasn't mixed with letters
                    s = str(client_socket_keyboard.recv(int(ln))).split(",")  #the key's new data
                    if s and len(s) > 0:
                        if s[0] == 'key down':  #checks if the key is pressed down on the server side
                            press_key(int(s[1]))  #presses it if it was
                        else:
                            release_key(int(s[1]))  #otherwise, releases it
        except socket.error:
            client_socket_keyboard.close()
            print "Client's keyboard Disconnected!"
            break
    try:
        client_socket_keyboard.close()
        client_socket_keyboard = socket.socket()
        print "Client's keyboard IS Disconnected!\n"
    finally:
        pass


def run_IO_threads():
    """
    running all I/O (inout output) threads
    """

    screen_input_thread = threading.Thread(name='screen_input', target=handle_screen_input, args=())
    screen_output_thread = threading.Thread(name='screen_output', target=handle_screen_output, args=())
    mouse_thread = threading.Thread(name='mouse', target=handle_mouse, args=())
    keyboard_thread = threading.Thread(name='keyboard', target=handle_keyboard_input, args=())

    screen_input_thread.start()
    screen_output_thread.start()
    mouse_thread.start()
    keyboard_thread.start()

    #mouse_thread.setDaemon(True)

    #screen_input_thread.join()
    #screen_output_thread.join()


def help_commands():
    '''This function presents the RCPC commands'''
    print '\nHello, this is the RCPC program commands!.\n' \
          '*Press enter to clear your command.\n*Type help to get the client commands.' \
          '\n*Type "exit" to close the RCPC program!'


def main():
    '''
    activates the threads, and by that, activates the program
    '''
    print 'Starting the RCPC Client!'
    #instead of the time.sleep, to close the threads if the client disconnects
    run_IO_threads()
    cmd = ''
    global check_stop
    help_commands()
    try:
        client_socket_disconnect.connect((CLIENT_IP, CLIENT_PORT_DISCONNECT))
    finally:
        pass
    cmd = ''
    sys.stdout.write('\n')
    sys.stdout.write('\nYour Command:')
    while True:  #to keep the main func running
        if msvcrt.kbhit():
            ch = msvcrt.getch()
            cmd += ch
            sys.stdout.write(ch)
            if cmd == 'exit':
                sys.stdout.write('\nClosing The RCPC...')
                client_socket_disconnect.send('EXIT')
                check_stop = True
                time.sleep(1)
                os._exit(0)
            if cmd == 'help':
                help_commands()
                cmd = ''
                sys.stdout.write('\n')
                sys.stdout.write('Your Command:')
            if ch == '\n' or ch == '\r':
                cmd = ''
                sys.stdout.write('\n')
                sys.stdout.write('Your Command:')


if __name__ == '__main__':
    main()