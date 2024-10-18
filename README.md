# Python Remote-PC
![DALL·E 2024-10-18 14 59 09 - A split-screen image showing on the left side a human hand moving a mouse on a table, and on the right side a computer screen showing a cursor moving ](https://github.com/user-attachments/assets/c7881a2f-0a42-4444-bd85-c7114f98472a)


This project is a real-time program which displays the client’s screen and enables control of his keyboard and mouse functionality (where a technician’s machine acts as a server).
The backend is written in Python, using socket library for client -server communication and ctypes, Windll and Win32Api libraries for event handling and multithreading, and For the GUI (frontend) I used OpenCV.



## Features

- Event-handling
- Multithreading
- Input Encapsulation

## Installation

### Steps:
###### 1. Clone the repository.
###### 2. Navigate to the project directory.
###### 3. Download the required libraries:

```bash
  pip install win32api ctypes windll socket threading time StringIO msvcrt sys os mss PIL atexit win32gui collections win32con
```
