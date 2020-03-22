import argparse
import ctypes
import os
import subprocess
import sys
import time

import pyautogui
import pytesseract
import win32con
import win32gui
from PIL import ImageOps

SendInput = ctypes.windll.user32.SendInput

# C struct redefinitions
PUL = ctypes.POINTER(ctypes.c_ulong)


class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]


class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class InputI(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                ("mi", MouseInput),
                ("hi", HardwareInput)]


class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", InputI)]


def press_key(hex_key_code):
    extra = ctypes.c_ulong(0)
    ii_ = InputI()
    ii_.ki = KeyBdInput(0, hex_key_code, 0x0008, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def release_key(hex_key_code):
    extra = ctypes.c_ulong(0)
    ii_ = InputI()
    ii_.ki = KeyBdInput(0, hex_key_code, 0x0008 | 0x0002, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def auto_press_key(hey_key_code):
    press_key(hey_key_code)
    time.sleep(.08)
    release_key(hey_key_code)


def window_enumeration_handler(hwnd, top_windows):
    """Add window title, ID and rect to array."""
    top_windows.append({
        'handle': hwnd,
        'title': win32gui.GetWindowText(hwnd),
        'rect': win32gui.GetWindowRect(hwnd),
    })


# Print a line preceded by a timestamp
def print_log(message: object) -> None:
    print(f'{time.strftime("%Y-%m-%d %H:%M:%S")} # {str(message)}')


# Move mouse cursor to provided x and y coordinates
def mouse_move(x: int, y: int) -> None:
    ctypes.windll.user32.SetCursorPos(x, y)
    time.sleep(.2)


# Perform a left click with the mouse
def mouse_left_click() -> None:
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    time.sleep(.2)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)


# Find a window in the windows array by its title
def find_window_by_title(search_title: str) -> dict:
    # Call window enumeration handler
    win32gui.EnumWindows(window_enumeration_handler, top_windows)
    found_window = None
    for window in top_windows:
        if search_title in window['title']:
            found_window = window

    return found_window


# Check if a process with the given PID is running/responding
def is_responding_pid(pid: int) -> bool:
    """Check if a program (based on its PID) is responding"""
    cmd = 'tasklist /FI "PID eq %d" /FI "STATUS eq running"' % pid
    status = subprocess.Popen(cmd, stdout=subprocess.PIPE).stdout.read()
    return str(pid) in str(status)


# Forcefully terminate a process by its PID
def taskkill_pid(pid: int) -> bool:
    cmd = 'taskkill /F /PID %d' % pid
    output = subprocess.Popen(cmd, stdout=subprocess.PIPE).stdout.read()
    return 'has been terminated' in str(output)


# Take a screenshot of the given region and run the result through OCR
def ocr_screenshot_region(x: int, y: int, w: int, h: int, invert: bool = False, show: bool = False,
                          config: str = r'--oem 3 --psm 7') -> str:
    screenshot = pyautogui.screenshot(region=(x, y, w, h))
    if invert:
        screenshot = ImageOps.invert(screenshot)
    if show:
        screenshot.show()
    ocr_result = pytesseract.image_to_string(screenshot, config=config)
    # print_log(f'OCR result: {ocr_result}')
    return ocr_result.lower()


# Check if an error message is present in the game
def error_message_present(top: int, left: int) -> bool:
    return 'error' in ocr_screenshot_region(
        top + 622,
        left + 340,
        55,
        15,
        True,
        False,
        r'--oem 3 --psm 8'
    )


# Close an error message in the game
def close_error_message(top: int, left: int) -> None:
    mouse_move(top + 648, left + 410)
    mouse_left_click()


# Set up argument parsing and parse
parser = argparse.ArgumentParser(description='Control a Call of Duty®: Modern Warfare®'
                                             'game instance to spectate players playing Warzone')
parser.add_argument('--tesseract-path', help='Path to Tesseract install folder',
                    type=str, default='C:\\Program Files\\Tesseract-OCR\\')
args = parser.parse_args()

# Init global vars/settings
pytesseract.pytesseract.tesseract_cmd = os.path.join(args.tesseract_path, 'tesseract.exe')
top_windows = []

# Make sure provided Tesseract path is valid
if not os.path.isfile(pytesseract.pytesseract.tesseract_cmd):
    sys.exit(f'Could not find tesseract.exe in given install folder: {pytesseract.pytesseract.tesseract_cmd}')

# Init game instance
# TODO

# Find game window
print_log('Finding Warzone window')
gameWindow = find_window_by_title('Call of Duty®: Modern Warfare®')
print_log(f'Found window: {gameWindow}')

# Make sure a game window was found
if gameWindow is None:
    sys.exit('No game window found, please start an instance of the game first')

# Bring game window to foreground and scale it to 720p
try:
    win32gui.ShowWindow(gameWindow['handle'], win32con.SW_SHOW)
    win32gui.SetForegroundWindow(gameWindow['handle'])
    win32gui.MoveWindow(gameWindow['handle'], 50, 50, 1296, 759, True)
    time.sleep(1)
except Exception as e:
    print_log(str(e))
    sys.exit('Error in handling game window, exiting')

# Spectate indefinitely
while True:
    try:
        win32gui.ShowWindow(gameWindow['handle'], win32con.SW_SHOW)
        win32gui.SetForegroundWindow(gameWindow['handle'])
        time.sleep(1)
    except Exception as e:
        print_log(str(e))
        sys.exit('Error in handling game window, exiting')

    # Close initial pop up
    mouse_move(gameWindow['rect'][0] + 569, gameWindow['rect'][1] + 200)
    mouse_left_click()

    time.sleep(1)

    # Click battle royal
    mouse_move(gameWindow['rect'][0] + 233, gameWindow['rect'][1] + 342)
    mouse_left_click()

    time.sleep(30)

    # Bring window back to front (to be sure, and to enable alt-tabbing between controller actions)
    try:
        win32gui.ShowWindow(gameWindow['handle'], win32con.SW_SHOW)
        win32gui.SetForegroundWindow(gameWindow['handle'])
        time.sleep(1)
    except Exception as e:
        print_log(str(e))
        sys.exit('Error in handling game window, exiting')

    # Wait for pre-game to start
    inPreGame = False
    errorMessagePresent = False
    while not inPreGame and not errorMessagePresent:
        # Check if we are in pre-game already
        ocrResult = ocr_screenshot_region(
            gameWindow['rect'][0] + 525,
            gameWindow['rect'][1] + 106,
            250,
            18
        )

        inPreGame = 'waiting' in ocrResult and 'lobby' in ocrResult or \
                    'match' in ocrResult and 'full' in ocrResult
        print_log(f'inPreGame: {inPreGame}')

        # Check for any error messages
        errorMessagePresent = error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])

        time.sleep(8)

    # If an error message is present, close it and start over
    if errorMessagePresent:
        print_log('Error message is present, closing it and starting over')
        close_error_message(gameWindow['rect'][0], gameWindow['rect'][1])
        continue

    print_log('Entered pre-game without errors, awaiting game start')
    time.sleep(70)

    # Bring window back to front (to be sure, and to enable alt-tabbing between controller actions)
    try:
        win32gui.ShowWindow(gameWindow['handle'], win32con.SW_SHOW)
        win32gui.SetForegroundWindow(gameWindow['handle'])
        time.sleep(1)
    except Exception as e:
        print_log(str(e))
        sys.exit('Error in handling game window, exiting')

    print_log('Game should start any second, starting ocr attempts')

    # Wait for jump button indicator to appear
    canJump = False
    while not canJump and not errorMessagePresent:
        canJump = 'space' in ocr_screenshot_region(
            gameWindow['rect'][0] + 558,
            gameWindow['rect'][1] + 662,
            40,
            15,
            False,
            False,
            r'--oem 3 --psm 8'
        )
        print_log(f'canJump: {canJump}')

        # Check for any error messages
        errorMessagePresent = error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])

        time.sleep(2)

    # If an error message is present, close it and start over
    if errorMessagePresent:
        print_log('Error message is present, closing it and starting over')
        close_error_message(gameWindow['rect'][0], gameWindow['rect'][1])
        continue

    time.sleep(1.5)
    print_log('Game started, jumping')

    auto_press_key(0x39)
    time.sleep(4)
    print_log('Deploying parachute')
    auto_press_key(0x39)
    time.sleep(3)
    print_log('Cutting parachute')
    auto_press_key(0x2e)
    time.sleep(2)

    # Wait to fall to ground
    print_log('Holding w to fall faster/forward')
    press_key(0x11)
    time.sleep(11)
    release_key(0x11)

    # Wait for gulag to pass
    print_log('Waiting for gulag to pass')
    time.sleep(20)

    # Attempt to suicide by arming and holding grenade
    # TODO

    # Bring window back to front (to be sure, and to enable alt-tabbing between controller actions)
    try:
        win32gui.ShowWindow(gameWindow['handle'], win32con.SW_SHOW)
        win32gui.SetForegroundWindow(gameWindow['handle'])
        time.sleep(1)
    except Exception as e:
        print_log(str(e))
        sys.exit('Error in handling game window, exiting')

    # Look for "Spectate"-button
    spectateButtonPresent = False
    while not spectateButtonPresent and not errorMessagePresent:
        spectateButtonPresent = 'spectate' in ocr_screenshot_region(
            gameWindow['rect'][0] + 994,
            gameWindow['rect'][1] + 392,
            60,
            16,
            True,
            False,
            r'--oem 3 --psm 8'
        )
        print_log(f'spectateButtonPresent: {spectateButtonPresent}')

        # Check for any error messages
        errorMessagePresent = error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])

        if not spectateButtonPresent:
            time.sleep(4)

    # If an error message is present, close it and start over
    if errorMessagePresent:
        print_log('Error message is present, closing it and starting over')
        close_error_message(gameWindow['rect'][0], gameWindow['rect'][1])
        continue

    # Click spectate
    print_log('Clicking "Spectate"-button')
    mouse_move(gameWindow['rect'][0] + 1043, gameWindow['rect'][1] + 402)
    mouse_left_click()

    # Toggle through players while game is running
    print_log('Entering player spectate rotation')

    # Move cursor to center of screen to avoid accidentally clicking buttons
    mouse_move(gameWindow['rect'][0] + 640, gameWindow['rect'][1] + 360)

    # Spectate first player for 22 seconds
    time.sleep(22)

    onInMemoriam = False
    leaveGameButtonPresent = False
    iterationsOnPlayer = 0
    while not onInMemoriam and not leaveGameButtonPresent:
        # Bring window back to front (to be sure, and to enable alt-tabbing between controller actions)
        try:
            win32gui.ShowWindow(gameWindow['handle'], win32con.SW_SHOW)
            win32gui.SetForegroundWindow(gameWindow['handle'])
            time.sleep(1)
        except Exception as e:
            print_log(str(e))
            sys.exit('Error in handling game window, exiting')

        # Check if "In Memoriam" title is present
        onInMemoriam = 'in memo' in ocr_screenshot_region(
            gameWindow['rect'][0] + 134,
            gameWindow['rect'][1] + 120,
            102,
            14,
            False,
            False
        )

        # Check if "Leave Game"-button is present
        leaveGameButtonPresent = 'leave' in ocr_screenshot_region(
            gameWindow['rect'][0] + 994,
            gameWindow['rect'][1] + 392,
            60,
            16,
            True,
            False,
            r'--oem 3 --psm 8'
        )
        print_log(f'onInMemoriam: {onInMemoriam}; leaveGameButtonPresent: {leaveGameButtonPresent}')

        # Check for any error messages
        errorMessagePresent = error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])

        # Switch to next player if current player has reached iteration limit
        # and game is not over
        if iterationsOnPlayer >= 10 and not onInMemoriam and not leaveGameButtonPresent and not errorMessagePresent:
            # Click to rotate
            print_log('Rotating')
            mouse_left_click()

            # Reset counter
            iterationsOnPlayer = 0
        elif errorMessagePresent:
            # An error message is present, break spectate loop and start over
            break

        iterationsOnPlayer += 1
        time.sleep(2)

    # If an error message is present, close it and start over
    if errorMessagePresent:
        print_log('Error message is present, closing it and starting over')
        close_error_message(gameWindow['rect'][0], gameWindow['rect'][1])
        continue

    # Hit ESC
    print_log('Game over, hitting ESC')
    auto_press_key(0x01)
    time.sleep(1)

    # Click "Leave game"-button
    print_log('Clicking "Leave game"-button')
    mouse_move(gameWindow['rect'][0] + 150, gameWindow['rect'][1] + 286)
    mouse_left_click()
    time.sleep(1)

    # Confirm
    print_log('Confirming "Leave game"-dialogue')
    mouse_move(gameWindow['rect'][0] + 643, gameWindow['rect'][1] + 386)
    mouse_left_click()
    time.sleep(5)
