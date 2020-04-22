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
def mouse_move(left: int, top: int) -> None:
    ctypes.windll.user32.SetCursorPos(left, top)
    time.sleep(.2)


# Perform a left click with the mouse
def mouse_left_click() -> None:
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    time.sleep(.2)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)


# Find a window in the windows array by its title
def find_window_by_title(search_title: str) -> dict:
    # Reset top windows array
    top_windows = []

    # Call window enumeration handler
    win32gui.EnumWindows(window_enumeration_handler, top_windows)
    found_window = None
    for window in top_windows:
        if search_title in window['title']:
            found_window = window

    return found_window


def close_window(handle) -> None:
    win32gui.PostMessage(handle, win32con.WM_CLOSE, 0, 0)


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
def ocr_screenshot_region(left: int, top: int, width: int, height: int, invert: bool = False, show: bool = False,
                          config: str = r'--oem 3 --psm 7') -> str:
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    if invert:
        screenshot = ImageOps.invert(screenshot)
    if show:
        screenshot.show()
    ocr_result = pytesseract.image_to_string(screenshot, config=config)
    # print_log(f'OCR result: {ocr_result}')
    return ocr_result.lower()


# Check if an error message is present in the game
def in_game_error_message_present(left: int, top: int) -> bool:
    ocr_result = ocr_screenshot_region(
        left + 590,
        top + 330,
        115,
        34,
        True,
        False,
        r'--oem 3 --psm 7'
    )

    return 'error' in ocr_result or 'notice' in ocr_result


# Close an error message in the game
def close_in_game_error_message(left: int, top: int) -> None:
    mouse_move(left + 648, top + 400)
    mouse_left_click()


def blizzard_error_message_present(left: int, top: int) -> bool:
    ocr_result = ocr_screenshot_region(
        left + 233,
        top + 258,
        365,
        30,
        True
    )

    return 'server disconnected' in ocr_result or 'connection failed' in ocr_result


def blank_screen_present(left: int, top: int) -> bool:
    screenshot = pyautogui.screenshot(region=(left + 30, top + 50, 50, 50))
    colors = screenshot.getcolors()

    # Check if only one color is present and if that color is black
    return colors is not None and len(colors) == 1 and colors[0][1] == (0, 0, 0)


def launch_game_instance(install_path: str) -> bool:
    # Use launcher to open game in client
    print_log('Running launcher')
    subprocess.run(os.path.join(install_path, 'Modern Warfare Launcher.exe'))
    print_log('Waiting for client window')
    time.sleep(10)

    # Get a handle for the client window
    print_log('Getting client window details')
    client_window = find_window_by_title('Blizzard Battle.net')

    try:
        print_log('Bringing client window to foreground')
        win32gui.ShowWindow(client_window['handle'], win32con.SW_SHOW)
        win32gui.SetForegroundWindow(client_window['handle'])
        time.sleep(1)
    except Exception as e:
        print_log(str(e))
        print_log('Error in handling client window')
        return False

    # Look for client's "Games"-tab
    print_log('Waiting for client to finish startup')
    attempts = 0
    max_attempts = 30
    games_tab_present = False
    while not games_tab_present and attempts < max_attempts:
        games_tab_present = 'games' in ocr_screenshot_region(
            client_window['rect'][0] + 137,
            client_window['rect'][1] + 40,
            80,
            22,
            True
        )
        attempts += 1
        time.sleep(2)

    if not games_tab_present:
        print_log('Looks like client did not start correctly')
        return False

    # >>Click<< "Play"-button by hitting enter
    print_log('Hitting enter to trigger "Play"-button')
    auto_press_key(0x1c)
    print_log('Waiting for game window')
    time.sleep(20)

    # Get a handle for the game window
    print_log('Getting game window details')
    game_window = find_window_by_title('Call of Duty®: Modern Warfare®')

    # Bring game window to foreground and scale it to 720p
    try:
        print_log('Bringing game window to foreground')
        win32gui.ShowWindow(game_window['handle'], win32con.SW_SHOW)
        win32gui.SetForegroundWindow(game_window['handle'])
        print_log('Scaling game window to 1280x720')
        win32gui.MoveWindow(game_window['handle'], 10, 30, 1296, 759, True)
        time.sleep(1)
    except Exception as e:
        print_log(str(e))
        print_log('Error in handling game window')
        return False

    # Update game window rect
    game_window['rect'] = win32gui.GetWindowRect(game_window['handle'])

    # Move cursor to center of game window to prepare Warzone selection
    mouse_move(game_window['rect'][0] + 640, game_window['rect'][1] + 360)

    # Wait until game has fully started
    print_log('Waiting for game to finish startup')
    attempts = 0
    max_attempts = 30
    game_title_present = False
    while not game_title_present and attempts < max_attempts:
        ocr_result = ocr_screenshot_region(
            game_window['rect'][0] + 420,
            game_window['rect'][1] + 102,
            455,
            106,
            True,
            False,
            r'--oem 3 --psm 8'
        )
        game_title_present = 'war' in ocr_result and 'zone' in ocr_result
        attempts += 1
        time.sleep(2)

    if not game_title_present:
        print_log('Looks like game did not start correctly')
        return False

    # Hit space once to close promotional message
    print_log('Hitting space to close promo message')
    auto_press_key(0x39)
    time.sleep(2)

    # Hit space again to choose Warzone
    print_log('Hitting space to select Warzone')
    auto_press_key(0x39)
    time.sleep(5)

    return 'warzone' in ocr_screenshot_region(
        game_window['rect'][0] + 64,
        game_window['rect'][1] + 64,
        160,
        30,
        True,
    )


# Set up argument parsing and parse
parser = argparse.ArgumentParser(description='Control a Call of Duty®: Modern Warfare®'
                                             'game instance to spectate players playing Warzone')
parser.add_argument('--game-path', help='Path to game install folder',
                    type=str, default=r'C:\Program Files (x86)\Call of Duty Modern Warfare')
parser.add_argument('--tesseract-path', help='Path to Tesseract install folder',
                    type=str, default=r'C:\Program Files\Tesseract-OCR')
args = parser.parse_args()

# Init global vars/settings
pytesseract.pytesseract.tesseract_cmd = os.path.join(args.tesseract_path, 'tesseract.exe')
top_windows = []

# Make sure provided Tesseract path is valid
if not os.path.isfile(pytesseract.pytesseract.tesseract_cmd):
    sys.exit(f'Could not find tesseract.exe in given install folder: {args.tesseract_path}')
elif not os.path.isfile(os.path.join(args.game_path, 'Modern Warfare Launcher.exe')):
    sys.exit(f'Could not find Modern Warfare Launcher.exe in given game install folder: {args.game_path}')

# Find game window
print_log('Finding Warzone window')
gameWindow = find_window_by_title('Call of Duty®: Modern Warfare®')
print_log(f'Found window: {gameWindow}')

# If no game window was found, start a new game instance
if gameWindow is None:
    print_log('No game window found, starting a new game instance')
    gameLaunched = launch_game_instance(args.game_path)

    print_log(f'gameLaunched: {gameLaunched}')

    if gameLaunched:
        gameWindow = find_window_by_title('Call of Duty®: Modern Warfare®')

# Make sure we now have a game window
if gameWindow is None:
    sys.exit("Failed to find/start game instance")

# Spectate indefinitely
restartRequired = False
while True:
    if restartRequired:
        # If window still exists but is showing error message, close it
        try:
            # Close existing game instance
            print_log('Closing existing game instance')
            close_window(gameWindow['handle'])
            time.sleep(10)
        except Exception as e:
            print_log('Failed to close game, restarting anyways')

        # Start new game instance
        print_log('Starting new game instance')
        gameLaunched = launch_game_instance(args.game_path)
        print_log(f'gameLaunched: {gameLaunched}')

        if gameLaunched:
            print_log('Updating window details')
            gameWindow = find_window_by_title('Call of Duty®: Modern Warfare®')
            restartRequired = False
        else:
            print_log('Failed to restart game, trying again in 60s')
            time.sleep(60)
            continue

    try:
        win32gui.ShowWindow(gameWindow['handle'], win32con.SW_SHOW)
        win32gui.SetForegroundWindow(gameWindow['handle'])
        time.sleep(1)
    except Exception as e:
        print_log(str(e))
        print_log('Error in handling game window, restarting game')
        restartRequired = True
        continue

    try:
        print_log('Scaling game window to 1280x720')
        win32gui.MoveWindow(gameWindow['handle'], 10, 30, 1296, 759, True)
        time.sleep(1)
    except Exception as e:
        print_log(str(e))
        print_log('Error in handling game window')
        restartRequired = True
        continue

    # Initial check for error message
    if in_game_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1]):
        print_log('In game error message present, closing it and starting over')
        close_in_game_error_message(gameWindow['rect'][0], gameWindow['rect'][1])
        continue
    elif blizzard_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1]):
        print_log('Blizzard error message present, restarting game')
        restartRequired = True
        continue
    elif blank_screen_present(gameWindow['rect'][0], gameWindow['rect'][1]):
        print_log('Blank screen is present, restarting game')
        restartRequired = True
        continue

    # Click battle royal
    print_log('Clicking "Battle Royale"-option')
    mouse_move(gameWindow['rect'][0] + 233, gameWindow['rect'][1] + 210)
    mouse_left_click()

    time.sleep(30)

    # Bring window back to front (to be sure, and to enable alt-tabbing between controller actions)
    try:
        win32gui.ShowWindow(gameWindow['handle'], win32con.SW_SHOW)
        win32gui.SetForegroundWindow(gameWindow['handle'])
        time.sleep(1)
    except Exception as e:
        print_log(str(e))
        print_log('Error in handling game window, restarting game')
        restartRequired = True
        continue

    # Wait for pre-game to start
    inPreGame = False
    inGameErrorMessagePresent = False
    blizzardErrorMessagePresent = False
    blankScreenCounter = 0
    blankScreenLimit = 5
    while not inPreGame and not inGameErrorMessagePresent and \
            not blizzardErrorMessagePresent and blankScreenCounter < blankScreenLimit:
        # Check if we are in pre-game already
        ocrResult = ocr_screenshot_region(
            gameWindow['rect'][0] + 525,
            gameWindow['rect'][1] + 106,
            250,
            18,
            True
        )

        inPreGame = 'waiting' in ocrResult and 'lobby' in ocrResult or \
                    'match' in ocrResult and 'full' in ocrResult or \
                    'deployment' in ocrResult and 'begin' in ocrResult
        print_log(f'inPreGame: {inPreGame}')

        # Check for any error messages
        inGameErrorMessagePresent = in_game_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])
        blizzardErrorMessagePresent = blizzard_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])

        if blank_screen_present(gameWindow['rect'][0], gameWindow['rect'][1]):
            blankScreenCounter += 1

        time.sleep(8)

    # If an error game message is present, close it and start over
    if inGameErrorMessagePresent:
        print_log('Error message is present, closing it and starting over')
        close_in_game_error_message(gameWindow['rect'][0], gameWindow['rect'][1])
        continue

    # If blizzard error message is present, close game and start a new instance
    if blizzardErrorMessagePresent:
        print_log('Blizzard error message present, restarting game')
        restartRequired = True
        continue

    # If blank screen limit has been reached, close game and start a new instance
    if blankScreenCounter >= blankScreenLimit:
        print_log('Blank screen limit reached, restarting game')
        restartRequired = True
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
        print_log('Error in handling game window, exiting')
        restartRequired = True
        continue

    print_log('Game should start any second, starting ocr attempts')

    # Wait for jump button indicator to appear
    canJump = False
    while not canJump and not inGameErrorMessagePresent and \
            not blizzardErrorMessagePresent and blankScreenCounter < blankScreenLimit:
        canJump = 'space' in ocr_screenshot_region(
            gameWindow['rect'][0] + 558,
            gameWindow['rect'][1] + 624,
            40,
            15,
            False,
            False,
            r'--oem 3 --psm 8'
        )
        print_log(f'canJump: {canJump}')

        # Check for any error messages
        inGameErrorMessagePresent = in_game_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])
        blizzardErrorMessagePresent = blizzard_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])

        if blank_screen_present(gameWindow['rect'][0], gameWindow['rect'][1]):
            blankScreenCounter += 1

        time.sleep(2)

    # If an error message is present, close it and start over
    if inGameErrorMessagePresent:
        print_log('Error message is present, closing it and starting over')
        close_in_game_error_message(gameWindow['rect'][0], gameWindow['rect'][1])
        continue

    # If blizzard error message is present, close game and start a new instance
    if blizzardErrorMessagePresent:
        print_log('Blizzard error message present, restarting game')
        restartRequired = True
        continue

    # If blank screen limit has been reached, close game and start a new instance
    if blankScreenCounter >= blankScreenLimit:
        print_log('Blank screen limit reached, restarting game')
        restartRequired = True
        continue

    time.sleep(5)
    print_log('Game started, jumping')

    auto_press_key(0x39)
    time.sleep(5)
    print_log('Deploying parachute')
    auto_press_key(0x39)
    time.sleep(4)
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
        print_log('Error in handling game window, restarting game')
        restartRequired = True
        continue

    # Look for "Spectate"-button
    spectateButtonPresent = False
    while not spectateButtonPresent and not inGameErrorMessagePresent and \
            not blizzardErrorMessagePresent and blankScreenCounter < blankScreenLimit:
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
        inGameErrorMessagePresent = in_game_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])
        blizzardErrorMessagePresent = blizzard_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])

        if blank_screen_present(gameWindow['rect'][0], gameWindow['rect'][1]):
            blankScreenCounter += 1

        if not spectateButtonPresent:
            time.sleep(4)

    # If an error message is present, close it and start over
    if inGameErrorMessagePresent:
        print_log('Error message is present, closing it and starting over')
        close_in_game_error_message(gameWindow['rect'][0], gameWindow['rect'][1])
        continue

    # If blizzard error message is present, close game and start a new instance
    if blizzardErrorMessagePresent:
        print_log('Blizzard error message present, restarting game')
        restartRequired = True
        continue

    # If blank screen limit has been reached, close game and start a new instance
    if blankScreenCounter >= blankScreenLimit:
        print_log('Blank screen limit reached, restarting game')
        restartRequired = True
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
    windowError = False
    iterationsOnPlayer = 0
    while not onInMemoriam and not leaveGameButtonPresent and \
            not inGameErrorMessagePresent and not blizzardErrorMessagePresent and \
            not windowError and blankScreenCounter < blankScreenLimit:
        # Check if "In Memoriam" title is present
        onInMemoriam = 'in memo' in ocr_screenshot_region(
            gameWindow['rect'][0] + 138,
            gameWindow['rect'][1] + 120,
            98,
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
        inGameErrorMessagePresent = in_game_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])
        blizzardErrorMessagePresent = blizzard_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1])

        if blank_screen_present(gameWindow['rect'][0], gameWindow['rect'][1]):
            blankScreenCounter += 1

        # Switch to next player if current player has reached iteration limit
        # and game is not over
        if iterationsOnPlayer >= 10 and not onInMemoriam and not leaveGameButtonPresent and \
                not inGameErrorMessagePresent and not blizzardErrorMessagePresent \
                and blankScreenCounter < blankScreenLimit:
            # Bring window back to front (to be sure, and to enable alt-tabbing between controller actions)
            try:
                win32gui.ShowWindow(gameWindow['handle'], win32con.SW_SHOW)
                win32gui.SetForegroundWindow(gameWindow['handle'])
                time.sleep(1)
            except Exception as e:
                print_log(str(e))
                windowError = True
                continue

            # Click to rotate
            print_log('Rotating')
            mouse_left_click()

            # Reset counter
            iterationsOnPlayer = 0

        iterationsOnPlayer += 1
        time.sleep(2)

    # If an error message is present, close it and start over
    if inGameErrorMessagePresent:
        print_log('Error message is present, closing it and starting over')
        close_in_game_error_message(gameWindow['rect'][0], gameWindow['rect'][1])
        continue

    # If blizzard error message is present, close game and start a new instance
    if blizzardErrorMessagePresent:
        print_log('Blizzard error message present, restarting game')
        restartRequired = True
        continue

    # If there was an error handling the game window, close game an start a new instance
    if windowError:
        print_log('Error in handling game window, restarting game')
        restartRequired = True
        continue

    # If blank screen limit has been reached, close game and start a new instance
    if blankScreenCounter >= blankScreenLimit:
        print_log('Blank screen limit reached, restarting game')
        restartRequired = True
        continue

    if onInMemoriam:
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
