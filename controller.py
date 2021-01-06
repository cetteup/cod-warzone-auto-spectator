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

from gameinstancestate import GameInstanceState

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
        'title': win32gui.GetWindowText(hwnd).replace('\u200b', ''),
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
    ocr_result = pytesseract.image_to_string(screenshot, config=config).strip(' \n\x0c')
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


def main_menu_present(left: int, top: int) -> bool:
    return 'warzone' in ocr_screenshot_region(
        left + 64,
        top + 58,
        160,
        38,
        True
    )


def spectate_button_present(left: int, top: int) -> bool:
    return 'spectate' in ocr_screenshot_region(
        left + 974,
        top + 479,
        60,
        35,
        True
    )


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
            game_window['rect'][0] + 475,
            game_window['rect'][1] + 222,
            355,
            84,
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

    return main_menu_present(game_window['rect'][0], game_window['rect'][1])


# Set up argument parsing and parse
parser = argparse.ArgumentParser(description='Control a Call of Duty®: Modern Warfare®'
                                             'game instance to spectate players playing Warzone')
parser.add_argument('--version', action='version', version='cod-warzone-auto-spectator v0.1.6')
parser.add_argument('--game-path', help='Path to game install folder',
                    type=str, default=r'C:\Program Files (x86)\Call of Duty Modern Warfare')
parser.add_argument('--tesseract-path', help='Path to Tesseract install folder',
                    type=str, default=r'C:\Program Files\Tesseract-OCR')
parser.add_argument('--blank-screen-limit', help='How many times a (mostly) blank/black screen can be detected before'
                                                 'the game is restarted', type=int, default=5)
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

gameInstanceState = GameInstanceState()

# If no game window was found, start a new game instance
if gameWindow is None:
    print_log('No game window found, starting a new game instance')
    gameLaunched = launch_game_instance(args.game_path)

    print_log(f'gameLaunched: {gameLaunched}')

    if gameLaunched:
        gameWindow = find_window_by_title('Call of Duty®: Modern Warfare®')
else:
    try:
        print_log('Scaling game window to 1280x720')
        win32gui.MoveWindow(gameWindow['handle'], 10, 30, 1296, 759, True)
        time.sleep(1)
    except Exception as e:
        print_log(str(e))
        print_log('Error in handling game window')
        gameInstanceState.set_error_restart_required(True)

# Make sure we now have a game window
if gameWindow is None:
    sys.exit("Failed to find/start game instance")

# Spectate indefinitely
iterationsOnPlayer = 0
while True:
    # Try to bring game window to foreground
    try:
        win32gui.ShowWindow(gameWindow['handle'], win32con.SW_SHOW)
        win32gui.SetForegroundWindow(gameWindow['handle'])
        time.sleep(1)
    except Exception as e:
        print_log(str(e))
        print_log('Error in handling game window, restarting game')
        gameInstanceState.set_error_restart_required(True)

    if gameInstanceState.error_restart_required():
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
            gameInstanceState.set_error_restart_required(False)
        else:
            print_log('Failed to restart game, trying again in 60s')
            time.sleep(60)
            continue

    # Initial check for error message
    if in_game_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1]):
        print_log('In game error message present, closing it and starting over')
        close_in_game_error_message(gameWindow['rect'][0], gameWindow['rect'][1])
        # TODO: Add handling of GameInstanceState
        continue
    elif blizzard_error_message_present(gameWindow['rect'][0], gameWindow['rect'][1]):
        print_log('Blizzard error message present, restarting game')
        gameInstanceState.set_error_restart_required(True)
        continue
    elif gameInstanceState.get_error_blank_screen_count() >= args.blank_screen_limit:
        print_log('Blank screen limit reached, restarting game')
        gameInstanceState.set_error_restart_required(True)
        continue
    elif blank_screen_present(gameWindow['rect'][0], gameWindow['rect'][1]):
        gameInstanceState.increase_error_blank_screen_count()
        time.sleep(2)
    elif gameInstanceState.get_error_blank_screen_count() > 0:
        gameInstanceState.reset_error_blank_screen_count()

    if main_menu_present(gameWindow['rect'][0], gameWindow['rect'][1]) and not gameInstanceState.round_rotation_started():
        # Move mouse battle royal
        print_log('Clicking "Battle Royale"-option')
        mouse_move(gameWindow['rect'][0] + 233, gameWindow['rect'][1] + 210)

        # Click on quads
        mouse_move(gameWindow['rect'][0] + 442, gameWindow['rect'][1] + 183)
        mouse_left_click()

        gameInstanceState.set_round_rotation_started(True)
        gameInstanceState.set_searching_for_game(True)

        time.sleep(30)
    elif gameInstanceState.searching_for_game():
        # Check if we are in pre-game already
        ocrResult = ocr_screenshot_region(
            gameWindow['rect'][0] + 525,
            gameWindow['rect'][1] + 106,
            250,
            18,
            True
        )

        gameInstanceState.set_in_pre_game('waiting' in ocrResult and 'lobby' in ocrResult or
                                          'match' in ocrResult and 'full' in ocrResult or
                                          'deployment' in ocrResult and 'begin' in ocrResult)
        if gameInstanceState.in_pre_game():
            gameInstanceState.set_searching_for_game(False)
            print_log('Entered pre-game without errors, awaiting game start')
            time.sleep(30)
            print_log('Game could start any second, starting to check of jump prompt')
        else:
            time.sleep(8)

        print_log(f'inPreGame: {gameInstanceState.in_pre_game()}')
    elif gameInstanceState.in_pre_game():
        canJump = 'space' in ocr_screenshot_region(
            gameWindow['rect'][0] + 568,
            gameWindow['rect'][1] + 624,
            40,
            15,
            False,
            False,
            r'--oem 3 --psm 8'
        )
        print_log(f'canJump: {canJump}')

        if canJump:
            gameInstanceState.set_in_pre_game(False)

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
        else:
            time.sleep(2)
    elif not gameInstanceState.round_started_spectation():
        # Look for "Spectate"-button
        spectateButtonPresent = spectate_button_present(gameWindow['rect'][0], gameWindow['rect'][1])
        print_log(f'spectateButtonPresent: {spectateButtonPresent}')

        if spectateButtonPresent:
            print_log('Clicking "Spectate"-button')
            mouse_move(gameWindow['rect'][0] + 1043, gameWindow['rect'][1] + 494)
            mouse_left_click()

            time.sleep(.5)

            gameInstanceState.set_round_started_spectation(
                not spectate_button_present(gameWindow['rect'][0], gameWindow['rect'][1])
            )

            if gameInstanceState.round_rotation_started():
                # Toggle through players while game is running
                print_log('Entering player spectate rotation')

                # Move cursor to center of screen to avoid accidentally clicking buttons
                mouse_move(gameWindow['rect'][0] + 640, gameWindow['rect'][1] + 360)
        else:
            time.sleep(4)
    elif gameInstanceState.round_started_spectation() and iterationsOnPlayer >= 10:
        # Click to rotate
        print_log('Rotating')
        mouse_left_click()

        # Reset counter
        iterationsOnPlayer = 0
    else:
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

        # Switch to next player if current player has reached iteration limit
        # and game is not over
        if not onInMemoriam and not leaveGameButtonPresent:
            iterationsOnPlayer += 1
            time.sleep(2)
        elif onInMemoriam:
            # Hit ESC
            print_log('Game over, hitting ESC')
            auto_press_key(0x01)
            time.sleep(1)

            # Click "Leave game"-button
            print_log('Clicking "Leave game"-button')
            mouse_move(gameWindow['rect'][0] + 150, gameWindow['rect'][1] + 312)
            mouse_left_click()
            time.sleep(1)

            # Confirm
            print_log('Confirming "Leave game"-dialogue')
            mouse_move(gameWindow['rect'][0] + 643, gameWindow['rect'][1] + 386)
            mouse_left_click()
            time.sleep(5)

            # Reset game instance state
            gameInstanceState.round_end_reset()
        elif leaveGameButtonPresent:
            # Click "Leave game"-button
            print_log('Clicking "Leave game"-button')
            mouse_move(gameWindow['rect'][0] + 1043, gameWindow['rect'][1] + 402)
            mouse_left_click()
            time.sleep(1)

            # Confirm
            print_log('Confirming "Leave game"-dialogue')
            mouse_move(gameWindow['rect'][0] + 643, gameWindow['rect'][1] + 386)
            mouse_left_click()
            time.sleep(5)

            # Reset game instance state
            gameInstanceState.round_end_reset()
