# cod-warzone-auto-spectator
An automated spectator for Call of Duty®: Warzone written in Python 🐍

Ever wanted to have something to fill your break time when you stream? Have an opening act for your stream? Have your stream fade out nicely? Here you go!

## Features
- automatic start of Battle.net client and Call of Duty®: Warzone itself
- automatic search for battle royale match
- automatic deployment in match
- automatic skipping of round end/credits
- (some) error detection
- automatic game restart on error

## Setup
1. Download and install Tesseract v4.1.0.20190314 from the [Uni Mannheim server](https://digi.bib.uni-mannheim.de/tesseract/) (be sure to use the 64 bit version, [direct link])
2. Download the [latest release](https://github.com/cetteup/cod-warzone-auto-spectator/releases/latest)
3. Enable windowed mode in the game
4. Disable the "fill squad"-option in the game

## How to run
1. Open CMD or Powershell
2. Enter the path to the controller.exe (can be done by dragging & dropping the .exe onto the CMD/Powershell window)
   1. If you have installed Warzone to a path other than `C:\Program Files (x86)\Call of Duty Modern Warfare`, enter `--game-path` followed by a space and the path to your game install folder
   2. If you have installed Tesseract to something other than `C:\Program Files\Tesseract-OCR`, enter `--tesseract-path` followed by a space and the path to your Tesseract install folder
3. Hit enter to run

If you want to stop the spectator, hit CTRL + C at any time.

**Please note: You cannot (really) use the computer while the spectator is running. It relies on having control over mouse and keyboard and needs the game window to be focused and in the foreground.** You do, however, have small time windows between the spectator's actions in which you can start/stop the stream, stop the spectator etc.

## Known limitations
- cod-warzone-auto-spectator currently only supports running the game in 720p
- Windows display scaling has to be set to 100%
- player spectate rotation is limited to the current squad
- current/next squad to watch is chosen by the game, whichever squad eliminates the current squad will be the next squad to spectate (starting with the squad of the spectator)
