import pyautogui as pygui
import time
from datetime import datetime as dt

def waitFor(hours, minutes, interval):
    finalTime = hours * 60 + minutes

    print('Waiting for ' + str(hours) + ':' + str(minutes) + ' (' + str(finalTime) + ') in intervals of ' + str(interval) + ' minutes:')

    currentTime = dt.now()
    currTime = currentTime.hour * 60 + currentTime.minute

    print('currTime < finalTime:')
    print(currTime < finalTime)
    print('currTime: ' + str(currTime))
    print('finalTime: ' + str(finalTime))

    while(currTime < finalTime):
        time.sleep(interval * 60)
        currentTime = dt.now()
        currTime = currentTime.hour * 60 + currentTime.minute
        print('currTime: ' + str(currTime))
    
    print('Done waiting for ' + str(hours) + ':' + str(minutes) + '.')



def recordClass():
    pygui.FAILSAFE = False

    pygui.click(958,1057)   # Open zoom
    time.sleep(10)
    pygui.click(803,447)    # Click join
    time.sleep(1)
    pygui.click(881,481)    # Click text field
    time.sleep(1)
    # pygui.typewrite('4317179320') # Test id
    pygui.typewrite('7373508675') # Write id
    pygui.click(980,649)    # Click join meeting
    time.sleep(60)  # Wait for meeting to load
    pygui.click(941,519)    # Click join computer audio
    pygui.click(394,780)    # Mute myself
    pygui.click(1501,19)    # Fullscreen mode

    # Start screen record
    pygui.hotkey('alt','s')
    pygui.moveTo(0,0)
    pygui.dragTo(1919,1079)
    pygui.click(96,1024)
    
    waitFor(10,00,5)

    pygui.hotkey('alt','q')    # Exit meeting
    pygui.click(952,490)    # Confirm exit meeting
    pygui.rightClick(958,1057)   # Right click zoom
    pygui.rightClick(941,907)   # Exit zoom

    # Save video
    pygui.hotkey('alt','s') # End screen record
    time.sleep(5 * 60)   # Give more time for window to pop up
    pygui.click(555,746)    # Save video to drive
    time.sleep(2)
    pygui.click(994,704)    # Confirm save location

    # Exit zoom
    pygui.rightClick(958,1057)   # Right click zoom
    time.sleep(1)
    pygui.click(953,904)    # Quit zoom

if __name__ == '__main__':
    waitFor(8,30,1)
    
    recordClass()

# pyinstaller --onefile .\classrec.py