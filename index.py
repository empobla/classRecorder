import pynput
import pyautogui as pygui
import xlsxwriter as excel
from os import listdir, path
# from pynput.mouse import Listener
from pynput import mouse

tmpCoords = []
isPressed = False

def on_click(x, y, button, pressed):
    global tmpCoords
    global isPressed

    if pressed:
        isPressed = True
        tmpCoords = [x, y]


def demo():
    workbook = excel.Workbook('./Automation/demo.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.set_column('A:A', 20)

    bold = workbook.add_format({ 'bold': True })

    worksheet.write('A1', 'Hello')

    worksheet.write('A2', 'World', bold)

    worksheet.write(2, 0, 123)
    worksheet.write(3, 0, 123.456)

    workbook.close()

def getCoordinates():
    global tmpCoords
    global isPressed

    listener = mouse.Listener(on_click=on_click)
    listener.start()

    stopListening = False
    while stopListening == False:
        if isPressed:
            isPressed = False
            stopListening = True
    
    listener.stop()
    listener.join()


def createAutomation():
    global tmpCoords
    global isPressed

    # Prompt for file name
    automationName = input('Input class recorder driver file name (without extension): ')
    automationName = str(automationName.lower().replace(' ', '_'))

    # Create excel file
    workbook = excel.Workbook('./classrec/' + automationName + '.xlsx')
    openClass = workbook.add_worksheet('open_class')

    bold = workbook.add_format({ 'bold': True })

    openClass.write(0, 0, 'X', bold)
    openClass.write(0, 1, 'Y', bold)
    openClass.write(0, 2, 'Wait', bold)

    print('\nFirstly, we need to know how to join your class through mouse click coordinates:')

    print('\nClick on the appropriate location on screen to save first-click coordinates.')

    getCoordinates()

    openClass.write(1, 0, tmpCoords[0])
    openClass.write(1, 1, tmpCoords[1])

    print('First-click coordinates saved on ', tmpCoords)
    wait = abs(int(input('Wait how many seconds until next click? (0 for no seconds) ')))

    openClass.write(1, 2, wait)
    
    count = 2
    endFlag = False
    while endFlag == False:
        moreCoords = str(input('\nAdd more coordinates? (y/n) ')).lower()
        
        if moreCoords == 'y':
            print('Click where you want the next-click to automatically click next.')
            getCoordinates()
            openClass.write(count, 0, tmpCoords[0])
            openClass.write(count, 1, tmpCoords[1])
            print('Coordinates logged on ', tmpCoords)
            
            wait = abs(int(input('Wait how many seconds until next click? (0 for no seconds) ')))
            openClass.write(count, 2, wait)
            count += 1
        else:
            print('Finished creating file, exiting...')
            endFlag = True

    workbook.close()
    return


if __name__ == '__main__':
    # Get all files in ./classrec
    fileArray = []
    for file in listdir('./classrec'):
        if file.endswith('.xlsx'):
            fileArray.append(file)
    
    # If no class recorders are found, prompt to create one
    if len(fileArray) == 0:
        createAutomationInput = input('No .xlsx files found in ./classrec/ directory. Create a new class recorder driver? (y/n) ')
        createAutomationInput = createAutomationInput.lower()
        if createAutomationInput == 'y':
            createAutomation()
        elif createAutomationInput == 'n':
            print('Exiting...')
        else:
            print('No valid input was received.')