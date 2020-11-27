from os import listdir, path
import xlsxwriter as excel
from pynput import mouse
import xlrd
from datetime import datetime
import time
import pyautogui as pygui


tmpCoords = []
isPressed = False


def on_click(x, y, button, pressed):
    global tmpCoords
    global isPressed

    if pressed:
        isPressed = True
        tmpCoords = [x, y]


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


def buildFunctions(worksheet):
    global tmpCoords
    global isPressed

    count = 1
    endFlag = False
    while endFlag == False:
        if count > 1:
            endInput = str(input('Add another function? (y/n) ')).lower()

            if endInput == 'y':
                pass
            elif endInput == 'n':
                endFlag = True
                continue
            else:
                print('No valid option was selected.\n')
                continue


        print('Choose a function:\n\t1. Click\n\t2. Right-Click\n\t3. Write\n\t4. Write Class ID\n\t5. Hotkey\n\t6. Wait')
        choice = input('Your Choice (1-5): ')

        try:
            choice = int(choice)
        except:
            print('Only numbers are accepted.')
            continue

        if choice == 1:
            worksheet.write(count, 0, 'click')
            print('Click on where you want to click to log the coordinates.')
            getCoordinates()
            worksheet.write(count, 1, tmpCoords[0])
            worksheet.write(count, 2, tmpCoords[1])
            print('Coordinates logged on ', tmpCoords)
            
            delay = str(input('Add a mouse movement delay? (y/n) ')).lower()

            if delay == 'y':
                delayNumber = input('How much mouse movement delay do you want (number in seconds)? ')
                
                try:
                    delayNumber = int(delayNumber)
                except:
                    print('A number was not inputted, setting mouse movement delay to 0 (default).\n')
                    worksheet.write(count, 3, 0)
                    return
                
                worksheet.write(count, 3, delayNumber)
                print('A mouse movement delay of ' + str(delayNumber) + ' seconds was set.\n')

            elif delay == 'n':
                worksheet.write(count, 3, 0)
                print()

            else:
                print('No valid choice was selected, movement delay will be 0 (default).\n')
                worksheet.write(count, 3, 0)
            
            count += 1

        elif choice == 2:
            worksheet.write(count, 0, 'rightclick')
            print('Click on where you want to right click to log the coordinates.')
            getCoordinates()
            worksheet.write(count, 1, tmpCoords[0])
            worksheet.write(count, 2, tmpCoords[1])
            print('Coordinates logged on ', tmpCoords)

            delay = str(input('Add a mouse movement delay? (y/n) ')).lower()

            if delay == 'y':
                delayNumber = input('How much mouse movement delay do you want (number in seconds)? ')
                
                try:
                    delayNumber = int(delayNumber)
                except:
                    print('A number was not inputted, setting mouse movement delay to 0 (default).\n')
                    worksheet.write(count, 3, 0)
                    return
                
                worksheet.write(count, 3, delayNumber)
                print('A mouse movement delay of ' + str(delayNumber) + ' seconds was set.\n')

            elif delay == 'n':
                worksheet.write(count, 3, 0)
                print()

            else:
                print('No valid choice was selected, movement delay will be 0 (default).\n')
                worksheet.write(count, 3, 0)
            
            count += 1

        elif choice == 3:
            worksheet.write(count, 0, 'write')
            text = str(input('What do you want to write? '))
            worksheet.write(count, 1, text)
            print()

            count += 1

        elif choice == 4:
            worksheet.write(count, 0, 'writeid')
            print()
            
            count += 1

        elif choice == 5:
            worksheet.write(count, 0, 'hotkey')
            print('Hotkey Syntax:')
            print('\tctrl - control\n\tcommand - command (mac)\n\twin - windows key\n\talt- alt\n\toption - option (alt mac)')
            print('\tesc - escape\n\tenter - enter\n\tshift - shift\n\tdel - delete (backspace)\n\tspace - space')

            hotkeyCount = 1
            while hotkeyCount <= 3:
                if hotkeyCount > 1:
                    newKey = str(input('Add another key? (y/n) ')).lower()

                    if newKey == 'y':
                        pass
                    elif newKey == 'n':
                        hotkeyCount = 4
                        print()
                        continue
                    else:
                        print('No valid input was recieved.\n')
                        continue

                hotkey = str(input('Key Press: '))
                worksheet.write(count, hotkeyCount, hotkey)
                hotkeyCount += 1
            
            print()
            count += 1

        
        elif choice == 6:
            worksheet.write(count, 0, 'wait')
            waitTime = input('How many seconds do you want to wait until the next action (number)? ')
            
            try:
                waitTime = int(waitTime)
                worksheet.write(count, 1, waitTime)
                print()
            except:
                print('A number was not inputted, setting wait time to 0 seconds (default).\n')
                worksheet.write(count, 0, 0)
                continue
            
            count += 1

        else:
            print('No valid option was selected.\n')
            continue


def createAutomation():
    # Prompt for file name
    automationName = input('Input class recorder driver file name (without extension): ')
    automationName = str(automationName.lower().replace(' ', '_'))

    # Create excel file
    workbook = excel.Workbook('./classrec/' + automationName + '.xlsx')
    openClass = workbook.add_worksheet('open_class')

    bold = workbook.add_format({ 'bold': True })

    openClass.write(0, 0, 'function', bold)
    openClass.write(0, 1, 'Param 1', bold)
    openClass.write(0, 2, 'Param 2', bold)
    openClass.write(0, 3, 'Param 3', bold)

    print('\nFirstly, we need to know how to join your class and start recording.')
    
    buildFunctions(openClass)
    
    closeClass = workbook.add_worksheet('close_class')
    closeClass.write(0, 0, 'function', bold)
    closeClass.write(0, 1, 'Param 1', bold)
    closeClass.write(0, 2, 'Param 2', bold)
    closeClass.write(0, 3, 'Param 3', bold)

    print('\nSecondly, we need to know how to leave your class and stop recording.')

    buildFunctions(closeClass)

    classHours = workbook.add_worksheet('class_hours')
    classHours.write(0, 0, 'Class Name', bold)
    classHours.write(0, 1, 'Class Id', bold)
    classHours.write(0, 2, 'Class Password', bold)
    classHours.write(0, 3, 'Monday Start', bold)
    classHours.write(0, 4, 'Monday End', bold)
    classHours.write(0, 5, 'Tuesday Start', bold)
    classHours.write(0, 6, 'Tuesday End', bold)
    classHours.write(0, 7, 'Wednesday Start', bold)
    classHours.write(0, 8, 'Wednesday End', bold)
    classHours.write(0, 9, 'Thursday Start', bold)
    classHours.write(0, 10, 'Thursday End', bold)
    classHours.write(0, 11, 'Friday Start', bold)
    classHours.write(0, 12, 'Friday End', bold)
    classHours.set_column(0, 0, 40)
    classHours.set_column(1, 12, 30)

    print('\nLastly, we need to know your class hours.')

    weekdayDict = {
        3: 'Monday',
        4: 'Tuesday',
        5: 'Wednesday',
        6: 'Thursday',
        7: 'Friday'
    }

    count = 1
    endFlag = False
    while endFlag == False:
        if count > 1:
            endInput = str(input('Add another class? (y/n) ')).lower()

            if endInput == 'y':
                pass
            elif endInput == 'n':
                endFlag = True
                continue
            else:
                print('No valid option was selected.\n')
                continue

        className = input('Class Name: ')
        classHours.write(count, 0, className)

        classId = input('Class Id: ')
        classHours.write(count, 1, classId)

        classPass = input('Class Password (write "-" if there is no class password): ')
        classHours.write(count, 2, classPass)

        colCount = 3
        for i in range(3, 8):
            print('\n' + weekdayDict[i] + ' class start time (write "-" if there is no start time):')
            startHour = input('Start hour from 0:00 to 23:00 : ')

            if startHour == '-':
                classHours.write(count, colCount, '-')
                classHours.write(count, colCount+1, '-')
                colCount += 2
                continue

            print(weekdayDict[i] + ' class end time (write "-" if there is no end time):')
            endHour = input('End hour from 0:00 to 23:00 : ')

            classHours.write(count, colCount, startHour)
            classHours.write(count, colCount+1, endHour)
            
            colCount += 2
        
        count += 1
        
        print()
    
    print('\nCreating class recorder driver...')
    workbook.close()
    print('Created excel file under ./classrec called ' + automationName + '.xlsx')


def mergesortClasses(arr):
    if len(arr) > 1:
        mid = len(arr) // 2

        left = arr[:mid]
        right = arr[mid:]

        mergesortClasses(left)
        mergesortClasses(right)

        idxL = idxR = idxA = 0

        while idxL < len(left) and idxR < len(right):
            leftTime = int(left[idxL].start.split(':')[0])
            rightTime = int(right[idxR].start.split(':')[0])

            if leftTime < rightTime:
                arr[idxA] = left[idxL]
                idxL += 1
            else:
                arr[idxA] = right[idxR]
                idxR += 1
            
            idxA += 1
        
        while idxL < len(left):
            arr[idxA] = left[idxL]
            idxL += 1
            idxA += 1
        
        while idxR < len(right):
            arr[idxA] = right[idxR]
            idxR += 1
            idxA += 1
    
def waitFor(hours, minutes, interval):
    finalTime = hours * 60 + minutes

    print('Waiting for ' + str(hours) + ':' + str(minutes) + ' (' + str(finalTime) + ') in intervals of ' + str(interval) + ' minutes:')

    currentTime = datetime.now()
    currTime = currentTime.hour * 60 + currentTime.minute

    print('currTime < finalTime: ' + str(currTime < finalTime))
    print('currTime: ' + str(currTime))
    print('finalTime: ' + str(finalTime))

    while(currTime < finalTime):
        time.sleep(interval * 60)
        currentTime = datetime.now()
        currTime = currentTime.hour * 60 + currentTime.minute
        print('currTime: ' + str(currTime))
    
    print('Done waiting for ' + str(hours) + ':' + str(minutes) + '.')


class Class:
    def __init__(self, name, classId, password, start, end):
        self.name = name
        self.classId = classId
        self.password = password
        self.start = start
        self.end = end

class Function:
    def __init__(self, name, params):
        self.name = name
        self.params = params

def runClassRec(path):
    pygui.FAILSAFE = False

    currentWeekday = datetime.today().weekday() # 0 for Monday, 6 for sunday

    dayDict = {
        0: 3,
        1: 5,
        2: 7,
        3: 9,
        4: 11
    }

    driver = xlrd.open_workbook(path)
    openSheet = driver.sheet_by_index(0)
    closeSheet = driver.sheet_by_index(1)
    classSheet = driver.sheet_by_index(2)

    if openSheet.nrows <= 1:
        print('\nNo open class functions were found. Please add them to ' + path)
        return
    
    openFunctions = []
    for row in range(openSheet.nrows):
        if row == 0:
            continue
        
        functionName = openSheet.cell_value(row, 0)
        params = []
        for col in range(1, 4):
            if openSheet.cell_value(row, col) != '':
                params.append(openSheet.cell_value(row, col))
        
        functionObj = Function(functionName, params)
        openFunctions.append(functionObj)
        
    if closeSheet.nrows <= 1:
        print('\nNo close class functions were found. Please add them to ' + path)
        return
    
    closeFunctions = []
    for row in range(closeSheet.nrows):
        if row == 0:
            continue
        
        functionName = closeSheet.cell_value(row, 0)
        params = []
        for col in range(1, 4):
            if closeSheet.cell_value(row, col) != '':
                params.append(closeSheet.cell_value(row, col))
        
        functionObj = Function(functionName, params)
        closeFunctions.append(functionObj)


    if classSheet.nrows <= 1:
        print('\nNo classes were found. Please add them to ' + path)
        return
    
    # Make Class temp objects into array
    currentWeekday = 3
    classes = [] 
    for row in range(classSheet.nrows):
        if row == 0 or classSheet.cell_value(row, dayDict[currentWeekday]) == '-':
            continue
        
        className = classSheet.cell_value(row, 0)
        classId = classSheet.cell_value(row, 1)
        classPass = classSheet.cell_value(row, 2)
        classStart = classSheet.cell_value(row, dayDict[currentWeekday])
        classEnd = classSheet.cell_value(row, dayDict[currentWeekday] + 1)
        
        classObj = Class(className, classId, classPass, classStart, classEnd)

        classes.append(classObj)

    # Mergesort classes by start time
    mergesortClasses(classes)

    print('Today\'s classes: ')
    for classObj in classes:
        print(classObj.name + ' (' + classObj.start + '-' + classObj.end + ')')
    
    currentTime = datetime.now()
    print('\nCurrent Time: ' + currentTime.strftime('%H:%M'))

    # Program run before today's first class
    recordClasses = classes.copy()
    firstClass = currentTime.replace(hour=int(classes[0].start.split(':')[0]), minute=int(classes[0].start.split(':')[1]))
    
    if firstClass > currentTime:
        pass
    
    else:
        i = 0
        while i < len(recordClasses):
            classStartTime = currentTime.replace(hour=int(recordClasses[i].start.split(':')[0]), minute=int(recordClasses[i].start.split(':')[1]))
            classEndTime = currentTime.replace(hour=int(recordClasses[i].end.split(':')[0]), minute=int(recordClasses[i].end.split(':')[1]))

            if classStartTime < currentTime and currentTime < classEndTime:
                break
            elif classEndTime < currentTime:
                recordClasses.pop(i)
                continue
            
            i += 1
    
    if len(recordClasses) == 0:
        print('\nNo classes left to record today, exitting...')
        return
    
    for classObj in recordClasses:
        print('\nWaiting for ' + classObj.name + ' class to start.')
        waitFor(int(classObj.start.split(':')[0]), int(classObj.start.split(':')[1]), 1)
        print('Opening ' + classObj.name + ' class and starting recording...')

        for function in openFunctions:
            if function.name == 'click':
                if len(function.params) != 3:
                    print('Not enough params for \'click\' function. Skipping to next function...')
                    continue
                else:
                    pygui.click(int(function.params[0]), int(function.params[1]), int(function.params[2]))
            elif function.name == 'rightclick':
                if len(function.params) != 3:
                    print('Not enough params for \'rightclick\' function. Skipping to next function...')
                    continue
                else:
                    pygui.rightClick(int(function.params[0]), int(function.params[1]), int(function.params[2]))
            elif function.name == 'write':
                if len(function.params) != 1:
                    print('Not enough params for \'write\' function. Skipping to next function...')
                    continue
                else:
                    pygui.typewrite(function.params[0])
            elif function.name == 'writeid':
                pygui.typewrite(classObj.classId)
            elif function.name == 'hotkey':
                if len(function.params) == 0:
                    print('Not enough params for \'hotkey\' function. Skipping to next function...')
                    continue
                elif len(function.params) == 1:
                    pygui.hotkey(function.params[0])
                elif len(function.params) == 2:
                    pygui.hotkey(function.params[0], function.params[1])
                else:
                    pygui.hotkey(function.params[0], function.params[1], function.params[2])
            elif function.name == 'wait':
                if len(function.params) == 0:
                    print('Not enough params for \'wait\' function. Skipping to next function...')
                    continue
                else:
                    time.sleep(int(function.params[0]))
            else:
                continue
            
        print('Waiting for ' + classObj.name + ' class to finish.')
        waitFor(int(classObj.end.split(':')[0]), int(classObj.end.split(':')[1]), 1)
        print('Stopping recording and closing ' + classObj.name + ' class...')

        for function in closeFunctions:
            if function.name == 'click':
                if len(function.params) != 3:
                    print('Not enough params for \'click\' function. Skipping to next function...')
                    continue
                else:
                    pygui.click(int(function.params[0]), int(function.params[1]), int(function.params[2]))
            elif function.name == 'rightclick':
                if len(function.params) != 3:
                    print('Not enough params for \'rightclick\' function. Skipping to next function...')
                    continue
                else:
                    pygui.rightClick(int(function.params[0]), int(function.params[1]), int(function.params[2]))
            elif function.name == 'write':
                if len(function.params) != 1:
                    print('Not enough params for \'write\' function. Skipping to next function...')
                    continue
                else:
                    pygui.typewrite(function.params[0])
            elif function.name == 'writeid':
                pygui.typewrite(classObj.classId)
            elif function.name == 'hotkey':
                if len(function.params) == 0:
                    print('Not enough params for \'hotkey\' function. Skipping to next function...')
                    continue
                elif len(function.params) == 1:
                    pygui.hotkey(function.params[0])
                elif len(function.params) == 2:
                    pygui.hotkey(function.params[0], function.params[1])
                else:
                    pygui.hotkey(function.params[0], function.params[1], function.params[2])
            elif function.name == 'wait':
                if len(function.params) == 0:
                    print('Not enough params for \'wait\' function. Skipping to next function...')
                    continue
                else:
                    time.sleep(int(function.params[0]))
            else:
                continue


    # now = datetime.now()
    # print(now.strftime('%H:%M'))
    # dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

    # print(openSheet.nrows)
    # print(openSheet.ncols)




if __name__ == '__main__':
    # Get all files in ./classrec
    fileArray = []
    for file in listdir('./classrec'):
        if file.endswith('.xlsx'):
            fileArray.append(file)
    
    print('Welcome to Class Recorder!')
    print('--------------------------')

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
    
    else:
        selectedFilePath = './classrec/' + fileArray[0]
        runClassRec(selectedFilePath)