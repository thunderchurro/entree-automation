from numbers_parser import Document
import pandas as pd
import numpy as np
from time import perf_counter, sleep
from AppKit import NSWorkspace
from rich.console import Console
from math import ceil
import pyautogui


pyautogui.FAILSAFE = True 

console = Console()


def isWindowFocused():
    focus = False

    activeAppName = NSWorkspace.sharedWorkspace().activeApplication()['NSApplicationName']

    if activeAppName == "Microsoft Remote Desktop":
        focus = True
    else:
        while focus is False:
            console.print("Window Not Found")
            sleep(2)
            activeAppName = NSWorkspace.sharedWorkspace().activeApplication()['NSApplicationName']
            if activeAppName == "Microsoft Remote Desktop":
                focus = True


def send_adjustment(item, quantity):
    isWindowFocused()
    console.log(f"Updating [bold]{item}[/] --> {quantity}")

    pyautogui.typewrite(item.lower())
    pyautogui.press('enter')
    sleep(0.20)

    pyautogui.press('tab')
    sleep(0.20)

    # pyautogui.hotkey('option', 'a')
    sleep(0.55)
    # pyautogui.write(quantity)
    # pyautogui.hotkey('option', 'o')
    sleep(1)


    pyautogui.hotkey('shift', 'tab')
    pyautogui.keyUp('shift')
    
    sleep(0.35)


def check_for_changes():
    t1_start = perf_counter()
    doc = Document("testdoc.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0] # for writing
    data = tables[0].rows(values_only=True)

    df = pd.DataFrame(data[1:], columns=data[0])

    changes = False

    for index, row in df.iterrows():
        data_index = index+1
        done_col = 7
        change_col = 6

        if row['Actual'] >= 0 and row['Done'] != 'OK':
            if changes == False: 
                changes = True


            if row['WHON'] == row['Actual']:
                table.write(data_index, change_col, "false")
            else:
                actual_count = int(row["Actual"])
                table.write(data_index, change_col, "true")
                console.log(f"Updating [bold]{row['Item']}[/] --> {actual_count}")

                # send to entree
                sleep(2.4)


            table.write(data_index, done_col, "OK")
    
    if changes == True:
        # doc.save("testdoc.numbers")
        t1_stop = perf_counter()
        console.print("Elapsed time during the whole program in seconds:", t1_stop-t1_start)

    changes = False


def check_for_changes2():
    t1_start = perf_counter()
    doc = Document("testdoc.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0] # for writing
    data = tables[0].rows(values_only=True)

    df = pd.DataFrame(data[1:], columns=data[0])
    changes = df.loc[(df['Actual'] >= 0) & (df['Done'] != "OK")]
    changes = changes[changes['Done'] != ""]

    if changes.shape[0] > 0:
        for index, row in changes.iterrows():
            data_index = index+1
            done_col, change_col = 7, 6

            if row['WHON'] == row['Actual']:
                table.write(data_index, change_col, "false")
            else:
                actual_count = int(row["Actual"])

                # send to entree
                send_adjustment(row['Item'], actual_count)
                
                table.write(data_index, change_col, "true")
                table.write(data_index, done_col, "OK")
        
        doc.save("testdoc.numbers")
        t1_stop = perf_counter()
        console.print("Elapsed time during the whole program in seconds:", ceil(t1_stop-t1_start))


def add_to_fmkt():
    t1_start = perf_counter()
    items = ["4PLA0001", "4PLA0002", "4PLA0003", "4PLA0004", "4PLA0005", "4PLA0006", "4PLA0007", "4PLA0009", "4PLA0012", "4POM0155", "4POM0157", "4PRC0001", "4RIM0001", "4VAL0163", "5GAI0662", "8DLR0001", "9SQS0150", "9SSC0110", "9SSC0115", "9SYH0050", "9THE0001", "9THE0002", "9THE0004", "9THE0005", "9THE0006", "9THE0008", "U1SA509", "U3AD525", "U3CA002", "U3CA003", "U3CA244", "U3DC001", "U3DI802", "U3JO034", "U3JO036", "U3NI001", "U3RI015", "U3RO435", "U3TU601", "U4CO702", "U9SA002", "U9SA005", "U9SA006"]

    for item in items:
        isWindowFocused()
        pyautogui.hotkey('option', 'a')
        sleep(.15)
        pyautogui.typewrite(item.lower())
        sleep(3)
        pyautogui.press('tab')
        sleep(.15)
        pyautogui.press('enter')
        sleep(.15)

            
        
    t1_stop = perf_counter()
    console.print("Elapsed time during the whole program in seconds:", ceil(t1_stop-t1_start))



if __name__ == "__main__":
    while True:
        # console.print("Checking for changes..", style="red")
        check_for_changes2()
        sleep(5)