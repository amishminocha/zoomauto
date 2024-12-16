import schedule
import time
import pyautogui
import xlrd

loc = "C:\\Users\\amish\\Desktop\\zoomExcel.xls"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)


meeting_time = str(sheet.cell_value(1, 0))
meeting_id = sheet.cell_value(1, 1)
meeting_password = sheet.cell_value(1, 2)


def join_meeting():
    pyautogui.press("win")
    time.sleep(2)
    pyautogui.write("zoom", interval=0.25)
    pyautogui.press("enter")
    time.sleep(4)
    pyautogui.moveTo(743, 438, 2)
    pyautogui.leftClick()
    time.sleep(4)
    pyautogui.write(meeting_id, interval=0.25)
    pyautogui.press("enter")
    # time.sleep(2)
    # pyautogui.write(meeting_password, interval=0.25)
    # pyautogui.press("enter")


schedule.every().day.at("18:18:00").do(join_meeting)

while True:
    schedule.run_pending()
    time.sleep(1)
