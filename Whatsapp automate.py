import keyboard as k
import pyautogui
import time
import pandas as pd
import webbrowser as web
import os
from urllib.parse import quote

def send_whatsapp(data_file_excel, message_file_text, image_file=None, x_cord=830, y_cord=954):
    df = pd.read_excel(data_file_excel, dtype={"Contact": str})
    name = df['Name'].values
    contact = df['Contact'].values
    files = message_file_text

    with open(files) as f:
        file_data = f.read()

    zipped = zip(name, contact)

    counter = 0

    for (a, b) in zipped:
        msg = file_data.format(a)

        web.open(f"https://web.whatsapp.com/send?phone={b}")
        time.sleep(15)  

        if image_file and os.path.exists(image_file):
            pyautogui.click(795, 1380)  
            time.sleep(1)

            pyautogui.click(893, 977)  
            time.sleep(1)

            pyautogui.write(image_file)  
            pyautogui.press('enter')
            time.sleep(2)

        pyautogui.write(msg, interval=0.2)
        time.sleep(1)

        pyautogui.click(x_cord, y_cord)  
        time.sleep(2)
        k.press_and_release('enter')
        time.sleep(2)

        k.press_and_release('ctrl+w')
        time.sleep(1)

        counter += 1
        print(counter, "- Message sent..!!")

    print("Done!")

excel_path = r"C:\Users\braya\Desktop\0. Whatsapp Web Automation\Whatsapp List_Main.xlsx"
text_path = r"C:\Users\braya\Desktop\0. Whatsapp Web Automation\WHATSDRAFT.txt"
image_path = r"C:\Users\braya\Desktop\0. Whatsapp Web Automation\capibara.jpg"  

send_whatsapp(excel_path, text_path, image_file=image_path)


