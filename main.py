import pandas as pd
import pywhatkit
import pyautogui
import time

excel_file_path = 'Number_test.xlsx'

try:
    df = pd.read_excel(excel_file_path)

    df['Status'] = ''

    for index, row in df.iterrows():
        mobile_number = "+880" + str(row['NUMBER'])
        image_path = "image.png"
        caption = "Ramadan mubarak."

        print(f"Sending image to {mobile_number}")

        try:
            pywhatkit.sendwhats_image(mobile_number, image_path, caption, 10)
            df.at[index, 'Status'] = 'Sent'
        except pywhatkit.exceptions.InvalidNumberException as invalid_number_error:
            df.at[index, 'Status'] = 'Invalid Number'
            print(f"Error sending image to {mobile_number}: {invalid_number_error}")
        except Exception as send_error:
            df.at[index, 'Status'] = 'Error'
            print(f"Error sending image to {mobile_number}: {send_error}")

        time.sleep(2)  

       
        pyautogui.hotkey('ctrl', 'w')

        time.sleep(2)

       
        pyautogui.hotkey('ctrl', 't')

    
    df.to_excel(excel_file_path, index=False)

except Exception as e:
    print(f"An error occurred: {e}")
