import requests
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
import os


API_KEY = "016da20c-a163-4c14-853f-79470ceb6266" # Your API KEY here
API_URL = "http://api.airvisual.com/v2/city" 
EXCEL_FILE_PATH = "air_quality_data.xlsx"

def get_air_quality_data():
    params = {
        "city": "Sofia", # Your city
        "state": "Sofia", # Your state
        "country": "Bulgaria", # Your country
        "key": API_KEY
    }

    response = requests.get(API_URL, params=params)
    data = response.json()

    if response.status_code == 200 and data.get("status") == "success":
        return data["data"]["current"]["pollution"]
    else:
        print("Failed to retrieve air quality data.")
        return None


def save_to_excel():
    current_time = datetime.now()
    aq_data = get_air_quality_data()

    if aq_data:
        formatted_time = current_time.strftime("%Y-%m-%d %H:%M:%S")
        aqius = aq_data["aqius"]

        df = pd.DataFrame({"Timestamp": [formatted_time], "AQI": [aqius]})

        try:
            # Check if the file exists
            if not os.path.exists(EXCEL_FILE_PATH):
                df.to_excel(EXCEL_FILE_PATH, index=False, sheet_name="AirQuality", header=True)
                print(f"Excel file {EXCEL_FILE_PATH} created.")
            else:
                # Check if the sheet exists
                with pd.ExcelFile(EXCEL_FILE_PATH) as xls:
                    sheet_exists = "AirQuality" in xls.sheet_names

                with pd.ExcelWriter(EXCEL_FILE_PATH, mode="a", engine="openpyxl") as writer:
                    if not sheet_exists:
                        df.to_excel(writer, index=False, sheet_name="AirQuality", header=True)
                        print(f"Sheet 'AirQuality' created in {EXCEL_FILE_PATH}.")
                    else:
                        df.to_excel(writer, index=False, sheet_name="AirQuality", startrow=writer.sheets["AirQuality"].max_row, header=None)
                        print(f"Data saved to {EXCEL_FILE_PATH} at {formatted_time}.")
        except Exception as e:
            print(f"Error saving data to Excel: {e}")
     
                                   
def run_scheduler():
    scheduler = BlockingScheduler()

    # ตั้งเวลาทำงานทุกๆ 1 ชั่วโมง
    scheduler.add_job(save_to_excel, 'interval', hours=1, id='save_to_excel_job')

    print("Program is running.")

    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        pass


if __name__ == "__main__":
    run_scheduler()