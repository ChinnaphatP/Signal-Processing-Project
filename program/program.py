import requests
import pandas as pd
from datetime import datetime, timedelta
from apscheduler.schedulers.blocking import BlockingScheduler

API_KEY = "016da20c-a163-4c14-853f-79470ceb6266"
API_URL = "http://api.airvisual.com/v2/city"
EXCEL_FILE_PATH = "air_quality_data.xlsx"

def get_air_quality_data():
    params = {
        "city": "Sofia",
        "state": "Sofia",
        "country": "Bulgaria",
        "key": API_KEY
    }

    response = requests.get(API_URL, params=params)
    data = response.json()

    if response.status_code == 200 and data.get("status") == "success":
        return data["data"]["current"]["pollution"]
    else:
        print("Failed to retrieve air quality data.")
        return None

def save_to_excel(data):
    current_time_thailand = datetime.utcnow() + timedelta(hours=7)
    formatted_time = current_time_thailand.strftime("%Y-%m-%d %H:%M:%S")
    aqius = data["aqius"]

    df = pd.DataFrame({"Timestamp": [formatted_time], "AQI": [aqius]})
    
    try:
        with pd.ExcelWriter(EXCEL_FILE_PATH, mode="a", engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="AirQuality", startrow=writer.sheets["AirQuality"].max_row, header=writer.sheets["AirQuality"].max_row == 0)
        print(f"Data saved to {EXCEL_FILE_PATH} at {formatted_time}.")
    except Exception as e:
        print(f"Error saving data to Excel: {e}")

if __name__ == "__main__":
    scheduler = BlockingScheduler()
    scheduler.add_job(get_air_quality_data, 'interval', hours=1, id='get_air_quality_job')
    scheduler.add_job(save_to_excel, 'interval', hours=1, args=[get_air_quality_data()], id='save_to_excel_job')
    scheduler.start()
