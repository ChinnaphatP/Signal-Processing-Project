import requests
import time
import pandas as pd
from datetime import datetime


API_KEY = "016da20c-a163-4c14-853f-79470ceb6266" # Your API KEY here
API_URL = "http://api.airvisual.com/v2/city" 
EXCEL_FILE_PATH = "air_quality_data.xlsx"

def fetch_data_from_api():
    params = {
        "city": "Sofia", # Your city
        "state": "Sofia", # Your state
        "country": "Bulgaria", # Your country
        "key": API_KEY
    }

    try:
        response = requests.get(API_URL, params=params)
        if response.status_code == 200:
            data = response.json()
            return data
        else:
            print("Error:", response.status_code)
            return None

    except Exception as e:
        print("An error occurred:", str(e))
        return None


# Set the end date and time (February 1, 2024, 00:00:00)
end_datetime = pd.to_datetime("2024-01-26 00:00:00")

columns = ["City", "State", "Country", "Latitude", "Longitude", "AQI (US)", "Main (US)", "AQI (CN)", "Main (CN)", "Timestamp", "Temperature", "Pressure", "Humidity", "Wind Speed", "Wind Direction", "Weather Icon"]
data_df = pd.DataFrame(columns=columns)
# Count 1 hour to fetch the API
fetch_interval_seconds = 3600

while pd.to_datetime("now") < end_datetime:
    data = fetch_data_from_api()
    if data:
        city = data['data']['city']
        state = data['data']['state']
        country = data['data']['country']
        coordinates = data['data']['location']['coordinates']
        pollution_data = data['data']['current']['pollution']
        weather_data = data['data']['current']['weather']
        
        current_timestamp = pd.to_datetime("now")

        data_df = data_df._append({
            "City": city,
            "State": state,
            "Country": country,
            "Latitude": coordinates[1],
            "Longitude": coordinates[0],
            "AQI (US)": pollution_data['aqius'],
            "Main (US)": pollution_data['mainus'],
            "AQI (CN)": pollution_data['aqicn'],
            "Main (CN)": pollution_data['maincn'],
            "Timestamp": current_timestamp,
            "Temperature": weather_data['tp'],
            "Pressure": weather_data['pr'],
            "Humidity": weather_data['hu'],
            "Wind Speed": weather_data['ws'],
            "Wind Direction": weather_data['wd'],
            "Weather Icon": weather_data['ic']
        }, ignore_index=True)
        print("Now Data : " , datetime.now() , '\n')
        print(data_df)
        # Change your name of CSV file here
        data_df.to_csv('AQI_PM2_5_Bulgaria.csv', mode='a', index=False, header=not pd.DataFrame(columns=columns).empty)

    # Wait for the specified interval before making the next request
    time.sleep(fetch_interval_seconds)

data_df.to_csv('data_until_2024-02-01.csv', index=False)