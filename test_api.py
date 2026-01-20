import requests
import os

url = "http://127.0.0.1:8000/pdf"
file_path = "template_dummy.xlsx" # I hope this exists or I'll use another one

# Check if file exists
if not os.path.exists(file_path):
    print(f"Error: {file_path} not found. Creating a dummy one...")
    import pandas as pd
    df = pd.DataFrame(columns=["Date", "Ticket No.", "REQ No.", "Type", "Requested for", "Assign To", "Request Detail", "Time - Arrive", "Time - Close", "Remarks", "Status"])
    df.to_excel(file_path, index=False)

files = {'file': open(file_path, 'rb')}
data = {
    'begin_date': '2026-01-12',
    'end_date': '2026-01-16'
}

try:
    print(f"Sending request to {url}...")
    response = requests.post(url, files=files, data=data)
    print(f"Status Code: {response.status_code}")
    print(f"Headers: {response.headers}")
    if response.status_code != 200:
        print(f"Response Body: {response.text}")
    else:
        print("Success! Received a response.")
except Exception as e:
    print(f"Request failed: {str(e)}")
