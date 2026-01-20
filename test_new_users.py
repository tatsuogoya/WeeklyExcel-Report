import requests

url = "http://127.0.0.1:8000/generate"
file_path = "sample_daily_work.xlsx"

files = {'file': open(file_path, 'rb')}
data = {
    'begin_date': '2026-01-05',
    'end_date': '2026-01-11'
}

try:
    response = requests.post(url, files=files, data=data)
    print(f"Status Code: {response.status_code}")
    if response.status_code == 200:
        res_json = response.json()
        print("\\n--- New Users Section ---")
        print(f"new_users_section type: {type(res_json.get('new_users_section'))}")
        print(f"new_users_section length: {len(res_json.get('new_users_section', []))}")
        print(f"Data: {res_json.get('new_users_section')}")
    else:
        print(f"Error: {response.text}")
except Exception as e:
    print(f"Request failed: {e}")
