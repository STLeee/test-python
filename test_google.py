import requests

api_key = ""
cse_id = ""

query = "多采多姿的植物"
num_results = 10

response = requests.get(
    "https://www.googleapis.com/customsearch/v1",
    params={
        "key": api_key,
        "cx": cse_id,
        "q": query,
        "num": num_results,
        "lr": "lang_zh-TW",
    },
)

response.raise_for_status()
data = response.json()

google_links = data.get("items", [])
for google_link in google_links:
    print("==================================================")
    for key, value in google_link.items():
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                print(f"        {sub_key}: {sub_value}")
        else:
            print(f"    {key}: {value}")
