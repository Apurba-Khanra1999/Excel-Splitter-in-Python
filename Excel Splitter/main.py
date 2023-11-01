

import os
import requests
import pandas as pd

# Step 1: Make an API request and get JSON data
api_url = "https://randomuser.me/api/?results=5"
response = requests.get(api_url)

if response.status_code != 200:
    print("Error: Failed to retrieve data from the API")
    exit()

api_data = response.json()

# Step 2: Parse JSON data
# Extract relevant information from the JSON response
data_list = []
for user in api_data['results']:
    user_info = {
        'First Name': user['name']['first'],
        'Last Name': user['name']['last'],
        'Email': user['email'],
        'Phone': user['phone'],
        'City': user['location']['city'],
        'Country': user['location']['country']
    }
    data_list.append(user_info)

# Step 3: Read existing data from Excel (if it exists)
excel_file_path = "random_users.xlsx"
if os.path.exists(excel_file_path):
    existing_data = pd.read_excel(excel_file_path)
else:
    existing_data = pd.DataFrame()

# Generate a unique sheet name
sheet_name = 'Sheet1'
i = 1
while sheet_name in pd.ExcelFile(excel_file_path).sheet_names:
    i += 1
    sheet_name = f'Sheet{i}'
# Append the new data to the existing data
updated_data = pd.concat([existing_data, pd.DataFrame(data_list)], ignore_index=True)

# Step 4: Write updated data to Excel
with pd.ExcelWriter(excel_file_path, engine="openpyxl", mode="a") as writer:
    updated_data.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Data successfully added to Excel, Sheet Name: {sheet_name}")
