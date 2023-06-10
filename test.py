from datetime import datetime
import os

client_name = os.path.splitext("LANTEIGNE Samuel.txt")[0].split(" ")
excel_entry = {}
excel_entry["fname"] = client_name[1]
excel_entry["lname"] = client_name[0]

print(excel_entry)
