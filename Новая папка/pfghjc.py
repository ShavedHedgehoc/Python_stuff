import requests
import pandas as pd
import json
response=requests.get("http://srv-webts:9000/MobileSMARTS/api/v1/Docs/Vzveshivanie")

data=response.json()
count=1
for value in data['value']:
    print(value['name'])
    count+=1
print(count)

    

