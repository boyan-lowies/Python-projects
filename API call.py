import requests

response = requests.get("https://api.guildwars2.com/v2/items/12452")
text = response.json()

print(text['name']) 