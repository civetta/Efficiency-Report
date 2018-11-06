import urllib2
import pandas as pd

url = "https://app.periscopedata.com/api/think-through-learning/chart/csv/701a76c4-5476-481d-bccf-4c53850fb9d8/487533"
response = urllib2.urlopen(url)

df = pd.read_csv(response)
print df.head()cd 