import requests

proxies = {
    "http": "http://antema103.gmail.com:9yucvu@gate2.proxyfuel.com:2000",
    "https": "http://antema103.gmail.com:9yucvu@gate2.proxyfuel.com:2000",
}

# Pretend to be Firefox
headers = {
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:87.0) Gecko/20100101 Firefox/87.0'
}

# Change the URL to your target website
url = 'https://bizzy.org/fr/fr/424264281'
try:
    r = requests.get(url, proxies=proxies, headers=headers, timeout=20)
    print(r.status_code)
except Exception as e:
    print(e)
