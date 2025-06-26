from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import urllib3
import time
import warnings

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)

proxy_host = "brd.superproxy.io"
proxy_port = 33335  # Doit être int ici pour le manifest
zone_name = "datacenter_proxy1"
customer_id = "hl_1d81edaa"
proxy_username = f"brd-customer-{customer_id}-zone-{zone_name}-ip-168.151.121.174"  # ajoute ip si besoin
proxy_password = "wur28vaq23lx"

# Cette ligne ne fonctionne pas avec auth proxy dans Chrome
proxy_url = f"http://{proxy_username}:{proxy_password}@{proxy_host}:{proxy_port}"

chrome_options = Options()
chrome_options.add_argument(f"--proxy-server=http://{proxy_host}:{proxy_port}")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--ignore-ssl-errors")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    driver.get("https://annuaire.pagesjaunes.fr")
    print(driver.page_source[:1000])
    time.sleep(10)
finally:
    driver.quit()
