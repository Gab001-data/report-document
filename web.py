import requests
from bs4 import BeautifulSoup


# start a session to maintain cookie

session = requests.session()

login_url = 'https://accounts.diligentoneplatform.com/login'

payload = {
    'email': 'gabrielo@galvanizeafrica.com',
    'password': 'Angel176391$'
}

# Common headers (make it look like a browser)
headers = {
    "User-Agent": "Edg/120.0.2210.91",
    "Referer": "https://accounts.diligentoneplatform.com/login",
    "Origin": "https://accounts.diligentoneplatform.com"
}

dashboard_url = 'https://cts.home.highbond.com/'
login_resp = session.post(login_url, data=payload, headers=headers)

if login_resp.ok:
    print(login_resp.text)
    dashboard_resp = session.get(dashboard_url, headers=headers)

    if dashboard_resp.ok:
        print('login successful! and dashboard loaded!')
        print('Current cookies:', session.cookies.get_dict())
        print(dashboard_resp.text)
    else:
        print('dashboard failed to load', dashboard_resp.status_code)
else:
    print('login failed!', login_resp.status_code, login_resp.text)

control_url = 'https://cts.projects.diligentoneplatform.com/audits/492704/objectives/1896423/controls/14421654'
controls_resp = session.get(control_url, headers=headers)

print(controls_resp.text)
print(session.cookies.get_dict())