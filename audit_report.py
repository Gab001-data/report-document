from docxtpl import DocxTemplate
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION, WD_ORIENTATION
from docx.shared import RGBColor, Mm, Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_ROW_HEIGHT_RULE
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from urllib.parse import urljoin
import json
from datetime import datetime as dt


####################################################################################
# fetch data
####################################################################################

BEARER_TOKEN = '0620e9abc1c94f1d0c4ab137ef1c4647cb8c4875b25450742af54742609f514c'
org_id = '49959' #'52734'
server_url = 'https://apis-eu.diligentoneplatform.com'
BASE_URL = f'https://apis-eu.diligentoneplatform.com/v1/orgs/{org_id}/projects'
headers = {'Authorization' : f'Bearer {BEARER_TOKEN}', 'Content-Type': 'application/vnd.api+json'}
params = {'filter[status]':'active', 'filter[start_date][lte]': dt.today().isoformat(), 'include':'project, project.fieldwork'}

def fetch_data(url, params=params, max_retries=5, timeout=10, backoff_factor=1):
    """
        fetch data from endpoint with maxixum retries, timeout and session management

        Args:
            url(str): targeted endpoint
            max_retries(int): no of retry attempt upon non-response from the server
            timeout(int): prevent prolonged server connection
            backoff_factor = wait multiplier between request
    """
    retries = Retry(total= max_retries, 
                    backoff_factor=backoff_factor, 
                    status_forcelist=[429, 500, 502, 503, 504], #retries on these http errors
                    allowed_methods=['GET'])
    session = requests.session()
    adapter = HTTPAdapter(max_retries=retries)
    session.mount('http://',adapter)
    session.mount('https://', adapter)
    all_data = []
    while url:
        try:
            resp = session.get(url, headers=headers, params=params, timeout=10)
            resp.raise_for_status()
        except Exception as e:
            print(f'failure to fetch data: {e}')
        all_data = all_data + resp.json()['data']

        next_page = resp.json()['links']['next']
        url = urljoin(server_url,next_page) if next_page else None
        params = None
    session.close()
    return all_data

# fetch project data ***************************************************************
projects = fetch_data(BASE_URL)

# fetch planning data **************************************************************
project_plannings = {}
for project in projects:
    project_id = project['id']
    BASE_URL = f'https://apis-eu.diligentoneplatform.com/v1/orgs/{org_id}/projects/{project_id}/planning_files'
    plannings = fetch_data(BASE_URL, params=None)
    project_plannings[project_id] = plannings

print(project_plannings, len(plannings))
