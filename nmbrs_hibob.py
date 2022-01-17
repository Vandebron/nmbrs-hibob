import argparse
import requests
import xml.etree.ElementTree as ET
import base64
from halo import Halo
from progress.bar import ShadyBar
from dataclasses import dataclass
from zipfile import ZipFile
import sys

parser = argparse.ArgumentParser(description='Export salary slip PDFs from Visma Nmbrs into Hibob')
parser.add_argument('--user', '-u', help=f'API user', default='IT@vandebron.nl')
parser.add_argument('--token', '-t', help=f'The API token https://support.nmbrs.com/hc/en-us/articles/'
                                          f'360013384371-Nmbrs-API', required=True)
parser.add_argument('--run', '-r', help='The run to download. Prints all runs for the year if not set.')
parser.add_argument('--year', '-y', help='The year in which the run took place', required=True)
parser.add_argument('--company', '-c', help='Select specific company number')
args = parser.parse_args()

user_arg = args.user
password_arg = args.token
run_arg = args.run
year_arg = args.year
company_id = args.company


@dataclass
class RunInfo:
    id: str
    number: str
    description: str
    period_start: str
    period_end: str


@dataclass
class Employee:
    id: str
    number: str


@dataclass
class EmployeeDetails:
    id: str
    number: str
    email: str


com = {'ns': 'com', 'url': 'https://api.nmbrs.nl/soap/v2.1/CompanyService'}
emp = {'ns': 'emp', 'url': 'https://api.nmbrs.nl/soap/v2.1/EmployeeService'}


def create_request(user, password, payload, namespace):
    ns_name = f"{namespace['ns']}"
    ns_prefix = f"{ns_name}:"
    url = namespace['url']
    template = f"""
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:{ns_name}="{url}">
    <soapenv:Header>
      <{ns_prefix}AuthHeader>
         <{ns_prefix}Username>{user}</{ns_prefix}Username>
         <{ns_prefix}Token>{password}</{ns_prefix}Token>
      </{ns_prefix}AuthHeader>
   </soapenv:Header>
   <soapenv:Body>
      {payload}
   </soapenv:Body>
</soapenv:Envelope>
"""
    return template


def get_company_info():
    return """<com:List_GetAll xmlns="https://api.nmbrs.nl/soap/v3/CompanyService" />"""


def get_runs(year):
    return f"""
<com:Run_GetList>
    <com:CompanyId>{company_id}</com:CompanyId>
    <com:Year>{year}</com:Year>
</com:Run_GetList>
"""


def get_employees(run, year):
    return f"""<com:Run_GetEmployeesByRunCompany>
      <com:CompanyId>{company_id}</com:CompanyId>
      <com:RunId>{run}</com:RunId>
      <com:Year>{year}</com:Year>
    </com:Run_GetEmployeesByRunCompany>"""


def get_payslip(employee_id, run_id, year):
    return f"""<com:SalaryDocuments_GetEmployeePayslipsPDFByRunCompany_v2>
         <com:EmployeeID>{employee_id}</com:EmployeeID>
         <com:CompanyId>{company_id}</com:CompanyId>
         <com:RunID>{run_id}</com:RunID>
         <com:intYear>{year}</com:intYear>
      </com:SalaryDocuments_GetEmployeePayslipsPDFByRunCompany_v2>"""


def get_employee(employee_id):
    return f"""<emp:PersonalInfo_GetCurrent>
         <emp:EmployeeId>{employee_id}</emp:EmployeeId>
      </emp:PersonalInfo_GetCurrent>"""


def do_request(body, service):
    req = create_request(user_arg, password_arg, body, service)
    # print(req)
    response = requests.post(f"{service['url']}.asmx", data=req,
                             headers={'content-type': 'text/xml; charset=utf-8'})
    if not response.ok:
        sys.exit(response.text)
    # print(response.text)
    return ET.ElementTree(ET.fromstring(response.text))


ns = {
    'soap': "http://schemas.xmlsoap.org/soap/envelope/",
    'xsi': "http://www.w3.org/2001/XMLSchema-instance",
    'xsd': "http://www.w3.org/2001/XMLSchema",
    'cs': 'https://api.nmbrs.nl/soap/v2.1/CompanyService',
    'emp': 'https://api.nmbrs.nl/soap/v2.1/EmployeeService'
}
spinner = Halo(text='Determining company ID..', spinner='dots')
spinner.start()

pdf_tree = do_request(get_company_info(), com)
companies = pdf_tree.findall('.//cs:Company', namespaces=ns)
if len(companies) > 1:
    companies = list(
        map(lambda c: c.find('./cs:ID', namespaces=ns).text + ": " + c.find('./cs:Name', namespaces=ns).text,
            companies))
    spinner.fail(f"More than one company found. Choose between {companies} via --company.")

company_id = companies.pop().find('./cs:ID', namespaces=ns).text


def to_run_info(element) -> RunInfo:
    id = element.find('./cs:ID', namespaces=ns).text
    number = element.find('./cs:Number', namespaces=ns).text
    description = element.find('./cs:Description', namespaces=ns).text
    period_start = element.find('./cs:PeriodStart', namespaces=ns).text
    period_end = element.find('./cs:PeriodEnd', namespaces=ns).text
    return RunInfo(id, number, description, period_start, period_end)


def to_employee_details(tree) -> EmployeeDetails:
    element = tree.find('.//emp:PersonalInfo_GetCurrentResult', namespaces=ns)
    id = element.find('./emp:Id', namespaces=ns).text
    number = element.find('./emp:Number', namespaces=ns).text
    email_element = element.find('./emp:EmailWork', namespaces=ns)
    email = number if email_element is None else email_element.text
    return EmployeeDetails(id, number, email)


def get_run_info(run):
    global spinner
    spinner.text = f"Getting run information for company {company_id} and year {year_arg}"
    result = do_request(get_runs(year_arg), com)
    spinner.succeed(f"Got run information for company {company_id} and year {year_arg}")
    if run:
        info = result.find(f'.//cs:RunInfo/[cs:ID="{run}"]', namespaces=ns)
        return to_run_info(info)
    run_infos = result.findall('.//cs:RunInfo', namespaces=ns)
    for run_info_element in run_infos:
        r = to_run_info(run_info_element)
        print(f'{r.id} {r.description} {r.period_start}-{r.period_end}')
    sys.exit()


run_info = get_run_info(run_arg)

spinner.text = f"Finding employees for {run_info.number} {run_info.description}"
pdf_tree = do_request(get_employees(run_arg, year_arg), com)
employeesElements = pdf_tree.findall('.//cs:EmployeeIdNumber', namespaces=ns)
employees = list(
    map(lambda c: Employee(c.find('./cs:EmployeeId', namespaces=ns).text,
                           c.find('./cs:EmployeeNumber', namespaces=ns).text),
        employeesElements))
spinner.succeed(f"Found {len(employees)} employees for run {run_info.description}")

zip_file_name = f"run_{run_info.number}_{year_arg}.zip"
with ShadyBar("Fetching salary slip for employees", max=len(employees)) as bar:
    with ZipFile(zip_file_name, 'w') as zip_file:
        for i, employee in enumerate(employees):
            employee_response = do_request(get_employee(employee.id), emp)
            employee_details = to_employee_details(employee_response)

            pdf_tree = do_request(get_payslip(employee.id, run_arg, year_arg), com)
            pdf = pdf_tree.find('.//cs:PDF', namespaces=ns)
            pdfBinary = base64.b64decode(pdf.text)

            zip_file.writestr(f"{employee_details.email}/{year_arg}_{run_info.number}_Salary.pdf", pdfBinary)
            bar.next()

spinner.succeed(f"Wrote results to {zip_file_name}")
