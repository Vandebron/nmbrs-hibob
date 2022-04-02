#!/usr/bin/env python3
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
parser.add_argument('--user', '-u', help='API user', default='IT@vandebron.nl')
parser.add_argument('--token', '-t', help='The API token https://support.nmbrs.com/hc/en-us/articles/'
                                          '360013384371-Nmbrs-API', required=True)
parser.add_argument('--run', '-r', help='The run to download. Prints all runs for the year if not set.', type=int)
parser.add_argument('--year', '-y', help='The year in which the run took place', required=True, type=int)
parser.add_argument('--annual', '-a', help='Fetch annual statements for the employees in a run', dest='annual',
                    action='store_true', default=False)
parser.add_argument('--company', '-c', help='Select specific company number')
parser.add_argument('--description', '-d', help='Will be appended to the PDFs instead of the run description')
parser.add_argument('--email', '-e', help='Indicates whether to use company email as folder name. Otherwise employee '
                                          'Id will be used', dest='email', action='store_true', default=False)
args = parser.parse_args()

user_arg = args.user
password_arg = args.token
run_arg = args.run
year_arg = args.year
company_id = args.company
description_arg = args.description
email_arg = args.email
annual_arg = args.annual


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
    startDate: str
    endDate: str


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


def get_all_employees():
    return f"""<emp:Employment_GetAll_AllEmployeesByCompany >
      <emp:CompanyID>{company_id}</emp:CompanyID>
    </emp:Employment_GetAll_AllEmployeesByCompany>"""


def get_payslip(employee_id, run_id, year):
    return f"""<com:SalaryDocuments_GetEmployeePayslipsPDFByRunCompany_v2>
         <com:EmployeeID>{employee_id}</com:EmployeeID>
         <com:CompanyId>{company_id}</com:CompanyId>
         <com:RunID>{run_id}</com:RunID>
         <com:intYear>{year}</com:intYear>
      </com:SalaryDocuments_GetEmployeePayslipsPDFByRunCompany_v2>"""


def get_annual_statement(employee_id, year):
    return f"""<emp:SalaryDocuments_GetAnnualStatementPDF>
         <emp:EmployeeID>{employee_id}</emp:EmployeeID>
         <emp:intYear>{year}</emp:intYear>
      </emp:SalaryDocuments_GetAnnualStatementPDF>"""


def get_employee(employee_id):
    return f"""<emp:PersonalInfo_GetCurrent>
         <emp:EmployeeId>{employee_id}</emp:EmployeeId>
      </emp:PersonalInfo_GetCurrent>"""


def do_request(body, service, print_response=False):
    req = create_request(user_arg, password_arg, body, service)
    # print(req)
    response = requests.post(f"{service['url']}.asmx", data=req,
                             headers={'content-type': 'text/xml; charset=utf-8'})
    if not response.ok:
        sys.exit(response.text)
    if print_response:
        print(response.text)
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
    run_id = element.find('./cs:ID', namespaces=ns).text
    number = element.find('./cs:Number', namespaces=ns).text
    description = element.find('./cs:Description', namespaces=ns).text
    period_start = element.find('./cs:PeriodStart', namespaces=ns).text
    period_end = element.find('./cs:PeriodEnd', namespaces=ns).text
    return RunInfo(run_id, number, description, period_start, period_end)


def to_employee_details(tree) -> EmployeeDetails:
    element = tree.find('.//emp:PersonalInfo_GetCurrentResult', namespaces=ns)
    employee_id = element.find('./emp:Id', namespaces=ns).text
    number = element.find('./emp:Number', namespaces=ns).text
    email_element = element.find('./emp:EmailWork', namespaces=ns)
    email = number if email_element is None else email_element.text
    return EmployeeDetails(employee_id, number, email)


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


def find_employees_for_run(run, year, run_info_arg):
    global spinner
    spinner.text = f"Finding employees for {run_info_arg.number} {run_info_arg.description}"
    employee_tree = do_request(get_employees(run, year), com)
    employees_elements = employee_tree.findall('.//cs:EmployeeIdNumber', namespaces=ns)
    employees = list(
        map(lambda c: Employee(c.find('./cs:EmployeeId', namespaces=ns).text,
                               c.find('./cs:EmployeeNumber', namespaces=ns).text, "", ""),
            employees_elements))
    spinner.succeed(f'Found {len(employees)} employees for run "{run_info_arg.description}"')
    return employees


def find_employees_for_year(year):
    global spinner
    spinner.text = f"Finding employees active in {year}"
    employee_tree = do_request(get_all_employees(), emp)
    employees_elements = employee_tree.findall('.//emp:EmployeeEmploymentItem', namespaces=ns)
    all_employees = list(
        map(lambda c: Employee(c.find('./emp:EmployeeId', namespaces=ns).text, "",
                               c.find('.//emp:StartDate', namespaces=ns).text,
                               None if c.find('.//emp:EndDate', namespaces=ns) is None else c.find('.//emp:EndDate',
                                                                                                   namespaces=ns).text),
            employees_elements))
    this_year = f'{year}-01-01T00:00:00'
    next_year = f'{year + 1}-01-01T00:00:00'
    active_employees = list(
        filter(lambda c: c.startDate < next_year and (c.endDate is None or c.endDate > this_year), all_employees))
    spinner.succeed(f'Found {len(active_employees)} employees with active contract in {year}')
    return active_employees


def fetch_annual_statements(year):
    employees = find_employees_for_year(year)
    zip_file_name = f"yearly_statements_{year}.zip"
    with ShadyBar("Fetching annual statement for employees", max=len(employees)) as bar:
        with ZipFile(zip_file_name, 'w') as zip_file:
            for i, employee in enumerate(employees):
                employee_response = do_request(get_employee(employee.id), emp)
                employee_details = to_employee_details(employee_response)

                request_body = get_annual_statement(employee.id, year)

                response = do_request(request_body, emp)
                pdf = response.find(f'.//emp:PDF', namespaces=ns)

                if pdf is not None:
                    description = "Annual_Statement" if description_arg is None else description_arg
                    folder_name = employee_details.email if email_arg else employee_details.number
                    file_name = f"{folder_name}/{year}_{description}.pdf"
                    zip_file.writestr(file_name, base64.b64decode(pdf.text))

                bar.next()
        spinner.succeed(f"Wrote annual statements to {zip_file_name}")


def fetch_salary_slips(year, run):
    run_info = get_run_info(run)
    employees = find_employees_for_run(run, year, run_info)
    zip_file_name = f"run_{run_info.number}_{year}.zip"
    with ShadyBar("Fetching salary slip for employees", max=len(employees)) as bar:
        with ZipFile(zip_file_name, 'w') as zip_file:
            for i, employee in enumerate(employees):
                employee_response = do_request(get_employee(employee.id), emp)
                employee_details = to_employee_details(employee_response)

                request_body = get_payslip(employee.id, run_arg, year)
                response = do_request(request_body, com)
                pdfs = response.findall(f'.//cs:PDF', namespaces=ns)

                for index, pdf in enumerate(pdfs):
                    description = run_info.description.replace(' ', '_') if description_arg is None else description_arg
                    folder_name = employee_details.email if email_arg else employee_details.number
                    file_name = f"{folder_name}/{year}_{run_info.number}_{description}_{index}.pdf"
                    zip_file.writestr(file_name, base64.b64decode(pdf.text))

                bar.next()
        spinner.succeed(f"Wrote payslips to {zip_file_name}")


if annual_arg:
    fetch_annual_statements(year_arg)
else:
    fetch_salary_slips(year_arg, run_arg)
