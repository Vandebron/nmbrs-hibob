import argparse
import requests
import xml.etree.ElementTree as ET
import base64

parser = argparse.ArgumentParser(description='Discover MPL project.yml definition files')
parser.add_argument('--user', '-u', help=f'User', default='IT@vandebron.nl')
parser.add_argument('--token', '-t', help=f'The API token https://support.nmbrs.com/hc/en-us/articles/'
                                          f'360013384371-Nmbrs-API', required=True)
parser.add_argument('--run', '-r', help='The run to download. Prints all runs for the year if not set.')
parser.add_argument('--year', '-y', help='The year in which the run took place', required=True)
args = parser.parse_args()
user_arg = args.user
password_arg = args.token
run_arg = args.run
year_arg = args.year

company_id = 62880


def create_request(user, password, payload):
    template = f"""
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:com="https://api.nmbrs.nl/soap/v2.1/CompanyService">
    <soapenv:Header>
      <com:AuthHeader>
         <com:Username>{user}</com:Username>
         <com:Token>{password}</com:Token>
      </com:AuthHeader>
   </soapenv:Header>
   <soapenv:Body>
      {payload}
   </soapenv:Body>
</soapenv:Envelope>
"""
    return template


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


ns = {
    'soap': "http://schemas.xmlsoap.org/soap/envelope/",
    'xsi': "http://www.w3.org/2001/XMLSchema-instance",
    'xsd': "http://www.w3.org/2001/XMLSchema",
    'cs': 'https://api.nmbrs.nl/soap/v2.1/CompanyService'
}

if not run_arg:
    r = requests.post('https://api.nmbrs.nl/soap/v2.1/CompanyService.asmx', data=create_request(user_arg, password_arg,
                                                                                                get_runs(year_arg)),
                      headers={'content-type': 'text/xml; charset=utf-8'})

    tree = ET.ElementTree(ET.fromstring(r.text))
    run_infos = tree.findall('.//cs:RunInfo', namespaces=ns)
    for run_info in run_infos:
        number = run_info.find('./cs:ID', namespaces=ns).text
        description = run_info.find('./cs:Description', namespaces=ns).text
        period_start = run_info.find('./cs:PeriodStart', namespaces=ns).text
        period_end = run_info.find('./cs:PeriodEnd', namespaces=ns).text
        print(f'{number} {description} {period_start}-{period_end}')
    exit()

request = create_request(user_arg, password_arg, get_employees(run_arg, year_arg))
r = requests.post('https://api.nmbrs.nl/soap/v2.1/CompanyService.asmx', data=request,
                  headers={'content-type': 'text/xml; charset=utf-8'})

tree = ET.ElementTree(ET.fromstring(r.text))
employees = tree.findall('.//cs:EmployeeId', namespaces=ns)
for employee in employees:
    employee_id = employee.text
    if employee_id == "490957":
        print(employee_id)
        request = create_request(user_arg, password_arg, get_payslip(employee_id, run_arg, year_arg))
        r = requests.post('https://api.nmbrs.nl/soap/v2.1/CompanyService.asmx', data=request,
                          headers={'content-type': 'text/xml; charset=utf-8'})
        tree = ET.ElementTree(ET.fromstring(r.text))
        pdf = tree.find('.//cs:PDF', namespaces=ns)
        pdfBinary = base64.b64decode(pdf.text)
        with open(f'{employee_id}.pdf', "wb") as slip_file:
            slip_file.write(pdfBinary)