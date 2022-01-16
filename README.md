# Nmbrs to Hibob salary slip export

![./resources/nmbrs_hibob.png](./resources/nmbrs_hibob_64x64.png)
```
usage: nmbrs_hibob [-h] [--user USER] --token TOKEN [--run RUN] --year YEAR [--company COMPANY]

Export salary slip PDFs from Visma Nmbrs into Hibob

optional arguments:
  -h, --help            show this help message and exit
  --user USER, -u USER  API user
  --token TOKEN, -t TOKEN
                        The API token https://support.nmbrs.com/hc/en-us/articles/360013384371-Nmbrs-API
  --run RUN, -r RUN     The run to download. Prints all runs for the year if not set.
  --year YEAR, -y YEAR  The year in which the run took place
  --company COMPANY, -c COMPANY
                        Select specific company number
```