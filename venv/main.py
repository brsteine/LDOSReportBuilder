import logging
import sys
from TableBuilder import JSONItem


def do_something() :

    j = JSONItem('/Users/brsteine/Library/CloudStorage/OneDrive-Personal/Documents/Applications/LDOSReportBuilder'
                 '/venv/Data/itemsExample.json')
    entireJSON = j.json
    accounts = [
        "OAG",
        "DIR",
        "LCRA",
        "TWC",
        "DFPS",
        "HHS",
        "TxDOT",
        "DMV",
        "DPS",
        "CPA",
        "Other"
    ]

    for account in accounts:
        j.json = list(filter(lambda x : x['common'] == account, entireJSON))
        #j.itemToRow(f'LDOS Item - {account}.xlsx')


if __name__ == '__main__' :
    do_something()
