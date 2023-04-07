import json
import xlsxwriter
#from LDOSItem import majorItem, configItem

class JSONItem:
    def __init__(self, path):
        self.path = path
        self.json = self.loadJSON()
        self.rowItems = []
        self.headings = [{'header':'Common Name'}, {'header':'SAV Name'}, {'header':'Business Entity'},
                         {'header':'Sub Business Entity'}, {'header':'Product Type'},{'header':'Product ID'},
                         {'header':'Item Description'}, {'header':'Current Covered Status'},
                         {'header':'Current Install Status'}, {'header':'Qty'}, {'header':'LDOS Year'},
                         {'header':'2021'}, {'header':'2022'}, {'header':'2023'}, {'header':'2024'},
                         {'header':'2025'},  {'header':'Configuration'}
                         ]

    def loadJSON(self):
        f = open(self.path, 'r')
        data = json.load(f)
        f.close()
        return data

    def itemToRow(self, fileName='LDOS Info.xlsx'):
        path = '/Users/brsteine/Library/CloudStorage/OneDrive-Personal/Documents/Applications/LDOSReportBuilder/venv' \
               '/Output/' + str(fileName)

        BES =[
            'Cloud and Compute',
            'Cloud Networking',
            'Collaboration',
            'Enterprise Routing',
            'Enterprise Switching',
            'IOT',
            'Meraki',
            'Other',
            'Security',
            'Service Provider Routing',
            'Wireless'
        ]

        wb = xlsxwriter.Workbook(path)

        totalItems = len(self.json)
        currItem = 1

        for BE in BES:

            ws = wb.add_worksheet(BE)

            BEjson = list(filter(lambda x : x['businessEntity'] == BE, self.json))

            rownum = 1
            for i in BEjson:
                print(f'Writing {currItem} of {totalItems}')
                ws.set_row(rownum, None, None, {'level' : 2})
                yr = self.getLDOSYr(i['ldosYear'])
                row = [i['common'], i['SAV'], i['businessEntity'], i['subBusEntity'], i['productType'], i['PID'],
                       i['description'], i['coveredStatus'], i['installStatus'], '1', i['ldosYear'],
                       yr[0], yr[1], yr[2], yr[3], yr[4], i['configID']]
                ws.write_row(rownum, 0, row)

                rownum += 1
                currItem += 1
                if len(i['subconfig']) > 0:
                    for s in i['subconfig']:
                        ws.set_row(rownum, None, None, {'level' : 3, 'hidden': True})
                        yr = self.getLDOSYr(s['ldosYear'], s['qty'])
                        config = str(s['parentConfig']) + '.' + str(s['configID'])
                        row = ['', '', '', '', '', s['PID'], s['description'], s['coveredStatus'], s['installStatus'],
                               s['qty'], s['ldosYear'], yr[0], yr[1], yr[2], yr[3], yr[4], config
                               ]
                        ws.write_row(rownum, 0, row)
                        rownum += 1

            ws.add_table(f'A1:Q{rownum}', {'columns': self.headings})
        wb.close()


    def getLDOSYr(self, ldosYr, qty=1):
        years = []

        year = ldosYr

        if year != '2021' :
            years.append('')
        else:
            years.append(qty)

        if year != '2022' :
            years.append('')
        else:
            years.append(qty)

        if year != '2023' :
            years.append('')
        else:
            years.append(qty)

        if year != '2024' :
            years.append('')
        else:
            years.append(qty)

        if year != '2025' :
            years.append('')
        else:
            years.append(qty)

        return years




class LDOSRow:
    def __init__(self, common, SAV, BE, subBE, prodType, configID, PID, desc, coveredStatus, installStatus, ldosYR,
                 qty=1):
        self.common = common
        self.SAV = SAV
        self.BE = BE
        self.subBE = subBE
        self.prodType = prodType
        self.PID = PID
        self.desc = desc
        self.coveredStatus = coveredStatus
        self.installStatus = installStatus
        self.qty = qty
        self.ldosYR = ldosYR
        self.yr1 = 0
        self.yr2 = 0
        self.yr3 = 0
        self.yr4 = 0
        self.yr5 = 0
        self.configID = configID

        self.loadLDOSYr()

    def loadLDOSYr(self):
        if self.ldosYR != '':
            year = self.ldosYR

            if year == '2021' :
                self.yr1 = self.qty
            elif year == '2022' :
                self.yr2 = self.qty
            elif year == '2023' :
                self.yr3 = self.qty
            elif year =='2024' :
                self.yr4 = self.qty
            elif year =='2025' :
                self.yr5 = self.qty

    def toArray(self):
        test = list(self)






