import random as r
import pandas as pd
from usedGood import UsedGood
from mitarbeiterKosten import MitarbeiterKosten
import datetime

with open('rechnung.tex','r',encoding='utf-8') as texFile:
    tex = texFile.read()
def getUsedGood(filePath):
    data = pd.read_excel(filePath)
    output = []
    for row in range(59,69):
        rowData = data.iloc[row]
        obj = UsedGood(
            leistung=rowData[1],
            mwst=rowData[2],
            einzelpreis=rowData[3],
            anzahl=rowData[4],
            gesamtpreis=rowData[5]
        )
        output.append(obj)
    return output
def genGoodsTable(filePath):
    usedGoods = getUsedGood(filePath)
    tableHead = r"\textbf{Pos} & \textbf{Leistung} & \textbf{MwSt.} & \textbf{Einzelpreis} & \textbf{Anzahl} & \textbf{Gesamtpreis} \\"
    content = ""
    for i in range(len(usedGoods)):
        content += fr"""
            \hline
            {i+1} & {usedGoods[i].leistung} & {usedGoods[i].mwst}\% & {usedGoods[i].einzelpreis} EUR & {usedGoods[i].anzahl} & {usedGoods[4].gesamtpreis} EUR \\
        """
    content+= r'\hline'
    
    with open('goodsTable.tex','w') as outputFile:
        outputFile.write(tableHead + content)
def getCosts(filePath):
    data = pd.read_excel(filePath)
    netto = round(data.iloc[49][1], 2)
    mwst = round((data.iloc[50][1]/100) * netto, 2)
    brutto = round(data.iloc[51][1], 2)
    return [netto, mwst, brutto]
def genCostsTable(filePath):
    data = getCosts(filePath)
    content = f"""
                \\textbf{{Nettobetrag:}} & {data[0]} EUR \\\\
                zzgl. 19\,\\% MwSt: & {data[1]} EUR \\\\
                \\textbf{{Gesamtbetrag:}} & \\textbf{data[2]} EUR \\\\
                """
    with open('gesBetraege.tex','w')as file:
        file.write(content)
def genRandomCustomer(filePath):
    rechnungsnummer = r.randint(1000000000,3999999999)
    kundennummer = r.randint(10000,99999)
    datum = datetime.datetime.now().strftime("%d.%m.%Y")
    content = fr"""
                Rechnung Nr. {rechnungsnummer} & Kunden-Nr.: {kundennummer} & Datum: {datum} \\
                """
    with open('customerData.tex','w') as file:
        file.write(content)

    data = pd.read_excel(filePath)
    ansprechpartner = data.iloc[4][8]
    firma = data.iloc[5][8]
    addresse = data.iloc[6][8]

    contentAddresse = fr"""
                        \textbf{{{firma}}} \\
                        {ansprechpartner} \\
                        {addresse}
                        """
    with open('customerAnschrift.tex','w') as file:
        file.write(contentAddresse) 
def getEmployeeCost(filePath):
    output=[]
    data = pd.read_excel(filePath)
    index = 73
    
    while True:
        if data.iloc[index][0] == "Summe netto":
            break
        leistung = data.iloc[index][0]
        stunde = data.iloc[index][1]
        satz = data.iloc[index][2]
        kosten = data.iloc[index][3]
        output.append(MitarbeiterKosten(leistung=leistung, stunden=stunde, satz=satz,gesamtkosten=kosten))
        index = index + 1
    return output

def genEmployeeCostTable(filePath):
    costArr = getEmployeeCost('excelData.xlsx')
    header= r"""
                \begin{longtable}{|c|p{6cm}|c|r|r|}
                \hline
                \rowcolor{gray!30}
                \textbf{Pos} & \textbf{Leistung} & \textbf{Stunden} & \textbf{Satz} & \textbf{Gesamtkosten} \\
                \hline
                \endfirsthead
            """
    content = ""

    for i in range(len(costArr)):
        content +=fr"""
                    \hline 
                    {i+1} & {costArr[i].leistung} & {costArr[i].stunden} & {costArr[i].satz} & {costArr[i].gesamtkosten} \\
                    """
    fotter = r"""
                \end{longtable} 
            """
    with open('mitarbeiterKosten.tex','w') as file:
        file.write(header+"\n"+content+"\n"+fotter)
#getEmployeeCost('excelData.xlsx')
genEmployeeCostTable('excelData.xlsx')
#genRandomCustomer('excelData.xlsx')
#genCostsTable('excelData.xlsx')
#genGoodsTable('excelData.xlsx')
#automatisch rechnung.tex rendern und in output ordner speichern