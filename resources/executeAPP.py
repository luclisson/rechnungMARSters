import random as r
import pandas as pd
from resources.usedGood import UsedGood
from resources.mitarbeiterKosten import MitarbeiterKosten
import datetime
import os
import json

#functions
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
        print(obj.gesamtpreis)
    return output
def genGoodsTable(filePath):
    usedGoods = getUsedGood(filePath)
    tableHead = r"""\textbf{Pos} & \textbf{Leistung}  & \textbf{Einzelpreis} & \textbf{Anzahl} & \textbf{Gesamtpreis} \\
                    \hline
                """
    content = ""
    for i in range(len(usedGoods)):
        content += fr"""
            {i+1} & {usedGoods[i].leistung} &  {usedGoods[i].einzelpreis:.2f} EUR & {usedGoods[i].anzahl} & {usedGoods[i].gesamtpreis:.2f} EUR \\
            \hline
        """

    with open('resources/texComps/goodsTable.tex','w') as outputFile:
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
                \\textbf{{Gesamtbetrag:}} & \\textbf{{{data[2]:.2f}}} EUR \\\\
                """
    with open('resources/texComps/gesBetraege.tex','w')as file:
        file.write(content)
def genRandomCustomer(filePath):
    rechnungsnummer = r.randint(1000000000,3999999999)
    with open('resources/letzteRechnungsnummer.json','w') as json_file:
        json.dump({"num":rechnungsnummer},json_file)
    kundennummer = r.randint(10000,99999)
    datum = datetime.datetime.now().strftime("%d.%m.%Y")
    content = fr"""
                Offer Nr. {rechnungsnummer} & Customer Nr.: {kundennummer} & Date: {datum} \\
                """
    with open('resources/texComps/customerData.tex','w') as file:
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
    with open('resources/texComps/customerAnschrift.tex','w') as file:
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
    costArr = getEmployeeCost(filePath)
    header= r"""
                \begin{tabularx}{\textwidth}{|l|X|l|r|r|}
                \hline
                \rowcolor{gray!30}
                \textbf{Pos} & \textbf{Leistung} & \textbf{Stunden} & \textbf{Satz/h} & \textbf{Gesamtkosten} \\
                \hline 
            """
    content = ""

    for i in range(len(costArr)):
        content +=fr"""
                    {i+1} & {costArr[i].leistung} & {costArr[i].stunden} & {costArr[i].satz:.2f} EUR & {costArr[i].gesamtkosten:.2f} EUR \\
                    \hline 
                    """
    fotter = r"""
                \end{tabularx} 
            """
    with open('resources/texComps/mitarbeiterKosten.tex','w') as file:
        file.write(header+"\n"+content+"\n"+fotter)
def execute(filePath):
    genRandomCustomer(filePath)
    genCostsTable(filePath)
    genGoodsTable(filePath)
    genEmployeeCostTable(filePath)

    #folders
    project_folder = os.getcwd()
    output_dir = os.path.join(project_folder, "output")
    resources_dir = os.path.join(project_folder, "resources")
    os.chdir(resources_dir)

    output_dir_relative = "../output"
    file_path_relative = "rechnung.tex"
    
    with open('letzteRechnungsnummer.json','r') as json_file:
        data = json.load(json_file)
        num = data["num"]
    
    jobname = f"Rechnung_{num}"
    command = f'xelatex -output-directory={output_dir_relative} -jobname={jobname} {file_path_relative}'

    #command execution
    os.system(command)
    os.chdir(project_folder)
