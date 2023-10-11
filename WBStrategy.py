import json
import time
from urllib.request import urlopen
import urllib.request
import openpyxl
import pandas as pd
from functools import lru_cache
import numpy as np
import traceback

def get_jsonparsed_data(url):
    response = urlopen(url)
    data = response.read().decode("utf-8")
    return json.loads(data)

def weight(list):
    return [1 / (i ** (0.5)) / sum([1 / (i ** (0.5)) for i in range(1, len(list) + 1)]) * 100 for i in range(1, len(list) + 1)]

@lru_cache
def afficher_barre_chargement(pourcentage):
    bar_length = 100  # Longueur totale de la barre de chargement
    blocs_charges = int(round(bar_length * pourcentage / 100))

    # Créer la barre de chargement en utilisant le caractère '=' pour les blocs chargés et '-' pour les blocs vides
    barre = '=' * blocs_charges + '-' * (bar_length - blocs_charges)

    # Afficher la barre de chargement et le pourcentage
    print(f"[{barre}] {pourcentage}%")

df = pd.read_excel('Actions.xlsx')

actions = list(get_jsonparsed_data("https://financialmodelingprep.com/api/v3/financial-statement-symbol-lists?apikey"
                                   "=ce82b6a14287d6b24fdcaf5468401b12"))

Indispo = ["8109.TWO","PHA.PA","SMTPC.PA","DGE.PA","HMI.DE","MSF.DE","IXD1.DE","FAS.DE","NOVC.DE","COZ.DE","ADB.DE",
           "ABEC.DE","ZOE.DE","2M6.DE","ABEA.DE","AKX.DE","CBHD.DE","FB2A.DE","MSF.F","NNND.F","ABEA.F","ABEC.F",
           "HMSB.DE","ORC.DE","XIX.DE","CCC3.DE","JNJ.DE","PCX.DE","NKE.DE","6MK.DE","ABL.DE","ABT.SW","BBY.DE","RTO",
           "LLY.DE","MMM.DE","PHM7.DE","SYK.DE","CPA.DE","3RB.DE","4I1.DE","ABC.L","FB2A.F","HMB.SW","MOV.F","RKLIF",
           "FLTR.L","BRM.DE","GSK.SW","CIS.DE","GS7.DE","MXI.DE","BZ7A.F","CIS.BR","FFV.DE","ATCO-B.ST","ACO1.F",
           "BKNG.SW","CIT.DE","PCE1.DE","FPE.DE","FPE3.DE","ITU.DE","ASME.DE","MSF.BR","NNN1.VI","NNND.VI","RBGPF",
           "KO.SW","ORCL.SW","LLY.SW","MMM.SW","MRK.PA","NEMA.SW","IGE.PA","MEDI.OL","ORNAV.HE","RMS.VI","RAA.SW",
           "ORNAV.HE","PMI.SW","RB.SW","ROG.SW","M4I.DE","BLQA.DE","3V64.DE","SO.PA","SNGR.L","ZWACK.BD","ABCZF",
           "ACEHF","AJMPF","CHSYF","ANPDF","ANPDY","CEO","BBBY","BF-A","BOMXF","CLPBF","CLPBY","DOTDF","NHMAF","TVBCF",
           "MTHRF","MTHRY","KKKUF","EVGGF","EVVTY","SATLF","SRTTY","KOZAY","PLSQF","OCLCF","RTLLF","IDEXF","PNDZF",
           "IDEXY","TBNGY","NONOF","NVO","ORINY","MRPLY","SLVFF","HRGLF","HRGLY","KIROY","TCEHY","TCTZF","PANDY",
           "SAPGF","RTOXF","YAHOF","YAHOY","SAP","MDT","GLAXF","RBGLY","GSK","HMRZF","HNNMY","PITPY","PSGTY","TVBCY",
           "ALSER.PA","NOKIA.PA","HESAF","HESAY","4338.HK","5274.TWO","4333.HK","4966.TWO","MSFT.NE","6146.TWO",
           "5289.TWO","3611.TWO","FB","B1C.F"]


def analyse(ticker):
    global weight
    try:
        try:
            info = get_jsonparsed_data(
                f"https://financialmodelingprep.com/api/v3/profile/{ticker}?apikey=ce82b6a14287d6b24fdcaf5468401b12")[0]
        except:
            info = {"companyName": "NONE", "industry": "NONE", "country": "NONE"}
        financial_statement = get_jsonparsed_data(f"https://financialmodelingprep.com/api/v3/income-statement/{ticker}?limit=120&apikey=ce82b6a14287d6b24fdcaf5468401b12")
        balance_sheet = get_jsonparsed_data(f"https://financialmodelingprep.com/api/v3/balance-sheet-statement/{ticker}?apikey=ce82b6a14287d6b24fdcaf5468401b12")
        recommandation = get_jsonparsed_data(f"https://financialmodelingprep.com/api/v4/score?symbol={ticker}&apikey=ce82b6a14287d6b24fdcaf5468401b12")
        ligne = [ticker]
        nom = info["companyName"]
        pays = info["country"]

        for debut in range(max(0,min(len(balance_sheet),len(financial_statement))-10)):

            equity = []
            inventory = []
            debt = []
            operating_profit = []
            revenus_brut = []
            revenus_variation = []
            eps = []
            eps_variation = []
            stockholder_equity = []
            stockholder_equity_variation = []
            cash_and_cash_equivalents = []
            cash_and_cash_equivalents_variation = []
            retained_earnings = []
            retained_earnings_variation = []
            research_and_development_expenses = []
            cost_of_revenue = []
            total_current_liabilities = []
            total_current_assets = []
            net_income = []
            depreciationAndAmortization = []
            interestExpense = []
            assets = []
            investment = []
            cost_of_investment = []
            score = 0

            if info["companyName"] == "NONE":
                ligne.append(financial_statement[debut]["date"])
                ligne.append(0)
            elif info["industry"] in ["NONE", "", None]:
                ligne.append(financial_statement[debut]["date"])
                ligne.append(0)

            elif min(len(balance_sheet),len(financial_statement)) < 10:
                ligne.append(financial_statement[debut]["date"])
                ligne.append(0)
            elif int(financial_statement[0]["date"][:4])<2021:
                ligne.append(financial_statement[debut]["date"])
                ligne.append(0)

            else:
                for i in range(debut,min(len(financial_statement), len(balance_sheet))):
                    research_and_development_expenses.append(financial_statement[i]["researchAndDevelopmentExpenses"])
                    cost_of_revenue.append(financial_statement[i]["costOfRevenue"])
                    net_income.append(financial_statement[i]["netIncome"])
                    depreciationAndAmortization.append(financial_statement[i]["depreciationAndAmortization"] if financial_statement[i]["depreciationAndAmortization"] is not None else 0)
                    interestExpense.append(financial_statement[i]["interestExpense"] if financial_statement[i]["interestExpense"] is not None else 0)
                    operating_profit.append(financial_statement[i]["operatingIncome"] if financial_statement[i]["operatingIncome"] is not None else 0)
                    revenus_brut.append(financial_statement[i]["revenue"])
                    eps.append(financial_statement[i]["eps"])
                    investment.append(balance_sheet[i]["totalInvestments"])
                    debt.append(balance_sheet[i]["netDebt"] if balance_sheet[i]["netDebt"] is not None else 0)
                    inventory.append(balance_sheet[i]["inventory"] / max(1, revenus_brut[-1]))
                    stockholder_equity.append(balance_sheet[i]["totalStockholdersEquity"] if balance_sheet[i]["totalStockholdersEquity"] is not None else 0)
                    cash_and_cash_equivalents.append(balance_sheet[i]["cashAndCashEquivalents"])
                    equity.append(balance_sheet[i]["totalEquity"] if balance_sheet[i]["totalEquity"] is not None else 0)
                    assets.append(balance_sheet[i]["totalAssets"])
                    total_current_assets.append(balance_sheet[i]["totalCurrentAssets"])
                    total_current_liabilities.append(balance_sheet[i]["totalCurrentLiabilities"])
                    retained_earnings.append(balance_sheet[i]["retainedEarnings"])

                for i in range(len(revenus_brut) - 1):
                    cost_of_investment.append(investment[i]-investment[i + 1])
                    revenus_variation.append((revenus_brut[i] - revenus_brut[i + 1]) / max(1, revenus_brut[i + 1]))
                    eps_variation.append((eps[i] - eps[i + 1]) / max(1, eps[i + 1]))
                    stockholder_equity_variation.append((stockholder_equity[i] - stockholder_equity[i + 1]) / max(1, stockholder_equity[i + 1]))
                    cash_and_cash_equivalents_variation.append((cash_and_cash_equivalents[i] - cash_and_cash_equivalents[i + 1]) / max(cash_and_cash_equivalents[i + 1], 1))
                    retained_earnings_variation.append((retained_earnings[i] - retained_earnings[i + 1]) / max(retained_earnings[i + 1], 1)), cost_of_investment.append(investment[-1])


                #Interprétation du compte de résultat
                #Gross Margin Study \ Etude Marge Brute \ Studie zur Bruttomarge
                score += sum([int((x - y) / max(1, x) > 0.4) * weight for x, y, weight in zip(revenus_brut, cost_of_revenue, weight(revenus_brut))])
                score += sum([int((x - y) / max(1, x) > 0.2) * weight for x, y, weight in zip(revenus_brut, cost_of_revenue, weight(revenus_brut))])

                #Frais SG&A
                score += sum([int(max(0.1, x) / max(1, y) < 0.3) * weight for x, y, weight in zip(research_and_development_expenses, revenus_brut, weight(research_and_development_expenses))])
                score += sum([int(max(0.1, x) / max(1, y) < 0.8) * weight for x, y, weight in zip(research_and_development_expenses, revenus_brut, weight(research_and_development_expenses))])

                #Les dépréciations et amortissements
                score += sum([int(z / max(1,(x - y)) < 0.1) * weight for x, y, z, weight in zip(revenus_brut, cost_of_revenue, depreciationAndAmortization, weight(cost_of_revenue))])
                score += sum([int(z / max(1,(x - y)) < 0.2) * weight for x, y, z, weight in zip(revenus_brut, cost_of_revenue, depreciationAndAmortization, weight(cost_of_revenue))])

                #Les charges d'intérêts financiers
                score += sum([int(x / max(y,1) < 0.15) * weight for x, y, weight in zip(interestExpense, operating_profit, weight(interestExpense))])
                score += sum([int(x / max(y,1) < 0.25) * weight for x, y, weight in zip(interestExpense, operating_profit, weight(interestExpense))])

                #Marge nette
                if pays in ["FR", "DE", "AT", "BE", "BG", "CY", "HR", "DK", "ES", "EE", "FI", "GR", "HU", "IE", "IT", "LV", "LT", "LU", "MT", "NL", "GB"]:
                    score += sum([int(x / max(1, y) > 0.1) * weight for y, x, weight in zip(revenus_brut, net_income, weight(revenus_brut))])
                    score += sum([int(x / max(1, y) > 0.05) * weight for y, x, weight in zip(revenus_brut, net_income, weight(revenus_brut))])

                else:
                    score += sum([int(x / max(1, y) > 0.2) * weight for y, x, weight in zip(revenus_brut, net_income, weight(revenus_brut))])
                    score += sum([int(x / max(1, y) > 0.1) * weight for y, x, weight in zip(revenus_brut, net_income, weight(revenus_brut))])

                #Interprétation des bilans
                #Stocks
                vInventory = []
                for year in range(len(inventory)-5):
                    vInventory.append(100*np.std([inventory[year],inventory[year+1],inventory[year+2],inventory[year+3],inventory[year+4]], ddof=1))  # Calcul de la volatilité sur 5 années glissantes
                score += sum([int(x<10) * weight for x, weight in zip(vInventory, weight(vInventory))])
                score += sum([int(x<15) * weight for x, weight in zip(vInventory, weight(vInventory))])

                #Actifs
                if assets[0]>5_000_000_000:
                    score += 200
                elif assets[0]>1_000_000_000:
                    score += 100

                #ROA
                score += sum([int(x / max(1,y) > 0.15) * weight for x, y, weight in zip(net_income, assets, weight(net_income))])
                score += sum([int(x / max(1,y) > 0.10) * weight for x, y, weight in zip(net_income, assets, weight(net_income))])

                #Debt
                score += sum([int(x / max(1,(y+z)) < 3) * weight for x, y, z, weight in zip(debt, operating_profit, depreciationAndAmortization , weight(depreciationAndAmortization))])
                score += sum([int(x / max(1,(y+z)) < 4) * weight for x, y, z, weight in zip(debt, operating_profit, depreciationAndAmortization , weight(debt))])

                #Fonds propres
                score += sum([int(x / max(1,y) < 3) * weight for x, y, weight in zip(debt, equity , weight(equity))])

                #Retained earning
                score += 2*sum([int(x > 0) * weight for x, weight in zip(retained_earnings_variation, weight(retained_earnings_variation))])

                #ROE
                score += sum([int(x/max(1,y) > 0.2) * weight for x, y, weight in zip(net_income, equity, weight(net_income))])
                score += sum([int(x/max(1,y) > 0.1) * weight for x, y, weight in zip(net_income, equity, weight(net_income))])

                #Flux de trésorerie
                #Free ashs flow
                score += sum([int(x > 0) * weight for x, weight in zip(cash_and_cash_equivalents, weight(cash_and_cash_equivalents))])

                #Investment
                score += sum([int(max(0,x)/max(1,y) < 0.25) * weight for x, y, weight in zip(cost_of_investment, net_income, weight(net_income))])
                score += sum([int(max(0,x)/max(1,y) < 0.50) * weight for x, y, weight in zip(cost_of_investment, net_income, weight(net_income))])

                score += sum([int(x + y > z) * weight for x, y, z, weight in zip(total_current_assets, cash_and_cash_equivalents, total_current_liabilities, weight(total_current_assets))])
                score += sum([int(x > 0) * weight for x, weight in zip(revenus_variation, weight(revenus_variation))])
                score += sum([int(x > 0) * weight for x, weight in zip(eps_variation, weight(eps_variation))])
                score += sum([int(x > 0) * weight for x, weight in zip(stockholder_equity_variation, weight(stockholder_equity_variation))])
                score += sum([int(x > 0) * weight for x, weight in zip(cash_and_cash_equivalents_variation, weight(cash_and_cash_equivalents_variation))])
                score += sum([int(x / max(y,1) > 2) * weight for x, y, weight in zip(total_current_assets, total_current_liabilities, weight(total_current_assets))])
                score += sum([int(x / max(1, (y + z)) > 0.12) * weight for x, y, z, weight in zip(net_income, total_current_liabilities, total_current_assets, weight(net_income))])

                score=round(score/33)

                ligne.append(financial_statement[debut]["date"])
                ligne.append(score / 100)

        sheet.append(ligne)

        workbook.save("Actions.xlsx")

        print(f"\033[32m {ticker} Nom : {nom.ljust(20)[:20]:<25} \033[0m")

    except urllib.error.HTTPError as e:
        if e.code == 502:
            print("HTTP Error 502: Bad Gateway")
        traceback.print_exc()
    except IndexError:
        traceback.print_exc()


try:
    workbook = openpyxl.load_workbook("Actions.xlsx")
    sheet = workbook.active
except:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Ticker", "Date", "Score", "Nom", "Secteur", "Pays"])

#actions = ['EVO.ST', 'RMS.PA', 'RAA.DE', '3092.T', 'ITX.MC', 'BF-B', '600519.SS', '603369.SS', '000858.SZ', '002304.SZ', '000568.SZ', '603896.SS', 'CSLLY', '600436.SS', '002821.SZ', '3613.HK', '0867.HK', 'FAE.MC', 'NOVO-B.CO', 'ORNBV.HE', '4587.T', 'BVXP.L', 'TECH', '603658.SS', '002932.SZ', '603387.SS', '002022.SZ', 'DIA.MI', '1858.HK', '7730.T', 'ANSS', 'LVC.WA', 'DOTD.L', 'MSFT', '2413.T', 'GOOG', '603444.SS', '3901.T', '2371.T', '4684.T', '002415.SZ', '1523.HK', '603203.SS', '002222.SZ', '600563.SS', '002690.SZ', 'FAST', '603568.SS', 'FDS']


for nombre, i in enumerate(actions):
    afficher_barre_chargement(round(nombre * 100 / len(actions)))
    if i not in Indispo + list(df['Ticker']):
            try:
                analyse(i)
            except Exception as e:
                print("Une erreur s'est produite :", e)
                traceback.print_exc()
