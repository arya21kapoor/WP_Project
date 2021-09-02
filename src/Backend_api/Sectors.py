from flask import Flask
from flask_restful import Resource, Api, reqparse
import pandas as pd
import ast
import os
from datetime import datetime

# Initialising the flask api
app = Flask(__name__)
api = Api(app)

# Created Sector_company as an endpoint
#This class will be returning data about each company in it's respective sector

class Sector_company(Resource):
    def get(self):

        print("Sector_company class running...")
        # Initializing the parser
        parser = reqparse.RequestParser()
        # Adding parser arguments
        parser.add_argument('sector_name', required=False)
        parser.add_argument('company_name', required=False)
        parser.add_argument('document_type', required=False)
        # Creating argument dictionary
        args = parser.parse_args()
        # sector_list=["Automobile","Constructon","Finance","FMCG","HealthCare","Metals_Chemicals","Power","Technology"]
        sector_file_list = list(os.listdir('D:\WP_project\src\csv_files'))
        # print(sector_file_list)
        sector = args['sector_name']
        company = args['company_name']
        document = args['document_type']
        # Created list of csv_files under each sector
        csv_list = []
        return_df = pd.DataFrame()
        # If else ladder to identify the secotr
        if(sector == "Automobile"):
            # To find the list of csv files under each seactors directory
            csv_list = list(os.listdir(
                'D:\WP_project\src\csv_files\Automobile'))
            # Retreiving the company name from the url and appending ".xlsx" to make the path
            comp = company + ".xlsx"
            if comp in csv_list:
                # Created a string path
                string_path = 'D:\WP_project\src\csv_files\Automobile\\'+comp
                print("File path:", string_path)
                # pd.set_option('display.max_rows',None,'display.max_columns',None)
                # This is the principle data sheet from where different data has been fetched
                data_sheet_df = pd.read_excel(
                    string_path, sheet_name="Data Sheet")

                if(document == "Profit Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                    return_df = profit_loss_df
                elif(document == "Balance Sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance sheet:")
                    return_df = balance_sheet_df
                    return_df = balance_sheet_df
                elif(document == "Cashflow"):
                    cashflow_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=80, skipfooter=8)
                    print("Cashflow Sheet:")
                    return_df = cashflow_df
                    return_df = cashflow_df
                else:
                    return {"data": "Invalid document"}

            else:
                return {"data": "Invalid company"}

        elif(sector == "Finance"):
            csv_list = list(os.listdir('D:\WP_project\src\csv_files\Finance'))
            print("List of csv files:", csv_list)
            comp = company + ".xlsx"
            if comp in csv_list:
                string_path = 'D:\WP_project\src\csv_files\Finance\\'+comp

                # pd.set_option('display.max_rows',None,'display.max_columns',None)
                data_sheet_df = pd.read_excel(
                    string_path, sheet_name="Data Sheet")
                #print("Data Sheet df:")
                # print(data_sheet_df)
                if(document == "Profit Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                    return_df = profit_loss_df
                elif(document == "Balance Sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance sheet:")
                    return_df = balance_sheet_df
                    return_df = balance_sheet_df
                elif(document == "Cashflow"):
                    cashflow_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=80, skipfooter=8)
                    print("Cashflow Sheet:")
                    return_df = cashflow_df
                    return_df = cashflow_df
                else:
                    return {"data": "Invalid document"}
            else:
                return {"data": "Invalid company"}
        elif(sector == "FMCG"):
            csv_list = list(os.listdir('D:\WP_project\src\csv_files\FMCG'))
            print("List of csv files:", csv_list)
            comp = company + ".xlsx"
            if comp in csv_list:
                string_path = 'D:\WP_project\src\csv_files\FMCG\\'+comp
                print("File path:", string_path)
                # pd.set_option('display.max_rows',None,'display.max_columns',None)
                data_sheet_df = pd.read_excel(
                    string_path, sheet_name="Data Sheet")
                #print("Data Sheet df:")
                # print(data_sheet_df)
                if(document == "Profit Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                    return_df = profit_loss_df
                elif(document == "Balance Sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance sheet:")
                    return_df = balance_sheet_df
                    return_df = balance_sheet_df
                elif(document == "Cashflow"):
                    cashflow_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=80, skipfooter=8)
                    print("Cashflow Sheet:")
                    return_df = cashflow_df
                    return_df = cashflow_df
                else:
                    return {"data": "Invalid document"}

            else:
                return {"data": "Invalid documnet"}
        elif(sector == "HealthCare"):
            csv_list = list(os.listdir(
                'D:\WP_project\src\csv_files\HealthCare'))
            print("List of csv files:", csv_list)
            comp = company + ".xlsx"
            if comp in csv_list:
                string_path = 'D:\WP_project\src\csv_files\HealthCare\\'+comp
                print("File path:", string_path)
                # pd.set_option('display.max_rows',None,'display.max_columns',None)
                data_sheet_df = pd.read_excel(
                    string_path, sheet_name="Data Sheet")
                #print("Data Sheet df:")
                # print(data_sheet_df)
                if(document == "Profit Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                    return_df = profit_loss_df
                elif(document == "Balance Sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance sheet:")
                    return_df = balance_sheet_df
                    return_df = balance_sheet_df
                elif(document == "Cashflow"):
                    cashflow_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=80, skipfooter=8)
                    print("Cashflow Sheet:")
                    return_df = cashflow_df
                    return_df = cashflow_df
                else:
                    return {"data": "Invalid document"}
            else:
                return {"data": "Invalid company"}
        elif(sector == "Metals_Chemicals"):
            csv_list = list(os.listdir(
                'D:\WP_project\src\csv_files\Metals_Chemicals'))
            print("List of csv files:", csv_list)
            comp = company + ".xlsx"
            if comp in csv_list:
                string_path = 'D:\WP_project\src\csv_files\Metals_Chemicals\\'+comp
                print("File path:", string_path)
                # pd.set_option('display.max_rows',None,'display.max_columns',None)
                data_sheet_df = pd.read_excel(
                    string_path, sheet_name="Data Sheet")
                #print("Data Sheet df:")
                # print(data_sheet_df)
                if(document == "Profit Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                    return_df = profit_loss_df
                elif(document == "Balance Sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance sheet:")
                    return_df = balance_sheet_df
                    return_df = balance_sheet_df
                elif(document == "Cashflow"):
                    cashflow_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=80, skipfooter=8)
                    print("Cashflow Sheet:")
                    return_df = cashflow_df

                else:
                    return {"data": "Invalid document"}
            else:
                return {"data": "Invalid company"}
        elif(sector == "Construction"):
            csv_list = list(os.listdir(
                'D:\WP_project\src\csv_files\Construction'))
            print("List of csv files:", csv_list)
            comp = company + ".xlsx"
            if comp in csv_list:
                string_path = 'D:\WP_project\src\csv_files\Construction\\'+comp
                print("File path:", string_path)
                # pd.set_option('display.max_rows',None,'display.max_columns',None)
                data_sheet_df = pd.read_excel(
                    string_path, sheet_name="Data Sheet")
                #print("Data Sheet df:")
                # print(data_sheet_df)
                if(document == "Profit Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                elif(document == "Balance Sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance sheet:")
                    return_df = balance_sheet_df
                elif(document == "Cashflow"):
                    cashflow_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=80, skipfooter=8)
                    print("Cashflow Sheet:")
                    return_df = cashflow_df
                else:
                    return {"data": "Invalid document"}
            else:
                return {"data": "Invalid company"}

        elif(sector == "Power"):
            csv_list = list(os.listdir('D:\WP_project\src\csv_files\Power'))
            print("List of csv files:", csv_list)
            comp = company + ".xlsx"
            if comp in csv_list:
                string_path = 'D:\WP_project\src\csv_files\Power\\'+comp
                print("File path:", string_path)
                # pd.set_option('display.max_rows',None,'display.max_columns',None)
                data_sheet_df = pd.read_excel(
                    string_path, sheet_name="Data Sheet")
                #print("Data Sheet df:")
                # print(data_sheet_df)
                if(document == "Profit Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                elif(document == "Balance Sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance sheet:")
                    return_df = balance_sheet_df
                elif(document == "Cashflow"):
                    cashflow_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=80, skipfooter=8)
                    print("Cashflow Sheet:")
                    return_df = cashflow_df
                else:
                    return {"data": "Invalid document"}
            else:
                return {"data": "Invalid company"}

        elif(sector == "Technology"):
            csv_list = list(os.listdir(
                'D:\WP_project\src\csv_files\Technology'))
            print("List of csv files:", csv_list)
            comp = company + ".xlsx"
            print(comp)
            if comp in csv_list:
                string_path = 'D:\WP_project\src\csv_files\Technology\\'+comp
                print("File path:", string_path)
                # pd.set_option('display.max_rows',None,'display.max_columns',None)
                data_sheet_df = pd.read_excel(
                    string_path, sheet_name="Data Sheet")
                #print("Data Sheet df:")
                # print(data_sheet_df)
                if(document == "Profit Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                elif(document == "Balance Sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance sheet:")
                    return_df = balance_sheet_df
                elif(document == "Cashflow"):
                    cashflow_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=80, skipfooter=8)
                    print("Cashflow Sheet:")
                    return_df = cashflow_df
                else:
                    return {"data": "Invalid document"}
                return {"data": True}
            else:
                return {"data": False}

        else:
            return {"data": "Invalid sector"}

        columns = list(return_df.columns.values)
        print("Columns:", columns)

        # rows=list(return_df.rows)
        no_rows = len(return_df.index)
        return_list = []
        for i in range(0, no_rows):
            return_dict = {}
            for j in range(1, len(columns)):
                # column_name=column[j]
                #print("i:",i," j:",j)
                # print(return_df.iloc[i][j])
                date = columns[j].strftime("%m/%d/%y")
                return_dict[date] = return_df.iloc[i][j]
            print("Return_dict", (i+1), ":", return_dict)
            return_list.append(return_dict)
        return {"data": return_list}

        #print("Return list:",return_list)
        # Returning the csv_files_dictionary
        # return {'data':csv_list}
# Added this class for the endpoint /sectors
api.add_resource(Sector_company, '/sectors')

class Sector_rankings(Resource):
    def get(self):
        #Initialsing the parser
        parser=reqparse.RequestParser()
        #Reading arguments from the url
        parser.add_argument('sector_name',required=True)
        #Creating argument dictionary
        args=parser.parse_args()
        #Fetching sector name
        sector=args['sector_name']
        sector_list=list(os.listdir('D:\WP_project\src\csv_files'))
        if sector in sector_list:
            print("Sector name:",sector)
            path="D:\WP_project\src\csv_files\\"+sector
            csv_list=list(os.listdir(path))
            market_cap_list=[]
            pe_ratio_list=[]
            profit_growth_list=[]
            for i in csv_list:
                company_path=path+"\\"+i
                company_df=pd.read_excel(company_path,sheet_name="Data Sheet")
                print("Company df:",company_df)
                market_cap=company_df.loc[7].iat[1]
                print("Market cap:",market_cap)
                market_cap_list.append(market_cap)
                pe_denominator=company_df.loc[28].iat[10]
                print("PE denominator:",pe_denominator)
                pe_ratio=(market_cap/pe_denominator)
                pe_ratio_list.append(pe_ratio)
                j_30=company_df.loc[28].iat[9]
                profit_growth=(pe_denominator - j_30)/j_30
                profit_growth_list.append(profit_growth)

            print("Market cap list:",market_cap_list)
            print("PE ratio list:",pe_ratio_list)
            print("Profit growth list:",profit_growth_list)
        
        else:
            return {"data":"Invalid sector"}
api.add_resource(Sector_rankings,'/sector_rankings')

if __name__ == "__main__":
    app.run()
