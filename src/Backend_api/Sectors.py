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
# This class will be returning data about each company in it's respective sector


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

                if(document == "Profit_Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                    return_df = profit_loss_df
                elif(document == "Balance_sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance_sheet:")
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
                if(document == "Profit_Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                    return_df = profit_loss_df
                elif(document == "Balance_sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance_sheet:")
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
                if(document == "Profit_Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                    return_df = profit_loss_df
                elif(document == "Balance_sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance_sheet:")
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
                if(document == "Profit_Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                    return_df = profit_loss_df
                elif(document == "Balance_sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance_sheet:")
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
                if(document == "Profit_Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                elif(document == "Balance_sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance_sheet:")
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
                if(document == "Profit_Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                elif(document == "Balance_sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance_sheet:")
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
                if(document == "Profit_Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                elif(document == "Balance_sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance_sheet:")
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
                if(document == "Profit_Loss"):

                    profit_loss_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=15, skipfooter=62)
                    print("Profit & Loss Sheet:")
                    return_df = profit_loss_df
                elif(document == "Balance_sheet"):
                    balance_sheet_df = pd.read_excel(
                        string_path, sheet_name="Data Sheet", header=55, skipfooter=21)
                    print("Balance_sheet:")
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
        name_list=list(return_df['Report Date'])
        print("Name list:",name_list)
        for i in range(0, no_rows):
            return_dict = {}
            return_parent_dict={}
            for j in range(1, len(columns)):
                # column_name=column[j]
                #print("i:",i," j:",j)
                # print(return_df.iloc[i][j])
                date = columns[j].strftime("%m/%d/%y")
                return_dict[date] = return_df.iloc[i][j]
            return_parent_dict[name_list[i]]=return_dict
            print("Return_dict", (i+1), ":", return_dict)
            return_list.append(return_parent_dict)
        return {"data": return_list}


        #print("Return list:",return_list)
        # Returning the csv_files_dictionary
        # return {'data':csv_list}
# Added this class for the endpoint /sectors
api.add_resource(Sector_company, '/sectors')


class Sector_rankings(Resource):
    def get(self):
        print("Sector_rankings class running...")
        # Initialsing the parser
        parser = reqparse.RequestParser()
        # Reading arguments from the url
        parser.add_argument('sector_name', required=True)
        parser.add_argument('rank_name', required=True)
        # Creating argument dictionary
        args = parser.parse_args()
        # Fetching sector name
        sector = args['sector_name']
        rank_type = args['rank_name']
        sector_list = list(os.listdir('D:\WP_project\src\csv_files'))
        if sector in sector_list:
            print("Sector name:", sector)
            path = "D:\WP_project\src\csv_files\\"+sector
            csv_list = list(os.listdir(path))
            unsorted_market_cap_list = []
            unsorted_pe_ratio_list = []
            unsorted_profit_growth_list = []
            mc_list = []
            pe_list = []
            pg_list = []
            for i in csv_list:
                t_market_cap = []
                t_pe_ratio = []
                t_pg = []

                company_path = path+"\\"+i
                name = i.replace('.xlsx', '')
                t_market_cap.append(name)
                t_pe_ratio.append(name)
                t_pg.append(name)
                company_df = pd.read_excel(
                    company_path, sheet_name="Data Sheet")
                #print("Company df:",company_df)
                market_cap = company_df.loc[7].iat[1]
                #print("Market cap:",market_cap)
                t_market_cap.append(market_cap)
                t_market_cap = tuple(t_market_cap)
                unsorted_market_cap_list.append(t_market_cap)
                mc_list.append(market_cap)
                pe_denominator = company_df.loc[28].iat[10]
                #print("PE denominator:",pe_denominator)
                pe_ratio = (market_cap/pe_denominator)
                t_pe_ratio.append(pe_ratio)
                t_pe_ratio = tuple(t_pe_ratio)
                unsorted_pe_ratio_list.append(t_pe_ratio)
                pe_list.append(pe_ratio)
                j_30 = company_df.loc[28].iat[9]
                profit_growth = (pe_denominator - j_30)/j_30
                t_pg.append(profit_growth)
                t_pg = tuple(t_pg)
                pg_list.append(profit_growth)
                unsorted_profit_growth_list.append(t_pg)

            print("\nMarket cap list:", unsorted_market_cap_list)
            print("\nPE ratio list:", unsorted_pe_ratio_list)
            print("\nProfit growth list:", unsorted_profit_growth_list)
            mc_sorted = sorted(mc_list, reverse=True)
            pe_sorted = sorted(pe_list, reverse=True)
            pg_sorted = sorted(pg_list, reverse=True)
            sorted_market_cap_list = [
                tuple for x in mc_sorted for tuple in unsorted_market_cap_list if tuple[1] == x]
            print("\nSorted makret cap list: ", sorted_market_cap_list)
            sorted_pe_ratio_list = [
                tuple for x in pe_sorted for tuple in unsorted_pe_ratio_list if tuple[1] == x]
            print("\nSorted pe ratio list: ", sorted_pe_ratio_list)
            sorted_profit_growth_list = [
                tuple for x in pg_sorted for tuple in unsorted_profit_growth_list if tuple[1] == x]
            print("\nSorted profit growth list: ", sorted_profit_growth_list)
            target_list = []
            if rank_type.lower() == "Market_cap".lower():
                target_list = list(sorted_market_cap_list)
            elif rank_type.lower() == "PE_ratio".lower():
                target_list = list(sorted_pe_ratio_list)
            elif rank_type.lower() == "Profit_growth".lower():
                target_list = list(sorted_profit_growth_list)
            else:
                return {"Error": "Invalid rank type"}

            return_list = []

            for i in range(0, len(target_list)):
                return_dict = {}
                t_item = target_list[i]
                return_dict["Name"] = t_item[0]
                return_dict["Value"] = t_item[1]
                return_dict["Ranking"]=(i+1)
                return_list.append(return_dict)

            # return {"Market Cap":sorted_market_cap_list,"PE Ratio":sorted_pe_ratio_list,"Profit growth":sorted_profit_growth_list}
            return {"data": return_list}
        else:
            return {"Error": "Invalid sector"}


api.add_resource(Sector_rankings, '/sector_rankings')

if __name__ == "__main__":
    app.run()
