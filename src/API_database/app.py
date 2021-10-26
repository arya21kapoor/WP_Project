from flask import Flask
from flask_restful import Resource, Api, reqparse
from skcriteria.madm import simple
from skcriteria import MAX, MIN, Data
import json
import os
import pandas as pd

app = Flask(__name__)
api = Api(app)


class Sector_and_Companies_List(Resource):
    def get(self):
        print('Running Sector_and_Companies_List class ...')
        return {'data': _get_sectors_and_companies(get_icons = True)}

api.add_resource(Sector_and_Companies_List, '/sectors_and_companies_list')

def _get_sectors_and_companies(get_icons = True):
    current_directory = os.getcwd() # get the current working directory
    new_directory = current_directory + '/csv_files' # appending the path
    # Here, all the csv files of the different companies of each sector is stored

    os.chdir(new_directory) # to go to the csv folder
    sector_list = os.listdir() # get the folder names (which are the Sector Names)

    icons_dict = {}

    try:
        sector_list.remove('icons.json')
        with open('icons.json', 'r', encoding='utf-8') as f:
            icons_dict = json.loads(f.read())
    except:
        print('Icons.json missing')

    sector_company_dict = {}

    if get_icons :
        for sector in sector_list:
            os.chdir(new_directory + '/' + sector) # go 1 level deeper
            sector_company_dict[sector] = {
            "icon" : icons_dict.get(sector, 'bi bi-app'), # retreiving the icon of each sector
            "companies" : [x[:x.index('.xlsx')] for x in os.listdir()] # retreiving the company names of each sector
            }
    else:
        for sector in sector_list:
            os.chdir(new_directory + '/' + sector) # go 1 level deeper
            sector_company_dict[sector] = [x[:x.index('.xlsx')] for x in os.listdir()] # retreiving the company names of each sector

    os.chdir(current_directory)
    return sector_company_dict


class CompanyDetails(Resource):
    def get(self, sector, company):
        print('Running Company Details Class ...')
        # Reading arguments from the url

        path = f'{os.getcwd()}\csv_files\\{sector}\\{company}.xlsx'

        if not os.path.isfile(path):
            return {'data':{'key': 'File Does not exist'}}

        try:
            profit_loss_df = pd.read_excel(path, sheet_name="Data Sheet", header=15, skipfooter=62)
            quarters_df = pd.read_excel(path, sheet_name="Data Sheet", header=40, skipfooter=43)
            balancesheet_df = pd.read_excel(path, sheet_name="Data Sheet", header=55, skipfooter=21)
            cashflow_df = pd.read_excel(path, sheet_name="Data Sheet", header=80, skipfooter=8)
        except:
            return {'data':{'key': 'Error in API (reading data)'}}

        #Above extracts the specfic tables into dataframes

        def convertToJSON(df):
            df = df.dropna(axis=1, how='all')
            df.rename(columns={x:x.strftime("%b-%y") for x in df.columns.values[1:]}, inplace=True)
            #converting the column names into a good datetime format
            no_of_columns = len(df.columns.values)
            no_of_rows = len(df)

            keys = df.columns.values[1:]


            t = {} #TODO add comment proper

            for i in range(0, no_of_rows):
                row_element = df['Report Date'][i] # attribute name (unique for each column)
                temp_dict = {}
                for j in range(0, no_of_columns-1):
                    temp_dict[keys[j]] = df.iloc[i,j+1] # column name : cell value as key-value pair
                t[row_element] = temp_dict # attribute name : row dict as key_value pair

            d = {'headers': ['Report Date'] + list(keys), 'values' : t}

            return d


        myList = ['Profit & Loss', 'Quarters', 'Balance Sheet', 'Cash Flow']
        df_list = [profit_loss_df, quarters_df, balancesheet_df, cashflow_df]
        q = {x : convertToJSON(y) for x, y in zip(myList, df_list)}

        return {'data' : q}

api.add_resource(CompanyDetails, '/company_details/sector=<sector>/company=<company>')


class SectorDetails(Resource):
    def get(self, sector):
        # add self afterwards
        print('Running Sector Class ...')
        current_directory = os.getcwd() # get the current working directory

        path = f'{current_directory}\csv_files\\{sector}'

        if not os.path.isdir(path):
            return {'data' : {'key' : 'Folder does not exist'}}

        company_list = _get_sectors_and_companies(get_icons = False)[sector]

        new_directory = current_directory + '/csv_files' + f'/{sector}'

        myList = []
        for comp in company_list:
            path = new_directory + f'/{comp}.xlsx'
            company_df = pd.read_excel(path, sheet_name="Data Sheet")

            market_cap_value = company_df.iloc[7][1]

            current_annual_earnings = company_df.iloc[28][10]
            price_to_earnings = round(market_cap_value/current_annual_earnings, 6)

            last_annual_earnings = company_df.iloc[28][9]
            profit_growth = (current_annual_earnings - last_annual_earnings)/last_annual_earnings
            profit_growth = round(profit_growth, 6)

            myList.append({
                'Company' : comp,
                'Market Cap.' : market_cap_value,
                'P/E Ratio' : price_to_earnings,
                'Profit Growth (%)' : profit_growth
            })

        headers = list(myList[0].keys())

        os.chdir(current_directory)

        return {'headers' : headers, 'data' : myList}

api.add_resource(SectorDetails, '/sector=<sector>')
class MCDA_rankings(Resource):

    def get(self):
        print("MCDA_ranking class running..")

        parser = reqparse.RequestParser()
        parser.add_argument('Rank_type', required=True)

        args = parser.parse_args()

        rank_type = args['Rank_type']
        print("Rank_type:", rank_type)
        company_names = []
        return_list = []
        name=""
        #Getting the orignal path
        orignal_path=os.getcwd()
        #Adding the '\csv files\' to go into this directory
        csv_path=orignal_path + '\\csv_files'
        sector_list = list(os.listdir(csv_path))
        # Removing financial sector and icons.json
        sector_list.remove('Finance')
        sector_list.remove('icons.json')
        # This segment takes care of making a dataframe for MCDA data object
        for i in sector_list:
            sector_path = csv_path+'\\'+i
            csv_list = list(os.listdir(sector_path))
            for j in csv_list:
                company_dict = {}
                company_name = j.replace('.xlsx', '')
                company_path = sector_path+'\\'+j
                df = pd.read_excel(company_path, sheet_name='Data Sheet')
                # Type of rank
                if rank_type.lower() == 'stat_cheap':
                    name="Value"
                    # Fetching the data values
                    b9 = df.loc[7].iat[1]
                    k30 = df.loc[28].iat[10]
                    # Applying the formula
                    pe = b9/k30
                    k57 = df.loc[55].iat[10]
                    k58 = df.loc[56].iat[10]
                    p_bv = b9/(k57+k58)
                    k17 = df.loc[15].iat[10]
                    p_s = b9/k17
                    k28 = df.loc[26].iat[10]
                    k27 = df.loc[25].iat[10]
                    k26 = df.loc[24].iat[10]
                    icr = (k28+k27+k26)/k26
                    # Assigning weights
                    w = [0.4, 0.4, 0.2]
                    # Assigning column names
                    c = ['PE ratio', 'Price to Book value', 'Price to sales']
                    max_min_criteria = [MIN, MIN, MIN]
                    # Necessary condition
                    if pe < 30 and pe >= 0:
                        # This dictionary contains the attribute
                        attribute_dict = {}
                        company_names.append(company_name)
                        attribute_dict['Company name'] = company_name
                        attribute_dict['PE Ratio'] = '%.3f'%(pe)
                        attribute_dict['Price to Book Value'] = '%.3f'%(p_bv)
                        attribute_dict['Price to Sales'] = '%.3f'%(p_s)
                        # This is the return list in which we append the dictionary
                        return_list.append(attribute_dict)

                elif rank_type.lower() == 'high_growth':

                    name="Growth"
                    k17 = df.loc[15].iat[10]
                    f17 = df.loc[15].iat[5]
                    sg_five = ((k17-f17)/f17)*100
                    h17 = df.loc[15].iat[7]
                    sg_three = ((k17-h17)/f17)*100
                    k30 = df.loc[28].iat[10]
                    f30 = df.loc[28].iat[5]
                    pg_five = ((k30-f30)/f30)*100
                    h30 = df.loc[28].iat[5]
                    pg_three = ((k30-h30)/h30)*100
                    k28 = df.loc[26].iat[10]
                    k27 = df.loc[25].iat[10]
                    k26 = df.loc[24].iat[10]
                    icr = (k28+k27+k26)/k26
                    max_min_criteria = [MAX, MAX, MAX, MAX]
                    w = [0.3, 0.2, 0.3, 0.2]
                    c = ['Sales growth after 5years', 'Sales growth after 3 years',
                         'Profit growth after 5 years', 'Profit growht after 3 years']
                    if icr > 4:
                        attribute_dict = {}
                        company_names.append(company_name)
                        attribute_dict['Company name'] = company_name
                        attribute_dict['Sales growth after 5years (%)'] = '%.3f'%(sg_five)
                        attribute_dict['Sales growth after 3years (%)'] = '%.3f'%(sg_three)
                        attribute_dict['Profit growth after 5years (%)'] = '%.3f'%(pg_five)
                        attribute_dict['Profit growth after 3years (%)'] = '%.3f'%(pg_three)
                        return_list.append(attribute_dict)

                elif rank_type.lower() == "debt_reduction":
                    name="Debt Reduction"
                    k57 = df.loc[55].iat[10]
                    k58 = df.loc[56].iat[10]
                    k59 = df.loc[57].iat[10]
                    k60 = df.loc[58].iat[10]
                    de = ((k59+k60)/(k57+k58))
                    k28 = df.loc[26].iat[10]
                    k27 = df.loc[25].iat[10]
                    k26 = df.loc[24].iat[10]
                    icr = (k28+k27+k26)/k26
                    h59 = df.loc[57].iat[7]
                    h60 = df.loc[58].iat[7]
                    dr_three = (k59 + k60) - (h59 + h60)
                    f59 = df.loc[57].iat[5]
                    f60 = df.loc[58].iat[5]
                    f57 = df.loc[55].iat[5]
                    f58 = df.loc[56].iat[5]
                    dr_five = (k59 + k60) - (f59 + f60)
                    debt_to_er_five = ((f59 + f60)/(f57 + f58)) - \
                        ((k59 + k60)/(k57 + k58))
                    max_min_criteria = [MAX, MAX, MAX]
                    w = [0.3, 0.4, 0.3]
                    c = ['Debt Reduction 3 years', 'Debt Reduction 5 years',
                         'Debt to Equity Reduction 5 years']
                    if icr > 4 and 0 <= de < 1.1 and company_name != 'Infosys':
                        attribute_dict = {}
                        company_names.append(company_name)
                        attribute_dict['Company name'] = company_name
                        attribute_dict['Debt Reduction 3 years (Cr)'] = '%.3f'%(dr_three)
                        attribute_dict['Debt Reduction 5 years (Cr)'] = '%.3f'%(dr_five)
                        attribute_dict['Debt to Equity Reduction 5 years (%)'] = '%.3f'%(debt_to_er_five)
                        return_list.append(attribute_dict)

                elif rank_type.lower() == "magic_formula":
                    name="Magic Formula"
                    k30 = df.loc[28].iat[10]
                    j30 = df.loc[28].iat[9]
                    i30 = df.loc[28].iat[8]
                    k58 = df.loc[56].iat[10]
                    j58 = df.loc[56].iat[9]
                    i58 = df.loc[56].iat[8]
                    roe_three = ((k30 + j30 + i30)/(k58 + j58 + i58)) * 100
                    k28 = df.loc[26].iat[10]
                    k27 = df.loc[25].iat[10]
                    k26 = df.loc[24].iat[10]
                    icr = (k28+k27+k26)/k26
                    b9 = df.loc[7].iat[1]
                    k30 = df.loc[28].iat[10]
                    pe = b9/k30
                    earning_yield = (1/pe)
                    w = [0.5, 0.5]
                    max_min_criteria = [MAX, MAX]
                    c = ['Return on Equity 3 years', 'Earning Yield']
                    if icr > 4 and company_name not in ['B P C L','Coal India','I O C L','NTPC','O N G C','Power Grid Corpn']:
                        attribute_dict = {}
                        company_names.append(company_name)
                        attribute_dict['Company name'] = company_name
                        attribute_dict['Return on Equity 3 years (%)'] = '%.3f'%(roe_three)
                        attribute_dict['Earning Yield'] = '%.3f'%(earning_yield)
                        return_list.append(attribute_dict)

                else:
                    return {"Error": "Invalid rank type"}
        mcda_df = pd.DataFrame(return_list)
        mcda_df.drop('Company name', axis='columns', inplace=True)
        # Loading the weights,mcda_df,anames and cnames on the Data Object
        data = Data(mcda_df,
                    max_min_criteria,
                    weights=w,
                    anames=company_names,
                    cnames=c)
        # Taking a copy of mcda_df so that we can append the results
        score_df = mcda_df.copy()

        # Using weighted sum for ranking
        # Using sum normalisation
        dm = simple.WeightedSum(mnorm="sum")
        dec = dm.decide(data)
        #score_df.loc[:, 'Algorithm points'] = map(lambda x:'%.3f'%(x),list(dec.e_.points))
        score_df.loc[:, 'Algorithm points'] = ['%.3f'%(x) for x in list(dec.e_.points)]
        score_df.loc[:, 'Rank'] = dec.rank_

        pd.set_option('display.max_rows', None, 'display.max_columns', None)
        score_df.insert(0, "Name", company_names)
        score_df = score_df.sort_values(by="Rank")
        score_df = score_df.head(10)
        #print("Score df:\n", score_df)

        # Making the return_dict for the front end
        return_list = []
        return_dict = score_df.to_dict("records")
        headers = list(return_dict[0].keys())


        return {"name":name,"headers":headers,"data": return_dict}


api.add_resource(MCDA_rankings, '/mcda_rankings')


if __name__ == "__main__":
    app.run()
