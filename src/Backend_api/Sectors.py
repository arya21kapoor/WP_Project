from flask import Flask
from flask_restful import Resource,Api,reqparse
import pandas as pd 
import ast
import os

#Initialising the flask api
app= Flask(__name__)
api=Api(app)

#Created Sector_company as an endpoint
class Sector_company(Resource):
    def get(self):

        print("Sector_company class running...")
        #Initializing the parser
        parser=reqparse.RequestParser()
        #Adding parser arguments
        parser.add_argument('sector_name',required=False)
        #parser.add_argument('company_name',required=True)
        #Creating argument dictionary
        args=parser.parse_args()
        #sector_list=["Automobile","Constructon","Finance","FMCG","HealthCare","Metals_Chemicals","Power","Technology"]
        sector_file_list=list(os.listdir('D:\WP_project\src\csv_files'))
        print(sector_file_list)
        sector=args['sector_name']
        #Created list of csv_files under each sector
        csv_list=[]
        if(sector=="Automobile"):
            csv_list=list(os.listdir('D:\WP_project\src\csv_files\Automobile'))
            print("List of excel files:",csv_list)
        elif(sector=="Construction"):
            csv_list=list(os.listdir('D:\WP_project\src\csv_files\Construction'))
            print("List of excel files:",csv_list)
        elif(sector=="Finance"):
            csv_list=list(os.listdir('D:\WP_project\src\csv_files\Finance'))
            print("List of csv files:",csv_list)
        elif(sector=="FMCG"):
            csv_list=list(os.listdir('D:\WP_project\src\csv_files\FMCG'))
            print("List of csv files:",csv_list)
        elif(sector=="HealthCare"):
            csv_list=list(os.listdir('D:\WP_project\src\csv_files\HealthCare'))
            print("List of csv files:",csv_list)
        elif(sector=="Metals_Chemicals"):
            csv_list=list(os.listdir('D:\WP_project\src\csv_files\Metals_Chemicals'))
            print("List of csv files:",csv_list)
        elif(sector=="Power"):
            csv_list=list(os.listdir('D:\WP_project\src\csv_files\Power'))
            print("List of csv files:",csv_list)
        elif(sector=="Technology"):
            csv_list=list(os.listdir('D:\WP_project\src\csv_files\Technology'))
            print("List of csv files:",csv_list)
        else:
            csv_list=[]

        #Returning the csv_files_dictionary
        return {'data':csv_list}
#Added this class for the endpoint /sectors
api.add_resource(Sector_company,'/sectors')

if __name__ =="__main__":
    app.run()
