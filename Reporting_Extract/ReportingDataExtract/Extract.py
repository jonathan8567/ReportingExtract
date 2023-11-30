from ReportingDataExtract import datetime, date, dt, timedelta, pd, relativedelta, load_workbook, Workbook, os, timeit

class GetReporting:
    def __init__(self, tgtticker, tgtday, function, file_path = "P:\\Product Specialists\\Tools\\Python Tools\\db_trial\\"):
        self.first_bd = str(tgtday) #first business day that is used to get webalto data
        self.ticker = tgtticker
        self.file_path = file_path + tgtticker + "_Slides_Data.xlsx"
        self.last_month = tgtday.replace(day=1) - datetime.timedelta(days=1)
        self.function = function        

        
        data = [self.first_bd, self.last_month]
        df = pd.DataFrame(data, columns=['Date'])
        
        # Check if there is already a excel file
        
        try: # try open the existing excel file
            wb = load_workbook(self.file_path)
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name = "Update_Date", header=True, index=False)
                
        except: # if not excel found, create a new one
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Update_Date"
            workbook.save(self.file_path)
            
            wb = load_workbook(self.file_path)
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name = "Update_Date", header=True, index=False)
        
        # find the file that specified urls 
        url_list = self.__ReadURL(self.ticker, self.last_month, self.first_bd)
        url_list_hist = self.__ReadURL_hist(self.ticker, self.last_month)
        url_list_hist_w = self.__ReadURL_hist_w(self.ticker, self.first_bd)
        
        # Creat an empty dataframe that is going to store all the data loaded
        df_main =  pd.DataFrame()            
        
        # Get end of month data and store in df_main
        print(self.ticker + " generating month end data...")
        for index in range(len(url_list)): # loop through each url in the url_list                   
            try:
                df_main[index] = self.__GetSlideData(url_list[index][1], url_list[index][0]) # Get slide data using the url
                pass
            except:
                pass
        
        # Get historical monthly data and store in df_main
        
        if self.function == "Generate": # if we want to generate historical data from the start
            print(self.ticker + " generating historic data...")
            
            for index in range(len(url_list_hist)):
                # generate historical data for fact sheet tables
                df_main[len(df_main)+index] = self.__GenerateHistData(url_list_hist[index][1], url_list_hist[index][0], self.ticker, self.file_path, self.last_month)
                
            for index in range(len(url_list_hist_w)):           
                # generate historical data for webalto tables                
                df_main[len(df_main)+index] = self.__GenerateHistData_W(url_list_hist_w[index][1], url_list_hist_w[index][0], self.ticker, self.file_path, self.last_month)

                  
        # Update historical monthly data and store in df_main
        
        elif self.function == "Update": # if we want to only update historical data
            print(self.ticker + " updating historic data...")
            
            for index in range(len(url_list_hist)):
                # generate historical data for fact sheet tables
                df_main[len(df_main)+index] = self.__UpdateHistData(url_list_hist[index][1], url_list_hist[index][0], self.ticker, self.file_path, self.last_month)

            for index in range(len(url_list_hist_w)): 
                # generate historical data for webalto tables
                df_main[len(df_main)+index] = self.__UpdateHistData_W(url_list_hist_w[index][1], url_list_hist_w[index][0], self.ticker, self.file_path, self.first_bd, self.last_month)

        
        
        # Save df_main (all the data loaded) in the excel file
        for index, df in enumerate(df_main):            
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                try:
                    df_main[index][1].to_excel(writer, sheet_name = df_main[index][0], header=True, index=False)
                except:
                    pass
            
            
    def __GetSlideData(self, url, sheetname): 
        # Get the table of month end data 
        try:        
            if "webalto" in url:
                # webalto table, try formatting the table by dropping unused rows & columns
                table = pd.read_html(url)
                df = table[0]
                df.columns = df.iloc[1, :]
                df = df.drop([0, 1])
                
                try:
                    df_new = df.astype('float64') # try to save the table in floating numbers
                except:
                    df_new = df        
                    
                return [sheetname, df_new]
            
            else:
                # fact sheet table, save directly            
                table = pd.read_html(url)
                df = table[0]                    
                return [sheetname, df]
        except:
            pass
        
        
    def __GetHistData(self, url, last_month):
        try:              
            table = pd.read_html(url)
            df1 = table[0]        
            df1 = df1.fillna(0)


            ##### Get column names and fill empty with i-1 #####
            ##### This prevents error cause by duplicate column names later #####
            column_name=[]
            for i in range(0,len(df1)):
                for j in range(1,len(df1.columns)):
                    if df1.iat[i,0] == 0:
                        df1.iat[i,0] = str(i-1)
                    column_name.append(df1.columns[j] + '_' + df1.iat[i,0])

            ##### Drop unused row and turn dataframe into a list #####                  
            df1.drop(df1.columns[0],axis = 1, inplace=True)
            df2 = df1.stack().tolist() 


            ##### Turn list to dataframe for later use (appending the dataframe)#####
            data = {last_month: df2}
            value_df = pd.DataFrame.from_dict(data, orient='index')
            value_df.columns = column_name
            return value_df
        
        except:
            pass        
        
        
    def __GenerateHistData(self, url, sheetname, ticker, file_path, last_month):
        # Define an empty list to store the data
        results = []
        
        last_month = last_month + relativedelta(months = 1)
        
        # loop through the past 50 months and get the data
        for i in range(50):
            last_month = last_month.replace(day=1) - datetime.timedelta(days=1)
            url_tgt = url.format(ticker, last_month)
            results.append(self.__GetHistData(url_tgt, last_month))

        df= pd.DataFrame()

        ##### Combine all output in 'results' into df #####
        for i in range(len(results)):
            try:
            ##### Handle duplicate values first #####
                results[i] = results[i].T[~results[i].T.index.duplicated(keep='last')].T
                if len(results[i].T) < 500:
                    df=df.append(results[i])
            except:
                pass
        df = df.fillna(0)

        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name = sheetname)
            
            
    def __UpdateHistData(self, url, sheetname, ticker, file_path, last_month):
        try:
            # Read the existing historic data if available
            url = url.format(ticker, last_month)       
            df= pd.read_excel(file_path,sheet_name = sheetname, index_col=0)
        except:
            df = pd.DataFrame()
               
        try:
            # Get a new line of the data and combine it with the old data
            results = self.__GetHistData(url, last_month)
            results = results.T[~results.T.index.duplicated(keep='last')].T
            df=pd.concat([results,df])
        except:
            pass       

        df = df.fillna(0)
        df = df[~df.index.duplicated(keep='first')]

        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:        
            df.to_excel(writer, sheet_name = sheetname)    
    

    def __GetHistData_W(self, url, last_month):
        try:              
            table_MN = pd.read_html(url)
            df1 = table_MN[0]        
            df1 = df1.fillna(0)            

            df1.drop(df1.columns[2:],axis = 1, inplace=True)
            df1.drop(df1.columns[:2],axis = 0, inplace=True)                                    
            
            ##### Get column names and fill empty with i-1 #####
            ##### This prevents error cause by duplicate column names later #####
            column_name=[]
            for i in range(0,len(df1)):
                if df1.iat[i,0] == 0:
                    df1.iat[i,0] = str(i)
                column_name.append(df1.iat[i,0])
            column_name = column_name[1:]

            ##### Drop unused row and turn values to a list #####                  
            df1.drop(df1.columns[0],axis = 1, inplace=True)
            df1.drop([2], inplace=True)
            df2 = df1.stack().tolist()
            
            ##### Turn list to dataframe #####
            data = {last_month: df2}
            value_df = pd.DataFrame.from_dict(data, orient='index')
            value_df.columns = column_name

            return value_df
        except:
            pass
    
    
    def __GenerateHistData_W(self, url, sheetname, ticker, file_path, last_month):
        results = []
        
        # define the list of business days
        start = date(date.today().year-5, date.today().month, 1)
        end = date.today()
        bussiness_days_rng =pd.date_range(start, end, freq='BMS')

        # loop through the business days in the past 5 years and get the data in webalto
        for bussiness_day in bussiness_days_rng:
            bussiness_day = bussiness_day.date()
            last_month = bussiness_day.replace(day=1) - datetime.timedelta(days=1)
            str_bussiness_day = str(bussiness_day).replace('-', '')
            url_tgt = url.format(str_bussiness_day, ticker)
            results.append(self.__GetHistData_W(url_tgt, last_month))

        df= pd.DataFrame()

        ##### Combine all output in 'results' into df #####
        for i in range(len(results)):
        ##### Handle duplicate values first #####
            try:
                df=df.append(results[i])
            except:
                pass

        df = df.fillna(0)

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name = sheetname)
        except:
            pass
        
        
    def __UpdateHistData_W(self, url, sheetname, ticker, file_path, first_bd, last_month):
        str_bussiness_day = str(first_bd).replace('-', '')
        url = url.format(str_bussiness_day, ticker)
        try:
            # Read the existing historic data if available
            df = pd.read_excel(file_path,sheet_name = sheetname, index_col=0)
        except:
            df = pd.DataFrame()
        
        try:
            results = __GetHistData_W(url, last_month)
            results = results.T[~results.T.index.duplicated(keep='last')].T
            df=df.append(results)
        except:
            pass

        df = df.fillna(0)
        df = df[~df.index.duplicated(keep='first')]

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:        
                df.to_excel(writer, sheet_name = sheetname)
        except:
            pass                
        
                
    def __ReadURL(self, ticker, last_month, first_bd):        
        #Read URL file
        first_bd = first_bd.replace("-","")
        file_path = "P:\\Product Specialists\\Tools\\Python Tools\\db_trial\\URL list\\" + ticker + "_URLs.xlsx"
        url_list_factsheet = pd.read_excel(file_path ,sheet_name = "Factsheet URLS")
        url_list_webalto = pd.read_excel(file_path ,sheet_name = "Webalto URLS")
        
        #Create a list to store the url
        url_list = []

        for i in range(len(url_list_factsheet)):
            url_list.append([url_list_factsheet.iloc[i, 0], url_list_factsheet.iloc[i, 1].format(ticker, last_month)])

        for i in range(len(url_list_webalto)):
            url_list.append([url_list_webalto.iloc[i, 0], url_list_webalto.iloc[i, 1].format(first_bd, ticker)])

        return url_list
                
        
    def __ReadURL_hist(self, ticker, last_month):        
        #Read URL file
        file_path = "P:\\Product Specialists\\Tools\\Python Tools\\db_trial\\URL list\\" + ticker + "_URLs.xlsx"
        url_list_hist = pd.read_excel(file_path ,sheet_name = "Hist URLS")
        
        #Create a list to store the url
        url_list = []

        for i in range(len(url_list_hist)):
            url_list.append([url_list_hist.iloc[i, 0], url_list_hist.iloc[i, 1]])

        return url_list
    
    
    def __ReadURL_hist_w(self, ticker, first_bd):        
        #Read URL file
        first_bd = first_bd.replace("-","")
        file_path = "P:\\Product Specialists\\Tools\\Python Tools\\db_trial\\URL list\\" + ticker + "_URLs.xlsx"
        url_list_hist_w = pd.read_excel(file_path ,sheet_name = "Hist URLS Webalto")

        #Create a list to store the url
        url_list = []
        
        for i in range(len(url_list_hist_w)):
            url_list.append([url_list_hist_w.iloc[i, 0], url_list_hist_w.iloc[i, 1]])

        return url_list