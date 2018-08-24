import pandas as pd

#Paths for all excel files
momo_path = "C:/Users/Peg/Desktop/mkopa-dlight may june 2019 (1).xlsx"

# A class of momo excel sheets with functions to read each sheet
class Sheet:
    
    @staticmethod 
    def read_MK170():
            return pd.read_excel(momo_path, sheet_name='MK 170')
        
    @staticmethod
    def read_MK3456():
            return pd.read_excel(momo_path, sheet_name='MK 3456')
        
    @staticmethod
    def read_MKPOP():
            return pd.read_excel(momo_path, sheet_name='MK POP')
        
    @staticmethod
    def read_VodafoneMK():
            return pd.read_excel(momo_path, sheet_name='Vodafone MK')
        
    @staticmethod
    def read_TigoMK():
            return pd.read_excel(momo_path, sheet_name='Tigo MK')
        
    @staticmethod
    def read_DL170():
            return pd.read_excel(momo_path, sheet_name='DL 170')
        
    @staticmethod
    def read_DL3456():
            return pd.read_excel(momo_path, sheet_name='DL 3456')
        
    @staticmethod
    def read_DLPOP():
            return pd.read_excel(momo_path, sheet_name='DL POP')
        
    @staticmethod
    def read_TigoDL():
            return pd.read_excel(momo_path, sheet_name='Tigo DL')
        
    @staticmethod
    def read_VodafoneDL():
            return pd.read_excel(momo_path, 'Vodafone DL')


   
## Function to filter for the New Channels in the momo data
#def filter_NewChannel(sheet):
#        return sheet.loc[sheet["Channel"] == "NEW"]
#    
#
## Function to filter for the Old Channels in the momo data
#def filter_OldChannel(sheet):
#        return sheet.loc[sheet["Channel"] == "OLD"]
#    
#def distinct_Account(sheet):
#       return sheet.AccountNumber.nunique()
#   
#def distinct_PerMonth(sheet):
#       return sheet.groupby(by=['Month']).AccountNumber.nunique()

#   
##Function to return Number of New Channel Users per Month
#def channel_Users(sheet):
#        return distinct_PerMonth(sheet)

def convert_to_datetime(frames):
    #Formatting Date for all
    for frame in frames:
        frame['TransactionDate'] = pd.to_datetime(frame['TransactionDate'])
        frame['ActivationDate'] = pd.to_datetime(frame['ActivationDate'])
        frame['LastDayOfMonth'] = pd.to_datetime(frame['LastDayOfMonth'])
        frame['Month'] = frame['TransactionDate'].dt.month
        frame["SinceActivation"] = (frame['LastDayOfMonth']-frame["ActivationDate"]).dt.days
    return frames
    
#Function to count the number of payment transactions by Network Operator
def payments_collected_count(sheet):
        month1 = sheet[sheet.Month == 5]
        month1_Payment_Count = month1.Amount.count()
        return month1_Payment_Count

#Function to aggregate total amount received per Network Operator or Channel     
def amounts_total(sheet):
        month1 = sheet[sheet.Month == 5]
        month1_amounts_total = month1.agg({"Amount": "sum"})
        return month1_amounts_total[0]

def customer_class(sheet):
        month1 = sheet[sheet.Month == 5]
        customer_status = month1.SinceActivation <=30
        customer_status = customer_status.value_counts()
        new_customers = customer_status[True]
        old_customers = customer_status[False]
        return(new_customers, old_customers)
 
def channel_aggregation(ussd,star, pop):
        ussd = ussd.groupby(by=['AccountNumber','Channel'])['Amount'].agg(['count', 'sum']).reset_index().rename(columns={'count':'Frequency of payment', 'sum':'Sum Paid','Channel':'Payment Method (POP/STAR/USSD)'})
        pop = pop.groupby(by=['AccountNumber','Channel'])['Amount'].agg(['count', 'sum']).reset_index().rename(columns={'count':'Frequency of payment', 'sum':'Sum Paid','Channel':'Payment Method (POP/STAR/USSD)'})
        star = star.groupby(by=['AccountNumber','Channel'])['Amount'].agg(['count', 'sum']).reset_index().rename(columns={'count':'Frequency of payment', 'sum':'Sum Paid','Channel':'Payment Method (POP/STAR/USSD)'})
        frame = [ussd, pop, star]
        frame = pd.concat(frame, sort=False)
        return frame      


def excute():
    
    #Saving Read sheets as variables
    MK170 = Sheet.read_MK170()
    MK3456 = Sheet.read_MK3456()
    MKPOP = Sheet.read_MKPOP()
    VodafoneMK = Sheet.read_VodafoneMK()
    TigoMK = Sheet.read_TigoMK()
    DL170 = Sheet.read_DL170()
    DL3456 = Sheet.read_DL3456()
    DLPOP = Sheet.read_DLPOP()
    TigoDL = Sheet.read_TigoDL()
    VodafoneDL = Sheet.read_VodafoneDL()
   
    #Concatinating all files into one frame
    frames = [MK170,MK3456,MKPOP,VodafoneMK,TigoMK,DL170,DL3456,DLPOP,TigoDL,VodafoneDL]
    
    #Calling to convert frames date to datetime
    convert_to_datetime(frames)
 
    #A dataframe of New Channel Users
    new_channel = [MK3456,DL3456,MKPOP,DLPOP]
    new_channel = pd.concat(new_channel, sort=False)
    
    #A dataframe of Old Channel Users
    ussd = [MK170,DL170,TigoDL,TigoMK,VodafoneDL,VodafoneMK]
    ussd = pd.concat(ussd, sort=False)
    
    #Star Channel
    star = [MK3456,DL3456]
    star = pd.concat(star, sort=False)
    
    #POP Channel
    pop = [MKPOP,DLPOP]
    pop = pd.concat(pop, sort=False)
    
    #MTN 
    mtn = [MK170, MK3456, MKPOP, DL170, DL3456, DLPOP]
    mtn = pd.concat(mtn, sort=False)
    
    #VODAFONE
    vodafone = [VodafoneMK, VodafoneDL]
    vodafone = pd.concat(vodafone, sort=False)
    
    #TIGO
    tigo = [TigoDL, TigoMK]
    tigo = pd.concat(tigo, sort=False)
    
    
    
    #Payment Collected per MTN
    mtn_payaments_Collected = payments_collected_count(mtn)
    print("This is the payment collected by MTN: {0}".format(mtn_payaments_Collected))
    
    
    #Payment Collected per VODAFONE
    vodafone_payaments_Collected = payments_collected_count(vodafone)
    print("This is the payment collected by Vodafone: {0}".format(vodafone_payaments_Collected))

     #Payment Collected per TIGO
    tigo_payaments_Collected = payments_collected_count(tigo)
    print("This is the payment collected by TIGO: {0}".format(tigo_payaments_Collected))
    
    
    #Amounts Collected by MTN
    mtn_amounts_total = amounts_total(mtn)
    print("This is total Amount Received from MTN: {:f}".format(mtn_amounts_total))
    
    #Amounts Collected by VODAFONE
    vodafone_amounts_total = amounts_total(vodafone)
    print("This is total Amount Received from VODAFONE: {:f}".format(vodafone_amounts_total))
    
    #Amounts Collected by TIGO
    tigo_amounts_total = amounts_total(tigo)
    print("This is total Amount Received from TIGO: {:f}".format(tigo_amounts_total))
    
    #New and Existing Customers in POP
    pop_customers = customer_class(pop)
    print("New POP customers: {0}\nExisting POP customers: {1}".format(pop_customers[0], pop_customers[1]))
    
    new_pop_customers, old_pop_customers = pop_customers
   
    #New and Existing Customers in STAR
    star_customers = customer_class(star)
    print("New STAR customers: {0}\nExisting STAR customers: {1}".format(star_customers[0], star_customers[1])) 
    new_star_customers, old_star_customers = star_customers
    
    
    #New and Existing Customers in USSD Channels
    ussd_customers = customer_class(ussd)
    print("New USSD customers: {0}\nExisting USSD customers: {1}".format(ussd_customers[0], ussd_customers[1]))
    new_ussd_customers, old_ussd_customers = ussd_customers
    
    channel_analysis = channel_aggregation(pop,star,ussd)
    
    
      #Saving return values from functions as a dictionary with the KPIs as the keys
    kpi =   {'KPI':['Month1'],
             'Payment Collected from MTN':[mtn_payaments_Collected],
             'Payment Collected from Vodafone':[vodafone_payaments_Collected],
             'Payment Collected from Tigo':[tigo_payaments_Collected],
             'Total Amount Received from MTN':[mtn_amounts_total],
             'Total Amount Received from Vodafone':[vodafone_amounts_total],
             'Total Amount Received from Tigo': [tigo_amounts_total],
             'New POP customers': [new_pop_customers],
             'Existing POP customers': [old_pop_customers],
             'New Star Customers': [new_star_customers],
             'Existing STAR customers': [old_star_customers],
             'New USSD Customers': [new_ussd_customers],
             'Existing USSD customers':[old_ussd_customers]
             }
    
        
    df = pd.DataFrame(kpi, index=None)
    
    channel_analysis_df = pd.DataFrame(channel_analysis)
    channel_analysis_df.set_index('AccountNumber', inplace=True)
    
    df = df.T

    performance_writer = pd.ExcelWriter('MayJuneMobileMoneyAnalysis.xlsx', engine='xlsxwriter')
    

    df.to_excel(performance_writer, sheet_name='Results')
    
    channel_analysis_df.to_excel(performance_writer, sheet_name='Channel Analysis')
    
    
#Calling The main Function 
if __name__ == "__main__":
    excute()
