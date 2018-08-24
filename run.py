# -*- coding: utf-8 -*-
"""
Created on Thu Jul 26 14:41:02 2018

@author: Dev
"""
import pandas as pd
import calendar

#Paths for all excel files
momo_path =  "C:/Users/Dev/Documents/PEG AFRICA/Data Team/Data Sets/Mobile Money Data/MainDatasetJune.xlsx"
momo_analysis = "C:/Users/Dev/Documents/PEG AFRICA/Data Team/Data Sets/Mobile Money Data/FINAL ANALYSIS/MomoAnalysisJune1.xlsx"
staffdata_path = "C:/Users/Dev/Documents/PEG AFRICA/Data Team/Data Sets/Mobile Money Data/PEGSTAFF.xlsx"


# A class of momo excel sheets with functions to read each sheet
class Sheet:

    @staticmethod
    def read_MK170():
            return pd.read_excel(momo_path, sheet_name='MK170')

    @staticmethod
    def read_MK3456():
            return pd.read_excel(momo_path, sheet_name='MK3456')

    @staticmethod
    def read_MKPOP():
            return pd.read_excel(momo_path, sheet_name='MKPOP')

    @staticmethod
    def read_VodafoneMK():
            return pd.read_excel(momo_path, sheet_name='VodafoneMK')

    @staticmethod
    def read_TigoMK():
            return pd.read_excel(momo_path, sheet_name='TigoMK')

    @staticmethod
    def read_DL170():
            return pd.read_excel(momo_path, sheet_name='DL170')

    @staticmethod
    def read_DL3456():
            return pd.read_excel(momo_path, sheet_name='DL3456')

    @staticmethod
    def read_DLPOP():
            return pd.read_excel(momo_path, sheet_name='DLPOP')

    @staticmethod
    def read_TigoDL():
            return pd.read_excel(momo_path, sheet_name='TigoDL')

    @staticmethod
    def read_VodafoneDL():
            return pd.read_excel(momo_path, 'VodafoneDL')

    @staticmethod
    def read_staff():
        return pd.read_excel(staffdata_path, sheet_name='staff')
    



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
        month1_Payment_Count = sheet.Amount.count()
        return month1_Payment_Count

#Function to aggregate total amount received per Network Operator or Channel
def amounts_total(sheet):
        month1_amounts_total = sheet.agg({"Amount": "sum"})
        return month1_amounts_total[0]

def customer_class(sheet):
        customer_status = sheet.SinceActivation <=30
        customer_status = customer_status.value_counts()
        new_customers = customer_status[True]
        old_customers = customer_status[False]
        return(new_customers, old_customers)



def channel_aggregation(ussd,star, pop):

        #First Aggregation
        ussd = ussd.groupby(by=['AccountNumber','Channel'])['Amount'].agg(['count', 'sum']).rename(columns={'count': 'Frequency of Payment', 'sum': 'Sum Paid'})
        pop = pop.groupby(by=['AccountNumber','Channel'])['Amount'].agg(['count', 'sum']).rename(columns={'count': 'Frequency of Payment', 'sum': 'Sum Paid'})
        star = star.groupby(by=['AccountNumber','Channel'])['Amount'].agg(['count', 'sum']).rename(columns={'count': 'Frequency of Payment', 'sum': 'Sum Paid'})
        frame = [ussd, pop, star]
        frame = pd.concat(frame, sort=False)
        frame.reset_index(inplace=True)

        #Second Aggregation
        frame = frame.groupby(by=['AccountNumber','Channel']).agg({'Frequency of Payment':['sum'],'Sum Paid':['sum']})
        frame.columns = frame.columns.droplevel(1)
        return frame


def payment_by_peg_staff(staff, momo_all):
        staff = staff.drop_duplicates()
        staff = momo_all.Mobile.isin(staff)
        staff = staff.value_counts()
        staff = staff[True]
        return staff


def payment_by_customers(momo_all):
        customerPrimaryNumber = momo_all.PrimaryNumber.drop_duplicates()
        transactionMobile = momo_all.Mobile.drop_duplicates()
        ownWallet = transactionMobile.isin(customerPrimaryNumber)
        ownWallet = ownWallet.value_counts()
        ownWallet = ownWallet[True]
        return ownWallet



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
    staff = Sheet.read_staff()


    #Concatinating all files into one frame
    frames = [MK170,MK3456,MKPOP,VodafoneMK,TigoMK,DL170,DL3456,DLPOP,TigoDL,VodafoneDL]

    #Calling to convert frames date to datetime
    convert_to_datetime(frames)

    #All excel data mashedup as one
    momo_all = pd.concat(frames, sort=False)
    momo_all['Mobile'] = momo_all['Mobile'].fillna(0).astype('int64')
    momo_all['PrimaryNumber'] = momo_all['PrimaryNumber'].fillna(0).astype('int64')
    momo_all['SecondaryNumber1'] = momo_all['SecondaryNumber1'].fillna(0).astype('int64')
    momo_all['SecondaryNumber2'] = momo_all['SecondaryNumber2'].fillna(0).astype('int64')
    
    #print(momo_all)


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


    #Staff
    staff = staff.fillna(0).astype('int64')
    ro = staff.RO
    cx = staff.CX
    go = staff.GO
    sfm = staff.SFM
    asm = staff.ASM
    dsr = staff.DSR






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


    #Customer Phone numbers
    own_wallet = payment_by_customers(momo_all)

    #RO
    ro = payment_by_peg_staff(ro, momo_all)

    #DSR
    dsr = payment_by_peg_staff(dsr,momo_all)

    #GO
    go  = payment_by_peg_staff(go, momo_all)

    #ASM
    asm = payment_by_peg_staff(asm, momo_all)

    #CX
    cx = payment_by_peg_staff(cx, momo_all)

    #SFM
    sfm = payment_by_peg_staff(sfm, momo_all)


    #Sorting by Channel
    channel_analysis = channel_aggregation(pop,star,ussd)
    
    
    #Own wallet analysis - Primary Number
    other_wallet1 = momo_all[['AccountNumber','PrimaryNumber','Mobile','Amount']]
    other_wallet1 = other_wallet1.groupby(['AccountNumber','PrimaryNumber','Mobile'], as_index=False)['Amount'].agg(['count', 'sum'])

    #Own wallet analysis - Secondary Number 1
    other_wallet2 = momo_all[['AccountNumber','SecondaryNumber1','Mobile','Amount']]
    other_wallet2 = other_wallet2.groupby(['AccountNumber','SecondaryNumber1','Mobile'], as_index=False)['Amount'].agg(['count', 'sum'])

    #Own wallet analysis - Secondary Number 2
    other_wallet3 = momo_all[['AccountNumber','SecondaryNumber2','Mobile','Amount']]
    other_wallet3 = other_wallet3.groupby(['AccountNumber','SecondaryNumber2','Mobile'], as_index=False)['Amount'].agg(['count', 'sum'])



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
             'Existing USSD customers':[old_ussd_customers],
             ' ':[' '],
             'Number of customers who made payments from their own wallet':[own_wallet],
             '  ':[' '],
             'Number of payments done by ROs ': [ro],
             'Number of payments done by CXs ': [cx],
             'Number of payments done by GOs ': [go],
             'Number of payments done by SFMs ': [sfm],
             'Number of payments done by ASMs ': [asm]
             }
    

    #Saving data as dataframes to be read by excel
    kpi_df = pd.DataFrame(kpi);
    channel_analysis_df = pd.DataFrame(channel_analysis)
    other_wallet1_df = pd.DataFrame(other_wallet1)
    other_wallet2_df = pd.DataFrame(other_wallet2)
    other_wallet3_df = pd.DataFrame(other_wallet3)

    #Transposing dataframe
    kpi_df = kpi_df.T

    #Writing to Excel
    performance_writer = pd.ExcelWriter(momo_analysis, engine='xlsxwriter')
    kpi_df.to_excel(performance_writer, sheet_name='Results')
    channel_analysis_df.to_excel(performance_writer, sheet_name='Channel Analysis')
    other_wallet1_df.to_excel(performance_writer, sheet_name='OtherWallet1')
    other_wallet2_df.to_excel(performance_writer, sheet_name='OtherWallet2')
    other_wallet3_df.to_excel(performance_writer, sheet_name='OtherWallet3')


#Calling The main Function
if __name__ == "__main__":
    excute()

