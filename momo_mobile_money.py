
import pandas as pd

#Paths for all excel files
momo_path =  "C:/Users/Peg/Desktop/dataset.xlsx"


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


   
# Function to filter for the New Channels in the momo data
def filter_NewChannel(sheet):
        return sheet.loc[sheet["Channel"] == "NEW"]
    

# Function to filter for the Old Channels in the momo data
def filter_OldChannel(sheet):
        return sheet.loc[sheet["Channel"] == "OLD"]
    
def distinct_Account(sheet):
       return sheet.AccountNumber.nunique()
   
def distinct_PerMonth(sheet):
       return sheet.groupby(by=['Month']).AccountNumber.nunique()

   
#Function to return Number of New Channel Users per Month
def channel_Users(sheet):
        return distinct_PerMonth(sheet)

def convert_to_datetime(frames):
    #Formatting Date for all
    for frame in frames:
        frame['TransactionDate'] = pd.to_datetime(frame['TransactionDate'])
        frame['ActivationDate'] = pd.to_datetime(frame['ActivationDate'])
        frame['Month'] = frame['TransactionDate'].dt.month
    return frames
    
#Function to Calculate Percentage Growth
def percentage_growth(sheet):
    existingCustomers = distinct_PerMonth(sheet)
    oldCustomers = existingCustomers[5]
    newCustomers = existingCustomers[6]
    diff =  newCustomers - oldCustomers
    percentage_growth = float(diff/oldCustomers)*100
    print("This is percentage growth: %.2f%%" % percentage_growth)
    return [oldCustomers, newCustomers, percentage_growth]

#Function to Calculate Percentage Growth
def percentage_growth1(sheet):
        month1 = sheet['AccountNumber'][sheet["Month"]==5]
        month2 = sheet['AccountNumber'][sheet["Month"]==6]   
        month1 = month1.drop_duplicates()
        month2 = month2.drop_duplicates()    
        check1 = month2.isin(month1)     
        status = check1.value_counts()    
        accounts_in_month2 = status[True]
        accounts_not_in_month1 = status[False]     
        percentage_growth = float(accounts_not_in_month1/accounts_in_month2)*100
        print("This is percentage growth: %.2f%%" % percentage_growth)
        return (accounts_in_month2, accounts_not_in_month1, percentage_growth)

#Function to find percentage of Customers paying with a particular Channel
def customers_Paying_Channel(sheet):
    existingCustomers = distinct_PerMonth(sheet)
    oldCustomers = existingCustomers[5]
    newCustomers = existingCustomers[6]
    totalCustomers = oldCustomers + newCustomers
    percentageExistingCustomers = float(oldCustomers/totalCustomers)*100
    print("This is the percentage of existing customers paying with a particular channel: %.2f%%" % percentageExistingCustomers)
    percentageExistingCustomers = [oldCustomers, newCustomers, percentageExistingCustomers]
    percentageNewCustomers = float(newCustomers/totalCustomers)*100
    print("This is the percentage of new customers paying with a particular channel: %.2f%%" % percentageNewCustomers)
    percentageNewCustomers = [oldCustomers, newCustomers, percentageNewCustomers]
    return (percentageExistingCustomers, percentageNewCustomers)

#Function to find the drop in stickiness to channnels
def stickiness_measure(sheet): #Stickiness is always going to be negative because we are analysing two months of existing customers
        month1 = sheet['AccountNumber'][sheet["Month"]==5]
        month2 = sheet['AccountNumber'][sheet["Month"]==6]
        month1 = month1.drop_duplicates()
        month2 = month2.drop_duplicates()
        check1 = month1.isin(month2)
        status = check1.value_counts()
        accounts_in_month1 = status[True]
        accounts_not_in_month2 = status[False]
        accounts_check = accounts_not_in_month2 + accounts_in_month1
        stickness_percentage = float(accounts_not_in_month2/accounts_in_month1)*100
        print("This is the percentage drop in Stickiness: %.2f%%" % stickness_percentage)
        return [accounts_in_month1, accounts_check, stickness_percentage]


#Function to Calculate the %Change in Value payment
def value_of_payment(sheet):
        month1 = sheet[sheet.Month == 5]
        month2 = sheet[sheet.Month == 6]
        month1_Amount = month1.Amount.sum()
        month2_Amount = month2.Amount.sum()
        diff = month2_Amount - month1_Amount
        value_diff = float(diff/month1_Amount)*100
        print("This is percentage change in value payment: %.2f%%" % value_diff)
        return [month1_Amount, month2_Amount, value_diff]

 #Function to Calculate the %Change in Frequency of Payment      
def frequency_of_payment(sheet):
        month1 = sheet[sheet.Month == 5]
        month2 = sheet[sheet.Month == 6]
        month1['isin_status'] = month1.AccountNumber.isin(month2.AccountNumber)
        month1_trans_count1 = month1['AccountNumber'][month1["isin_status"]==True]
        month1_trans_count = len(month1_trans_count1)  #We used len because count() was omitting some of the accountnumbers
        month2['isin_status'] = month2.AccountNumber.isin(month1.AccountNumber)
        month2_trans_count1 = month2['AccountNumber'][month2["isin_status"]==True]
        month2_trans_count = len(month2_trans_count1)
        diff = month2_trans_count - month1_trans_count
        freq_payment_percentage = float(diff/month1_trans_count)*100
        print("This is percentage change in frequency payment: %.2f%%" % freq_payment_percentage)
        return [month1_trans_count, month2_trans_count, freq_payment_percentage]
        
 
# Determing new customers
def thirtyDayMargin(sheet):
        return (sheet["SinceActivation"] <= 30)


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
    
    #All excel data mashedup as one
    #momo_all = pd.concat(frames, sort=False)
    #print(momo_all)
   
 
    #A dataframe of New Channel Users
    new_channel = [MK3456,DL3456,MKPOP,DLPOP]
    new_channel = pd.concat(new_channel, sort=False)
    
    #A dataframe of Old Channel Users
    old_channel = [MK170,DL170,TigoDL,TigoMK,VodafoneDL,VodafoneMK]
    old_channel = pd.concat(old_channel, sort=False)
    
    #Star Channel
    star = [MK3456,DL3456]
    star = pd.concat(star, sort=False)
    
    #POP Channel
    pop = [MKPOP,DLPOP]
    pop = pd.concat(pop, sort=False)
    
    
    #Print New Channel Users
#    print("Number of Unique users of new channel:")
#    print(channel_Users(new_channel))
    
    
    #Print Old Channel Users
#    print("Number of Unique users of old channel: ")
#    print(channel_Users(old_channel))
    
    
    #Calling Percentage Growth Function and passing it data
    print("\n### Percentage Growth New channel - method 1 ####")
    number_of_old_customers, number_of_new_customers, percentage_change_one = percentage_growth(new_channel)
    
    
     #Calling Percentage Growth Function and passing it data
    print("\n### Percentage Growth New Channel method 2 ####")
    accounts_second_month, accounts_not_first_month, percentage_change_two = percentage_growth1(new_channel)

    
    #Calling Percentage Growth Function and passing it data
    print("\n### Percentage Growth old Channel method 1 ####")
    number2_of_old_customers, number2_of_new_customers, percentage_change_three = percentage_growth(old_channel)
    
    
    #Calling Percetage Growth Function and passing it data
    print("\n### Percentage Growth old Channel method 2 ####")
    accounts1_second_month, accounts1_not_first_month, percentage_change_four = percentage_growth1(old_channel)
    
    
    #Calling Existing Customers Paying with a particular Channel and passing that Channel data
    print("\n### Existing Customers Paying with NEW Channel ####")
    existingCustomersNewChannel, newCustomersNewChannel = customers_Paying_Channel(new_channel)
    fresh_channel_customers, already_channel_customers, percentage_change_five = existingCustomersNewChannel
    fresh1_channel_customers, already1_channel_customers, percentage_change_six = newCustomersNewChannel
    
    
    #Calling Existing Customers Paying with a particular Channel and passing that Channel data
    print("\n### Existing Customers Paying with STAR Channel ####")
    existingCustomerSTAR, newCustomerSTAR = customers_Paying_Channel(star) 
    old_star_customers, new_star_customers, percentage_change_seven = existingCustomerSTAR
    old1_star_customers, new1_star_customers, percentage_change_eight = newCustomerSTAR
    
    #Calling Existing Customers Paying with a particular Channel and passing that Channel data
    print("\n### Existing Customers Paying with POP Channel ####")
    existingCustomersPOP, newCustomersPOP = customers_Paying_Channel(pop)
    old_pop_customers, new_pop_customers, percentage_change_nine = existingCustomersPOP
    old1_pop_customers, new1_pop_customers, percentage_change_ten = newCustomersPOP
    
    
    #Calling Stickiness measure to check the stickness in use of channel
    print("\n### Percentage drop in stickness of new solution ####")
    first_month_new_channel, new_regular_accounts, new_percent_stickiness = stickiness_measure(new_channel)
    
    
    #Calling Stickiness measure to check the stickness in use of channel
    print("\n### Percentage drop in stickness of old Channel ####")
    first_month_old_channel, old_regular_accounts, old_percent_stickiness = stickiness_measure(old_channel)
    
        
    #Calling value of payment function to calculate the %Change in value of payment    
    print("\n### Percentage in value of payment for STAR ####")
    star_first_amount, star_second_amount, star_difference_value = value_of_payment(star)
    
    #Calling value of payment function to calculate the %Change in value of payment    
    print("\n### Percentage in value of payment for POP ####")
    pop_first_amount, pop_second_amount, pop_difference_value = value_of_payment(pop)

    #Calling frequency of payment function to calculate the %Change in frequency of payment    
    print("\n### Percentage Change in frequency of payment for POP ####")   
    pop_month1_transaction, pop_month2_transaction, pop_frequency = frequency_of_payment(pop)
    
    #Calling frequency of payment function to calculate the %Change in frequency of payment    
    print("\n### Percentage Change in frequency of payment for STAR ####")   
    star_month1_transaction, star_month2_transaction, star_frequency = frequency_of_payment(star)
    
    perform_indicators = pd.Series([percentage_change_one, percentage_change_two, percentage_change_three, percentage_change_four,
                                    percentage_change_five, percentage_change_six, percentage_change_seven, percentage_change_eight,
                                    percentage_change_nine, percentage_change_ten,new_percent_stickiness, old_percent_stickiness,
                                    star_difference_value, pop_difference_value,pop_frequency, star_frequency])
    
    initial_month = pd.Series([number_of_old_customers, accounts_second_month, number2_of_old_customers, accounts1_second_month,
                               fresh_channel_customers, fresh1_channel_customers, old_star_customers, old1_star_customers,
                               old_pop_customers, old1_pop_customers, first_month_new_channel, first_month_old_channel,
                               star_first_amount, pop_first_amount, pop_month1_transaction, star_month1_transaction])
    
    current_month = pd.Series([number_of_new_customers, accounts_not_first_month, number2_of_new_customers, accounts1_not_first_month,
                               already_channel_customers, already1_channel_customers, new_star_customers, new1_star_customers,
                               new_pop_customers, new1_pop_customers,new_regular_accounts, old_regular_accounts,
                               star_second_amount, pop_second_amount, pop_month2_transaction, star_month2_transaction])

    
    key_performance_indicators = pd.Series(['Percentage Growth/Decline in New Channel',
                                           'Percentage Growth/Decline in New Channel 2',
                                           'Growth/Decline in USSD users',
                                           'Growth/Decline in USSD users 2',
                                           'Existing customers paying with STAR',
                                           'New customers paying with STAR',
                                           'Existing customers paying with POP',
                                           'New customers paying with POP',
                                           'Existing customers paying with New Channel',
                                           'New customers paying with New Channel',
                                           'Stickiness of new solutions',
                                           'Stickiness of old soltuions',
                                           'STAR users with  improved/declined value of payment',
                                           'POP users with  improved/declined value of payment',
                                           'STAR customers with improved/declined frequency of payment',
                                           'POP customers with improved/declined frequency of payment'])
    
  
    kpi = pd.DataFrame({'Key Performance Indicators':key_performance_indicators, 
                        'First Month':initial_month,
                        'Second Month': current_month,
                        'Percentage(%)':perform_indicators})
    #df = pd.DataFrame(kpi);
    #df = df.iloc[0][1][2]
    #print(df)
    
    performance_writer = pd.ExcelWriter('MayJunePerformance2.xlsx', engine='xlsxwriter')   

    kpi.to_excel(performance_writer, sheet_name='Results')
        
        

#Calling The main Function 
if __name__ == "__main__":
    excute()

