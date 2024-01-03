# import
import pandas as pd
import re
import os
import numpy as np

# function 
def clean_and_format(df):
    column_to_clean = ['Card Issuing Bank (System)', 'Card Type']
    df[column_to_clean] = df[column_to_clean].replace('--', '')

    df = df.drop('Status', axis=1)

    df['Card Type'] = df['Card Type'].apply(lambda x : 'Credit Card' if x == 'CREDIT' else ('Debit Card' if x == 'DEBIT' else x ))
    
    df['Exp Year 2'] = df['Exp Year'].str[2:]

    df['Exp Date'] = np.where((df['Exp Month'] != '') & (df['Exp Year 2'] != ''),
                                df['Exp Month'] + '/' + df['Exp Year 2'], 
                                '')
    
    df = df.drop('Exp Year 2', axis=1)

    df['Card Issuing Bank (System)'] = df['Card Issuing Bank (System)'].str.title()
    df['Card Issuing Bank (System)'] = df['Card Issuing Bank (System)'].apply(lambda x: format_spelling(x))

    return df

def format_spelling(x):
    if x == 'Aeon Credit Service (M) Berhad':
        return 'Aeon Credit'
    elif x == 'Affin Bank Berhad':
        return 'AFFIN Bank Berhad'
    elif x == 'Alliance Bank Malaysia Berhad':
        return 'Alliance Bank Malaysia Berhad'
    elif x == 'Ambank (M) Berhad':
        return 'AmBank (M) Berhad'
    elif x == 'Baiduri Bank Berhad':
        return 'Baiduri bank'
    elif x == 'Baiduri Bank Bhd':
        return 'Baiduri bank'
    elif x == 'Bank Islam Brunei Darussalam Berhad':
        return 'Bank Islam Brunei Darussalam'
    elif x == 'Bank Islam Malaysia Berhad':
        return 'Bank Islam Malaysia Berhad'
    elif x == 'Bank Muamalat Malaysia Berhad':
        return 'Bank Muamalat Malaysia Berhad'
    elif x == 'Bank Of China (Malaysia) Berhad':
        return 'Bank Of China (Malaysia) Berhad'
    elif x == 'Bank Simpanan Nasional':
        return 'Bank Simpanan National'
    elif x == 'Cimb Bank Berhad':
        return 'CIMB Bank Berhad'
    elif x == 'Citibank Berhad':
        return 'Citibank Berhad'
    elif x == 'Hong Leong Bank Berhad':
        return 'Hong Leong Bank Berhad'
    elif x == 'Hsbc Bank Malaysia Berhad':
        return 'HSBC Bank Malaysia Berhad'
    elif x == 'Malayan Banking Berhad':
        return 'Malayan Banking Berhad'
    elif x == 'Ocbc Bank (Malaysia) Berhad':
        return 'OCBC Bank (Malaysia) Berhad'
    elif x == 'Public Bank Berhad':
        return 'Public Bank Berhad'
    elif x == 'Rhb Bank Berhad':
        return 'RHB Bank Berhad'
    elif x == 'Standard Chartered Bank Malaysia Berhad':
        return 'Standard Chartered Bank Malaysia Bhd'
    elif x == 'United Overseas Bank (Malaysia) Berhad':
        return 'United Overseas Bank (Malaysia) Berhad'
    elif x == 'United Overseas Bank, Ltd.':
        return 'United Overseas Bank, Ltd.'
    else:
        return x



# main function
def main():
    # input folder path. Edit path accordingly.
    folder_path = r'C:\Users\mfmohammad\OneDrive - UNICEF\Desktop\Paydollar vs SF Task\Dec\191223-311223\test'

    file_name = 'order.xlsx'

    files = os.listdir(folder_path)

    file_path = os.path.join(folder_path, file_name)

    df = pd.read_excel(file_path)

    df = clean_and_format(df)


    #print('Enter start date:')
    #day1 = str(input('Enter day for start date in dd: '))
    #month1 = str(input('Enter month for the start date in mm: '))
    #year1 = str(input('Enter year for the start date in yy: '))

    #print('Enter end date:')
    #day2 = str(input('Enter day for start date in dd: '))
    #month2 = str(input('Enter month for the start date in mm: '))
    #year2 = str(input('Enter year for the start date in yy: '))


    #new_file_name = f'Online Donation - {day1}{month1}{year1}-{day2}{month2}{year2} - Paydollar.xlsx'
    new_file_name = 'test.xlsx'
 
    new_file_path = os.path.join(folder_path, new_file_name)
    

    df.to_excel(new_file_path, index=False)
    print(f'File {new_file_name} been saved in the folder')




if __name__ == "__main__":
    main()
    