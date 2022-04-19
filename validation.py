# %%
import pandas as pd
import numpy as np
import datetime

# %%
def open_files(constellation, upload_template):
    try:
        const = pd.read_excel(constellation)
        up = pd.read_excel(upload_template, 'Add Bills-Non Electric')
        return const, up
    except:
        print("ERROR: Error trying to read-in this file. Check file name and directory are correct.\n\n")
        
def create_matching_column(up):
    up['const_customer_id'] = up['Meter Name\n(Pre-filled)'].str.split('__', expand=True)[1]
    up['const_meter_number'] = up['Meter Name\n(Pre-filled)'].str.split('__', expand=True)[2].astype('int64')
    
def compare_meters(const, up):
    #this finds the amount of meters in constellation that aren't in ESPM
    difference = len(set(const['MeterNumber']) - set(up['const_meter_number']))
    #this finds which specific meters are in the constellation raw data but not in the ESPM template
    const_not_ESPM_meters = set(const['MeterNumber']) - set(up['const_meter_number'])
    if difference == 0:
        return ""
    return (
    f'WARNING: There are **{difference}** more unique meters in Constellation than there are in the ESPM download template.\n'
    f'This mismatch could be due to unfinished meter mapping between Constellation meters and Energy Star\n'
    f'and/or there are any non-gas meters in the data, and/or there are outdated meters.\n'
    f'It is expected that there will be some difference, but here are the meters that are in constellation\n'
    f'that are not in ESPM for your own investigation if necessary:\n'
    f'{const_not_ESPM_meters}')

def find_overlapping_meters(const, up):
    overlap_meters = list(set(const['MeterNumber']) & set(up['const_meter_number']))
    assert len(overlap_meters) > 0, "There are no overlapping meters between this constellation data and ESPM upload template. Double check that these are the right files!"
    messed_up_meters = set(up['const_meter_number']) - set(overlap_meters)
    if len(messed_up_meters) > 0:
        warning = f"WARNING: This/these meter(s) {messed_up_meters} are not having their data updated, likely because of a meter naming error.\nDouble check it is named correctly in ESPM and that it is 10 digits.\n\n"
    else:
        warning = ""
    return warning, overlap_meters

def format_check(overlap_meters, const, up):
    #establish a place to store errors
    errors = ""
    
    #looping through all meters that appear in both spreadsheets
    for meter in overlap_meters:

        #grabbing relevant rows from each spreadsheet that have the same meter number
        #grabbing last occurence to get most recent bill data
        const_row = const.loc[const['MeterNumber'] == meter].iloc[-1]
        #grabbing 2nd occurence to get most recent billing data
        up_row = up.loc[up['const_meter_number'] == meter].iloc[1]

        #asserting that the row values are the same between the sheets, including Customer Id
        if const_row['CustomerId'] != up_row['const_customer_id']:
            errors += f"\nERROR: Customer ID mismatch, meter #{meter}\n"
            errors += f"\tConst Customer ID: {const_row['CustomerId']}"
            errors += f"\n\tESPM Customer ID: {up_row['const_customer_id']}\n"

        #asserting that most recent bill data has correct cycle start date
        if const_row['CycleStartDate'] != up_row['Start Date\n(Required)']:
            errors += f"\nERROR: Start Date mismatch, meter #{meter}"
            errors += f"\n\tConst Start Date: {const_row['CycleStartDate']}"
            errors += "\n\tESPM Start Date: {}\n".format(up_row['Start Date\n(Required)'])

        #asserting that most recent bill data has correct cycle end date
        if const_row['CycleEndDate'] != up_row['End Date\n(Required)']:
            errors += f"\nERROR: End Date mismatch, meter #{meter}"
            errors += f"\n\tConst End Date: {const_row['CycleEndDate']}"
            errors += "\n\tESPM End Date: {}\n".format(up_row['End Date\n(Required)'])

        #asserting that most recent bill data has correct quantity
        if const_row['FeeVolume'] != up_row['Quantity\n(Required)']:
            errors += f"\nERROR: Quantity mismatch, meter #{meter}"
            errors += f"\n\tConst Quantity: {const_row['FeeVolume']}"
            errors += "\n\tESPM Quantity: {}\n".format(up_row['Quantity\n(Required)'])

        #checking if estimation value is correct 
        if const_row['EndReadType'] == 'Actual':
            if up_row['Estimation (Required)'] != 'No':
                errors += f"\nERROR: Estimation value incorrect, meter #{meter}"
                errors += f"\n\tConst Estimation value: Actual (No)"
                errors += f"\n\tESPM Estimation value: Yes\n"
        if const_row['EndReadType'] == 'Estimate':
            if up_row['Estimation (Required)'] != 'Yes':
                errors += f"\nERROR: Estimation value incorrect, meter #{meter}"
                errors += f"\n\tConst Estimation value: Estimate (Yes)"
                errors += f"\n\tESPM Estimation value: No\n"

        #asserting that the most recent bill data has correct cost, need to manually go over NaNs
        if np.isnan(const_row['TotalCharges']) & np.isnan(up_row['Cost\n(Optional)']):
            continue
        if const_row['TotalCharges'] != up_row['Cost\n(Optional)']:
                errors += f"\nERROR: Cost mismatch, meter #{meter}"
                errors += f"\n\tConst Cost: {const_row['TotalCharges']}"
                errors += "\n\tESPM Cost: {}\n".format(up_row['Cost\n(Optional)'])
                
    if errors != "":
        errors += "\n\nGiven there are ERRORS in this section, there may be a problem with the automation tool,\n"
        errors += "because data values are not lining up where they should be between spreadsheets.\n"
        errors += "**DOUBLE CHECK** the right files are being input, and if so then contact your 2030\n"
        errors += "District data lead to sort these errors, and determine if this is the tool for you.\n\n"
    return errors

def write_errors_to_file(errors):
    with open('warnings_and_errors.txt', 'w') as f:
        f.write(errors)

# %%
def main():
    
    #establish global variable of text made up of found errors to export to a txt at the end
    warnings_and_errors = ""
    
    #import in the raw constellation data here, and the ESPM filled in upload template
    const_excel = 'Constellation_Excel_input\APPS_CONST_3.31.2022.xlsx'
    output_file = 'Output.xlsx'
    const, up = open_files(const_excel, output_file)
    
    #separating out the constellation meter name from our
    #upload sheet to compare row data from const sheet
    create_matching_column(up)
    
    #assert that these are the right files and that meter numbers overlap,
    #and grab appropriate meter #s to loop through
    warning, overlap_meters = find_overlapping_meters(const, up)
    warnings_and_errors += warning
    
    #find which meters are in Constellation that aren't in ESPM, if there are any, and print them
    warnings_and_errors += compare_meters(const, up)
    
    #ensure data is being put in the right places and in the right
    #format by checking the most recent data of every meter
    warnings_and_errors += format_check(overlap_meters, const, up)
    
    #output the collected warnings and errors
    if warnings_and_errors != "":
        write_errors_to_file(warnings_and_errors)
    
main()   


