g# -*- coding: utf-8 -*-
"""

Authored by Andy Alba
email: andy_alba@apple.com
LinkedIn: /in/aialba

The purpose of this script is to extract the recommended labeling definition,
and the list of addressable localities/addresses from an LES Report .xlsx file.

Possible improvements:
    Delete rows where lmz value is not numeric
    User-proofing
    Option to upload .csv contents to PG
    Phase out usage of Pandas to make script more accessible
    Include option to specify alternate sheet names if defaults don't work

"""

import sys
import pandas as pd
import re

def extract_labeling_info(df, output_path):
    
    df.columns = df.iloc[0]
    df = df[df['Category'] != 'Category']
    df = df[df['Desired Zoom level \n(based on logic)'] != '-']
    df = df.dropna(axis=0, subset=['Desired Zoom level \n(based on logic)'])
    df = df.dropna(axis=0, subset=['Official Name \n(English)'])
    df= df.reset_index(drop=True)
    
    
    output = []
    for i in range(0, len(df)):
        output.append({
            'territory_name': df['Official Name \n(English)'][i],
            'logic_lmz': df['Desired Zoom level \n(based on logic)'][i],
            'local_lmz': df['Desired Zoom level (based on local expertise)'][i]
        })
    
    pd.DataFrame(output).to_csv(output_path, index=False, sep = '\t')
    return 'File has been made at: ' + output_path

def extract_address_info(df, output_path):
        
    df = df.dropna(axis=0, subset=['Address']) # remove rows where there is no address
    df = df.reset_index(drop = True) 
    
    output = [] 
    for i in range(0,len(df)):
        output.append({

            'address': df['Address'][i].strip(),

            # gets the name of locality
            'locality_name': df['Address'][i][
                re.search(
                    r'\d{1}(?!\d|.*\d{1})', 
                    df['Address'][i].replace(', ', '')
                ).start() + 4:
            ].strip(),

            'latitude': df['Location \n(X, Y coordinate)'][i].split(', ')[0],

            'longitude': df['Location \n(X, Y coordinate)'][i].split(', ')[1]
        })

    pd.DataFrame(output).to_csv(output_path, sep = ',', index = False)
    
    return 'File has been made at: ' + output_path


def extract_info(input_path, iso3):
    
    address_path = input_path[0:input_path.rfind('/')+1] + iso3 + '_les_locality_locations.csv'
    labeling_path = input_path[0:input_path.rfind('/')+1] + iso3 + '_les_locality_labeling.csv'
    
    df = pd.read_excel(input_path, sheet_name = None)
    
    df_address = df['Addressing List']
    
    print(extract_address_info(df_address, address_path))
    
    matches = [df[string] for string in list(df.keys()) if re.match(r'Display Summary.*', string)]
    print(extract_labeling_info(pd.concat(matches), labeling_path))
    
    return

def checkUsage():
    if len(sys.argv) != 3:
        print("Usage: python les_data_extraction.py <path_to_excel_file> <iso3>")
        sys.exit()
    elif '.xlsx' not in sys.argv[1] :
        print("First argument must be a path to a .xlsx file. Argument must contain .xlsx")
        sys.exit()
    elif len(sys.argv[2]) != 3:
        print("Second argument must be ISO Alpha 3 Country Code. Argument must contain 3 characters long")
        sys.exit()

def main():
    
    checkUsage()
    
    extract_info(
            sys.argv[1], # input_path
            sys.argv[2] # iso3
    )
    
    

if __name__ == '__main__':
    main()