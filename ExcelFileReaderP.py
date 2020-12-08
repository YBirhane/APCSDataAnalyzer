import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile


# retrieves the sheet with the sheetname



stateNames = ['alabama', 'alaska', 'arizona', 'arkansas', 'california', 'colorado', 'connecticut', 'delaware', 'district-of-columbia', 'florida',
'georgia', 'hawaii', 'idaho', 'illinois', 'indiana', 'iowa', 'kansas', 'kentucky', 'louisiana', 'maine', 'maryland',
'massachusetts', 'michigan', 'minnesota', 'mississippi', 'missouri', 'montana', 'nebraska', 'nevada', 'new-hampshire',
'new-jersey', 'new-mexico', 'new-york', 'north-carolina', 'north-dakota', 'ohio', 'oklahoma', 'oregon', 'pennsylvania',
'rhode-island', 'south-carolina', 'south-dakota', 'tennessee', 'texas', 'utah', 'vermont', 'virginia', 'washington',
'west-virginia', 'wisconsin', 'wyoming']

stateNames1 = ['alabama', 'alaska', 'arizona', 'arkansas', 'california', 'colorado', 'connecticut', 'delaware', 'district-of-columbia', 'florida',
'georgia', 'hawaii', 'idaho', 'illinois', 'indiana', 'iowa', 'kansas', 'kentucky', 'louisiana', 'maine', 'maryland',
'massachusetts', 'michigan', 'minnesota', 'mississippi', 'missouri', 'montana']
stateNames2 = ['nebraska', 'nevada', 'new-hampshire',
'new-jersey', 'new-mexico', 'new-york', 'north-carolina', 'north-dakota', 'ohio', 'oklahoma', 'oregon', 'pennsylvania',
'rhode-island', 'south-carolina', 'south-dakota', 'tennessee', 'texas', 'utah', 'vermont', 'virginia', 'washington',
'west-virginia', 'wisconsin', 'wyoming']
# list of lists with the demographic numbers 
data = []
sheetnameAll = ''
sheetnameFemale = ''

for x in stateNames:
    # since formatting is different for half the states check sheetname 
    if x in stateNames1:
        sheetnameAll = 'ALL'
        sheetnameFemale = 'FEMALES'
    else: 
        sheetnameAll = 'All'
        sheetnameFemale = 'Females'

    # read the sheets for All and Females 
    df1 = pd.read_excel(x + '-summary-2019.xlsx', sheetname=sheetnameAll)
    df2 = pd.read_excel(x + '-summary-2019.xlsx', sheetname=sheetnameFemale)
    
    # drop the first 3 rows 
    df1 = df1.iloc[4:]
    df2 = df2.iloc[4:]

  
    # store each demographic data 
    totalNumberOfBlack = df1['Unnamed: 11'].iloc[20]
    blackAvg = df1['Unnamed: 11'].iloc[21]

    totalNumberOfLat = df1['Unnamed: 11'].iloc[27]
    latAvg = df1['Unnamed: 11'].iloc[28]

    totalNumberOfAmIndian = df1['Unnamed: 11'].iloc[6]
    amIndAvg = df1['Unnamed: 11'].iloc[7]


    totalNumberOfWhite = df1['Unnamed: 11'].iloc[41]
    whiteAvg = df1['Unnamed: 11'].iloc[42]

    totalNumberOfPacific = df1['Unnamed: 11'].iloc[34]
    pacificAvg = df1['Unnamed: 11'].iloc[35]

    totalNumberOfAsian = df1['Unnamed: 11'].iloc[13]
    asianAvg = df1['Unnamed: 11'].iloc[14]

    totalNumberOfOther = df1['Unnamed: 11'].iloc[55]
    otherAvg = df1['Unnamed: 11'].iloc[56]

    totalNumberOfFemales = df2['Unnamed: 11'].iloc[69]

    totalNumberOfStudents = df1['Unnamed: 11'].iloc[69]
    studentAvg = df1['Unnamed: 11'].iloc[70]

    


    # append it to the list of lists of data 
    data.append([x,totalNumberOfStudents, studentAvg, totalNumberOfFemales, 
    totalNumberOfBlack, blackAvg,totalNumberOfAmIndian, amIndAvg, 
    totalNumberOfLat, latAvg, totalNumberOfAsian, asianAvg, 
    totalNumberOfWhite, whiteAvg, totalNumberOfPacific, pacificAvg, totalNumberOfOther, otherAvg])


# Create the pandas DataFrame to hold the extracted info
df = pd.DataFrame(data, columns = ['State', 'Total', 'Total Average', 'Females', 
    'Black', 'Black Average', 'American Indian', 'American Indian Average', 'Latino/Hispanic', 
    'Latino/Hispanic Average', 'Asian', 'Asian Average', 'White', 'White Average', 'Pacific Islander', 
    'Pacific Islander Average', 'Other', 'Other Average']) 
 

# create an excel file from dataframe
df.to_excel("Computer Science P 2019.xlsx")





'''
#  Experimentation 

df = pd.read_excel('alabama-summary-2019.xlsx', sheetname='ALL')
df = df.iloc[4:]

# totalNumberOfBlack = df['Unnamed: 10'].iloc[20]
# totalNumberOfHispanic = df['Unnamed: 10'].iloc[27]
# totalNumberOfFemales = df['Unnamed: 10'].iloc[69]
# totalNumberOfStudents = df['Unnamed: 10'].iloc[69]


print(df['Unnamed: 11'].iloc[0])
df.to_excel("Computer Science A Test.xlsx")
'''



