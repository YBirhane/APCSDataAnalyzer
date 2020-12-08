import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile


# retrieves the sheet with the sheetname



# states to loop through 
stateNames = ['ALABAMA', 'ALASKA', 'ARIZONA', 'ARKANSAS', 'CALIFORNIA', 'COLORADO', 'CONNECTICUT', 'DELAWARE', 'DC', 'FLORIDA',
'GEORGIA', 'HAWAII', 'IDAHO', 'ILLINOIS', 'INDIANA', 'IOWA', 'KANSAS', 'KENTUCKY', 'LOUISIANA', 'MAINE', 'MARYLAND',
'MASSACHUSETTS', 'MICHIGAN', 'MINNESOTA', 'MISSISSIPPI', 'MISSOURI', 'MONTANA', 'NEBRASKA', 'NEVADA', 'NEW_HAMPSHIRE',
'NEW_JERSEY', 'NEW_MEXICO', 'NEW_YORK', 'NORTH_CAROLINA', 'NORTH_DAKOTA', 'OHIO', 'OKLAHOMA', 'OREGON', 'PENNSYLVANIA',
'RHODE_ISLAND', 'SOUTH_CAROLINA', 'SOUTH_DAKOTA', 'TENNESSEE', 'TEXAS', 'UTAH', 'VERMONT', 'VIRGINIA', 'WASHINGTON',
'WEST_VIRGINIA', 'WISCONSIN', 'WYOMING']

# list of lists with the demographic numbers 
data = []

for x in stateNames:
    # read the sheets for All and Females 
    df1 = pd.read_excel(x + '_Summary_10.xls', sheetname='all')
    df2 = pd.read_excel(x + '_Summary_10.xls', sheetname='females')
    
    # drop the first 3 rows 
    df1 = df1.iloc[3:]
    df2 = df2.iloc[3:]
    
    # store each demographic data 
    totalNumberOfBlack = df1['Unnamed: 8'].iloc[27]
    blackAvg = df1['Unnamed: 8'].iloc[28]

    totalNumberOfAmIndian = df1['Unnamed: 8'].iloc[13]
    amIndAvg = df1['Unnamed: 8'].iloc[14]

    totalNumberOfMexican = df1['Unnamed: 8'].iloc[34]
    mexAvg = df1['Unnamed: 8'].iloc[35]

    totalNumberOfPuerto = df1['Unnamed: 8'].iloc[55]
    puertoAvg = df1['Unnamed: 8'].iloc[56]


    totalNumberOfLatOther = df1['Unnamed: 8'].iloc[48]
    latOthAvg = df1['Unnamed: 8'].iloc[49]


    totalNumberOfAsian = df1['Unnamed: 8'].iloc[20]
    asianAvg = df1['Unnamed: 8'].iloc[21]

    totalNumberOfOther = df1['Unnamed: 8'].iloc[41]
    otherAvg = df1['Unnamed: 8'].iloc[42]


    totalNumberOfWhite = df1['Unnamed: 8'].iloc[62]
    whiteAvg = df1['Unnamed: 8'].iloc[63]


    totalNumberOfFemales = df2['Unnamed: 8'].iloc[69]

    totalNumberOfStudents = df1['Unnamed: 8'].iloc[69]
    studentAvg = df1['Unnamed: 8'].iloc[70]
  
 
    # append it to the list of lists of data 
    data.append([x,totalNumberOfStudents, studentAvg, totalNumberOfFemales, 
    totalNumberOfBlack, blackAvg,totalNumberOfAmIndian, amIndAvg, 
    totalNumberOfMexican, mexAvg, totalNumberOfPuerto, puertoAvg, 
    totalNumberOfLatOther, latOthAvg, totalNumberOfAsian, asianAvg, 
    totalNumberOfWhite, whiteAvg, totalNumberOfOther, otherAvg])


# Create the pandas DataFrame to hold the extracted info
df = pd.DataFrame(data, columns = ['State', 'Total', 'Total Average', 'Females', 
    'Black', 'Black Average', 'American Indian', 'American Indian Average', 'Latino: Mexican', 
    'Latino:Mexican Average', 'Latino: Puerto Rican', 'Latino: Puerto Rican Average', 
    'Latino: Other', 'Latino: Other Average', 'Asian', 'Asian Average', 'White', 'White Average',
    'Other', 'Other Average']) 

# create an excel file from dataframe
df.to_excel("Computer Science A 2010.xlsx")






'''

#  Experimentation 

df = pd.read_excel('ALABAMA_Summary_10.xls', sheetname='all')
df = df.iloc[3:]

# totalNumberOfBlack = df['Unnamed: 8'].iloc[27]
# totalNumberOfMexican = df['Unnamed: 8'].iloc[34]
# totalNumberOfPuerto = df['Unnamed: 8'].iloc[55]
# totalNumberOfLatOther = df['Unnamed: 8'].iloc[48]
# totalNumberOfFemales = df['Unnamed: 8'].iloc[69]
# totalNumberOfStudents = df['Unnamed: 8'].iloc[69]


print(df['Unnamed: 8'].iloc[0])
df.to_excel("Computer Science A Test.xlsx")


'''



