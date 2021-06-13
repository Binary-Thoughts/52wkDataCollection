import os
import pandas as pd
import datetime



### Getting Data Folder Location ###
datapath = input("\n\ninsert folders :- ").replace('"','')
# print(datapath)

### for total file read counter #
fileread = 0


### DateTime Stamp For Excel File Name ###
ts = datetime.datetime.now()
ts = ts.strftime("%d%b%Y_%H%M%S")



### recursively check for csv file in datapath folder and make a list ###
def getListOfFiles(dirName):

    # create a list of file and sub directories 
    # names in the given directory 
    listOfFile = os.listdir(dirName)
    allFiles = list()


    ### Iterate over all the entries ###
    for entry in listOfFile:
        # Create full path
        fullPath = os.path.join(dirName, entry)
        # If entry is a directory then get the list of files in this directory 
        if os.path.isdir(fullPath):
            allFiles = allFiles + getListOfFiles(fullPath)
        else:
            allFiles.append(fullPath)
                
    return allFiles


filelist = getListOfFiles(datapath)



### Column Name to select from csv files ###
fields = ['Symbol','New 52W/H price']




### Output file ###
outputfile = pd.DataFrame(columns=['Symbol'])







# Loop to read every csv file available one by one
for filenum in filelist :

    # print("#"*50)
    

    colname = filenum.split('.')

    
    ### read File ###

    try:
        file = pd.read_csv(f'{filenum}',usecols=fields)
    except Exception as e:
        print(f"file reading {filenum} :- error occured {e}")
        print("#"*50)

    
    # Chnaging Column Name with CSV File Name

    column_name_temp = colname[0].replace("\\",",")

    file.rename(columns = {'New 52W/H price':column_name_temp.rsplit(',',1)[1]}, inplace = True)

    

    # Merge All Files Result With Outputfile
    outputfile = pd.merge( outputfile , file , on = 'Symbol', how = 'outer')


    ### file read Counter ###
    fileread = fileread + 1


print(f"\n\nread total {fileread} files out of {len(filelist)}")






### main index is Symbol and changing column to to datetime format ###
df = outputfile.set_index("Symbol")
df.columns = pd.to_datetime(df.columns)




### grouping columns ###
out = (
    df
    .T
    .groupby(pd.Grouper(level=0, freq="MS"))
    .agg(lambda xs: ", ".join(map(str, filter(pd.notnull, xs))))
    .T
)

### changing datetime format from 01-01-2021 to Jan-2021 ###
out.columns = out.columns.strftime("%b-%Y")






### for total counts and % chnage column ###
out_max = outputfile.max(axis = 1)
out_min = outputfile.min(axis = 1)

#add new col counts
outputfile['Counts'] = outputfile.apply(lambda x: x.count()-1, axis=1)

#add new col percent change
outputfile['% Change'] = round(((out_max - out_min)/out_min)*100,2)






### writing excel file ### 
writer = pd.ExcelWriter('output_'+ts+'.xlsx', engine='openpyxl')


if os.path.exists('output_'+ts+'.xlsx'):
    book = openpyxl.load_workbook('output_'+ts+'.xlsx')
    writer.book = book

out.to_excel(writer, sheet_name='Monthly', index = True)
outputfile.to_excel(writer, sheet_name='Daily', index = False)

writer.save()
writer.close()





print(f"\n\nfile created with name -> {'output_'+ts+'.xlsx'}")