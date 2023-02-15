import requests
import openpyxl
import os
from openpyxl import Workbook

def getCustomerInfo(url):                           # extract api info from website
    customer={}
    data=requests.get(url)
    customer=data.json()
    return customer

# Idea: what if it's a text file
def getCustomerText(path):                              # extract from text file
    f = open(path, "r")                                     # open file
    test_string=f.read()                                    # read file
    f.close()   
    removethis = ["\n","\t","        "]                     # declaring what to remove, remove the extra characters like endline and the indentation

    for i in removethis:
        test_string = test_string.replace(i, '')            # remove aforementioned characters
    ts=str(test_string)
    x=ts.split("}")                                         # identify that each mapping is separated by "}", identify a {} as a separate client
    x.pop()                                                 # remove last element, last element should be empty
    count=0
    customer=[]
    for i in x:
        customer.append("")
        i="customer["+str(count)+"]="+i+"}"
        exec(i)
        count=count+1
    return customer

def datasheet(path, customer_info):                     # update excel with info from mapping
    data = openpyxl.load_workbook(path,keep_vba=True)       # open excel template
    sheet1 = data.worksheets[0]                             # info to be updated is in the first sheet
    first_section(sheet1,customer_info)                     # update the corresponding cells
    filename= customer_info['projectName']+\
        ".xlsm"
    print(filename)
    data.save(filename)                                     # save file under this name

def assignCell(x,sheet,customer):                       # assign to cell if parameter exists
    num_only=['zip','latitude']
    for key in x:
        if x[key] in customer:
            if x[key] in num_only:
                sheet[key]=float(customer[x[key]])
            else:
                sheet[key]=customer[x[key]]       

def first_section(sheet,customer_info):                 # customer section, assigning customer info to sheet and cell
    legend={        
        #"C5": "",                                           # first name
        "C6": "LastName",                                   # last name
        "C7": "streetAddress",                              # street address
        "C8": "city",                                       # city
        "C9": "state",                                      # state
        "C10": "zip",                                       # zip
        "C23": "latitude",                                  # latitude
        "C24": "longitude",                                 # longitude
    }
    assignCell(legend, sheet, customer_info)            # use legend to assign to the appropriate cell
   
if __name__ == "__main__":                              # if file is run directly
    customer=getCustomerText("Downloads/demofile.txt")      # get api info from text file
    os.chdir("OneDrive\Desktop")                            # redirect to where template is
    count=0
    for i in customer:                                  # for each customer api found in text file
        datasheet("name.xlsm", customer[count])             # save template (as new file) with customer info, name.xlsm is template excel 
        count=count+1
    

