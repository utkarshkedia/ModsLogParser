import sys,os
import numpy as np
import xlsxwriter

#to parse the inputs to a csv file
def parser(keyWord, values, unit):
    length = len(values)
    indexes = []
    for i in range(length):
        indexes.append(i)
    valuesWithHeader = [keyWord+"({})".format(unit)] + values
    indexesWithHeader = ["Index"] + indexes

    #adding the data to the xlsx file
    worksheet = workbook.add_worksheet(keyWord)
    worksheet.write_column(0, 0, indexesWithHeader)
    worksheet.write_column(0, 1, valuesWithHeader)
    worksheet.write_column(5,3,["MAX",max(values)])
    worksheet.write_column(5,4,["MIN",min(values)])
    worksheet.write_column(5,5,["AVERAGE",sum(values)/len(values)])
    worksheet.write_column(5,6,["MAX-MIN",max(values)-min(values)])
    worksheet.write_column(5,7,["MAX-AVG",max(values)-(sum(values)/len(values))])
    worksheet.write_column(5,8,["MIN-AVG",min(values)-(sum(values)/len(values))])
    worksheet.write_column(5,9,["STANDARD DEVIATION",np.std(values)])

    #creating and adding chart
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({'values': '={}!$B$2:$B${}'.format(keyWord,str(len(values)+1)),'line':   {'width': 1},})
    chart.set_title({'name': keyWord})
    chart.set_y_axis({'name': keyWord})
    worksheet.insert_chart('D10', chart)

#main code starts here
#reading back the arguments
fileName = sys.argv[1]
keyWords = sys.argv[2].split(",")

#Reading and storing the input file
cwd = os.getcwd()
filePath = os.path.join(cwd,fileName)
outputFilePath = os.path.join(cwd,fileName.split(".")[0]+"_Output.xlsx")
with open(filePath,'r') as f:
    fileData = f.readlines()

#creating an xlsx file
workbook = xlsxwriter.Workbook(outputFilePath)

for keyWord in keyWords:
    #Reading and storing the values
    values = []
    unitFound = False
    for line in fileData:
        if " " + keyWord + " " in line:
            line = line.split(":")
            value = line[-1]
            i1 = value.index("[")
            i2 = value.index("]")
            value = value[i1+1:i2]
            values.append(float(value))

            #finding the unit
            if unitFound == False:
                line[1] = line[1].split("[")
                line[1] = line[1][1].split("]")
                unit = line[1][0]
                unitFound = True

    parser(keyWord, values, unit)

#closing the xlsx file
workbook.close()