import xlrd
import csv
import matplotlib.pyplot as mt
import numpy as np

def trapezoid(b1, b2, h):
    return h*(b1 + b2)*0.5

#gives the location of an item in a list of items.
def indexer(value, list):
    index = 0
    indices = []
    listvalues = []
    if value in list:
        while (index + 1) <= len(list):
            if value == list[index]:
                indices.append(index)
                listvalues.append(list[index])
            index += 1
        return indices, listvalues
    return 'no such value in list'

#given any value, returns the location (index) of which number in a list of numbers is closest to the value specified.
def approx_index(value, list):
    counter = 0
    residual = []
    if value == None or value in list:
        return indexer(value, list)
    else:
        #make vector of the form [X1-value, X2-value,..., Xn-value...]
        #find index of min value of vector (location of the number closest to value)
        for every in list:
            residual.append(abs(value - every))
        return indexer( min(residual), residual)

def baseline(array, low, high): #use to average baseline given some array of values.
    counter = 0
    sum = 0
    adjusted_array = []
    subarray = array[low : high]
    for every in subarray:
        sum += every
        counter += 1
    average = sum / float(counter)
    for every in array:
        value = every - average
        adjusted_array.append(value)
    return average , np.array(adjusted_array)    

def format_list(list):
    counter = 0
    print 'Dataset','\t','Integral', '\n',
    print '--------------------------'
    while counter + 1 <= len(list):
        print counter+1, '\t', list[counter] ,'\n',
        counter += 1
    print '-----------END------------'
        
def integrate_sub(filename, excelsheet, startx, stopx, xaxis, yaxis, ydata):
    data = xlrd.open_workbook(filename)
    sheet = data.sheet_by_index(excelsheet)
    rowpos = 0 #tracks current row
    rows = sheet.row_values(0) #returns values of row n as a list
    colcounter = 1
    xdata = sheet.col_values(0)
    df = 0
    rowcount = []
    integrals = []
    if (startx == 'start' and stopx == 'end'):
        xdata2 = xdata
        ydata2 = ydata
        print 'condition 1'
        while rowpos+1 < len(xdata2):
            df += trapezoid(ydata2[rowpos], ydata2[rowpos + 1], (xdata2[rowpos + 1] - xdata2[rowpos]) )
            rowpos += 1
            #rowcount.append(rowpos) #for debugging
        integrals.append(df)
        mt.fill_between(xdata2, ydata2)
        mt.plot(xdata, ydata, 'r-', linewidth = 1.0)
        mt.xlabel(xaxis, fontsize = 14, style = 'normal', fontweight = 'bold')
        mt.ylabel(yaxis, fontsize = 14, style = 'normal', fontweight = 'bold')
        mt.show()
        return integrals
    elif (startx == 'start' and type(stopx) == int):
        print 'condition 2'
        #find index where xdata equals stop value
        end = approx_index(stopx, xdata)[0][0]  + 1
        #modify list lengths
        xdata2 = xdata[:end]
        ydata2 = ydata[:end]
        while rowpos+1 < len(xdata2):
            df += trapezoid(ydata2[rowpos], ydata2[rowpos + 1], (xdata2[rowpos + 1] - xdata2[rowpos]) )
            rowpos += 1
        integrals.append(df)
        mt.fill_between(xdata2, ydata2)
        mt.plot(xdata, ydata, 'r-', linewidth = 1.0)
        mt.xlabel(xaxis, fontsize = 14, style = 'normal', fontweight = 'bold')
        mt.ylabel(yaxis, fontsize = 14, style = 'normal', fontweight = 'bold')
        mt.show()
        return integrals
    elif (type(startx) == int and stopx == 'end'):
        print 'condition 3'
        #find index where xdata equals start value
        begin = approx_index(startx, xdata)[0][0]
        #modify list lengths
        xdata2 = xdata[begin:]
        ydata2 = ydata[begin:]
        while rowpos+1 < len(xdata2):
            df += trapezoid(ydata2[rowpos], ydata2[rowpos + 1], (xdata2[rowpos + 1] - xdata2[rowpos]) )
            rowpos += 1
            #rowcount.append(rowpos) #for debugging
        integrals.append(df)
        mt.fill_between(xdata2, ydata2)
        mt.plot(xdata, ydata, 'r-', linewidth = 1.0)
        mt.xlabel(xaxis, fontsize = 14, style = 'normal', fontweight = 'bold')
        mt.ylabel(yaxis, fontsize = 14, style = 'normal', fontweight = 'bold')
        mt.show()
        return integrals
    else:
        print 'condition 4'
        #find index where xdata equals start value
        begin = approx_index(startx, xdata)[0][0]
        #find index where xdata equals stop value
        end = approx_index(stopx, xdata)[0][0] +1
        #modify list lengths
        xdata2 = xdata[begin:end]
        ydata2 = ydata[begin:end]
        while rowpos+1 < len(xdata2):
            df += trapezoid(ydata2[rowpos], ydata2[rowpos + 1], (xdata2[rowpos + 1] - xdata2[rowpos]) )
            rowpos += 1
            #rowcount.append(rowpos) #for debugging
        integrals.append(df)
        mt.fill_between(xdata2, ydata2)
        mt.plot(xdata, ydata, 'r-', linewidth = 1.0)
        mt.xlabel(xaxis, fontsize = 25, style = 'normal', fontweight = 'bold')
        mt.ylabel(yaxis, fontsize = 25, style = 'normal', fontweight = 'bold')
        mt.show()
        return integrals

def integrate():
    print
    print
    print 'General Directions:'
    print 'Spreadsheets  to be parsed cannot contain any strings values (numbers only)'
    print 'When entering strings, be sure to surround them in quotes.'
    print 'Filenames must be entered with the file extension i.e. "integrator.xls"'
    print 'This program treats the first sheet tab in a spreadsheet as sheet zero.'
    print
    print
    filename = input("Enter filename of the spreadsheet: ")
    excelsheet = input("In which sheet tab are the data located (first sheet = 0)? ")
    start = input("At what value will you start integration? (Enter 'start' if unsure or integrating whole spectrum) ")
    stop = input("At what value will you end integration? (Enter 'end' if unsure or integrating whole spectrum) ")
    xaxis = input("Enter an x-axis label: ")
    yaxis = input("Enter a y-axis label: ")
    data = xlrd.open_workbook(filename)
    sheet = data.sheet_by_index(excelsheet)
    colpos = 1 #tracks current row
    rows = sheet.row_values(0) #returns values of row n as a list
    yarray = []
    integrals = []
    sub = []
    xdata = sheet.col_values(0)
    corrected_data = [xdata]
    responses = ['yes', 'Yes', 'y', 'Y', 'YES']
    base = input("Do need to baseline adjust? (y/n) ")
    if base in responses:
        low = input("Enter lower bound for averaging: ")
        high = input("Enter upper bound for averaging: ")
        low_ind = approx_index(low, xdata)[0][0]
        high_ind = approx_index(high, xdata)[0][0]
    else:
        low = ''
        high = ''
    while colpos < len(rows):
        ydata = sheet.col_values(colpos)
        yarray.append(ydata)
        colpos += 1
    if type(low) == int or type(low) == float:
        for every in yarray:
            adjusted_array = baseline(every, low_ind, high_ind)[1]
            corrected_data.append(adjusted_array)
            integrals += integrate_sub(filename, excelsheet, start, stop, xaxis, yaxis, adjusted_array)
        c_array = np.array(corrected_data)
        c_arrayT = np.transpose(c_array)
        #writing adjusted data to file:
        saveto = input("Enter a filename (*.csv) to save to: ")
        savefile = open(saveto, 'w')
        writer = csv.writer(savefile)
        for every in c_arrayT:
            writer.writerow(every)
        savefile.close()
    else:
        for every in yarray:
            integrals += integrate_sub(filename, excelsheet, start, stop, xaxis, yaxis, every)
    return format_list(integrals)
    
