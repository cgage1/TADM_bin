# Import necessary python libraries
import os
import csv
import pyodbc
import numpy as np
import struct
import sys
import subprocess
import matplotlib
matplotlib.use('Agg')
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from time import sleep
import ctypes  

#------INPUT PARAMS----------#
genereate_guardbandsBATLocation = r'c:\Program Files (x86)\HAMILTON\Methods\TADM\generate_guardbands.bat'
guardbandCSVLocation = '//ussd-file/Depts/Ops/MFG/ReagentsFill_Protocols/cLLD/Guardband.csv'

 # 'C:/Program Files (x86)/HAMILTON/Methods/TADM/Guardband.csv' 
 
 

#----------------------------#

# Python GUI. Used in original script to ask operator to select .mdb file. No longer used since .mdb file selection is automatic. Used in tadm_guardbands.py to request user to select file/folder
root = tk.Tk()
root.withdraw()

# Clear command prompt window
os.system('cls')

# Retrieve mdb file path from Venus > Batch > Script command line argument. File path is broken into separate Strings so need to join entire path into a single String variable.
mdb_path = ''.join(sys.argv[1:])

# Save file path for generated plots graphs. Only makes directory if it doesn't already exist. File path includes '.mdb' extension, so ignoring last 4 characters
mdb_path = mdb_path.replace("'","")
directory = mdb_path[:len(mdb_path)-4]
if not os.path.isdir(directory):
	os.mkdir(directory)

##########################################
# Initialize variables

# Used to generate and identify tube position. columnCounter increments from 1-6 after every aspirate/dispense cycle. carrierCounterBase serves as a base for carriers 1 and 2 for the first 48 tubes (for channels 1-4 and 5-8, respectively) and carriers 3 and 4 for the next 48
columnCounter = 1
columnCounterMax = 6
carrierCounterBase = 1
# fillCounter counts number of entries in results.csv file. Used to retrieve the next 96 tubes worth of TADM data
fillCounter = '-1'
guardbandIdxBase = -1

# Fill volumes over 1000uL are split into two aspirate/dispense steps with half of the volume. If volume exceeds 1mL, set flag
splitVolumeFlag = 0
splitVolumeCounter = 0

# startBoundary and endBoundary ignore the starting and ending milliseconds based on assigned value
startBoundary = 25
endBoundary = 35
press_length = 0

##########################################
# Check if csv file exists. If it exists, update counter to retrieve only new data since last run #

if os.path.isfile(directory + '/results.csv'):
	# fileResultsReader can only read the data from the files, not able to make any additions/modifications to file
	fileResultsReader = open(directory + '/results.csv')
	readerResults = csv.reader(fileResultsReader)
	# Example: this the third rack, so 192 tubes have been completed. Each tube has an aspirate and dispense, so 2 rows per tube. 192 * 2 = 384, so there should be 384 rows in the results.csv file. fillCounter is set to 384, so the next run it will only pull results from 385+
	fillCounter = str(len(list(readerResults)) - 1)
	fileResultsReader.close()

# fileResults is able to read from and write to the files. the 'a' denotes "append" so the writer knows to append data to the end of the file rather than overwriting existing rows
fileResults = open(directory + '/results.csv', 'a', newline='')
writerResults = csv.writer(fileResults, delimiter=',')
# Insert column headers if file is empty. fillCounter is initially set to -1 in the Initialize Variables section above. If it's still -1, we know the file is empty
if fillCounter == '-1':
	# Since file is empty, insert the column headers
	# All column headers explained further down when results are actually being saved, since each of these columns will be explained
	writerResults.writerow(['Step', 'Step Type', 'Result', 'Carrier', 'Channel', 'Column', 'Coordinate', 'Processed', 'CurveId'])

# Update the current working directory of the script to the folder path, so graphs will be saved in this folder
save_path=os.path.abspath(directory)
os.chdir(save_path)

##########################################
# Retrieve guardband data #

# Checking for guardband.csv file within a while loop to give operators option to re-attempt opening the guardband.csv file
retryReadGuardband = True
while retryReadGuardband:
	try:
		# Update filepath of guardband file
		fileGuardband = open(guardbandCSVLocation)
		readerGuardband = csv.reader(fileGuardband)
		# List to hold all of the rows in the guardband.csv file
		l_guardband=[]
		for idxGuardband, row in enumerate(readerGuardband): 
		# Filter out empty values from guardband since length of all rows is determined by the length of the longest row
			l_guardband.append(list(filter(None, row)))
		# Disable flag to allow while loop to end if guardband file has been found and read
		retryReadGuardband = False
	# If guardband can't be found/opened, allow operators the option to retry
	except Exception as e:
		while True:
			print(' Error '.center(120, '#'))
			inputRetry = input('Guardband file not found. Retry? (Y/N): ')
			if inputRetry.lower() in ('y', 'n'):
				break
			print('Invalid entry.')
		if inputRetry.lower() == 'y':
			print('Retrying..')
			continue
		elif inputRetry.lower() == 'n':
			print('Aborting..')
			# If operator decides to not retry, aborts method and returns exit code 101 to the method
			sys.exit(101)

##########################################
# Retrieve TADM data

# Checking for the {run}.mdb file within a while loop to give operators option to re-attempt opening the .mdb file
retryReadMDB = True
while retryReadMDB:
	try:
		# Open a connection to the .mdb file to be able to query data from it
		MDB = mdb_path
		DRV = '{Microsoft Access Driver (*.mdb, *.accdb)}'
		conn = pyodbc.connect('DRIVER={};DBQ={}'.format(DRV,MDB))
		cursor = conn.cursor()
		# Querying top 192 to get the aspirate and dispense data for 96 tubes in a rack
		# TadmCurve is the name of the table in the .mdb file
		# CurveId is a unique identifer for each row, basically the counter for the row number
		# fillCounter value assigned above based on existing rows in results.csv file. CurveId > fillCounter gets the next 192 rows that have not yet been read
		# Query results saved to 'cursor' variable
		cursor.execute('select top 192 * from TadmCurve where CurveId > ' + fillCounter + ' order by CurveId asc')
		# Disable flag to allow while loop to end if mdb file has been found and read
		retryReadMDB = False
	# If mdb can't be found/opened, allow operators the option to retry
	except Exception as e:
		while True:
			print(' Error '.center(120, '#'))
			inputRetry = input('MDB file not found. Retry? (Y/N): ')
			if inputRetry.lower() in ('y', 'n'):
				break
			print('Invalid entry.')
		if inputRetry.lower() == 'y':
			print('Retrying..')
			continue
		elif inputRetry.lower() == 'n':
			print('Aborting..')
			# If operator decides to not retry, aborts method and returns exit code 102 to the method
			sys.exit(102)

##########################################
# Parse TADM mdb data and save as local list variables

# Save all results from above query into a list
l_mdb=[]
for row in cursor.fetchall():
    l_mdb.append(row)
print('l_mdb length: ', len(l_mdb))
liq_class = l_mdb[0][1]
liq_class_vol_index = liq_class.rfind('_')
liq_class_vol = liq_class[liq_class_vol_index + 1:]
if int(liq_class_vol) > 1000:
	splitVolumeFlag = 1
	cursor.execute('select top 384 * from TadmCurve where CurveId > ' + fillCounter + ' order by CurveId asc')
	l_mdb=[]
	for row in cursor.fetchall():
	    l_mdb.append(row)
	print('l_mdb length: ', len(l_mdb))

#Unpack the pressure data. Pressure data is saved in a long binary form within mdb file and needs to be unpacked, which results in a list of pressure measurements, one value for each ms
l_data=[]
for row in l_mdb:
	# Unpack bytes from column 9, which is CurvePoints
	tadm_byt=row[8]
	num_byt=len(tadm_byt)/2
	fmt=str(int(num_byt))+'h'
	tadm_num=struct.unpack(fmt,tadm_byt)
	
	# Translate column 3 (StepType) to Aspirate or Dispense based on value
	if row[2]==-533331728:
		step_type='Aspirate'
	elif row[2]==-533331727:
		step_type='Dispense'
	else:
		step_type='Unknown'

	# Round volume to integer
	volume = round(row[3])
	
	# Save data to a new list, replacing StepType, Volume, and CurvePoints with the above changes
	# LiquidClassName, StepType, Volume, TimeStamp, StepNumber, CurvePoints, ChannelNumber, CurveId
	l_data.append([row[1],step_type,volume,row[4],row[7],tadm_num,row[10],row[0]])

# Save the above list into a tadm_mdb.csv file. Not used for anything within this script, just to view the data
np.savetxt('tadm_mdb.csv',l_data,fmt='%s', delimiter=',')

##########################################

"""
# Create a list of all liquid classes used in the run. Only relevant if there are multiple liquid classes used in a single run, which isn't the case in HSL.
l_lc=[]
for row in l_data:
	l_lc.append(row[0])
s_lc=set(l_lc)
print('Set of liquid classes: ', s_lc)

# Main loop to process data
# For loop through every liquid class
for item in s_lc:
liq_class = item
print('Liquid class currently processing: ',liq_class)

# Remove MN and Volume suffices from LC name. Disabled for now to use unique guardbands for every liquid class
# Enable if guardbands can be shared across same liquid classes for different MN/Volume combinations
liq_class_vol_index = liq_class.rfind('_')
liq_class_vol = liq_class[liq_class_gen_index:]
liq_class_gen_index = liq_class_gen.rfind('_')
liq_class_gen = liq_class_gen[0:liq_class_gen_index]
"""

# Testing unique liquid classes
liq_class_gen = liq_class
print('General liquid class currently processing: ',liq_class_gen)
print('Volume: ', volume)
print('LC Volume: ', liq_class_vol)

# Identify guardband index value from .csv based on liquid class name and volume
# guardbandIdxBase will be used later to identify which row in the guardband.csv file to retrieve data from
# This will continuously loop until the guardband.csv contains the liquidclass for this order 
fileGuardband.close()
while guardbandIdxBase == -1:
    for idxGuardbandBase, guardbandRow in enumerate(l_guardband):
    	if int(guardbandRow[1]) == int(volume):
    		if guardbandRow[2] == liq_class_gen:
    			guardbandIdxBase = idxGuardbandBase
    			break
    print('Liquid class index: ', guardbandIdxBase)
    if guardbandIdxBase == -1:
        ctypes.windll.user32.MessageBoxW(0, "****** ESCALATE TO RTO PROCESS ENGINEERING ******\nIssue: No TADM guardbands exist for this liquid class yet.", "****** ALERT ******", 1)
        print('Guardband for liquid class not found. Follow prompted instructions: \n\nPress 0 -> Select most recent MDB file (which should be the MDB file for this run).\n')
        # Now run bat file to generate bands
        ## subprocess.call([genereate_guardbandsBATLocation])   this errors out due to permissions errors 
        tempStopInput = input('Once completed, press any key to continue method...')
        # Now reloop through the newlyupdated guardband file and check
        retryReadGuardband = True
        while retryReadGuardband:
            try:
        		# Update filepath of guardband file
                fileGuardband = open(guardbandCSVLocation)
                readerGuardband = csv.reader(fileGuardband)
                # List to hold all of the rows in the guardband.csv file
                l_guardband=[]
                for idxGuardband, row in enumerate(readerGuardband): 
                    # Filter out empty values from guardband since length of all rows is determined by the length of the longest row
                    l_guardband.append(list(filter(None, row)))
                # Disable flag to allow while loop to end if guardband file has been found and read
                retryReadGuardband = False
            # If guardband can't be found/opened, allow operators the option to retry
            except Exception as e:
                while True:
                    print(' Error '.center(120, '#'))
                    inputRetry = input('Guardband file not found. Retry? (Y/N): ')
                    if inputRetry.lower() in ('y', 'n'):
                        break
                    print('Invalid entry.')
                if inputRetry.lower() == 'y':
                    print('Retrying..')
                    continue
                elif inputRetry.lower() == 'n':
                    print('Aborting..')
        			# If operator decides to not retry, aborts method and returns exit code 101 to the method
                    sys.exit(101)
    	#quit()

# Pull data from the overall list for steps that use the current liquid class
# Should be all of the data since the entire run uses a single liquid class
lc_data=[]
for row in l_data:
	if row[0]==liq_class:
		lc_data.append(row)

# Determine each step ID used for commands using the current liquid class. Step ID will be used to help identify the exact tube
# One aspirate or dispense step is saved as one StepNumber. Since 8 channels are aspirated/dispensed at a time, there will be 8 rows per StepNumber
step_list=[]
for row in lc_data:
	step_list.append(row[4])
# Remove duplicates from list of StepNumbers since each one will be repeated 8 times
step_list_ordered=list(dict.fromkeys(step_list))
print('Step set: ', step_list_ordered)

# For loop through each step
for step in step_list_ordered:
	print('Step currently processing: ', step)
	# Save data(StepNumber, CurvePoints, ChannelNumber) for all 8 rows of current step being processed in a list
	# StepType and TimeStamp are saved locally since it should be the same for all 8 channels
	tadm=[]
	for row in lc_data:
		if row[4]==step:
			tadm.append([row[5],row[6],row[7]])
			type_step=row[1]
			time_step=row[3]

	#Convert time format to YY-mm-DD HH-MM-SS
	s_ts=time_step.strftime('%Y-%m-%d %H:%M:%S')
	l_ts=list(s_ts)
	for n, i in enumerate(l_ts):
		if i==':':
			l_ts[n]='-'
	s_ts=''.join(l_ts)

	# Set up to generate the pressure curve graphs
	f1=plt.figure(1, figsize=(12,6))
	ax1=plt.subplot(111)
	ax1.set_title(type_step)
	ax1.yaxis.grid(True, linestyle='-', which='major', color='grey', alpha=0.5)
	ax1.xaxis.grid(True, linestyle='-', which='major', color='grey', alpha=0.5)
	ax1.axhline(0, color='black')
	ax1.set_ylabel('Pressure [Pa]',fontsize=14)
	ax1.set_xlabel('Time [ms]',fontsize=14)


	# For every liquid class and volume, there are 6 rows in the guardband.csv file
	# Lines 1 - 3: Liquid Class, Volume, and StepType (Aspirate) identifier, followed by Upper and Lower limit for Aspirate
	# Lines 4 - 6: Liquid Class, Volume, and StepType (Dispense) identifier, followed by Upper and Lower limit for Dispense
	# Depending on step type, set the starting point to either Line 1 or Line 4 for that liquid class
	if type_step == 'Aspirate':
		guardbandIdx = guardbandIdxBase + 1
	elif type_step == 'Dispense':
		guardbandIdx = guardbandIdxBase + 4

	# Retrieves pressure data either 1 or 2 rows below the starting point identified above to retrieve upper or lower limit
	press_upper = [int(i) for i in l_guardband[guardbandIdx]]
	press_lower = [int(i) for i in l_guardband[guardbandIdx + 1]]

	# Plot upper and lower guardbands in red, using duration of the upper guardband as the base duration
	time=np.arange(0,len(press_upper),1)
	plt.plot(time,press_upper,color = 'red')
	plt.plot(time,press_lower,color = 'red')

	# For loop through each channel in this aspirate/dispense step
	for idxTadm, row in enumerate(tadm):
		ch_num = str(row[1])
		# ch_num_adj will be used as the row identifier when generating the coordinate for the tube since the 8 channels are split into 2 carriers of 4 tubes each
		ch_num_adj = ch_num # Tried ch_num % 4 but didn't work for some reason
		# Result set to Pass by default, update to Fail if failure detected. fail_flag to identify when tubes failed in either aspirate or dispense
		result = 'Pass'
		fail_flag = 0
		# Save pressure of actual run to press, compare each pressure measurement to corresponding upper/lower guardband value
		press=np.asarray(row[0],dtype='int')
		# Flag any instances where upper/lower guardbands are breached
		# Time saved as length of pressure, as each measurement denotes the pressure at that ms
		time=np.arange(0,len(press),1)
		press_length = len(press)
		# For loop through each pressure value
		for inx, row_press in enumerate(press):
			# If failed at any point, mark tube as fail and move on
			if fail_flag == 0:
				# Check if current time is within guardband duration
				if inx < len(press_upper):
					# Only compare to upper/lower limit if time is within the start and end boundary, explained above in variables
					if time[inx] > startBoundary:
						if time[inx] < press_length - endBoundary: 
							# Flag tube if current pressure is above the upper limit or below the lower limit
							if row_press > press_upper[inx]:
								print('Upper limit passed @ channel ', ch_num, ', Actual: ', row_press, ' Expected: ', press_upper[inx], ' at ', inx, 'ms')
								fail_flag = 1
								result = 'Fail'
							if row_press < press_lower[inx]:
								print('Lower limit passed @ channel ', ch_num, ', Actual: ', row_press, ' Expected: ', press_lower[inx], ' at ', inx, 'ms')
								fail_flag = 1
								result = 'Fail'
		# 8 channels are used to dispense 4 tubes into 2 carriers at a time. Identify if this channel is in the first vs second carrier (or third vs fourth for the next 48 tubes)
		if int(ch_num) <= 4:
			carrierCounter = carrierCounterBase
		if int(ch_num) > 4:
			carrierCounter = carrierCounterBase + 1
			ch_num_adj = str(int(ch_num) - 4)
		# Tube coordinates are written as (Carrier)-(Row)(Column) i.e. 2-A4 would be carrier 2, row A, column 4
		# Need to convert channel number to corresponding row (1 -> A, 2 -> B, etc.) Look up ASCII Table to find character to integer conversion for reference
		ch_num_char = chr(ord('A') + int(ch_num_adj) - 1)

		# Write results into the results.csv file for this channel
		# Step, StepType, Result, Carrier, Channel, Column, Coordinate, Processed, CurveId
		# Step: Step Number, used to identify the 8 channels that aspirated/dispensed together
		# StepType: Aspirate or Dispense
		# Result: Result of comparison to guardband limits
		# Carrier: Identify if first or second carrier, depending on which channel (1-4 = first, 5-8 = second)
		# Channel: Identify which channel, converted into A-B-C-D depending on channel to identify which row in the carrier
		# Column: Identify which column on the carrier. Aspirates/Dispenses to Column 1, then Column 2, then Column 3, etc. until the last column (6). After 6, will move to next set of two carriers and reset column to 1
		# Coordinate: Combination of Carrier, Channel, and Column in #-## format
		# Processed: Initially written as 'No'. Once Python script writes the results.csv file and ends, Venus method will read the results.csv file and look only at rows where Processed is 'No'. Once it finishes, it'll overwrite to 'Yes' so subsequent calls to the same results.csv file don't produce duplicates
		# CurveId: Unique identifier for each row, essentially row number - 1 to account for column header row
		writerResults.writerow([step, type_step, result, carrierCounter, ch_num_char, columnCounter, str(carrierCounter) + '-' + ch_num_char + str(columnCounter), 'No', row[2]])
		
		# Plot the pressure curve, using the channel number as the point
		plt.scatter(time,press,label=str(carrierCounter) + '-' + ch_num_char + str(columnCounter),marker='$'+ch_num+'$',s=16, linewidths=0.05)
		# If this is step is the Dispense for the last channel (last of the set of 8 tubes), reset some of the variables
		if (type_step == 'Dispense') and (ch_num == '8'):
			# Increment the column. If this is a split volume aspirate/dispense, increment the column every other step
			if splitVolumeFlag == 0:
				columnCounter = columnCounter + 1
			elif splitVolumeFlag == 1:
				columnCounter = columnCounter + splitVolumeCounter
				splitVolumeCounter = (splitVolumeCounter + 1) % 2
			# If this is the last column for this carrier, increment the carrier to the next 2 (1,2 -> 3,4) and reset the column
			if columnCounter > columnCounterMax:
				columnCounter = 1
				carrierCounterBase = carrierCounterBase + 2
				if carrierCounter >= 4:
					carrierCounterBase = 1

	# Plot vertical blue lines for the start/end boundary.
	plt.axvline(x = startBoundary)
	plt.axvline(x = press_length - endBoundary)
	handles, labels = ax1.get_legend_handles_labels()
	f1.text(0.15, 0.93, time_step, ha='center')
	plt.tight_layout(rect=[0.05, 0.05, 0.95, 0.95])
	plt.suptitle(str(step)+'_'+liq_class, fontweight='bold')
	plt.legend(loc = 'upper center', bbox_to_anchor = (0.5, -0.125), fancybox = True, ncol = 8, fontsize = 'small')
	os.chdir(save_path)
	plt.savefig(str(s_ts)+'_'+liq_class+'_step'+str(step)+'.png',dpi=300)
	#plt.show()
	plt.close(f1)
# Close out the results.csv and guardband.csv files
fileResults.close()
fileGuardband.close()
sys.exit(0)
