# v2 asks for user input for file/folder and unique liquid classes vs categorized

import sys
import os
import csv
import numpy as np
import struct
import pyodbc
import matplotlib
matplotlib.use('Agg')
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import glob

root = tk.Tk()
root.withdraw()

int_stepAspirate = '-533331728'
int_stepDispense = '-533331727'
counter = 1
startBoundary = 15
endBoundary = 30

guardbandOffset = 350

def unpackBinaryData(list_mdb):
	# list_mdb = l_mdb_asp
	l_data_mean = [ 0 for _ in range(10000) ]
	l_data_std = [ [] for _ in range(10000) ]
	l_data_std_temp = [ [] for _ in range(10000) ]
	l_data_std_val = [ 0 for _ in range(10000) ]
	l_data_mean_std_upper = [ 0 for _ in range(10000) ]
	l_data_mean_std_lower = [ 0 for _ in range(10000) ]

	#Unpack the pressure data
	l_data=[]
	for idx, row in enumerate(list_mdb):
		#Unpack bytes
		tadm_byt=row[8]
		num_byt=len(tadm_byt)/2
		fmt=str(int(num_byt))+'h'
		tadm_num=struct.unpack(fmt,tadm_byt)
		
		#Dispense or Aspirate
		if row[2]==-533331728:
			step_type='Aspirate'
		elif row[2]==-533331727:
			step_type='Dispense'
		else:
			step_type='Unknown'
		
		volume = round(row[3])

		#LiquidClassName, StepType, Volume, TimeStamp, StepNumber, CurvePoints, ChannelNumber
		l_data.append([row[1], step_type, volume, row[4], row[7], tadm_num, row[10]])

		l_tadm_num = list(tadm_num)
		for idxTADM, rowTADM in enumerate(l_tadm_num):
			# print('rowTADM: ', rowTADM) # Pressure value at time = idxTADM
			l_data_mean[idxTADM] = l_data_mean[idxTADM] + rowTADM # Adding to list to later calculate mean
			l_data_std_temp = l_data_std[idxTADM] # Temp list to append rowTADM to
			# print('l_data_std_temp pre-append: ', l_data_std_temp)
			l_data_std_temp.append(rowTADM)
			# print('l_data_std_temp post-append: ', l_data_std_temp)
			l_data_std[idxTADM] = l_data_std_temp
			# print('l_data_std[idxTADM]: ', l_data_std[idxTADM])
			# l_data_std[idxTADM] = l_data_std[idxTADM].append(rowTADM)

	#setMDBPath()

	#unpackBinaryData(l_mdb)

	l_data_std_len = []
	for idxLength, rowLength in enumerate(l_data_std):
		l_data_std_len.append(len(rowLength))
	# print('l_data_std_len: ', l_data_std_len)
	for idxMean, rowMean in enumerate(l_data_mean):
		if l_data_std_len[idxMean] != 0:
			l_data_mean[idxMean] = round(l_data_mean[idxMean] / l_data_std_len[idxMean])
			l_data_std_val[idxMean] = round(np.std([l_data_std[idxMean]]))

	l_data_mean = np.trim_zeros(l_data_mean)
	for idxMeanStd, rowMeanStd in enumerate(l_data_mean):
		# l_data_mean_std_upper[idxMeanStd] =  l_data_mean[idxMeanStd] + (l_data_std_val[idxMeanStd] * 2)
		# l_data_mean_std_lower[idxMeanStd] =  l_data_mean[idxMeanStd] - (l_data_std_val[idxMeanStd] * 2)
		l_data_mean_std_upper[idxMeanStd] =  l_data_mean[idxMeanStd] + guardbandOffset
		l_data_mean_std_lower[idxMeanStd] =  l_data_mean[idxMeanStd] - guardbandOffset
	# print('l_data_std: ', l_data_std)
	l_data_mean_std_upper = np.trim_zeros(l_data_mean_std_upper)
	l_data_mean_std_lower = np.trim_zeros(l_data_mean_std_lower)
	l_data_std_val = np.trim_zeros(l_data_std_val)

	# print('l_data_mean: ', l_data_mean)
	# print('l_data_std_val: ', l_data_std_val)
	# print('list of mean + std: ', l_data_mean_std_upper)
	# print('list of mean - std: ', l_data_mean_std_lower)
	# """
	f1=plt.figure(1, figsize=(12,6))
	ax1=plt.subplot(111)
	ax1.set_title('Testing')
	ax1.yaxis.grid(True, linestyle='-', which='major', color='grey', alpha=0.5)
	ax1.xaxis.grid(True, linestyle='-', which='major', color='grey', alpha=0.5)
	ax1.axhline(0, color='black')
	ax1.set_ylabel('Pressure [Pa]',fontsize=14)
	ax1.set_xlabel('Time [ms]',fontsize=14)

	time_upper=np.arange(0,len(l_data_mean_std_upper),1)
	time_lower=np.arange(0,len(l_data_mean_std_lower),1)
	time_mean=np.arange(0,len(l_data_mean),1)
	plt.plot(time_upper,l_data_mean_std_upper,color = 'red')
	plt.plot(time_lower,l_data_mean_std_lower,color = 'red')
	plt.axvline(x = startBoundary)
	plt.axvline(x = len(l_data_mean) - endBoundary)
	plt.scatter(time_mean,l_data_mean,marker='x', color='black',s=16, linewidths=0.1)

	global counter
	handles, labels = ax1.get_legend_handles_labels()
	plt.tight_layout(rect=[0.05, 0.05, 0.95, 0.95])
	plt.suptitle(liq_class_gen, fontweight='bold')
	plt.savefig(liq_class_gen + '_testing.png',dpi=300)
	# plt.savefig(stepType + '_' + str(counter) + '_testing.png',dpi=300)
	#plt.show()
	plt.close(f1)
	# """
	save_path = os.path.abspath('C:/Program Files (x86)/HAMILTON/Methods/TADM')
	os.chdir(save_path)
	# np.savetxt(stepType + '_tadm_mdb.csv',l_data,fmt='%s', delimiter=',')

	writer.writerow([step_type, volume, liq_class_gen])
	writer.writerow(l_data_mean_std_upper)
	writer.writerow(l_data_mean_std_lower)

# mdb_path = ''
l_mdb_asp = []
l_mdb_disp = []

log_path = 'C:/Program Files (x86)/HAMILTON/Logfiles/'

fileOrFolder = '0'

fileOrFolder = input('Guardband for single file(0) or entire directory(1)? (0/1): ')
if fileOrFolder.lower() not in ('0', '1'):
	sys.exit()

# def setMDBPath():
# Prompt for file/folder location
print("Please select .mdb TADM file/folder")
if fileOrFolder == '0':
	mdb_path = filedialog.askopenfilename()
	mdbfiles = []
	mdbfiles.append(mdb_path)
else:
	### For entire directory ###
	mdb_path = filedialog.askdirectory()
	currdir = os.getcwd()
	print('cwd: ', currdir)
	os.chdir(mdb_path)
	mdbfiles = []
	for file in glob.glob('*.mdb'):
			mdbfiles.append(log_path + file)

	print('mdbfiles: ', mdbfiles)
	mdbfiles.sort(key=os.path.getctime)
	print('mdbfiles: ', mdbfiles)
	mdbfiles.reverse()
	print('mdbfiles: ', mdbfiles)
	os.chdir(currdir)

os.system('cls')

if mdb_path == '':
	print('No file/folder selected. Aborting.')
	sys.exit()

for mdb_path in mdbfiles:
	#Read the .mdb file and create a list
	skipFlag = 0
	print('mdb_path: ', mdb_path)
	MDB = mdb_path
	DRV = '{Microsoft Access Driver (*.mdb, *.accdb)}'

	mdbSQLQuery = 'SELECT * FROM TadmCurve WHERE StepType = '

	conn = pyodbc.connect('DRIVER={};DBQ={}'.format(DRV,MDB))
	cursor = conn.cursor()

	mdbSQLQuery_asp = mdbSQLQuery + int_stepAspirate
	cursor.execute(mdbSQLQuery_asp)
	l_mdb_asp = []
	for row in cursor.fetchall():
		l_mdb_asp.append(row)

	mdbSQLQuery_disp = mdbSQLQuery + int_stepDispense
	cursor.execute(mdbSQLQuery_disp)
	l_mdb_disp = []
	for row in cursor.fetchall():
		l_mdb_disp.append(row)

	conn.close()

	liq_class = l_mdb_asp[0][1]
	run_vol = round(l_mdb_asp[0][3])
	# Remove MN and Volume suffices from LC name
	liq_class_gen_index = liq_class.rfind('_')
	liq_class_gen = liq_class[0:liq_class_gen_index]
	liq_class_gen_index = liq_class_gen.rfind('_')
	liq_class_gen = liq_class_gen[0:liq_class_gen_index]

	# Testing unique liquid classes
	liq_class_gen = liq_class
	print('General liquid class currently processing: ',liq_class_gen)

	guardbandFilePath = '//ussd-file/Depts/Ops/MFG/ReagentsFill_Protocols/cLLD/Guardband.csv'
	if os.path.isfile(guardbandFilePath):
		fileGuardbandReader = open(guardbandFilePath, newline='')
		reader = csv.reader(fileGuardbandReader)
		for row in reader:
			if len(row) > 0:
				if row[2] == liq_class_gen:
					if int(row[1]) == int(run_vol):
						print('Guardband for liquid class already exists. Aborting.')
						skipFlag = 1
						break
		fileGuardbandReader.close()

	if skipFlag == 0:
		fileGuardbandWriter = open(guardbandFilePath, 'a', newline='')
		writer = csv.writer(fileGuardbandWriter) # csv Writer to store data to be used as guardband
		counter = counter + 1
		stepType = 'a'
		unpackBinaryData(l_mdb_asp)
		stepType = 'b'
		unpackBinaryData(l_mdb_disp)
		fileGuardbandWriter.close()

	# getMDBData(int_stepAspirate)

"""


#Create a list of all liquid classes used in the run
l_lc=[]
for row in l_data:
	l_lc.append(row[0])
s_lc=set(l_lc)

#Main loop to process data
for item in s_lc:
	liq_class=item
	print('Step currently processing: ',liq_class)
	
	#Pull data from the overall list for steps that use the liquid class under consideration
	lc_data=[]
	for row in l_data:
		if row[0]==liq_class:
			lc_data.append(row)
	
	#Determine each step ID used for commands using the liquid class under consideration
	step_list=[]
	for row in lc_data:
		step_list.append(row[3])
	step_set=set(step_list)
	#step_set=[1543,5129]
	
	#Create list of pressure data for each step ID
	for entry in step_set:
		step=entry
		print('Step currently processing: ',step)

		tadm=[]
		for row in lc_data:
			if row[3]==step:
				tadm.append([row[4],row[5]])
				type_step=row[1]
				time_step=row[2]
				
		s_ts=time_step.strftime('%Y-%m-%d %H:%M:%S')
		
		l_ts=list(s_ts)
		for n, i in enumerate(l_ts):
			if i==':':
				l_ts[n]='-'
		s_ts=''.join(l_ts)
			
		for idx, row in enumerate(tadm):
			press=np.asarray(row[0],dtype='int')
			time=np.arange(0,len(press),1)
			ch_num=str(row[1])

"""

