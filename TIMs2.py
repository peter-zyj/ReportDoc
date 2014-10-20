#######################################################################
#
# An example of creating Excel Bar charts with Python and XlsxWriter.
#
# Copyright 2013-2014, John McNamara, jmcnamara@cpan.org #
import xlsxwriter
import xlrd
import os,sys
import re,time,datetime
wb = xlrd.open_workbook('tim.xls')
sh = wb.sheet_by_name(u'Summary Report')
whole_status = {}

# for colnum in range(sh.ncols):
# 	if sh.cell(0,colnum).value == 'Folders':
# 		whole_status['US_number'] = sh.cell(1,colnum).value

# 	if sh.cell(0,colnum).value == 'Executed':
# 		whole_status['Exec_Perc'] = sh.cell(1,colnum).value

# 	if sh.cell(0,colnum).value == 'Quality':
# 		whole_status['Quality'] = sh.cell(1,colnum).value

# 	if sh.cell(0,colnum).value == 'Defects':
# 		whole_status['Defects'] = sh.cell(1,colnum).value

for colnum in range(sh.ncols):
	if sh.cell(0,colnum).value == 'Folders':
		whole_status['US_number'] = sh.cell(1,colnum).value

	# if sh.cell(0,colnum).value == 'Executed':
	# 	whole_status['Executed'] = sh.cell(1,colnum).value

	if sh.cell(0,colnum).value == 'Pending':
		whole_status['Pending'] = sh.cell(1,colnum).value

	if sh.cell(0,colnum).value == 'Passed':
		whole_status['Passed'] = sh.cell(1,colnum).value

	if sh.cell(0,colnum).value == 'Pass w/ X':
		whole_status['PassX'] = sh.cell(1,colnum).value

	if sh.cell(0,colnum).value == 'Failed':
		whole_status['Failed'] = sh.cell(1,colnum).value

	if sh.cell(0,colnum).value == 'Dropped':
		whole_status['Dropped'] = sh.cell(1,colnum).value

	if sh.cell(0,colnum).value == 'Blocked':
		whole_status['Blocked'] = sh.cell(1,colnum).value

# print ("CZ:debug::",whole_status)

US_status = {}
index = 0

for rownum in range(sh.nrows):
	if re.compile(r'US\d+').findall(sh.cell(rownum,0).value)!= []:
		US_status[index]= {}

		for colnum in range(sh.ncols):
			if sh.cell(0,colnum).value == 'Folders':
				US_status[index]['US_number'] = sh.cell(rownum,colnum).value
				
			if sh.cell(0,colnum).value == 'Executed':
				US_status[index]['Exec_Perc'] = sh.cell(rownum,colnum).value

			if sh.cell(0,colnum).value == 'Quality':
				US_status[index]['Quality'] = sh.cell(rownum,colnum).value

			if sh.cell(0,colnum).value == 'Defects':
				US_status[index]['Defects'] = sh.cell(rownum,colnum).value
		index = index + 1
# print ("CZ:debug::",US_status)

whole_list = [float(re.compile(r'(?<=\().*?(?=\%\))').findall(str(whole_status['Pending']))[0]),\
			  float(re.compile(r'(?<=\().*?(?=\%\))').findall(str(whole_status['Passed']))[0]),\
			  float(re.compile(r'(?<=\().*?(?=\%\))').findall(str(whole_status['PassX']))[0]),\
			  float(re.compile(r'(?<=\().*?(?=\%\))').findall(str(whole_status['Failed']))[0]),\
			  float(re.compile(r'(?<=\().*?(?=\%\))').findall(str(whole_status['Dropped']))[0]),\
			  float(re.compile(r'(?<=\().*?(?=\%\))').findall(str(whole_status['Blocked']))[0]),\
			 ]



print ("CZ:debug::",whole_list)

US_list = []
Exec_list = []
Quality_list = []
Defects_list = []

for i in US_status:
	US_list.append(US_status[i]['US_number'].strip())
	try:
		v = float(re.compile(r'(?<=\().*?(?=\%\))').findall(str(US_status[i]['Exec_Perc']))[0])
	except:
		v = ''
	Exec_list.append(v)
	Quality_list.append(US_status[i]['Quality']*100)
	Defects_list.append(US_status[i]['Defects'])


print ("CZ:debug::",US_list)
print ("CZ:debug::",Exec_list)
print ("CZ:debug::",Quality_list)
print ("CZ:debug::",Defects_list)



workbook = xlsxwriter.Workbook('TIMs.xlsx')
cc = time.localtime()[0:3]
sheetName = str(datetime.date(cc[0], cc[1], cc[2]).isocalendar()[1])+'th_week_quality'
sheetName2 = str(datetime.date(cc[0], cc[1], cc[2]).isocalendar()[1])+'th_week_percentage'
worksheet = workbook.add_worksheet(sheetName)
worksheet2 = workbook.add_worksheet(sheetName2)
bold = workbook.add_format({'bold': 1})

# Add the worksheet data that the charts will refer to.
headings = ['US_number', 'Exec_Perc', 'Quality','Defects']
headings2 = ['Category', 'Values']

data2 = [
		    ['Pending','Passed','PassX','Failed','Dropped','Blocked'],
		    whole_list,

	    ]

data = [
		    US_list,
		    Exec_list,
			Quality_list,
			Defects_list,
	    ]
worksheet.write_row('A1', headings, bold)
worksheet.write_column('A2', data[0])
worksheet.write_column('B2', data[1])
worksheet.write_column('C2', data[2])
worksheet.write_column('D2', data[3])

worksheet2.write_row('A1', headings2, bold)
worksheet2.write_column('A2', data2[0])
worksheet2.write_column('B2', data2[1])
# #######################################################################

#
# Create a new bar chart.
#


chart1 = workbook.add_chart({'type': 'bar'})
chart2 = workbook.add_chart({'type': 'pie'})


# Configure the 1st series.
chart2.add_series({
    'name':       'Pie execution Percentage',
    'categories': [sheetName2, 1, 0, len(whole_list), 0],
    'values':     [sheetName2, 1, 1, len(whole_list), 1],
     'points': [
        {'fill': {'color': 'gray'}},
        {'fill': {'color': 'green'}},
        {'fill': {'color': 'orange'}},
        {'fill': {'color': 'red'}},
        {'fill': {'color': 'purple'}},
        {'fill': {'color': 'black'}},
		],
})

# Configure the 1st series.
# chart1.add_series({
# 	    	'name':       '=%s!$B$1' % (sheetName),
# 	        'categories': '=%s!$A$2:$A$%s' % (sheetName,len(US_list)),
# 		    'values':     '=%s!$B$2:$B$%s' % (sheetName,len(US_list)),
# 		    'fill':		  {'color': 'yellow'},

# 		    })

chart1.add_series({
	    	'name':       [sheetName, 0, 1],
	        'categories': [sheetName, 1, 0, len(US_list), 0],
		    'values':     [sheetName, 1, 1, len(US_list), 1],
		    'fill':		  {'color': 'yellow'},

		    })
# Configure a 2nd series. Note use of alternative syntax to define rang
chart1.add_series({
	    	'name':       [sheetName, 0, 2],
	        'categories': [sheetName, 1, 0, len(US_list), 0],
		    'values':     [sheetName, 1, 2, len(US_list), 2],
		    'fill':		  {'color': 'green'},
		    })
# Configure a 3rd series. Note use of alternative syntax to define rang
chart1.add_series({
	    	'name':       [sheetName, 0, 3],
	        'categories': [sheetName, 1, 0, len(US_list), 0],
		    'values':     [sheetName, 1, 3, len(US_list), 3],
		    'line':       {'color': 'red'},
		    'fill':		  {'none': True},
		    'y2_axis': 1,
		    })





# Add a chart title and some axis labels.
chart1.set_title ({'name': 'Quality Track'})
chart1.set_x_axis({'name': 'Cover Percentage %'})
chart1.set_y2_axis({'name': 'Defects Number'})
chart1.set_y_axis({'name': 'User Story'})

chart2.set_title ({'name': 'Whole EXEC %'})

# chart1.set_plotarea({
#     'layout': {
# 	'x': 0.2,
# 	'y': 0.26,
#     'width':  0.73,
#     'height': 0.57,
# } })


chart2.set_chartarea({
    'border': {'width': 20},
})

chart2.set_legend({
    'layout': {
		'x': 0.80,
		'y': 0.37,
        'width':  0.12,
        'height': 0.25,
} 
})

# Set an Excel chart style.
# chart1.set_size({'y_scale': 2})
# chart1.set_size({'x_scale': 2})
chart1.set_style(11)
# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart('G2', chart1, {'x_offset': 250, 'y_offset': 100, 'y_scale': 2, 'x_scale': 3})
worksheet2.insert_chart('C2', chart2, {'x_offset': 250, 'y_offset': 100, 'y_scale': 2, 'x_scale': 2})
workbook.close()

