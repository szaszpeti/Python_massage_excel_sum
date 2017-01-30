import os, openpyxl, csv, time

startTime = time.time()
massageArt = {'Pantha Jama':-5, 'Sport':-5, 'Gesicht':-5, 'Klassische':-5, 'Entspannung':-5, 
			'Bein':-5, 'Fussreflexzonen':-5, 'Heisse Steine':-5, 'Lomi Lomi':-5, 'Rücken-Nacken':-5,
			'Spez. Schulter-Nacken':-5, 'P.J. Gesicht Fuss':-5, 'Dorn Breuss':-5, 'Kristalle':-5, 'Migräne':-5,
			'Peeling':-5, 'Teilkörper Steine':-5}

massageExtern = {'Extern': 0, 'Whymper House':0}			
			
massagePrice = {'Pantha Jama':190, 'Sport':90, 'Gesicht':60, 'Klassische':110, 'Entspannung':160, 
			'Bein':60, 'Fussreflexzonen':60, 'Heisse Steine':190, 'Lomi Lomi':110, 'Rücken-Nacken':60,
			'Spez. Schulter-Nacken':60, 'P.J. Gesicht Fuss':130, 'Dorn Breuss':110, 'Kristalle':110, 'Migräne':90,
			'Peeling':100, 'Teilkörper Steine':100}
			

massageTime = {'08:00-12:00': 0, '12:00-15:00':0, '15:00-20:00':0}
			
for excelFile in os.listdir('.'):
	if not excelFile.endswith('.xlsx'):
		continue
	wb = openpyxl.load_workbook(excelFile)
	for sheetName in wb.get_sheet_names():
		sheet = wb.get_sheet_by_name(sheetName)
		for i in range(1, 55):
			for x in range(1, 25):
				c = sheet.cell(row=i, column=x).value
				time = sheet.cell(row=i, column=1).value
				if c in massageArt:
					massageArt[c] += 1
					if str(time) >= '15:00:00':
						massageTime['15:00-20:00'] += 1 
					elif '12:00:00' <= str(time) < '15:00:00':
						massageTime['12:00-15:00'] += 1
					else:
						massageTime['08:00-12:00'] += 1
				elif str(c).startswith('ext'):
					massageExtern['Extern'] += 1
					print(c)
				elif str(c).startswith('wh') or str(c).startswith('wy'):
					massageExtern['Whymper House'] += 1
					print (c)
					
outputFile = open('my_massages.txt', 'w')
outputFile.write('MASSAGES BOOKED'.center(40, '-') + '\n\n')
for k, v in massageArt.items():
	outputFile.write((' ' * 10) + k.ljust(22, '.') + ':' + ('%d' %v).rjust(5) + '\n')

outputFile.write('\n\n')

outputFile.write(('TOTAL MASSAGES'.ljust(32, '.')) + ':' + ('%d' % sum(list(massageArt.values()))).rjust(5))  
outputFile.write('\n\n')

price = 0
for k, v in massageArt.items():
	price += massagePrice[k] * v


outputFile. write('TOTAL INCOME'.ljust(32, '.') + ':' + ('%d' %price).rjust(5) + '\n\n')

outputFile.write('COUSTAMERS FROM OTHER HOTEL'.center(40, '-') + '\n\n')
for k, v in massageExtern.items():
	outputFile.write((' ' * 10) + k.ljust(22, '.') + ':' + ('%d' %v).rjust(5) + '\n')
	
outputFile.write('\n')

outputFile.write('MASSAGE BOOKING TIME'.center(40, '-') + '\n\n')
outputFile.write((' ' * 10) + '08:00-12:00'.ljust(22, '.') + ':' + ('%d' % (massageTime['08:00-12:00'])).rjust(5) + '\n') 
outputFile.write((' ' * 10) + '12:00-15:00'.ljust(22, '.') + ':' + ('%d' % (massageTime['12:00-15:00'])).rjust(5) + '\n') 
outputFile.write((' ' * 10) + '15:00-20:00'.ljust(22, '.') + ':' + ('%d' % (massageTime['15:00-20:00'])).rjust(5) + '\n') 
				
outputFile.close()

print('Done')

