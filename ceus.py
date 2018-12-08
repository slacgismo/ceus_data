import glob, os, datetime
import xlrd, csv, pandas
import numpy

#
# Building data
#
def load_xls(file):
	if not os.path.isfile(file):
		raise Exception("%s: open failed" % file)
	xls = xlrd.open_workbook(file)
	data = {}
	for sheet in xls.sheets():
		data[sheet.name] = []
		for row in range(sheet.nrows):
			datarow = []
			for col in range(sheet.ncols):
				cell = sheet.cell(row,col)
				if cell.ctype == xlrd.XL_CELL_NUMBER:
					value = float(cell.value)
				elif cell.ctype == xlrd.XL_CELL_TEXT:
					value = str(cell.value)
				elif cell.ctype == xlrd.XL_CELL_DATE:
					value = datetime.datetime(1900,1,1) + datetime.timedelta(cell.value)
				elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
					value = bool(cell.value)
				elif cell.ctype == xlrd.XL_CELL_BLANK:
					value = []
				elif cell.ctype == xlrd.XL_CELL_EMPTY:
					value = None
				elif cell.ctype == xlrd.XL_CELL_ERROR:
					value = int(cell.value) # represents internal Excel error code
				else:
					raise Exception("unable to read cell type {} value {}".format(cell.type,cell.value))
				datarow.append(value)
			data[sheet.name].append(datarow)
	return data

def convert_to_loadshape(data,to_csv) :
	print("updating from %s" % to_csv)

	# segment information
	segment = data["ctrlSEGINFO"]
	assert(segment[1][0]=="Description")
	description = segment[1][1]
	assert(segment[2][0]== "AnalysisYear")
	analysis_year = int(segment[2][1])
	#print("  segment........... %s" % description)
	#print("  year.............. %s" % analysis_year)

	# summary
	summary = data["Summary"]
	enduse_label = ["Heating","Cooling","Ventilation","Water Heating","Cooking","Refrigeration",
		"Exterior Lighting","Interior Lighting","Office Equipment","Miscellaneous","Process","Motors","Air Compressors"]
	floorarea_list = []
	for row in range(5,18) :
		rowdata = summary[row]
		assert(rowdata[1]==enduse_label[row-5])
		floorarea_list.append(rowdata[2])	
	#print("  enduse labels..... %s" % enduse_label)
	#print("  floorarea......... %s" % floorarea_list)

	# enduse
	enduse_data = data["expMnthDT"]
	assert(enduse_data[0][0]=="SegID")
	assert(enduse_data[0][1]=="Mth")
	assert(enduse_data[0][2]=="Dy")
	assert(enduse_data[0][3]=="Hr")
	enduse_heading = ["Heating","Cooling","Vent","WaterHeat","Cooking","Refrig",
		"ExtLight","IntLight","OfficeEquip","Misc","Process","Motors","AirComp"]
	enduse_loads = []
	active_enduses = []
	header = ["Month","Daytype","Hour"]
	for col in range(4,17) :
		assert(enduse_data[0][col]==enduse_heading[col-4])
		if floorarea_list[col-4] > 0 :
			active_enduses.append(col)
			header.append(enduse_label[col-4].replace(" ","_"))
	daytype_name = ["WEEKDAY","SATURDAY","SUNDAY","HOLIDAY"]
	for row in range(1,2017) :
		rowdata = enduse_data[row]
		if rowdata[2] not in (10,11,12,13) :
			continue;
		csvdata = [int(rowdata[1]),daytype_name[int(rowdata[2])-10],int(rowdata[3]-1)]
		for col in active_enduses :
			csvdata.append("%.4f"%(rowdata[col]/floorarea_list[col-4]))
		enduse_loads.append(csvdata)
	with open(to_csv,"w") as csvfile:
		writer = csv.writer(csvfile,delimiter=',',quoting=csv.QUOTE_MINIMAL,quotechar='"')
		writer.writerow(header)
		for row in enduse_loads:
			writer.writerow(row)
	return None

def update_csv() :
	for xls in os.listdir("xls/") :
		if xls.endswith(".xls"):
			csv = "enduse/%s.csv" % os.path.splitext(xls)[0]
			data = None
			if not os.path.isfile(csv) :
				data = load_xls("xls/"+xls)
				convert_to_loadshape(data=data,to_csv=csv)
	print("enduse/*.csv up to date")

def find(data) :
	result = []
	for key, item in data.items() :
		if not type(item) is bool :
			exception("test result of item %d must be boolean" % key)
		if item == True :
			result.append(key)
	return result

#
# Weather data
#
def get_weather(station) :
	data = pandas.read_csv('weather/lcd.csv', dtype={'HOURLYDRYBULBTEMPF':str}, usecols=['STATION','DATE','HOURLYDRYBULBTEMPF'])
	ndx = find(data['STATION']==station)
	return data.ix[ndx,['DATE','HOURLYDRYBULBTEMPF']].dropna()

def update_weather() :
	zones = pandas.read_csv('weather_zones.csv')
	for ndx, zone in zones.iterrows() :
		name = zone['AREA']
		file = 'weather/%s.csv' % name
		if os.path.isfile(file) :
			continue
		try :
			station = zone['STATION']
			data =get_weather(station)
			if len(data) == 0 :
				print("no data for %s (%s)" % (name,station))
				continue
			else :
				print("processing %s..." % name)
			date = data['DATE']
			first = data.iterrows().next()[0]
			year = datetime.datetime.strptime(date[first],'%Y-%m-%d %H:%M').year
			starttime = datetime.datetime(year,1,1,0,0,0)
			stoptime = datetime.datetime(year+1,1,1,0,0,0)
			dt = list(map(lambda x: (datetime.datetime.strptime(x,'%Y-%m-%d %H:%M')-starttime).total_seconds()/3600.0,date))
			hour = numpy.arange(0,8760,1)
			timestamp = list(map(lambda x: (starttime+datetime.timedelta(hours=x)).strftime('%Y-%m-%d %H:%M:%S'),hour))
			Tdb = numpy.around(numpy.interp(numpy.arange(0,8760,1),dt,list(map(lambda x: numpy.float64(x),data['HOURLYDRYBULBTEMPF']))),1)
			with open(file,'w') as csvfile:
				writer = csv.writer(csvfile)
				writer.writerow(['hour','drybulb'])
				for row in sorted(dict(zip(timestamp,Tdb)).items()):
					writer.writerow(row)
		except :
			print(Exception("unable to process zone '%s'" % name))
			raise 
	print("weather/*.csv up to date")

#
# Sensititivy Analysis
#
def update_sensitivity() :
	print("sensitivity/*.csv up to date")

#
# MAIN
#
def main():
	update_weather()
	update_csv()
	update_sensitivity()

if __name__ == '__main__':
	main()