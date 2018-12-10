import sys, os
print(sys.version)

import glob, datetime
import xlrd, csv, pandas, numpy

enduse_dict = dict(zip(["Heating","Cooling","Vent","WaterHeat","Cooking","Refrig",
		"ExtLight","IntLight","OfficeEquip","Misc","Process","Motors","AirComp"],
		["Heating","Cooling","Ventilation","Water Heating","Cooking","Refrigeration",
		"Exterior Lighting","Interior Lighting","Office Equipment","Miscellaneous","Process","Motors","Air Compressors"]
		))
Theat = 55.0
Tcool = 65.0

#
# Building data
#
def load_xls(file,sheets='all'):
	if not os.path.isfile(file):
		raise Exception("%s: open failed" % file)
	xls = xlrd.open_workbook(file)
	data = {}
	for sheet in xls.sheets():
		if sheets != 'all' and not sheet.name in sheets :
			continue;
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
	floorarea_list = []
	for row in range(5,18) :
		rowdata = summary[row]
		assert(rowdata[1]==enduse_dict.values()[row-5])
		floorarea_list.append(rowdata[2])	

	# enduse
	enduse_data = data["expMnthDT"]
	assert(enduse_data[0][0]=="SegID")
	assert(enduse_data[0][1]=="Mth")
	assert(enduse_data[0][2]=="Dy")
	assert(enduse_data[0][3]=="Hr")
	enduse_loads = []
	active_enduses = []
	header = ["Month","Daytype","Hour"]
	for col in range(4,17) :
		assert(enduse_data[0][col]==enduse_dict.keys()[col-4])
		if floorarea_list[col-4] > 0 :
			active_enduses.append(col)
			header.append(enduse_dict.values()[col-4].replace(" ","_"))
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
				data = load_xls(file="xls/"+xls,sheets=['ctrlSEGINFO','Summary','expMnthDT'])
				convert_to_loadshape(data=data,to_csv=csv)
	print("enduse/*.csv up to date")

#
# Weather data
#
def get_weather(station) :
	data = pandas.read_csv('weather/lcd.csv', dtype={'HOURLYDRYBULBTEMPF':str}, usecols=['STATION','DATE','HOURLYDRYBULBTEMPF'])
	ndx = find(data['STATION']==station)
	result = data.ix[ndx,['DATE','HOURLYDRYBULBTEMPF']].dropna()
	if len(result) == 0 :
		print('get_weather(station=%s): no data found in LCD repository' % station)
	return result

def load_weather(station) :
	data = pandas.read_csv('weather/%s.csv'%station)
	return data

def find(data) :
	result = []
	for key, item in data.items() :
		if not type(item) is bool :
			exception("test result of item %d must be boolean" % key)
		if item == True :
			result.append(key)
	return result

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
	fcz = None
	result = {}
	for xls in os.listdir("xls/") :
		if xls.endswith(".xls") :
			csv = "sensitivity/%s.csv" % os.path.splitext(xls)[0]
			if not os.path.isfile(csv) :
				data = load_xls(file="xls/"+xls,sheets=['ctrlSEGINFO','expEndUse8760'])
				if type(fcz) is None or fcz != xls.split('_')[0] :
					fcz = xls.split('_')[0]
					weather = load_weather(fcz)
				#if fcz[0] in result.keys() :
				#	result[fcz[0]] = {}
				#result[fcz[0]][fcz[1]] = 
				get_sensitivity(data,weather)
	print("sensitivity/*.csv up to date")


def get_sensitivity(data,weather) :
	segment = data['ctrlSEGINFO']
	assert(segment[0][0]=='SegID')
	segment_name = segment[0][1]
	year = int(segment[2][1])
	day0 = datetime.date(year,1,1)
	enduses = data['expEndUse8760']
	A = {}
	y = {}
	found = {}
	senscols = {}
	#result = {}
	remap = {"OffEquip":"OfficeEquip", "Cook":"Cooking", "Cool":"Cooling", "Heat":"Heating", "HotWater":"WaterHeat"} # fix enduse inconsitencies
	samap = { # identify segments over which sensitivity is to be computed
		"Heating" : {'Theat': 55.0},
		"Cooling" : {'Tcool': 65.0},
		"Vent" : {'Theat': 55.0,'Tcool': 65.0},
		"WaterHeat" : {},
		"Cooking" : {},
		"Refrig" : {},
		"ExtLight" : {},
		"IntLight" : {},
		"OfficeEquip": {},
		"Misc" : {},
		"Process" : {},
		"Motors" : {},
		"AirComp" : {},
		}
	enduse_keys = enduse_dict.keys()
	for enduse_name in enduse_keys :
		A[enduse_name] = numpy.zeros((8760,48 + len(samap[enduse_name])))
		y[enduse_name] = numpy.zeros((8760,1))
		found[enduse_name] = []
		senscols[enduse_name] = {}
		if 'Theat' in samap[enduse_name] :
			senscols[enduse_name]['Theat'] = 48
		if 'Tcool' in samap[enduse_name] :
			senscols[enduse_name]['Tcool'] = 48+len(senscols[enduse_name].keys())
	print("processing %s (%d rows, %d cols)..." % (segment_name,len(enduses),len(enduses[0])))

	for row in enduses[1:] :
		enduse_name = row[1]
		if enduse_name in remap.keys() :
			enduse_name = remap[enduse_name]
		fuel = row[2]
		month = int(row[3])
		day = int(row[4])
		load = row[5:29]
		heat_col = None
		if 'Theat' in senscols[enduse_name] :
			heat_col = senscols[enduse_name]['Theat']
		cool_col = None
		if 'Tcool' in senscols[enduse_name] :
			cool_col = senscols[enduse_name]['Tcool']
		if fuel == 'Elec' :
			if enduse_name not in enduse_keys :
				raise Exception("%s.%s(%d,%d): enduse '%s' not found in enduse_dict" % (segment_name,enduse_name,month,day,enduse_name))
			else:
				date = datetime.date(year,month,day)
				doy = date.toordinal() - day0.toordinal()
				if date.weekday() < 5 :
					hour0 = 0
				else:
					hour0 = 24
				for hour in range(0,24) :
					r = doy*24 + hour
					if load[hour] > 0.0 :
						T = weather["drybulb"][r]
						c = hour0 + hour
						#print("%s.%s(%2d,%2d): row %3d, col %2d, hour %2d, T %.1f, y %.1f" % (segment_name,enduse_name,month,day,r,c,hour,T,load[hour]))
						if c > 0 :
							A[enduse_name][r,0] = 1.0
						A[enduse_name][r,c] = 1.0
						if T < Theat and heat_col != None :
							A[enduse_name][r,heat_col] = T - Theat
							#print("%s.%s(%2d,%2d): row %3d, col %2d, hour %2d, Theat %.1f, y %.1f" % (segment_name,enduse_name,month,day,r,c,hour,T,load[hour]))
						elif T > Tcool and cool_col != None:
							A[enduse_name][r,cool_col] = Tcool - T
							#print("%s.%s(%2d,%2d): row %3d, col %2d, hour %2d, Tcool %.1f, y %.1f" % (segment_name,enduse_name,month,day,r,c,hour,T,load[hour]))
						#else :
							#print("%s.%s(%2d,%2d): row %3d, col %2d, hour %2d, Tnone %.1f, y %.1f" % (segment_name,enduse_name,month,day,r,c,hour,T,load[hour]))
						y[enduse_name][r] = load[hour]
						found[enduse_name].append(r)
	for enduse_name in enduse_keys :
		if len(found[enduse_name]) > 0 :
			heat_col = None
			if 'Theat' in senscols[enduse_name] :
				heat_col = senscols[enduse_name]['Theat']
			cool_col = None
			if 'Tcool' in senscols[enduse_name] :
				cool_col = senscols[enduse_name]['Tcool']
			#for row in range(len(found[enduse_name])) :
				#print(AA[row,:],yy[row])	
			try :
				cols = []
				AA = A[enduse_name][found[enduse_name],:]
				for c in range(0,48) :
					if numpy.count_nonzero(AA[:,c]) > 0 :
						cols.append(c)
				if 'Theat' in senscols[enduse_name] :
					cols.append(senscols[enduse_name]['Theat'])
				if 'Tcool' in senscols[enduse_name] :
					cols.append(senscols[enduse_name]['Tcool'])
				#print('Columns: %s' % cols)
				AA = AA[:,cols]
				yy = y[enduse_name][found[enduse_name]]
				#print('Enduse %s, %d samples, A is %dx%d, y is %dx%d'
				#	% (enduse_name,len(found[enduse_name]),len(AA),len(cols), len(yy), len(cols)))
				At = AA.transpose()
				AtA = numpy.dot(At,AA)
				AtAi = numpy.linalg.inv(AtA)
				AtAiAt = numpy.dot(AtAi,At)
				x = numpy.dot(AtAiAt,yy)

				# output
				xx = numpy.zeros(50)
				for h,xh in dict(zip(cols,x)).items() :
					#print(h,xh[0],xx)
					xx[h] = xh[0]
				xx[1:47] += x[0]
				rs = pandas.DataFrame()
				rs['WeekdayLoad'] = xx[0:24]
				rs['WeekendLoad'] = xx[24:48]
				if 'Theat' in senscols[enduse_name] :
					rs['HeatingSensitivity'] = xx[senscols[enduse_name]['Theat']]
				else :
					rs['HeatingSensitivity'] = 0.0
				if 'Tcool' in senscols[enduse_name] :
					rs['CoolingSensitivity'] = xx[senscols[enduse_name]['Tcool']]
				else :
					rs['CoolingSensitivity'] = 0.0
				rs.index.name = 'HourOfDay'
				path = 'loadshape/%s' %  segment_name.replace('_','/')
				if not os.path.exists(path) :
					os.makedirs(path)
				if os.path.isdir(path) :
					rs.to_csv('%s/%s.csv' % (path,enduse_name))
				#print(rs)
				#result[enduse_name] = rs
			except :
				if not os.path.exists('dump') :
					os.mkdir('dump')
				if os.path.isdir('dump') :
					print('Enduse %s, %d samples, A is %dx%d, y is %dx%d: sensitivity analysis failed -- dumping data to tmp/%s_%s_[Ay].csv'
						% (enduse_name,len(found[enduse_name]),len(AA),len(AA[0]), len(yy), len(yy[0]), segment_name, enduse_name))
					pandas.DataFrame(AA).to_csv('dump/%s_%s_A.csv'%(segment_name,enduse_name))
					pandas.DataFrame(yy).to_csv('dump/%s_%s_y.csv'%(segment_name,enduse_name))
				raise
	#return result
#
# MAIN
#
def main():
	update_weather()
	update_csv()
	update_sensitivity()

if __name__ == '__main__':
	main()