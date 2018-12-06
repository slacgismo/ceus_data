import glob, os, datetime, xlrd, csv

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

def convert(from_xls,to_csv) :
	print("%s: updating from %s" % (to_csv,from_xls))
	data = load_xls(from_xls)

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
	enduse_label = ["Heating","Cooling","Ventilation","Water Heating","Cooking","Refrigeration","Exterior Lighting","Interior Lighting","Office Equipment","Miscellaneous","Process","Motors","Air Compressors"]
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
	enduse_heading = ["Heating","Cooling","Vent","WaterHeat","Cooking","Refrig","ExtLight","IntLight","OfficeEquip","Misc","Process","Motors","AirComp"]
	enduse_loads = []
	active_enduses = []
	header = ["Month","Weekend","Hour"]
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

def main():
	for xls in os.listdir("xls/") :
		if xls.endswith(".xls"):
			csv = "csv/%s.csv" % os.path.splitext(xls)[0]
			if not os.path.isfile(csv) :
				xls = "xls/"+xls
				convert(from_xls=xls,to_csv=csv)
	print("Files are up to date")

if __name__ == '__main__':
	main()