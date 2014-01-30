import sys
import datetime
import gspread
import csv
from locale import *

territorydict = { "620  GAC" : "Georgia", "613  Ft Lauderdale" : "Florida", 
				"612  Tampa/Ft Myers" : "Florida", "610  Orlando" : "Florida",
				"611  JAX/Panhandle" : "Florida" } 

gusername = "paland@gmail.com"
gpassword = "YTkriOZ5YWNa2H"			
gsheetname = sys.argv[1]

docompare = False
setlocale(LC_NUMERIC, '')
if len(sys.argv) > 2:
	docompare = True
	compfile = sys.argv[2]
	infile = open(compfile, "rb")
	reader = csv.DictReader(infile)
	rownum = 0
	actdict = {}
	for row in reader:
		actdict[row["Resource"]] = atof(row["Project"])

gc = gspread.login(gusername, gpassword)
hours_spreadsheet = gc.open(gsheetname)
ws = hours_spreadsheet.get_worksheet(0)
allitems = ws.get_all_records()
region_dict = {}
for item in allitems:
	try:
		hours = float(item["Hours"])
	except ValueError:
		print "Blank entry found for %s" % item["Changepoint Name"]
		hours = 0.0
		item["Notes"] = "Not reported by supervisor [auto]"
	if item["Region"] not in territorydict.keys():
		item["Region"] = "Unknown Region"
	region = territorydict[item["Region"]]
	engtech = item["Resource Tech"]
	if region not in region_dict.keys():
		region_dict[region] = { "Headcount" : 0, "Hours Count" : 0, 
								"techs" : {}, "underutil" : {}, "badforecast": {},
								"Actual Hours": 0 }
	region_dict[region]["Headcount"] += 1
	region_dict[region]["Hours Count"] += hours
	if docompare:
		try:
			acthours = float(actdict[item["Changepoint Name"]])
		except ValueError:
			print "No actuals found for %s" % item["Changepoint Name"]
			acthours += 0
		region_dict[region]["Actual Hours"] += acthours
	if engtech not in region_dict[region]["techs"].keys():
		region_dict[region]["techs"][engtech] = { "Headcount" : 0, 
												"Hours Count" : 0, "Actual Hours": 0}
	region_dict[region]["techs"][engtech]["Headcount"] += 1
	region_dict[region]["techs"][engtech]["Hours Count"] += hours
	if docompare:
		region_dict[region]["techs"][engtech]["Actual Hours"] += acthours
		if hours - acthours > 8:
			region_dict[region]["badforecast"][item["Changepoint Name"]] = [hours, acthours]
	if hours < 8:
		region_dict[region]["underutil"][item["Changepoint Name"]] = { "Hours" : hours, 
		"Notes" : item["Notes"] }

for key, value in region_dict.iteritems():
	print "Engineer Forecast for %s" % key
	print "\tUnion 900 Engineers: %s" % value["Headcount"]
	print "\tForecast Hours: %s" % value["Hours Count"]
	if docompare:
		print "\tActual Hours: %.2f" % value["Actual Hours"]
	print "\tForecast Utilization: %.2f%%" % (float(value["Hours Count"]) / 
		(float(value["Headcount"]) * 40.0 ) * 100)
	if docompare:
		print "\tActual Utilization: %.2f%%" % (float(value["Actual Hours"]) / 
				(float(value["Headcount"]) * 40.0 ) * 100)
	print "\t**Technology Team Breakdown**"
	for tkey, tvalue in value["techs"].iteritems():
		print "\t\tTechnology: %s" % tkey
		print "\t\t\tUnion 900 Engineers: %s" % tvalue["Headcount"]
		print "\t\t\tForecast Hours: %s" % tvalue["Hours Count"]
		if docompare:
			print "\t\t\tActual Hours: %s" % tvalue["Actual Hours"]
		print "\t\t\tForecast Utilization: %.2f%%" % (float(tvalue["Hours Count"]) / 
			(float(tvalue["Headcount"]) * 40.0 ) * 100)
		if docompare:
			print "\t\t\tActual Utilization: %.2f%%" % (float(tvalue["Actual Hours"]) / 
					(float(tvalue["Headcount"]) * 40.0 ) * 100)
	print "\t**Under utilized engineers**"
	for tkey, tvalue in value["underutil"].iteritems():
		print "\t\t%s %.2f hours (%s)" % (tkey, tvalue["Hours"], tvalue["Notes"])
	print ""
	if docompare:
		print "\t**Incorrect Forecast Engineers (neg only)"
		for tkey, tvalue in value["badforecast"].iteritems():
			print "\t\t%s %.2f Forecast / %.2f Actual" % (tkey, tvalue[0], tvalue[1])