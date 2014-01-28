import sys
import datetime
import gspread

territorydict = { "620  GAC" : "Georgia", "613  Ft Lauderdale" : "Florida", 
				"612  Tampa/Ft Myers" : "Florida", "610  Orlando" : "Florida",
				"611  JAX/Panhandle" : "Florida" } 

			
gc = gspread.login(sys.argv[1], sys.argv[2])
hours_spreadsheet = gc.open("Forecast Week of 2014-01-26")
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
		region_dict[region] = { "Headcount" : 0, "Hours Count" : 0, "techs" : {}, "underutil" : {} }
	region_dict[region]["Headcount"] += 1
	region_dict[region]["Hours Count"] += hours
	if engtech not in region_dict[region]["techs"].keys():
		region_dict[region]["techs"][engtech] = { "Headcount" : 0, "Hours Count" : 0}
	region_dict[region]["techs"][engtech]["Headcount"] += 1
	region_dict[region]["techs"][engtech]["Hours Count"] += hours
	if hours < 8:
		region_dict[region]["underutil"][item["Changepoint Name"]] = { "Hours" : hours, 
		"Notes" : item["Notes"] }

for key, value in region_dict.iteritems():
	print "Engineer Forecast for %s" % key
	print "\tUnion 900 Engineers: %s" % value["Headcount"]
	print "\tForecast Hours: %s" % value["Hours Count"]
	print "\tEngineering Utilization: %.2f%%" % (float(value["Hours Count"]) / 
		(float(value["Headcount"]) * 40.0 ) * 100)
	print "\t**Technology Team Breakdown**"
	for tkey, tvalue in value["techs"].iteritems():
		print "\t\tTechnology: %s" % tkey
		print "\t\t\tUnion 900 Engineers: %s" % tvalue["Headcount"]
		print "\t\t\tForecast Hours: %s" % tvalue["Hours Count"]
		print "\t\t\tEngineering Utilization: %.2f%%" % (float(tvalue["Hours Count"]) / 
			(float(tvalue["Headcount"]) * 40.0 ) * 100)
	print "\t**Under utilized engineers**"
	for tkey, tvalue in value["underutil"].iteritems():
		print "\t\t%s %.2f hours (%s)" % (tkey, tvalue["Hours"], tvalue["Notes"])
	print ""