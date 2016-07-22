#coding=utf-8
import os
import sys
import codecs 
import xlrd

def script_path():
	import inspect, os
	caller_file = inspect.stack()[1][1]         # caller's filename
	path = os.path.abspath(os.path.dirname(caller_file))# path	
	return path
	
def ChangePath(path):
	return path.decode('gbk').encode('gbk')

def read_val(val):
	if isinstance(val, float):
		return str(int(val))
	if isinstance(val, int):
		return str(val)
	return val.rstrip()
	
def ReadTemplate(file):
	path = os.path.join(local.decode('gbk').encode('utf8')+"/tmp/", file)
	f = open(path)
	text = ""
	for i in f:
		text += i
	return text
	
def createAndOpenFile(name):
	path = os.path.join(local.decode('gbk').encode('utf8')+"/src/", name)
	if os.path.exists(path):
		os.remove(path)
	return codecs.open(path, 'w+', 'gbk')	
	
def export(file, text, script):
	csv=createAndOpenFile(file)
	csv.write(text)
	csv.close()
	print script
	
local = script_path()
print "local:" + local
xlsPath = os.path.join(local,"DBConfig.xls")
print "excel:" + xlsPath
xlsPath = ChangePath(xlsPath)
print "excel:" + xlsPath

try:
	book = xlrd.open_workbook(xlsPath)
except Exception , e:
	print "err:",
	print e

print "book:" + str(book)

Rows = []
dictName = {}
dictIndex = {}

def read_json(row, dcIndex):
	text = ""	
	for column in range(len(dcIndex)):
		if text != "":
			text += ","
		text += '"' + dcIndex[column] + '":"' + read_val(row[column].value) + '"'
	return "{" + text + "}"
class RowInfo:
	def __init__(self):
		self.cname = ""
		self.table = ""
		self.remark = ""
	def to_string(self):
		return "name:" + self.cname + ";table:" + self.table + ";remark:" + self.remark
	def to_json(self):
		text = ReadTemplate("json.xml")
		controls_search = self.read_controls_search()
		if controls_search != "":
			print controls_search
		return text.replace("#remark#", self.remark).replace("#cname#", self.cname).replace("#controls_search#", controls_search)
	def to_sql_xml(self):
		text = ReadTemplate("sql.xml")
		return text.replace("#remark#", self.remark).replace("#cname#", self.cname).replace("#table#", self.table)
	def read_sheet(self, name):
		sheet_name = self.remark + "." + name
		for sheet in book.sheets():
			if sheet.name == sheet_name:
				text = ""	
				for i in range(2, sheet.nrows):
					if text != "":
						text += ",\n"
					text += read_json(sheet.row(i), dictIndex)
				return "[" + text + "]"
		return ""
	def read_controls(self):
		return self.read_sheet("Controls")
	def read_search(self):
		return self.read_sheet("Search")
	def read_form(self):
		text = ReadTemplate("json.xml")
		return text.replace("#remark#", self.remark).replace("#cname#", self.cname)
	def read_controls_search(self):
		controls = self.read_controls()
		search = self.read_search()
		if controls == "" or search == "":
			return ""
		text = ReadTemplate("controls.xml")		
		return text.replace("#controls#", controls).replace("#search#", search).replace("#remark#", self.remark)
		
for sheet in book.sheets():
	if sheet.name == "Table":			
		for i in range(2, sheet.nrows):
			row = RowInfo()
			row.cname = read_val(sheet.row(i)[0].value)
			row.table = read_val(sheet.row(i)[1].value)
			row.remark = read_val(sheet.row(i)[2].value)
			Rows.append(row)
	if sheet.name == "ClolumnInfo":		
		colnames =  sheet.row_values(0) 
		for column in range(len(colnames)):
			dictName[colnames[column]] = column
			dictIndex[column] = colnames[column]
	for i in dictName.keys():
		print i + ":" + str(dictName[i])
	print "**********"
	for i in dictIndex.keys():
		print str(i) + ":" + dictIndex[i]
	
sql = ""
json = ""
for i in Rows:
	json += i.to_json()
	sql += i.to_sql_xml()
	#print i.read_controls()
	#print i.read_search()
	
export("json.xml", json, "json success")
export("sql.xml", sql, "sql success")