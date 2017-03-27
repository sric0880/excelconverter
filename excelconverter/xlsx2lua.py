# encoding: utf-8
from slpp import slpp as lua
from xlsx2x import parseObject

def convert(xlsxfile, tablename):
	o = parseObject(xlsxfile)
	if isinstance(o, list) and type(o[0]) is dict:
		if 'Key' in o[0] and 'Value' in o[0]:
			o = dict([ (item['Key'], item['Value']) for item in o ])
	tempLuaFilename = xlsxfile.replace('.xlsx', '.lua')
	with open(tempLuaFilename, 'w') as tempLuaFile:
		tempLuaFile.write(tablename + ' = ' + lua.encode(o))
