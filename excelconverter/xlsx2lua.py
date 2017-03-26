# encoding: utf-8
from slpp import slpp as lua
from .xlsx2x import parseObject

def convert(xlsxfile, tablename):
	tempLuaFilename = xlsxfile.replace('.xlsx', '.lua')
	with open(tempLuaFilename, 'w') as tempLuaFile:
		tempLuaFile.write(tablename + ' = ' + lua.encode(parseObject(xlsxfile)))
