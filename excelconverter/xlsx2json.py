# encoding: utf-8
import json
from .xlsx2x import parseObject

def convert(xlsxfile):
	tempJsonFilename = xlsxfile.replace('.xlsx', '.json')
	with open(tempJsonFilename, 'w') as tempJsonFile:
		tempJsonFile.write(json.dumps(parseObject(xlsxfile), indent = 4))
