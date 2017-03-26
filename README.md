# Excelconverter
## Install
```shell
## install depends first and then
cd excelconverter
pip install .
```

## Depends
 * [Simple lua-python parser](https://github.com/sric0880/slpp)

## Usage
 * convert from xlsx to json:
```py
from excelconverter import convertXlsx2Json
convertXlsx2Json(xlsx_filepath_name)
```

* convert from xlsx to lua:
```py
from excelconverter import convertJson2Xlsx
convertJson2Xlsx(json_filepath_name)
```

* convert from json to xlsx:
```py
from excelconverter import convertJson2Xlsx
convertJson2Xlsx(json_filepath_name)
```
