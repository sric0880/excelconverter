from setuptools import setup

setup(name='excelconverter',
      version='1.2.2',
      description='python tool for coverting from xlsx to json/lua and from json to xlsx, support complex json format',
      url='https://github.com/sric0880/excelconverter.git',
      author='sric0880',
      author_email='justgotpaid88@qq.com',
      license='MIT',
      packages=['excelconverter'],
      scripts= [
        'bin/json2xlsx',
      ],
      install_requires=[
          'argparse>=1.2.1',
          'openpyxl>=2.3',
          'pyparsing>=1.5.5'
      ],
      zip_safe=False)