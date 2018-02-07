from setuptools import setup

setup(name='excelconverter',
      version='1.1.2',
      description='python tool for coverting from xlsx to json/lua and from json to xlsx, support complex json format',
      url='https://github.com/sric0880/excelconverter.git',
      author='sric0880',
      author_email='justgotpaid88@qq.com',
      license='MIT',
      packages=['excelconverter'],
      install_requires=[
          'json2xlsx',
      ],
      zip_safe=False)