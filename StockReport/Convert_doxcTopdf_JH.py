import win32com.client
import os
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx2pdf import convert

import datetime


today_time = datetime.date.today().strftime("%Y%m%d")  

################################ pdf 변환 #################################################### 
inputFile = "C:\\PYTHONWORKSPACE\\webscraping_basic\\webscraping_project\\2022\\Y&R_리포트_{}_장마감.docx".format(today_time)
outputFile = "C:\\PYTHONWORKSPACE\\webscraping_basic\\webscraping_project\\2022\\Y&R_리포트_{}_장마감.pdf".format(today_time)
file = open(outputFile, "w")
file.close()

convert(inputFile, outputFile)
