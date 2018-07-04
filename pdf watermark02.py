import PyPDF2
import os
os.chdir('C:\\Users\\peter\\新增資料夾')

minutesFile = open('meetingminutes.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(minutesFile)
watermarkFile = open('watermark.pdf', 'rb')
pdfWatermarkReader = PyPDF2.PdfFileReader(watermarkFile)

pdfWriter = PyPDF2.PdfFileWriter()

for pageNum in range(0, pdfReader.numPages):
    minutesFirstPage = pdfReader.getPage(pageNum)
    minutesFirstPage.mergePage(pdfWatermarkReader.getPage(0))
    pdfWriter.addPage(minutesFirstPage)



#for pageNum in range(1, pdfReader.numPages):
#    pageObj = pdfReader.getPage(pageNum)
#    pdfWriter.addPage(pageObj)

resultPdfFile = open('watermarkedCover.pdf', 'wb')
pdfWriter.write(resultPdfFile)
minutesFile.close()
resultPdfFile.close()
watermarkFile.close()
