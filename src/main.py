import xlrd
import tabula
import fpdf
import sys
import os.path

class myPDF(fpdf.FPDF):
    new_pdf = fpdf.FPDF()
    flags = []

    def __init__(self):
        pass

    def readPDF(filename):
        text = tabula.io.read_pdf(filename)
        if len(text[0]) == 0:
            raise
        my_text = text[0].values.tolist()
        return my_text

    def compResolve(self, pdfObject, XLSObject):
        self.flags = [0]*len(pdfObject)
        index = 0
        for pdf in pdfObject:
            for xls in XLSObject:
                if pdf[0] == xls[0].value:
                    pdf[0] = xls[1].value
                    self.flags[index] = 1
            index = index + 1
        return pdfObject


    def writeFile(self, mypdfObject, newPath):
        new_pdf = fpdf.FPDF()
        new_pdf.add_page()
        new_pdf.set_font("Arial", size = 12)

        for index, flag in zip(mypdfObject, self.flags):
            if flag == 1:
                new_pdf.set_fill_color(0,150,0)
                new_pdf.cell(100,10, str(''.join([str(elem).strip() + "  " for elem in index])), border =1 , ln=1, align = 'C', fill = True)
            else:
                new_pdf.set_fill_color(230,0,0)
                new_pdf.cell(100,10, str(''.join([str(elem).strip() + "  " for elem in index])), border =0 , ln=1, align = 'C', fill = True)
        new_pdf.output(newPath)


class myXLS:
    def __init__(self):
        pass

    def readXLS( filename):
        wb= xlrd.open_workbook(filename)
        sheet = wb.sheet_by_index(0)
        pageObject = []
        for i in range(sheet.nrows):
            pageObject.append(sheet.row(i))
        wb.release_resources()
        return pageObject


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise ValueError("Not enough arguments")
    pdfPath = sys.argv[1]
    xlsObjectPath = sys.argv[2]
    newPath = sys.argv[3]
    try:
        os.path.exists(pdfPath)
        os.path.exists(xlsObjectPath)
        os.path.isdir(newPath)
    except FileNotFoundError:
        raise FileNotFoundError

    pdfObject = myPDF.readPDF(pdfPath)
    xlsObject = myXLS.readXLS(xlsObjectPath)

    my_ret = myPDF.compResolve(myPDF, pdfObject,xlsObject)
    myPDF.writeFile(myPDF,my_ret,newPath)


