from PyPDF2 import PdfFileReader, PdfFileWriter


def split(path, name_of_split):
    pdf = PdfFileReader(path)

    for page in range(pdf.getNumPages()):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))

        output = f'{name_of_split}{page}.pdf'
        with open(output, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)


if __name__ == '__main__':
    path = r'D:\Users\Administrator\Desktop\非洲猪瘟常态下猪场生物安全体系建设.pdf'
    split(path, '0')