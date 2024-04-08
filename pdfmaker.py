#!/usr/bin/env python
import openpyxl

from PyPDF2 import PdfWriter, PdfReader
import io

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

from reportlab.lib import colors
from reportlab.graphics.shapes import (Drawing, Rect, String, Line, Group)


from reportlab.pdfbase.pdfmetrics import registerFont, stringWidth
from reportlab.pdfbase.ttfonts import TTFont

# font
registerFont(TTFont("Times", "/System/Library/Fonts/Times.ttc"))
registerFont(TTFont('Lumios', 'LumiosMarker.ttf'))
 
# drawing = Drawing(400, 200)
# beige rectangle
# r1 = Rect(0, 0, 400, 200, 0, 0)
# r1.fillColor = colors.beige
# drawing.add(r1)
# save
# drawing.save(formats=['pdf', 'png'], outDir=".", fnRoot="card")



# Blank sheet
file_name ="Certificate_PM.pdf"




def card(cert_name):

	drawing = Drawing(400, 200)

	# name
	name = Group(
	    String(
	        0,
	        100,
	        cert_name,
	        textAnchor='middle',
	        fontName='Times',
	        fontSize=18,
	        fillColor=colors.black
	    )
	)
	name.translate(290, 10)
	drawing.add(name)
	return drawing


# d = card("First Sample")

# save
# d.save(formats=['pdf', 'png'], outDir=".", fnRoot=certificate_name)

# d.save(
#             formats=['png'],
#             outDir="company/",
#             fnRoot="%s-%s" % (row['First Name'], row['Last Name'])
#         )



user_name = "Mr Girish, Organization"


def prepare_certificate(user_name, cert_name):
	font_name = "Lumios"
	font_size = 30

	x = 0
	y = 380 #8.5 * 72

	packet = io.BytesIO()

	# read your existing PDF
	existing_pdf = PdfReader(open(file_name, "rb"))
	output = PdfWriter()
	# add the "watermark" (which is the new pdf) on the existing page
	page = existing_pdf.pages[0]

	print(page.mediabox.width)


	page_size =  (page.mediabox.width, page.mediabox.height)  # 20 inch width and 10 inch height.
	print(page_size)
	can = canvas.Canvas(packet, pagesize=page_size)
	can.setFillColorRGB(194/255, 153/255, 96/255) 
	# can.fillColor = colors.beige

	can.setFont(font_name, font_size)

	textWidth = stringWidth(user_name, font_name, font_size) 
	print(textWidth)
	x = (float(page.mediabox.width)/2.0) - (float(textWidth)/2.0) 

	can.drawString(x, y, user_name)


	can.save()

	#move to the beginning of the StringIO buffer
	packet.seek(0)

	# create a new PDF with Reportlab
	new_pdf = PdfReader(packet)

	page.merge_page(new_pdf.pages[0])
	output.add_page(page)
	# finally, write "output" to a real file
	certificate_path = "Certificates/"+cert_name+".pdf"
	output_stream = open(certificate_path, "wb")
	output.write(output_stream)
	output_stream.close()

# prepare_certificate(user_name)




def getCertificateLabels():

	# Excel sheet File with no col name or numbering - just the list of text
	list_file_name = "UserList.xlsx"

	sheet_name = "Sheet 1"

	text_list = []

	workbook = openpyxl.load_workbook(list_file_name)
	sheets = workbook.active


	first_row = list(sheets.rows)
	for items in first_row:
		name = ""
		for item in items:
			if (item.value!=None):
				name = name+item.value + ", "
			else:
				continue	
		if len(name) > 2:
			text_list.append(name[:-2])

	for labels in text_list:
		print(labels)
		cert_name = labels.split(',', 1)[0]
		prepare_certificate(labels,cert_name)


getCertificateLabels()		

