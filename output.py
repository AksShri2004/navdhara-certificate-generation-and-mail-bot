from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
import io
from data import *


def add_center_aligned_text(input_path, output_path, text, coordinates, font_size=25):
    # Create a temporary PDF with the text using reportlab
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    width, height = letter  # Size of the page

    # Set font and size
    can.setFont("Times-Bold", font_size)

    # Measure text width
    text_width = can.stringWidth(text, "Times-Bold", font_size)

    # Center the text
    x = coordinates[0] - (text_width / 2)
    y = coordinates[1]

    # Draw text
    can.drawString(x, y, text)
    can.save()

    # Move to the beginning of the StringIO buffer
    packet.seek(0)

    # Read the existing PDF and the new PDF with the text
    existing_pdf = PdfReader(input_path)
    new_pdf = PdfReader(packet)
    output = PdfWriter()

    # Add the "watermark" (which is the new PDF) on the existing PDF
    for i in range(len(existing_pdf.pages)):
        page = existing_pdf.pages[i]
        page.merge_page(new_pdf.pages[0])
        output.add_page(page)

    # Write the output PDF to a file
    with open(output_path, 'wb') as outputStream:
        output.write(outputStream)


# Example Usage
input_path = 'parti_.pdf'
output_path = 'vedant rane_.pdf'
text = 'Yahs'
coordinates = [420, 258]  # Coordinates (x, y) where you want to add the text

add_center_aligned_text(input_path, output_path, text, coordinates)
