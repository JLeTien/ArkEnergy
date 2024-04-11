from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


def generate_pdf(file_name):
    # Create a canvas
    c = canvas.Canvas(file_name, pagesize=letter)

    # Set font and size
    c.setFont("Times-Bold", 12)

    # Title
    text = "Ark Energy Quarterly Report"
    c.drawString(20, 765, text)
    

    # Close the canvas
    c.save()

if __name__ == "__main__":
    # Output file name
    output_pdf = "sample_pdf.pdf"

    # Generate PDF
    generate_pdf(output_pdf)
    print("PDF generated successfully.")