import os
from fpdf import FPDF # Fpdf v1.7.2
from PyPDF2 import PdfFileWriter, PdfFileReader # PyPDF2 v2.10.9

# Watermark mask for PC
def make_watermark_pdf_PC():
    pdf = FPDF()
    pdf.add_page()

    # razitko2.png
    png1_w, png1_h = 7522, 3121

    # razitko_vyhovuje.png
    png2_w, png2_h = 7500, 2867

    ration = 90

    pdf.image('./png/razitko2.png', x=70, y=150, w=png1_w/ration, h=png1_h/ration)
    pdf.image('./png/razitko_vyhovuje.png', x=100, y=190, w=png2_w / ration, h=png2_h / ration)

    pdf.output("watermark_PC.pdf", "F")


# Watermark mask for PCS
def make_watermark_pdf_PCS():
    pdf = FPDF()
    pdf.add_page()

    # razitko2.png
    png1_w, png1_h = 7522, 3121

    # razitko_vyhovuje.png
    png2_w, png2_h = 7500, 2867

    ration = 90

    pdf.image('./png/razitko2.png', x=70, y=160, w=png1_w/ration, h=png1_h/ration)
    pdf.image('./png/razitko_vyhovuje.png', x=100, y=210, w=png2_w / ration, h=png2_h / ration)

    pdf.output("watermark_PCS.pdf", "F")


# Watermark mask for IP
def make_watermark_pdf_IP():
    pdf = FPDF()
    pdf.add_page()

    # razitko2.png
    png1_w, png1_h = 7522, 3121

    # razitko_vyhovuje.png
    png2_w, png2_h = 7500, 2867

    ration = 95

    pdf.image('./png/razitko2.png', x=77, y=230, w=png1_w/ration, h=png1_h/ration)
    pdf.image('./png/razitko_vyhovuje.png', x=100, y=230 - png1_h/ration * 1.5, w=png2_w / ration, h=png2_h / ration)

    pdf.output("watermark_IP.pdf", "F")


# Watermark mask for Power Supply
def make_watermark_pdf_SZ():
    pdf = FPDF()
    pdf.add_page()

    # razitko2.png
    png1_w, png1_h = 7522, 3121

    # razitko_vyhovuje.png
    png2_w, png2_h = 7500, 2867

    ration = 95

    pdf.image('./png/razitko2.png', x=120, y=265, w=png1_w/ration, h=png1_h/ration)
    pdf.image('./png/razitko_vyhovuje.png', x=40, y=265, w=png2_w / ration, h=png2_h / ration)

    pdf.output("watermark_SZ.pdf", "F")


# Watermark mask for ICPP
def make_watermark_pdf_ICPP():
    pdf = FPDF()
    pdf.add_page()

    # razitko2.png
    png1_w, png1_h = 7522, 3121

    # razitko_vyhovuje.png
    png2_w, png2_h = 7500, 2867

    ration = 95

    pdf.image('./png/razitko2.png', x=130, y=115, w=png1_w/ration, h=png1_h/ration)
    pdf.image('./png/razitko_vyhovuje.png', x=90, y=175, w=png2_w / ration, h=png2_h / ration)

    pdf.output("watermark_ICPP.pdf", "F")


# Watermark mask for PCP
def make_watermark_pdf_PCP():
    pdf = FPDF()
    pdf.add_page()

    # razitko2.png
    png1_w, png1_h = 7522, 3121

    # razitko_vyhovuje.png
    png2_w, png2_h = 7500, 2867

    ration = 95

    pdf.image('./png/razitko2.png', x=130, y=200, w=png1_w/ration, h=png1_h/ration)
    pdf.image('./png/razitko_vyhovuje.png', x=15, y=255, w=png2_w / ration, h=png2_h / ration)

    pdf.output("watermark_PCP.pdf", "F")


# Watermark mask for Space
def make_watermark_pdf_space():
    pdf = FPDF()
    pdf.add_page()

    # razitko2.png
    png1_w, png1_h = 7522, 3121

    # razitko_vyhovuje.png
    png2_w, png2_h = 7500, 2867

    ration = 95

    pdf.image('./png/razitko2.png', x=130, y=76, w=png1_w/ration, h=png1_h/ration)
    pdf.image('./png/razitko_vyhovuje.png', x=100, y=135, w=png2_w / ration, h=png2_h / ration)

    pdf.output("watermark_space.pdf", "F")


# Sign masks
def make_sign_pdf(sign_file_path, pdf_name, pdf_type):
    pdf = FPDF()
    pdf.add_page()

    # Universal coords
    png_sign_w, png_sign_h = 972, 1010
    ration = 70

    # Norm coordinates
    if pdf_type == '':
        x = 179
        y = 240
        pdf_name += '.pdf' 

    # PC coordinates           
    if pdf_type == 'PC':
        x = 170
        y = 164
        pdf_name += '_PC.pdf'

    # PSP coordinates        
    if pdf_type == 'PSP':
        x = 170
        y = 105
        pdf_name += '_PSP.pdf'

    # SZ coordinates        
    if pdf_type == 'SZ':
        x = 180
        y = 244
        pdf_name += '_SZ.pdf' 

    # ICPP/ICPS coordinates      
    if pdf_type == 'ICPP':
        x = 170
        y = 145
        pdf_name += '_ICPP.pdf'

    # PCP coordinates
    if pdf_type == 'PCP':
        x = 179
        y = 230
        pdf_name += '_PCP.pdf'

    # PCS coordinates
    if pdf_type == 'PCS':
        x = 183
        y = 172
        pdf_name += '_PCS.pdf'

    pdf.image(sign_file_path, x=x, y=y, w=png_sign_w / ration, h=png_sign_h / ration)

    sign_pdf_path = r'.\source\signs\\' + pdf_name

    pdf.output(sign_pdf_path, "F")
    return sign_pdf_path


# Adding watermark masks to pdfs
def add_watermark_to_pdf(pdf_file, watermark, result):
    with open(pdf_file, "rb") as input_file, open(watermark, "rb") as watermark_file:
        input_pdf = PdfFileReader(input_file)
        watermark_pdf = PdfFileReader(watermark_file)
        watermark_page = watermark_pdf.pages[0]

        output = PdfFileWriter()

        for i in range(input_pdf.getNumPages()):
            pdf_page = input_pdf.pages[i]
            if i == input_pdf.getNumPages() - 1:
                watermark_page.mergePage(watermark_page)
                pdf_page.mergePage(watermark_page)

            output.addPage(pdf_page)

        with open(result, "wb") as merged_file:
            output.write(merged_file)


# Adding sign masks to pdfs
def add_sign_to_pdf(pdf_file, watermark, result):
    with open(pdf_file, "rb") as input_file, open(watermark, "rb") as watermark_file:
        input_pdf = PdfFileReader(input_file)
        watermark_pdf = PdfFileReader(watermark_file)
        watermark_page = watermark_pdf.pages[0]

        output = PdfFileWriter()

        for i in range(input_pdf.getNumPages()):
            pdf_page = input_pdf.pages[i]
            if i == input_pdf.getNumPages() - 1:
                watermark_page.mergePage(watermark_page)
                pdf_page.mergePage(watermark_page)

            output.addPage(pdf_page)

        with open(result, "wb") as merged_file:
            output.write(merged_file)


# Main function
if __name__ == '__main__':
    """
        pyinstaller -F -c add_temp_to_pdf.py -n Sign_generator
    """
    make_watermark_pdf_SZ()
    make_watermark_pdf_ICPP()
    make_watermark_pdf_PCP()
    make_watermark_pdf_space()
    make_watermark_pdf_PC()
    make_watermark_pdf_PCS()

    for sign in os.listdir(r'.\png'):
        if '_podpis' in sign:
            print(sign.split('_')[0])
            for form_type in ['', 'PC', 'PSP', 'SZ', 'ICPP', 'PCP', 'PCS']:
                sign_pdf = make_sign_pdf(sign_file_path=r'.\png\\' + sign,
                                         pdf_name='sign_' + sign.split('_')[0].lower(),
                                         pdf_type=form_type)

                print(sign_pdf)


