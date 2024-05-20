from PyPDF2 import PdfWriter, PdfReader # PyPDF2 v2.10.9
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject, NumberObject, create_string_object
from PyPDF2.constants import FieldFlag
import pandas as pd # Pandas v1.5.3
import json
import textwrap


def set_need_appearances_writer(writer: PdfWriter):
    try:
        catalog = writer._root_object

        if "/AcroForm" not in catalog:
            writer._root_object.update({
                NameObject("/AcroForm"): IndirectObject(len(writer._objects), 0, writer)
            })

        writer._root_object["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)})

        return writer

    except Exception as e:
        print('set_need_appearances_writer() catch : ', repr(e))
        return writer


# Update values in pdfs, adding signs and watermarks
def update_form_values(infile, outfile, newvals, sign_file_name,
                       watermark_file_path=r'./source/watermarks/watermark.pdf'):
    sign_found = False

    # Opening file
    with open(infile, 'rb') as input_file, open(watermark_file_path, 'rb') as watermark_file:
        
        # Assigning names of technicians
        if '_' in sign_file_name:
            if sign_file_name.split('_')[0] in name_to_sign:
                sign_name = name_to_sign[sign_file_name.split('_')[0]] + '_' + sign_file_name.split('_')[1] + '.pdf'
                sign_found = True
        else:
            if sign_file_name.split('_')[0] in name_to_sign:
                sign_name = name_to_sign[sign_file_name] + '.pdf'
                sign_found = True

        # Finding correct sign mask
        if sign_found:
            sing_file = open('./source/signs/sign_' + sign_name, 'rb')
            sign_pdf = PdfReader(sing_file)
            sign_page = sign_pdf.pages[0]

        pdf = PdfReader(input_file)

        watermark_pdf = PdfReader(watermark_file)
        watermark_page = watermark_pdf.pages[0]

        writer = PdfWriter()
        writer.append_pages_from_reader(pdf)
        writer.set_need_appearances_writer()

        # Update form values
        for index, page in enumerate(writer.pages):
            for j in range(0, len(page['/Annots'])):
                writer_annot = page['/Annots'][j].get_object()

            try:
                writer.update_page_form_field_values(page=page, fields=newvals, flags=FieldFlag.READ_ONLY)
                for j in range(0, len(page['/Annots'])):
                    writer_annot = page['/Annots'][j].get_object()
                    if NameObject("/Ff") not in writer_annot:
                        writer_annot[NameObject("/Ff")] = NumberObject(1)
                    writer_annot.update({
                        NameObject("/Ff"): NumberObject(1)  # make ReadOnly
                    })

                # Add watermarks
                if page == writer.pages[-1]:
                    page.merge_page(watermark_page)

                    if sign_found:
                        page.merge_page(sign_page)

            except Exception as e:
                print(repr(e))

        with open(outfile, 'wb') as out:
            writer.set_need_appearances_writer()
            writer.write(out)

        if sign_found:
            sing_file.close()

        return outfile


# Printing headers
def print_header(header_text: str):
    print('=' * (len(header_text) + 5))
    print(header_text)
    print('-' * len(header_text))


#Infusomat P
def iP_gb_3_0_STK():
    print_header('Generating "iP_gb_3_0_STK" forms')

    # Name of excel sheet and pdf file
    sheet_name = 'IP'
    _pdf_file_name = 'iP_gb_3_0_STK-Form_prefilled.pdf'

    # Reading data from excel 
    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name)
    excel_data_df = excel_data_df.iloc[1:]

    # Transforming date format and setting the offset
    excel_data_df['Datum'] = pd.to_datetime(excel_data_df['Datum'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['Nächster Termin'] = pd.to_datetime(excel_data_df['Datum'], dayfirst=True) + pd.offsets.DateOffset(years=2)
    excel_data_df['Kalibriert'] = pd.to_datetime(excel_data_df['Kalibriert'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['Nächster Termin'] = pd.to_datetime(excel_data_df['Nächster Termin'], dayfirst=True).dt.strftime('%d.%m.%Y')

    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)

    rows_to_generate = list()

    # Writing the additional documentation in multiple rows (hours, sw, parts)
    for index, row in enumerate(json_rows):
        new_ell = str()
        for cell in row.copy():
            if 'Besonderheiten  Dokumentation 1' in cell:
                new_ell += str(row[cell]) + ',    '
                row.pop(cell)
            if 'Besonderheiten  Dokumentation 2' in cell:
                new_ell += '\n'+ str(row[cell]) 
                row.pop(cell)
        new_ell += '\n'

        row['Besonderheiten  Dokumentation 1'] = new_ell
        rows_to_generate.append(row)

    generated_pdfs = dict()

    # Looping over lines and updating the pdf forms
    for index, row in enumerate(rows_to_generate):
        row['BetreiberRow1'] = textwrap.fill(row['BetreiberRow1'], 20)
        date_split = str(row['Datum']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,
                                 outfile='./' + '8712344_' + str(row['GeräteNrRow1']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['Kontrolle durchgeführt von'] ,
                                 watermark_file_path=r'./source/watermarks/watermark_IP.pdf')

        generated_pdfs[out] = row
        print('\t> File: ' + out + ' generated.')

    # Merging separate pdf files (if number of pdfs > 1)
    pdf_merge(generated_pdfs, sheet_name + '_merged.pdf')


#Perfusor Space
def STK_Form_PSP_10_0_en():
    print_header('Generating "STK Form PSP 10 0" forms')

    # Name of excel sheet
    sheet_name = 'PSP'

    # Reading data from excel
    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name)
    excel_data_df = excel_data_df.iloc[1:]

    # Transforming date format and setting the offset
    excel_data_df['3-25'] = pd.to_datetime(excel_data_df['3-24'],dayfirst=True) + pd.offsets.DateOffset(years=2)
    excel_data_df['3-24'] = pd.to_datetime(excel_data_df['3-24'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['2-106b'] = pd.to_datetime(excel_data_df['2-106b'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['3-25'] = pd.to_datetime(excel_data_df['3-25'], dayfirst=True).dt.strftime('%d.%m.%Y')

    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)

    rows_to_generate = list()

    # Writing the additional documentation in multiple rows (hours, sw, parts)
    for index, row in enumerate(json_rows):
        new_ell = str()
        for cell in row.copy():
            if str(cell).startswith('3-15'):
                new_ell += str(row[cell]) + '\r\n'
                row.pop(cell)

            if str(cell).startswith('1-01'):
                row[cell] = textwrap.fill(row[cell], 20)

        row['3-15'] = new_ell
        rows_to_generate.append(row)

    generated_pdfs = dict()

    # Looping over lines and updating the pdf forms
    for index, row in enumerate(rows_to_generate):
        date_split = str(row['3-24']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]

        if row['T3a'] == 'y':
            _pdf_file_name = 'STK-Form_PSP_10_0_en_prefilled_y.pdf'
        elif row['T3a'] == 'n' and row['T3b'] == 'p':
            _pdf_file_name = 'STK-Form_PSP_10_0_en_prefilled_n_p.pdf'
        else:
            _pdf_file_name = 'STK-Form_PSP_10_0_en_prefilled_n_f.pdf'

        row.pop('T3a')
        row.pop('T3b')

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,
                                 outfile='./' + '8713030_' + str(row['1-04']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['3-21'] + '_PSP',
                                 watermark_file_path=r'./source/watermarks/watermark_space.pdf')
        generated_pdfs[out] = row
        print('\t> File: ' + out + ' generated.')

    # Merging separate pdf files (if number of pdfs > 1)
    pdf_merge(generated_pdfs, sheet_name + '_merged.pdf')


#Infusomat Space
def STK_Form_ISP_7_0_en3():
    print_header('Generating "STK Form ISP 7 0 en3" forms')

    # Name of excel sheet
    sheet_name = 'ISP'

    # Reading data from excel
    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name)
    excel_data_df = excel_data_df.iloc[2:]

    # Transforming date format and setting the offset
    excel_data_df['3-25'] = pd.to_datetime(excel_data_df['3-24']) + pd.offsets.DateOffset(years=2)
    excel_data_df['3-24'] = pd.to_datetime(excel_data_df['3-24']).dt.strftime('%d.%m.%Y')
    excel_data_df['2-102b'] = pd.to_datetime(excel_data_df['2-102b']).dt.strftime('%d.%m.%Y')
    excel_data_df['3-25'] = pd.to_datetime(excel_data_df['3-25']).dt.strftime('%d.%m.%Y')

    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)

    # Writing the additional documentation in two rows (hours, sw, parts)
    rows_to_generate = list()
    for index, row in enumerate(json_rows):
        new_ell = str()
        for cell in row.copy():
            if str(cell).startswith('3-15'):
                new_ell += str(row[cell]) + '\r\n'
                row.pop(cell)

            if str(cell).startswith('1-01'):
                row[cell] = textwrap.fill(row[cell], 20)

        row['3-15'] = new_ell
        rows_to_generate.append(row)

    generated_pdfs = dict()
    for index, row in enumerate(rows_to_generate):
        date_split = str(row['3-24']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]

        if row['T3a'] == 'y':
            _pdf_file_name = 'STK-Form_ISP_7_0_en3_y.pdf'
        elif row['T3a'] == 'n' and row['T3b'] == 'p':
            _pdf_file_name = 'STK-Form_ISP_7_0_en3_n_p.pdf'
        else:
            _pdf_file_name = 'STK-Form_ISP_7_0_en3_n_f.pdf'

        row.pop('T3a')
        row.pop('T3b')

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,
                                 outfile='./' + '8713050_' + str(row['1-04']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['3-21'] + '_PSP',
                                 watermark_file_path=r'./source/watermarks/watermark_space.pdf')

        generated_pdfs[out] = row
        print('\t> File: ' + out + ' generated.')

    # Merging separate pdf files (if number of pdfs > 1)
    pdf_merge(generated_pdfs, sheet_name + '_merged.pdf')


#nfusomat Space P
def STK_Form_ISPP_6_0_enver2():
    print_header('Generating "STK-Form_ISPP_6_0_enver2" forms')

    # Name of excel sheet
    sheet_name = 'ISPP'
    
    # Reading data from excel
    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name)
    excel_data_df = excel_data_df.iloc[2:]

    # Transforming date format and setting the offset
    excel_data_df['3-25'] = pd.to_datetime(excel_data_df['3-24']) + pd.offsets.DateOffset(years=2)
    excel_data_df['3-24'] = pd.to_datetime(excel_data_df['3-24']).dt.strftime('%d.%m.%Y')
    excel_data_df['2-102b'] = pd.to_datetime(excel_data_df['2-102b']).dt.strftime('%d.%m.%Y')
    excel_data_df['3-25'] = pd.to_datetime(excel_data_df['3-25']).dt.strftime('%d.%m.%Y')

    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)

    rows_to_generate = list()

    # Writing the additional documentation in two rows (hours, sw, parts)
    for index, row in enumerate(json_rows):
        new_ell = str()

        for cell in row.copy():
            if str(cell).startswith('3-15'):
                new_ell += str(row[cell]) + '\r\n'
                row.pop(cell)

            if str(cell).startswith('1-01'):
                row[cell] = textwrap.fill(row[cell], 20)

        row['3-15'] = new_ell
        rows_to_generate.append(row)

    generated_pdfs = dict()

    # Looping over lines and updating the pdf forms
    for index, row in enumerate(rows_to_generate):
        date_split = str(row['3-24']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]

        if row['T3a'] == 'y':
            _pdf_file_name = 'STK-Form_ISPP_6_0_enver2_y.pdf'
        elif row['T3a'] == 'n' and row['T3b'] == 'p':
            _pdf_file_name = 'STK-Form_ISPP_6_0_enver2_n_p.pdf'
        else:
            _pdf_file_name = 'STK-Form_ISPP_6_0_enver2_n_f.pdf'

        row.pop('T3a')
        row.pop('T3b')

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,
                                 outfile='./' + '8713070_' + str(row['1-04']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['3-21'] + '_PSP',
                                 watermark_file_path=r'./source/watermarks/watermark_space.pdf')

        generated_pdfs[out] = row
        print('\t> File: ' + out + ' generated.')

    # Merging separate pdf files (if number of pdfs > 1)
    pdf_merge(generated_pdfs, sheet_name + '_merged.pdf')


#Perfusor Compact
def Perfusor_compact():
    print_header('Generating "Perfusor_compact_3_1_gb_onlinetsc" forms')

    # Name of excel sheet and pdf file
    sheet_name = 'PC'
    _pdf_file_name = 'Perfusor_compact_3_1_gb_onlinetsc_prefilled.pdf'

    # Reading data from excel
    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name)
    excel_data_df = excel_data_df.iloc[1:]

    # Transforming date format and setting the offset
    excel_data_df['Next deadline'] = pd.to_datetime(excel_data_df['Date  Signature'], dayfirst=True) + pd.offsets.DateOffset(years=2)
    excel_data_df['Date  Signature'] = pd.to_datetime(excel_data_df['Date  Signature'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['Next deadline'] = pd.to_datetime(excel_data_df['Next deadline'], dayfirst=True).dt.strftime('%d.%m.%Y')

    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)
    generated_pdfs = dict()

    # Looping over lines and updating the pdf forms
    for index, row in enumerate(json_rows):
        date_split = str(row['Date  Signature']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]
        row['User'] = textwrap.fill(row['User'], 20)

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,      # nazev vygenerovaneho pdf
                                 outfile='./' + '8714827_' + str(row['Unit NoRow1']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['Inspection performed by'] + '_PC',
                                 watermark_file_path=r'./source/watermarks/watermark_PC.pdf')
        generated_pdfs[out] = row

        print('\t> File: ' + out + ' generated.')

    # Merging separate pdf files (if number of pdfs > 1)
    pdf_merge(generated_pdfs, sheet_name + '_merged.pdf')


#Perfusor Compact S
def Perfusor_compact_s():
    print_header('Generating "Perfusor_compact_s_3_1_gb_onlinetsc" forms')

    # Name of excel sheet and pdf file
    sheet_name = 'PCS'
    _pdf_file_name = 'TSC-Perf_Com._S_EN_filled.pdf'

    # Reading data from excel
    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name) #precteni excel listu
    excel_data_df = excel_data_df.iloc[1:]

    # Transforming date format and setting the offset
    excel_data_df['Next deadline for TSC'] = pd.to_datetime(excel_data_df['Date  Signature'], dayfirst=True) + pd.offsets.DateOffset(years=2)
    excel_data_df['Date  Signature'] = pd.to_datetime(excel_data_df['Date  Signature'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['Next deadline for TSC'] = pd.to_datetime(excel_data_df['Next deadline for TSC'], dayfirst=True).dt.strftime('%d.%m.%Y')

    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)

    # Writing the additional documentation in two rows (hours, sw, parts)
    rows_to_generate = list()
    for index, row in enumerate(json_rows):
        new_ell = str()
        for cell in row.copy():
            if 'Copy fillin and attach to documentation 1' in cell:
                new_ell += str(row[cell]) + ',    '
                row.pop(cell)
            if 'Copy fillin and attach to documentation 2' in cell:
                new_ell += '\n' + str(row[cell])
                row.pop(cell)

        row['Copy fillin and attach to documentation'] = new_ell
        rows_to_generate.append(row)

    generated_pdfs = dict()

    # Looping over lines and updating the pdf forms
    for index, row in enumerate(json_rows):
        date_split = str(row['Date  Signature']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]
        row['Textfield'] = textwrap.fill(row['Textfield'], 20)

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,      # nazev vygenerovaneho pdf
                                 outfile='./' + '8714843_' + str(row['Unit No']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['Inspection performed by'] + '_PCS',
                                 watermark_file_path=r'./source/watermarks/watermark_PCS.pdf')
        generated_pdfs[out] = row

        print('\t> File: ' + out + ' generated.')

    # Merging separate pdf files (if number of pdfs > 1)
    pdf_merge(generated_pdfs, sheet_name + '_merged.pdf')


# Infusomat Compact Plus P
def infusomat_compact_plus_p():
    print_header('Generating "Infusomat_compact_plus_3_1_gb_onlinetsc" forms')

    # Name of excel sheet and pdf file
    sheet_name = 'ICPP'
    _pdf_file_name = 'STK-Form_ICPP_2_0_en_prefilled.pdf'

    # Reading data from excel
    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name)
    excel_data_df = excel_data_df.iloc[1:]
    
    # Transforming date format and setting the offset
    excel_data_df['2-63'] = pd.to_datetime(excel_data_df['2-63'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['2-64'] = pd.to_datetime(excel_data_df['2-63'], dayfirst=True) + pd.offsets.DateOffset(years=2)
    excel_data_df['2-21b'] = pd.to_datetime(excel_data_df['2-21b'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['2-30b'] = pd.to_datetime(excel_data_df['2-30b'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['2-64'] = pd.to_datetime(excel_data_df['2-64'], dayfirst=True).dt.strftime('%d.%m.%Y')    
    
    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)

    # Writing the additional documentation in two rows (hours, sw, parts)
    rows_to_generate = list()
    for index, row in enumerate(json_rows):
        new_ell = str()
        for cell in row.copy():
            if '2-54 1' in cell:
                new_ell += str(row[cell]) + ',    '
                row.pop(cell)
            if '2-54 2' in cell:
                new_ell += '\n'+ str(row[cell])
                row.pop(cell)

        row['2-54'] = new_ell
        rows_to_generate.append(row)

    generated_pdfs = dict()
    
    # Looping over lines and updating the pdf forms
    for index, row in enumerate(rows_to_generate):
        row['1-01'] = textwrap.fill(row['1-01'], 20)
        date_split = str(row['2-63']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,
                                 outfile='./' + '8717070_' + str(row['1-05']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['2-60'] + '_ICPP' ,
                                 watermark_file_path=r'./source/watermarks/watermark_ICPP.pdf')

        generated_pdfs[out] = row
        print('\t> File: ' + out + ' generated.')

    # Merging separate pdf files (if number of pdfs > 1)
    pdf_merge(generated_pdfs, sheet_name + '_merged.pdf')


# Infusomat Compact Plus S
def infusomat_compact_plus_s():
    print_header('Generating "Infusomat_compact_plus_3_1_gb_onlinetsc" forms')

    # Name of excel sheet and pdf file
    sheet_name = 'ICPS'
    _pdf_file_name = 'STK-Form_ICPS_3_0_en_prefilled.pdf'

    # Reading data from excel
    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name)
    excel_data_df = excel_data_df.iloc[1:]
    
    # Transforming date format and setting the offset
    excel_data_df['2-64'] = pd.to_datetime(excel_data_df['2-64'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['2-65'] = pd.to_datetime(excel_data_df['2-64'], dayfirst=True) + pd.offsets.DateOffset(years=2)
    excel_data_df['2-21b'] = pd.to_datetime(excel_data_df['2-21b'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['2-30b'] = pd.to_datetime(excel_data_df['2-30b'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['2-65'] = pd.to_datetime(excel_data_df['2-65'], dayfirst=True).dt.strftime('%d.%m.%Y')    
    
    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)

    # Writing the additional documentation in two rows (hours, sw, parts)
    rows_to_generate = list()
    for index, row in enumerate(json_rows):
        new_ell = str()
        for cell in row.copy():
            if '2-55 1' in cell:
                new_ell += str(row[cell]) + ',    '
                row.pop(cell)
            if '2-55 2' in cell:
                new_ell += '\n'+ str(row[cell])
                row.pop(cell)

        row['2-55'] = new_ell
        rows_to_generate.append(row)

    generated_pdfs = dict()
    
    # Looping over lines and updating the pdf forms
    for index, row in enumerate(rows_to_generate):
        row['1-01'] = textwrap.fill(row['1-01'], 20)
        date_split = str(row['2-64']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,
                                 outfile='./' + '8717050_' + str(row['1-05']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['2-61'] + '_ICPP' ,
                                 watermark_file_path=r'./source/watermarks/watermark_ICPP.pdf')

        generated_pdfs[out] = row
        
        print('\t> File: ' + out + ' generated.')
        
    # Merging separate pdf files (if number of pdfs > 1)
    pdf_merge(generated_pdfs, sheet_name + '_merged.pdf')


# Infusomat FMS
def infusomat_FMS():
    print_header('Generating "Infusomat_FMS_3_1_gb_onlinetsc" forms')

    # Name of excel sheet
    sheet_name = 'IFMS'

    _pdf_file_name = "ifmS_gb_3_0_STK-Form_prefilled.pdf"

    # Reading data from excel, transforming date format and setting the offset
    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name)
    excel_data_df = excel_data_df.iloc[1:]
    excel_data_df['Datum'] = pd.to_datetime(excel_data_df['Datum'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['Nächster Termin'] = pd.to_datetime(excel_data_df['Datum'], dayfirst=True) + pd.offsets.DateOffset(years=2)
    excel_data_df['Kalibriert'] = pd.to_datetime(excel_data_df['Kalibriert'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['Nächster Termin'] = pd.to_datetime(excel_data_df['Nächster Termin'], dayfirst=True).dt.strftime('%d.%m.%Y')

    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)

    rows_to_generate = list()

    # Writing the additional documentation in two rows (hours, sw, parts)
    for index, row in enumerate(json_rows):
        new_ell = str()
        for cell in row.copy():
            if 'Besonderheiten  Dokumentation 1' in cell:
                new_ell += str(row[cell]) + ',    '
                row.pop(cell)
            if 'Besonderheiten  Dokumentation 2' in cell:
                new_ell += '\n' + str(row[cell]) + '\n'
                row.pop(cell)

        row['Besonderheiten  Dokumentation 1'] = new_ell
        rows_to_generate.append(row)

    generated_pdfs = dict()

    # Looping over lines and updating the pdf forms
    for index, row in enumerate(rows_to_generate):
        row['BetreiberRow1'] = textwrap.fill(row['BetreiberRow1'], 20)
        date_split = str(row['Datum']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,
                                 outfile='./' + '8715580_' + str(row['GeräteNrRow1']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['Kontrolle durchgeführt von'] ,
                                 watermark_file_path=r'./source/watermarks/watermark_IP.pdf')

        generated_pdfs[out] = row
        print('\t> File: ' + out + ' generated.')

    # Merging separate pdf files (if number of pdfs > 1)
    pdf_merge(generated_pdfs, sheet_name + '_merged.pdf')


# Perfusor Compact Plus
def perfusor_compact_plus():
    print_header('Generating "Perfusor_compact_plus_3_1_gb_onlinetsc" forms')

    # Name of excel sheet and pdf file
    sheet_name = 'PCP'
    _pdf_file_name = 'STK-Form_PCP_3_0_en_prefilled.pdf'

    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name)
    excel_data_df = excel_data_df.iloc[1:]

    # Reading data from excel, transforming date format and setting the offset
    excel_data_df['2-63'] = pd.to_datetime(excel_data_df['2-63'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['2-64'] = pd.to_datetime(excel_data_df['2-63'], dayfirst=True) + pd.offsets.DateOffset(years=2)
    excel_data_df['2-30b'] = pd.to_datetime(excel_data_df['2-30b'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['2-31b'] = pd.to_datetime(excel_data_df['2-31b'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['2-64'] = pd.to_datetime(excel_data_df['2-64'], dayfirst=True).dt.strftime('%d.%m.%Y')    
    
    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)

    # Writing the additional documentation in two rows (hours, sw, parts)
    rows_to_generate = list()
    for index, row in enumerate(json_rows):
        new_ell = str()
        for cell in row.copy():
            if '2-54 1' in cell:
                new_ell += str(row[cell]) + ',    '
                row.pop(cell)
            if '2-54 2' in cell:
                new_ell += '\n' + str(row[cell])
                row.pop(cell)

        row['2-54'] = new_ell
        rows_to_generate.append(row)

    generated_pdfs = dict()
    
    # Looping over lines and updating the pdf forms
    for index, row in enumerate(rows_to_generate):
        row['1-01'] = textwrap.fill(row['1-01'], 20)
        date_split = str(row['2-63']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,
                                 outfile='./' + '8717030_' + str(row['1-05']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['2-60'] + '_PCP',
                                 watermark_file_path=r'./source/watermarks/watermark_PCP.pdf')

        generated_pdfs[out] = row
        print('\t> File: ' + out + ' generated.')

    # Merging separate pdf files (if number of pdfs > 1)
    pdf_merge(generated_pdfs, sheet_name + '_merged.pdf')


# Power Supply Space
def Power_Supply():
    print_header('Generating "Power_supply_SP_3_1_gb_onlinetsc" forms')

    # Name of excel sheet and pdf file
    sheet_name = 'SZ'
    _pdf_file_name = 'STK-Form-PowerSupply_10_0_en_prefilled.pdf'

    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name)
    excel_data_df = excel_data_df.iloc[1:]

    # Reading data from excel, transforming date format and setting the offset
    excel_data_df['1-44'] = pd.to_datetime(excel_data_df['1-44'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['1-45'] = pd.to_datetime(excel_data_df['1-44'], dayfirst=True) + pd.offsets.DateOffset(years=2)
    excel_data_df['1-45'] = pd.to_datetime(excel_data_df['1-45'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['1-25b'] = pd.to_datetime(excel_data_df['1-25b'], dayfirst=True).dt.strftime('%d.%m.%Y')

    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)

    generated_pdfs = dict()

    # Looping over lines and updating the pdf forms
    for index, row in enumerate(json_rows):
        date_split = str(row['1-44']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]
        row['1-01'] = textwrap.fill(row['1-01'], 20)

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,
                                 outfile='./' + '8713110A_' + str(row['1-04']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['1-41'] + '_SZ',
                                 watermark_file_path=r'./source/watermarks/watermark_SZ.pdf')
        generated_pdfs[out] = row

        print('\t> File: ' + out + ' generated.')

    # Merging separate pdf files (if number of pdfs > 1)
    pdf_merge(generated_pdfs, sheet_name + '_merged.pdf')


# TODO SPS_en_6_0_STK-Form
#Space Station --nefunkcni
def SPS_en_6_0_STK():
    print_header('Generating "SPS_en_6_0_STK-Form" forms')

    sheet_name = 'SPST'
    _pdf_file_name = 'SPS_en_6_0_STK-Form.pdf'
    _pdf_file_name = 'SPS_en_6_0_STK-Form_new.pdf'

    excel_data_df = pd.read_excel('./source/source_tables.xlsx', sheet_name=sheet_name, usecols='A:Q')
    excel_data_df = excel_data_df.iloc[1:]

    excel_data_df['Datum'] = pd.to_datetime(excel_data_df['Datum'], dayfirst=True).dt.strftime('%d.%m.%Y')
    excel_data_df['Kalibriert bis'] = pd.to_datetime(excel_data_df['Kalibriert bis'], dayfirst=True).dt.strftime('%d.%m.%Y')

    json_str = excel_data_df.to_json(orient='records')
    json_rows = json.loads(json_str)
    generated_pdfs = dict()

    for index, row in enumerate(json_rows):
        print(row)
        num_spacestations = 3

        for i in range(2, 5):
            if str(row['ArtikelNrRow' + str(i)]) == '8713142 ':
                ...

        for i in range(3, 5):
            print(row['ArtikelNrRow' + str(i)])
            if row['ArtikelNrRow' + str(i)] is None:
                row.pop('ArtikelNrRow' + str(i))
                num_spacestations -= 1
            if row['GeräteNrRow' + str(i)] is None:
                row.pop('GeräteNrRow' + str(i))

        if str(row['ArtikelNrRow1']) == '8713145':
            ...

        date_split = str(row['Datum']).split('.')
        date = date_split[2] + date_split[1] + date_split[0]

        row['BetreiberRow1'] = textwrap.fill(row['BetreiberRow1'], 20)

        print(f'{num_spacestations=}')
        if num_spacestations == 1:
            row['fill_2'] = 10
            row['fill_3'] = 35
            row['fill_4'] = 0.12
        else:
            row[str(num_spacestations) + ' SpaceStations'] = '/On'
            if num_spacestations == 2:
                row['fill_5'] = 10.0
                row['fill_6'] = 35
                row['fill_7'] = 0.12
            if num_spacestations == 3:
                row['fill_8'] = 10
                row['fill_9'] = 35
                row['fill_10'] = 0.12

        _pdfs_to_merge = dict()
        file_name = './' + str(row['ArtikelNrRow2']) + '_' + str(row['SerienNr 1']) + '_BTK_' + date + '.pdf'

        out = update_form_values(infile='./source/source_forms/' + _pdf_file_name,
                                 outfile='./' + str(row['ArtikelNrRow2']) + '_' + str(
                                     row['SerienNr 1']) + '_BTK_' + date + '.pdf',
                                 newvals=row,
                                 sign_file_name=row['Kontrolle durchgeführt von'])

        generated_pdfs[out] = row
        print('\t> File: ' + file_name + ' generated.')
        exit(0)


# Merging pdfs if there is more than one of the same type
def pdf_merge(pdfs_to_merge: dict, merged_file_name: str):
    if len(pdfs_to_merge) <= 1:
        return

    pdf = PdfReader(list(pdfs_to_merge)[0])
    add_blank_page = False
    num_pages = len(pdf.pages)
    if num_pages % 2 == 1:
        add_blank_page = True

    writer = PdfWriter()
    writer = set_need_appearances_writer(writer)

    for index_file, pdf_to_merge in enumerate(pdfs_to_merge):
        print(f'{pdf_to_merge=}')
        pdf = PdfReader(pdf_to_merge)
        for index_page, page in enumerate(pdf.pages):

            writer.update_page_form_field_values(page, pdfs_to_merge[pdf_to_merge])

            for j in range(0, len(page['/Annots'])):
                writer_annot = page['/Annots'][j].get_object()

                print(f'{writer_annot.get("/T")=}')
                writer_annot.update({
                    NameObject("/Ff"): NumberObject(1) # make ReadOnly
                })
                if '/T' in writer_annot:
                    writer_annot.update({
                        NameObject('/T'): create_string_object(writer_annot.get('/T') + '_' + str(index_file) + str(index_page))
                    })
                print(f'{writer_annot.get("/T")=}')
                print('-' * 25)
            writer.add_page(page)

        if add_blank_page:
            writer.add_blank_page()

    with open(merged_file_name, 'wb') as out:
        writer.write(out)

    print('\t> File: ' + merged_file_name + ' generated.')


def make_forms():
    iP_gb_3_0_STK() # OK
    STK_Form_PSP_10_0_en() # OK
    STK_Form_ISP_7_0_en3()  # OK
    STK_Form_ISPP_6_0_enver2() # OK
    Perfusor_compact() # OK
    Power_Supply() # OK
    infusomat_FMS() # OK
    infusomat_compact_plus_p() # OK
    infusomat_compact_plus_s() # OK
    perfusor_compact_plus() # OK
    Perfusor_compact_s() # OK


# Main function
if __name__ == '__main__':
    """
        pyinstaller -F -c test_pdf.py -n Form_generator
    """
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 2000)

    print('Awesome Forms generator v2.0')
    print('=' * 20)

    with open('./source/sign_names.txt', encoding='utf-8') as f:
        name_to_sign = json.loads(f.read())

    make_forms()

    print('\nAll forms were successfully generated.')
    #input('Press any key to close this window...')
