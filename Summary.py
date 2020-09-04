import os
import xlrd
import xlwt

def isNum(num):
    try:
        a = float(num)
        return round(a,2)
    except ValueError:
        return '-'


print('Start...')
loc = os.getcwd()

wb = xlwt.Workbook()
proj = loc[(loc.find('\\',12)+1):(loc.find('\\',12)+1+6)]
print(proj)
path = loc + '\Summary ' + proj + '.xls'

style_text_wrap_font_bold_black_color = xlwt.Style.easyxf('font: bold on, color-index black; align: vert centre, horiz center')
style_centre = xlwt.Style.easyxf('align: vert centre, horiz center')
col_width = 128*30


#Classification
dataClassc = []
for file in os.listdir(loc):
    if file.endswith("Classification.xlsm"):
        workbook = xlrd.open_workbook(os.path.join(loc, file))
        sheet = workbook.sheet_by_name("Report")
        nameFile = file[(file.find('-'))+1:-5]
        a = {'id': nameFile,
             'Unit Weight': isNum(str(sheet.cell_value(12, 7))),
             'Gs': isNum(str(sheet.cell_value(14, 7))),
             'MC': isNum(str(sheet.cell_value(15, 7))),
             'PL': isNum(str(sheet.cell_value(16, 7))),
             'LL': isNum(str(sheet.cell_value(17, 7))),
             'PI': isNum(str(sheet.cell_value(18, 7))),
             'Fines(#200)': isNum(str(sheet.cell_value(19, 7)))}
        dataClassc.append(a)

ws2 = wb.add_sheet('Classification')

ws2.write(0,0, 'Borehole', style_text_wrap_font_bold_black_color)
ws2.write(0,1, 'Unit Weight (kN/m3)', style_text_wrap_font_bold_black_color)
ws2.write(0,2, 'Gs', style_text_wrap_font_bold_black_color)
ws2.write(0,3, 'MC (%)', style_text_wrap_font_bold_black_color)
ws2.write(0,4, 'PL (%)', style_text_wrap_font_bold_black_color)
ws2.write(0,5, 'LL (%)', style_text_wrap_font_bold_black_color)
ws2.write(0,6, 'PI (%)', style_text_wrap_font_bold_black_color)
ws2.write(0,7, 'Fines(#200) (%)', style_text_wrap_font_bold_black_color)

ws2.col(0).width = 256 * (len(loc) + 10)
for i in range(1, 8):
    ws2.col(i).width = col_width

index = 1
for i in dataClassc:
    ws2.write(index+1,0,i['id'])
    ws2.write(index+1,1,i['Unit Weight'], style_centre)
    ws2.write(index+1,2,i['Gs'], style_centre)
    ws2.write(index+1,3,i['MC'], style_centre)
    ws2.write(index+1,4,i['PL'], style_centre)
    ws2.write(index+1,5,i['LL'], style_centre)
    ws2.write(index+1,6,i['PI'], style_centre)
    ws2.write(index+1,7,i['Fines(#200)'], style_centre)
    index+=1



#Triaxial UU
dataTxUU = []
for file in os.listdir(loc):
    if file.endswith("TX UU.xlsm"):
        workbook = xlrd.open_workbook(os.path.join(loc, file))
        sheet = workbook.sheet_by_index(2)
        nameFile = file[(file.find('-'))+1:-5]
        a = {'id': nameFile,
             'Cohession': str(round(sheet.cell_value(18, 16),2)),
             'Phi' : str(round(sheet.cell_value(19, 16), 2))}
        dataTxUU.append(a)

ws = wb.add_sheet('TX UU')

ws.write(0,0, 'Borehole', style_text_wrap_font_bold_black_color)
ws.write(0,1, 'Cohession (kPa)', style_text_wrap_font_bold_black_color)
ws.write(0,2, 'Phi (degree)', style_text_wrap_font_bold_black_color)

ws.col(0).width = 256 * (len(loc) + 10)
ws.col(1).width = col_width
ws.col(2).width = col_width

index = 1
for i in dataTxUU:
    ws.write(index+1,0,i['id'])
    ws.write(index+1,1,i['Cohession'])
    ws.write(index+1,2,i['Phi'])
    index+=1


#Consolidation
dataConsol = []
for file in os.listdir(loc):
    if file.endswith("Consolidation.xlsm"):
        workbook = xlrd.open_workbook(os.path.join(loc, file))
        sheet = workbook.sheet_by_name("Report")
        nameFile = file[(file.find('-'))+1:-5]
        a = {'id': nameFile,
             'eo': str(round(sheet.cell_value(52, 7),2)),
             'Cc': str(round(sheet.cell_value(53, 7),2)),
             'Cr': str(round(sheet.cell_value(54, 7),2)),
             'Cs': str(round(sheet.cell_value(55, 7),2)),
             'Pc': str(round(sheet.cell_value(56, 7),2))}
        dataConsol.append(a)

ws3 = wb.add_sheet('Consolidation')

ws3.write(0,0, 'Borehole', style_text_wrap_font_bold_black_color)
ws3.write(0,1, 'eo', style_text_wrap_font_bold_black_color)
ws3.write(0,2, 'Cc', style_text_wrap_font_bold_black_color)
ws3.write(0,3, 'Cr', style_text_wrap_font_bold_black_color)
ws3.write(0,4, 'Cs', style_text_wrap_font_bold_black_color)
ws3.write(0,5, 'Pc (kPa)', style_text_wrap_font_bold_black_color)

ws3.col(0).width = 256 * (len(loc) + 10)
for i in range(1, 6):
    ws3.col(i).width = col_width

index = 1
for i in dataConsol:
    ws3.write(index+1,0,i['id'])
    ws3.write(index+1,1,i['eo'], style_centre)
    ws3.write(index+1,2,i['Cc'], style_centre)
    ws3.write(index+1,3,i['Cr'], style_centre)
    ws3.write(index+1,4,i['Cs'], style_centre)
    ws3.write(index+1,5,i['Pc'], style_centre)
    index+=1


#UCS poisson
dataUcsPois = []
for file in os.listdir(loc):
    if file.endswith("UCS (Poisson).xlsm"):
        workbook = xlrd.open_workbook(os.path.join(loc, file))
        sheet = workbook.sheet_by_name("Report")
        nameFile = file[(file.find('-'))+1:-5]
        a = {'id': nameFile,
             'SigmaMax': str(round(sheet.cell_value(12, 6),2)),
             'v': str(round(sheet.cell_value(13, 6),2)),
             'e': str(round(sheet.cell_value(14, 6),2))}
        dataUcsPois.append(a)

ws4 = wb.add_sheet('UCS (Poisson)')

ws4.write(0,0, 'Borehole', style_text_wrap_font_bold_black_color)
ws4.write(0,1, 'Sigma Max (MPa)', style_text_wrap_font_bold_black_color)
ws4.write(0,2, 'v', style_text_wrap_font_bold_black_color)
ws4.write(0,3, 'E (MPa)', style_text_wrap_font_bold_black_color)

ws4.col(0).width = 256 * (len(loc) + 10)
for i in range(1, 3):
    ws4.col(i).width = col_width

index = 1
for i in dataUcsPois:
    ws4.write(index+1,0,i['id'])
    ws4.write(index+1,1,i['SigmaMax'], style_centre)
    ws4.write(index+1,2,i['v'], style_centre)
    ws4.write(index+1,3,i['e'], style_centre)
    index+=1


#UCT
dataUct = []
for file in os.listdir(loc):
    if file.endswith("UCT.xlsm"):
        workbook = xlrd.open_workbook(os.path.join(loc, file))
        sheet = workbook.sheet_by_name("Report")
        nameFile = file[(file.find('-'))+1:-5]
        a = {'id': nameFile,
             'st': str(round(sheet.cell_value(14, 15),2)),             
             'cohession': str(round(sheet.cell_value(15, 15),2)),
             'ei': str(round(sheet.cell_value(16, 15),2))}
        dataUct.append(a)

ws5 = wb.add_sheet('UCT')

ws5.write(0,0, 'Borehole', style_text_wrap_font_bold_black_color)
ws5.write(0,1, 'Cohession (kPa)', style_text_wrap_font_bold_black_color)
ws5.write(0,2, 'Sensitivity (-)', style_text_wrap_font_bold_black_color)
ws5.write(0,3, 'Ei (kPa)', style_text_wrap_font_bold_black_color)

ws5.col(0).width = 256 * (len(loc) + 10)
for i in range(1, 4):
    ws5.col(i).width = col_width

index = 1
for i in dataUct:
    ws5.write(index+1,0,i['id'])
    ws5.write(index+1,1,i['cohession'], style_centre)
    ws5.write(index+1,2,i['st'], style_centre)
    ws5.write(index+1,3,i['ei'], style_centre)
    index+=1


#Triaxial CU
dataCu = []
for file in os.listdir(loc):
    if file.endswith("TX CU.xlsm"):
        workbook = xlrd.open_workbook(os.path.join(loc, file))
        sheet = workbook.sheet_by_name("Report")
        nameFile = file[(file.find('-'))+1:-5]
        a = {'id': nameFile,
             'ctotal': str(round(sheet.cell_value(18, 14),2)),
             'phitotal': str(round(sheet.cell_value(19, 14),2)),
             'ceffective': str(round(sheet.cell_value(18, 16),2)),
             'phieffective': str(round(sheet.cell_value(19, 16),2))}
        dataCu.append(a)

ws6 = wb.add_sheet('Triaxial CU')

ws6.write(0,0, 'Borehole', style_text_wrap_font_bold_black_color)
ws6.write(0,1, 'c (kPa)', style_text_wrap_font_bold_black_color)
ws6.write(0,2, 'phi (degree)', style_text_wrap_font_bold_black_color)
ws6.write(0,3, "c' (kPa)", style_text_wrap_font_bold_black_color)
ws6.write(0,4, "phi' (degree)", style_text_wrap_font_bold_black_color)

ws6.col(0).width = 256 * (len(loc) + 10)
for i in range(1, 5):
    ws6.col(i).width = col_width

index = 1
for i in dataCu:
    ws6.write(index+1,0,i['id'])
    ws6.write(index+1,1,i['ctotal'], style_centre)
    ws6.write(index+1,2,i['phitotal'], style_centre)
    ws6.write(index+1,3,i['ceffective'], style_centre)
    ws6.write(index+1,4,i['phieffective'], style_centre)
    index+=1

print('Finish...')
wb.save(path)

