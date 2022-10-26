# coding=utf-8
# pzw
# 20211028

import WordWriter_beta as wwb

list_col=['基因','外显子','核苷酸变异','氨基酸变异','丰度']

resultsDict = {}
resultsDict["#[TABLE-medic_7col]#"] = "用药相关位点.tsv"
resultsDict["#[TABLE-medic2_5col]#"] = "用药无关位点.tsv"
docx=wwb.WordWriter("报告模板.docx" ,resultsDict)

def mergeFirstColumnCell(cell_main, cell):
    text=cell_main.text
    cell_main.merge(cell)
    cell_main.text=text

global tmp_cell_hrzt
Horizontal_dic={}
for table in docx.tables:
    for row in table.rows:
        for cell in row.cells:
            if cell.text != '#<#':
                tmp_cell_hrzt=cell
            else :
                Horizontal_dic.update({tmp_cell_hrzt:cell})

for k,v in Horizontal_dic.items():
    mergeFirstColumnCell(k,v)


global tmp_cell_vtc
vertical_dic = {}
for table in docx.tables:
    for column in table.columns:
        tmp_cell_vtc = '' #每读一列重置一次
        if column.cells[0].text in list_col:
            
            for cell in column.cells:
                if tmp_cell_vtc == ''  or cell.text != tmp_cell_vtc.text:
                    tmp_cell_vtc=cell
                else :
                    vertical_dic.update({tmp_cell_vtc:cell})
for k,v in vertical_dic.items():
    mergeFirstColumnCell(k,v)

#text=[[[cell.text for cell in row.cells ] for row in table.columns] for table in docx.tables]

docx.save('merge_new.docx')

