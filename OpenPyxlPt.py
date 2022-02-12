from dataclasses import dataclass
from sqlite3 import Row
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
#from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.reader.excel import load_workbook
from openpyxl.styles.borders import Border, Side


import os
import subprocess
import sys


BordaPadrao = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))

def DevolveNomeCelula(ws, L, I):
    return ws.cell(row=L, column=I).coordinate

def DevolveLetraColuna(ws, L, I)    :
    return ws.cell(row=L, column=I).column_letter

def AjustaLarguraColunas(ws, PrimeiraLinha = 1):
    thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))
    TamanhoMinimo = 11
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter # Get the column name
        for cell in col:
            cell.border = thin_border
            if cell.row >= PrimeiraLinha:
                #if column == "G":
                #    print(len(str(cell.value)))

                try: # Necessary to avoid error on empty cells
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass



                adjusted_width = (max_length + 3) * 1.1

                ws.column_dimensions[column].width = adjusted_width



def AbrirArquivo(NomeArquivo):
    if (os.path.isfile(NomeArquivo + '.xlsx')):
        return load_workbook(NomeArquivo + '.xlsx')
    else:
        Arquivo = Workbook()
        Arquivoxlsx = NomeArquivo + '.xlsx'
        Arquivo.save(Arquivoxlsx)
        return Arquivo

def Fechar(Arquivo):
    Arquivo.close()

def ConverteXlxsParaOds(NomeArquivo):
    #$ /usr/bin/libreoffice --headless --invisible -convert-to ods /home/cwgem/Downloads/QTL_Sample_data.xls
    #convert /home/cwgem/Downloads/QTL_Sample_data.xls -> /home/cwgem/QTL_Sample_data.ods using OpenDocument Spreadsheet Flat XML
    #$ /usr/bin/libreoffice --headless --invisible -convert-to xls /home/cwgem/QTL_Sample_data.ods
    #convert /home/cwgem/QTL_Sample_data.ods -> /home/cwgem/QTL_Sample_data.xls using
    #os.system('soffice --headless --invisible -convert-to ods ' + str(NomeArquivo))

    output = os.popen('soffice --headless --invisible --convert-to ods "' + str(NomeArquivo)).read() +'"'
    #output = os.popen('soffice --convert-to ods ' + str(NomeArquivo) + ' --headless')
    print(output)
    #subprocess.call(["soffice", "--headless", "--invisible", "-convert-to", "ods", NomeArquivo])

def Salvar(Arquivo, NomeArquivo, ConverteParaOds = True):
    Arquivoxlsx = NomeArquivo + '.xlsx'
    #Arquivo.save(Arquivoxlsx)

    if "Sheet" in Arquivo.sheetnames: #caso exista uma aba sem nome(inicio do arquiv) deleta
        Arquivo.remove(Arquivo['Sheet'])

    Arquivo.save(Arquivoxlsx)
    if (ConverteParaOds):
        ConverteXlxsParaOds(Arquivoxlsx)

def AdicionaCabecalho(ws, Campos = [], linha = 1, Alinhamento = "", CorFundo = ""):
    Alfabeto = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
    font = Font(name='Arial',
                size=10,
                 bold=True)

    thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))


    I = 0
    for label in Campos:
        ws[Alfabeto[I] + str(linha)] = label
        ws[Alfabeto[I] + str(linha)].font = font
        if (Alinhamento != ""):
            Alinhar(ws, Alfabeto[I] + str(linha), Alinhamento)
        #ws[Alfabeto[I] + '1'].Border = thin_border
        ws[Alfabeto[I] + str(linha)].border = thin_border
        if (CorFundo != ""):
            ws[Alfabeto[I] + str(linha)].fill = PatternFill("solid", start_color=CorFundo)


        I += 1
    ws.append([''])

def AdicionaRegistro(ws, Campos = [], Ordenacao = [], DataTypes = [], Borda = False):
    I = 1
    L = DevolveUltimaLinha(ws) + 1

    for Conteudo in Campos:
        Celula = ws.cell(row=L, column=I)

        if(len(DataTypes) > 0):
            if DataTypes[I-1] == 'numeric':
                if ConverteNumerico(Conteudo):
                    Conteudo = '{0:.2f}'.format(float(Conteudo))

                    Conteudo = float(Conteudo)#str(Conteudo).replace(".", ",")
                if Conteudo == '0':
                    Conteudo = float('0.00')

                Celula.number_format = '#,##0.00'
            elif DataTypes[I-1] == "string":
                Conteudo = str(Conteudo).strip()


        Celula.value = Conteudo

        if(len(Ordenacao) > 0):
            NomeCelula = str(Celula.coordinate)
            Alinhar(ws,NomeCelula,Ordenacao[I-1], 'center')

        if (Borda):
            Celula.border = BordaPadrao

        I += 1
    #ws.append(Campos)
    #ws.cell(row=4, column=9).value = 2
    #sys.exit

def LimpaAba(ws):
    ws.delete_cols(1, 100)
    ws.delete_rows(1, 300)

def FormataCelula(ws, cell, formato = ""):
    ws[cell].number_format = formato

def FormataPorcentagem(ws, cell):
    ws[cell].number_format = '0.00%'

def MesclarCelulas(ws, cells):
    ws.merge_cells(cells)

def Alinhar(ws, cell, horizontal='center', vertical='center'):
    currentCell = ws[cell]
    currentCell.alignment = Alignment(horizontal=horizontal, vertical=vertical)

def DevolveUltimaLinha(ws, Coluna = ""):
    if Coluna == "":
        return ws.max_row
    else:
        return len(ws[Coluna])

def UltimaLinha(ws):
    return ws.max_row

def UltimaColuna(ws):
    return ws.max_column

def ConverteNumerico(Valor):
    try:
        return float(Valor)
    except:
        return False
        pass

def AjustaLarguraColuna(ws, Coluna, Tamanho = 10):
    ws.column_dimensions[Coluna].width = Tamanho

def AdicionaRegistroUnico(ws, Linha = 1, Coluna = 1, Conteudo = '', Negrito = False, SobrePor = True, CorFundo = "", Borda=False):
    '''
        Caso utulizado para formula, deve ser utilizado partindo do principio que está sendo escrito um excel
        Logo, SOMA do BROffice deve ser escrito como SUM que é a formula do excel
    '''
    Celula = ws.cell(row=Linha, column=Coluna)
    #print(Celula.value, Conteudo)
    if Celula.value != None:
        if SobrePor:
            Celula.value = Conteudo
    else:
        Celula.value = Conteudo

    if (CorFundo != ""):
        Celula.fill = PatternFill("solid", start_color=CorFundo)

    if (Negrito):
        Celula.font = Font(name='Arial', size=10, bold=True)
    if (Borda):
        Celula.border = BordaPadrao

def AdicionaBorda(ws, Linha = 1, Coluna = 1):
    ws.cell(row=Linha, column=Coluna).border = BordaPadrao
    #ws[cell].border = BordaPadrao

def CriaAba(Arquivo, NomeAba = "", Sobrescrever = True):
    if NomeAba in Arquivo.sheetnames:
        ws = Arquivo[NomeAba]
        if(Sobrescrever):
            LimpaAba(ws)
    else:
        ws = Arquivo.create_sheet(NomeAba)

    return ws