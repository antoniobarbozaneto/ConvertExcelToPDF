import sys
import os

import win32com.client

if __name__ == '__main__':

    arquivo_caminho = str(sys.argv[1])

    if arquivo_caminho == "help":
        print("     Conversor Excel para PDF v1.0"
              "\n \n Execute o programa ConvertXlsxToPDF.py passando como parâmetro o caminho completo, "
              "\n incluindo o nome do arquivo Excel e sua extensão que deseja converter para PDF, "
              "\n o arquivo de saída será salvo com o mesmo nome na mesma pasta do arquivo de entrada."
              "\n \n Exemplo:"
              '\n ConvertXlsxToPDF.py C:/ThomsonReuters/automations/Excel_To_PDF/RelatorioExcel.xls'
              "\n \n Saída:"
               '\n C:/ThomsonReuters/automations/Excel_To_PDF/RelatorioExcel.pdf'
              "\n")
        os.system("pause")

    #Caminho do arquivo Excel
    #WB_PATH = r'{caminho_completo_arquivo}'
    WB_PATH = r'' + arquivo_caminho + ''
    nome_arquivo = WB_PATH.split('/')[-1] #nome do arquivo com extensão
    index = nome_arquivo.index('.') #seta o ponto como index
    nome_arquivo = nome_arquivo[:index] #remove a extensão
    # Caminho para salvar o PDF
    PATH_TO_PDF = r'' + arquivo_caminho + ''+nome_arquivo+'.pdf'
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    print ("nome arquivo:" + nome_arquivo)
    print('Start conversion to PDF')
    # Abre o Excel
    wb = excel.Workbooks.Open(WB_PATH)
    # Especificar o sheet que deseja converter, pode ser um array ex: [1, 2, 3, 4, ...]
    ws_index_list = [1]
    wb.WorkSheets(ws_index_list).Select()
    # salva
    wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
    print('Arquivo Gerado com sucesso no caminho: '+PATH_TO_PDF)
    wb.Close()
    excel.Quit()

#python3 -m pip install pywin32