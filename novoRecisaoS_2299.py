import pandas as pd
from xml.dom import minidom
import os


def pegando_arquivos(arquivo_xml):
    # Lendo o arquivo XML
        dom = minidom.parse(arquivo_xml)
        # Inicializando listas para armazenar os dados
        cpfs = []
        insCs = []
        matriculas = []
        mtvDesligs = []
        datDesligs = []
        indPagtoAPIs = []
        codrubs = []
        valoresRubs = []
        arq = []


        # Obtendo os elementos das tags
        cpf_elements = dom.getElementsByTagName('cpfTrab')
        matriculas_elements = dom.getElementsByTagName('matricula')[0].firstChild.data
        mtvDeslig_elements = dom.getElementsByTagName('mtvDeslig')[0].firstChild.data 
        datDesligs_elements = dom.getElementsByTagName('dtDeslig')[0].firstChild.data
        indPagtoAPIs_elements= dom.getElementsByTagName('indPagtoAPI')[0].firstChild.data
        codrub_elements = dom.getElementsByTagName('codRubr')
        valoresRubs_elements = dom.getElementsByTagName('vrRubr')


        # Iterando pelas tags no XML
        for cpf_elem in cpf_elements:
            cpf = cpf_elem.firstChild.nodeValue.strip()  # Valor da tag cpfTrab
            if not codrub_elements:
                cpfs.append(cpf)
                insCs = dom.getElementsByTagName('nrInsc')[0].firstChild.data
                matriculas.append(matriculas_elements)
                mtvDesligs.append(mtvDeslig_elements) 
                datDesligs.append(datDesligs_elements)
                indPagtoAPIs.append(indPagtoAPIs_elements)
                codrubs.append(0)
                valoresRubs.append(0)
                arq.append(arquivo_xml[61:])
            else:
                for codrub_elem in codrub_elements:
                    codrub = codrub_elem.firstChild.nodeValue.strip()  # Valor da tag codRub
                    cpfs.append(cpf)
                    insCs = dom.getElementsByTagName('nrInsc')[0].firstChild.data
                    matriculas.append(matriculas_elements)
                    mtvDesligs.append(mtvDeslig_elements) 
                    datDesligs.append(datDesligs_elements)
                    indPagtoAPIs.append(indPagtoAPIs_elements)
                    codrubs.append(codrub)
                    arq.append(arquivo_xml[61:])
                for valores_elem in valoresRubs_elements:
                    valorRubrica = valores_elem.firstChild.nodeValue.strip() #valor da tag valor rubrica
                    valoresRubs.append(valorRubrica)
   
        
        return cpfs,insCs,matriculas,mtvDesligs, datDesligs,indPagtoAPIs, codrubs,valoresRubs,arq

def main(folder_path):
    # Lista para armazenar os dados

    data = []

    # Percorrer todos os arquivos XML na pasta
    for filename in os.listdir(folder_path):
        if filename.lower().endswith("s-2299.xml"):
            arq_xml = os.path.join(folder_path, filename)
            cpfs,insCs,matriculas,mtvDesligs, datDesligs,indPagtoAPIs, codrubs,valoresRubs,arq = pegando_arquivos(arq_xml)
            if cpfs and insCs and matriculas and mtvDesligs and datDesligs and indPagtoAPIs and codrubs and valoresRubs and arq:
                data.append({'cpf': cpfs,'insCs':insCs,'matriculas':matriculas,
                             'mtvDesligs':mtvDesligs, 'datDesligs':datDesligs,'indPagtoAPIs':indPagtoAPIs,
                             'codRub': codrubs, 'valorRub': valoresRubs,'arq':arq})
        else:
            pass
            
    if len(data) == 0:
        pass
    else:
        print("arquivo salvo")
        df = pd.DataFrame(data)
    # Salvar em um arquivo Excel (substitua 'dados.xml' pelo nome desejado)
        df.to_excel('Recisao S-2299.xlsx', index=False)
  




def formatacao(arquivo_Excel):
    
    try:
        # Nomes das colunas que contêm as listas
        cpfs = 'cpf'
        matriculas = 'matriculas'
        mtvDesligs = 'mtvDesligs'
        datDesligs = 'datDesligs'
        indPagtoAPIs = 'indPagtoAPIs'
        codRub = 'codRub'
        valorRub = 'valorRub'
        arq = 'arq'
        
        # Carrega a planilha Excel
        df = pd.read_excel(arquivo_Excel)


        # Lista para armazenar os novos dados
        novos_dados_S_2299 = []

        # Itera sobre cada linha do DataFrame
        for index, row in df.iterrows():
            # Converte as strings das listas em listas reais
            lista_valores1 = eval(row[cpfs])
            lista_valores2 = eval(row[codRub])
            lista_valores3 = eval(row[valorRub])
            lista_valores4 = eval(row[matriculas])
            lista_valores5 = eval(row[mtvDesligs])
            lista_valores6 = eval(row[datDesligs])
            lista_valores7 = eval(row[indPagtoAPIs]) 
            lista_valores8 = eval(row[arq]) 
                  

            
            # Garante que ambas as listas tenham o mesmo tamanho
            max_len = max(len(lista_valores1), len(lista_valores2),len(lista_valores3),len(lista_valores4),len(lista_valores5),len(lista_valores6),len(lista_valores7),len(lista_valores8))
            lista_valores1.extend([None] * (max_len - len(lista_valores1)))
            lista_valores2.extend([None] * (max_len - len(lista_valores2)))
            lista_valores3.extend([None] * (max_len - len(lista_valores3)))
            lista_valores4.extend([None] * (max_len - len(lista_valores4)))
            lista_valores5.extend([None] * (max_len - len(lista_valores5)))
            lista_valores6.extend([None] * (max_len - len(lista_valores6)))
            lista_valores7.extend([None] * (max_len - len(lista_valores7)))
            lista_valores8.extend([None] * (max_len - len(lista_valores8)))

            for valor1, valor2,valor3,valor4,valor5,valor6,valor7,valor8 in zip(lista_valores1, lista_valores2,lista_valores3,lista_valores4,lista_valores5,lista_valores6,lista_valores7,lista_valores8):
                nova_linha = row.copy()
                nova_linha[cpfs] = valor1
                nova_linha[codRub] = valor2
                nova_linha[valorRub] = valor3
                nova_linha[matriculas] = valor4
                nova_linha[mtvDesligs] = valor5
                nova_linha[datDesligs] = valor6
                nova_linha[indPagtoAPIs] = valor7
                nova_linha[arq] = valor8
                novos_dados_S_2299.append(nova_linha)
                
        # Cria um novo DataFrame com os novos dados
        novo_df_S_2299 = pd.DataFrame(novos_dados_S_2299)

        # Salva o novo DataFrame em uma nova planilha Excel
        novo_df_S_2299.to_excel(f'{arquivo_Excel}', index=False)
        print("Planilha Recisao S-2299 salva com sucesso!")
    except:
        print("Não é o arquivo S-2299")
        pass


