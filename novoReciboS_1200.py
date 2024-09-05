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
        competencias = []
        cnpjs = []
        codrubs = []
        valoresRubs = []
        arq = []


        # Obtendo os elementos das tags
        cpf_elements = dom.getElementsByTagName('cpfTrab')
        competentcia_elements = dom.getElementsByTagName('perApur')[0].firstChild.data
        cnpj_elements = dom.getElementsByTagName('nrInsc')[1].firstChild.data 
        matriculas_elements = dom.getElementsByTagName('nrInsc')[0].firstChild.data
        codrub_elements = dom.getElementsByTagName('codRubr')
        valoresRubs_elements = dom.getElementsByTagName('vrRubr')
        


        # Iterando pelas tags no XML
        for cpf_elem in cpf_elements:
            cpf = cpf_elem.firstChild.nodeValue.strip()  # Valor da tag cpfTrab
            if not codrub_elements:
                cpfs.append(cpf)
                insCs = dom.getElementsByTagName('nrInsc')[0].firstChild.data
                matriculas.append(matriculas_elements)
                competencias.append(competentcia_elements) 
                cnpjs.append(cnpj_elements)
                codrubs.append('')
                valoresRubs.append('')
                arq.append(arquivo_xml[61:])
            else:
                for codrub_elem in codrub_elements:
                    codrub = codrub_elem.firstChild.nodeValue.strip()  # Valor da tag codRub
                    cpfs.append(cpf)
                    insCs = dom.getElementsByTagName('nrInsc')[0].firstChild.data
                    matriculas.append(matriculas_elements)
                    competencias.append(competentcia_elements) 
                    cnpjs.append(cnpj_elements)
                    codrubs.append(codrub)
                    arq.append(arquivo_xml[61:])
                for valores_elem in valoresRubs_elements:
                    valorRubrica = valores_elem.firstChild.nodeValue.strip() #valor da tag valor rubrica
                    valoresRubs.append(valorRubrica)
   
        
        return cpfs,insCs,matriculas,competencias, cnpjs, codrubs,valoresRubs,arq

def mainS_1200(folder_path):
    # Lista para armazenar os dados
    data = []

    # Percorrer todos os arquivos XML na pasta
    for filename in os.listdir(folder_path):
        if filename.lower().endswith("s-1200.xml"):
            arq_xml = os.path.join(folder_path, filename)
            cpfs,insCs,matriculas,competencias, cnpjs, codrubs,valoresRubs,arq = pegando_arquivos(arq_xml)
            if cpfs and insCs and matriculas and competencias and cnpjs and  codrubs and valoresRubs and arq:
                data.append({'cpf': cpfs,'insCs':insCs,'matriculas':matriculas,
                             'competencias':competencias, 'cnpjs':cnpjs,
                             'codRub': codrubs, 'valorRub': valoresRubs, 'arq':arq})
        else:
            pass

    if len(data) == 0:
        pass   
    else:   
        df = pd.DataFrame(data)
        
        # Salvar em um arquivo Excel (substitua 'dados.xml' pelo nome desejado)
        df.to_excel('Recibos S-1200.xlsx', index=False)




def formatacaoS_1200(arquivo_excel):


    try:
        # Nomes das colunas que contêm as listas
        cpfs = 'cpf'
        matriculas = 'matriculas'
        competencias = 'competencias'
        cnpjs = 'cnpjs'
        codRub = 'codRub'
        valorRub = 'valorRub'
        arq = 'arq'

        # Carrega a planilha Excel
        df = pd.read_excel(arquivo_excel)

        # Lista para armazenar os novos dados
        novos_dados_S_1200 = []

        # Itera sobre cada linha do DataFrame
        for index, row in df.iterrows():
            # Converte as strings das listas em listas reais
            lista_valores1 = eval(row[cpfs])
            lista_valores2 = eval(row[codRub])
            lista_valores3 = eval(row[valorRub])
            lista_valores4 = eval(row[matriculas])
            lista_valores5 = eval(row[competencias])
            lista_valores6 = eval(row[cnpjs])
            lista_valores7 = eval(row[arq])
            
            

            
            # Garante que ambas as listas tenham o mesmo tamanho
            max_len = max(len(lista_valores1), len(lista_valores2),len(lista_valores3),len(lista_valores4),len(lista_valores5),len(lista_valores6),len(lista_valores7))
            lista_valores1.extend([None] * (max_len - len(lista_valores1)))
            lista_valores2.extend([None] * (max_len - len(lista_valores2)))
            lista_valores3.extend([None] * (max_len - len(lista_valores3)))
            lista_valores4.extend([None] * (max_len - len(lista_valores4)))
            lista_valores5.extend([None] * (max_len - len(lista_valores5)))
            lista_valores6.extend([None] * (max_len - len(lista_valores6)))
            lista_valores7.extend([None] * (max_len - len(lista_valores7)))

            for valor1, valor2,valor3,valor4,valor5,valor6,valor7 in zip(lista_valores1, lista_valores2,lista_valores3,lista_valores4,lista_valores5,lista_valores6,lista_valores7):
                nova_linha = row.copy()
                nova_linha[cpfs] = valor1
                nova_linha[codRub] = valor2
                nova_linha[valorRub] = valor3
                nova_linha[matriculas] = valor4
                nova_linha[competencias] = valor5
                nova_linha[cnpjs] = valor6
                nova_linha[arq] = valor7
                novos_dados_S_1200.append(nova_linha)
                
                

        # Cria um novo DataFrame com os novos dados
        novo_df_S_1200 = pd.DataFrame(novos_dados_S_1200)

        # Salva o novo DataFrame em uma nova planilha Excel
        novo_df_S_1200.to_excel(f'{arquivo_excel}', index=False)
        print("Planilha Recibo S-1200 salva com sucesso!")

    except:
        print("Não é o arquivo S-1200")
        pass
