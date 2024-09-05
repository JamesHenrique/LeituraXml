
# formatacaoS1010(caminho)
import pandas as pd
import os
from xml.dom import minidom




def pegando_arquivos(arquivo_xml):
    # Lendo o arquivo XML
        dom = minidom.parse(arquivo_xml)
        # Inicializando listas para armazenar os dados
        arq = []
        cpfs = []
        dtAlteracao = []
        nmTrab = []
        estCiv = []
        nmDeps = []
        datNascimentoDep = []
        depIRRF = []
        tpDepTodos =[]
        
        
        
        cpfs_elements = dom.getElementsByTagName('cpfTrab')[0].firstChild.data
        dtAlteracao_elements = dom.getElementsByTagName('dtAlteracao')[0].firstChild.data
        nmTrab_elements = dom.getElementsByTagName('nmTrab')[0].firstChild.data 
        nmDep_elements = dom.getElementsByTagName('nmDep')
        estado_elements = dom.getElementsByTagName('estCiv')[0].firstChild.data
        nascimento = dom.getElementsByTagName('nascimento')
        
    
        if not nmDep_elements:
            cpfs.append(cpfs_elements)
            dtAlteracao.append(dtAlteracao_elements)
            nmTrab.append(nmTrab_elements)
            estCiv.append(estado_elements)
            nmDeps.append('')
            depIRRF.append('')
            tpDepTodos.append('')
            datNascimentoDep.append('')
            arq.append(arquivo_xml[61:])
        else:
            for tag in nmDep_elements:
                cpfs.append(cpfs_elements)
                dtAlteracao.append(dtAlteracao_elements)
                nmTrab.append(nmTrab_elements)
                estCiv.append(estado_elements)
                nomes = tag.firstChild.nodeValue.strip() 
                nmDeps.append(nomes)
                arq.append(arquivo_xml[61:])
                
            depIRRF_elements = dom.getElementsByTagName('depIRRF')
            for tag in depIRRF_elements:
                depIRRF.append(tag.firstChild.nodeValue)
                
            tpDepTodos_elements = dom.getElementsByTagName('tpDep')
            for tag in tpDepTodos_elements:
                tpDepTodos.append(tag.firstChild.nodeValue)
             
            if not nascimento:  
                datNascimentoDep_elem = dom.getElementsByTagName('dtNascto')
                for tag in datNascimentoDep_elem:
                    datNascimentoDep.append(tag.firstChild.nodeValue)
            else: 
                datNascimentoDep_elem = dom.getElementsByTagName('dtNascto')[1:]
                for tag in datNascimentoDep_elem:
                    datNascimentoDep.append(tag.firstChild.nodeValue)
                
        return cpfs,dtAlteracao,nmTrab,estCiv,nmDeps,depIRRF,tpDepTodos,datNascimentoDep,arq

def mainS_2205(folder_path):
    # Lista para armazenar os dados
    data = []

    # Percorrer todos os arquivos XML na pasta
    for filename in os.listdir(folder_path):
        if filename.lower().endswith("s-2205.xml"):
            arq_xml = os.path.join(folder_path, filename)
            cpfs,dtAlteracao,nmTrab,estCiv,nmDeps,depIRRF,tpDepTodos, datNascimentoDep,arq = pegando_arquivos(arq_xml) 
            # ,depIRRF,tpDepTodos,datNascimentoDep
            if cpfs and dtAlteracao and nmTrab and estCiv and nmDeps and depIRRF and tpDepTodos and datNascimentoDep and arq:
                # and depIRRF and tpDepTodos and datNascimentoDep:
                data.append({'cpf': cpfs,
                             'dtAlteracao':dtAlteracao,
                             'nmTrab':nmTrab,
                             'estCiv':estCiv,
                             'nmDeps':nmDeps,
                             'depIRRF':depIRRF,
                             'tpDepTodos':tpDepTodos,
                             'datNascimentoDep':datNascimentoDep,
                             'arq' : arq
            
                             })
        else:
            pass


    if len(data) == 0:
        pass   
    else:   
        df = pd.DataFrame(data)
        df.to_excel('Inclusao Dependente S-2205.xlsx', index=False)
        




def formatacaoS_2205(arquivo_excel):
    # Nomes das colunas que contêm as listas
    cpfs = 'cpf'
    dtAlteracao = 'dtAlteracao'
    nmTrab = 'nmTrab'
    estCiv = 'estCiv'
    nmDeps = 'nmDeps'
    depIRRF = 'depIRRF'
    tpDepTodos = 'tpDepTodos'
    datNascimentoDep = 'datNascimentoDep'
    arq = 'arq'

    # Carrega a planilha Excel
    df = pd.read_excel(arquivo_excel)

    # Lista para armazenar os novos dados
    novos_dados_S_1200 = []

    # Itera sobre cada linha do DataFrame
    for index, row in df.iterrows():
        # Converte as strings das listas em listas reais
        lista_valores1 = eval(row[cpfs])
        lista_valores2 = eval(row[dtAlteracao])
        lista_valores3 = eval(row[nmTrab])
        lista_valores4 = eval(row[estCiv])
        lista_valores5 = eval(row[nmDeps])

        lista_valores6 = eval(row[depIRRF])
        lista_valores7 = eval(row[tpDepTodos])
        lista_valores8 = eval(row[datNascimentoDep])
        lista_valores9 = eval(row[arq])
        
        
        

        
        # Garante que ambas as listas tenham o mesmo tamanho
        max_len = max(len(lista_valores1), len(lista_valores2),len(lista_valores3),len(lista_valores4),len(lista_valores5),len(lista_valores6),len(lista_valores7),len(lista_valores8),len(lista_valores9))
        # ,len(lista_valores6),len(lista_valores7),len(lista_valores8)
        lista_valores1.extend([None] * (max_len - len(lista_valores1)))
        lista_valores2.extend([None] * (max_len - len(lista_valores2)))
        lista_valores3.extend([None] * (max_len - len(lista_valores3)))
        lista_valores4.extend([None] * (max_len - len(lista_valores4)))
        lista_valores5.extend([None] * (max_len - len(lista_valores5)))
        lista_valores6.extend([None] * (max_len - len(lista_valores6)))
        lista_valores7.extend([None] * (max_len - len(lista_valores7)))
        lista_valores8.extend([None] * (max_len - len(lista_valores8)))
        lista_valores8.extend([None] * (max_len - len(lista_valores9)))
        

        # ,valor6,valor7,valor8
        for valor1, valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9 in zip(lista_valores1, lista_valores2,lista_valores3,lista_valores4,lista_valores5,lista_valores6,lista_valores7,lista_valores8,lista_valores9):
            # ,lista_valores6,lista_valores7,lista_valores8
            nova_linha = row.copy()
            nova_linha[cpfs] = valor1
            nova_linha[dtAlteracao] = valor2
            nova_linha[nmTrab] = valor3
            nova_linha[estCiv] = valor4
            nova_linha[nmDeps] = valor5
        
            
            nova_linha[depIRRF] = valor6
            
            nova_linha[tpDepTodos] = valor7
            nova_linha[datNascimentoDep] = valor8
            nova_linha[arq] = valor9
            novos_dados_S_1200.append(nova_linha)
            
            

    # Cria um novo DataFrame com os novos dados
    novo_df_S_1200 = pd.DataFrame(novos_dados_S_1200)

    # Salva o novo DataFrame em uma nova planilha Excel
    novo_df_S_1200.to_excel(f'{arquivo_excel}', index=False)
    print("Planilha Inclusao Dependente S-2205 salva com sucesso!")








 