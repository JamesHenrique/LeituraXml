
import pandas as pd
from xml.dom import minidom
import os



def pegando_arquivos(arquivo_xml):
    # Lendo o arquivo XML
        
        dom = minidom.parse(arquivo_xml)
      
        # Inicializando listas para armazenar os dados
        cpfs = []
        nomes = []
        sexos = []
        datas_nascimento = []
        dtAdmissao = []
        nome_dependentes = []
        cpf_Deps = []
        data_dep = []
        dep_IRRFs = []
        arq = []

        # Obtendo os elementos das tags
        cpf_elements = dom.getElementsByTagName('cpfTrab')[0].firstChild.data 
        nome_elements = dom.getElementsByTagName('nmTrab')[0].firstChild.data 
        sexo_elements = dom.getElementsByTagName('sexo')[0].firstChild.data
        data_nascimento_elements = dom.getElementsByTagName('dtNascto')[0].firstChild.data
        nome_dependentes_elements = dom.getElementsByTagName('nmDep')
 

        
        


        
        if not nome_dependentes_elements:
            try:
                cpfs.append(cpf_elements)
                nomes.append(nome_elements)
                sexos.append(sexo_elements)
                datas_nascimento.append(data_nascimento_elements) 
                dtAdm_elemets = dom.getElementsByTagName('dtAdm')[0].firstChild.data 
                dtAdmissao.append(dtAdm_elemets)
                nome_dependentes.append('')
                cpf_Deps.append('')
                data_dep.append('')
                dep_IRRFs.append('')
                arq.append(arquivo_xml[61:])
            except:
                dtAdm_elemets = dom.getElementsByTagName('dtAdm')[1].firstChild.data
                dtAdmissao.append(dtAdm_elemets)
        else:
            for nome_elem in nome_dependentes_elements:
                nome_dep = nome_elem.firstChild.nodeValue.strip() 
                cpfs.append(cpf_elements)
                nomes.append(nome_elements)
                sexos.append(sexo_elements)
                datas_nascimento.append(data_nascimento_elements) 
                
                try:
                    dtAdm_elemets = dom.getElementsByTagName('dtAdm')[0].firstChild.data
                    dtAdmissao.append(dtAdm_elemets)  
                except:
                    dtAdm_elemets = dom.getElementsByTagName('dtAdm')[1].firstChild.data
                    dtAdmissao.append(dtAdm_elemets)  
                
                nome_dependentes.append(nome_dep)   
                cpf_dep_elements = dom.getElementsByTagName('cpfDep') 
                arq.append(arquivo_xml[61:])    
                
            if not cpf_dep_elements:
                cpf_Deps.append('')
            else:
                for tag in cpf_dep_elements:
                    cpf_Deps.append(tag.firstChild.nodeValue)
               
            
                
            data_dep_elements = dom.getElementsByTagName('dtNascto')[1:]
            for tag in data_dep_elements:
                data_dep.append(tag.firstChild.nodeValue)

            IRRFs_elements = dom.getElementsByTagName('depIRRF')        
            for tag in IRRFs_elements:
                dep_IRRFs.append(tag.firstChild.nodeValue)
            
        
        return cpfs,nomes,sexos,datas_nascimento,dtAdmissao,nome_dependentes, cpf_Deps,data_dep,dep_IRRFs,arq


def mainS_2200(folder_path):
    # Lista para armazenar os dados
    data = []

    # Percorrer todos os arquivos XML na pasta
    for filename in os.listdir(folder_path):
        if filename.lower().endswith("s-2200.xml"):
            arq_xml = os.path.join(folder_path, filename)
            cpfs,nomes,sexos,datas_nascimento,dtAdmissao,nome_dependentes, cpf_Deps,data_dep,dep_IRRFs,arq = pegando_arquivos(arq_xml) #,nome_dependentes,,data_dep,dep_IRRFs
            if cpfs and nomes and sexos and datas_nascimento and dtAdmissao and nome_dependentes and  cpf_Deps and data_dep and dep_IRRFs and arq: # and nome_dependentes and cpf_Deps and data_dep and dep_IRRFs
                data.append({'cpf': cpfs,
                             'nomes':nomes,
                             'sexos':sexos,
                             'datas_nascimento':datas_nascimento,
                             'dtAdmissao':dtAdmissao,
                             'nome_dependentes':nome_dependentes,
                             'cpf_Deps':cpf_Deps,
                             'data_dep':data_dep,
                             'dep_IRRFs':dep_IRRFs,
                             'arq':arq
                             })
        else:
            pass

    if len(data) == 0:
        pass   
    else:   
        df = pd.DataFrame(data)
        # Salvar em um arquivo Excel (substitua 'dados.xml' pelo nome desejado)
        df.to_excel('Admissao S-2200.xlsx', index=False)



def formatacaoS_2200(arquivo_excel):    


    # Nomes das colunas que contêm as listas
    cpfs = 'cpf'
    nomes = 'nomes'
    sexos = 'sexos'
    datas_nascimento = 'datas_nascimento'
    dtAdmissao = 'dtAdmissao'
    nome_dependentes = 'nome_dependentes'
    cpf_Deps = 'cpf_Deps'
    data_dep = 'data_dep'
    dep_IRRFs ='dep_IRRFs'
    arq = 'arq'

    # Carrega a planilha Excel
    df = pd.read_excel(arquivo_excel)

    # Lista para armazenar os novos dados
    novos_dados = []

    # Itera sobre cada linha do DataFrame
    for index, row in df.iterrows():
        # Converte as strings das listas em listas reais
        lista_valores1 = eval(row[cpfs])
        lista_valores2 = eval(row[nomes])
        lista_valores3 = eval(row[sexos])
        lista_valores4 = eval(row[datas_nascimento])
        lista_valores5 = eval(row[dtAdmissao])
        lista_valores6 = eval(row[dtAdmissao])
        lista_valores7 = eval(row[cpf_Deps])
        lista_valores8 =  eval(row[data_dep])
        lista_valores9 =  eval(row[dep_IRRFs])
        lista_valores10 =  eval(row[arq])
        
        

        
        # Garante que ambas as listas tenham o mesmo tamanho
        max_len = max(len(lista_valores1), len(lista_valores2),len(lista_valores3),len(lista_valores4),len(lista_valores5),len(lista_valores6), len(lista_valores7),len(lista_valores8),len(lista_valores9),len(lista_valores10))
        lista_valores1.extend([None] * (max_len - len(lista_valores1)))
        lista_valores2.extend([None] * (max_len - len(lista_valores2)))
        lista_valores3.extend([None] * (max_len - len(lista_valores3)))
        lista_valores4.extend([None] * (max_len - len(lista_valores4)))
        lista_valores5.extend([None] * (max_len - len(lista_valores5)))
        lista_valores6.extend([None] * (max_len - len(lista_valores6)))
        lista_valores7.extend([None] * (max_len - len(lista_valores7)))
        lista_valores8.extend([None] * (max_len - len(lista_valores8)))
        lista_valores9.extend([None] * (max_len - len(lista_valores9))) 
        lista_valores10.extend([None] * (max_len - len(lista_valores10))) 

        for valor1, valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10 in zip(lista_valores1, lista_valores2,lista_valores3,lista_valores4,lista_valores5,lista_valores6,lista_valores7,lista_valores8,lista_valores9,lista_valores10):
            nova_linha = row.copy()
            nova_linha[cpfs] = valor1
            nova_linha[nomes] = valor2
            nova_linha[sexos] = valor3
            nova_linha[datas_nascimento] = valor4
            nova_linha[dtAdmissao] = valor5
            nova_linha[nome_dependentes] = valor6
            nova_linha[cpf_Deps] = valor7
            nova_linha[data_dep] = valor8
            nova_linha[dep_IRRFs] = valor9
            nova_linha[arq] = valor10
            
            novos_dados.append(nova_linha)
                        
    # Cria um novo DataFrame com os novos dados
    novo_df_S_1200 = pd.DataFrame(novos_dados)

    # Salva o novo DataFrame em uma nova planilha Excel
    novo_df_S_1200.to_excel(f'{arquivo_excel}', index=False)
    print("Planilha Admissão S-2200 criada com sucesso!🆗")




