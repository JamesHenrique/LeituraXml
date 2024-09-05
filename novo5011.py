import pandas as pd
from xml.dom import minidom
import os


def pegando_arquivos(arquivo_xml):
    # Lendo o arquivo XML
        dom = minidom.parse(arquivo_xml)
        # Inicializando listas para armazenar os dados
        
        perApur = []
        insCs = []
        vrDescCP = []
        vrCpSeg = []
        classTrib = []
        cnaePrep = []
        aliqRat = []
        fap = []
        aliqRatAjust = []
        fpas = []
        codTercs = []
        codCateg = []
        vrBcCp00 = []
        vrBcCp25 = []
        vrSalFam = []
        vrDescSest = []
        vrCalcSest = []
        vrDescSenat = []
        vrCalcSenat = []
        tpCR = []
        vrCR = []
        
    
        id = []


        # # Obtendo os elementos das tags
        perApur_elements = dom.getElementsByTagName('perApur')[0].firstChild.data
        insCs_elements = dom.getElementsByTagName('nrInsc')[1].firstChild.data
        vrDescCP_elements = dom.getElementsByTagName('vrDescCP')[0].firstChild.data
        vrCpSeg_elements = dom.getElementsByTagName('vrCpSeg')[0].firstChild.data
        classTrib_elements = dom.getElementsByTagName('classTrib')[0].firstChild.data
        cnaePrep_elements =  dom.getElementsByTagName('cnaePrep')[0].firstChild.data
        aliqRat_elements = dom.getElementsByTagName('aliqRat')[0].firstChild.data
        fap_elements = dom.getElementsByTagName('fap')[0].firstChild.data
        aliqRatAjust_elements = dom.getElementsByTagName('aliqRatAjust')[0].firstChild.data
        fpas_elements = dom.getElementsByTagName('fpas')[0].firstChild.data
        codTercs_elements = dom.getElementsByTagName('codTercs')[0].firstChild.data
        id_elements = dom.getElementsByTagName('evtCS')[0]
        id_value = id_elements.getAttribute('Id')
        codCateg_elements = dom.getElementsByTagName('codCateg')
        vrBcCp00_elements = dom.getElementsByTagName('vrBcCp00')
        vrBcCp25_elements = dom.getElementsByTagName('vrBcCp25')
        vrSalFam_elements = dom.getElementsByTagName('vrSalFam')
        vrDescSest_elements = dom.getElementsByTagName('vrDescSest')
        vrCalcSest_elements = dom.getElementsByTagName('vrCalcSest')
        vrDescSenat_elements = dom.getElementsByTagName('vrDescSenat')
        vrCalcSenat_elements = dom.getElementsByTagName('vrCalcSenat')
        tpCR_elements = dom.getElementsByTagName('tpCR')
        vrCR_elements = dom.getElementsByTagName('vrCR')
        
        

        
            
        for tag in tpCR_elements:
                tpCR_elementss = tag.firstChild.nodeValue.strip() 
                tpCR.append(tpCR_elementss)
                perApur.append(perApur_elements)
                insCs.append(insCs_elements)
                vrDescCP.append(vrDescCP_elements)
                vrCpSeg.append(vrCpSeg_elements)
                classTrib.append(classTrib_elements)
                cnaePrep.append(cnaePrep_elements)
                aliqRat.append(aliqRat_elements)
                fap.append(fap_elements)
                aliqRatAjust.append(aliqRatAjust_elements)
                fpas.append(fpas_elements)
                codTercs.append(codTercs_elements)
                id.append(id_value)
                
        for codCategs_elem in codCateg_elements:
            codCategs = codCategs_elem.firstChild.nodeValue.strip() 
            codCateg.append(codCategs)        
  
        for vrBcCp00_elem in vrBcCp00_elements:
            tags = vrBcCp00_elem.firstChild.nodeValue.strip() 
            vrBcCp00.append(tags)
                
        for tag in vrBcCp25_elements:
            vrBcCp25s = tag.firstChild.nodeValue.strip() 
            vrBcCp25.append(vrBcCp25s)

        for tag in vrSalFam_elements:
            rSalFams = tag.firstChild.nodeValue.strip() 
            vrSalFam.append(rSalFams)            
        
        for tag in vrDescSest_elements:
            vrDescSests = tag.firstChild.nodeValue.strip() 
            vrDescSest.append(vrDescSests) 
        
        for tag in vrCalcSest_elements:
            rCalcSests = tag.firstChild.nodeValue.strip() 
            vrCalcSest.append(rCalcSests)
            
        for tag in vrDescSenat_elements:
            vrDescSenats = tag.firstChild.nodeValue.strip() 
            vrDescSenat.append(vrDescSenats)       
        
        for tag in vrCalcSenat_elements:
            vrCalcSenats = tag.firstChild.nodeValue.strip() 
            vrCalcSenat.append(vrCalcSenats)        
                    
                
        for tag in vrCR_elements:
            rCRs = tag.firstChild.nodeValue.strip() 
            vrCR.append(rCRs)        
        
        
        
        
        
   
        
        return perApur,insCs,vrDescCP,vrCpSeg,classTrib,cnaePrep,aliqRat,fap,aliqRatAjust,fpas,codTercs,codCateg,vrBcCp00,vrBcCp25,vrSalFam,vrDescSest,vrCalcSest,vrDescSenat,vrCalcSenat,tpCR,vrCR, id #,,,,tpCR,,

def mainS5011(folder_path):
    # Lista para armazenar os dados

    data = []

    # Percorrer todos os arquivos XML na pasta
    for filename in os.listdir(folder_path):
        if filename.lower().endswith("s-5011.xml"):
            arq_xml = os.path.join(folder_path, filename)
            perApur,insCs,vrDescCP,vrCpSeg,classTrib,cnaePrep,aliqRat,fap,aliqRatAjust,fpas,codTercs,codCateg,vrBcCp00,vrBcCp25,vrSalFam,vrDescSest,vrCalcSest,vrDescSenat,vrCalcSenat,tpCR,vrCR, id = pegando_arquivos(arq_xml) #codCateg,,
            if perApur and insCs and vrDescCP and vrCpSeg and classTrib and cnaePrep and aliqRat and fap and aliqRatAjust and fpas and codTercs and codCateg and vrBcCp00  and vrBcCp25 and vrSalFam and vrDescSest and vrCalcSest and vrDescSenat and vrCalcSenat and tpCR and vrCR and id: 
                data.append({'perApur': perApur,
                             'insCs':insCs,
                             'vrDescCP':vrDescCP,
                             'vrCpSeg':vrCpSeg,
                             'classTrib':classTrib,
                             'cnaePrep':cnaePrep,
                             'aliqRat':aliqRat,
                             'fap':fap,
                             'aliqRatAjust':aliqRatAjust,
                             'fpas':fpas,
                             'codTercs':codTercs,
                             'codCateg':codCateg,
                             'vrBcCp00':vrBcCp00,
                             'vrBcCp25':vrBcCp25,
                             'vrSalFam':vrSalFam,
                             'vrDescSest':vrDescSest,
                             'vrCalcSest':vrCalcSest,
                             'vrDescSenat':vrDescSenat,
                             'vrCalcSenat':vrCalcSenat,
                             'tpCR':tpCR,
                             'vrCR':vrCR,
                             'id':id 
                             
                             })
        else:
            pass
            
    if len(data) == 0:
        pass
    else:
        print("arquivo salvo")
        df = pd.DataFrame(data)
    # Salvar em um arquivo Excel (substitua 'dados.xml' pelo nome desejado)
        df.to_excel('Planilha S-5011.xlsx', index=False)
  





def formatacaoS5011(arquivo_Excel):
    

    # Nomes das colunas que contêm as listas
    perApur = 'perApur'
    insCs = 'insCs'
    vrDescCP = 'vrDescCP'
    vrCpSeg = 'vrCpSeg'
    classTrib = 'classTrib'
    cnaePrep = 'cnaePrep'
    aliqRat = 'aliqRat'
    fap = 'fap'
    aliqRatAjust = 'aliqRatAjust'
    fpas = 'fpas'
    codTercs = 'codTercs'
    codCateg = 'codCateg'
    vrBcCp00 = 'vrBcCp00'
    id = 'id'
    vrBcCp25 = 'vrBcCp25'
    vrSalFam = 'vrSalFam'
    vrDescSest = 'vrDescSest'
    vrCalcSest = 'vrCalcSest'
    vrDescSenat = 'vrDescSenat'
    vrCalcSenat = 'vrCalcSenat'
    tpCR = 'tpCR'
    vrCR = 'vrCR'
    
   
    
    
    
    # Carrega a planilha Excel
    df = pd.read_excel(arquivo_Excel)


    # Lista para armazenar os novos dados
    novos_dados_S_2299 = []

    # Itera sobre cada linha do DataFrame
    for index, row in df.iterrows():
        # Converte as strings das listas em listas 
        lista_valores1 = eval(row[perApur])
        lista_valores2 = eval(row[insCs])
        lista_valores3 = eval(row[vrDescCP])
        lista_valores4 = eval(row[vrCpSeg])
        lista_valores5 = eval(row[classTrib])
        lista_valores6 = eval(row[cnaePrep])
        lista_valores7 = eval(row[aliqRat])
        lista_valores8 = eval(row[fap])
        lista_valores9 = eval(row[aliqRatAjust])
        lista_valores10 = eval(row[fpas])
        lista_valores11 = eval(row[codTercs])
        lista_valores12 = eval(row[codCateg])
        lista_valores13 = eval(row[vrBcCp00])
        
        lista_valores14 = eval(row[vrBcCp25])
        lista_valores15 = eval(row[vrSalFam])
        lista_valores16 = eval(row[vrDescSest])
        lista_valores17 = eval(row[vrCalcSest])
        lista_valores18 = eval(row[vrDescSenat])
        lista_valores19 = eval(row[vrCalcSenat])
        lista_valores20 = eval(row[tpCR])
        lista_valores21 = eval(row[vrCR])
        lista_valores22 = eval(row[id]) 
        
                

        
        # Garante que ambas as listas tenham o mesmo tamanho
        max_len = max(len(lista_valores1), len(lista_valores2),len(lista_valores3),len(lista_valores4),len(lista_valores5),len(lista_valores6),len(lista_valores7),len(lista_valores8),len(lista_valores9),len(lista_valores10),len(lista_valores11),len(lista_valores12),len(lista_valores13),len(lista_valores14),len(lista_valores15),len(lista_valores16),len(lista_valores17),len(lista_valores18),len(lista_valores19),len(lista_valores20),len(lista_valores21),len(lista_valores22))#,,,,
        
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
        lista_valores11.extend([None] * (max_len - len(lista_valores11)))
        lista_valores12.extend([None] * (max_len - len(lista_valores12)))
        lista_valores13.extend([None] * (max_len - len(lista_valores13)))
        lista_valores22.extend([None] * (max_len - len(lista_valores22)))
        lista_valores14.extend([None] * (max_len - len(lista_valores14)))
        lista_valores15.extend([None] * (max_len - len(lista_valores15)))
        lista_valores16.extend([None] * (max_len - len(lista_valores16)))
        lista_valores17.extend([None] * (max_len - len(lista_valores17)))
        lista_valores18.extend([None] * (max_len - len(lista_valores18)))
        lista_valores19.extend([None] * (max_len - len(lista_valores19)))
        lista_valores20.extend([None] * (max_len - len(lista_valores20)))
        lista_valores21.extend([None] * (max_len - len(lista_valores21)))
        
        
        for valor1, valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9, valor10,valor11,valor12,valor13,valor14,valor15,valor16,valor17,valor18,valor19,valor20,valor21,valor22 in zip(lista_valores1, lista_valores2,lista_valores3,lista_valores4,lista_valores5,lista_valores6,lista_valores7,lista_valores8,lista_valores9,lista_valores10,lista_valores11,lista_valores12,lista_valores13,lista_valores13,lista_valores15,lista_valores16,lista_valores17,lista_valores18,lista_valores19,lista_valores20,lista_valores21,lista_valores22):#,,lista_valores20,lista_valores21

            nova_linha = row.copy()
            nova_linha[perApur] = valor1
            nova_linha[insCs] = valor2
            nova_linha[vrDescCP] = valor3
            nova_linha[vrCpSeg] = valor4
            nova_linha[classTrib] = valor5
            nova_linha[cnaePrep] = valor6
            nova_linha[aliqRat] = valor7
            nova_linha[fap] = valor8
            nova_linha[aliqRatAjust] = valor9
            nova_linha[fpas] = valor10
            nova_linha[codTercs] = valor11
            nova_linha[codCateg] = valor12
            nova_linha[vrBcCp00] = valor13
            
            nova_linha[vrBcCp25] = valor14
            nova_linha[vrSalFam] = valor15
            nova_linha[vrDescSest] = valor16
            nova_linha[vrCalcSest] = valor17
            nova_linha[vrDescSenat] = valor18
            nova_linha[vrCalcSenat] = valor19
            nova_linha[tpCR] = valor20
            nova_linha[vrCR] = valor21
            nova_linha[id] = valor22

            
            
            
            novos_dados_S_2299.append(nova_linha)
            
    # Cria um novo DataFrame com os novos dados
    novo_df_S_2299 = pd.DataFrame(novos_dados_S_2299)

    # Salva o novo DataFrame em uma nova planilha Excel
    novo_df_S_2299.to_excel(f'{arquivo_Excel}', index=False)
    print("Planilha S-5011 salva com sucesso!")



