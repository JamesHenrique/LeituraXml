import pandas as pd
import os
from xml.dom import minidom




def pegar_valores(arq,valores,pasta):
    
    with open(f'{pasta}\{arq}','rb') as arquivo_xml:
        info_xmls = minidom.parse(arquivo_xml)
        cpfTrab = info_xmls.getElementsByTagName('cpfTrab')[0].firstChild.data 
        matricula = info_xmls.getElementsByTagName('matricula')[0].firstChild.data 
        nrInsc = info_xmls.getElementsByTagName('nrInsc')[0].firstChild.data 
        nmCargo = info_xmls.getElementsByTagName('nmCargo')[0].firstChild.data 
        vrSalFx = info_xmls.getElementsByTagName('vrSalFx')[0].firstChild.data
        codCargo = info_xmls.getElementsByTagName('codCargo')[0].firstChild.data
        try:
            CBOCargo = info_xmls.getElementsByTagName('codCBO')[0].firstChild.data
        except:
            CBOCargo = " "
        dtAlteracao = info_xmls.getElementsByTagName('dtAlteracao')[0].firstChild.data
        
    valores.append([cpfTrab,matricula,nrInsc,nmCargo,vrSalFx,CBOCargo,codCargo,dtAlteracao])
    
    

def formatacaoS2206(pasta):
    lista_arquivos = os.listdir(pasta)

    colunas = ["cpf_trabalhador","matricula","nrInsc","nmCargo","vrSalFx","CBOCargo","codCargo","dtAlteracao"]
    valores  = []

    
    for aquivos in lista_arquivos:
        if 'S-2206.xml' in aquivos:
            pegar_valores(aquivos,valores,pasta)
            tabela = pd.DataFrame(columns=colunas,data=valores)
        else:
            pass
    

    if len(valores) == 0:
        return False
    else:  
        tabela.to_excel('AlteracaoCargo S-2206.xlsx', index=False)
        return True


