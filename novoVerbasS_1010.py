#xml verbas s-1010
import xmltodict
import os
import pandas as pd 



def pegar_infos(arq,valores,caminho):
        with open(f'{caminho}/{arq}', "rb") as arquivo_xml:
            print(f'abriu arquivo: {arq}')
            dic_arquivo = xmltodict.parse(arquivo_xml) #tranforma xml em dicionario python
            infos_xmls = dic_arquivo['eSocial']['retornoProcessamentoDownload']['evento']['eSocial']['evtTabRubrica']['infoRubrica']
            if "inclusao" in infos_xmls:
                codRubr = infos_xmls['inclusao']['ideRubrica']['codRubr']
                dscRubr = infos_xmls['inclusao']['dadosRubrica']['dscRubr']
                natRubr = infos_xmls['inclusao']['dadosRubrica']['natRubr']
                tpRubr =  infos_xmls['inclusao']['dadosRubrica']['tpRubr']
                codIncCP = infos_xmls['inclusao']['dadosRubrica']['codIncCP']
                codIncIRRF = infos_xmls['inclusao']['dadosRubrica']['codIncIRRF']
                codIncFGTS = infos_xmls['inclusao']['dadosRubrica']['codIncFGTS']
                iniValid = infos_xmls['inclusao']['ideRubrica']['iniValid']
                try:
                    fimValid = infos_xmls['inclusao']['ideRubrica']['fimValid']
                except:
                    fimValid = "Sem informação"
                    
            if "alteracao" in infos_xmls:
                codRubr = infos_xmls['alteracao']['ideRubrica']['codRubr']
                dscRubr = infos_xmls['alteracao']['dadosRubrica']['dscRubr']
                natRubr = infos_xmls['alteracao']['dadosRubrica']['natRubr']
                tpRubr =  infos_xmls['alteracao']['dadosRubrica']['tpRubr']
                codIncCP = infos_xmls['alteracao']['dadosRubrica']['codIncCP']
                codIncIRRF = infos_xmls['alteracao']['dadosRubrica']['codIncIRRF']
                codIncFGTS = infos_xmls['alteracao']['dadosRubrica']['codIncFGTS']
                iniValid = infos_xmls['alteracao']['ideRubrica']['iniValid']
                try:
                    fimValid = infos_xmls['alteracao']['ideRubrica']['fimValid']
                except:
                    fimValid = "Sem informação"
            if "exclusao" in infos_xmls:
                codRubr = infos_xmls['exclusao']['ideRubrica']['codRubr']
                dscRubr = "Sem informação"
                natRubr = "Sem informação"
                tpRubr =  "Sem informação"
                codIncCP = "Sem informação"
                codIncIRRF = "Sem informação"
                codIncFGTS = "Sem informação"
                iniValid = infos_xmls['exclusao']['ideRubrica']['iniValid']
                try:
                    fimValid = infos_xmls['exclusao']['ideRubrica']['fimValid']
                except:
                    fimValid = "Sem informação"
        valores.append([codRubr,dscRubr,natRubr,tpRubr,codIncCP,codIncIRRF,codIncFGTS,iniValid,fimValid,arq])


   
caminho = r"C:\Users\TI\Documents\James\PROJETO ESOCIAL\Arquivos\zip\zip"

def formatacaoS1010(caminho):   
    lista_aquivos = os.listdir(caminho)
    colunas = ["codRubr","dscRubr","natRubr","tpRubr","codIncCP","codIncIRRF","codIncFGTS","iniValid","fimValid","arq"]
    valores  = []
    
    for arquivo in lista_aquivos:
        if "S-1010.xml" in arquivo:
            pegar_infos(arquivo,valores,caminho)
            tabela = pd.DataFrame(columns=colunas,data=valores)
        else:
            pass  
    if len(valores) == 0:
        return False
    else:  
        tabela.to_excel('Verbas S-1010.xlsx', index=False)
        return True    

formatacaoS1010(caminho)
