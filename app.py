import logging
import customtkinter as ct
from customtkinter import filedialog
from tkinter import scrolledtext
import os
from novoRecisaoS_2299 import main, formatacao
from novoReciboS_1200 import mainS_1200,formatacaoS_1200
from novoAdmissaoS_2200 import mainS_2200,formatacaoS_2200
from novoVerbasS_1010 import formatacaoS1010
from novoAlteracaoCargo import formatacaoS2206
from novoInclusaoDepS_2205 import formatacaoS_2205, mainS_2205
from novo5011 import mainS5011, formatacaoS5011

import time
    

# Configuração do logger
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(message)s')
logger = logging.getLogger()
tempo = 2

def openFile():

    pasta = filedialog.askdirectory() #escolha a pasta que tem os arquivos
    
    label3.configure(text="Procurando S-2299") 
    janela.update_idletasks()  
    main(pasta)
    if os.path.exists('Recisao S-2299.xlsx'):
        formatacao('Recisao S-2299.xlsx') 
        teste_logger("Planilha Recisao S-2299 criada com sucesso!🆗") 
    else:
        teste_logger("xml S-2299 não encontrado") 
    time.sleep(tempo)
    
   
    label3.configure(text="Procurando S-1200") 
    janela.update_idletasks()  
    mainS_1200(pasta)
    if os.path.exists('Recibos S-1200.xlsx'):
        formatacaoS_1200('Recibos S-1200.xlsx')
        teste_logger("Planilha Recibos S-1200 criada com sucesso!🆗") 
    else:
        teste_logger("xml S-1200 não encontrado") 
    time.sleep(tempo)
    
    label3.configure(text="Procurando S-2200")
    janela.update_idletasks()  
    mainS_2200(pasta)
    if os.path.exists('Admissao S-2200.xlsx'):
        formatacaoS_2200('Admissao S-2200.xlsx') 
        teste_logger("Planilha Admissão S-2200 criada com sucesso!🆗")
    else:
        teste_logger('xml S-2200 não encontrado')
        pass
    time.sleep(tempo)
     
    label3.configure(text="Procurando S-2205")
    janela.update_idletasks()  
    mainS_2205(pasta)
    if os.path.exists('Inclusao Dependente S-2205.xlsx'):
        formatacaoS_2205('Inclusao Dependente S-2205.xlsx') 
        teste_logger("Planilha Inclusao Dependente S-2205 criada com sucesso!🆗")
    else:
        teste_logger("xml S-2205 não encontrado") 
    
    time.sleep(tempo)
    
    label3.configure(text="Procurando S-1010")
    janela.update_idletasks()  
    formatacaoS1010(pasta)
    if formatacaoS1010(pasta) == False:
        teste_logger("xml S-1010 não encontrado") 
    else:
        teste_logger("Planilha Verbas S-1010 criada com sucesso!🆗")

        
    time.sleep(tempo)
    label3.configure(text="Procurando S-2206")
    janela.update_idletasks()  
    formatacaoS2206(pasta)
    if formatacaoS2206(pasta) == False:
        teste_logger("xml S-2206 não encontrado") 
    else:
        teste_logger("Planilha AlteracaoCargo S-2206 criada com sucesso!🆗")

    time.sleep(tempo)
    
    label3.configure(text="Procurando S-5011")
    janela.update_idletasks()  
    mainS5011(pasta)
    if os.path.exists('Planilha S-5011.xlsx'):
        formatacaoS5011('Planilha S-5011.xlsx') 
        teste_logger("Planilha  S-5011 criada com sucesso!🆗")
    else:
        teste_logger('xml S-5011 não encontrado')
        pass
    time.sleep(tempo)
     

    label3.configure(text='Busca finalizada ✔️')
    

# funçao para criar mensagem no leg  
def teste_logger(message):
    textbox.config(state='normal')
    textbox.insert(ct.END, message + '\n')
    textbox.config(state='disabled')
    textbox.yview(ct.END) 
    janela.update_idletasks()    

#Centraliza a tela
def centralizar_janela(janela):
    largura_janela = 500
    altura_janela = 350
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    pos_x = (largura_tela // 2) - (largura_janela // 2)
    pos_y = (altura_tela // 2) - (altura_janela // 2)
    janela.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")



# Inicializando a aplicação
janela = ct.CTk()
janela.title("App Busca XMLs")
ct.set_appearance_mode('dark')
centralizar_janela(janela)



# Centralizando a janela


label1 = ct.CTkLabel(janela,text='ESCOLHA A PASTA QUE CONTEM OS XMLs:\n\n S-2200 / S-1010 / S-1200 / S-2205 / S-2206 / S-2299',
                    text_color="aliceblue",
                    font = ("", 15)
                    )
label1.pack(pady=5)

label3 = ct.CTkLabel(janela,text='',
                    text_color="aliceblue",
                    font = ("", 15)
                    )
label3.pack()


textbox = scrolledtext.ScrolledText(janela, state='disabled', wrap='word', width=50, height=5,font = ("", 12))
textbox.pack(pady=20)

# textbox.insert("0.0","Processo..: \n\n")

btn = ct.CTkButton(janela, text="ESCOLHER A PASTA",command=openFile,width=200)
btn.pack(pady=5)


janela.mainloop()






            

