#Pacote para fazer uma interface gráfica com Python
import tkinter as tk

#Importando o pacote para modificar as fontes dos widgets
from tkinter import font

import win32com.client as win32

#Importando as funções para rodar o Pregão Eletrônico de acordo com o checkbox
from pregraoallocs import bec_pregaoeletronico
from pregraoonlyfromXtoY import bec_pregaoeletronico2
from pregraowithdates import bec_pregaoeletronico3


                                                                    ######################################
                                                                    ###RODANDO O CÓDIGO INTEIRO COM GUI###
                                                                    ######################################

#Verificação da entrada do usuário (Quantas OCs irá ser buscadas)
def total_numbersOCs(entry_number, entry_number2):
    field_value = entry_number.get()
    field_value2 = entry_number2.get()
    print(field_value)
    print(field_value2)
    
    if field_value.isdigit() and field_value2.isdigit():
        field_value = int(field_value)
        field_value2 = int(field_value2)
        if field_value >=1 and field_value2<=500:
            run_onlyXtoYOCs(field_value, field_value2) #Passando o valor do parâmetro para a próxima função usar
            #print("Valor menor ou igual à 20.")    
        else:
            secondary_screen = tk.Tk()
            screen_width2 = root.winfo_screenwidth()
            screen_height2 = root.winfo_screenheight()

            width2 = 500
            height2 = 400
            x = (screen_width2/2) - (width2/2)
            y = (screen_height2/2) - (height2/2)
            secondary_screen.geometry("{}x{}+{}+{}".format(width2, height2, int(x), int(y)))

            secondary_screen.title("ERROR! - Campo(s) não preenchido(s)")

            content2 = tk.Frame(secondary_screen, background="#E44D2D")
            content2.pack(fill=tk.BOTH, expand=True)

            font_backroot = font.Font(family="Arial", size=18, weight="bold")
            label = tk.Label(content2, text="POR FAVOR,\n COLOQUE NO PRIMEIRO CAMPO\nUM VALOR NUMÉRICO MAIOR QUE 0!\nE NO SEGUNDO CAMPO UM VALOR MENOR OU IGUAL\nDO QUE 100.", font=font_backroot, background="#E44D2D", fg="snow")
            label.pack(side=tk.TOP, anchor=tk.CENTER, padx=5, pady=100)

            back_button = tk.Button(content2, bg="#DEA228", font=font_backroot, fg="snow", text="VOLTAR", relief="raised", borderwidth=5, command=secondary_screen.destroy)
            back_button.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)

            secondary_screen.mainloop()
    else:        
            secondary_screen = tk.Tk()
            screen_width2 = root.winfo_screenwidth()
            screen_height2 = root.winfo_screenheight()
    
            width2 = 500
            height2 = 400
            x = (screen_width2/2) - (width2/2)
            y = (screen_height2/2) - (height2/2)
            secondary_screen.geometry("{}x{}+{}+{}".format(width2, height2, int(x), int(y)))
            
            secondary_screen.title("ERROR! - Campo(s) não preenchido(s)")
            
            content2 = tk.Frame(secondary_screen, background="#E44D2D")
            content2.pack(fill=tk.BOTH, expand=True)
            
            font_backroot = font.Font(family="Arial", size=18, weight="bold")
            label = tk.Label(content2, text="POR FAVOR,\n PREENCHA OS CAMPOS\nPARA QUE O SCRAPING\nSEJA FEITO!", font=font_backroot, background="#E44D2D", fg="snow")
            label.pack(side=tk.TOP, anchor=tk.CENTER, padx=5, pady=100)
            
            back_button = tk.Button(content2, bg="#DEA228", font=font_backroot, fg="snow", text="VOLTAR", relief="raised", borderwidth=5, command=secondary_screen.destroy)
            back_button.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)
            
            secondary_screen.mainloop()
     

#Função para pegar uma quantia delimitada de OCs que será de acordo com a entrada do usuário
def run_onlyXtoYOCs(field_value, field_value2):

            print(field_value)
            print(field_value2)
            
            bec_pregaoeletronico2(field_value, field_value2)
            
            finished_screen = tk.Tk()
            screen_width_finished = finished_screen.winfo_screenwidth()
            screen_height_finished = finished_screen.winfo_screenheight()

            width_finished = 500
            height_finished = 400

            x = (screen_width_finished/2) - (width_finished/2)
            y = (screen_height_finished/2) - (height_finished/2)

            finished_screen.geometry("{}x{}+{}+{}".format(width_finished, height_finished, int(x), int(y)))
            finished_screen.title("SUCESSO! - SCRAPING REALIZADO")

            finished_content = tk.Frame(finished_screen, background="#E44D2D")
            finished_content.pack(fill=tk.BOTH, expand=True)

            font_backroot2 = font.Font(family="Arial", size=20, weight="bold")
            label_finished = tk.Label(master=finished_content, text="FINALIZADO COM SUCESSO! \n O SCRAPING DAS OCs \n FORA EXECUTADO COM ÊXITO!", font=font_backroot2, background="#E44D2D", fg="snow")
            label_finished.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=70)

            back_button = tk.Button(master=finished_content, text="VOLTAR", font=font_backroot2, bg="#DEA228", fg="snow", relief="raised", borderwidth=5, command=finished_screen.destroy)
            back_button.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)

            finished_screen.mainloop()



#Função para pegar todos os valores de todas as Ofertas de Compra existentes (demorado)
def run_allOCs():
        bec_pregaoeletronico()
        
        finished_screen2 = tk.Tk()
        screen_width_finished2 = finished_screen2.winfo_screenwidth()
        screen_height_finished2 = finished_screen2.winfo_screenheight()

        width_finished2 = 500
        height_finished2 = 400

        x = (screen_width_finished2/2) - (width_finished2/2)
        y = (screen_height_finished2/2) - (height_finished2/2)

        finished_screen2.geometry("{}x{}+{}+{}".format(width_finished2, height_finished2, int(x), int(y)))
        finished_screen2.title("SUCESSO! - Scraping finalizado com sucesso")

        finished_content2 = tk.Frame(finished_screen2, background="#E44D2D")
        finished_content2.pack(fill=tk.BOTH, expand=True)

        font_backroot2 = font.Font(family="Arial", size=20, weight="bold")
        label_finished = tk.Label(master=finished_content2, text="FINALIZADO COM SUCESSO! \n O SCRAPING DE TODAS AS OCs \n FORA EXECUTADO COM ÊXITO!", font=font_backroot2, background="#E44D2D", fg="snow")
        label_finished.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=70)

        back_button = tk.Button(master=finished_content2, text="VOLTAR", font=font_backroot2, bg="#DEA228", fg="snow", relief="raised", borderwidth=5, command=finished_screen2.destroy)
        back_button.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)

        finished_screen2.mainloop()
        
        
#Função para verificar os campos com as datas e invocar a função bec_pregaoeletronico3()        
def check_dates(first_date_entry, second_date_entry):
    start_date = first_date_entry.get()
    end_date = second_date_entry.get()
    
    if len(start_date)==10 and len(end_date)==10:
        
        #Separando as partes da data pelas barras
        start_date_parts = start_date.split("/")
        end_date_parts = end_date.split("/")
        
        #Verificando se cada parte da primeira data está correta, se não está faltando ou sobrando nenhum número
        if len(start_date_parts[0]) == 2 and len(start_date_parts[1]) == 2 and len(start_date_parts[2])==4 :
            
            #Verificando a segunda data
            if len(end_date_parts[0]) == 2 and len(start_date_parts[1])== 2 and len(start_date_parts[2]) ==4 : 
                
                #Retirando as barras das datas para colocar nos campos e funcionar
                start_date_combined = start_date_parts[0] + start_date_parts[1] + start_date_parts[2]
                end_date_combined = end_date_parts[0] + end_date_parts[1] + end_date_parts[2] 
                bec_pregaoeletronico3(start_date_combined, end_date_combined)
                
                finished_screen2 = tk.Tk()
                screen_width_finished2 = finished_screen2.winfo_screenwidth()
                screen_height_finished2 = finished_screen2.winfo_screenheight()

                width_finished2 = 500
                height_finished2 = 400

                x = (screen_width_finished2/2) - (width_finished2/2)
                y = (screen_height_finished2/2) - (height_finished2/2)

                finished_screen2.geometry("{}x{}+{}+{}".format(width_finished2, height_finished2, int(x), int(y)))
                finished_screen2.title("SUCESSO! - Scraping finalizado com sucesso")

                finished_content2 = tk.Frame(finished_screen2, background="#E44D2D")
                finished_content2.pack(fill=tk.BOTH, expand=True)

                font_backroot2 = font.Font(family="Arial", size=20, weight="bold")
                label_finished = tk.Label(master=finished_content2, text="FINALIZADO COM SUCESSO! \n O SCRAPING DE TODAS AS OCs \n FORA EXECUTADO COM ÊXITO!", font=font_backroot2, background="#E44D2D", fg="snow")
                label_finished.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=70)

                back_button = tk.Button(master=finished_content2, text="VOLTAR", font=font_backroot2, bg="#DEA228", fg="snow", relief="raised", borderwidth=5, command=finished_screen2.destroy)
                back_button.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)

                finished_screen2.mainloop()
            
            else:
                secondary_screen = tk.Tk()
                screen_width2 = root.winfo_screenwidth()
                screen_height2 = root.winfo_screenheight()

                width2 = 500
                height2 = 400
                x = (screen_width2/2) - (width2/2)
                y = (screen_height2/2) - (height2/2)
                secondary_screen.geometry("{}x{}+{}+{}".format(width2, height2, int(x), int(y)))

                secondary_screen.title("ERROR! - Campo(s) preenchido(s) errado")

                content2 = tk.Frame(secondary_screen, background="#E44D2D")
                content2.pack(fill=tk.BOTH, expand=True)

                font_backroot = font.Font(family="Arial", size=18, weight="bold")
                label = tk.Label(content2, text="POR FAVOR,\n COLOQUE A DATA NO FORMATO CORRETO\nSENDO O FORMATO: DD/MM/AAAA.", font=font_backroot, background="#E44D2D", fg="snow")
                label.pack(side=tk.TOP, anchor=tk.CENTER, padx=5, pady=100)

                back_button = tk.Button(content2, bg="#DEA228", font=font_backroot, fg="snow", text="VOLTAR", relief="raised", borderwidth=5, command=secondary_screen.destroy)
                back_button.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)
        else:
                secondary_screen = tk.Tk()
                screen_width2 = root.winfo_screenwidth()
                screen_height2 = root.winfo_screenheight()

                width2 = 500
                height2 = 400
                x = (screen_width2/2) - (width2/2)
                y = (screen_height2/2) - (height2/2)
                secondary_screen.geometry("{}x{}+{}+{}".format(width2, height2, int(x), int(y)))

                secondary_screen.title("ERROR! - Campo(s) não preenchido(s)")

                content2 = tk.Frame(secondary_screen, background="#E44D2D")
                content2.pack(fill=tk.BOTH, expand=True)

                font_backroot = font.Font(family="Arial", size=18, weight="bold")
                label = tk.Label(content2, text="POR FAVOR,\n PREENCHA CORRETAMENTE OS DOIS CAMPOS\nSENDO NO FORMATO: DD/MM/AAAA.", font=font_backroot, background="#E44D2D", fg="snow")
                label.pack(side=tk.TOP, anchor=tk.CENTER, padx=5, pady=100)

                back_button = tk.Button(content2, bg="#DEA228", font=font_backroot, fg="snow", text="VOLTAR", relief="raised", borderwidth=5, command=secondary_screen.destroy)
                back_button.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)
    else:
                secondary_screen = tk.Tk()
                screen_width2 = root.winfo_screenwidth()
                screen_height2 = root.winfo_screenheight()

                width2 = 500
                height2 = 400
                x = (screen_width2/2) - (width2/2)
                y = (screen_height2/2) - (height2/2)
                secondary_screen.geometry("{}x{}+{}+{}".format(width2, height2, int(x), int(y)))

                secondary_screen.title("ERROR! - Campo(s) não preenchido(s) ou formato errado")

                content2 = tk.Frame(secondary_screen, background="#E44D2D")
                content2.pack(fill=tk.BOTH, expand=True)

                font_backroot = font.Font(family="Arial", size=18, weight="bold")
                label = tk.Label(content2, text="POR FAVOR,\n PREENCHA OS DOIS CAMPOS CORRETAMENTE\nSENDO NO FORMATO: DD/MM/AAAA.", font=font_backroot, background="#E44D2D", fg="snow")
                label.pack(side=tk.TOP, anchor=tk.CENTER, padx=5, pady=100)

                back_button = tk.Button(content2, bg="#DEA228", font=font_backroot, fg="snow", text="VOLTAR", relief="raised", borderwidth=5, command=secondary_screen.destroy)
                back_button.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)

        
        
                                                                    ######################################
                                                                    ###CRIAÇÃO DE UMA INTERFACE GRÁFICA###
                                                                    ######################################

#Criação da Janela onde estará todos os Widgets (controles, elementos, janelas,...) - É o elemento Pai da hierarquia de widgets
root = tk.Tk()

root.iconbitmap("Imagens/bec_ico_blue.ico")

#Pegando a Largura e Altura da tela do monitor
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

#Definindo os valores para a Largura e Altura da janela
width = 800
height = 600

#Calculando as coordenadas de X e Y para o centro da tela com base no tamanho do monitor e da janela do aplicativo
x = (screen_width/2) - (width/2)
y = (screen_height/2) - (height/2)

#Definindo/Formatando os valores geométricos (tamanhos) da janela 
root.geometry("{}x{}+{}+{}".format(width, height, int(x), int(y)))
root.title("WebScrapping - Pregão BEC")

#Criação do elemento filho do root
content = tk.Frame(root, background="#E44D2D")
content.pack(fill=tk.BOTH, expand=True)

#Criação de uma padronização das fontes dos widgets
font_all = font.Font(family="Arial", size=18, weight="bold")

#Criação dos próximos elementos da hierarquia

#Criação do campo de entrada (widget)
label = tk.Label(master = content, text="Digite o número da linha inicial:", font=font_all, background="#E44D2D", fg="snow")
label.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)

entry_number = tk.Entry(master=content)
entry_number.pack(side=tk.TOP, anchor=tk.CENTER, padx=20, pady=5)

label2 = tk.Label(master = content, text="Digite o número da linha final:", font=font_all, background="#E44D2D", fg="snow")
label2.pack(side=tk.TOP, anchor=tk.CENTER, padx=10, pady=10)

entry_number2 = tk.Entry(master = content)
entry_number2.pack(side=tk.TOP, anchor=tk.CENTER, padx=10)

#Criação do botão que irá invocar a função para rodar a função bec_pregaoeletronico2()
button_onlyxtoy = tk.Button(master = content, text="Scraping", font=font_all, command=lambda:total_numbersOCs(entry_number, entry_number2), bg="#DEA228", fg="snow", relief="raised", borderwidth=5)
button_onlyxtoy.pack(side=tk.TOP, anchor=tk.CENTER, padx=20, pady=20)

#Criação dos campos para colocar as datas
date_label1 = tk.Label(master = content, text="Primeira data:", font=font_all, background="#E44D2D", fg="snow")
date_label1.pack(side=tk.TOP, anchor=tk.CENTER, padx=20, pady=5)

first_date_entry = tk.Entry(master = content)
first_date_entry.pack(side=tk.TOP, anchor=tk.CENTER, padx=20, pady=5)
first_date_entry.focus()

date_label2 = tk.Label(master = content, text="Segunda data:", font=font_all, background="#E44D2D", fg="snow")
date_label2.pack(side=tk.TOP, anchor=tk.CENTER, padx=20, pady=5)

second_date_entry = tk.Entry(master = content)
second_date_entry.pack(side=tk.TOP, anchor=tk.CENTER, padx=20, pady=5)

#Botão para fazer o Scraping utilizando as datas
button_dates = tk.Button(master = content, text="Scraping", font=font_all, command=lambda:check_dates(first_date_entry, second_date_entry), bg="#DEA228", fg="snow", relief="raised", borderwidth=5)
button_dates.pack(side=tk.TOP, anchor=tk.CENTER, padx=20, pady=10)

#Botão para fazer o Scraping de todas as OCs
button_all = tk.Button(master = content, text="Scraping de todas OCs", font=font_all, command=run_allOCs, bg="#DEA228", fg="snow", relief="raised", borderwidth=5)
button_all.pack(side=tk.TOP, anchor=tk.CENTER, padx=20, pady=20)


root.mainloop()


            
    
    

