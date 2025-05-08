import pywhatkit
import pyautogui
import time
import os
import openpyxl
import threading
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

executando = False
PROGRESSO_ARQUIVO = "progresso.txt"

def enviar_mensagem_whatsapp_auto(numero, nome, estado, municipio):
    try:
        numero_formatado = f"+55{numero.replace('-', '').replace(' ', '')}"

        mensagem = (
            f"Olá *{nome}*, meu nome é *Felipe*, sou Diretor Pedagógico do DEPARTAMENTO DE ASSISTÊNCIA A EDUCAÇÃO (SAEB) do Estado de *{estado}*, "
            f"estou entrando em contato devido ao *seu bom desempenho e comprometimento com a Educação, "
            f"através do Município de {municipio}.* Você ainda atua como educador?"
        )

        print(f"Enviando mensagem para {nome} ({numero_formatado})...")
        pywhatkit.sendwhatmsg_instantly(numero_formatado, mensagem, wait_time=15, tab_close=False) # Reduzi o wait_time

        print("Aguardando para enviar...")
        time.sleep(5)  # Tempo para garantir que a caixa de texto esteja focada

        pyautogui.press("enter")  # Enviar mensagem
        print("Mensagem enviada com sucesso!")

        time.sleep(5)  # Espaço entre mensagens

    except Exception as e:
        print(f"[ERRO] Falha ao enviar mensagem para {nome} ({numero}): {e}")

def ler_progresso(planilha_atual):
    if os.path.exists(PROGRESSO_ARQUIVO):
        with open(PROGRESSO_ARQUIVO, "r") as f:
            linhas = f.readlines()
            if len(linhas) >= 2:
                nome_salvo = linhas[0].strip()
                linha_salva = int(linhas[1].strip())
                if nome_salvo == planilha_atual:
                    return linha_salva
    return 5

def salvar_progresso(planilha_atual, linha):
    with open(PROGRESSO_ARQUIVO, "w") as f:
        f.write(f"{planilha_atual}\n{linha}")

def zerar_progresso(planilha_atual):
    with open(PROGRESSO_ARQUIVO, "w") as f:
        f.write(f"{planilha_atual}\n5")
    messagebox.showinfo("Reiniciado", "Progresso zerado com sucesso.")

def iniciar_envio():
    global executando
    executando = True
    thread = threading.Thread(target=main)
    thread.start()

def parar_envio():
    global executando
    executando = False
    messagebox.showinfo("Parar", "Envio de mensagens interrompido.")

def main():
    global executando
    arquivos_excel = [arquivo for arquivo in os.listdir() if arquivo.endswith(('.xlsx', '.xls'))]

    if not arquivos_excel:
        messagebox.showerror("Erro", "Nenhum arquivo Excel encontrado na pasta atual.")
        return

    arquivo_excel = arquivos_excel[0]

    # INSTRUÇÃO IMPORTANTE PARA O USUÁRIO
    messagebox.showinfo("Atenção", "Por favor, abra o WhatsApp Web no seu navegador e deixe-o carregado antes de clicar em 'Iniciar'.")

    try:
        workbook = openpyxl.load_workbook(arquivo_excel)
        sheet = workbook.active

        numero_coluna = nome_coluna = estado_coluna = municipio_coluna = None

        # Começar a leitura dos cabeçalhos a partir da linha 5 (índice 4)
        for i, coluna in enumerate(sheet[4]):
            if coluna.value is not None:
                coluna_lower = str(coluna.value).lower().strip()
                if "numero" in coluna_lower or "número" in coluna_lower:
                    numero_coluna = i
                elif "nome" in coluna_lower:
                    nome_coluna = i
                elif "estado" in coluna_lower:
                    estado_coluna = i
                elif "município" in coluna_lower or "municipio" in coluna_lower:
                    municipio_coluna = i

        if None in (numero_coluna, nome_coluna, estado_coluna, municipio_coluna):
            messagebox.showerror("Erro", "As colunas 'Número', 'Nome', 'Estado' ou 'Município' não foram encontradas na linha 5.")
            return

        linha_inicio = ler_progresso(arquivo_excel)

        for linha_num in range(linha_inicio, sheet.max_row + 1):
            if not executando:
                print("Envio interrompido pelo usuário.")
                break

            try:
                numero = str(sheet.cell(row=linha_num, column=numero_coluna + 1).value)
                nome = sheet.cell(row=linha_num, column=nome_coluna + 1).value
                estado = sheet.cell(row=linha_num, column=estado_coluna + 1).value
                municipio = sheet.cell(row=linha_num, column=municipio_coluna + 1).value

                if None in (numero, nome, estado, municipio):
                    print(f"Linha {linha_num} ignorada: dados incompletos.")
                    continue

                enviar_mensagem_whatsapp_auto(numero, nome, estado, municipio)
                salvar_progresso(arquivo_excel, linha_num + 1)

            except Exception as inner_e:
                print(f"[ERRO] Erro ao processar linha {linha_num}: {inner_e}")
                continue  # Tenta processar a próxima linha mesmo com erro

        if executando:
            messagebox.showinfo("Concluído", "Envio de mensagens finalizado.")
        executando = False

    except Exception as e:
        messagebox.showerror("Erro ao ler a planilha", str(e))

# Interface gráfica
janela = tk.Tk()
janela.title("Envio de WhatsApp Automático")

# Carregar a imagem de fundo
try:
    imagem_pil = Image.open("teladefundoboot.jpg")
    imagem_tk = ImageTk.PhotoImage(imagem_pil)
    fundo_label = tk.Label(janela, image=imagem_tk)
    fundo_label.place(x=0, y=0, relwidth=1, relheight=1)
except FileNotFoundError:
    messagebox.showerror("Erro", "A imagem 'bootplay.png' não foi encontrada.")
    janela.geometry("320x200") # Define um tamanho padrão se a imagem não for carregada



# Frame para os botões
frame_botoes = tk.Frame(janela, bg="#ADD8E6", bd=0) # Cor de fundo e borda removida
frame_botoes.pack(side=tk.BOTTOM, pady=20) # Botões na parte inferior

# Botões
botao_iniciar = tk.Button(frame_botoes, text="Iniciar", width=10, command=iniciar_envio, bg="green", fg="white")
botao_iniciar.pack(side=tk.LEFT, padx=5)

botao_parar = tk.Button(frame_botoes, text="Parar", width=10, command=parar_envio, bg="red", fg="white")
botao_parar.pack(side=tk.LEFT, padx=5)

def recomecar_zero():
    arquivos_excel = [arquivo for arquivo in os.listdir() if arquivo.endswith(('.xlsx', '.xls'))]
    if arquivos_excel:
        zerar_progresso(arquivos_excel[0])
    else:
        messagebox.showerror("Erro", "Nenhuma planilha encontrada para reiniciar.")

botao_recomecar = tk.Button(frame_botoes, text="Recomeçar", width=10, command=recomecar_zero, bg="blue", fg="white")
botao_recomecar.pack(side=tk.LEFT, padx=5)

janela.mainloop()