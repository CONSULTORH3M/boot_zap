

import openpyxl
import webbrowser
import pyautogui
import time
import tkinter as tk
from urllib.parse import quote
from threading import Thread
from datetime import datetime  # Para pegar a hora atual

# Função para enviar mensagens no WhatsApp Web
def enviar_mensagem(empresa, nome, telefone, inicio, mensagem):
    try:
        # Gerar o link do WhatsApp com a mensagem
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        time.sleep(15)  # Espera o WhatsApp carregar

        # Pressionar 'Enter' para enviar a mensagem
        pyautogui.press("enter")
        time.sleep(2)  # Pausa de 2 segundos para o envio ser concluído

        # Fechar a aba do WhatsApp
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(3)

        # Obter a hora atual para exibir
        hora_atual = datetime.now().strftime('%H:%M:%S')

        # Atualizar o status na interface gráfica com a hora
        status_message.insert(tk.END, f"Mensagem enviada para {empresa} ({telefone}) às {hora_atual}\n")
        status_message.yview(tk.END)  # Rolagem automática para o final
    except Exception as e:
        status_message.insert(tk.END, f"Erro ao enviar para {empresa}: {e}\n")
        status_message.yview(tk.END)


# Função para iniciar o envio
def start_send_messages():
    global current_row
    # Abrir o WhatsApp Web
    webbrowser.open('https://web.whatsapp.com/')
    time.sleep(15)  # Espera o WhatsApp carregar

    # Ler a planilha e obter as informações
    workbook = openpyxl.load_workbook('Mala_Whats.xlsx')
    pagina_clientes = workbook['Enviar']

    # Iniciar o envio de mensagens
    send_next_message(pagina_clientes)


# Função para enviar mensagens com intervalo de 2 minutos entre elas
def send_next_message(pagina_clientes):
    global current_row
    try:
        # Verificar se há mais clientes para enviar
        if current_row < len(list(pagina_clientes.rows)):
            linha = pagina_clientes[current_row]
            empresa = linha[0].value
            nome = linha[1].value
            telefone = linha[2].value
            inicio = linha[3].value

            # Preparar a mensagem
            mensagem = f'{inicio}, {nome} da Empresa {empresa}, Nós da *GDI Informática*, trabalhamos com um *Sistema de Gestão: Simples e Prático*, atraves do Sistema *EvoluTI*,\
            Prestamos um Suporte Rápido e Atuante, pode sempre contar com a nossa Ajuda, no Whats do Suporte *55 9119 4370*. Somos Parceiros de sua empresa, mas caso não queira \
            mais receber esse tipo de mensagem, basta enviar a palavra *Sair*!'

            # Enviar mensagem
            enviar_mensagem(empresa, nome, telefone, inicio, mensagem)

            # Atualizar a linha para o próximo cliente
            current_row += 1

            # Agendar o envio da próxima mensagem após 2 minutos (120000 ms)
            janela.after(120000, lambda: send_next_message(pagina_clientes))
        else:
            status_message.insert(tk.END, "Todos os clientes receberam a mensagem.\n")
            status_message.yview(tk.END)
            print("Envio de mensagens finalizado.")
    except Exception as e:
        status_message.insert(tk.END, f"Erro no envio: {e}\n")
        status_message.yview(tk.END)


# Função para parar a execução
def stop_execution():
    global stop_flag
    stop_flag = True
    janela.quit()


# Variáveis de controle
current_row = 2  # Começa da segunda linha (ignorando o cabeçalho)
stop_flag = False


# Interface gráfica
janela = tk.Tk()
janela.title('Envio Mensagens Automatico no WhatsApp')
janela.geometry("780x420")  # Aumentar o tamanho da janela
janela.iconbitmap("icon.ico")
janela.resizable(False, False)

# Caixa de texto para exibir o status
status_message = tk.Text(janela, height=20, width=90)  # Aumentar a altura e largura da caixa
status_message.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

# Botões
botao_start = tk.Button(janela, text='START', color=("red"), command=start_send_messages)
botao_start.grid(row=2, column=0, padx=10, pady=10, ipadx=50)

botao_stop = tk.Button(janela, text='STOP', command=stop_execution)
botao_stop.grid(row=2, column=1, padx=10, pady=10, ipadx=50)

# Label para mostrar o status
resultado_label = tk.Label(janela, text="")
resultado_label.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

janela.mainloop()
