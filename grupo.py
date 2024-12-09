import tkinter as tk
from tkinter import ttk
import threading
import time

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Envio de Mensagens")
        self.root.geometry("600x400")

        # Variáveis
        self.is_running = False
        self.selected_group = tk.StringVar(value="Grupo 1")
        
        # Seleção do grupo
        self.group_label = tk.Label(root, text="Selecione o Grupo")
        self.group_label.grid(row=0, column=0, padx=10, pady=10)
        
        self.group_menu = ttk.Combobox(root, textvariable=self.selected_group, values=["Grupo 1", "Grupo 2", "Grupo 3"])
        self.group_menu.grid(row=0, column=1, padx=10, pady=10)
        
        # Botões
        self.start_button = tk.Button(root, text="Iniciar Envio", command=self.start_sending)
        self.start_button.grid(row=1, column=0, padx=10, pady=10)

        self.stop_button = tk.Button(root, text="Parar", command=self.stop_sending, state=tk.DISABLED)
        self.stop_button.grid(row=1, column=1, padx=10, pady=10)
        
        self.close_button = tk.Button(root, text="Fechar", command=self.close_app)
        self.close_button.grid(row=1, column=2, padx=10, pady=10)

        # Grid para mostrar mensagens
        self.tree = ttk.Treeview(root, columns=("Grupo", "Mensagem", "Status"), show="headings")
        self.tree.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
        
        self.tree.heading("Grupo", text="Grupo")
        self.tree.heading("Mensagem", text="Mensagem")
        self.tree.heading("Status", text="Status")

        # Configura layout da grid
        root.grid_rowconfigure(2, weight=1)
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=1)
        root.grid_columnconfigure(2, weight=1)

    def start_sending(self):
        """Inicia o envio das mensagens"""
        if not self.is_running:
            self.is_running = True
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            
            # Criar thread para envio simulado
            self.thread = threading.Thread(target=self.send_messages)
            self.thread.start()

    def send_messages(self):
        """Simula o envio de mensagens e mostra status na grid"""
        grupo = self.selected_group.get()
        for i in range(1, 11):
            if not self.is_running:  # Se for interrompido
                break
            mensagem = f"Mensagem {i}"
            status = "Enviando"
            self.tree.insert("", "end", values=(grupo, mensagem, status))
            self.root.update_idletasks()  # Atualiza a interface

            time.sleep(1)  # Simula o envio com delay de 1 segundo
            
            # Atualizar status
            self.tree.item(self.tree.get_children()[-1], values=(grupo, mensagem, "Enviado"))
            self.root.update_idletasks()  # Atualiza a interface

        # Finaliza envio
        self.finish_sending()

    def stop_sending(self):
        """Interrompe o envio de mensagens"""
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

    def finish_sending(self):
        """Finaliza a operação de envio"""
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

    def close_app(self):
        """Fecha a aplicação"""
        self.is_running = False
        self.root.quit()

# Função principal para rodar o aplicativo
def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
