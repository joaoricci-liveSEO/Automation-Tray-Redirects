# Este é o automation_gui.py

import customtkinter as ctk
import threading
import queue
import sys

# Importa as funções do seu outro arquivo
try:
    import automation_logic
except ImportError:
    print("ERRO: Não foi possível encontrar o arquivo 'automation_logic.py'")
    sys.exit()

# --- Nossas cores ---
COR_LARANJA = "#ff970f"
COR_PRETO = "#000000"
COR_CINZA_ESCURO = "#242424"
COR_CINZA_CLARO = "#DCE4EE"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Configuração da Janela Principal ---
        self.title("Automação de Redirecionamento Tray")
        self.geometry("700x500")
        self.config(bg=COR_PRETO)
        
        # Define o tema
        ctk.set_appearance_mode("Dark")

        # Fila para logs (automation_logic -> GUI)
        self.log_queue = queue.Queue()
        # Fila para decisões (GUI -> automation_logic)
        self.decision_queue = queue.Queue()
        
        self.automation_thread = None
        
        # --- Layout da GUI ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- Frame Superior (Título e Botão) ---
        self.top_frame = ctk.CTkFrame(self, fg_color=COR_PRETO)
        self.top_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        self.top_frame.grid_columnconfigure(0, weight=1)

        self.title_label = ctk.CTkLabel(self.top_frame, text="AUTOMAÇÃO REDIRECT TRAY",
                                        text_color=COR_LARANJA,
                                        font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.start_button = ctk.CTkButton(self.top_frame, text="INICIAR AUTOMAÇÃO",
                                          font=ctk.CTkFont(size=14, weight="bold"),
                                          fg_color=COR_LARANJA,
                                          text_color=COR_PRETO,
                                          hover_color="#E08000",
                                          command=self.iniciar_automacao_thread)
        self.start_button.grid(row=0, column=1, padx=10, pady=10, sticky="e")

        # --- Caixa de Logs ---
        self.log_textbox = ctk.CTkTextbox(self, fg_color=COR_CINZA_ESCURO,
                                          text_color=COR_CINZA_CLARO,
                                          font=ctk.CTkFont(family="Consolas", size=12),
                                          state="disabled") # Começa desabilitada
        self.log_textbox.grid(row=1, column=0, padx=20, pady=(10, 20), sticky="nsew")

        # Inicia o "ouvinte" da fila de logs
        self.after(100, self.processar_fila_logs)

    def log(self, message):
        """ Adiciona uma mensagem ao log de forma segura para a thread """
        self.log_textbox.configure(state="normal") # Habilita para edição
        
        # Remove a quebra de linha extra se houver
        message = message.strip() + "\n"
        
        self.log_textbox.insert("end", message)
        self.log_textbox.configure(state="disabled") # Desabilita
        self.log_textbox.see("end") # Rola até o fim

    def iniciar_automacao_thread(self):
        """ Dispara a automação em uma thread separada """
        if self.automation_thread and self.automation_thread.is_alive():
            self.log("A automação já está em execução.")
            return
            
        self.log("Iniciando thread de automação...")
        self.start_button.configure(state="disabled", text="Executando...")
        
        # Limpa as filas de decisão
        self.decision_queue = queue.Queue()

        self.automation_thread = threading.Thread(
            target=automation_logic.executar_automacao,
            args=(self.log_queue, self.decision_queue),
            daemon=True # Garante que a thread feche se a GUI fechar
        )
        self.automation_thread.start()

    def processar_fila_logs(self):
        """ Verifica a fila de logs da outra thread e atualiza a GUI """
        try:
            while True:
                # Pega a mensagem da fila (sem bloquear)
                message = self.log_queue.get_nowait()
                
                if isinstance(message, tuple) and message[0] == "SHOW_PAUSE_DIALOG":
                    # É uma tupla de comando, não um log
                    self.mostrar_popup_pausa(motivo=message[1])
                else:
                    self.log(str(message))
                    if "Thread de automação finalizada." in message:
                        self.start_button.configure(state="normal", text="INICIAR NOVAMENTE")

        except queue.Empty:
            pass # Fila vazia, tudo certo
        finally:
            # Re-agenda a verificação
            self.after(100, self.processar_fila_logs)
            
    def mostrar_popup_pausa(self, motivo):
        """ Cria a janela modal (pop-up) de pausa """
        
        popup = ctk.CTkToplevel(self)
        popup.title("Automação Pausada")
        popup.geometry("450x200")
        popup.transient(self) # Mantém no topo
        popup.grab_set() # Bloqueia a janela principal
        popup.configure(fg_color=COR_CINZA_ESCURO)

        popup.grid_columnconfigure(0, weight=1)
        
        label_motivo = ctk.CTkLabel(popup, text="🚨 AUTOMAÇÃO PAUSADA 🚨",
                                    font=ctk.CTkFont(size=16, weight="bold"), text_color=COR_LARANJA)
        label_motivo.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        label_detalhe = ctk.CTkLabel(popup, text=motivo, wraplength=400)
        label_detalhe.grid(row=1, column=0, padx=20, pady=5)
        
        # Frame dos botões
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.grid(row=2, column=0, padx=20, pady=20)

        # Funções lambda para enviar a decisão E fechar o popup
        def choice_s():
            self.decision_queue.put('S')
            popup.destroy()

        def choice_t():
            self.decision_queue.put('T')
            popup.destroy()

        def choice_f():
            self.decision_queue.put('F')
            popup.destroy()

        # Botões
        btn_s = ctk.CTkButton(btn_frame, text="Salvar e Sair (S)", fg_color=COR_LARANJA, text_color=COR_PRETO, command=choice_s)
        btn_s.grid(row=0, column=0, padx=5)
        
        btn_t = ctk.CTkButton(btn_frame, text="Tentar Novamente (T)", command=choice_t)
        btn_t.grid(row=0, column=1, padx=5)

        btn_f = ctk.CTkButton(btn_frame, text="Forçar Saída (F)", fg_color="#D00000", hover_color="#A00000", command=choice_f)
        btn_f.grid(row=0, column=2, padx=5)


if __name__ == "__main__":
    app = App()
    app.mainloop()