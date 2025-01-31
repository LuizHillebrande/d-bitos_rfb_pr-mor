import customtkinter as ctk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import messagebox
from PIL import Image

# Credenciais padrão
DEFAULT_EMAIL = "legal@contabilprimor.com.br"
DEFAULT_SENHA = "q7ne5k0la0VJ"

# Função para iniciar o WebDriver com as credenciais
def iniciar_webdriver(email, senha):
    try:
        driver = webdriver.Chrome()
        driver.get('https://app.digiliza.com.br/login')

        input_email = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@id='email']"))
        )
        input_email.send_keys(email)

        input_senha = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@id='password']"))
        )
        input_senha.send_keys(senha)

        botao_login = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
        )
        botao_login.click()

        messagebox.showinfo("Sucesso", "Login realizado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha no login: {e}")

# Criando a interface gráfica
def criar_interface():
    def fazer_login():
        email = email_entry.get()
        senha = senha_entry.get()
        iniciar_webdriver(email, senha)

    def toggle_password():
        """ Alterna entre exibir ou ocultar a senha """
        if senha_entry.cget("show") == "*":
            senha_entry.configure(show="")
            toggle_button.configure(image=eye_open)
        else:
            senha_entry.configure(show="*")
            toggle_button.configure(image=eye_closed)

    app = ctk.CTk()
    app.title("Login - Digiliza")
    app.geometry(f"{app.winfo_screenwidth()}x{app.winfo_screenheight()}+0+0")  # Tela cheia
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    # Fundo estilizado
    bg_frame = ctk.CTkFrame(master=app, fg_color="#1E1E1E")
    bg_frame.pack(fill="both", expand=True)

    # Container do login centralizado
    frame = ctk.CTkFrame(master=bg_frame, width=400, height=500, corner_radius=20, fg_color="#2E2E2E")
    frame.place(relx=0.5, rely=0.5, anchor="center")

    titulo = ctk.CTkLabel(master=frame, text="Login no Digiliza", font=("Arial", 24, "bold"), text_color="#00A3FF")
    titulo.pack(pady=20)

    email_label = ctk.CTkLabel(master=frame, text="E-mail:", text_color="white")
    email_label.pack()
    email_entry = ctk.CTkEntry(master=frame, width=300, height=40, corner_radius=10)
    email_entry.insert(0, DEFAULT_EMAIL)
    email_entry.pack(pady=5)

    senha_label = ctk.CTkLabel(master=frame, text="Senha:", text_color="white")
    senha_label.pack()

    # Campo de senha com botão de exibição
    senha_frame = ctk.CTkFrame(master=frame, fg_color="transparent")
    senha_frame.pack()

    senha_entry = ctk.CTkEntry(master=senha_frame, width=260, height=40, corner_radius=10, show="*")
    senha_entry.insert(0, DEFAULT_SENHA)
    senha_entry.pack(side="left", pady=5)

    # Ícones para alternar a visibilidade da senha
    eye_open = ctk.CTkImage(light_image=Image.open("imgs\eye_open.png"), size=(24, 24))
    eye_closed = ctk.CTkImage(light_image=Image.open("imgs\eye_closed.png"), size=(24, 24))

    toggle_button = ctk.CTkButton(master=senha_frame, width=40, height=40, text="", image=eye_closed,
                                  fg_color="transparent", hover_color="#444", command=toggle_password)
    toggle_button.pack(side="right", padx=5)

    # Botão estilizado
    def on_enter(e):
        login_button.configure(fg_color="#0088CC")  # Azul mais vibrante ao passar o mouse

    def on_leave(e):
        login_button.configure(fg_color="#00A3FF")  # Retorna ao azul original

    login_button = ctk.CTkButton(master=frame, text="Login", command=fazer_login, 
                                 width=300, height=50, corner_radius=10, fg_color="#00A3FF", text_color="white",
                                 hover_color="#0088CC")
    login_button.pack(pady=20)

    login_button.bind("<Enter>", on_enter)
    login_button.bind("<Leave>", on_leave)

    app.mainloop()

# Iniciar a interface
criar_interface()
