import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import subprocess
import sys
import os
import string
import secrets

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def run_powershell(command):
    process = subprocess.run(
        ["powershell", "-Command", command],
        capture_output=True,
        text=True,
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    return process

def gerar_senha_segura():
    caracteres_especiais = "!@#$%*&?"
    alfabeto_completo = string.ascii_letters + string.digits + caracteres_especiais
    
    while True:
        senha = ''.join(secrets.choice(alfabeto_completo) for i in range(10))
        if (any(c.islower() for c in senha) and
            any(c.isupper() for c in senha) and
            any(c.isdigit() for c in senha) and
            any(c in caracteres_especiais for c in senha)):
            break 

    entry_pass.configure(show="") 
    entry_pass.delete(0, 'end')
    entry_pass.insert(0, senha)
    
    app.clipboard_clear()
    app.clipboard_append(senha)
    app.update() 
    
    messagebox.showinfo("Senha Gerada", "Senha gerada e COPIADA com sucesso!\n\nBasta clicar no botão 'Resetar Senha' e depois dar um Ctrl+V no chamado.")

def consult_user():
    search_term = entry_user.get().strip()
    if not search_term:
        messagebox.showwarning("Aviso", "Por favor, digite o nome ou login.")
        return

    # Limpando os campos para mostrar que está carregando
    lbl_status_ativo.configure(text="Status da Conta: Consultando...", text_color="gray")
    lbl_status_expirada.configure(text="Senha Expirada: -", text_color="gray")
    lbl_status_bloqueada.configure(text="Conta Bloqueada: -", text_color="gray")
    lbl_status_tentativas.configure(text="Tentativas Erradas: -", text_color="gray")
    lbl_status_hora_erro.configure(text="Hora do Bloqueio: -", text_color="gray")
    lbl_status_troca.configure(text="Última Troca: -", text_color="gray")
    app.update()

    # Usando ANR para busca ultrarrápida e puxando BadLogonCount / AccountLockoutTime
    cmd = (
        f"$search = '{search_term}'; "
        f"$users = @(Get-ADUser -Filter \"anr -eq '$search'\" "
        f"-Properties Enabled, PasswordExpired, PasswordLastSet, LockedOut, BadLogonCount, AccountLockoutTime -ErrorAction SilentlyContinue); "
        f"if ($users.Count -eq 0) {{ Write-Output 'ERRO_NENHUM' }} "
        f"elseif ($users.Count -gt 1) {{ Write-Output 'ERRO_MULTIPLOS' }} "
        f"else {{ "
        f"$u = $users[0]; "
        f"Write-Output \"ENCONTRADO|$($u.SamAccountName)|Ativo:$($u.Enabled)|Expirada:$($u.PasswordExpired)|UltimaTroca:$($u.PasswordLastSet)|Bloqueada:$($u.LockedOut)|Tentativas:$($u.BadLogonCount)|HoraBloqueio:$($u.AccountLockoutTime)\" "
        f"}}"
    )
    
    result = run_powershell(cmd)
    output = result.stdout.strip()

    if "ERRO_NENHUM" in output or not output:
        messagebox.showerror("Erro", f"Nenhum usuário encontrado para '{search_term}'.")
        lbl_status_ativo.configure(text="Status da Conta: Não encontrado", text_color="#ef476f")
        return
    
    elif "ERRO_MULTIPLOS" in output:
        messagebox.showwarning("Atenção", "Vários usuários encontrados.\nDigite nome e sobrenome.")
        lbl_status_ativo.configure(text="Status: Múltiplos encontrados", text_color="#ffd166")
        return

    try:
        # Tratamento mais seguro dos dados separados por pipe (|)
        parts = output.split("|")
        real_samaccountname = parts[1]
        ativo_str = parts[2].replace("Ativo:", "").strip()
        expirada_str = parts[3].replace("Expirada:", "").strip()
        troca_str = parts[4].replace("UltimaTroca:", "").strip()
        bloqueada_str = parts[5].replace("Bloqueada:", "").strip()
        tentativas_str = parts[6].replace("Tentativas:", "").strip()
        hora_bloqueio_str = parts[7].replace("HoraBloqueio:", "").strip()

        entry_user.delete(0, 'end')
        entry_user.insert(0, real_samaccountname)

        # Regras Visuais
        txt_ativo = "ATIVO" if ativo_str == "True" else "DESATIVADO"
        cor_ativo = "#06d6a0" if ativo_str == "True" else "#ef476f"
        
        txt_expirada = "SIM" if expirada_str == "True" else "NÃO"
        cor_expirada = "#ef476f" if expirada_str == "True" else "#06d6a0"

        txt_bloqueada = "SIM" if bloqueada_str == "True" else "NÃO"
        cor_bloqueada = "#ef476f" if bloqueada_str == "True" else "#06d6a0"
        
        txt_tentativas = tentativas_str if tentativas_str else "0"
        # Destaca em amarelo/laranja se tiver erros na senha, mas ainda não bloqueou
        cor_tentativas = "#f4a261" if (int(txt_tentativas) > 0 and bloqueada_str != "True") else ("black", "white")

        txt_hora_bloqueio = hora_bloqueio_str if hora_bloqueio_str else "Nenhum registro recente"
        txt_troca = troca_str if troca_str else "Desconhecida"

        # Aplicando na Interface
        lbl_status_ativo.configure(text=f"Status da Conta: {txt_ativo}", text_color=cor_ativo)
        lbl_status_expirada.configure(text=f"Senha Expirada: {txt_expirada}", text_color=cor_expirada)
        lbl_status_bloqueada.configure(text=f"Conta Bloqueada: {txt_bloqueada}", text_color=cor_bloqueada)
        
        # Os novos campos do "Dedo Duro"
        lbl_status_tentativas.configure(text=f"Tentativas Erradas: {txt_tentativas}", text_color=cor_tentativas)
        lbl_status_hora_erro.configure(text=f"Hora do Bloqueio: {txt_hora_bloqueio}", text_color=("black", "white"))
        
        lbl_status_troca.configure(text=f"Última Troca: {txt_troca}", text_color=("black", "white"))

    except Exception as e:
        messagebox.showerror("Erro Crítico", f"Falha ao processar:\n{e}")

def unlock_user():
    username = entry_user.get().strip()
    if not username:
        messagebox.showwarning("Aviso", "Por favor, consulte um usuário primeiro.")
        return

    cmd = (
        f"try {{ "
        f"$user = Get-ADUser -Identity '{username}' -Properties LockedOut -ErrorAction Stop; "
        f"if ($user.LockedOut) {{ Unlock-ADAccount -Identity '{username}'; Write-Output 'BLOQUEADO' }} "
        f"else {{ Write-Output 'LIVRE' }} "
        f"}} catch {{ Write-Output 'ERRO_USUARIO' }}"
    )

    result = run_powershell(cmd)
    output = result.stdout.strip()

    if "BLOQUEADO" in output:
        messagebox.showinfo("Sucesso", f"Usuário '{username}' foi liberado!")
        consult_user() 
    elif "LIVRE" in output:
        messagebox.showinfo("Informação", f"A conta '{username}' NÃO estava bloqueada.")
    elif "ERRO_USUARIO" in output:
        messagebox.showerror("Erro", f"Usuário '{username}' não encontrado.")
    else:
        messagebox.showerror("Erro Crítico", f"Falha inesperada:\n{result.stderr}")

def reset_password():
    username = entry_user.get().strip()
    password = entry_pass.get().strip()
    force_change = var_force_change.get()

    if not username or not password:
        messagebox.showwarning("Aviso", "Preencha o usuário e a nova senha.")
        return

    ps_force_change = "$true" if force_change else "$false"
    cmd = (
        f'$pw = ConvertTo-SecureString "{password}" -AsPlainText -Force; '
        f'Set-ADAccountPassword -Identity "{username}" -NewPassword $pw -Reset; '
        f'Set-ADUser -Identity "{username}" -ChangePasswordAtLogon {ps_force_change}'
    )

    result = run_powershell(cmd)

    if result.returncode == 0:
        messagebox.showinfo("Sucesso", f"Senha alterada para '{username}'.")
        entry_pass.delete(0, 'end')
        entry_pass.configure(show="*")
        consult_user() 
    else:
        messagebox.showerror("Erro", f"Falha ao redefinir a senha:\n{result.stderr}")

# --- Configuração da Interface Moderna (CustomTkinter) ---
app = ctk.CTk()
app.title("Painel AD - Suporte N1")
app.geometry("450x630") # Aumentei a altura para caber o dedo duro
app.resizable(False, False)

try:
    app.iconbitmap(resource_path("logo.ico"))
except Exception:
    pass 

# Título Superior
titulo = ctk.CTkLabel(app, text="Gestão de Contas AD", font=ctk.CTkFont(size=20, weight="bold"))
titulo.pack(pady=(20, 10))

# Frame de Busca
frame_busca = ctk.CTkFrame(app, fg_color="transparent")
frame_busca.pack(pady=5, padx=20, fill="x")

ctk.CTkLabel(frame_busca, text="Nome, Sobrenome ou Login:", font=ctk.CTkFont(size=13)).pack(anchor="w")
entry_user = ctk.CTkEntry(frame_busca, placeholder_text="Ex: maria.silva", width=250)
entry_user.pack(side="left", pady=5)

btn_consult = ctk.CTkButton(frame_busca, text="Buscar", command=consult_user, width=100)
btn_consult.pack(side="right", padx=(10, 0))

# Frame de Informações
frame_info = ctk.CTkFrame(app)
frame_info.pack(pady=15, padx=20, fill="x")

ctk.CTkLabel(frame_info, text="Resultados da Consulta", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(10, 5))

lbl_status_ativo = ctk.CTkLabel(frame_info, text="Status da Conta: -", font=ctk.CTkFont(size=13))
lbl_status_ativo.pack(anchor="w", padx=15, pady=2)

lbl_status_expirada = ctk.CTkLabel(frame_info, text="Senha Expirada: -", font=ctk.CTkFont(size=13))
lbl_status_expirada.pack(anchor="w", padx=15, pady=2)

lbl_status_bloqueada = ctk.CTkLabel(frame_info, text="Conta Bloqueada: -", font=ctk.CTkFont(size=13))
lbl_status_bloqueada.pack(anchor="w", padx=15, pady=2)

# --- INÍCIO DOS CAMPOS DEDO DURO ---
lbl_status_tentativas = ctk.CTkLabel(frame_info, text="Tentativas Erradas: -", font=ctk.CTkFont(size=13))
lbl_status_tentativas.pack(anchor="w", padx=15, pady=2)

lbl_status_hora_erro = ctk.CTkLabel(frame_info, text="Hora do Bloqueio: -", font=ctk.CTkFont(size=13))
lbl_status_hora_erro.pack(anchor="w", padx=15, pady=2)
# --- FIM DOS CAMPOS DEDO DURO ---

lbl_status_troca = ctk.CTkLabel(frame_info, text="Última Troca: -", font=ctk.CTkFont(size=13))
lbl_status_troca.pack(anchor="w", padx=15, pady=(2, 10))

# Frame de Ações (Senha)
frame_acoes = ctk.CTkFrame(app, fg_color="transparent")
frame_acoes.pack(pady=5, padx=20, fill="x")

ctk.CTkLabel(frame_acoes, text="Nova Senha Temporária:", font=ctk.CTkFont(size=13)).pack(anchor="w")

frame_senha_input = ctk.CTkFrame(frame_acoes, fg_color="transparent")
frame_senha_input.pack(pady=5, anchor="w")

entry_pass = ctk.CTkEntry(frame_senha_input, show="*", width=250)
entry_pass.pack(side="left")

btn_gerar = ctk.CTkButton(frame_senha_input, text="🎲 Gerar", command=gerar_senha_segura, width=70, fg_color="#457b9d", hover_color="#1d3557")
btn_gerar.pack(side="left", padx=(10, 0))

var_force_change = ctk.BooleanVar(value=True)
chk_force = ctk.CTkCheckBox(frame_acoes, text="Exigir troca no próximo logon", variable=var_force_change)
chk_force.pack(pady=10, anchor="w")

# Botões de Execução
frame_botoes = ctk.CTkFrame(app, fg_color="transparent")
frame_botoes.pack(pady=10)

btn_unlock = ctk.CTkButton(frame_botoes, text="🔓 Desbloquear", command=unlock_user, fg_color="#f4a261", hover_color="#e76f51", text_color="white")
btn_unlock.grid(row=0, column=0, padx=10)

btn_reset = ctk.CTkButton(frame_botoes, text="🔑 Resetar Senha", command=reset_password, fg_color="#e63946", hover_color="#d62828", text_color="white")
btn_reset.grid(row=0, column=1, padx=10)

app.mainloop()