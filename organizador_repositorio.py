import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import sv_ttk
import threading
import subprocess
import stat
import time
import sys

class OrganizadorRepositorio:
    def __init__(self, root):
        self.root = root
        self.root.title("Organizador de Repositório")
        self.root.geometry("800x600")
        
        # Aplicar tema Sun Valley
        sv_ttk.set_theme("dark")
        
        # Variáveis
        self.diretorio_selecionado = tk.StringVar()
        
        # Frame principal
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Seleção de diretório
        dir_frame = tk.Frame(main_frame)
        dir_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(dir_frame, text="Diretório:").pack(side=tk.LEFT, padx=5)
        tk.Entry(dir_frame, textvariable=self.diretorio_selecionado, width=60).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(dir_frame, text="Selecionar", command=self.selecionar_diretorio).pack(side=tk.LEFT, padx=5)
        
        # Frame para os botões de ações
        actions_frame = tk.Frame(main_frame)
        actions_frame.pack(fill=tk.BOTH, expand=True, pady=20)
        
        # Título das ações
        tk.Label(actions_frame, text="Ações disponíveis:", font=("Arial", 12, "bold")).pack(anchor=tk.W, pady=10)
        
        # Botões para cada ação
        self.criar_botao_acao(actions_frame, "Remover arquivos PDF 'SEI_XXXXX'", self.remover_sei_pdfs)
        self.criar_botao_acao(actions_frame, "Remover arquivos 'ChecklistOcupante_'", self.remover_checklist_ocupante)
        self.criar_botao_acao(actions_frame, "Remover arquivos PDF 'N_PRXXXXXXXXX'", self.remover_pr_pdfs)
        self.criar_botao_acao(actions_frame, "Remover pastas '01_docsRecebidosEmail_Wpp'", self.remover_pastas_docs_recebidos)
        self.criar_botao_acao(actions_frame, "Mover pareceres para pastas de lotes", self.mover_pareceres)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Pronto")
        status_bar = tk.Label(main_frame, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def criar_botao_acao(self, parent, texto, comando):
        """Cria um botão para uma ação específica"""
        frame = tk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        
        btn = tk.Button(frame, text=texto, command=lambda: self.executar_acao(comando), width=50)
        btn.pack(side=tk.LEFT, padx=5)
        
    def executar_acao(self, comando):
        """Executa uma ação em uma thread separada"""
        if not self.diretorio_selecionado.get():
            messagebox.showerror("Erro", "Selecione um diretório primeiro!")
            return
            
        # Iniciar thread para não travar a interface
        threading.Thread(target=comando, daemon=True).start()
        
    def selecionar_diretorio(self):
        """Abre o diálogo para selecionar um diretório"""
        diretorio = filedialog.askdirectory()
        if diretorio:
            self.diretorio_selecionado.set(diretorio)
            self.status_var.set(f"Diretório selecionado: {diretorio}")
    
    def atualizar_status(self, mensagem):
        """Atualiza a barra de status"""
        self.status_var.set(mensagem)
        self.root.update_idletasks()
    
    def remover_sei_pdfs(self):
        """Remove arquivos PDF que seguem o padrão 'SEI_XXXXX'"""
        diretorio = self.diretorio_selecionado.get()
        padrao = re.compile(r"SEI_\d+\.\d+_\d+_\d+.*\.pdf$", re.IGNORECASE)
        removidos = 0
        
        self.atualizar_status("Removendo arquivos SEI...")
        
        for root, _, arquivos in os.walk(diretorio):
            for arquivo in arquivos:
                if padrao.search(arquivo):
                    try:
                        caminho_completo = os.path.join(root, arquivo)
                        os.remove(caminho_completo)
                        removidos += 1
                        self.atualizar_status(f"Removido: {arquivo}")
                    except Exception as e:
                        self.atualizar_status(f"Erro ao remover {arquivo}: {str(e)}")
        
        messagebox.showinfo("Concluído", f"{removidos} arquivos SEI foram removidos.")
        self.atualizar_status(f"Concluído: {removidos} arquivos SEI removidos.")
    
    def remover_checklist_ocupante(self):
        """Remove arquivos que contêm 'ChecklistOcupante_' no nome"""
        diretorio = self.diretorio_selecionado.get()
        padrao = re.compile(r"ChecklistOcupante_", re.IGNORECASE)
        removidos = 0
        
        self.atualizar_status("Removendo arquivos de checklist...")
        
        for root, _, arquivos in os.walk(diretorio):
            for arquivo in arquivos:
                if padrao.search(arquivo):
                    try:
                        caminho_completo = os.path.join(root, arquivo)
                        os.remove(caminho_completo)
                        removidos += 1
                        self.atualizar_status(f"Removido: {arquivo}")
                    except Exception as e:
                        self.atualizar_status(f"Erro ao remover {arquivo}: {str(e)}")
        
        messagebox.showinfo("Concluído", f"{removidos} arquivos de checklist foram removidos.")
        self.atualizar_status(f"Concluído: {removidos} arquivos de checklist removidos.")
    
    def remover_pr_pdfs(self):
        """Remove arquivos PDF que seguem o padrão 'N_PRXXXXXXXXX'"""
        diretorio = self.diretorio_selecionado.get()
        padrao = re.compile(r"\d+_PR\d+\.pdf$", re.IGNORECASE)
        removidos = 0
        
        self.atualizar_status("Removendo arquivos PR...")
        
        for root, _, arquivos in os.walk(diretorio):
            for arquivo in arquivos:
                if padrao.search(arquivo):
                    try:
                        caminho_completo = os.path.join(root, arquivo)
                        os.remove(caminho_completo)
                        removidos += 1
                        self.atualizar_status(f"Removido: {arquivo}")
                    except Exception as e:
                        self.atualizar_status(f"Erro ao remover {arquivo}: {str(e)}")
        
        messagebox.showinfo("Concluído", f"{removidos} arquivos PR foram removidos.")
        self.atualizar_status(f"Concluído: {removidos} arquivos PR removidos.")
    
    def remover_pastas_docs_recebidos(self):
        """Remove pastas que contenham 'docsRecebidosEmail_Wpp'"""
        diretorio = self.diretorio_selecionado.get()
        padrao_pasta = "docsRecebidosEmail_Wpp"
        removidos = 0
        falhas = 0
        
        self.atualizar_status("Removendo pastas de docs recebidos...")
        
        # Verificar pasta específica no diretório raiz
        pasta_raiz = os.path.join(diretorio, padrao_pasta)
        if os.path.exists(pasta_raiz) and os.path.isdir(pasta_raiz):
            if self.remover_pasta_com_retry(pasta_raiz):
                removidos += 1
                self.atualizar_status(f"Removida pasta na raiz: {pasta_raiz}")
            else:
                falhas += 1
                self.atualizar_status(f"Falha persistente ao remover pasta: {pasta_raiz}")
        
        # Coletar todos os caminhos para evitar problemas de modificação durante a iteração
        pastas_para_remover = []
        
        for root, dirs, _ in os.walk(diretorio, topdown=True):
            dirs_copy = dirs.copy()  # Criar uma cópia para evitar modificação durante iteração
            for dir_name in dirs_copy:
                if padrao_pasta in dir_name:
                    caminho_completo = os.path.join(root, dir_name)
                    pastas_para_remover.append(caminho_completo)
        
        # Agora remover as pastas
        for caminho in pastas_para_remover:
            if self.remover_pasta_com_retry(caminho):
                removidos += 1
                self.atualizar_status(f"Removida pasta: {caminho}")
            else:
                falhas += 1
                self.atualizar_status(f"Falha persistente ao remover pasta: {caminho}")
        
        if falhas > 0:
            messagebox.showinfo("Concluído com avisos", 
                               f"{removidos} pastas foram removidas. {falhas} pastas não puderam ser removidas devido a restrições de acesso.\n\n"
                               f"Você pode tentar remover manualmente ou executar este programa como administrador.")
        else:
            messagebox.showinfo("Concluído", f"{removidos} pastas de docs recebidos foram removidas.")
        
        self.atualizar_status(f"Concluído: {removidos} pastas removidas, {falhas} falhas.")
    
    def remover_pasta_com_retry(self, caminho_pasta, max_tentativas=3):
        """Tenta remover uma pasta com várias abordagens e tentativas"""
        # Primeira tentativa - método padrão
        for tentativa in range(max_tentativas):
            try:
                shutil.rmtree(caminho_pasta)
                return True
            except PermissionError:
                if tentativa < max_tentativas - 1:
                    self.atualizar_status(f"Erro de permissão, tentando abordagem alternativa {tentativa+1}...")
                    
                    # Tentar alterar permissões dos arquivos
                    try:
                        for root, dirs, files in os.walk(caminho_pasta):
                            for arquivo in files:
                                caminho_arquivo = os.path.join(root, arquivo)
                                try:
                                    os.chmod(caminho_arquivo, stat.S_IWRITE)
                                except:
                                    pass
                        
                        # Dar um pequeno intervalo antes de tentar novamente
                        time.sleep(1)
                    except:
                        pass
                    
            except Exception as e:
                self.atualizar_status(f"Erro ao remover pasta {caminho_pasta}: {str(e)}")
                return False
        
        # Se todas as tentativas acima falharem, tenta com comando do sistema
        try:
            if sys.platform == "win32":
                # No Windows, tenta usar o comando rd
                resultado = subprocess.run(['rd', '/s', '/q', caminho_pasta], 
                                          shell=True, 
                                          stdout=subprocess.PIPE, 
                                          stderr=subprocess.PIPE)
                return resultado.returncode == 0
            else:
                # Em sistemas Unix, tenta usar rm -rf
                resultado = subprocess.run(['rm', '-rf', caminho_pasta], 
                                          stdout=subprocess.PIPE, 
                                          stderr=subprocess.PIPE)
                return resultado.returncode == 0
        except Exception as e:
            self.atualizar_status(f"Erro ao usar comando do sistema para remover {caminho_pasta}: {str(e)}")
            return False
    
    def mover_pareceres(self):
        """Move pareceres e relatórios PDF para as respectivas pastas dos lotes"""
        diretorio = self.diretorio_selecionado.get()
        
        # Expressão regular principal - atualizada para capturar formato sem underscore entre ParecerConclusivo e PA
        padrao_parecer = re.compile(r"L(\d+)(?:e\d+)*(?:_\d+)?_(?:ParecerConclusivo(?:Ocupante)?|Relatorio)_?PA([^_]+).*\.pdf$", re.IGNORECASE)
        
        # Expressão de backup para casos mais difíceis
        padrao_backup = re.compile(r"L(\d+)(?:e\d+)*(?:_\d+)?_(?:ParecerConclusivo|Relatorio).*_?PA?([A-Za-z]+).*\.pdf$", re.IGNORECASE)
        
        # Casos especiais que precisamos tratar individualmente - usando apenas parte do nome para ser mais flexível
        casos_especiais = {
            "L149_ParecerConclusivo_PAEDUARDORADUAN": {"lote": "149", "pa": "EDUARDORADUAN"}
        }
        
        movidos = 0
        nao_movidos = 0
        
        self.atualizar_status("Movendo pareceres e relatórios PDF para pastas de lotes...")
        
        for root, _, arquivos in os.walk(diretorio):
            for arquivo in arquivos:
                # Processa apenas arquivos PDF
                if not arquivo.lower().endswith(".pdf"):
                    continue
                
                # Exibir informações de debug para cada arquivo PDF que começa com L
                if arquivo.startswith("L"):
                    self.atualizar_status(f"Analisando arquivo: {arquivo}")
                    
                    # Verificar se o arquivo corresponde a algum caso especial (verificação parcial)
                    caso_especial_encontrado = False
                    for caso, valores in casos_especiais.items():
                        if caso in arquivo:
                            self.atualizar_status(f"Caso especial encontrado: {caso} em {arquivo}")
                            num_lote = valores["lote"]
                            nome_pa = valores["pa"]
                            caso_especial_encontrado = True
                            break
                            
                    if not caso_especial_encontrado:
                        # Tentar fazer a correspondência com a expressão regular principal
                        match = padrao_parecer.search(arquivo)
                        if match:
                            num_lote = match.group(1)
                            nome_pa = match.group(2)
                            self.atualizar_status(f"Padrão principal corresponde: {arquivo}")
                        else:
                            # Tentar com o padrão de backup
                            match_backup = padrao_backup.search(arquivo)
                            if match_backup:
                                num_lote = match_backup.group(1)
                                nome_pa = match_backup.group(2)
                                self.atualizar_status(f"Padrão de backup corresponde: {arquivo}")
                            else:
                                self.atualizar_status(f"⚠️ Arquivo PDF não reconhecido por nenhum padrão: {arquivo}")
                                # Para debug, vamos mostrar o que os padrões estão tentando capturar
                                self.atualizar_status(f"Debug - Tentando: {padrao_parecer.pattern} ou {padrao_backup.pattern}")
                                continue
                    
                    # Debug - mostrar o que foi capturado
                    self.atualizar_status(f"Capturado: Lote={num_lote}, PA={nome_pa} do arquivo {arquivo}")
                    
                    # Padronizar número do lote (ex: 3 -> Lote0003, 091 -> Lote0091, 032 -> Lote0032)
                    lote_formatado = f"Lote{int(num_lote):04d}"
                    
                    # Buscar pasta destino
                    caminho_origem = os.path.join(root, arquivo)
                    pasta_destino = None
                    
                    # Listar todas as pastas encontradas para diagnóstico
                    pastas_encontradas = []
                    
                    # Tentar encontrar pasta com base no nome PA (mais flexível com case-insensitive)
                    for raiz, _, _ in os.walk(diretorio):
                        nome_pa_lower = nome_pa.lower()
                        raiz_lower = raiz.lower()
                        
                        # Registrar todas as pastas que contêm o lote
                        if lote_formatado.lower() in raiz_lower:
                            pastas_encontradas.append(raiz)
                        
                        # Verificar se o nome do PA está na pasta e se a pasta contém o lote
                        if nome_pa_lower in raiz_lower and "lote" in raiz_lower and lote_formatado.lower() in raiz_lower:
                            pasta_destino = raiz
                            self.atualizar_status(f"Pasta encontrada: {raiz}")
                            break
                    
                    if not pasta_destino and pastas_encontradas:
                        # Se não encontrou pasta com PA+lote, mas encontrou pastas com o lote,
                        # use a primeira pasta do lote encontrada
                        pasta_destino = pastas_encontradas[0]
                        self.atualizar_status(f"Usando pasta alternativa: {pasta_destino}")
                    
                    if pasta_destino:
                        try:
                            caminho_origem = os.path.join(root, arquivo)
                            caminho_destino = os.path.join(pasta_destino, arquivo)
                            
                            # Verificar se o arquivo já existe no destino
                            if os.path.exists(caminho_destino):
                                self.atualizar_status(f"Arquivo já existe no destino: {caminho_destino}")
                                
                                # Comparar as datas de modificação
                                data_origem = os.path.getmtime(caminho_origem)
                                data_destino = os.path.getmtime(caminho_destino)
                                
                                nome_base, extensao = os.path.splitext(arquivo)
                                
                                # Se o arquivo de destino é mais recente, renomeá-lo com "_2" antes da extensão
                                if data_destino > data_origem:
                                    # Renomear o arquivo de destino (mais recente) para incluir _2
                                    novo_nome = f"{nome_base}_2{extensao}"
                                    novo_caminho_destino = os.path.join(pasta_destino, novo_nome)
                                    
                                    self.atualizar_status(f"Arquivo no destino é mais recente. Renomeando destino para: {novo_nome}")
                                    
                                    # Verificar se o novo nome também já existe
                                    if os.path.exists(novo_caminho_destino):
                                        self.atualizar_status(f"O nome {novo_nome} já está em uso na pasta de destino.")
                                        i = 3
                                        while True:
                                            novo_nome = f"{nome_base}_{i}{extensao}"
                                            novo_caminho_destino = os.path.join(pasta_destino, novo_nome)
                                            if not os.path.exists(novo_caminho_destino):
                                                break
                                            i += 1
                                        self.atualizar_status(f"Usando nome alternativo: {novo_nome}")
                                    
                                    # Renomear o arquivo de destino
                                    try:
                                        os.rename(caminho_destino, novo_caminho_destino)
                                        self.atualizar_status(f"Arquivo no destino renomeado para: {novo_nome}")
                                        
                                        # Agora mover o arquivo original
                                        shutil.copy2(caminho_origem, caminho_destino)
                                        os.remove(caminho_origem)
                                        movidos += 1
                                        self.atualizar_status(f"Arquivo original (mais antigo) movido para: {caminho_destino}")
                                    except Exception as e:
                                        self.atualizar_status(f"Erro ao renomear arquivo no destino: {str(e)}")
                                        nao_movidos += 1
                                else:
                                    # Se o arquivo da origem é mais recente, renomeá-lo com "_2" antes da extensão
                                    novo_nome = f"{nome_base}_2{extensao}"
                                    novo_caminho_destino = os.path.join(pasta_destino, novo_nome)
                                    
                                    self.atualizar_status(f"Arquivo de origem é mais recente. Renomeando origem para: {novo_nome}")
                                    
                                    # Verificar se o novo nome também já existe
                                    if os.path.exists(novo_caminho_destino):
                                        self.atualizar_status(f"O arquivo renomeado já existe: {novo_nome}")
                                        i = 3
                                        while True:
                                            novo_nome = f"{nome_base}_{i}{extensao}"
                                            novo_caminho_destino = os.path.join(pasta_destino, novo_nome)
                                            if not os.path.exists(novo_caminho_destino):
                                                break
                                            i += 1
                                        self.atualizar_status(f"Usando nome alternativo: {novo_nome}")
                                    
                                    shutil.copy2(caminho_origem, novo_caminho_destino)
                                    os.remove(caminho_origem)
                                    movidos += 1
                                    self.atualizar_status(f"Movido com novo nome: {novo_nome}")
                                continue
                                
                            # Se não existe arquivo duplicado, move normalmente
                            shutil.copy2(caminho_origem, caminho_destino)
                            os.remove(caminho_origem)
                            movidos += 1
                            self.atualizar_status(f"Movido: {arquivo} para {pasta_destino}")
                        except Exception as e:
                            self.atualizar_status(f"Erro ao mover {arquivo}: {str(e)}")
                            nao_movidos += 1
                    else:
                        if pastas_encontradas:
                            self.atualizar_status(f"Pastas do lote {lote_formatado} encontradas, mas nenhuma corresponde ao PA {nome_pa}:")
                            for pasta in pastas_encontradas:
                                self.atualizar_status(f"  - {pasta}")
                        else:
                            self.atualizar_status(f"Nenhuma pasta encontrada para o lote {lote_formatado}")
                        self.atualizar_status(f"Não foi encontrada pasta para o arquivo: {arquivo}")
                        nao_movidos += 1
        
        messagebox.showinfo("Concluído", f"{movidos} arquivos PDF foram movidos. {nao_movidos} não puderam ser movidos.")
        self.atualizar_status(f"Concluído: {movidos} arquivos PDF movidos, {nao_movidos} não movidos.")

if __name__ == "__main__":
    root = tk.Tk()
    app = OrganizadorRepositorio(root)
    root.mainloop() 