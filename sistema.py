import os
import sys
import json
import shutil
import logging
from pathlib import Path
from tkinter import filedialog, messagebox
from datetime import datetime
from PIL import Image

import customtkinter as ctk

from modules.photo_handler import PhotoHandler, PhotoInfo
from modules.template_handler import TemplateHandler
from modules.obras_module import ObraHandler

# ─── Configuração de Logging ───────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    _dir_app = Path(sys.executable).parent.resolve()
else:
    _dir_app = Path(__file__).parent.resolve()
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
    handlers=[
        logging.FileHandler(str(_dir_app / 'sistema_poda.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('Sistema')


class TextHandler(logging.Handler):
    """Handler de logging que direciona mensagens para o widget de texto."""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.setFormatter(logging.Formatter(
            '%(asctime)s [%(levelname)s] %(message)s'
        ))

    def emit(self, record):
        msg = self.format(record)
        self.text_widget.after(0, self._append, msg)

    def _append(self, msg):
        try:
            self.text_widget.configure(state='normal')
            self.text_widget.insert('end', msg + '\n')
            self.text_widget.see('end')
            self.text_widget.configure(state='disabled')
        except Exception:
            pass


class SistemaPodaApp(ctk.CTk):
    """
    Aplicação principal do Sistema de Relatórios de Poda.
    """

    def __init__(self):
        super().__init__()

        # ─── Configuração da Janela ──────────────────────────────────────────
        self.title('Sistema Automático de Relatórios - Equatorial')
        self.geometry('680x680')
        self.resizable(False, False)

                # ─── Diretórios do Projeto ───────────────────────────────────────────
        if getattr(sys, 'frozen', False):
            self.dir_atual = Path(sys.executable).parent.resolve()
        else:
            self.dir_atual = Path(__file__).parent.resolve()
        self.pasta_modelo = self.dir_atual / 'modelo'
        self.pasta_relatorios = self.dir_atual / 'relatorios_gerados'

        tema_inicial = self._carregar_config_tema()
        ctk.set_appearance_mode(tema_inicial)
        ctk.set_default_color_theme('green')

        # Garante que os diretórios existam
        self.pasta_relatorios.mkdir(parents=True, exist_ok=True)

        # Verifica e cria a estrutura necessária (modelo, etc.)
        self._verificar_estrutura()

        # ─── Variáveis de Estado ─────────────────────────────────────────────
        self.pasta_fotos_selecionada = None
        self.photo_handler = None
        self.processando = False
        self.tema_atual = tema_inicial

        # ─── Estado do Menu / Módulos ───────────────────────────────────────
        self.modo_atual = 'menu'
        self.pasta_obra_selecionada = None
        self.obra_handler = None

        # ─── Construção da Interface ─────────────────────────────────────────
        self._construir_interface()

        # ─── Adiciona handler de log no Text Widget ──────────────────────────
        self._log_handler = TextHandler(self.log_text)
        logging.getLogger().addHandler(self._log_handler)

        self.log_text.configure(state='normal')
        self.log_text.insert('end', 'Sistema Automático de Relatórios\n')
        self.log_text.insert('end', 'FK Engenharia e Serviços LTDA\n')
        self.log_text.insert('end', '─' * 60 + '\n')
        self.log_text.insert('end', 'Pronto. Aguardando...\n')
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

        # Centraliza a janela na tela
        self.after(100, self._centralizar)

        logger.info('Sistema iniciado')

    def _resource_path(self, relative_path):
        """Obtém caminho absoluto do recurso, compatível com PyInstaller."""
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = self.dir_atual
        return str(Path(base_path) / relative_path)

    def _verificar_estrutura(self):
        """Verifica e cria automaticamente a estrutura necessária do sistema."""
        # Cria a pasta modelo se não existir
        if not self.pasta_modelo.exists():
            self.pasta_modelo.mkdir(parents=True, exist_ok=True)
            logger.info('Pasta modelo criada.')

        # Verifica se o template já existe na pasta modelo
        templates = list(self.pasta_modelo.glob('*_Relatorio_de_Podas_FK_Eng_PHB.docx'))
        if not templates:
            origem_modelo = Path(self._resource_path('modelo'))
            templates_origem = list(origem_modelo.glob('*_Relatorio_de_Podas_FK_Eng_PHB.docx'))
            if templates_origem:
                shutil.copy2(str(templates_origem[0]), str(self.pasta_modelo / templates_origem[0].name))
                logger.info('Modelo padrão copiado.')

        logger.info('Estrutura do sistema verificada.')

    def _carregar_config_tema(self):
        """Carrega a preferência de tema do arquivo config.json."""
        config_path = self.dir_atual / 'config.json'
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    return json.load(f).get('theme', 'light')
            except Exception:
                pass
        return 'light'

    def _salvar_config_tema(self, tema):
        """Salva a preferência de tema no arquivo config.json."""
        config_path = self.dir_atual / 'config.json'
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump({'theme': tema}, f, indent=2)
        except Exception as e:
            logger.warning(f'Erro ao salvar tema: {e}')

    def _centralizar(self):
        """Centraliza a janela no monitor."""
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f'+{x}+{y}')

    def _construir_interface(self):
        self._construir_cabecalho()

        self.frame_conteudo = ctk.CTkFrame(self, corner_radius=10)
        self.frame_conteudo.pack(fill='both', expand=True, padx=20, pady=(5, 5))

        self._construir_toolbar()
        self._construir_menu_principal()
        self._construir_interface_podas()
        self._construir_interface_obras()

        self._construir_log()
        self._construir_rodape()

        self._mostrar_menu()

    # ─── Toolbar ─────────────────────────────────────────────────────────────
    def _construir_toolbar(self):
        self.frame_toolbar = ctk.CTkFrame(self.frame_conteudo, fg_color='transparent', height=26)
        self.btn_voltar = ctk.CTkButton(
            self.frame_toolbar,
            text='← Voltar ao Menu',
            font=ctk.CTkFont(size=11),
            width=110,
            height=24,
            fg_color=('gray70', 'gray40'),
            hover_color=('gray50', 'gray30'),
            text_color=('#333333', '#CCCCCC'),
            corner_radius=3,
            command=self._voltar_menu
        )
        self.btn_voltar.pack(side='left', padx=(3, 0))

    # ─── Menu Principal ──────────────────────────────────────────────────────
    def _construir_menu_principal(self):
        self.frame_menu = ctk.CTkFrame(self.frame_conteudo, fg_color='transparent')
        self.frame_menu.grid_rowconfigure(0, weight=1)
        self.frame_menu.grid_rowconfigure(4, weight=1)
        self.frame_menu.grid_columnconfigure(0, weight=1)

        label_modulo = ctk.CTkLabel(
            self.frame_menu,
            text='Selecione o módulo desejado:',
            font=ctk.CTkFont(size=13),
            text_color=('gray50', 'gray60')
        )
        label_modulo.grid(row=1, column=0, pady=(0, 20))

        btn_podas = ctk.CTkButton(
            self.frame_menu,
            text='Relatório de Podas',
            font=ctk.CTkFont(size=13),
            height=38,
            width=280,
            fg_color='#2E7D32',
            hover_color='#1B5E20',
            text_color='#FFFFFF',
            corner_radius=4,
            command=self._mostrar_podas
        )
        btn_podas.grid(row=2, column=0, pady=(0, 8))

        btn_obras = ctk.CTkButton(
            self.frame_menu,
            text='Relatório de Obras',
            font=ctk.CTkFont(size=13),
            height=38,
            width=280,
            fg_color='#1B3A5C',
            hover_color='#2C5F8A',
            text_color='#FFFFFF',
            corner_radius=4,
            command=self._mostrar_obras
        )
        btn_obras.grid(row=3, column=0, pady=(0, 0))

    # ─── Interface de Podas (extraída sem alterações) ─────────────────────────
    def _construir_interface_podas(self):
        self.frame_podas = ctk.CTkFrame(self.frame_conteudo, fg_color='transparent')

        # ─── ID do Projeto ───────────────────────────────────────────────────
        label_id = ctk.CTkLabel(
            self.frame_podas,
            text='Nota do Projeto:',
            font=ctk.CTkFont(size=14),
            anchor='w'
        )
        label_id.pack(fill='x', padx=30, pady=(15, 0))

        self.entry_id = ctk.CTkEntry(
            self.frame_podas,
            placeholder_text='Ex: 440100001',
            font=ctk.CTkFont(size=14),
            height=35
        )
        self.entry_id.pack(fill='x', padx=30, pady=(3, 10))

        # ─── Pasta das Fotos ─────────────────────────────────────────────────
        label_pasta = ctk.CTkLabel(
            self.frame_podas,
            text='Pasta das Fotos:',
            font=ctk.CTkFont(size=14),
            anchor='w'
        )
        label_pasta.pack(fill='x', padx=30)

        frame_pasta = ctk.CTkFrame(self.frame_podas, fg_color='transparent')
        frame_pasta.pack(fill='x', padx=30, pady=(3, 10))

        self.entry_pasta = ctk.CTkEntry(
            frame_pasta,
            placeholder_text='Selecione a pasta com as fotos...',
            font=ctk.CTkFont(size=12),
            height=35,
            state='disabled'
        )
        self.entry_pasta.pack(side='left', fill='x', expand=True, padx=(0, 8))

        btn_selecionar = ctk.CTkButton(
            frame_pasta,
            text='Selecionar',
            font=ctk.CTkFont(size=12),
            width=100,
            height=35,
            command=self._selecionar_pasta
        )
        btn_selecionar.pack(side='right')

        # ─── Contagem de Fotos ───────────────────────────────────────────────
        frame_info = ctk.CTkFrame(self.frame_podas, fg_color='transparent')
        frame_info.pack(fill='x', padx=30, pady=(0, 10))

        label_encontradas = ctk.CTkLabel(
            frame_info,
            text='Fotos Encontradas:',
            font=ctk.CTkFont(size=14)
        )
        label_encontradas.pack(side='left')

        self.label_contagem = ctk.CTkLabel(
            frame_info,
            text='0',
            font=ctk.CTkFont(size=14, weight='bold')
        )
        self.label_contagem.pack(side='left', padx=(5, 0))

        # ─── Linha Separadora ────────────────────────────────────────────────
        separador = ctk.CTkFrame(self.frame_podas, height=2, fg_color=('gray70', 'gray30'))
        separador.pack(fill='x', padx=30, pady=(5, 15))

        # ─── Barra de Progresso ──────────────────────────────────────────────
        self.progress_bar = ctk.CTkProgressBar(self.frame_podas, height=8)
        self.progress_bar.pack(fill='x', padx=30, pady=(0, 5))
        self.progress_bar.set(0)

        # ─── Status da Foto Atual ────────────────────────────────────────────
        self.label_progresso = ctk.CTkLabel(
            self.frame_podas,
            text='0 / 0 fotos processadas',
            font=ctk.CTkFont(size=12),
            text_color=('gray40', 'gray60')
        )
        self.label_progresso.pack(fill='x', padx=30)

        self.label_foto_atual = ctk.CTkLabel(
            self.frame_podas,
            text='',
            font=ctk.CTkFont(size=11),
            text_color=('gray50', 'gray50')
        )
        self.label_foto_atual.pack(fill='x', padx=30, pady=(0, 10))

        # ─── Botão GERAR RELATÓRIO ───────────────────────────────────────────
        self.btn_gerar = ctk.CTkButton(
            self.frame_podas,
            text='GERAR RELATÓRIO',
            font=ctk.CTkFont(size=16, weight='bold'),
            height=45,
            command=self._gerar_relatorio
        )
        self.btn_gerar.pack(fill='x', padx=30, pady=(0, 5))

        # ─── Status Final ────────────────────────────────────────────────────
        self.label_status = ctk.CTkLabel(
            self.frame_podas,
            text='Pronto. Aguardando...',
            font=ctk.CTkFont(size=12),
            text_color=('gray40', 'gray60')
        )
        self.label_status.pack(fill='x', padx=30, pady=(0, 10))

    # ─── Interface de Obras ──────────────────────────────────────────────────
    def _construir_interface_obras(self):
        self.frame_obras = ctk.CTkFrame(self.frame_conteudo, fg_color='transparent')

        # ─── NOTA / PROJETO ──────────────────────────────────────────────────
        label_id_obra = ctk.CTkLabel(
            self.frame_obras,
            text='NOTA / PROJETO:',
            font=ctk.CTkFont(size=14),
            anchor='w'
        )
        label_id_obra.pack(fill='x', padx=30, pady=(15, 0))

        self.entry_id_obra = ctk.CTkEntry(
            self.frame_obras,
            placeholder_text='Ex: 440182639',
            font=ctk.CTkFont(size=14),
            height=35
        )
        self.entry_id_obra.pack(fill='x', padx=30, pady=(3, 10))

        # ─── Pasta Principal da Obra ─────────────────────────────────────────
        label_pasta_obra = ctk.CTkLabel(
            self.frame_obras,
            text='PASTA PRINCIPAL DA OBRA:',
            font=ctk.CTkFont(size=14),
            anchor='w'
        )
        label_pasta_obra.pack(fill='x', padx=30)

        frame_pasta_obra = ctk.CTkFrame(self.frame_obras, fg_color='transparent')
        frame_pasta_obra.pack(fill='x', padx=30, pady=(3, 10))

        self.entry_pasta_obra = ctk.CTkEntry(
            frame_pasta_obra,
            placeholder_text='Clique em Selecionar para escolher a pasta...',
            font=ctk.CTkFont(size=12),
            height=35,
            state='disabled'
        )
        self.entry_pasta_obra.pack(side='left', fill='x', expand=True, padx=(0, 8))

        btn_selecionar_obra = ctk.CTkButton(
            frame_pasta_obra,
            text='Selecionar',
            font=ctk.CTkFont(size=12),
            width=100,
            height=35,
            command=self._selecionar_pasta_obra
        )
        btn_selecionar_obra.pack(side='right')

        # ─── Linha Separadora ────────────────────────────────────────────────
        separador_obra = ctk.CTkFrame(self.frame_obras, height=2, fg_color=('gray70', 'gray30'))
        separador_obra.pack(fill='x', padx=30, pady=(5, 15))

        # ─── Botão ANALISAR OBRA ─────────────────────────────────────────────
        self.btn_analisar_obra = ctk.CTkButton(
            self.frame_obras,
            text='ANALISAR OBRA',
            font=ctk.CTkFont(size=16, weight='bold'),
            height=45,
            command=self._analisar_obra
        )
        self.btn_analisar_obra.pack(fill='x', padx=30, pady=(0, 5))

        # ─── Botão GERAR RELATÓRIO ───────────────────────────────────────────
        self.btn_gerar_obra = ctk.CTkButton(
            self.frame_obras,
            text='GERAR RELATÓRIO',
            font=ctk.CTkFont(size=16, weight='bold'),
            height=45,
            state='disabled',
            command=self._gerar_relatorio_obra
        )
        self.btn_gerar_obra.pack(fill='x', padx=30, pady=(0, 5))

        # ─── Barra de Progresso ──────────────────────────────────────────────
        self.progress_bar_obra = ctk.CTkProgressBar(self.frame_obras, height=8)
        self.progress_bar_obra.pack(fill='x', padx=30, pady=(0, 5))
        self.progress_bar_obra.set(0)

        # ─── Status da Foto Atual ────────────────────────────────────────────
        self.label_progresso_obra = ctk.CTkLabel(
            self.frame_obras,
            text='0 / 0 fotos processadas',
            font=ctk.CTkFont(size=12),
            text_color=('gray40', 'gray60')
        )
        self.label_progresso_obra.pack(fill='x', padx=30)

        self.label_foto_atual_obra = ctk.CTkLabel(
            self.frame_obras,
            text='',
            font=ctk.CTkFont(size=11),
            text_color=('gray50', 'gray50')
        )
        self.label_foto_atual_obra.pack(fill='x', padx=30, pady=(0, 10))

        # ─── Status ──────────────────────────────────────────────────────────
        self.label_status_obra = ctk.CTkLabel(
            self.frame_obras,
            text='Preencha os campos acima e clique em ANALISAR OBRA.',
            font=ctk.CTkFont(size=12),
            text_color=('gray40', 'gray60')
        )
        self.label_status_obra.pack(fill='x', padx=30, pady=(0, 10))

        # ─── Prévia dos Postes ───────────────────────────────────────────────
        self.frame_previa = ctk.CTkScrollableFrame(
            self.frame_obras,
            corner_radius=8,
            height=200
        )
        self.frame_previa.pack(fill='both', expand=True, padx=30, pady=(0, 10))

        self.label_previa = ctk.CTkLabel(
            self.frame_previa,
            text='',
            font=ctk.CTkFont(size=12),
            text_color=('gray50', 'gray50')
        )
        self.label_previa.pack(pady=20)

    # ─── Navegação entre Telas ────────────────────────────────────────────────
    def _mostrar_menu(self):
        self.frame_toolbar.pack_forget()
        self.frame_podas.pack_forget()
        self.frame_obras.pack_forget()
        self.frame_menu.pack(fill='both', expand=True)
        self.modo_atual = 'menu'

    def _mostrar_podas(self):
        self.frame_menu.pack_forget()
        self.frame_obras.pack_forget()
        self.frame_toolbar.pack(fill='x')
        self.frame_podas.pack(fill='both', expand=True)
        self.modo_atual = 'podas'
        logger.info('Módulo de Podas selecionado.')

    def _mostrar_obras(self):
        self.frame_menu.pack_forget()
        self.frame_podas.pack_forget()
        self.frame_toolbar.pack(fill='x')
        self.frame_obras.pack(fill='both', expand=True)
        self.modo_atual = 'obras'
        logger.info('Módulo de Obras selecionado.')

    def _voltar_menu(self):
        self._mostrar_menu()
        logger.info('Voltou ao menu principal.')

    # ─── Seleção de Pasta de Obra ─────────────────────────────────────────────
    def _selecionar_pasta_obra(self):
        pasta = filedialog.askdirectory(
            title='Selecione a pasta principal da obra'
        )
        if not pasta:
            return

        self.pasta_obra_selecionada = pasta
        self.entry_pasta_obra.configure(state='normal')
        self.entry_pasta_obra.delete(0, 'end')
        self.entry_pasta_obra.insert(0, pasta)
        self.entry_pasta_obra.configure(state='disabled')

        projeto_id = self.entry_id_obra.get().strip()
        if not projeto_id:
            projeto_id = Path(pasta).name
            self.entry_id_obra.delete(0, 'end')
            self.entry_id_obra.insert(0, projeto_id)

        self.label_status_obra.configure(
            text='Pasta selecionada. Clique em ANALISAR OBRA.',
            text_color=('gray40', 'gray60')
        )

    # ─── Análise da Obra ─────────────────────────────────────────────────────
    def _analisar_obra(self):
        projeto_id = self.entry_id_obra.get().strip()
        if not projeto_id:
            messagebox.showwarning('Aviso', 'Por favor, informe a Nota / Projeto.')
            self.entry_id_obra.focus()
            return

        if not self.pasta_obra_selecionada:
            messagebox.showwarning('Aviso', 'Por favor, selecione a pasta da obra.')
            return

        try:
            self.obra_handler = ObraHandler(self.pasta_obra_selecionada, projeto_id)
            logger.info(f'Projeto: {self.obra_handler.projeto_id}')

            self.label_status_obra.configure(
                text=f'Projeto: {self.obra_handler.projeto_id}  |  '
                     f'Postes: {self.obra_handler.total_postes}  |  '
                     f'Fotos: {self.obra_handler.total_fotos}',
                text_color=('#2E7D32', '#4CAF50')
            )

            self._exibir_previa_obra()
            self.btn_gerar_obra.configure(state='normal')

        except (FileNotFoundError, ValueError) as e:
            self.obra_handler = None
            self.btn_gerar_obra.configure(state='disabled')
            self.label_status_obra.configure(
                text=f'Erro: {str(e)}',
                text_color='#C62828'
            )
            logger.error(f'Erro ao carregar obra: {e}')
            self._limpar_previa_obra()

    def _limpar_previa_obra(self):
        for widget in self.frame_previa.winfo_children():
            widget.destroy()

    def _exibir_previa_obra(self):
        for widget in self.frame_previa.winfo_children():
            widget.destroy()

        handler = self.obra_handler

        # Card de resumo
        card_resumo = ctk.CTkFrame(self.frame_previa, fg_color=('gray92', 'gray25'), corner_radius=4)
        card_resumo.pack(fill='x', padx=10, pady=(8, 6))

        lbl_projeto = ctk.CTkLabel(
            card_resumo,
            text=f'Projeto: {handler.projeto_id}',
            font=ctk.CTkFont(size=13, weight='bold'),
            anchor='w'
        )
        lbl_projeto.pack(fill='x', padx=12, pady=(8, 1))

        lbl_resumo = ctk.CTkLabel(
            card_resumo,
            text=f'Total de Postes: {handler.total_postes}    |    Total de Fotos: {handler.total_fotos}',
            font=ctk.CTkFont(size=11),
            text_color=('gray40', 'gray60'),
            anchor='w'
        )
        lbl_resumo.pack(fill='x', padx=12, pady=(0, 8))

        # Separação visual
        sep = ctk.CTkFrame(self.frame_previa, height=1, fg_color=('gray75', 'gray40'))
        sep.pack(fill='x', padx=12, pady=(2, 6))

        # Lista de postes
        for poste in handler.postes:
            poste_card = ctk.CTkFrame(self.frame_previa, fg_color=('gray97', 'gray18'), corner_radius=3)
            poste_card.pack(fill='x', padx=12, pady=(1, 1))

            linha_poste = ctk.CTkFrame(poste_card, fg_color='transparent')
            linha_poste.pack(fill='x', padx=8, pady=4)

            lbl_nome = ctk.CTkLabel(
                linha_poste,
                text=poste.nome_formatado,
                font=ctk.CTkFont(size=12, weight='bold'),
                anchor='w'
            )
            lbl_nome.pack(side='left')

            lbl_fotos = ctk.CTkLabel(
                linha_poste,
                text=f'{poste.quantidade_fotos} fotos',
                font=ctk.CTkFont(size=11),
                text_color=('gray45', 'gray55'),
                anchor='e'
            )
            lbl_fotos.pack(side='right')

    # ─── Geração de Relatório de Obras ─────────────────────────────────────────
    def _gerar_relatorio_obra(self):
        """Executa o fluxo de geração do relatório de obra."""
        if self.processando:
            return

        projeto_id = self.entry_id_obra.get().strip()
        if not projeto_id:
            messagebox.showwarning('Aviso', 'Por favor, informe a Nota / Projeto.')
            self.entry_id_obra.focus()
            return

        if self.obra_handler is None:
            messagebox.showwarning('Aviso', 'Por favor, analise a obra primeiro.')
            return

        if self.obra_handler.total_fotos == 0:
            messagebox.showwarning('Aviso', 'Nenhuma foto encontrada na obra.')
            return

        if not messagebox.askyesno(
            'Confirmar',
            f'Projeto: {projeto_id}\n'
            f'Postes: {self.obra_handler.total_postes}\n'
            f'Fotos: {self.obra_handler.total_fotos}\n\n'
            'Deseja gerar o relatório?'
        ):
            return

        self.processando = True
        self.btn_gerar_obra.configure(state='disabled', text='PROCESSANDO...')
        self.progress_bar_obra.set(0)
        self.label_foto_atual_obra.configure(text='')
        self.label_status_obra.configure(text='Iniciando...', text_color='#1565C0')
        self.update_idletasks()

        try:
            self._executar_geracao_obra(projeto_id)
        except Exception as e:
            logger.exception('Erro durante a geração do relatório de obra')
            messagebox.showerror('Erro', f'Ocorreu um erro ao gerar o relatório:\n\n{str(e)}')
            self.label_status_obra.configure(
                text=f'Erro: {str(e)}',
                text_color='#C62828'
            )
        finally:
            self.processando = False
            self.btn_gerar_obra.configure(state='normal', text='GERAR RELATÓRIO')

    def _executar_geracao_obra(self, projeto_id: str):
        """Executa o pipeline de geração do relatório de obra."""
        logger.info(f'Iniciando geração do relatório de obra. Projeto: {projeto_id}')
        logger.info(f'Postes: {self.obra_handler.total_postes}')
        logger.info(f'Fotos totais: {self.obra_handler.total_fotos}')

        templates = list(self.pasta_modelo.glob('*Obra*.docx'))
        if not templates:
            raise FileNotFoundError(
                f'Template de Obras não encontrado em: {self.pasta_modelo}\n'
                'Certifique-se de que o arquivo "Relatório de Conclusão de Obra" '
                'esteja na pasta modelo/.'
            )
        template_path = templates[0]
        logger.info(f'Template encontrado: {template_path.name}')

        nome_saida = f'{projeto_id}_Relatorio_de_Conclusao_de_Obra_FK_Eng_PHB.docx'
        caminho_saida = self.pasta_relatorios / nome_saida

        fotos = []
        for poste in self.obra_handler.postes:
            for caminho_foto in poste.fotos:
                fotos.append(PhotoInfo(caminho_foto))

        logger.info('Ordem das fotos:')
        for i, foto in enumerate(fotos):
            logger.info(f'  {i + 1}. {foto.nome}')

        self.label_status_obra.configure(text='Abrindo template...', text_color='#1565C0')
        self.update_idletasks()

        handler = TemplateHandler(str(template_path))
        handler.substituir_id_projeto(projeto_id)
        handler.preencher_evidencias(fotos, callback_status=self._atualizar_progresso_obra)

        self.label_status_obra.configure(text='Salvando relatório...', text_color='#1565C0')
        self.update_idletasks()

        caminho_final = handler.salvar(str(caminho_saida))

        self.progress_bar_obra.set(1)
        self.label_progresso_obra.configure(text=f'{len(fotos)} / {len(fotos)} fotos processadas')
        self.label_foto_atual_obra.configure(text='')
        self.label_status_obra.configure(
            text='✓ Relatório gerado com sucesso!',
            text_color='#2E7D32'
        )

        logger.info(f'Relatório de Obra gerado com sucesso: {caminho_final}')

        if messagebox.askyesno(
            'Concluído',
            f'Relatório gerado com sucesso!\n\n'
            f'Arquivo: {nome_saida}\n'
            f'Local: {self.pasta_relatorios}\n\n'
            'Deseja abrir a pasta de relatórios?'
        ):
            os.startfile(str(self.pasta_relatorios))

    def _atualizar_progresso_obra(self, atual: int, total: int, nome_foto: str):
        """Callback para atualizar a barra de progresso durante a geração da obra."""
        self.after(0, self._renderizar_progresso_obra, atual, total, nome_foto)

    def _renderizar_progresso_obra(self, atual: int, total: int, nome_foto: str):
        """Atualiza os elementos visuais de progresso da obra."""
        progresso = atual / total if total > 0 else 0
        self.progress_bar_obra.set(progresso)
        self.label_progresso_obra.configure(text=f'{atual} / {total} fotos processadas')
        self.label_foto_atual_obra.configure(text=f'Foto atual: {nome_foto}')
        self.label_status_obra.configure(
            text=f'Processando foto {atual} de {total}...',
            text_color='#1565C0'
        )

    def _construir_cabecalho(self):
        """Constrói o cabeçalho com logos e título central."""
        header_frame = ctk.CTkFrame(self, height=85, corner_radius=10)
        header_frame.pack(fill='x', padx=20, pady=(15, 5))
        header_frame.pack_propagate(False)

        # Logo FK Engenharia (esquerda)
        try:
            fk_path = self._resource_path('assets/fk_logo.png')
            self._fk_img = ctk.CTkImage(
                light_image=Image.open(fk_path),
                dark_image=Image.open(fk_path),
                size=(110, 45)
            )
            fk_label = ctk.CTkLabel(header_frame, image=self._fk_img, text='')
            fk_label.pack(side='left', padx=(15, 5))
        except Exception as e:
            logger.warning(f'Logo FK não carregada: {e}')

        # Texto central
        center_frame = ctk.CTkFrame(header_frame, fg_color='transparent')
        center_frame.pack(side='left', fill='both', expand=True)

        titulo = ctk.CTkLabel(
            center_frame,
            text='Sistema Automático de Relatórios',
            font=ctk.CTkFont(size=18, weight='bold')
        )
        titulo.pack(expand=True, anchor='center')

        subtitulo = ctk.CTkLabel(
            center_frame,
            text='FK Engenharia e Serviços LTDA',
            font=ctk.CTkFont(size=11)
        )
        subtitulo.pack(anchor='center')

        # Logo Equatorial (direita)
        try:
            eq_path = self._resource_path('assets/equatorial_logo.png')
            self._eq_img = ctk.CTkImage(
                light_image=Image.open(eq_path),
                dark_image=Image.open(eq_path),
                size=(110, 45)
            )
            eq_label = ctk.CTkLabel(header_frame, image=self._eq_img, text='')
            eq_label.pack(side='right', padx=(5, 15))
        except Exception as e:
            logger.warning(f'Logo Equatorial não carregada: {e}')

    def _construir_log(self):
        """Constrói a área de log com scrollbar e botões."""
        log_frame = ctk.CTkFrame(self, corner_radius=10)
        log_frame.pack(fill='x', padx=20, pady=(0, 5))

        log_header = ctk.CTkFrame(log_frame, fg_color='transparent')
        log_header.pack(fill='x', padx=15, pady=(8, 0))

        lbl_log = ctk.CTkLabel(
            log_header,
            text='📋 Log do Sistema',
            font=ctk.CTkFont(size=13, weight='bold')
        )
        lbl_log.pack(side='left')

        self.btn_toggle_tema = ctk.CTkButton(
            log_header,
            text='🌙 Modo Escuro' if self.tema_atual == 'light' else '☀️ Modo Claro',
            font=ctk.CTkFont(size=11),
            width=130,
            height=28,
            fg_color=('gray60', 'gray30'),
            hover_color=('gray40', 'gray20'),
            command=self._toggle_tema
        )
        self.btn_toggle_tema.pack(side='right', padx=(5, 0))

        self.btn_limpar_log = ctk.CTkButton(
            log_header,
            text='🧹 Limpar Log',
            font=ctk.CTkFont(size=11),
            width=110,
            height=28,
            fg_color=('gray60', 'gray30'),
            hover_color=('gray40', 'gray20'),
            command=self._limpar_log
        )
        self.btn_limpar_log.pack(side='right')

        self.log_text = ctk.CTkTextbox(
            log_frame,
            height=130,
            font=ctk.CTkFont(size=11, family='Consolas'),
            state='disabled',
            wrap='word'
        )
        self.log_text.pack(fill='x', padx=15, pady=(5, 12))

    def _construir_rodape(self):
        """Constrói o rodapé com versão e créditos."""
        footer_frame = ctk.CTkFrame(self, fg_color='transparent')
        footer_frame.pack(fill='x', padx=20, pady=(0, 15))

        versao = ctk.CTkLabel(
            footer_frame,
            text='Versão 1.0',
            font=ctk.CTkFont(size=10),
            text_color=('gray50', 'gray50')
        )
        versao.pack(side='left')

        creditos = ctk.CTkLabel(
            footer_frame,
            text='Desenvolvido por Diego Gomes',
            font=ctk.CTkFont(size=10),
            text_color=('gray50', 'gray50')
        )
        creditos.pack(side='right')

    def _toggle_tema(self):
        """Alterna entre modo claro e escuro instantaneamente."""
        if self.tema_atual == 'light':
            ctk.set_appearance_mode('dark')
            self.tema_atual = 'dark'
            self.btn_toggle_tema.configure(text='☀️ Modo Claro')
        else:
            ctk.set_appearance_mode('light')
            self.tema_atual = 'light'
            self.btn_toggle_tema.configure(text='🌙 Modo Escuro')
        self._salvar_config_tema(self.tema_atual)
        logger.info(f'Tema alterado para: {self.tema_atual}')

    def _limpar_log(self):
        """Limpa apenas o conteúdo da área de log."""
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')
        logger.info('Log limpo pelo usuário')

    def _selecionar_pasta(self):
        """Abre o diálogo para selecionar a pasta de fotos."""
        pasta = filedialog.askdirectory(title='Selecione a pasta com as fotos das podas')
        if not pasta:
            return

        self.pasta_fotos_selecionada = pasta
        # Mostra apenas o nome da pasta no campo
        nome_pasta = Path(pasta).name
        self.entry_pasta.configure(state='normal')
        self.entry_pasta.delete(0, 'end')
        self.entry_pasta.insert(0, nome_pasta)
        self.entry_pasta.configure(state='disabled')

        # Tenta carregar as fotos para contagem
        try:
            self.photo_handler = PhotoHandler(pasta)
            self.label_contagem.configure(text=str(self.photo_handler.quantidade))
            self.label_status.configure(
                text=f'{self.photo_handler.quantidade} fotos encontradas. Pronto para gerar.',
                text_color='#2E7D32'
            )
            logger.info(f'Pasta selecionada: {pasta} ({self.photo_handler.quantidade} fotos)')
        except (FileNotFoundError, ValueError) as e:
            self.photo_handler = None
            self.label_contagem.configure(text='0')
            self.label_status.configure(
                text=f'Erro: {str(e)}',
                text_color='#C62828'
            )
            logger.error(f'Erro ao carregar fotos: {e}')

    def _atualizar_progresso(self, atual: int, total: int, nome_foto: str):
        """
        Callback para atualizar a barra de progresso durante o processamento.
        Executado na thread principal via after().
        """
        self.after(0, self._renderizar_progresso, atual, total, nome_foto)

    def _renderizar_progresso(self, atual: int, total: int, nome_foto: str):
        """Atualiza os elementos visuais de progresso."""
        progresso = atual / total if total > 0 else 0
        self.progress_bar.set(progresso)
        self.label_progresso.configure(text=f'{atual} / {total} fotos processadas')
        self.label_foto_atual.configure(text=f'Foto atual: {nome_foto}')
        self.label_status.configure(
            text=f'Processando foto {atual} de {total}...',
            text_color='#1565C0'
        )

    def _gerar_relatorio(self):
        """Executa o fluxo de geração do relatório."""
        if self.processando:
            return

        # ─── Validações ──────────────────────────────────────────────────────
        projeto_id = self.entry_id.get().strip()
        if not projeto_id:
            messagebox.showwarning('Aviso', 'Por favor, informe o ID do projeto.')
            self.entry_id.focus()
            return

        if not self.pasta_fotos_selecionada or self.photo_handler is None:
            messagebox.showwarning('Aviso', 'Por favor, selecione a pasta com as fotos.')
            return

        if self.photo_handler.quantidade == 0:
            messagebox.showwarning('Aviso', 'Nenhuma foto encontrada na pasta selecionada.')
            return

        # Confirmação
        if not messagebox.askyesno(
            'Confirmar',
            f'Projeto: {projeto_id}\n'
            f'Fotos: {self.photo_handler.quantidade}\n\n'
            'Deseja gerar o relatório?'
        ):
            return

        # ─── Execução ────────────────────────────────────────────────────────
        self.processando = True
        self.btn_gerar.configure(state='disabled', text='PROCESSANDO...')
        self.progress_bar.set(0)
        self.label_foto_atual.configure(text='')
        self.label_status.configure(text='Iniciando...', text_color='#1565C0')
        self.update_idletasks()

        try:
            self._executar_geracao(projeto_id)
        except Exception as e:
            logger.exception('Erro durante a geração do relatório')
            messagebox.showerror('Erro', f'Ocorreu um erro ao gerar o relatório:\n\n{str(e)}')
            self.label_status.configure(
                text=f'Erro: {str(e)}',
                text_color='#C62828'
            )
        finally:
            self.processando = False
            self.btn_gerar.configure(state='normal', text='GERAR RELATÓRIO')

    def _executar_geracao(self, projeto_id: str):
        """Executa o pipeline de geração do relatório."""
        # ─── Localiza o template ─────────────────────────────────────────────
        templates = list(self.pasta_modelo.glob('*_Relatorio_de_Podas_FK_Eng_PHB.docx'))
        if not templates:
            raise FileNotFoundError(
                f'Template não encontrado em: {self.pasta_modelo}\n'
                'Certifique-se de que o arquivo *_Relatorio_de_Podas_FK_Eng_PHB.docx '
                'esteja na pasta modelo/.'
            )
        template_path = templates[0]

        logger.info(f'Template carregado: {template_path.name}')

        # ─── Nome do arquivo de saída ────────────────────────────────────────
        nome_saida = f'{projeto_id}_Relatorio_de_Podas_FK_Eng_PHB.docx'
        caminho_saida = self.pasta_relatorios / nome_saida

        # ─── Processa as fotos (ordenadas cronologicamente) ──────────────────
        logger.info(f'Processando {self.photo_handler.quantidade} fotos...')
        fotos = self.photo_handler.listar_fotos()

        # Mostra a ordem cronológica no log
        logger.info('Ordem cronológica das fotos:')
        for i, foto in enumerate(fotos):
            logger.info(f'  {i + 1}. {foto.nome} - {foto.data_hora}')

        # ─── Preenche o template ─────────────────────────────────────────────
        self.label_status.configure(text='Abrindo template...', text_color='#1565C0')
        self.update_idletasks()

        handler = TemplateHandler(str(template_path))

        # Substitui o ID do projeto
        handler.substituir_id_projeto(projeto_id)

        # Preenche as evidências
        handler.preencher_evidencias(fotos, callback_status=self._atualizar_progresso)

        # ─── Salva o relatório ───────────────────────────────────────────────
        self.label_status.configure(text='Salvando relatório...', text_color='#1565C0')
        self.update_idletasks()

        caminho_final = handler.salvar(str(caminho_saida))

        # ─── Concluído ───────────────────────────────────────────────────────
        self.progress_bar.set(1)
        self.label_progresso.configure(text=f'{len(fotos)} / {len(fotos)} fotos processadas')
        self.label_foto_atual.configure(text='')
        self.label_status.configure(
            text=f'✓ Relatório gerado com sucesso!',
            text_color='#2E7D32'
        )

        logger.info(f'Relatório gerado com sucesso: {caminho_final}')

        # Pergunta se deseja abrir a pasta
        if messagebox.askyesno(
            'Concluído',
            f'Relatório gerado com sucesso!\n\n'
            f'Arquivo: {nome_saida}\n'
            f'Local: {self.pasta_relatorios}\n\n'
            'Deseja abrir a pasta de relatórios?'
        ):
            os.startfile(str(self.pasta_relatorios))


# ─── Ponto de Entrada ───────────────────────────────────────────────────────
if __name__ == '__main__':
    app = SistemaPodaApp()
    app.mainloop()
