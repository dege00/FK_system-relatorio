import os
import sys
import json
import logging
from pathlib import Path
from tkinter import filedialog, messagebox
from datetime import datetime
from PIL import Image

import customtkinter as ctk

from modules.photo_handler import PhotoHandler
from modules.template_handler import TemplateHandler

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
        self.title('Sistema de Relatórios de Poda - Equatorial')
        self.geometry('730x780')
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

        # ─── Variáveis de Estado ─────────────────────────────────────────────
        self.pasta_fotos_selecionada = None
        self.photo_handler = None
        self.processando = False
        self.tema_atual = tema_inicial

        # ─── Construção da Interface ─────────────────────────────────────────
        self._construir_interface()

        # ─── Adiciona handler de log no Text Widget ──────────────────────────
        self._log_handler = TextHandler(self.log_text)
        logging.getLogger().addHandler(self._log_handler)

        self.log_text.configure(state='normal')
        self.log_text.insert('end', 'Sistema Automático de Relatórios de Poda\n')
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
        """Constrói todos os elementos da interface gráfica."""
        # ─── Cabeçalho com Logos ─────────────────────────────────────────────
        self._construir_cabecalho()

        # ─── Frame Principal ─────────────────────────────────────────────────
        frame_principal = ctk.CTkFrame(self, corner_radius=10)
        frame_principal.pack(fill='both', expand=True, padx=20, pady=(5, 5))

        # ─── ID do Projeto ───────────────────────────────────────────────────
        label_id = ctk.CTkLabel(
            frame_principal,
            text='Nota do Projeto:',
            font=ctk.CTkFont(size=14),
            anchor='w'
        )
        label_id.pack(fill='x', padx=30, pady=(15, 0))

        self.entry_id = ctk.CTkEntry(
            frame_principal,
            placeholder_text='Ex: 440100001',
            font=ctk.CTkFont(size=14),
            height=35
        )
        self.entry_id.pack(fill='x', padx=30, pady=(3, 10))

        # ─── Pasta das Fotos ─────────────────────────────────────────────────
        label_pasta = ctk.CTkLabel(
            frame_principal,
            text='Pasta das Fotos:',
            font=ctk.CTkFont(size=14),
            anchor='w'
        )
        label_pasta.pack(fill='x', padx=30)

        frame_pasta = ctk.CTkFrame(frame_principal, fg_color='transparent')
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
        frame_info = ctk.CTkFrame(frame_principal, fg_color='transparent')
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
        separador = ctk.CTkFrame(frame_principal, height=2, fg_color=('gray70', 'gray30'))
        separador.pack(fill='x', padx=30, pady=(5, 15))

        # ─── Barra de Progresso ──────────────────────────────────────────────
        self.progress_bar = ctk.CTkProgressBar(frame_principal, height=8)
        self.progress_bar.pack(fill='x', padx=30, pady=(0, 5))
        self.progress_bar.set(0)

        # ─── Status da Foto Atual ────────────────────────────────────────────
        self.label_progresso = ctk.CTkLabel(
            frame_principal,
            text='0 / 0 fotos processadas',
            font=ctk.CTkFont(size=12),
            text_color=('gray40', 'gray60')
        )
        self.label_progresso.pack(fill='x', padx=30)

        self.label_foto_atual = ctk.CTkLabel(
            frame_principal,
            text='',
            font=ctk.CTkFont(size=11),
            text_color=('gray50', 'gray50')
        )
        self.label_foto_atual.pack(fill='x', padx=30, pady=(0, 10))

        # ─── Botão GERAR RELATÓRIO ───────────────────────────────────────────
        self.btn_gerar = ctk.CTkButton(
            frame_principal,
            text='GERAR RELATÓRIO',
            font=ctk.CTkFont(size=16, weight='bold'),
            height=45,
            command=self._gerar_relatorio
        )
        self.btn_gerar.pack(fill='x', padx=30, pady=(0, 5))

        # ─── Status Final ────────────────────────────────────────────────────
        self.label_status = ctk.CTkLabel(
            frame_principal,
            text='Pronto. Aguardando...',
            font=ctk.CTkFont(size=12),
            text_color=('gray40', 'gray60')
        )
        self.label_status.pack(fill='x', padx=30, pady=(0, 10))

        # ─── Área de Log ─────────────────────────────────────────────────────
        self._construir_log()

        # ─── Rodapé ──────────────────────────────────────────────────────────
        self._construir_rodape()

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
            text='Sistema Automático de Relatórios de Poda',
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
