================================================================================
  SISTEMA AUTOMÁTICO DE RELATÓRIOS DE PODA - EQUATORIAL
  FK Engenharia
================================================================================

  Manual do Usuário


================================================================================
1. O QUE O SISTEMA FAZ
================================================================================

Este sistema gera automaticamente relatórios de poda em formato Word (.docx).

O que ele faz:

- Você informa o número do projeto
- Você seleciona a pasta onde estão as fotos das podas
- O sistema lê todas as fotos, organiza por data/hora
- Insere automaticamente as fotos no modelo Word existente
- Gera um arquivo Word completo e pronto para uso

O que ele NÃO faz:

- Não cria o layout do relatório (ele usa um modelo já existente)
- Não edita fotos
- Não precisa de Microsoft Word instalado para funcionar

O sistema usa um arquivo modelo (.docx) que já possui:
- Capa com informações do projeto
- Páginas de evidências com espaços reservados para fotos
- Página final de assinatura

O sistema apenas preenche automaticamente esse modelo com as informações
e as fotos que você selecionar.


================================================================================
2. REQUISITOS
================================================================================

Para executar o sistema, você precisa ter instalado:

A) Python (versão 3.9 ou superior)
   - Baixe em: https://www.python.org/downloads/
   - Marque a opção "Add Python to PATH" durante a instalação

B) Bibliotecas Python (dependências)
   - Instaladas automaticamente via comando pip (veja seção 3)

C) Microsoft Word (OPCIONAL)
   - Não é necessário para gerar os relatórios
   - Apenas para visualizar/editar o arquivo .docx gerado

D) Sistema Operacional
   - Windows (recomendado), Linux ou macOS


================================================================================
3. COMO INSTALAR
================================================================================

PRIMEIRO PASSO - Verificar se o Python está instalado:

   Abra o terminal (Prompt de Comando ou PowerShell) e digite:

   python --version

   Se aparecer algo como "Python 3.12.x" está ok.

   Se aparecer "Python não encontrado", instale o Python primeiro
   (https://www.python.org/downloads/) e repita este passo.

SEGUNDO PASSO - Instalar as bibliotecas necessárias:

   No terminal, dentro da pasta do projeto, execute:

   pip install python-docx Pillow customtkinter piexif

   Aguarde a instalação concluir (pode levar alguns minutos).

TERCEIRO PASSO - Verificar se tudo está instalado:

   python -c "import docx; from PIL import Image; import customtkinter; import piexif; print('OK')"

   Se aparecer "OK" está tudo pronto.

   Se aparecer algum erro, repita o passo anterior.


================================================================================
4. ESTRUTURA DAS PASTAS
================================================================================

A pasta do projeto deve estar organizada da seguinte forma:

   SYSTEM/
   │
   ├── sistema.py                 ← Arquivo principal (executar este)
   ├── README.txt                 ← Este manual
   │
   ├── modelo/                    ← Coloque aqui o arquivo modelo .docx
   │   └── XXXXXXX_Relatorio_de_Podas_FK_Eng_PHB.docx
   │
   ├── relatorios_gerados/        ← Os relatórios prontos aparecem aqui
   │
   ├── assets/                    ← (uso futuro)
   │
   └── modules/                   ← Arquivos internos do sistema
       ├── __init__.py
       ├── photo_handler.py
       └── template_handler.py

Onde colocar cada coisa:

   ARQUIVO MODELO (.docx)
   Deve estar dentro da pasta "modelo/".
   O nome do arquivo deve conter "Relatorio_de_Podas_FK_Eng_PHB" no nome.
   Exemplo: modelo/440182639_Relatorio_de_Podas_FK_Eng_PHB.docx

   PASTA DAS FOTOS
   Pode estar em qualquer lugar do computador.
   Apenas fotos nos formatos JPG, JPEG e PNG são aceitas.
   Exemplo: C:\Users\SeuNome\Desktop\Fotos_Poda\440182639\

   RELATÓRIOS GERADOS
   Aparecem automaticamente dentro da pasta "relatorios_gerados/".
   O nome do arquivo será o ID do projeto seguido do nome do modelo.


================================================================================
5. COMO EXECUTAR
================================================================================

MODO 1 - PELO TERMINAL (recomendado para primeiro uso):

   Passo 1: Abra o terminal (Prompt de Comando ou PowerShell)

   Passo 2: Navegue até a pasta do projeto:
      cd C:\caminho\para\SYSTEM

   Passo 3: Execute o comando:
      python sistema.py

   A janela do sistema vai abrir.

MODO 2 - CLICANDO NO ARQUIVO (após configurado):

   Se o Python estiver configurado corretamente, você pode apenas
   dar duplo clique no arquivo "sistema.py" para executar.


================================================================================
6. COMO UTILIZAR
================================================================================

PASSO 1 - DIGITAR O ID DO PROJETO

   No campo "ID do Projeto", digite o número do projeto.
   Exemplo: 440182639

PASSO 2 - SELECIONAR A PASTA DAS FOTOS

   Clique no botão "Selecionar" e escolha a pasta onde estão as fotos.
   O sistema vai contar automaticamente quantas fotos encontrou.
   As fotos serão organizadas da mais antiga para a mais recente.

PASSO 3 - GERAR RELATÓRIO

   Clique no botão "GERAR RELATÓRIO".
   Uma barra de progresso mostra o andamento.
   Quando terminar, uma mensagem de sucesso aparece.

PASSO 4 - ENCONTRAR O RESULTADO

   O relatório gerado estará na pasta "relatorios_gerados/".
   O nome do arquivo será algo como:
      440182639_Relatorio_de_Podas_FK_Eng_PHB.docx
   O sistema pergunta se você quer abrir a pasta automaticamente.

DICAS IMPORTANTES:

   - As fotos são inseridas em ordem cronológica (da mais antiga para a mais nova)
   - Cada página de evidências comporta 4 fotos (2 colunas x 2 linhas)
   - Se houver mais fotos que o espaço disponível, novas páginas são
     criadas automaticamente
   - O sistema funciona com qualquer quantidade de fotos
     (10, 50, 100, 300, 500... sem limite prático)


================================================================================
7. COMO GERAR EXECUTÁVEL (.exe)
================================================================================

Se quiser gerar um arquivo .exe para executar sem precisar do Python:

PASSO 1 - Instalar o PyInstaller:

   pip install pyinstaller

PASSO 2 - Gerar o executável:

   pyinstaller --onefile --windowed sistema.py

   Explicação dos parâmetros:
   - --onefile    : Gera um único arquivo .exe
   - --windowed   : Não abre o terminal junto com o programa

PASSO 3 - Onde encontrar:

   O executável será criado na pasta:
   SYSTEM\dist\sistema.exe

   Você pode copiar esse arquivo para qualquer lugar e executar.

ATENÇÃO:
   - O arquivo .exe é maior (cerca de 30-50 MB)
   - A primeira execução pode ser mais lenta
   - Antivírus podem alertar falsamente (é normal em programas
     gerados com PyInstaller)


================================================================================
8. SOLUÇÃO DE PROBLEMAS
================================================================================

PROBLEMA: "Python não é reconhecido como comando interno"
SOLUÇÃO: Instale o Python e marque a opção "Add Python to PATH"
         Ou execute: C:\caminho\python.exe sistema.py

PROBLEMA: "ModuleNotFoundError: No module named 'docx'"
SOLUÇÃO: Instale as bibliotecas:
         pip install python-docx Pillow customtkinter piexif

PROBLEMA: "Erro ao abrir o template DOCX"
SOLUÇÃO: Verifique se o arquivo modelo está dentro da pasta "modelo/"
         O nome do arquivo deve conter "Relatorio_de_Podas_FK_Eng_PHB"
         no nome do arquivo

PROBLEMA: "Pasta de fotos vazia" ou "Nenhuma foto encontrada"
SOLUÇÃO: Verifique se a pasta selecionada contém arquivos .jpg, .jpeg ou .png
         Verifique se as fotos não estão com extensão errada (ex: .JPG maiúsculo)

PROBLEMA: "Erro ao processar imagem corrompida"
SOLUÇÃO: Verifique se alguma foto da pasta está danificada ou incompleta
         Tente abrir a foto no visualizador de imagens do Windows

PROBLEMA: "A janela abre e fecha rapidamente"
SOLUÇÃO: Execute pelo terminal para ver a mensagem de erro:
         python sistema.py

PROBLEMA: "customtkinter não funciona"
SOLUÇÃO: No Windows, pode ser necessário instalar:
         pip install customtkinter --upgrade

PROBLEMA: "O relatório gerado está muito pesado"
SOLUÇÃO: Fotos muito grandes (acima de 5 MB cada) podem deixar o arquivo
         pesado. Reduza o tamanho das fotos antes de usar.

PROBLEMA: "Erro de permissão ao salvar"
SOLUÇÃO: Execute o programa como administrador ou salve em uma pasta
         onde você tem permissão de escrita

PROBLEMA: "pyinstaller gerou um arquivo muito grande"
SOLUÇÃO: É normal. O arquivo .exe contém todo o Python e as bibliotecas
         dentro dele. O tamanho varia de 30 a 80 MB.


================================================================================
9. ATUALIZAÇÕES FUTURAS
================================================================================

O sistema foi desenvolvido de forma modular para facilitar alterações.

ESTRUTURA DOS ARQUIVOS:

   modules/photo_handler.py
   - Responsável por ler e organizar as fotos
   - Para adicionar novos formatos de imagem, edite a variável
     EXTENSOES_VALIDAS no início do arquivo

   modules/template_handler.py
   - Responsável por manipular o documento Word
   - Para alterar como as fotos são inseridas ou redimensionadas,
     edite os métodos _inserir_imagem_celula ou preencher_evidencias

   sistema.py
   - Interface gráfica (CustomTkinter)
   - Para adicionar novos campos na tela, altere o método
     _construir_interface

COMO ADICIONAR NOVAS FUNCIONALIDADES:

   1. Adicione novos campos na interface em sistema.py
   2. Crie novos placeholders no modelo Word (ex: CIDADE: XXXXXXX)
   3. Adicione o código para substituir no template_handler.py

EXEMPLO - Adicionar campo "Responsável Técnico":

   No modelo Word:  Responsável: XXXXXXX
   No template_handler.py, adicione na substituição:
       if 'Responsável:' in texto_completo and 'XXXXXXX' in texto_completo:
           for run in paragrafo.runs:
               if 'XXXXXXX' in run.text:
                   run.text = run.text.replace('XXXXXXX', responsavel)


================================================================================
10. COMO USAR EM 30 SEGUNDOS
================================================================================

1. Abra o terminal na pasta do projeto
2. Execute: python sistema.py
3. Digite o ID do projeto
4. Clique em "Selecionar" e escolha a pasta das fotos
5. Clique em "GERAR RELATÓRIO"
6. Pronto! O arquivo está em: relatorios_gerados/

Pré-requisito único (uma vez):
   pip install python-docx Pillow customtkinter piexif

================================================================================
  FK Engenharia e Serviços LTDA
  Sistema gerado automaticamente - 2024
================================================================================
