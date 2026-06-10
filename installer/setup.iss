; Script Inno Setup para o Sistema Automático de Relatórios de Poda
; FK Engenharia e Serviços LTDA
; Versão 1.0

#define MyAppName "Sistema Automático de Relatórios de Poda"
#define MyAppVersion "1.0"
#define MyAppPublisher "FK Engenharia e Serviços LTDA"
#define MyAppURL ""
#define MyAppExeName "sistema.exe"

[Setup]
; Informações do aplicativo
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

; Diretório padrão de instalação
DefaultDirName=C:\Program Files\Sistema Relatórios de Poda
DisableProgramGroupPage=yes

; Arquivo de saída do instalador
OutputDir=..
OutputBaseFilename=Setup_Sistema_Podas_v{#MyAppVersion}

; Compactação
Compression=lzma
SolidCompression=yes

; Permissões
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

; Ícone do instalador (opcional)
; SetupIconFile=..\assets\icone.ico

[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na &Área de Trabalho"; GroupDescription: "Ícones adicionais:"
Name: "startmenuicon"; Description: "Criar atalho no &Menu Iniciar"; GroupDescription: "Ícones adicionais:"

[Files]
; Executável principal e arquivos do _internal
Source: "..\dist\sistema\sistema.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\dist\sistema\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

; Assets (logos)
Source: "..\dist\sistema\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs

; Modelo (template Word)
Source: "..\dist\sistema\modelo\*"; DestDir: "{app}\modelo"; Flags: ignoreversion recursesubdirs createallsubdirs

; Configuração
Source: "..\dist\sistema\config.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Atalho no Menu Iniciar
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: startmenuicon

; Atalho na Área de Trabalho
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Opção para abrir o sistema após a instalação
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: postinstall nowait skipifsilent shellexec
