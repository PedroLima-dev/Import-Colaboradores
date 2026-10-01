# 🚀 Automação de Importação de Colaboradores & Gestão AD - Grupo Umuarama

> **Documentação Técnica para a Equipe de TI / Analistas**
> 
> *Este projeto automatiza o fluxo de extração, tratamento de dados (ETL), provisionamento de novos colaboradores, readmissões e gestão de contas no Active Directory (AD) do Grupo Umuarama.*

---

## 📌 Visão Geral da Arquitetura

O sistema integra três pilares essenciais para a operação de TI:

1. **Processamento e Higienização de Dados (ETL):**
   - Lê a base bruta de exportação do sistema Senior (`FPRE111-*.CSV`).
   - Normaliza nomes, CPFs, e-mails, telefones, departamentos, unidades e cargos (padronizando nomenclaturas para o formato neutro/gênero, ex: `SECRETARIO(A)`).
   - Filtra os colaboradores com base na data de admissão e gera a planilha formatada no caminho `DD-MM/importar-usuarios-DDMMYYYY.csv`.

2. **Provisionamento no Active Directory (AD) & Readmissão por CPF:**
   - Lê o arquivo CSV formatado e processa em lote.
   - Conecta ao Domain Controller (DC) via comandos integrados em PowerShell / Directory Services (`DirectoryEntry`).
   - Mapeia automaticamente a **Organizational Unit (OU)** correta no AD com base na Unidade e Concessionária (`ou_map.json`).
   - **Tratamento Inteligente de Readmissão:**
     - Se o CPF já existir no AD em uma conta **desativada**, o sistema executa a **readmissão**: reativa a conta, move para a OU correta (caso tenha mudado de unidade), atualiza os atributos e define nova senha temporária.
     - Se o CPF já existir em uma conta **ativa** com login diferente, registra `CONFLITO CPF` no log para evitar duplicidades.
   - Criação de novas contas com grupos VPN automáticos quando aplicável e registro detalhado de logs em CSV.

3. **Consulta & Gestão de Contas no AD:**
   - Interface ágil para pesquisa de contas por **CPF**, **Login (`sAMAccountName`)** ou **Nome Completo**.
   - Visualização do card de status em tempo real (**ATIVO** / **DESATIVADO**) com a localização física/OU legível.
   - **Ações Rápidas na Conta:**
     - **🔑 Alterar Senha:** gerador automático de senha padrão (`@Primeironome2026`) com opção de forçar troca no próximo logon.
     - **🚫 Desativar Conta:** bloqueio imediato da conta no AD (desativação via UAC).
     - **🔄 Reativar / Readmissão:** reabilitação de contas desligadas, com reset de senha e possibilidade de mover o colaborador para uma nova OU.

---

## 🛠️ Pré-requisitos & Configuração do Ambiente

### 1. Requisitos do Sistema
- **Sistema Operacional:** Windows 10 / 11 ou Windows Server (necessário para execução dos comandos PowerShell / AD).
- **Python:** Versão `3.8` ou superior (certifique-se de marcar a opção **"Add Python to PATH"** durante a instalação).

### 2. Dependências Python
Instale as bibliotecas necessárias via terminal (PowerShell ou Prompt de Comando):
```bash
pip install pandas openpyxl chardet
```

### 3. Cadastro de Credenciais (Primeira Execução)
Para garantir a segurança e evitar a exposição de senhas em scripts, o sistema utiliza o **Windows Credential Manager** (`cmdkey`).

Execute este script **apenas uma vez** no seu computador:
```cmd
CADASTRAR_CREDENCIAIS.bat
```
Ele solicitará seu usuário (ex: `seu.nome@umuarama.local`) e senha da rede para liberar a conexão com o servidor de arquivos (`\\10.56.43.28`).

---

## 🖥️ Módulos da Interface Gráfica (GUI)

O aplicativo [ImportarUsuariosUmuarama.py](file:///c:/Users/Pedro%20Lima/Documents/IMPORT_COLABORADORES/ImportarUsuariosUmuarama.py) (ou executável `ImportarUsuariosUmuarama.exe`) disponibiliza 3 abas de operação:

### 1. 📂 Aba "GERAR CSV" (ETL)
- Seleção manual ou automática do arquivo `FPRE111-*.CSV`.
- Filtro por data de admissão (data mais recente detectada automaticamente ou data customizada).
- Botão **"Gerar CSV"** que dispara o engine de normalização e salva o arquivo tratado na pasta por data.

### 2. 👥 Aba "IMPORTAR AD" (Criação & Readmissão em Lote)
- Carregamento do CSV gerado.
- Pré-visualização com tabela interativa de usuários, cargos, e-mails e status de validação.
- Resolução dinâmica de Domain Controller (DC) com fallback automático.
- Processamento em lote com tratamento automático de readmissões por CPF e resolução de duplicidades.
- Exportação de relatório CSV com o status individual de cada colaborador importado.

### 3. 🔍 Aba "CONSULTA & GESTÃO AD" (Administração Rápida)
- **Busca Rápida:** Campo unificado com suporte a busca por CPF (com ou sem máscara), login (`sAMAccountName`) ou nome.
- **Card de Informações:** Exibe Nome Completo, Login, E-mail, CPF, Caminho legível da Unidade/OU e Badge de status (**ATIVO** / **DESATIVADO**).
- **Ações Administrativas Diretas:**
  - **Alterar Senha:** Modal com pré-preenchimento da senha padronizada `@Primeironome2026`, botão de regerar e flag de expiração no logon.
  - **Desativar:** Confirmação segura para desativar colaborador desligado imediatamente.
  - **Reativar / Readmissão:** Reabilitação da conta com redefinição de senha e campo opcional para mover o colaborador de OU (caso tenha mudado de concessionária ou holding).

---

## 💻 Modos de Uso (Passo a Passo)

### 🔴 Método 1: Execução Automática (Fluxo Diário Padrão)

Recomendado para o dia a dia do analista. O script cuida do mapeamento de rede, busca do arquivo mais recente e abertura da interface.

1. **Exportar a Base no Senior:**
   - Gere o relatório **FPRE111** no formato **CSV**.
   - Salve o arquivo no servidor de arquivos (`\\10.56.43.28\Users\pedro.lima\Desktop\IMPORT`) com o padrão de nome: `FPRE111-DDMMYYYY.csv`.

2. **Rodar a Automação:**
   - Dê um duplo clique em:
     ```cmd
     EXECUTAR_IMPORT_COMPLETO.bat
     ```
   - O processo automático irá:
     - Autenticar no servidor de arquivos.
     - Localizar a base `FPRE111` mais recente.
     - Processar e gerar o CSV de admissões formatado (ex: `16-03/importar-usuarios-16032026.csv`).
     - Abrir a **Interface Gráfica (GUI)** com os dados pré-carregados.

3. **Conferência & Importação:**
   - Na tela da interface gráfica, revise a lista de usuários a serem criados ou readmitidos.
   - Digite suas credenciais com acesso de Admin no Active Directory (AD).
   - Clique em **"Iniciar Importação AD"**.

---

### 🟡 Método 2: Execução Manual / Linha de Comando (CLI)

#### 1. Executar o Tratamento dos Dados (ETL)
```bash
python gerar_importar_usuarios_uap.py "\\10.56.43.28\caminho\FPRE111-16032026.CSV"
```
*Dica:* Para filtrar por uma data de admissão específica:
```bash
python gerar_importar_usuarios_uap.py "FPRE111-16032026.CSV" "16/03/2026"
```

#### 2. Abrir a Interface Gráfica com Pré-carregamento
```bash
python ImportarUsuariosUmuarama.py --csv-preload "16-03/importar-usuarios-16032026.csv"
```

#### 3. Modo Autônomo / Batch (Sem Interface)
```bash
python ImportarUsuariosUmuarama.py --modo-bat --csv "caminho/arquivo.csv" --usuario "admin.user" --senha "sua_senha"
```

---

## 🔍 Checklist de Validação ("Double-Check")

Antes de efetuar o provisionamento no AD, o Analista deve verificar:

- [ ] **Padronização de Cargos:** Confirmar se os cargos foram higienizados (ex: sem sufixos como `SR`, `JR`, `PLENO` desnecessários e com formatação correta).
- [ ] **Mapeamento de OUs:** Garantir que a Unidade e Marca correspondem a uma OU válida no mapa (`ou_map.json`).
- [ ] **E-mails e Superior:** Confirmar se o e-mail do colaborador e o e-mail do gestor direto estão preenchidos no padrão `nome.sobrenome@grupoumuarama.com.br`.
- [ ] **Campos Especiais:** Verificar se o CPF e data de nascimento estão formatados adequadamente.
- [ ] **Contas Readmitidas:** Confirmar se as contas marcadas com `READMITIDO (CPF)` pertencem de fato ao colaborador readmitido.

---

## 🔐 Recomendações de Segurança (LGPD & Git)

> ⚠️ **ATENÇÃO:**
> 1. **Credenciais:** NUNCA escreva senhas no código-fonte ou em arquivos `.bat`. Utilize sempre as variáveis de ambiente ou a janela de login da interface.
> 2. **Dados Sensíveis (LGPD):** NUNCA faça commit de arquivos `.csv` de colaboradores no repositório. As bases de dados contêm dados pessoais (CPFs, salários, endereços). O arquivo `.gitignore` já está configurado para impedir a subida acidental desses arquivos.

---

## 📦 Como Gerar o Executável (.exe) & Publicação

O sistema pode ser compilado em um arquivo executável para rodar em computadores dos operadores sem necessidade de ter Python instalado:

- **Modo Arquivo Único (`GERAR_EXE_UNICO.bat`):** Gera um único executável `dist/ImportarUsuariosUmuarama.exe` portátil com todas as dependências e `ou_map.json` embutidos.
- **Modo Pasta (`GERAR_EXE_PASTA.bat`):** Gera uma pasta `dist/ImportarUsuariosUmuarama/` contendo o executável e o arquivo `ou_map.json` externo editável.

### 🔄 Controle de Versão & Atualização do `.exe` no Portal e VM

1. **Deploy Automático no Portal:** O script `GERAR_EXE_UNICO.bat` copia automaticamente a nova versão para `\\10.56.24.17\apps\portal\downloads\ImportarUsuariosUmuarama.exe`.
2. **Ambiente da VM:** O ambiente oficial em produção fica localizado em:
   ```text
   \\10.56.24.17\apps\IMPORT_COLABORADORES
   ```
3. Assim que a compilação finalizar, os operadores que baixarem o `.exe` através do **Portal de Serviços TI** (`http://10.56.24.17:9090`) ou acessarem a VM já receberão a versão atualizada com a nova aba de Gestão AD.

---

## 📁 Estrutura do Projeto

```text
IMPORT_COLABORADORES/
├── EXECUTAR_IMPORT_COMPLETO.bat   # Script principal de execução diária (Workflow completo)
├── CADASTRAR_CREDENCIAIS.bat      # Utilitário para salvar acesso ao servidor no Windows
├── GERAR_EXE_UNICO.bat            # Gera o executável em arquivo único (--onefile)
├── GERAR_EXE_PASTA.bat            # Gera o executável em modo pasta (--onedir)
├── ExecutarImportAD.bat           # Atalho para inicializar apenas a GUI do AD
├── EXECUTAR_IMPORT.bat            # Execução de importação do FPRE111
├── gerar_importar_usuarios_uap.py # Engine Python de ETL, limpeza e padronização do FPRE111
├── ImportarUsuariosUmuarama.py     # Aplicação Tkinter (3 abas) + Gestão e Provisionamento AD
├── ImportarUsuariosUmuarama.spec  # Especificação de build do PyInstaller
├── ler_credenciais.ps1            # Leitura segura de credenciais do Windows Credential Manager
├── ou_map.json                     # Mapeamento dinâmico e customizável de OUs do Active Directory
├── dist/                          # Executável compilado (ImportarUsuariosUmuarama.exe)
├── logs/                          # Logs de auditoria e relatórios de importação em CSV
├── old/                           # Scripts e módulos legados arquivados
├── .gitignore                      # Filtro de segurança contra vazamento de dados no Git
└── README.md                       # Guia de documentação técnica do sistema
```