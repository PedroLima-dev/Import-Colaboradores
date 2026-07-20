# 🚀 Automação de Importação de Colaboradores - Grupo Umuarama

> **Documentação Técnica para a Equipe de TI / Analistas**
> 
> *Este projeto automatiza o fluxo de extração, tratamento de dados (ETL) e provisionamento de novos colaboradores no Active Directory (AD) e sistemas internos do Grupo Umuarama.*

---

## 📌 Visão Geral da Arquitetura

O sistema opera em duas etapas encadeadas:

1. **Processamento e Higienização de Dados (ETL):**
   - Lê a base bruta de exportação do sistema Senior (`FPRE111-*.CSV`).
   - Normaliza nomes, CPFs, e-mails, telefones, departamentos, unidades e cargos (padronizando nomenclaturas para o formato neutro/gênero ex: `SECRETARIO(A)`).
   - Filtra os colaboradores com base na data de admissão e gera a planilha formatada no caminho `DD-MM/importar-usuarios-DDMMYYYY.csv`.

2. **Provisionamento no Active Directory (AD):**
   - Lê o arquivo formatado.
   - Conecta ao Domain Controller (DC) via comandos integrados em PowerShell / Directory Services (`DirectoryEntry`).
   - Mapeia automaticamente a **Organizational Unit (OU)** correta no AD com base na Unidade e Concessionária.
   - Cria a conta do usuário, gera a senha inicial temporária e configura a obrigatoriedade de alteração de senha no primeiro logon.

---

## 🛠️ Pré-requisitos & Configuração do Ambiente

### 1. Requisitos do Sistema
- **Sistema Operacional:** Windows 10 / 11 ou Windows Server (necessário para execução dos scripts PowerShell / AD).
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

## 💻 Modo de Usar (Guia Passo a Passo)

### 🔴 Método 1: Execução Automática (Fluxo Padrão no Dia a Dia)

Este é o modo recomendado para a rotina diária do analista. O arquivo em lote cuida da montagem de rede, busca do arquivo mais recente e abertura da interface.

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
     - Processar e gerar o CSV de admissões formatado dentro da pasta por data (ex: `16-03/importar-usuarios-16032026.csv`).
     - Abrir a **Interface Gráfica (GUI)** pré-carregada com os dados.

3. **Conferência & Importação:**
   - Na tela da interface gráfica, revise a lista de usuários a serem criados.
   - Digite suas credenciais com acesso de Admin no Active Directory (AD).
   - Clique no botão para realizar a importação e criação dos usuários no domínio.

---

### 🟡 Método 2: Execução Manual / CLI (Para Testes e Manutenção)

Se você precisar rodar os scripts separadamente via linha de comando:

#### 1. Executar o Tratamento dos Dados (ETL)
```bash
python gerar_importar_usuarios_uap.py "\\10.56.43.28\caminho\FPRE111-16032026.CSV"
```
*Dica:* Caso queira filtrar por uma data de admissão específica (diferente da mais recente):
```bash
python gerar_importar_usuarios_uap.py "FPRE111-16032026.CSV" "16/03/2026"
```

#### 2. Abrir a Interface do AD com Pré-carregamento
```bash
python ImportarUsuariosUmuarama.py --csv-preload "16-03/importar-usuarios-16032026.csv"
```

---

## 🔍 Checklist de Validação ("Double-Check")

Antes de efetuar o provisionamento no AD ou na Intranet, o Analista Júnior deve verificar:

- [ ] **Padronização de Cargos:** Confirmar se os cargos foram higienizados (ex: sem sufixos como `SR`, `JR`, `PLENO` desnecessários e com formatação correta).
- [ ] **Mapeamento de OUs:** Garantir que a Unidade e Marca correspondem a uma OU válida no mapa (`OU_MAP`).
- [ ] **E-mails e Superior:** Confirmar se o e-mail do colaborador e o e-mail do gestor direto estão preenchidos no padrão `nome.sobrenome@grupoumuarama.com.br`.
- [ ] **Campos Especiais:** Verificar se o CPF e data de nascimento estão formatados adequadamente.

---

## 🔐 Recomendações de Segurança (LGPD & Git)

> ⚠️ **ATENÇÃO JÚNIOR:**
> 1. **Credenciais:** NUNCA escreva senhas no código fonte ou em arquivos `.bat`. Utilize sempre as variáveis de ambiente ou a janela de login da interface.
> 2. **Dados Sensíveis (LGPD):** NUNCA faça commit de arquivos `.csv` de colaboradores no GitHub. As bases de dados contêm dados pessoais (CPFs, salários, endereços). O arquivo `.gitignore` já está configurado para impedir a subida acidental desses arquivos.

---

## 📁 Estrutura do Projeto

```text
IMPORT_COLABORADORES/
├── EXECUTAR_IMPORT_COMPLETO.bat   # Script principal de execução diária (Workflow completo)
├── CADASTRAR_CREDENCIAIS.bat      # Utilitário para salvar acesso ao servidor no Windows
├── ExecutarImportAD.bat           # Atalho para inicializar apenas a GUI do AD
├── gerar_importar_usuarios_uap.py # Engine Python de ETL, limpeza e padronização do FPRE111
├── ImportarUsuariosUmuarama.py     # Interface Tkinter + Script de provisionamento no AD via PowerShell
├── .gitignore                      # Filtro de segurança contra vazamento de dados no Git
└── README.md                       # Guia de documentação do sistema
```