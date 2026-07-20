@ECHO OFF

:: Extrai o dia, mês e ano da variável %DATE%
:: Nota: Isso assume que sua data está no formato DD/MM/AAAA (padrão Brasil)
SET DIA=%DATE:~0,2%
SET MES=%DATE:~3,2%
SET ANO=%DATE:~6,4%

:: Monta a variável da data no formato DDMMYYYY
SET DATA_HOJE=%DIA%%MES%%ANO%

:: Define o caminho completo usando a variável da data
SET CAMINHO_ARQUIVO="C:\Users\Pedro Lima\Documents\IMPORT_COLABORADORES\FPRE111-%DATA_HOJE%.CSV"

:: Executa o Python passando o caminho gerado
python gerar_importar_usuarios.py %CAMINHO_ARQUIVO%

PAUSE