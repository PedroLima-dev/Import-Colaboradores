"""
Automação: FPRE111 → importar-usuarios.csv
Uso: python gerar_importar_usuarios.py <arquivo_FPRE111.XLS> [data_admissao]

- Se data_admissao não for informada, usa a data mais recente do arquivo.
- Filtra apenas os funcionários com DATA DE ADMISSÃO igual à data alvo.
- Aplica todas as transformações de campos conforme mapeamento identificado.
"""

import sys
import os
import subprocess
import pandas as pd
from datetime import datetime

# ─── Mapeamentos ─────────────────────────────────────────────────────────────

# Cargos que devem ser padronizados com (A)
# Regra: substitui a palavra-chave pelo padrão PALAVRA(A)
CARGO_PADRONIZACAO = {
    "ADMINISTRADOR":    "ADMINISTRADOR(A)",
    "AGENTE":           "AGENTE",           # neutro, não se aplica
    "ANALISTA":         "ANALISTA",         # neutro
    "ASSESSOR":         "ASSESSOR(A)",
    "ASSISTENTE":       "ASSISTENTE",       # neutro
    "AUXILIAR":         "AUXILIAR",         # neutro
    "AVALIADOR":        "AVALIADOR(A)",
    "COMPRADOR":        "COMPRADOR(A)",
    "CONSULTOR":        "CONSULTOR(A)",
    "CONTROLADOR":      "CONTROLADOR(A)",
    "COORDENADOR":      "COORDENADOR(A)",
    "DESIGNER":         "DESIGNER",         # neutro
    "DIRETOR":          "DIRETOR(A)",
    "ENCARREGADO":      "ENCARREGADO(A)",
    "ENGENHEIRO":       "ENGENHEIRO(A)",
    "ENTREGADOR":       "ENTREGADOR(A)",
    "ESPECIALISTA":     "ESPECIALISTA",     # neutro
    "ESTAGIARIO":       "ESTAGIARIO(A)",
    "ESTOQUISTA":       "ESTOQUISTA",       # neutro
    "GERENTE":          "GERENTE",          # neutro
    "GESTOR":           "GESTOR(A)",
    "LAVADOR":          "LAVADOR(A)",
    "LIDER":            "LIDER",            # neutro
    "LÍDER":            "LIDER",            # neutro
    "MECANICO":         "MECANICO(A)",
    "MONTADOR":         "MONTADOR(A)",
    "MOTORISTA":        "MOTORISTA",        # neutro
    "PINTOR":           "PINTOR(A)",
    "PORTEIRO":         "PORTEIRO(A)",
    "PRECIFICADOR":     "PRECIFICADOR(A)",
    "PREPARADOR":       "PREPARADOR(A)",
    "RECEPCIONISTA":    "RECEPCIONISTA",    # neutro
    "SECRETARIA":       "SECRETARIO(A)",
    "SUPERVISOR":       "SUPERVISOR(A)",
    "TECNICO":          "TECNICO(A)",
}

DESABREVIACOES = {
    "SEG.": "SEGURANCA",
    "ADM.": "ADMINISTRACAO",
    "TEC.": "TECNICO"
    # Adicione outras abreviações conforme for mapeando a necessidade
}

# Qualificações de nível a remover de cargos ANALISTA (tokens individuais)
QUALIFICACOES_ANALISTA = {"SENIOR", "PLENO", "JUNIOR", "SR", "JR", "I", "II", "III", "IV"}

def padronizar_cargo(cargo):
    """Aplica padronização de gênero (A) ao cargo, preservando o restante da descrição.
    Cargos que começam com 'A.' são tratados como ANALISTA.
    Qualificações de nível (SENIOR, PLENO, JUNIOR, SR, I, II, III...) são removidas de ANALISTA.
    """
    if not cargo or str(cargo).strip() == "":
        return cargo
    cargo_str = str(cargo).strip().upper()

    # Normaliza abreviação A. → ANALISTA
    if cargo_str.upper().startswith("A."):
        cargo_str = "ANALISTA" + cargo_str[2:]
    
    # --- NOVA PARTE: Desabreviar termos no meio da string ---
    for abrev, completo in DESABREVIACOES.items():
        if abrev in cargo_str:
            cargo_str = cargo_str.replace(abrev, completo)

    # Já está padronizado (contém "(A)")
    if "(A)" in cargo_str.upper():
        return cargo_str

    upper = cargo_str.upper()
    for palavra, substituto in CARGO_PADRONIZACAO.items():
        if upper.startswith(palavra):
            sufixo = cargo_str[len(palavra):]
            # Remove qualificações de nível somente para ANALISTA
            if palavra == "ANALISTA":
                tokens = sufixo.split()
                tokens_limpos = [t for t in tokens if t.upper() not in QUALIFICACOES_ANALISTA]
                sufixo = " ".join(tokens_limpos)
                if sufixo:
                    sufixo = " " + sufixo.strip()
                return substituto + sufixo
            return substituto + sufixo
    return cargo_str


# FILIAL → prefixo de marca para o campo "unidade"
FILIAL_MARCA = {
    "UMUARAMA MOTORS COMERCIO E SERVICOS LTDA": "TOYOTA",
    "UMUARAMA MOTORS COMERCIO E SERVIÇOS LTDA": "TOYOTA",
    "UMUARAMA AUTOS LTDA": "VOLKS",
    "UMUARAMA AUTOMOVEIS LTDA": "FIAT",
    "UMUARAMA VEICULOS LTDA": "KIA",
    "UMUARAMA MOTOCICLETAS LTDA": "TRIUMPH",
    "UMUARAMA MOTOS LTDA": "HD",
    "UMUARAMA MOTOS UBR": "HD",
    "UMUARAMA MOTOSPORT LTDA": "KTM",
    "UMUARAMA ADM. E PART. CONCESSIONARIA LTD": "HOLDING UAC",
    "UMUARAMA ADM E CORRETORA DE SEGUROS LTDA": "CORRETORA DE SEGUROS",
    "UMUARAMA ADMINISTRACAO E PARTICIPACAO AG": "HOLDING UAC",
    "UMUARAMA ADMINISTRACAO E PARTICIPACAO LT": "HOLDING UAC",
    "UMUARAMA ADMINISTRACAO E PARTICIPACAO AG" : "HOLDING UAA",
    "UMUARAMA ADMINISTRACAO E PARTICIPACAO LT" : "HOLDING UAP",
}

# CENTRO DE CUSTO → departamento (normalização)
DEPTO_MAP = {
    "PEÇAS VAREJO"              : "POS VENDAS - PECAS E ACESSORIOS",
    "PC - PECAS"                : "POS VENDAS - PECAS E ACESSORIOS",
    "PECAS"                     :"POS VENDAS - PECAS E ACESSORIOS",
    "ACESSORIOS"                : "POS VENDAS - PECAS E ACESSORIOS",
    "MECANICA"                  : "POS VENDAS - MECANICA",
    "OFICINA MECANICA"          : "POS VENDAS - MECANICA",
    "VEICULOS NOVOS OFICINA"    : "POS VENDAS - MECANICA",
    "VU - VEICULOS USADOS"      : "VEICULOS USADOS",
    "VN - VEICULOS NOVOS"       : "VEICULOS NOVOS",
    "Veiculos Novos"            : "VEICULOS NOVOS",
    "FU - FUNILARIA"            : "FUNILARIA E PINTURA",
    "SC - SERVICOS CARROCERIA"  : "FUNILARIA E PINTURA",
    "AT-ASSISTENCIA TECNICA SERV. GERAIS": "ASSISTENCIA TECNICA",
    "AD - ADMINISTRACAO"        : "ADMINISTRACAO",
    "ADMINISTRAÇÃO"             : "ADMINISTRACAO",
    "ADMINISTRATIVO"            : "ADMINISTRACAO",
    "APOIO ADMINISTRATIVO"      : "ADMINISTRACAO",
    "FINACEIRO"                 : "FINANCEIRO",
    "AGENDAMENTO"               : "CENTRAL DE RELACIONAMENTO - AGENDAMENTO",
    "PV - POS VENDAS"           : "POS VENDAS",
    "corretora seguros"         : "CORRETORA DE SEGUROS",
    "ACESSOROIS E BOUTIQUE"     : "PECAS E BOUTIQUE",
    "SEGURANCA DO TRABALHO/AMBIENTAL" : "SEGURANCA DO TRABALHO",
}


def converter_xls_para_xlsx(xls_path):
    """Converte .XLS para .xlsx usando LibreOffice."""
    out_dir = "/home/claude"
    subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "xlsx", xls_path, "--outdir", out_dir],
        check=True, capture_output=True
    )
    base = os.path.splitext(os.path.basename(xls_path))[0]
    return os.path.join(out_dir, base + ".xlsx")


def normalizar_depto(centro_custo):
    """Aplica mapeamento de CENTRO DE CUSTO → departamento."""
    if pd.isna(centro_custo):
        return ""
    return DEPTO_MAP.get(str(centro_custo).strip(), str(centro_custo).strip())


def montar_unidade(filial, local):
    """Constrói o campo 'unidade' = marca + ' - ' + LOCAL, ou só 'HOLDING UAC' ou só 'HOLDING UAP'."""
    filial = str(filial).strip() if not pd.isna(filial) else ""
    local = str(local).strip() if not pd.isna(local) else ""
    marca = FILIAL_MARCA.get(filial, filial)
    if marca == "HOLDING UAC":
        return "HOLDING UAC"
    elif marca == "HOLDING UAP":
        return "HOLDING UAP"
    return f"{marca} - {local}" if local else marca


def limpar_telefone(numero):
    """Remove zeros à esquerda e zeros desnecessários."""
    if pd.isna(numero):
        return ""
    s = str(numero).strip()
    if s in ("0", "00", ""):
        return ""
    # Remove zero à esquerda (ex: 064... -> 64...)
    s = s.lstrip("0")
    # Recoloca o 0 se DDD ficou com 1 dígito (improvável, mas defensivo)
    return s


def processar(xls_path, data_alvo=None):
    ext = os.path.splitext(xls_path)[1].upper()

    if ext == ".CSV":
        df = None
        for enc in ("latin1", "utf-8", "utf-8-sig"):
            try:
                df = pd.read_csv(xls_path, dtype=str, encoding=enc, sep=None, engine="python")
                print(f"Lido como CSV (encoding={enc})") # Checagem de leitura
                break
            except Exception:
                continue
        if df is None:
            raise ValueError("Nao foi possivel ler o arquivo CSV.")
    elif ext == ".XLS":
        print(f"Convertendo {xls_path} para xlsx...")
        xlsx_path = converter_xls_para_xlsx(xls_path)
        df = pd.read_excel(xlsx_path, header=0, dtype=str)
    else:
        df = pd.read_excel(xls_path, header=0, dtype=str)

    df.columns = df.columns.str.strip()

    # Determina a data alvo
    if data_alvo:
        alvo = pd.to_datetime(data_alvo, dayfirst=True)
    else:
        datas = pd.to_datetime(df["DATA DE ADMISSÃO"], dayfirst=True, errors="coerce")
        alvo = datas.max()
        print(f"Data de admissão mais recente encontrada: {alvo.strftime('%d/%m/%Y')}")

    # Filtra funcionários da data alvo
    datas_col = pd.to_datetime(df["DATA DE ADMISSÃO"], dayfirst=True, errors="coerce")
    df_filtrado = df[datas_col == alvo].copy()
    print(f"Funcionários encontrados para {alvo.strftime('%d/%m/%Y')}: {len(df_filtrado)}")

    if df_filtrado.empty:
        print("Nenhum funcionário encontrado para essa data.")
        return None

    # Monta o CSV de saída
    rows = []
    for _, r in df_filtrado.iterrows():
        unidade = montar_unidade(r.get("FILIAL"), r.get("LOCAL"))
        departamento = normalizar_depto(r.get("CENTRO DE CUSTO"))
        subunidade = str(r.get("LOCAL", "")).strip()
        telefone = limpar_telefone(r.get("CELULAR"))

        row = {
            "nome": r.get("NOME", ""),
            "e-mail": "",
            "admissao": alvo.strftime("%d/%m/%Y"),
            "data de nascimento": str(r.get("DATA DE NASCIMENTO", "")).strip(),
            "matricula": str(r.get("MATRICULA", "")).strip() if not pd.isna(r.get("MATRICULA")) else "",
            "ramal": "",
            "cpf": str(r.get("CPF", "")).strip(),
            "telefone": "",
            "residencial": "",
            "celular": telefone,
            "unidade": unidade,
            "departamento": departamento,
            "subunidade": subunidade,
            "cargo": padronizar_cargo(r.get("CARGO", "")),
            "email superior": str(r.get("E-MAIL DO SUPERIOR", "")).strip() if not pd.isna(r.get("E-MAIL DO SUPERIOR")) else "",
            "email assessor": "",
            "sexo": str(r.get("SEXO", "")).strip(),
            "estado civil": str(r.get("ESTADO CIVIL", "")).strip(),
            "naturalidade": "",
            "tem filhos": "",
            "numero filhos": "",
            "fuma": "",
            "camiseta": "",
            "calcado": "",
            "email secundario": str(r.get("E-MAIL PARTICULAR", "")).strip() if not pd.isna(r.get("E-MAIL PARTICULAR")) else "",
            "cep": str(r.get("CEP", "")).strip(),
            "endereco": str(r.get("ENDEREÇO", "")).strip() if not pd.isna(r.get("ENDEREÇO")) else "",
            "bairro": str(r.get("BAIRRO", "")).strip(),
            "complemento": "",
            "cidade": str(r.get("CIDADE", "")).strip(),
            "uf": str(r.get("UF", "")).strip(),
            "contato emergencia": "",
            "telefone emergencia": "",
            "contato emergencia secundario": "",
            "telefone emergencia alternativo": "",
            "peso": "",
            "altura": "",
            "grupo sanguineo": "",
            "cnpj": "",
            "inativo": "",
            "codigo externo 1": str(r.get("MATRICULA (ESOCIAL)", "")).strip(),
            "codigo externo 2": "",
            "codigo externo 3": "",
            "desligamento": "",
            "motivo de desligamento": "",
        }
        rows.append(row)

    df_out = pd.DataFrame(rows)

    # Remove pontuações de todas as colunas de texto do output
    import re
    def remover_pontuacao(val):
        if isinstance(val, str):
            return re.sub(r"[^\w\s@.\-/()\u00C0-\u00FF]", "", val)
        return val
    df_out = df_out.map(remover_pontuacao)

    # Nome do arquivo de saída com a data
    data_str = alvo.strftime("%d%m%Y")
    data_pasta = alvo.strftime("%d-%m")
    base_dir = os.path.dirname(os.path.abspath(xls_path))
    out_dir = os.path.join(base_dir, data_pasta)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"importar-usuarios-{data_str}.csv")
    df_out.to_csv(out_path, index=False, encoding="latin1")

    # ─── Relatório de verificação ────────────────────────────────────────────
    CARGOS_ESPECIFICOS = [
        "AUXILIAR DE OFICINA", "MECANICO", "MOTORISTA", "PINTOR",
        "ALINHADOR DE VEICULOS", "AUXILIAR DE PINTOR", "COPEIRA",
        "LAVADOR", "PORTEIRO", "VIGIA",
    ]

    total = len(df_filtrado)

    # Cargos específicos — lista cada cargo encontrado e seus funcionários
    cargos_encontrados = []
    for cargo_ref in CARGOS_ESPECIFICOS:
        mask = df_filtrado["CARGO"].fillna("").str.upper().str.startswith(cargo_ref.upper())
        nomes = df_filtrado.loc[mask, "NOME"].tolist()
        if nomes:
            cargos_encontrados.append((cargo_ref, nomes))

    total_especificos = sum(len(n) for _, n in cargos_encontrados)

    # Sem e-mail do superior
    mask_sem_email = df_filtrado["E-MAIL DO SUPERIOR"].isna() | (df_filtrado["E-MAIL DO SUPERIOR"].str.strip() == "")
    nomes_sem_email = df_filtrado.loc[mask_sem_email, "NOME"].tolist()

    total_cadastrados = total - total_especificos

    print(f"")
    print(f"Arquivo gerado: {out_path}")
    print(f"─" * 55)
    print(f"Total de funcionarios: {total}")
    if cargos_encontrados:
        for cargo_ref, nomes in cargos_encontrados:
            for nome in nomes:
                print(f"Funcionarios em cargos especificos: {len(nomes)} - {nome} ({cargo_ref})")
    else:
        print(f"Funcionarios em cargos especificos: 0")
    if nomes_sem_email:
        for nome in nomes_sem_email:
            print(f"Funcionarios sem e-mail de superior cadastrado: {len(nomes_sem_email)} - {nome}")
    else:
        print(f"Funcionarios sem e-mail de superior cadastrado: 0")
    print(f"Total de funcionarios cadastrados: {total_cadastrados}")
    print(f"─" * 55)

    return out_path


FILIAIS_UAP = {
    "UMUARAMA ADMINISTRACAO E PARTICIPACAO AG",
    "UMUARAMA ADMINISTRACAO E PARTICIPACAO LT",
}


def processar_uap(xls_path, data_pasta="16-04"):
    """Gera importação exclusiva das filiais UAP, sem filtro de data de admissão."""
    ext = os.path.splitext(xls_path)[1].upper()

    if ext == ".CSV":
        df = None
        for enc in ("latin1", "utf-8", "utf-8-sig"):
            try:
                df = pd.read_csv(xls_path, dtype=str, encoding=enc, sep=None, engine="python")
                print(f"Lido como CSV (encoding={enc})")
                break
            except Exception:
                continue
        if df is None:
            raise ValueError("Nao foi possivel ler o arquivo CSV.")
    elif ext == ".XLS":
        print(f"Convertendo {xls_path} para xlsx...")
        xlsx_path = converter_xls_para_xlsx(xls_path)
        df = pd.read_excel(xlsx_path, header=0, dtype=str)
    else:
        df = pd.read_excel(xls_path, header=0, dtype=str)

    df.columns = df.columns.str.strip()

    # Filtra apenas as filiais UAP (sem filtro de data)
    df_filtrado = df[df["FILIAL"].str.strip().isin(FILIAIS_UAP)].copy()
    print(f"Funcionarios das filiais UAP encontrados: {len(df_filtrado)}")

    if df_filtrado.empty:
        print("Nenhum funcionario UAP encontrado.")
        return None

    # Monta o CSV de saída
    rows = []
    for _, r in df_filtrado.iterrows():
        unidade = montar_unidade(r.get("FILIAL"), r.get("LOCAL"))
        departamento = normalizar_depto(r.get("CENTRO DE CUSTO"))
        subunidade = str(r.get("LOCAL", "")).strip()
        telefone = limpar_telefone(r.get("CELULAR"))
        admissao = str(r.get("DATA DE ADMISSÃO", "")).strip()

        row = {
            "nome": r.get("NOME", ""),
            "e-mail": "",
            "admissao": admissao,
            "data de nascimento": str(r.get("DATA DE NASCIMENTO", "")).strip(),
            "matricula": str(r.get("MATRICULA", "")).strip() if not pd.isna(r.get("MATRICULA")) else "",
            "ramal": "",
            "cpf": str(r.get("CPF", "")).strip(),
            "telefone": telefone,
            "residencial": "",
            "celular": "",
            "unidade": unidade,
            "departamento": departamento,
            "subunidade": subunidade,
            "cargo": padronizar_cargo(r.get("CARGO", "")),
            "email superior": str(r.get("E-MAIL DO SUPERIOR", "")).strip() if not pd.isna(r.get("E-MAIL DO SUPERIOR")) else "",
            "email assessor": "",
            "sexo": str(r.get("SEXO", "")).strip(),
            "estado civil": str(r.get("ESTADO CIVIL", "")).strip(),
            "naturalidade": "",
            "tem filhos": "",
            "numero filhos": "",
            "fuma": "",
            "camiseta": "",
            "calcado": "",
            "email secundario": str(r.get("E-MAIL PARTICULAR", "")).strip() if not pd.isna(r.get("E-MAIL PARTICULAR")) else "",
            "cep": str(r.get("CEP", "")).strip(),
            "endereco": str(r.get("ENDEREÇO", "")).strip() if not pd.isna(r.get("ENDEREÇO")) else "",
            "bairro": str(r.get("BAIRRO", "")).strip(),
            "complemento": "",
            "cidade": str(r.get("CIDADE", "")).strip(),
            "uf": str(r.get("UF", "")).strip(),
            "contato emergencia": "",
            "telefone emergencia": "",
            "contato emergencia secundario": "",
            "telefone emergencia alternativo": "",
            "peso": "",
            "altura": "",
            "grupo sanguineo": "",
            "cnpj": "",
            "inativo": "",
            "codigo externo 1": str(r.get("MATRICULA (ESOCIAL)", "")).strip(),
            "codigo externo 2": "",
            "codigo externo 3": "",
            "desligamento": "",
            "motivo de desligamento": "",
        }
        rows.append(row)

    df_out = pd.DataFrame(rows)

    # Remove pontuações
    import re
    def remover_pontuacao(val):
        if isinstance(val, str):
            return re.sub(r"[^\w\s@.\-/()\u00C0-\u00FF]", "", val)
        return val
    df_out = df_out.map(remover_pontuacao)

    # Salva na pasta fixa 16-04
    base_dir = os.path.dirname(os.path.abspath(xls_path))
    out_dir = os.path.join(base_dir, data_pasta)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "import_usuarios_uap16042026.csv")
    df_out.to_csv(out_path, index=False, encoding="latin1")

    # ─── Relatório ───────────────────────────────────────────────────────────
    CARGOS_ESPECIFICOS = [
        "AUXILIAR DE OFICINA", "MECANICO", "MOTORISTA", "PINTOR",
        "ALINHADOR DE VEICULOS", "AUXILIAR DE PINTOR", "COPEIRA",
        "LAVADOR", "PORTEIRO", "VIGIA",
    ]

    total = len(df_filtrado)
    cargos_encontrados = []
    for cargo_ref in CARGOS_ESPECIFICOS:
        mask = df_filtrado["CARGO"].fillna("").str.upper().str.startswith(cargo_ref.upper())
        nomes = df_filtrado.loc[mask, "NOME"].tolist()
        if nomes:
            cargos_encontrados.append((cargo_ref, nomes))

    total_especificos = sum(len(n) for _, n in cargos_encontrados)

    mask_sem_email = df_filtrado["E-MAIL DO SUPERIOR"].isna() | (df_filtrado["E-MAIL DO SUPERIOR"].str.strip() == "")
    nomes_sem_email = df_filtrado.loc[mask_sem_email, "NOME"].tolist()

    total_cadastrados = total - total_especificos

    print(f"")
    print(f"Arquivo gerado: {out_path}")
    print(f"─" * 55)
    print(f"Total de funcionarios: {total}")
    if cargos_encontrados:
        for cargo_ref, nomes in cargos_encontrados:
            for nome in nomes:
                print(f"Funcionarios em cargos especificos: {len(nomes)} - {nome} ({cargo_ref})")
    else:
        print(f"Funcionarios em cargos especificos: 0")
    if nomes_sem_email:
        for nome in nomes_sem_email:
            print(f"Funcionarios sem e-mail de superior cadastrado: {len(nomes_sem_email)} - {nome}")
    else:
        print(f"Funcionarios sem e-mail de superior cadastrado: 0")
    print(f"Total de funcionarios cadastrados: {total_cadastrados}")
    print(f"─" * 55)

    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso:")
        print("  python gerar_importar_usuarios.py <arquivo> [dd/mm/aaaa]         # importacao semanal normal")
        print("  python gerar_importar_usuarios.py <arquivo> --uap [pasta]        # importacao filiais UAP")
        sys.exit(1)

    xls_path = sys.argv[1]
    if len(sys.argv) > 2 and sys.argv[2] == "--uap":
        pasta = sys.argv[3] if len(sys.argv) > 3 else "16-04"
        processar_uap(xls_path, pasta)
    else:
        data_alvo = sys.argv[2] if len(sys.argv) > 2 else None
        processar(xls_path, data_alvo)
