"""
Automação: FPRE111 → importar-usuarios.csv
Uso: python gerar_importar_usuarios.py <arquivo_FPRE111.XLS> [data_admissao]

- Se data_admissao não for informada, usa a data mais recente do arquivo.
- Filtra apenas os funcionários com DATA DE ADMISSÃO igual à data alvo.
- Aplica todas as transformações de campos conforme mapeamento identificado.
"""

import re
import sys
import os
import shutil
import tempfile
import subprocess
import unicodedata
import pandas as pd
from contextlib import contextmanager
from datetime import datetime

# ─── Resolução de diretório base (PyInstaller / Script) ──────────────────────
def _obter_dir_base():
    """Retorna o diretório base da aplicação (onde o .exe ou o .py está localizado)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

# ─── Segurança: isola arquivos de rede ───────────────────────────────────────

DIR_LOCAL_SEGURO = os.path.join(_obter_dir_base(), "_tmp_local")

@contextmanager
def arquivo_local_seguro(path):
    """Context manager de segurança para arquivos de rede UNC (\\\\servidor\\...).

    Se o caminho for de rede:
      1. Cria um arquivo temporário local com mkstemp
      2. Copia o conteúdo uma única vez (1 tráfego SMB)
      3. Fornece o caminho local ao bloco 'with'
      4. No bloco 'finally', SEMPRE destrói o temporário —
         inclusive se copy2 falhar (arquivo vazio do mkstemp),
         se o processamento lançar qualquer exceção, ou se
         o próprio os.remove falhar (nesse caso loga aviso sem
         mascarar a exceção original).

    Se o caminho for local, simplesmente repassa sem fazer nada.
    """
    if not (path.startswith("\\\\") or path.startswith("//")):
        yield path  # caminho local: nenhuma copia necessaria
        return

    os.makedirs(DIR_LOCAL_SEGURO, exist_ok=True)
    ext = os.path.splitext(path)[1]
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=ext, dir=DIR_LOCAL_SEGURO)
    os.close(tmp_fd)
    try:
        print(f"[SEGURANCA] Arquivo de rede detectado: {path}")
        print(f"[SEGURANCA] Copiando para diretorio local seguro antes de processar...")
        shutil.copy2(path, tmp_path)  # unico trafego SMB
        print(f"[SEGURANCA] Copia local criada: {tmp_path}")
        yield tmp_path                # fornece o caminho seguro ao chamador
    finally:
        # Destruicao garantida: cobre falha em copy2, excecao no processamento
        # e qualquer outro caminho de saida (return, raise, sys.exit)
        try:
            os.remove(tmp_path)
            print(f"[SEGURANCA] Arquivo temporario destruido: {tmp_path}")
        except OSError as e:
            # Nao mascara a excecao original — apenas registra o aviso
            print(f"[SEGURANCA][AVISO] Nao foi possivel destruir o temporario "
                  f"{tmp_path}: {e}")


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
    "SECRETARIO":       "SECRETARIO(A)",
    "SUPERVISOR":       "SUPERVISOR(A)",
    "TECNICO":          "TECNICO(A)",
    "ASSISTENTE ADM. FINANC. PONTA": "ASSISTENTE ADMIN FINANCEIRO PONTA",
    "ASSIST. FIN. DE CONTAS A RECEB": "ASSISTENTE FINANCEIRO",
    "CONSULTOR DE VENDAS PCS EXTERN": "CONSULTOR(A) DE VENDAS PECAS EXTERNO",
    "ANALISTA COM. DE SEMI.": "ANALISTA COMERCIAL DE SEMINOVOS",
    "ANALISTA COMERCIAL DE SEMI": "ANALISTA COMERCIAL DE SEMINOVOS",
}


# Qualificações de nível a remover de qualquer cargo (tokens individuais)
QUALIFICACOES_ANALISTA = {"SENIOR", "PLENO", "JUNIOR", "SR", "JR", "I", "II", "III", "IV"}

def _remover_acentos(texto):
    """Remove acentos e cedilha de uma string, preservando caracteres como '(' e ')'."""
    normalizado = unicodedata.normalize("NFD", texto)
    return "".join(c for c in normalizado if unicodedata.category(c) != "Mn")


def padronizar_cargo(cargo):
    """Aplica padronização de gênero (A) ao cargo, preservando o restante da descrição.
    Cargos que começam com 'A.' são tratados como ANALISTA.
    Qualificações de nível (SENIOR, PLENO, JUNIOR, SR, I, II, III...) são removidas de qualquer cargo.
    Hífens no sufixo são substituídos por espaço.
    Acentos e cedilha são removidos do resultado final.
    """
    if not cargo or str(cargo).strip() == "":
        return cargo
    cargo_str = str(cargo).strip()

    # Normaliza abreviação A. → ANALISTA
    if cargo_str.upper().startswith("A."):
        cargo_str = "ANALISTA" + cargo_str[2:]

    # Função auxiliar para normalizar as strings durante o processo de comparação (busca)
    def normalizar_para_busca(texto):
        import re
        # Remove "(A)" ou "(A) " de forma case-insensitive
        t = re.sub(r"\([aA]\)\s*", "", texto)
        # Remove acentos e cedilha
        t = _remover_acentos(t)
        # Substitui pontos e hífens por espaço
        t = t.replace(".", " ").replace("-", " ")
        # Transforma múltiplos espaços em um único espaço e remove extras nas pontas
        return " ".join(t.split()).upper()

    busca_input = normalizar_para_busca(cargo_str)

    # Itera da chave mais longa para a mais curta (com base no comprimento normalizado de busca)
    chaves_ordenadas = sorted(
        CARGO_PADRONIZACAO.items(),
        key=lambda item: len(normalizar_para_busca(item[0])),
        reverse=True
    )

    for palavra, substituto in chaves_ordenadas:
        palavra_busca = normalizar_para_busca(palavra)
        if busca_input.startswith(palavra_busca):
            # Encontra o comprimento correspondente no cargo_str original que gerou a palavra_busca
            prefix_len = 0
            for i in range(len(cargo_str) + 1):
                if normalizar_para_busca(cargo_str[:i]) == palavra_busca:
                    prefix_len = i

            if prefix_len == 0:
                prefix_len = len(palavra)

            sufixo = cargo_str[prefix_len:]
            # Substitui hífens por espaço e remove qualificações de nível (vale para qualquer cargo)
            sufixo = sufixo.replace("-", " ")
            tokens = sufixo.split()
            tokens_limpos = [t for t in tokens if t.upper() not in QUALIFICACOES_ANALISTA]
            sufixo = " ".join(tokens_limpos)
            if sufixo:
                sufixo = " " + sufixo.strip()
            return _remover_acentos(substituto + sufixo)

    # Se não deu match em nenhuma chave do dicionário, mas já possui "(A)"
    if "(A)" in cargo_str.upper():
        idx = cargo_str.upper().find("(A)")
        if idx != -1:
            prefixo = cargo_str[:idx + 3]
            sufixo = cargo_str[idx + 3:]
            sufixo = sufixo.replace("-", " ")
            tokens = sufixo.split()
            tokens_limpos = [t for t in tokens if t.upper() not in QUALIFICACOES_ANALISTA]
            sufixo = " ".join(tokens_limpos)
            if sufixo:
                sufixo = " " + sufixo.strip()
            return _remover_acentos(prefixo + sufixo)
        return _remover_acentos(cargo_str)

    # Para cargos sem mapeamento e sem "(A)", também limpamos hífens e qualificações
    cargo_limpo = cargo_str.replace("-", " ")
    tokens = cargo_limpo.split()
    tokens_limpos = [t for t in tokens if t.upper() not in QUALIFICACOES_ANALISTA]
    cargo_limpo = " ".join(tokens_limpos)
    return _remover_acentos(cargo_limpo)


# FILIAL → prefixo de marca para o campo "unidade"
FILIAL_MARCA = {
    "UMUARAMA MOTORS COMERCIO E SERVICOS LTDA": "TOYOTA",
    "UMUARAMA MOTORS COMERCIO E SERVIÇOS LTDA": "TOYOTA",
    "UMUARAMA AUTOS LTDA": "VOLKS",
    "UMUARAMA AUTOMOVEIS LTDA": "FIAT",
    "UMUARAMA AUTOMOTORES LTDA": "JEEP E RAM",
    "UMUARAMA VEICULOS LTDA": "KIA",
    "UMUARAMA MOTOCICLETAS LTDA": "TRIUMPH",
    "UMUARAMA MOTOS LTDA": "HD",
    "UMUARAMA MOTOS UBR": "HD",
    "UMUARAMA MOTOSPORT LTDA": "KTM",
    "UMUARAMA ADM. E PART. CONCESSIONARIA LTD": "HOLDING UAC",
    "UMUARAMA ADM E CORRETORA DE SEGUROS LTDA": "CORRETORA DE SEGUROS",
    "UMUARAMA ADMINISTRACAO E PARTICIPACAO AG": "HOLDING UAC",
    "UMUARAMA ADMINISTRACAO E PARTICIPACAO LT": "HOLDING UAC",
}

# CENTRO DE CUSTO → departamento (normalização)
DEPTO_MAP = {
    "PEÇAS VAREJO"              : "POS VENDAS - PECAS E ACESSORIOS",
    "PC - PECAS"                : "POS VENDAS - PECAS E ACESSORIOS",
    "PECAS"                     : "POS VENDAS - PECAS E ACESSORIOS",
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
    "AD - ADMINISTRACAO"        : "ADMINISTRATIVO",
    "ADMINISTRAÇÃO"             : "ADMINISTRATIVO",
    "ADMINISTRATIVO"            : "ADMINISTRATIVO",
    "APOIO ADMINISTRATIVO"      : "ADMINISTRATIVO",
    "FINACEIRO"                 : "FINANCEIRO",
    "AGENDAMENTO"               : "CENTRAL DE RELACIONAMENTO - AGENDAMENTO",
    "PESQUISA DE SATISFAÇÃO"    : "CENTRAL DE RELACIONAMENTO - PESQUISA DE SATISFACAO",
    "PV - POS VENDAS"           : "POS VENDAS",
    "POS-VENDA"                 : "POS VENDAS",
    "POS-VENDAS"                : "POS VENDAS",
    "corretora seguros"         : "CORRETORA DE SEGUROS",
    "ACESSOROIS E BOUTIQUE"     : "PECAS E BOUTIQUE",
    "LEADS"                     : "CENTRAL DE RELACIONAMENTO - VENDAS DIGITAIS"
}

# Normalização de acentos/typos em departamentos (aplicada APÓS DEPTO_MAP)
# Chave: valor com acento ou typo (uppercase) → Valor: forma correta sem acento
DEPTO_TYPO_MAP = {
    "ADMINISTRAÇÃO"             : "ADMINISTRATIVO",
    "FINANCEIRO"                : "FINANCEIRO",
    "FINACEIRO"                 : "FINANCEIRO",
    "PEÇAS E ACESSORIOS"        : "POS VENDAS - PECAS E ACESSORIOS",
    "PECAS E ACESSÓRIOS"        : "POS VENDAS - PECAS E ACESSORIOS",
    "FUNILÁRIA E PINTURA"       : "FUNILARIA E PINTURA",
    "ASSISTÊNCIA TÉCNICA"       : "ASSISTENCIA TECNICA",
    "VEÍCULOS NOVOS"            : "VEICULOS NOVOS",
    "VEÍCULOS USADOS"           : "VEICULOS USADOS",
}

# LOCAL (subunidade) → forma padronizada (sem acentos, sem typos)
SUBUNIDADE_MAP = {
    # Typos / singular-plural
    "MINEIRO"                   : "MINEIROS",
    "MINEIROS"                  : "MINEIROS",
    # Acentos
    "GOIANÉSIA"                 : "GOIANESIA",
    "GOIANESIA"                 : "GOIANESIA",
    "CATALÃO"                   : "CATALAO",
    "CATALAO"                   : "CATALAO",
    "ANÁPOLIS"                  : "ANAPOLIS",
    "ANAPOLIS"                  : "ANAPOLIS",
    "GOIÂNIA"                   : "GOIANIA",
    "GOIANIA"                   : "GOIANIA",
    "BRASÍLIA"                  : "BRASILIA",
    "BRASILIA"                  : "BRASILIA",
    "UBERLÂNDIA"                : "UBERLANDIA",
    "UBERLANDIA"                : "UBERLANDIA",
    "SÃO PAULO"                 : "SAO PAULO",
    "SAO PAULO"                 : "SAO PAULO",
    "UMUARAMA"                  : "UMUARAMA",
    # Cedilha
    "URUA\u00c7U"                   : "URUACU",
    "URUACU"                    : "URUACU",
    # Redenção
    "REDEN\u00c7\u00c3O"                  : "REDENCAO",
    "REDENCAO"                  : "REDENCAO",
}


def normalizar_depto(centro_custo):
    """Aplica mapeamento de CENTRO DE CUSTO → departamento e corrige acentos/typos."""
    if pd.isna(centro_custo):
        return ""
    valor = DEPTO_MAP.get(str(centro_custo).strip(), str(centro_custo).strip())
    # Aplica correção de acentos/typos sobre o resultado
    return DEPTO_TYPO_MAP.get(valor.upper(), valor)


def normalizar_subunidade(local):
    """Padroniza o campo LOCAL (subunidade) removendo acentos e corrigindo typos."""
    if pd.isna(local) or str(local).strip() == "":
        return ""
    local_upper = str(local).strip().upper()
    return SUBUNIDADE_MAP.get(local_upper, str(local).strip().upper())


def montar_unidade(filial, local):
    """Constrói o campo 'unidade' = marca + ' - ' + LOCAL normalizado, ou só a marca para HOLDING e CORRETORA."""
    filial = str(filial).strip() if not pd.isna(filial) else ""
    local_norm = normalizar_subunidade(local)
    marca = FILIAL_MARCA.get(filial, filial)
    if marca in ("HOLDING UAC", "HOLDING UAP", "CORRETORA DE SEGUROS", "CORRETORA"):
        return marca
    return f"{marca} - {local_norm}" if local_norm else marca


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


CARGOS_ESPECIFICOS = [
    "AUXILIAR DE OFICINA", "MECANICO", "MOTORISTA", "PINTOR",
    "ALINHADOR DE VEICULOS", "AUXILIAR DE PINTOR", "COPEIRA",
    "LAVADOR", "PORTEIRO", "VIGIA", "SERVICOS GERAIS",
]


def remover_pontuacao(val):
    if isinstance(val, str):
        return re.sub(r"[^\w\s@.\-/()\u00C0-\u00FF]", "", val)
    return val


ESTADO_CIVIL_MAP = {
    "DIVOLCIADO": "Divorciado",
    "DIVOLCIADA": "Divorciada",
    "DIVORCIADO": "Divorciado",
    "DIVORCIADA": "Divorciada",
    "SOLTEIRO": "Solteiro",
    "SOLTEIRA": "Solteira",
    "CASADO": "Casado",
    "CASADA": "Casada",
    "SEPARADO": "Separado",
    "SEPARADA": "Separada",
    "VIUVO": "Viuvo",
    "VIUVA": "Viuva",
    "VIÚVO": "Viuvo",
    "VIÚVA": "Viuva",
    "UNIAO ESTAVEL": "Uniao Estavel",
    "UNIÃO ESTÁVEL": "Uniao Estavel",
}

def normalizar_estado_civil(val):
    """Padroniza o estado civil e corrige erros comuns de digitação como Divolciado."""
    if pd.isna(val) or not str(val).strip():
        return ""
    s = str(val).strip()
    s_norm = _remover_acentos(s.upper())
    s_norm_corrigido = s_norm.replace("DIVOL", "DIVOR")
    if s_norm_corrigido in ESTADO_CIVIL_MAP:
        return ESTADO_CIVIL_MAP[s_norm_corrigido]
    if s_norm in ESTADO_CIVIL_MAP:
        return ESTADO_CIVIL_MAP[s_norm]
    return re.sub(r'(?i)divol', 'Divor', s)

def _montar_linha_dict(r, admissao, modo_uap=False):
    unidade = montar_unidade(r.get("FILIAL"), r.get("LOCAL"))
    departamento = normalizar_depto(r.get("CENTRO DE CUSTO"))
    subunidade = normalizar_subunidade(r.get("LOCAL", ""))
    telefone = limpar_telefone(r.get("CELULAR"))
    tel_val = "" if modo_uap else telefone
    cel_val = telefone if modo_uap else ""
    return {
        "nome": r.get("NOME", ""),
        "e-mail": "",
        "admissao": admissao,
        "data de nascimento": str(r.get("DATA DE NASCIMENTO", "")).strip(),
        "matricula": str(r.get("MATRICULA", "")).strip() if not pd.isna(r.get("MATRICULA")) else "",
        "ramal": "",
        "cpf": str(r.get("CPF", "")).strip(),
        "telefone": tel_val,
        "residencial": "",
        "celular": cel_val,
        "unidade": unidade,
        "departamento": departamento,
        "subunidade": subunidade,
        "cargo": padronizar_cargo(r.get("CARGO", "")),
        "email superior": str(r.get("E-MAIL DO SUPERIOR", "")).strip() if not pd.isna(r.get("E-MAIL DO SUPERIOR")) else "",
        "email assessor": "",
        "sexo": str(r.get("SEXO", "")).strip(),
        "estado civil": normalizar_estado_civil(r.get("ESTADO CIVIL", "")),
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


def _imprimir_relatorio(df_filtrado, out_path):
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

    print("")
    print(f"Arquivo gerado: {out_path}")
    print("─" * 55)
    print(f"Total de funcionarios: {total}")
    if cargos_encontrados:
        for cargo_ref, nomes in cargos_encontrados:
            for nome in nomes:
                print(f"Funcionarios em cargos especificos: {len(nomes)} - {nome} ({cargo_ref})")
    else:
        print("Funcionarios em cargos especificos: 0")
    if nomes_sem_email:
        for nome in nomes_sem_email:
            print(f"Funcionarios sem e-mail de superior cadastrado: {len(nomes_sem_email)} - {nome}")
    else:
        print("Funcionarios sem e-mail de superior cadastrado: 0")
    print(f"Total de funcionarios cadastrados: {total_cadastrados}")
    print("─" * 55)


def processar(xls_path, data_alvo=None):
    with arquivo_local_seguro(xls_path) as caminho_seguro:
        return _processar_interno(caminho_seguro, data_alvo)


def _processar_interno(xls_path, data_alvo=None):
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
    rows = [_montar_linha_dict(r, alvo.strftime("%d/%m/%Y"), modo_uap=False) for _, r in df_filtrado.iterrows()]
    df_out = pd.DataFrame(rows).map(remover_pontuacao)

    # Nome do arquivo de saída com a data
    data_str = alvo.strftime("%d%m%Y")
    data_pasta = alvo.strftime("%d-%m")
    base_dir = os.path.dirname(os.path.abspath(xls_path))
    # Se o arquivo vem de pasta somente leitura ou servidor remoto, salva localmente
    if base_dir.startswith("/mnt/user-data/uploads") or base_dir.startswith("\\\\") or base_dir.startswith("//"):
        base_dir = _obter_dir_base()
    out_dir = os.path.join(base_dir, data_pasta)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"importar-usuarios-{data_str}.csv")
    df_out.to_csv(out_path, index=False, encoding="latin1")

    _imprimir_relatorio(df_filtrado, out_path)
    return out_path


FILIAIS_UAP = {
    "UMUARAMA ADMINISTRACAO E PARTICIPACAO AG",
    "UMUARAMA ADMINISTRACAO E PARTICIPACAO LT",
}


def processar_uap(xls_path):
    """Gera importacao completa de TODOS os colaboradores do FPRE111,
    sem filtro de data de admissao nem de filial.
    Formato de saida compativel com import CSV da Intranet.
    """
    with arquivo_local_seguro(xls_path) as caminho_seguro:
        return _processar_uap_interno(caminho_seguro)


def _processar_uap_interno(xls_path):
    """Logica interna — sempre recebe caminho local seguro."""
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
    else:
        df = pd.read_excel(xls_path, header=0, dtype=str)

    df.columns = df.columns.str.strip()

    # Import Completo: usa TODOS os colaboradores (sem filtro de filial nem data)
    df_filtrado = df.copy()
    print(f"Total de colaboradores encontrados no arquivo: {len(df_filtrado)}")

    if df_filtrado.empty:
        print("Nenhum colaborador encontrado no arquivo.")
        return None

    # Monta o CSV de saída (modo_uap=True coloca telefone no campo celular)
    rows = [_montar_linha_dict(r, str(r.get("DATA DE ADMISSÃO", "")).strip(), modo_uap=True) for _, r in df_filtrado.iterrows()]
    df_out = pd.DataFrame(rows).map(remover_pontuacao)

    # Salva com nome baseado na data atual
    agora = datetime.now()
    data_str = agora.strftime("%d%m%Y")
    data_pasta = agora.strftime("%d-%m")
    base_dir = os.path.dirname(os.path.abspath(xls_path))
    # Se o arquivo vem de pasta somente leitura ou servidor remoto, salva localmente
    if base_dir.startswith("/mnt/user-data/uploads") or base_dir.startswith("\\\\") or base_dir.startswith("//"):
        base_dir = _obter_dir_base()
    out_dir = os.path.join(base_dir, data_pasta)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"import-completo-{data_str}.csv")
    df_out.to_csv(out_path, index=False, encoding="latin1")

    _imprimir_relatorio(df_filtrado, out_path)
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso:")
        print("  python gerar_importar_usuarios.py <arquivo> [dd/mm/aaaa]         # importacao semanal normal")
        print("  python gerar_importar_usuarios.py <arquivo> --uap               # import completo (todos os colaboradores)")
        sys.exit(1)

    xls_path = sys.argv[1]
    if len(sys.argv) > 2 and sys.argv[2] == "--uap":
        processar_uap(xls_path)
    else:
        data_alvo = sys.argv[2] if len(sys.argv) > 2 else None
        processar(xls_path, data_alvo)
