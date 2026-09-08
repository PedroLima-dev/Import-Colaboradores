"""
ImportarUsuariosUmuarama.py
Modos de uso:
  - Interface grafica (duplo clique ou via ExecutarImport.bat)
  - Linha de comando (chamado pelo EXECUTAR_IMPORT_COMPLETO.bat):
      python ImportarUsuariosUmuarama.py --modo-bat --csv <path> --usuario <user> --senha <pass> [--dc <dc>] [--sem-troca]
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import subprocess, threading, sys, os, csv, io, tempfile, unicodedata, chardet, argparse, json, re

# ── Resolucao de caminhos compativel com PyInstaller / Script ─────────────────
def obter_diretorio_base():
    """Retorna o diretorio base da aplicacao.
    Quando empacotado (.exe via PyInstaller), retorna a pasta do executavel.
    Quando executado como script (.py), retorna a pasta do script.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

def obter_caminho_ou_map():
    """Localiza o arquivo ou_map.json:
    1. Ao lado do executavel / script (permite customizacao sem recompilar)
    2. Embutido no bundle do PyInstaller (_MEIPASS)
    3. Fallback no diretorio base
    """
    caminho_externo = os.path.join(obter_diretorio_base(), "ou_map.json")
    if os.path.exists(caminho_externo):
        return caminho_externo
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        caminho_embutido = os.path.join(sys._MEIPASS, "ou_map.json")
        if os.path.exists(caminho_embutido):
            return caminho_embutido
    return caminho_externo

# ── Importacao opcional do modulo de geracao de CSV ───────────────────────────
_DIR_SCRIPTS = obter_diretorio_base()
if _DIR_SCRIPTS not in sys.path:
    sys.path.insert(0, _DIR_SCRIPTS)
try:
    from gerar_importar_usuarios_uap import (
        processar     as gerar_ultimas,
        processar_uap as gerar_completo,
    )
    _GERAR_OK  = True
    _GERAR_ERR = ""
except ImportError as _e:
    _GERAR_OK  = False
    _GERAR_ERR = str(_e)

# ── Servidor DC padrao ────────────────────────────────────────────────────────
DC_DEFAULT_IP    = "10.56.24.10"
DC_DEFAULT_LABEL = "umuarama.local"

# ──────────────────────────────────────────────────────────────────────────────
# MAPEAMENTO DE OUs
# ──────────────────────────────────────────────────────────────────────────────
BASE = "OU=GRUPO UMUARAMA,DC=umuarama,DC=local"
CONC = f"OU=CONCESSIONARIAS,{BASE}"

OU_MAP = {
    ("TOYOTA", "ARAGUAINA"):    f"OU=ARAGUAINA,OU=TOYOTA,{CONC}",
    ("TOYOTA", "BACABAL"):      f"OU=BACABAL,OU=TOYOTA,{CONC}",
    ("TOYOTA", "BALSAS"):       f"OU=BALSAS,OU=TOYOTA,{CONC}",
    ("TOYOTA", "CALDAS NOVAS"): f"OU=CALDAS NOVAS,OU=TOYOTA,{CONC}",
    ("TOYOTA", "CATALAO"):      f"OU=CATALAO,OU=TOYOTA,{CONC}",
    ("TOYOTA", "IMPERATRIZ"):   f"OU=IMPERATRIZ,OU=TOYOTA,{CONC}",
    ("TOYOTA", "ITUMBIARA"):    f"OU=ITUMBIARA,OU=TOYOTA,{CONC}",
    ("TOYOTA", "JATAI"):        f"OU=JATAI,OU=TOYOTA,{CONC}",
    ("TOYOTA", "MINEIROS"):     f"OU=MINEIROS,OU=TOYOTA,{CONC}",
    ("TOYOTA", "RIO VERDE"):    f"OU=RIO VERDE,OU=TOYOTA,{CONC}",
    ("VOLKSWAGEN", "ARAGUAINA"):  f"OU=ARAGUAINA,OU=VOLKSWAGEN,{CONC}",
    ("VOLKSWAGEN", "CATALAO"):    f"OU=CATALAO,OU=VOLKSWAGEN,{CONC}",
    ("VOLKSWAGEN", "GOIANESIA"):  f"OU=GOIANESIA,OU=VOLKSWAGEN,{CONC}",
    ("VOLKSWAGEN", "GURUPI"):     f"OU=GURUPI,OU=VOLKSWAGEN,{CONC}",
    ("VOLKSWAGEN", "ITUMBIARA"):  f"OU=ITUMBIARA,OU=VOLKSWAGEN,{CONC}",
    ("VOLKSWAGEN", "PALMAS"):     f"OU=PALMAS,OU=VOLKSWAGEN,{CONC}",
    ("VOLKSWAGEN", "PORANGATU"):  f"OU=PORANGATU,OU=VOLKSWAGEN,{CONC}",
    ("VOLKSWAGEN", "URUACU"):     f"OU=GOIANESIA,OU=VOLKSWAGEN,{CONC}",
    ("FIAT", "ARAGUAINA"):   f"OU=ARAGUAINA,OU=FIAT,{CONC}",
    ("FIAT", "PARAUAPEBAS"): f"OU=PARAUAPEBAS,OU=FIAT,{CONC}",
    ("FIAT", "REDENCAO"):    f"OU=REDENCAO,OU=FIAT,{CONC}",
    ("KIA", "IMPERATRIZ"):  f"OU=IMPERATRIZ,OU=KIA,{CONC}",
    ("KIA", "PALMAS"):      f"OU=PALMAS,OU=KIA,{CONC}",
    ("HARLEY", "GOIANIA"):    f"OU=GOIANIA,OU=HARLEY,{CONC}",
    ("HARLEY", "UBERLANDIA"): f"OU=UBERLANDIA,OU=HARLEY,{CONC}",
    ("HD", "GOIANIA"):    f"OU=GOIANIA,OU=HARLEY,{CONC}",
    ("HD", "UBERLANDIA"): f"OU=UBERLANDIA,OU=HARLEY,{CONC}",
    ("KTM", "GOIANIA"):  f"OU=GOIANIA,OU=KTM,{CONC}",
    ("TRIUMPH", "UBERLANDIA"):  f"OU=UBERLANDIA,OU=TRIUMPH,{CONC}",
    ("CITROEN-PEUGEOT", "RIO VERDE"):  f"OU=RIO VERDE,OU=CITROEN-PEUGEOT,{CONC}",
    ("CRT", "ARAGUAINA"):   f"OU=ARAGUAINA,OU=CRT,{CONC}",
    ("CRT", "IMPERATRIZ"):  f"OU=IMPERATRIZ,OU=CRT,{CONC}",
    ("CRT", "ITUMBIARA"):   f"OU=ITUMBIARA,OU=CRT,{CONC}",
    ("SEMINOVOS", "GO"):        f"OU=GO,OU=SEMINOVOS,{CONC}",
    ("SEMINOVOS", "MA"):        f"OU=MA,OU=SEMINOVOS,{CONC}",
    ("SEMINOVOS", "PA"):        f"OU=PA,OU=SEMINOVOS,{CONC}",
    ("SEMINOVOS", "TO"):        f"OU=TO,OU=SEMINOVOS,{CONC}",
    ("SEMINOVOS", "RIO VERDE"): f"OU=RIO VERDE,OU=SEMINOVOS,{CONC}",
    ("CORRETORA DE SEGUROS", ""): f"OU=CORRETORA DE SEGUROS,{CONC}",
    ("HOLDING UAC", "TI"):                        f"OU=TI,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "INFORMATICA"):               f"OU=TI,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "TECNOLOGIA"):                f"OU=TI,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "RECURSOS HUMANOS"):          f"OU=RECURSOS HUMANOS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "RH"):                        f"OU=RECURSOS HUMANOS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "DP"):                        f"OU=DP,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "DEPARTAMENTO PESSOAL"):      f"OU=DP,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "FINANCEIRO"):                f"OU=FINANCEIRO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "FINANCIAMENTO"):             f"OU=FINANCIAMENTO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "FINANCIAMENTOS"):            f"OU=FINANCIAMENTO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "F&I"):                       f"OU=FINANCIAMENTO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "MARKETING"):                 f"OU=MARKETING,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "MKT"):                       f"OU=MARKETING,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "PLANEJAMENTO"):              f"OU=PLANEJAMENTO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "COMPRAS"):                   f"OU=COMPRAS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "COMPRAS - OBRAS E INFRAESTRUTURA"): f"OU=COMPRAS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "OBRAS"):                     f"OU=COMPRAS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "INFRAESTRUTURA"):            f"OU=COMPRAS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "DIRETORIA"):                 f"OU=DIRETORIA,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "SEGURANCA DO TRABALHO"):     f"OU=SEGURANCA DO TRABALHO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "SESMT"):                     f"OU=SEGURANCA DO TRABALHO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "SEGURANCA"):                 f"OU=SEGURANCA DO TRABALHO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "REPASSE / SEMINOVOS"):       f"OU=REPASSE / SEMINOVOS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "REPASSE"):                   f"OU=REPASSE / SEMINOVOS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "VEICULOS USADOS"):           f"OU=REPASSE / SEMINOVOS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "SEMINOVOS"):                 f"OU=REPASSE / SEMINOVOS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "VENDAS DIRETAS"):            f"OU=VENDAS DIRETAS / MOBILIDADE / LICITACAO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "VENDA DIRETA"):              f"OU=VENDAS DIRETAS / MOBILIDADE / LICITACAO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "VENDAS DIRETAS / MOBILIDADE / LICITACAO"): f"OU=VENDAS DIRETAS / MOBILIDADE / LICITACAO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "VENDA DIRETA / MOBILIDADE / LICITACAO"):   f"OU=VENDAS DIRETAS / MOBILIDADE / LICITACAO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "MOBILIDADE"):                f"OU=VENDAS DIRETAS / MOBILIDADE / LICITACAO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "LICITACAO"):                 f"OU=VENDAS DIRETAS / MOBILIDADE / LICITACAO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "CENTRAL DE RELACIONAMENTO"): f"OU=CENTRAL DE RELACIONAMENTO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "CRC"):                       f"OU=CENTRAL DE RELACIONAMENTO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "RELACIONAMENTO"):            f"OU=CENTRAL DE RELACIONAMENTO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "CALL CENTER"):               f"OU=CENTRAL DE RELACIONAMENTO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "ADMINISTRACAO"):             f"OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "ADMINISTRATIVO"):            f"OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "ADM"):                       f"OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", ""):                          f"OU=HOLDING UAC,{BASE}",

    ("HOLDING UAP", "FINANCEIRO"):    f"OU=FINANCEIRO,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "CONTABIL"):      f"OU=CONTABIL,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "CONTABILIDADE"): f"OU=CONTABIL,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "CONTROLADORIA"): f"OU=CONTROLADORIA,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "FISCAL"):        f"OU=FISCAL,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "TRIBUTARIO"):    f"OU=FISCAL,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "AUDITORIA"):     f"OU=AUDITORIA,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "MEIO AMBIENTE"): f"OU=MEIO AMBIENTE,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "AMBIENTAL"):     f"OU=MEIO AMBIENTE,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "ADMINISTRACAO"): f"OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "ADMINISTRATIVO"): f"OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "ADM"):           f"OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", ""):              f"OU=HOLDING UAP,{BASE}",
}

# ── Carregamento dinamico de OUs via JSON externo ─────────────────────────────
def carregar_ou_map():
    """Carrega mapeamento de OUs do arquivo JSON externo.
    Se o JSON nao existir, retorna o dicionario hardcoded OU_MAP acima.
    """
    caminho_json = obter_caminho_ou_map()
    if not os.path.exists(caminho_json):
        print(f"[OU-MAP] Arquivo {caminho_json} nao encontrado, usando mapeamento hardcoded.")
        return dict(OU_MAP)  # copia do hardcoded

    try:
        with open(caminho_json, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception as e:
        print(f"[OU-MAP] Erro ao ler {caminho_json}: {e}  — usando mapeamento hardcoded.")
        return dict(OU_MAP)

    mapa = {}

    # Concessionarias: padrao OU=CIDADE,OU=MARCA,{CONC}
    for marca, cidades in cfg.get("concessionarias", {}).items():
        for cidade in cidades:
            if cidade:
                mapa[(marca, cidade)] = f"OU={cidade},OU={marca},{CONC}"
            else:
                mapa[(marca, "")] = f"OU={marca},{CONC}"

    # Overrides (casos especiais como VOLKSWAGEN|URUACU → OU de GOIANESIA)
    for chave, valor in cfg.get("overrides", {}).items():
        partes = chave.split("|", 1)
        if len(partes) == 2:
            mapa[(partes[0], partes[1])] = valor

    # Holdings UAC
    for depto, ou_nome in cfg.get("holding_uac", {}).items():
        if ou_nome is True:
            mapa[("HOLDING UAC", depto)] = f"OU={depto},OU=HOLDING UAC,{BASE}"
        elif ou_nome == "":
            mapa[("HOLDING UAC", depto)] = f"OU=HOLDING UAC,{BASE}"
        else:
            mapa[("HOLDING UAC", depto)] = f"OU={ou_nome},OU=HOLDING UAC,{BASE}"

    # Holdings UAP
    for depto, ou_nome in cfg.get("holding_uap", {}).items():
        if ou_nome is True:
            mapa[("HOLDING UAP", depto)] = f"OU={depto},OU=HOLDING UAP,{BASE}"
        elif ou_nome == "":
            mapa[("HOLDING UAP", depto)] = f"OU=HOLDING UAP,{BASE}"
        else:
            mapa[("HOLDING UAP", depto)] = f"OU={ou_nome},OU=HOLDING UAP,{BASE}"

    # Carrega aliases do JSON para atualizar os dicionarios globais
    for alias, real in cfg.get("alias_marca", {}).items():
        ALIAS_MARCA[norm(alias)] = real
    for alias, real in cfg.get("alias_cidade", {}).items():
        ALIAS_CIDADE[norm(alias)] = real

    print(f"[OU-MAP] Carregado de {caminho_json}: {len(mapa)} OUs mapeadas.")
    return mapa


def _gerar_ou_fallback(marca, cidade):
    """Gera OU automaticamente no padrao de concessionaria.
    Retorna a OU gerada ou None se nao for aplicavel.
    """
    if marca in ("HOLDING UAC", "HOLDING UAP") or not cidade:
        return None
    return f"OU={cidade},OU={marca},{CONC}"


def _salvar_ou_no_json(marca, cidade):
    """Salva a nova combinacao marca+cidade no ou_map.json para uso futuro.
    Grava no diretorio base da aplicacao para garantir persistencia entre execucoes.
    """
    try:
        caminho_leitura = obter_caminho_ou_map()
        caminho_destino = os.path.join(obter_diretorio_base(), "ou_map.json")
        cfg = {}
        if os.path.exists(caminho_leitura):
            with open(caminho_leitura, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        conc = cfg.get("concessionarias", {})
        if marca not in conc:
            conc[marca] = []
        if cidade not in conc[marca]:
            conc[marca].append(cidade)
            cfg["concessionarias"] = conc
            with open(caminho_destino, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            print(f"[OU-MAP] Salvo automaticamente no JSON: {marca} > {cidade} em {caminho_destino}")
    except Exception as e:
        print(f"[OU-MAP] Aviso: nao foi possivel salvar no JSON: {e}")


ALIAS_MARCA = {
    "VOLKS": "VOLKSWAGEN", "VW": "VOLKSWAGEN",
    "HARLEY-DAVIDSON": "HARLEY", "HD": "HARLEY",
    "CITROEN": "CITROEN-PEUGEOT", "PEUGEOT": "CITROEN-PEUGEOT",
    "CORRETORA": "CORRETORA DE SEGUROS", "CORRETORA DE SEGUROS": "CORRETORA DE SEGUROS",
    "JEEP": "JEEP E RAM", "RAM": "JEEP E RAM",
}

ALIAS_CIDADE = {
    "MINEIRO":  "MINEIROS",
    "URUACU":   "URUACU",
    "URUAÇU":   "URUACU",
}

def strip_accents(s):
    return unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode()

def norm(s):
    return strip_accents(s.strip()).upper()

def pertence_grupo_vpn(unidade, depto):
    u_norm = norm(unidade)
    if u_norm in ("HOLDING UAC", "HOLDING UAP") or u_norm.startswith("HOLDING UAC") or u_norm.startswith("HOLDING UAP"):
        return True
    d_norm = norm(depto)
    for termo in ("POS VENDAS", "POS-VENDAS", "VEICULOS NOVOS", "ADMINISTRACAO"):
        if d_norm.startswith(termo):
            return True
    return False

# ── Inicializacao dinamica do OU_MAP ──────────────────────────────────────────
# Carrega do JSON externo se disponivel; caso contrario usa hardcoded acima.
OU_MAP = carregar_ou_map()

def parse_unidade(unidade_raw, departamento_raw):
    u = unidade_raw.strip()
    d = departamento_raw.strip()
    u_norm = norm(u)

    # 1. Trata Holding UAC e UAP
    if u_norm.startswith("HOLDING UAC") or u_norm.startswith("HOLDING UAP"):
        holding = "HOLDING UAC" if u_norm.startswith("HOLDING UAC") else "HOLDING UAP"
        # Se a unidade veio no formato "HOLDING UAC - VENDA DIRETA"
        if " - " in u_norm:
            sub = u_norm.split(" - ", 1)[1].strip()
            depto = sub if not d else d
        else:
            depto = d
        depto_norm = norm(depto)
        return holding, depto_norm

    # 2. Corretora de Seguros (sem distinção de cidade na unidade)
    if u_norm.startswith("CORRETORA"):
        return "CORRETORA DE SEGUROS", ""

    # 3. Concessionárias
    if " - " in u:
        partes = u.split(" - ", 1)
        marca_raw, cidade_raw = partes[0].strip(), partes[1].strip()
    else:
        marca_raw, cidade_raw = u, ""
    marca  = ALIAS_MARCA.get(norm(marca_raw), norm(marca_raw))
    cidade = ALIAS_CIDADE.get(norm(cidade_raw), norm(cidade_raw))
    return marca, cidade

def resolver_ou(marca, depto_ou_cidade):
    """Resolve a OU para a combinacao marca+cidade ou holding+departamento.
    Retorna (ou_dn, is_auto) onde is_auto=True indica OU gerada por fallback.
    """
    # 1. Busca direta exata
    ou = OU_MAP.get((marca, depto_ou_cidade))
    if ou:
        return ou, False

    # 2. Casos de Holding (busca parcial/sinônimos)
    if marca in ("HOLDING UAC", "HOLDING UAP"):
        if depto_ou_cidade:
            for (m, d), val in OU_MAP.items():
                if m == marca and d and (d == depto_ou_cidade or d in depto_ou_cidade or depto_ou_cidade in d):
                    return val, False
        ou_vazio = OU_MAP.get((marca, ""))
        if ou_vazio:
            return ou_vazio, False

    # 3. Corretora de Seguros
    if marca == "CORRETORA DE SEGUROS":
        ou_corr = OU_MAP.get(("CORRETORA DE SEGUROS", ""))
        if ou_corr:
            return ou_corr, False

    # 4. Concessionárias (busca com cidade vazia se aplicável)
    ou_vazio = OU_MAP.get((marca, ""))
    if ou_vazio and not depto_ou_cidade:
        return ou_vazio, False

    # 5. Fallback: tenta gerar automaticamente no padrao concessionaria
    ou_auto = _gerar_ou_fallback(marca, depto_ou_cidade)
    if ou_auto:
        return ou_auto, True
    return None, False

def ou_para_caminho(ou_dn):
    """Converte DN de OU para formato legivel 'GRUPO UMUARAMA > CONCESSIONARIAS > TOYOTA > ITUMBIARA'."""
    partes = []
    for seg in ou_dn.split(","):
        seg = seg.strip()
        if seg.upper().startswith("OU="):
            partes.append(seg[3:].strip())
    partes.reverse()
    return " > ".join(partes) if partes else ou_dn

def gerar_login(nome_completo, marca=""):
    """Gera o login no padrao nome.sobrenome.
    Para FIAT, KIA e CORRETORA, nao aplica restricao de 15 caracteres.
    Para as demais marcas, trunca em 15 caracteres.
    """
    partes = nome_completo.strip().split()
    if not partes:
        return ""
    if len(partes) < 2:
        base = strip_accents(partes[0]).lower()
    else:
        base = f"{strip_accents(partes[0]).lower()}.{strip_accents(partes[-1]).lower()}"

    m_norm = norm(marca)
    # Unidades isentas da restricao de 15 caracteres: FIAT, KIA e CORRETORA
    if m_norm in ("FIAT", "KIA", "CORRETORA", "CORRETORA DE SEGUROS") or any(m_norm.startswith(p) for p in ("FIAT", "KIA", "CORRETORA")):
        return base
    return base[:15]

# ──────────────────────────────────────────────────────────────────────────────
# Parse CSV de admissoes
# ──────────────────────────────────────────────────────────────────────────────
def parse_csv_admissoes(path):
    with open(path, "rb") as f:
        raw = f.read()
    enc = chardet.detect(raw).get("encoding") or "utf-8"
    text = raw.decode(enc, errors="replace")
    primeira = text.splitlines()[0] if text.strip() else ""
    sep = ";" if primeira.count(";") >= primeira.count(",") else ","
    reader = csv.DictReader(io.StringIO(text), delimiter=sep)
    campo = {norm(k): k for k in (reader.fieldnames or [])}

    def get(row, *chaves):
        for c in chaves:
            k = campo.get(norm(c))
            if k and row.get(k, "").strip():
                return row[k].strip()
        return ""

    usuarios, erros = [], []
    for i, row in enumerate(reader, 2):
        nome_completo = get(row, "nome", "nome completo", "colaborador")
        if not nome_completo:
            continue
        unidade     = get(row, "unidade", "loja", "empresa")
        departamento= get(row, "departamento", "depto", "setor")
        cpf         = get(row, "cpf")
        email       = get(row, "e-mail", "email", "mail")

        if not unidade:
            erros.append(f"Linha {i} ({nome_completo}): coluna 'unidade' vazia")
            continue

        marca, cidade = parse_unidade(unidade, departamento)
        ou, is_auto = resolver_ou(marca, cidade)
        if not ou:
            erros.append(
                f"Linha {i} ({nome_completo}): OU nao mapeada "
                f"[unidade='{unidade}' -> marca='{marca}' cidade='{cidade}']"
            )
            continue

        partes = nome_completo.strip().split()
        login     = gerar_login(nome_completo, marca)
        primeiro  = partes[0]
        sobrenome = " ".join(partes[1:]) if len(partes) > 1 else ""
        if not email:
            email = f"{login}@umuarama.local"

        usuarios.append({
            "nome": primeiro, "sobrenome": sobrenome,
            "login": login, "email": email,
            "ou": ou, "cpf": cpf, "unidade": unidade,
            "vpn": pertence_grupo_vpn(unidade, departamento),
            "is_auto_ou": is_auto,
            "marca_auto": marca if is_auto else "",
            "cidade_auto": cidade if is_auto else "",
        })
    return usuarios, erros

# ──────────────────────────────────────────────────────────────────────────────
# Gera PowerShell inline
# ──────────────────────────────────────────────────────────────────────────────
def gerar_ps1(usuarios, servidor_dc, forcar_troca):
    linhas = ""
    for u in usuarios:
        nome_esc = u["nome"].replace('"', "'")
        sob_esc  = u["sobrenome"].replace('"', "'")
        vpn_val  = "$true" if u.get("vpn") else "$false"
        linhas += (
            f'    @{{Nome="{nome_esc}";Sobrenome="{sob_esc}";'
            f'Login="{u["login"]}";Email="{u["email"]}";'
            f'OU="{u["ou"]}";CPF="{u["cpf"]}";VPN={vpn_val}}},\n'
        )
    linhas = linhas.rstrip(",\n")
    dc_p  = f'"{servidor_dc}"' if servidor_dc else '""'
    troca = "$true" if forcar_troca else "$false"
    return f"""
$credUser   = $env:UMU_USER
$credSenha  = $env:UMU_PASS
$ServidorDC = {dc_p}
$ForcarTroca = {troca}
$BaseDN = "DC=umuarama,DC=local"

function New-Entry([string]$path) {{
    $escaped = $path -replace '/', '\\/'
    if ($ServidorDC) {{ $p = "LDAP://$ServidorDC/$escaped" }} else {{ $p = "LDAP://$escaped" }}
    return New-Object System.DirectoryServices.DirectoryEntry($p, $credUser, $credSenha)
}}
function Get-UsuarioAD([string]$sam) {{
    $ldap = if ($ServidorDC) {{ "LDAP://$ServidorDC/$BaseDN" }} else {{ "LDAP://$BaseDN" }}
    $e = New-Object System.DirectoryServices.DirectoryEntry($ldap, $credUser, $credSenha)
    $s = New-Object System.DirectoryServices.DirectorySearcher($e)
    $s.Filter = "(&(objectClass=user)(sAMAccountName=$sam))"
    $s.SearchScope = "Subtree"
    $res = $s.FindOne()
    if ($res -ne $null) {{
        $entry = $res.GetDirectoryEntry()
        $dispName = $entry.Properties["displayName"].Value
        if ($dispName -eq $null) {{ $dispName = $entry.Properties["cn"].Value }}
        if ($dispName -eq $null) {{ $dispName = "" }}
        $uac = $entry.Properties["userAccountControl"].Value
        if ($uac -eq $null) {{ $uac = 0 }}
        $dn = $entry.Properties["distinguishedName"].Value
        return [PSCustomObject]@{{
            Exists = $true
            DisplayName = [string]$dispName
            UAC = [int]$uac
            DN = [string]$dn
            Path = [string]$entry.Path
        }}
    }}
    return $null
}}
$usuarios = @(
{linhas}
)
foreach ($u in $usuarios) {{
    $nomeCompleto = "$($u.Nome) $($u.Sobrenome)".Trim()
    $sam = $u.Login
    $primeiroNome = $u.Nome.Split(" ")[0]
    $SenhaPadrao = "@" + $primeiroNome.Substring(0,1).ToUpper() + $primeiroNome.Substring(1).ToLower() + "2026"
    Write-Output "INICIO|$sam|$nomeCompleto"
    try {{
        $existe = Get-UsuarioAD $sam
        if ($existe -ne $null) {{
            $uac = $existe.UAC
            $is_disabled = (($uac -band 2) -eq 2)
            $existing_name = $existe.DisplayName
            
            $existing_norm = ($existing_name -replace '\\s+', ' ').Trim()
            $import_norm = ($nomeCompleto -replace '\\s+', ' ').Trim()
            $match_nome = ($existing_norm -ieq $import_norm)
            
            if ($match_nome) {{
                if (-not $is_disabled) {{
                    Write-Output "RESULTADO|$sam|OK_EXISTE_ATIVO|$($existe.DN)"
                    continue
                }} else {{
                    try {{
                        $userEntry = New-Object System.DirectoryServices.DirectoryEntry($existe.Path, $credUser, $credSenha)
                        $currUac = [int]$userEntry.Properties["userAccountControl"].Value
                        $newUac = ($currUac -band -bnot 2) -bor 512
                        $userEntry.Properties["userAccountControl"].Value = $newUac
                        $userEntry.CommitChanges()
                        $userEntry.Dispose()
                        Write-Output "RESULTADO|$sam|OK_ATIVADO|$($existe.DN)"
                        continue
                    }} catch {{
                        Write-Output "RESULTADO|$sam|ERRO|Falha ao ativar usuario: $($_.Exception.Message)"
                        continue
                    }}
                }}
            }} else {{
                $status = if ($is_disabled) {{ "INATIVO" }} else {{ "ATIVO" }}
                $ouDN = $existe.DN -replace '^CN=[^,]+,\\s*', ''
                $clean_name = $existing_name -replace '\\|', ' '
                $clean_ou = $ouDN -replace '\\|', ' '
                Write-Output "RESULTADO|$sam|CONFLITO_NOME|$clean_name|$clean_ou|$status"
                continue
            }}
        }}
        $entry = New-Entry $u.OU
        if (-not $entry.Guid) {{ throw "OU nao encontrada: $($u.OU)" }}
        $user = $entry.Children.Add("CN=$nomeCompleto", "user")
        $user.Properties["sAMAccountName"].Value    = $sam
        $user.Properties["userPrincipalName"].Value = $u.Email
        $user.Properties["givenName"].Value         = $u.Nome
        $user.Properties["sn"].Value                = $u.Sobrenome
        $user.Properties["displayName"].Value       = $nomeCompleto
        $user.Properties["mail"].Value              = $u.Email
        $user.Properties["description"].Value       = $u.CPF
        $user.CommitChanges()
        $user.Invoke("SetPassword", $SenhaPadrao)
        $user.Properties["userAccountControl"].Value = 512
        if ($ForcarTroca) {{ $user.Properties["pwdLastSet"].Value = 0 }}
        $user.CommitChanges()

        if ($u.VPN) {{
            try {{
                $grupoDN = "CN=UsuariosVPN,OU=OpenVPN,DC=umuarama,DC=local"
                $ldapGrupo = if ($ServidorDC) {{ "LDAP://$ServidorDC/$grupoDN" }} else {{ "LDAP://$grupoDN" }}
                $grupoEntry = New-Object System.DirectoryServices.DirectoryEntry($ldapGrupo, $credUser, $credSenha)
                $userDN = $user.Properties["distinguishedName"].Value
                if (-not $grupoEntry.Guid) {{ throw "Grupo VPN nao encontrado" }}
                if (-not $grupoEntry.Properties["member"].Contains($userDN)) {{
                    $grupoEntry.Properties["member"].Add($userDN) | Out-Null
                    $grupoEntry.CommitChanges()
                }}
                $grupoEntry.Dispose()
                Write-Output "VPN_OK|$sam"
            }} catch {{
                Write-Output "VPN_ERRO|$sam|$($_.Exception.Message)"
            }}
        }}

        $user.Dispose()
        $entry.Dispose()
        Write-Output "RESULTADO|$sam|OK|$($u.OU)"
    }} catch {{
        Write-Output "RESULTADO|$sam|ERRO|$($_.Exception.Message)"
    }}
}}
Write-Output "FIM"
"""

# ──────────────────────────────────────────────────────────────────────────────
# Modo batch (chamado pelo .bat, sem GUI)
# ──────────────────────────────────────────────────────────────────────────────
def modo_bat(args):
    print(f"\nCSV: {args.csv}")
    print(f"Usuario AD: {args.usuario}")
    print(f"DC: {args.dc or '(automatico)'}\n")

    usuarios, erros = parse_csv_admissoes(args.csv)

    if erros:
        print(f"[AVISO] {len(erros)} linha(s) com problema:")
        for e in erros:
            print(f"  - {e}")
        print()

    # Confirmacao de OUs auto-geradas no terminal
    auto_ous = {}
    for u in usuarios:
        if u.get("is_auto_ou"):
            chave = (u["marca_auto"], u["cidade_auto"])
            auto_ous.setdefault(chave, []).append(u)

    if auto_ous:
        print(f"[ATENCAO] {len(auto_ous)} OU(s) nao mapeada(s) detectada(s):\n")
        recusados = []
        for (marca, cidade), users_auto in auto_ous.items():
            ou_gerada = _gerar_ou_fallback(marca, cidade)
            caminho = ou_para_caminho(ou_gerada) if ou_gerada else "???"
            nomes = ", ".join(f"{u['nome']} {u['sobrenome']}" for u in users_auto)
            print(f"  {marca} > {cidade}")
            print(f"  OU gerada: {caminho}")
            print(f"  Funcionario(s): {nomes}")
            resp = input(f"  Usar essa OU? (S=sim e salvar / N=pular): ").strip().upper()
            if resp in ("S", "SIM", "Y", "YES"):
                print(f"  -> CONFIRMADO. Salvo no mapeamento.\n")
                _salvar_ou_no_json(marca, cidade)
                global OU_MAP
                OU_MAP = carregar_ou_map()
                for u in users_auto:
                    u["is_auto_ou"] = False
            else:
                print(f"  -> RECUSADO. Funcionarios removidos.\n")
                recusados.extend(users_auto)
        if recusados:
            ids_recusados = {id(u) for u in recusados}
            usuarios = [u for u in usuarios if id(u) not in ids_recusados]

    if not usuarios:
        print("[ERRO] Nenhum usuario valido para importar.")
        sys.exit(1)

    print(f"Importando {len(usuarios)} usuario(s)...\n")

    script = gerar_ps1(usuarios, args.dc or "", not args.sem_troca)
    tmp = tempfile.NamedTemporaryFile(suffix=".ps1", delete=False, mode="w", encoding="utf-8")
    tmp.write(script); tmp.close()

    env = os.environ.copy()
    env["UMU_USER"] = args.usuario
    env["UMU_PASS"] = args.senha

    c = {"ok": 0, "erro": 0, "existe": 0}
    registros_dict = {
        id(u): {
            "nome": f"{u['nome']} {u['sobrenome']}".strip(),
            "login": u["login"],
            "status": "PENDENTE",
            "vpn": "SIM" if u.get("vpn") else "NAO",
            "ou": u["ou"],
            "detalhe": "",
        }
        for u in usuarios
    }
    u_atual = None
    try:
        proc = subprocess.Popen(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmp.name],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            env=env, text=True, encoding="utf-8", errors="replace"
        )
        for linha in proc.stdout:
            linha = linha.rstrip()
            if linha.startswith("INICIO|"):
                _, sam, nome = linha.split("|", 2)
                u_atual = next((u for u in usuarios if u["login"] == sam), None)
                print(f"  -> {nome} ({sam})", end="")
            elif linha.startswith("RESULTADO|"):
                partes = linha.split("|")
                sam = partes[1]
                status = partes[2]
                r_item = registros_dict.get(id(u_atual)) if u_atual else None
                if status == "OK":
                    print(f"  [CRIADO]  →  {ou_para_caminho(partes[3])}"); c["ok"] += 1
                    if r_item: r_item.update({"status": "CRIADO", "ou": partes[3]})
                elif status == "OK_EXISTE_ATIVO":
                    print(f"  [JA EXISTE E ATIVO (PULADO)]  →  {ou_para_caminho(partes[3])}"); c["existe"] += 1
                    if r_item: r_item.update({"status": "JA EXISTIA (ATIVO)", "ou": partes[3], "detalhe": "Criacao pulada"})
                elif status == "OK_ATIVADO":
                    print(f"  [ATIVADO]  →  {ou_para_caminho(partes[3])}"); c["ok"] += 1
                    if r_item: r_item.update({"status": "ATIVADO", "ou": partes[3], "detalhe": "Reativado no AD"})
                elif status in ("JA_EXISTE", "CONFLITO_NOME"):
                    print(f"  [JA EXISTE]"); c["existe"] += 1
                    if r_item: r_item.update({"status": "JA EXISTIA"})
                else:
                    detalhe = partes[3] if len(partes) > 3 else ""
                    print(f"  [ERRO] {detalhe}"); c["erro"] += 1
                    if r_item: r_item.update({"status": "ERRO", "detalhe": detalhe})
            elif linha.startswith("VPN_OK|"):
                print(" [VPN OK]", end="")
                if u_atual and id(u_atual) in registros_dict:
                    registros_dict[id(u_atual)]["vpn"] = "SIM"
            elif linha.startswith("VPN_ERRO|"):
                _, sam, erro_msg = linha.split("|", 2)
                print(f" [VPN ERRO: {erro_msg}]", end="")
                if u_atual and id(u_atual) in registros_dict:
                    registros_dict[id(u_atual)]["vpn"] = f"ERRO: {erro_msg}"
            elif linha == "FIM":
                break
            elif linha.strip():
                print(f"  {linha}")
        proc.wait()
    finally:
        os.unlink(tmp.name)

    print(f"\n{'─'*55}")
    print(f"Criados: {c['ok']}  |  Ja existiam: {c['existe']}  |  Erros: {c['erro']}")
    print(f"{'─'*55}\n")

    if c["ok"] > 0 or c["existe"] > 0 or c["erro"] > 0:
        try:
            arq_log = App._salvar_log_csv(args.dc or "(automatico)", list(registros_dict.values()), c, args.csv)
            print(f"Log CSV salvo em: {arq_log}\n")
            try:
                os.startfile(arq_log)
            except Exception:
                pass
        except Exception as e_log:
            print(f"Aviso: nao foi possivel salvar log CSV: {e_log}\n")

    if c["erro"] > 0:
        sys.exit(2)
    sys.exit(0)

# ──────────────────────────────────────────────────────────────────────────────
# Redirecionamento de stdout para o log da GUI
# ──────────────────────────────────────────────────────────────────────────────
class _GuiStdout:
    """Redireciona sys.stdout para o log da GUI durante processamento em thread.
    Cada linha impressa pelo modulo de geracao aparece em tempo real no log."""
    def __init__(self, log_fn):
        self._log = log_fn

    def write(self, msg):
        stripped = msg.rstrip("\n")
        if stripped:
            self._log(stripped, "info")

    def flush(self):
        pass

# ──────────────────────────────────────────────────────────────────────────────
# Interface Grafica  —  proporcao 16:9  (1120 × 630)
# ──────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self, csv_preload=""):
        super().__init__()
        self.title("Importar Usuarios AD - Grupo Umuarama")
        self.geometry("1120x630")
        self.minsize(1120, 630)
        self.resizable(True, True)
        self.configure(bg="#1e1e2e")
        self._preview_realizado = False
        self._build_ui()
        # Se chamado pelo .bat com CSV ja definido, pre-carrega e faz preview
        if csv_preload and os.path.exists(csv_preload):
            self.var_csv.set(csv_preload)
            self.after(300, self._preview)

    # ── Construcao da UI ───────────────────────────────────────────────────────
    def _build_ui(self):
        COR_BG   = "#1e1e2e"
        COR_CARD = "#2a2a3e"
        COR_ACC  = "#7c6af7"
        COR_TXT  = "#cdd6f4"
        COR_SUB  = "#6c7086"

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TLabel",       background=COR_BG,   foreground=COR_TXT,  font=("Segoe UI", 10))
        style.configure("TEntry",       fieldbackground=COR_CARD, foreground=COR_TXT, font=("Segoe UI", 10))
        style.configure("TCheckbutton", background=COR_BG,   foreground=COR_TXT,  font=("Segoe UI", 10))
        style.configure("TFrame",       background=COR_BG)
        style.configure("Card.TFrame",  background=COR_CARD)
        style.configure("Accent.TButton", background=COR_ACC, foreground="white",
                        font=("Segoe UI", 10, "bold"), padding=6)
        style.map("Accent.TButton", background=[("active", "#6a58e0")])
        style.configure("TButton", background=COR_CARD, foreground=COR_TXT,
                        font=("Segoe UI", 10), padding=5)
        style.map("TButton", background=[("active", "#3a3a5e")])
        style.configure("TNotebook",     background=COR_BG, borderwidth=0)
        style.configure("TNotebook.Tab", background=COR_CARD, foreground=COR_SUB,
                        font=("Segoe UI", 10, "bold"), padding=[16, 7])
        style.map("TNotebook.Tab",
                  background=[("selected", "#252538")],
                  foreground=[("selected", COR_ACC)])
        style.configure("TRadiobutton", background=COR_CARD, foreground=COR_TXT,
                        font=("Segoe UI", 10))

        # ── Cabecalho ──────────────────────────────────────────────────────────
        tk.Label(self, text="Importar Usuarios - Grupo Umuarama",
                 bg=COR_BG, fg=COR_ACC, font=("Segoe UI", 14, "bold")).pack(pady=(12, 2))
        tk.Label(self,
                 text="Sistema integrado  ·  Geração de CSV Intranet   e   Importação no Active Directory",
                 bg=COR_BG, fg=COR_SUB, font=("Segoe UI", 9)).pack(pady=(0, 6))

        # ── Notebook ───────────────────────────────────────────────────────────
        self._nb = ttk.Notebook(self)
        self._nb.pack(fill="x", padx=24, pady=0)

        tab1 = ttk.Frame(self._nb, style="TFrame")
        self._nb.add(tab1, text="  GERAR CSV INTRANET  ")
        self._build_tab_gerar(tab1, COR_BG, COR_CARD, COR_ACC, COR_TXT, COR_SUB)

        tab2 = ttk.Frame(self._nb, style="TFrame")
        self._nb.add(tab2, text="  IMPORTAR AD  ")
        self._build_tab_importar(tab2, COR_BG, COR_CARD, COR_ACC, COR_TXT, COR_SUB)

        # ── Log compartilhado (abaixo das abas) ────────────────────────────────
        tk.Label(self, text="Log de execucao", bg=COR_BG, fg=COR_SUB,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=26, pady=(8, 0))
        self.log = scrolledtext.ScrolledText(
            self, bg="#11111b", fg=COR_TXT, font=("Consolas", 9),
            insertbackground=COR_TXT, relief="flat", bd=0)
        self.log.pack(fill="both", expand=True, padx=24, pady=(2, 0))
        self.log.tag_config("ok",     foreground="#a6e3a1")
        self.log.tag_config("erro",   foreground="#f38ba8")
        self.log.tag_config("aviso",  foreground="#f9e2af")
        self.log.tag_config("info",   foreground="#89b4fa")
        self.log.tag_config("titulo", foreground="#cba6f7", font=("Consolas", 9, "bold"))

        # ── Barra de status ────────────────────────────────────────────────────
        self.var_status = tk.StringVar(value="Pronto.")
        tk.Label(self, textvariable=self.var_status, bg="#11111b", fg=COR_SUB,
                 font=("Segoe UI", 8), anchor="w").pack(fill="x", padx=24, pady=(0, 4))

    # ── Aba 1: GERAR CSV INTRANET ──────────────────────────────────────────────
    def _build_tab_gerar(self, parent, COR_BG, COR_CARD, COR_ACC, COR_TXT, COR_SUB):
        card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        card.pack(fill="x", padx=0, pady=6)

        # Arquivo FPRE111
        ttk.Label(card, text="Arquivo FPRE111 (CSV / XLS / XLSX):", background=COR_CARD).grid(
            row=0, column=0, sticky="w", pady=5)
        self.var_fpre = tk.StringVar()
        ttk.Entry(card, textvariable=self.var_fpre, width=64).grid(
            row=0, column=1, padx=8, sticky="ew")
        ttk.Button(card, text="Procurar...", command=self._browse_fpre).grid(row=0, column=2)

        # Modo de geracao
        ttk.Label(card, text="Modo de geracao:", background=COR_CARD).grid(
            row=1, column=0, sticky="w", pady=5)
        self.var_modo = tk.StringVar(value="ultimas")
        mf = ttk.Frame(card, style="Card.TFrame")
        mf.grid(row=1, column=1, sticky="w", padx=8, columnspan=2)
        ttk.Radiobutton(mf, text="Últimas admissões",
                        variable=self.var_modo, value="ultimas",
                        command=self._toggle_data_field).pack(side="left", padx=(0, 28))
        ttk.Radiobutton(mf, text="Import Completo",
                        variable=self.var_modo, value="completo",
                        command=self._toggle_data_field).pack(side="left")

        # Campo de data (so para "Ultimas admissoes")
        self._lbl_data = ttk.Label(card, text="Data de admissao (dd/mm/aaaa):", background=COR_CARD)
        self._lbl_data.grid(row=2, column=0, sticky="w", pady=5)
        self.var_data = tk.StringVar()
        self._ent_data = ttk.Entry(card, textvariable=self.var_data, width=18)
        self._ent_data.grid(row=2, column=1, sticky="w", padx=8)
        self._lbl_hint = ttk.Label(
            card, text="(opcional — usa a data mais recente do arquivo se vazio)",
            background=COR_CARD, foreground=COR_SUB, font=("Segoe UI", 8))
        self._lbl_hint.grid(row=2, column=2, sticky="w")

        card.columnconfigure(1, weight=1)

        # Botoes
        bf = ttk.Frame(parent, style="TFrame")
        bf.pack(pady=6)
        ttk.Button(bf, text="Gerar CSV", style="Accent.TButton",
                   command=self._gerar_csv).pack(side="left", padx=6)
        ttk.Button(bf, text="Limpar Log",
                   command=lambda: self.log.delete("1.0", "end")).pack(side="left", padx=6)

        # Label de resultado (preenchido apos geracao bem-sucedida)
        self.var_lbl_gerado = tk.StringVar(value="")
        tk.Label(parent, textvariable=self.var_lbl_gerado,
                 bg=COR_BG, fg="#a6e3a1", font=("Segoe UI", 8),
                 anchor="w").pack(fill="x", padx=6, pady=(0, 2))

        # Aviso se o modulo nao foi importado
        if not _GERAR_OK:
            tk.Label(parent,
                     text=f"⚠  Modulo gerar_importar_usuarios_uap nao encontrado: {_GERAR_ERR}",
                     bg=COR_BG, fg="#f38ba8", font=("Segoe UI", 8),
                     wraplength=980, justify="left", anchor="w").pack(fill="x", padx=6)

    # ── Aba 2: IMPORTAR AD ─────────────────────────────────────────────────────
    def _build_tab_importar(self, parent, COR_BG, COR_CARD, COR_ACC, COR_TXT, COR_SUB):
        card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        card.pack(fill="x", padx=0, pady=6)

        # CSV de admissoes
        ttk.Label(card, text="Arquivo CSV de admissoes:", background=COR_CARD).grid(
            row=0, column=0, sticky="w", pady=5)
        self.var_csv = tk.StringVar()
        self.var_csv.trace_add("write", lambda *_: setattr(self, "_preview_realizado", False))
        ttk.Entry(card, textvariable=self.var_csv, width=64).grid(
            row=0, column=1, padx=8, sticky="ew")
        ttk.Button(card, text="Procurar...", command=self._browse).grid(row=0, column=2)

        # Usuario AD
        ttk.Label(card, text="Usuario AD (ex: umuarama\\admin):", background=COR_CARD).grid(
            row=1, column=0, sticky="w", pady=5)
        self.var_user = tk.StringVar()
        ttk.Entry(card, textvariable=self.var_user, width=64).grid(
            row=1, column=1, padx=8, sticky="ew")

        # Senha
        ttk.Label(card, text="Senha AD:", background=COR_CARD).grid(
            row=2, column=0, sticky="w", pady=5)
        self.var_pass = tk.StringVar()
        ttk.Entry(card, textvariable=self.var_pass, show="*", width=64).grid(
            row=2, column=1, padx=8, sticky="ew")

        # Servidor DC — pre-preenchido com label legivel; internamente resolve para IP
        ttk.Label(card, text="Servidor DC:", background=COR_CARD).grid(
            row=3, column=0, sticky="w", pady=5)
        self.var_dc = tk.StringVar(value=DC_DEFAULT_LABEL)
        ttk.Entry(card, textvariable=self.var_dc, width=64).grid(
            row=3, column=1, padx=8, sticky="ew")
        ttk.Label(card,
                  text=f"({DC_DEFAULT_LABEL}  →  {DC_DEFAULT_IP}  · editavel)",
                  background=COR_CARD, foreground=COR_SUB,
                  font=("Segoe UI", 8)).grid(row=3, column=2, sticky="w")

        # Forcar troca de senha
        self.var_troca = tk.BooleanVar(value=True)
        ttk.Checkbutton(card, text="Forcar troca de senha no 1 login",
                        variable=self.var_troca).grid(
            row=4, column=1, sticky="w", padx=8, pady=6)
        card.columnconfigure(1, weight=1)

        # Botoes
        bf = ttk.Frame(parent, style="TFrame")
        bf.pack(pady=6)
        ttk.Button(bf, text="Pre-visualizar", command=self._preview).pack(side="left", padx=6)
        ttk.Button(bf, text="Importar", style="Accent.TButton",
                   command=self._iniciar).pack(side="left", padx=6)
        ttk.Button(bf, text="Limpar Log",
                   command=lambda: self.log.delete("1.0", "end")).pack(side="left", padx=6)

    # ── Helpers gerais ─────────────────────────────────────────────────────────
    def _resolver_dc(self):
        """Retorna o host/IP real para conexao LDAP.
        'umuarama.local' e campo vazio resolvem para DC_DEFAULT_IP (10.56.24.10).
        Qualquer outro valor digitado pelo operador e usado diretamente.
        """
        val = self.var_dc.get().strip()
        return DC_DEFAULT_IP if (not val or val == DC_DEFAULT_LABEL) else val

    def _toggle_data_field(self):
        """Exibe ou oculta o campo de data conforme o modo selecionado."""
        if self.var_modo.get() == "ultimas":
            self._lbl_data.grid()
            self._ent_data.grid()
            self._lbl_hint.grid()
        else:
            self._lbl_data.grid_remove()
            self._ent_data.grid_remove()
            self._lbl_hint.grid_remove()

    def _log(self, msg, tag="info"):
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")

    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("CSV", "*.csv"), ("Todos", "*.*")])
        if p:
            self.var_csv.set(p)

    def _browse_fpre(self):
        p = filedialog.askopenfilename(
            filetypes=[("CSV/Excel", "*.csv *.xls *.xlsx"), ("Todos", "*.*")])
        if p:
            self.var_fpre.set(p)

    # ── Aba 1: Geracao de CSV ──────────────────────────────────────────────────
    def _gerar_csv(self):
        if not _GERAR_OK:
            messagebox.showerror("Erro", f"Modulo de geracao nao disponivel:\n{_GERAR_ERR}")
            return
        fpre = self.var_fpre.get().strip()
        if not fpre or not os.path.exists(fpre):
            messagebox.showerror("Erro", "Selecione um arquivo FPRE111 valido.")
            return
        modo      = self.var_modo.get()
        data_alvo = self.var_data.get().strip() if modo == "ultimas" else None

        self.log.delete("1.0", "end")
        self._log("=" * 72, "titulo")
        label_modo = "Ultimas admissoes" if modo == "ultimas" else "Import Completo"
        self._log(f"  GERANDO CSV  [{label_modo}]", "titulo")
        self._log("=" * 72, "titulo")
        self.var_status.set("Gerando CSV...")
        self.var_lbl_gerado.set("")

        threading.Thread(
            target=self._executar_gerar,
            args=(fpre, modo, data_alvo), daemon=True).start()

    def _executar_gerar(self, fpre_path, modo, data_alvo):
        """Executa a geracao de CSV em thread separada, redirecionando stdout para o log."""
        old_out = sys.stdout
        sys.stdout = _GuiStdout(self._log)
        resultado = None
        try:
            if modo == "ultimas":
                resultado = gerar_ultimas(fpre_path, data_alvo or None)
            else:
                resultado = gerar_completo(fpre_path)
        except Exception as e:
            self._log(f"ERRO na geracao: {e}", "erro")
        finally:
            sys.stdout = old_out

        if resultado and os.path.exists(resultado):
            self._log("", "info")
            self._log(f"  CSV gerado: {resultado}", "ok")
            self._log("=" * 72, "titulo")
            self.var_lbl_gerado.set(f"✔  Gerado: {resultado}")
            self.var_csv.set(resultado)          # preenche automaticamente a aba 2
            self.var_status.set(f"CSV gerado: {os.path.basename(resultado)}")
            try:
                os.startfile(resultado)
                self._log("  Arquivo CSV de saida aberto externamente.", "ok")
            except Exception as e_open:
                self._log(f"  Erro ao abrir o CSV gerado externamente: {e_open}", "erro")
        else:
            self._log("  Geracao concluida sem arquivo de saida.", "aviso")
            self.var_status.set("Geracao concluida.")

    # ── Aba 2: Pre-visualizacao e Importacao ───────────────────────────────────
    def _preview(self):
        path = self.var_csv.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Erro", "Selecione um arquivo CSV valido."); return
        usuarios, erros = parse_csv_admissoes(path)
        self.log.delete("1.0", "end")
        self._log("=" * 72, "titulo")
        self._log(f"  PRE-VISUALIZACAO - {len(usuarios)} valido(s)  /  {len(erros)} problema(s)", "titulo")
        self._log("=" * 72, "titulo")

        # Separa usuarios com OU auto-gerada para confirmacao
        auto_ous = {}  # {(marca, cidade): [lista de usuarios]}
        for u in usuarios:
            caminho = ou_para_caminho(u['ou'])
            if u.get("is_auto_ou"):
                self._log(f"  ??  {u['nome']} {u['sobrenome']:<25}  login: {u['login']:<18}  {caminho}  [OU AUTO-GERADA]", "aviso")
                chave = (u["marca_auto"], u["cidade_auto"])
                auto_ous.setdefault(chave, []).append(u)
            else:
                self._log(f"  OK  {u['nome']} {u['sobrenome']:<25}  login: {u['login']:<18}  {caminho}", "ok")
        if erros:
            self._log("")
            for e in erros:
                self._log(f"  XX  {e}", "erro")

        # Se ha OUs auto-geradas, pede confirmacao ao operador
        if auto_ous:
            self._log("")
            self._log(f"  {'─'*68}", "aviso")
            self._log(f"  {len(auto_ous)} OU(s) nao mapeada(s) detectada(s) - CONFIRMACAO NECESSARIA", "aviso")
            self._log(f"  {'─'*68}", "aviso")

            recusados = []
            for (marca, cidade), users_auto in auto_ous.items():
                ou_gerada = _gerar_ou_fallback(marca, cidade)
                caminho_gerado = ou_para_caminho(ou_gerada) if ou_gerada else "???"
                nomes = ", ".join(f"{u['nome']} {u['sobrenome']}" for u in users_auto)

                confirmado = messagebox.askyesno(
                    "OU nao mapeada — Confirmar?",
                    f"A combinacao {marca} > {cidade} nao esta no mapeamento.\n\n"
                    f"OU gerada automaticamente:\n"
                    f"  {caminho_gerado}\n\n"
                    f"Funcionario(s): {nomes}\n\n"
                    f"Essa OU existe no AD?\n"
                    f"SIM = usar essa OU e salvar no mapeamento\n"
                    f"NAO = remover esses funcionarios da importacao",
                )

                if confirmado:
                    self._log(f"  ✔  CONFIRMADO: {marca} > {cidade}  →  {caminho_gerado}", "ok")
                    # Salva no JSON para uso futuro
                    _salvar_ou_no_json(marca, cidade)
                    # Recarrega o OU_MAP global
                    global OU_MAP
                    OU_MAP = carregar_ou_map()
                    # Marca como nao-auto agora que foi confirmado
                    for u in users_auto:
                        u["is_auto_ou"] = False
                else:
                    self._log(f"  ✘  RECUSADO: {marca} > {cidade}  — funcionarios removidos", "erro")
                    recusados.extend(users_auto)

            # Remove usuarios recusados da lista
            if recusados:
                ids_recusados = {id(u) for u in recusados}
                usuarios[:] = [u for u in usuarios if id(u) not in ids_recusados]
                self._log("")
                self._log(f"  {len(recusados)} funcionario(s) removido(s) da importacao.", "erro")

        self._preview_realizado = True
        # Armazena a lista de usuarios validados para uso na importacao
        self._usuarios_preview = usuarios

    def _iniciar(self):
        if not getattr(self, "_preview_realizado", False):
            messagebox.showerror("Erro", "E obrigatorio realizar a Pre-visualizacao antes de importar para o AD.")
            return
        path  = self.var_csv.get().strip()
        user  = self.var_user.get().strip()
        senha = self.var_pass.get()
        dc    = self._resolver_dc()
        if not path or not os.path.exists(path):
            messagebox.showerror("Erro", "Selecione um arquivo CSV valido."); return
        if not user or not senha:
            messagebox.showerror("Erro", "Informe usuario e senha AD."); return

        usuarios, erros = parse_csv_admissoes(path)
        if erros:
            msg = "\n".join(erros[:12]) + ("\n..." if len(erros) > 12 else "")
            if not messagebox.askyesno("Avisos",
                f"{len(erros)} linha(s) com problema:\n\n{msg}\n\n"
                f"Continuar com os {len(usuarios)} registro(s) validos?"):
                return
        if not usuarios:
            messagebox.showinfo("Vazio", "Nenhum usuario valido para importar."); return

        self.log.delete("1.0", "end")
        self._log("=" * 72, "titulo")
        self._log(f"  IMPORTANDO - {len(usuarios)} usuario(s)  [DC: {dc}]", "titulo")
        self._log("=" * 72, "titulo")
        self.var_status.set("Importando...")
        threading.Thread(target=self._executar,
                         args=(usuarios, user, senha, dc), daemon=True).start()

    # ── Dialogo de login duplicado — com validacao e retry inline ─────────────
    def _dialog_novo_login(self, u_data, user, senha, dc, c, r_item=None):
        """Abre dialogo modal, valida o login em tempo real e tenta criar o usuario
        inline (sem fechar o dialogo). Rexibe mensagem de erro se o novo login
        tambem existir no AD. So fecha quando: OK criado, operador pulou (0 / X).

        Retorna o status final: 'OK', 'PULADO' ou 'ERRO'.
        """
        _SAM_REGEX = re.compile(r'^[a-zA-Z0-9._\-]+$')
        _SAM_MAX   = 20

        def _validar_fmt(login):
            if login == "0":
                return None
            if not login:
                return "O login não pode estar vazio."
            if len(login) > _SAM_MAX:
                return f"Login excede {_SAM_MAX} caracteres (limite do AD)."
            if not _SAM_REGEX.match(login):
                return "Caracteres inválidos. Use apenas letras, números, ponto, hífen ou underline."
            return None

        event      = threading.Event()
        resultado  = ["PULADO"]   # default

        COR_BG   = "#1e1e2e"
        COR_CARD = "#2a2a3e"
        COR_ACC  = "#7c6af7"
        COR_TXT  = "#cdd6f4"
        COR_SUB  = "#6c7086"
        COR_AVS  = "#f9e2af"

        def _mostrar():
            dlg = tk.Toplevel(self)
            dlg.title("Usuário duplicado")
            dlg.configure(bg=COR_BG)
            dlg.minsize(520, 420)
            dlg.resizable(True, True)

            # Centraliza a janela modal sobre a janela principal
            self.update_idletasks()
            largura_dlg = 540
            altura_dlg = 450
            pos_x = self.winfo_x() + max(0, (self.winfo_width() - largura_dlg) // 2)
            pos_y = self.winfo_y() + max(0, (self.winfo_height() - altura_dlg) // 2)
            dlg.geometry(f"{largura_dlg}x{altura_dlg}+{pos_x}+{pos_y}")

            dlg.grab_set()
            dlg.focus_force()

            nome_completo   = f"{u_data['nome']} {u_data['sobrenome']}"

            content = tk.Frame(dlg, bg=COR_BG)
            content.pack(fill="both", expand=True, padx=20, pady=16)

            tk.Label(content, text="\u26a0  Usuário duplicado", bg=COR_BG, fg=COR_AVS,
                     font=("Segoe UI", 12, "bold")).pack(pady=(0, 4))
            tk.Label(content, text=nome_completo, bg=COR_BG, fg=COR_TXT,
                     font=("Segoe UI", 11, "bold")).pack(pady=(0, 2))
            tk.Label(content, text=f"Login em conflito:  {u_data['login']}",
                     bg=COR_BG, fg=COR_SUB, font=("Segoe UI", 9)).pack(pady=(0, 8))

            conflict_name   = u_data.get("conflict_name")
            conflict_ou     = u_data.get("conflict_ou")
            conflict_status = u_data.get("conflict_status")
            if conflict_name:
                card_conflito = tk.Frame(content, bg="#2d2238", padx=12, pady=8,
                                         highlightbackground="#583c66", highlightthickness=1)
                card_conflito.pack(fill="x", pady=(0, 10))
                tk.Label(card_conflito, text="Já cadastrado no Active Directory:",
                         bg="#2d2238", fg="#f38ba8", font=("Segoe UI", 8, "bold")).pack(anchor="w")
                txt_conflict = f"{conflict_name}\n({conflict_ou})  ·  Status: {conflict_status}"
                tk.Label(card_conflito, text=txt_conflict, bg="#2d2238", fg="#cdd6f4",
                         font=("Segoe UI", 9), wraplength=480, justify="left").pack(anchor="w", pady=(2, 0))

            tk.Label(content, text="Informe um novo login  ou digite  0  para pular:",
                     bg=COR_BG, fg=COR_TXT, font=("Segoe UI", 9)).pack(pady=(0, 6))

            # ── Campo + contador ───────────────────────────────────────────────
            row_entry = tk.Frame(content, bg=COR_BG)
            row_entry.pack(pady=(0, 2))
            var_entrada = tk.StringVar()
            entry = tk.Entry(row_entry, textvariable=var_entrada, bg=COR_CARD, fg=COR_TXT,
                             insertbackground=COR_TXT, font=("Segoe UI", 11),
                             relief="flat", bd=6, width=24)
            entry.pack(side="left")
            var_contador = tk.StringVar(value=f"0/{_SAM_MAX}")
            lbl_contador = tk.Label(row_entry, textvariable=var_contador,
                                    bg=COR_BG, fg=COR_SUB, font=("Segoe UI", 9), width=6)
            lbl_contador.pack(side="left", padx=(6, 0))

            # ── Label de status (erro de formato OU resultado do AD) ───────────
            var_msg  = tk.StringVar(value="")
            lbl_msg  = tk.Label(content, textvariable=var_msg,
                                bg=COR_BG, font=("Segoe UI", 8),
                                wraplength=480, justify="center")
            lbl_msg.pack(pady=(4, 8))

            def _set_msg(txt, cor):
                var_msg.set(txt)
                lbl_msg.configure(fg=cor)

            # ── Botoes ─────────────────────────────────────────────────────────
            bf = tk.Frame(content, bg=COR_BG)
            bf.pack(pady=(4, 10))
            btn_ok = tk.Button(bf, text="Confirmar  [Enter]", bg=COR_ACC, fg="white",
                               activebackground="#6a58e0", activeforeground="white",
                               font=("Segoe UI", 10, "bold"), relief="flat", padx=16, pady=5, cursor="hand2")
            btn_ok.pack(side="left", padx=8)
            btn_pular = tk.Button(bf, text="Pular (0)  [Esc]", bg=COR_CARD, fg=COR_SUB,
                                  activebackground="#3a3a5e", activeforeground=COR_TXT,
                                  font=("Segoe UI", 10), relief="flat", padx=16, pady=5, cursor="hand2")
            btn_pular.pack(side="left", padx=8)

            # ── Validacao de formato em tempo real ────────────────────────────
            def _on_change(*_):
                val = var_entrada.get().strip()
                n   = len(val)
                cor_cnt = "#f38ba8" if n > _SAM_MAX else COR_SUB
                var_contador.set(f"{n}/{_SAM_MAX}")
                lbl_contador.configure(fg=cor_cnt)
                if val == "0":
                    _set_msg("", COR_SUB)
                    btn_ok.configure(state="normal", bg=COR_ACC, text="Pular (confirmar 0)")
                    return
                erro_fmt = _validar_fmt(val)
                if erro_fmt:
                    _set_msg(erro_fmt, "#f38ba8")
                    btn_ok.configure(state="disabled", bg="#44415a", text="Confirmar  [Enter]")
                else:
                    _set_msg("", COR_SUB)
                    btn_ok.configure(state="normal", bg=COR_ACC, text="Confirmar  [Enter]")

            var_entrada.trace_add("write", _on_change)
            _on_change()

            # ── Confirmar: tenta criar inline sem fechar o dialogo ────────────
            def _tentar_criar(_tk_ev=None):
                if btn_ok.cget("state") == "disabled":
                    return

                val = var_entrada.get().strip()
                if val == "0":
                    resultado[0] = "PULADO"
                    self._log(f"     Usuario pulado pelo operador: {nome_completo} ({u_data['login']})", "aviso")
                    if r_item: r_item.update({"status": "PULADO", "detalhe": "Pulado pelo operador"})
                    dlg.destroy()
                    return

                erro_fmt = _validar_fmt(val)
                if erro_fmt:
                    _set_msg(erro_fmt, "#f38ba8")
                    return

                # Desabilita controles durante a tentativa
                btn_ok.configure(state="disabled", text="Aguardando AD...")
                btn_pular.configure(state="disabled")
                entry.configure(state="disabled")
                _set_msg("Consultando o Active Directory...", COR_SUB)
                dlg.update_idletasks()

                def _executar_ps1():
                    u_retry = dict(u_data)
                    u_retry["login"] = val
                    u_retry["email"] = f"{val}@umuarama.local"
                    status = self._criar_usuario_individual(u_retry, user, senha, dc, c, r_item=r_item)

                    # Atualiza a UI de volta no main thread
                    def _atualizar():
                        if status == "OK":
                            resultado[0] = "OK"
                            dlg.destroy()   # log ja feito por _criar_usuario_individual
                        elif status == "JA_EXISTE":
                            # Limpa o campo e exibe erro — dialogo permanece aberto
                            var_entrada.set("")
                            entry.configure(state="normal")
                            btn_pular.configure(state="normal")
                            _set_msg(
                                f"\u26a0  '{val}' ja existe no AD. Informe um login diferente.",
                                "#f38ba8"
                            )
                            btn_ok.configure(state="disabled", bg="#44415a", text="Confirmar  [Enter]")
                            entry.focus_set()
                        else:  # ERRO
                            resultado[0] = "ERRO"
                            dlg.destroy()

                    self.after(0, _atualizar)

                threading.Thread(target=_executar_ps1, daemon=True).start()

            entry.bind("<Return>", _tentar_criar)
            entry.bind("<KP_Enter>", _tentar_criar)
            dlg.bind("<Return>", _tentar_criar)
            dlg.bind("<KP_Enter>", _tentar_criar)

            def _ao_fechar():
                resultado[0] = "PULADO"
                self._log(f"     Usuario pulado pelo operador: {nome_completo} ({u_data['login']})", "aviso")
                if r_item: r_item.update({"status": "PULADO", "detalhe": "Pulado pelo operador"})
                dlg.destroy()

            dlg.bind("<Escape>", lambda _e: _ao_fechar())
            btn_ok.configure(command=_tentar_criar)
            btn_pular.configure(command=_ao_fechar)
            dlg.protocol("WM_DELETE_WINDOW", _ao_fechar)
            entry.focus_set()

            def _aguardar():
                if dlg.winfo_exists():
                    self.after(50, _aguardar)
                else:
                    event.set()

            self.after(50, _aguardar)

        self.after(0, _mostrar)
        event.wait()
        return resultado[0]

    def _criar_usuario_individual(self, u_data, user, senha, dc, c, r_item=None):
        """Roda um mini PS1 para criar um unico usuario (usado apos conflito de login).
        Retorna o status obtido do AD: 'OK', 'JA_EXISTE' ou 'ERRO'."""
        script = gerar_ps1([u_data], dc, self.var_troca.get())
        tmp = tempfile.NamedTemporaryFile(suffix=".ps1", delete=False, mode="w", encoding="utf-8")
        tmp.write(script); tmp.close()
        env = os.environ.copy()
        env["UMU_USER"] = user
        env["UMU_PASS"] = senha
        _ultimo_status = "ERRO"   # default caso o PS1 nao retorne RESULTADO
        vpn_info = "NAO"          # default vpn status
        try:
            proc = subprocess.Popen(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmp.name],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                env=env, text=True, encoding="utf-8", errors="replace"
            )
            for linha in proc.stdout:
                linha = linha.rstrip()
                if linha.startswith("RESULTADO|"):
                    _, sam, status, detalhe = linha.split("|", 3)
                    _ultimo_status = status
                    if status in ("OK", "OK_EXISTE_ATIVO", "OK_ATIVADO"):
                        _ultimo_status = "OK"
                        caminho = ou_para_caminho(detalhe)
                        if status == "OK_EXISTE_ATIVO":
                            self._log(f"     ✔  JA EXISTIA E ATIVO  (novo login: {sam})  →  {caminho}", "ok")
                        elif status == "OK_ATIVADO":
                            self._log(f"     ✔  ATIVADO  (novo login: {sam})  →  {caminho}", "ok")
                        else:
                            self._log(f"     ✔  CRIADO  (novo login: {sam})  →  {caminho}", "ok")
                        
                        if vpn_info == "SIM":
                            self._log(f"     → Adicionado ao grupo VPN (UsuariosVPN)", "ok")
                        elif vpn_info.startswith("ERRO:"):
                            self._log(f"     ⚠ Falha ao adicionar ao grupo VPN: {vpn_info[5:].strip()}", "aviso")
                        else:
                            self._log(f"     → Sem VPN (Não elegível)", "info")

                        c["ok"] += 1; c["existe"] -= 1
                        if r_item is not None:
                            r_item.update({
                                "login": sam,
                                "ou": detalhe,
                                "status": "CRIADO (CONFLITO RESOLVIDO)" if status == "OK" else status,
                                "vpn": vpn_info,
                                "detalhe": f"Novo login em vez de {u_data.get('login','')}"
                            })
                    elif status in ("JA_EXISTE", "CONFLITO_NOME"):
                        _ultimo_status = "JA_EXISTE"
                        self._log(f"     ⚠  Login '{sam}' JA EXISTE no AD — informe outro login.", "aviso")
                    else:
                        self._log(f"     ✘  ERRO - {detalhe}", "erro"); c["erro"] += 1
                elif linha.startswith("VPN_OK|"):
                    _, sam = linha.split("|", 1)
                    vpn_info = "SIM"
                elif linha.startswith("VPN_ERRO|"):
                    _, sam, erro_msg = linha.split("|", 2)
                    vpn_info = f"ERRO: {erro_msg}"
                elif linha == "FIM":
                    break
            proc.wait()
        finally:
            os.unlink(tmp.name)
        return _ultimo_status

    # ── Log de importacao em arquivo .csv ─────────────────────────────────────
    @staticmethod
    def _salvar_log_csv(dc, registros, c, csv_origem):
        """Gera um arquivo .csv em logs/ com o resumo da importacao, ordenado por nome em ordem alfabetica."""
        import datetime
        pasta_logs = os.path.join(obter_diretorio_base(), "logs")
        os.makedirs(pasta_logs, exist_ok=True)
        agora = datetime.datetime.now()
        nome_arq = agora.strftime("importacao-%d%m%Y-%H%M%S.csv")
        caminho_arq = os.path.join(pasta_logs, nome_arq)

        # Ordena alfabeticamente por nome (sem considerar acentos nem case)
        registros_ordenados = sorted(
            registros,
            key=lambda r: strip_accents(r.get("nome", "")).lower()
        )

        campos = ["Nome Completo", "Login AD", "Status", "VPN", "Localizacao", "Observacao"]

        with open(caminho_arq, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=campos, delimiter=";")
            writer.writeheader()
            for r in registros_ordenados:
                writer.writerow({
                    "Nome Completo": r.get("nome", ""),
                    "Login AD": str(r.get("login", "")).upper(),
                    "Status": r.get("status", ""),
                    "VPN": r.get("vpn", "NAO"),
                    "Localizacao": ou_para_caminho(r.get("ou", "")),
                    "Observacao": r.get("detalhe", ""),
                })

        return caminho_arq

    def _executar(self, usuarios, user, senha, dc):
        csv_origem = self.var_csv.get().strip()
        script = gerar_ps1(usuarios, dc, self.var_troca.get())
        tmp = tempfile.NamedTemporaryFile(suffix=".ps1", delete=False, mode="w", encoding="utf-8")
        tmp.write(script); tmp.close()
        env = os.environ.copy()
        env["UMU_USER"] = user
        env["UMU_PASS"] = senha
        ja_existe_users = []    # acumula conflitos para tratamento interativo apos o batch
        vpn_status     = {}    # mapeia login -> status da vpn ("SIM", "NAO", "ERRO (msg)")

        # Mapeamento para relatorio final CSV (ordenado por nome)
        registros_dict = {
            id(u): {
                "nome": f"{u['nome']} {u['sobrenome']}".strip(),
                "login": u["login"],
                "status": "PENDENTE",
                "vpn": "SIM" if u.get("vpn") else "NAO",
                "ou": u["ou"],
                "detalhe": "",
            }
            for u in usuarios
        }

        try:
            proc = subprocess.Popen(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmp.name],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                env=env, text=True, encoding="utf-8", errors="replace"
            )
            c = {"ok": 0, "erro": 0, "existe": 0}
            u_atual = None   # usuario sendo processado no momento
            for linha in proc.stdout:
                linha = linha.rstrip()
                if linha.startswith("INICIO|"):
                    _, sam, nome = linha.split("|", 2)
                    u_atual = next((u for u in usuarios if u["login"] == sam), None)
                    vpn_status[sam] = "NAO"
                    self._log(f"  ->  {nome}  ({sam})", "info")
                elif linha.startswith("RESULTADO|"):
                    _, sam, status, detalhe = linha.split("|", 3)
                    r_item = registros_dict.get(id(u_atual)) if u_atual else None
                    if status in ("OK", "OK_EXISTE_ATIVO", "OK_ATIVADO"):
                        caminho = ou_para_caminho(detalhe)
                        v_stat = vpn_status.get(sam, "NAO")
                        if status == "OK_EXISTE_ATIVO":
                            self._log(f"     ✔  JA EXISTE E ATIVO (CRIACAO PULADA)  \u2192  {caminho}", "ok")
                            c["existe"] += 1
                            if r_item:
                                r_item.update({"status": "JA EXISTIA (ATIVO)", "ou": detalhe, "vpn": v_stat, "detalhe": "Criacao pulada"})
                        elif status == "OK_ATIVADO":
                            self._log(f"     ✔  ATIVADO E ATIVO (CRIACAO PULADA)  \u2192  {caminho}", "ok")
                            c["ok"] += 1
                            if r_item:
                                r_item.update({"status": "ATIVADO", "ou": detalhe, "vpn": v_stat, "detalhe": "Reativado no AD"})
                        else:
                            self._log(f"     OK CRIADO  \u2192  {caminho}", "ok")
                            c["ok"] += 1
                            if r_item:
                                r_item.update({"status": "CRIADO", "ou": detalhe, "vpn": v_stat})

                        if v_stat == "SIM":
                            self._log(f"     → Adicionado ao grupo VPN (UsuariosVPN)", "ok")
                        elif v_stat.startswith("ERRO:"):
                            self._log(f"     ⚠ Falha ao adicionar ao grupo VPN: {v_stat[5:].strip()}", "aviso")
                        else:
                            self._log(f"     → Sem VPN (Não elegível)", "info")

                    elif status == "JA_EXISTE":
                        u_conflict = next((u for u in usuarios if u["login"] == sam), None)
                        if u_conflict:
                            ja_existe_users.append(u_conflict)
                            if r_item:
                                r_item.update({"status": "CONFLITO DE LOGIN"})
                        self._log(f"     JA EXISTE - sera tratado apos o batch...", "aviso")
                        c["existe"] += 1
                    elif status == "CONFLITO_NOME":
                        conflict_name, conflict_ou_dn, conflict_status = detalhe.split("|", 2)
                        u_conflict = next((u for u in usuarios if u["login"] == sam), None)
                        if u_conflict:
                            u_conflict["conflict_name"] = conflict_name
                            u_conflict["conflict_ou"] = ou_para_caminho(conflict_ou_dn)
                            u_conflict["conflict_status"] = conflict_status
                            ja_existe_users.append(u_conflict)
                            if r_item:
                                r_item.update({"status": "CONFLITO DE LOGIN", "detalhe": f"Nome no AD: {conflict_name}"})
                        self._log(f"     CONFLITO DE LOGIN (Nome diferente no AD) - sera tratado apos o batch...", "aviso")
                        c["existe"] += 1
                    else:
                        self._log(f"     ERRO - {detalhe}", "erro"); c["erro"] += 1
                        if r_item:
                            r_item.update({"status": "ERRO", "detalhe": detalhe})
                elif linha.startswith("VPN_OK|"):
                    _, sam = linha.split("|", 1)
                    vpn_status[sam] = "SIM"
                    if u_atual and id(u_atual) in registros_dict:
                        registros_dict[id(u_atual)]["vpn"] = "SIM"
                elif linha.startswith("VPN_ERRO|"):
                    _, sam, erro_msg = linha.split("|", 2)
                    vpn_status[sam] = f"ERRO: {erro_msg}"
                    if u_atual and id(u_atual) in registros_dict:
                        registros_dict[id(u_atual)]["vpn"] = f"ERRO: {erro_msg}"
                elif linha == "FIM":
                    break
                elif linha.strip():
                    self._log(f"     {linha}", "aviso")
            proc.wait()
        finally:
            os.unlink(tmp.name)

        # ── Trata conflitos de login interativamente ───────────────────────────
        if ja_existe_users:
            self._log("")
            self._log(f"  {'-'*68}", "aviso")
            self._log(f"  {len(ja_existe_users)} usuario(s) com login em conflito - aguardando operador", "aviso")
            self._log(f"  {'-'*68}", "aviso")
            for u_conflict in ja_existe_users:
                nome_completo = f"{u_conflict['nome']} {u_conflict['sobrenome']}"
                self._log(f"  ?   {nome_completo}  (login atual: {u_conflict['login']})", "aviso")
                r_item = registros_dict.get(id(u_conflict))
                res = self._dialog_novo_login(u_conflict, user, senha, dc, c, r_item=r_item)

        self._log("")
        self._log("=" * 72, "titulo")
        self._log(f"  CONCLUIDO  OK {c['ok']} criados   {c['existe']} ja existiam   XX {c['erro']} erros", "titulo")
        self._log("=" * 72, "titulo")
        self.var_status.set(f"Concluido - {c['ok']} criados, {c['existe']} ja existiam, {c['erro']} erros.")

        # ── Gera arquivo de log .csv ───────────────────────────────────────────
        if c["ok"] > 0 or c["existe"] > 0 or c["erro"] > 0:
            try:
                arq_log = self._salvar_log_csv(dc, list(registros_dict.values()), c, csv_origem)
                self._log("")
                self._log(f"  Log CSV salvo em: {arq_log}", "ok")
                try:
                    os.startfile(arq_log)
                    self._log("  Arquivo de log CSV aberto externamente.", "ok")
                except Exception as e_open:
                    self._log(f"  Aviso: nao foi possivel abrir o arquivo de log automaticamente - {e_open}", "aviso")
            except Exception as e_log:
                self._log(f"  Aviso: nao foi possivel salvar o log CSV - {e_log}", "aviso")

# ──────────────────────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if "--modo-bat" in sys.argv:
        parser = argparse.ArgumentParser()
        parser.add_argument("--modo-bat",   action="store_true")
        parser.add_argument("--csv",        required=True)
        parser.add_argument("--usuario",    required=True)
        parser.add_argument("--senha",      required=True)
        parser.add_argument("--dc",         default="")
        parser.add_argument("--sem-troca",  action="store_true")
        args = parser.parse_args()
        modo_bat(args)
    else:
        csv_preload = ""
        if "--csv-preload" in sys.argv:
            idx = sys.argv.index("--csv-preload")
            if idx + 1 < len(sys.argv):
                csv_preload = sys.argv[idx + 1]
        app = App(csv_preload=csv_preload)
        app.mainloop()