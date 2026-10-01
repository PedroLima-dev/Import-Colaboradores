"""
ImportarUsuariosUmuarama.py
Modos de uso:
  - Interface grafica (duplo clique ou via ExecutarImport.bat)
  - Linha de comando (chamado pelo EXECUTAR_IMPORT_COMPLETO.bat):
      python ImportarUsuariosUmuarama.py --modo-bat --csv <path> --usuario <user> --senha <pass> [--dc <dc>] [--sem-troca]
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import subprocess, threading, sys, os, csv, io, tempfile, unicodedata, chardet, argparse, json, re, datetime

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

OU_MAP_FALLBACK = {
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
    Se o JSON nao existir, retorna o dicionario hardcoded OU_MAP_FALLBACK acima.
    """
    caminho_json = obter_caminho_ou_map()
    if not os.path.exists(caminho_json):
        print(f"[OU-MAP] Arquivo {caminho_json} nao encontrado, usando mapeamento hardcoded.")
        return dict(OU_MAP_FALLBACK)  # copia do fallback

    try:
        with open(caminho_json, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception as e:
        print(f"[OU-MAP] Erro ao ler {caminho_json}: {e}  — usando mapeamento hardcoded.")
        return dict(OU_MAP_FALLBACK)

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
    "CORRETORA": "CORRETORA DE SEGUROS",
    "JEEP": "JEEP E RAM", "RAM": "JEEP E RAM",
}

ALIAS_CIDADE = {
    "MINEIRO":  "MINEIROS",
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
    """Gera script PowerShell de importacao em lote.
    Agora inclui checagem por CPF (atributo description) ANTES de criar,
    para tratar readmissoes de colaboradores com a mesma conta inativa no AD.
    """
    linhas = ""
    for u in usuarios:
        nome_esc = u["nome"].replace('"', "'")
        sob_esc  = u["sobrenome"].replace('"', "'")
        vpn_val  = "$true" if u.get("vpn") else "$false"
        senha_esc = u.get("senha", "").replace('"', '`"')
        linhas += (
            f'    @{{Nome="{nome_esc}";Sobrenome="{sob_esc}";'
            f'Login="{u["login"]}";Email="{u["email"]}";'
            f'OU="{u["ou"]}";CPF="{u["cpf"]}";VPN={vpn_val};Senha="{senha_esc}"}},\n'
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

# Busca usuario no AD pelo CPF (atributo description) — para detectar readmissao
function Get-UsuarioPorCPF([string]$cpf) {{
    if (-not $cpf) {{ return $null }}
    $cpf_limpo = $cpf -replace '[^0-9]', ''
    if (-not $cpf_limpo) {{ return $null }}
    $cpf_com_zeros = if ($cpf_limpo.Length -lt 11) {{ $cpf_limpo.PadLeft(11, '0') }} else {{ $cpf_limpo }}
    $ldap = if ($ServidorDC) {{ "LDAP://$ServidorDC/$BaseDN" }} else {{ "LDAP://$BaseDN" }}
    $e = New-Object System.DirectoryServices.DirectoryEntry($ldap, $credUser, $credSenha)
    $s = New-Object System.DirectoryServices.DirectorySearcher($e)
    # Busca tanto CPF com mascara quanto so digitos e com zeros a esquerda
    $s.Filter = "(&(objectCategory=person)(objectClass=user)(|(description=$cpf)(description=$cpf_limpo)(description=$cpf_com_zeros)(description=*$cpf_limpo*)))"
    $s.SearchScope = "Subtree"
    $res = $s.FindOne()
    if ($res -ne $null) {{
        $entry = $res.GetDirectoryEntry()
        $dispName = $entry.Properties["displayName"].Value
        if ($dispName -eq $null) {{ $dispName = $entry.Properties["cn"].Value }}
        if ($dispName -eq $null) {{ $dispName = "" }}
        $uac = $entry.Properties["userAccountControl"].Value
        if ($uac -eq $null) {{ $uac = 0 }}
        $sam = $entry.Properties["sAMAccountName"].Value
        $dn  = $entry.Properties["distinguishedName"].Value
        return [PSCustomObject]@{{
            Exists      = $true
            DisplayName = [string]$dispName
            SAM         = [string]$sam
            UAC         = [int]$uac
            DN          = [string]$dn
            Path        = [string]$entry.Path
        }}
    }}
    return $null
}}

function Mover-UsuarioOU([System.DirectoryServices.DirectoryEntry]$userEntry, [string]$ouAlvo) {{
    try {{
        $ouEntry = New-Entry $ouAlvo
        if (-not $ouEntry.Guid) {{ throw "OU de destino nao encontrada: $ouAlvo" }}
        $userEntry.MoveTo($ouEntry)
        $ouEntry.Dispose()
        return $true
    }} catch {{
        Write-Output "AVISO_MOVE|$($_.Exception.Message)"
        return $false
    }}
}}

$usuarios = @(
{linhas}
)
foreach ($u in $usuarios) {{
    $nomeCompleto = "$($u.Nome) $($u.Sobrenome)".Trim()
    $sam = $u.Login
    $primeiroNome = $u.Nome.Split(" ")[0]
    $SenhaPadrao = if ($u.Senha) {{ $u.Senha }} else {{ "@" + $primeiroNome.Substring(0,1).ToUpper() + $primeiroNome.Substring(1).ToLower() + "2026" }}
    Write-Output "INICIO|$sam|$nomeCompleto"
    try {{
        # ── Etapa 1: busca pelo login (sAMAccountName) ──────────────────
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

        # ── Etapa 2: busca pelo CPF para detectar readmissao ────────────
        $porCPF = Get-UsuarioPorCPF $u.CPF
        if ($porCPF -ne $null) {{
            $uac_cpf = $porCPF.UAC
            $is_disabled_cpf = (($uac_cpf -band 2) -eq 2)
            if ($is_disabled_cpf) {{
                # Readmissao confirmada: reativa, move OU se necessario, reseta senha
                try {{
                    $userEntry = New-Object System.DirectoryServices.DirectoryEntry($porCPF.Path, $credUser, $credSenha)

                    # Move para nova OU se o DN da OU atual for diferente do alvo
                    $ouAtualDN = $porCPF.DN -replace '^CN=[^,]+,\\s*', ''
                    if ($ouAtualDN -ine $u.OU) {{
                        Mover-UsuarioOU $userEntry $u.OU | Out-Null
                        # Recarrega a entrada apos mover
                        $userEntry.Dispose()
                        $pesquisa2 = Get-UsuarioPorCPF $u.CPF
                        if ($pesquisa2 -ne $null) {{
                            $userEntry = New-Object System.DirectoryServices.DirectoryEntry($pesquisa2.Path, $credUser, $credSenha)
                        }}
                    }}

                    # Atualiza login, email e nome
                    $userEntry.Properties["sAMAccountName"].Value    = $sam
                    $userEntry.Properties["userPrincipalName"].Value = $u.Email
                    $userEntry.Properties["givenName"].Value         = $u.Nome
                    $userEntry.Properties["sn"].Value                = $u.Sobrenome
                    $userEntry.Properties["displayName"].Value       = $nomeCompleto
                    $userEntry.Properties["mail"].Value              = $u.Email

                    # Reativa e reseta senha
                    $currUac = [int]$userEntry.Properties["userAccountControl"].Value
                    $newUac = ($currUac -band -bnot 2) -bor 512
                    $userEntry.Properties["userAccountControl"].Value = $newUac
                    $userEntry.CommitChanges()
                    $userEntry.Invoke("SetPassword", $SenhaPadrao)
                    if ($ForcarTroca) {{ $userEntry.Properties["pwdLastSet"].Value = 0 }}
                    $userEntry.CommitChanges()

                    $userDN = $userEntry.Properties["distinguishedName"].Value
                    $userEntry.Dispose()

                    # VPN se elegivel
                    if ($u.VPN) {{
                        try {{
                            $grupoDN = "CN=UsuariosVPN,OU=OpenVPN,DC=umuarama,DC=local"
                            $ldapGrupo = if ($ServidorDC) {{ "LDAP://$ServidorDC/$grupoDN" }} else {{ "LDAP://$grupoDN" }}
                            $grupoEntry = New-Object System.DirectoryServices.DirectoryEntry($ldapGrupo, $credUser, $credSenha)
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

                    Write-Output "RESULTADO|$sam|OK_READMITIDO|$userDN"
                    continue
                }} catch {{
                    Write-Output "RESULTADO|$sam|ERRO|Falha na readmissao por CPF: $($_.Exception.Message)"
                    continue
                }}
            }} else {{
                # CPF encontrado em conta ATIVA com login diferente — apenas avisa
                $clean_name = $porCPF.DisplayName -replace '\\|', ' '
                Write-Output "RESULTADO|$sam|CONFLITO_CPF|$($porCPF.SAM)|$clean_name"
                continue
            }}
        }}

        # ── Etapa 3: criacao normal (usuario nao existe no AD) ───────────
        $entry = New-Entry $u.OU
        if (-not $entry.Guid) {{ throw "OU nao encontrada: $($u.OU)" }}
        $user = $entry.Children.Add("CN=$nomeCompleto", "user")
        $user.Properties["sAMAccountName"].Value    = $sam
        $user.Properties["userPrincipalName"].Value = $u.Email
        $user.Properties["givenName"].Value         = $u.Nome
        $user.Properties["sn"].Value                = $u.Sobrenome
        $user.Properties["displayName"].Value       = $nomeCompleto
        $user.Properties["mail"].Value              = $u.Email
        if ($u.CPF) {{ $user.Properties["description"].Value = $u.CPF }}
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
# Scripts PowerShell para a aba de Consulta & Gestao
# ──────────────────────────────────────────────────────────────────────────────
def gerar_ps1_consultar(termo, servidor_dc):
    """Gera PS1 que busca um usuario no AD por CPF, Login ou Nome (parcial).
    Emite: ENCONTRADO|sam|nome|email|cpf|uac|dn
           NAO_ENCONTRADO
           ERRO|mensagem
    """
    dc_p   = f'"{servidor_dc}"' if servidor_dc else '""'
    termo_esc = termo.replace('"', "'")
    return f"""
$credUser   = $env:UMU_USER
$credSenha  = $env:UMU_PASS
$ServidorDC = {dc_p}
$BaseDN     = "DC=umuarama,DC=local"
$Termo      = "{termo_esc}"

$ldap = if ($ServidorDC) {{ "LDAP://$ServidorDC/$BaseDN" }} else {{ "LDAP://$BaseDN" }}
try {{
    $root = New-Object System.DirectoryServices.DirectoryEntry($ldap, $credUser, $credSenha)
    if (-not $root.Guid) {{ throw "Falha de autenticacao com o DC." }}
    $s = New-Object System.DirectoryServices.DirectorySearcher($root)
    $s.SearchScope = "Subtree"
    $s.PageSize    = 50
    $s.SizeLimit   = 50

    function Escape-Ldap([string]$str) {{
        if (-not $str) {{ return "" }}
        return $str.Replace('\\', '\\5c').Replace('*', '\\2a').Replace('(', '\\28').Replace(')', '\\29').Replace("`0", '\\00')
    }}

    $termoLimpo = $Termo.Trim()
    $termoLdap  = Escape-Ldap $termoLimpo
    $termoDig   = $termoLimpo -replace '[^0-9]', ''

    # Remove acentos para busca flexivel
    $norm = $termoLimpo.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $norm.ToCharArray()) {{
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {{
            [void]$sb.Append($c)
        }}
    }}
    $termoSemAcento = $sb.ToString()
    $termoSemAcentoLdap = Escape-Ldap $termoSemAcento

    $clausulas = New-Object System.Collections.Generic.List[string]

    # 1. Ambiguous Name Resolution (ANR)
    if ($termoLdap) {{
        $clausulas.Add("(anr=$termoLdap)")
        if ($termoSemAcentoLdap -ne $termoLdap) {{
            $clausulas.Add("(anr=$termoSemAcentoLdap)")
        }}
    }}

    # 2. Login / sAMAccountName (exato e wildcard)
    $clausulas.Add("(sAMAccountName=$termoLdap)")
    $clausulas.Add("(sAMAccountName=*$termoLdap*)")
    if ($termoSemAcentoLdap -ne $termoLdap) {{
        $clausulas.Add("(sAMAccountName=*$termoSemAcentoLdap*)")
    }}

    # 3. Nome de exibicao, CN e Name
    $clausulas.Add("(displayName=*$termoLdap*)")
    $clausulas.Add("(cn=*$termoLdap*)")
    $clausulas.Add("(name=*$termoLdap*)")
    $clausulas.Add("(mail=*$termoLdap*)")
    if ($termoSemAcentoLdap -ne $termoLdap) {{
        $clausulas.Add("(displayName=*$termoSemAcentoLdap*)")
        $clausulas.Add("(cn=*$termoSemAcentoLdap*)")
        $clausulas.Add("(name=*$termoSemAcentoLdap*)")
    }}

    # 4. CPF / Description
    $clausulas.Add("(description=*$termoLdap*)")
    if ($termoDig.Length -ge 4) {{
        $clausulas.Add("(description=*$termoDig*)")
        if ($termoDig.Length -lt 11) {{
            $comZeros = $termoDig.PadLeft(11, '0')
            $clausulas.Add("(description=*$comZeros*)")
        }}
    }}

    $orFilter = "(|" + ($clausulas -join "") + ")"
    $s.Filter = "(&(objectCategory=person)(objectClass=user)$orFilter)"

    $resultados = $s.FindAll()
    if ($resultados.Count -eq 0) {{
        Write-Output "NAO_ENCONTRADO"
    }} else {{
        $vistos = @{{}}
        foreach ($res in $resultados) {{
            $e   = $res.GetDirectoryEntry()
            $sam = [string]$e.Properties["sAMAccountName"].Value
            if (-not $sam) {{ continue }}
            if ($vistos.ContainsKey($sam.ToLower())) {{ continue }}
            $vistos[$sam.ToLower()] = $true

            $dn   = [string]$e.Properties["distinguishedName"].Value
            $disp = [string]$e.Properties["displayName"].Value
            if (-not $disp) {{ $disp = [string]$e.Properties["cn"].Value }}
            if (-not $disp) {{ $disp = [string]$e.Properties["name"].Value }}
            $mail = [string]$e.Properties["mail"].Value
            $cpf  = [string]$e.Properties["description"].Value
            $uac  = $e.Properties["userAccountControl"].Value
            if ($uac -eq $null) {{ $uac = 512 }}

            # Escapa pipes
            $disp = $disp -replace '\\|', ' '
            $mail = $mail -replace '\\|', ' '
            $cpf  = $cpf  -replace '\\|', ' '
            $dn2  = $dn   -replace '\\|', '/'
            Write-Output "ENCONTRADO|$sam|$disp|$mail|$cpf|$uac|$dn2"
        }}
    }}
    $root.Dispose()
}} catch {{
    Write-Output "ERRO|$($_.Exception.Message)"
}}
Write-Output "FIM"
"""


def gerar_ps1_acao(acao, sam, servidor_dc, nova_senha="", forcar_troca=True, nova_ou=""):
    """Gera PS1 para executar uma acao administrativa em uma conta existente.
    acao: 'alterar_senha' | 'desativar' | 'reativar'
    Emite: ACAO_OK|sam|detalhe
           ACAO_ERRO|sam|mensagem
    """
    dc_p     = f'"{servidor_dc}"' if servidor_dc else '""'
    sam_esc  = sam.replace('"', "'")
    senha_esc = nova_senha.replace('"', "'")
    troca    = "$true" if forcar_troca else "$false"
    ou_esc   = nova_ou.replace('"', "'")
    return f"""
$credUser   = $env:UMU_USER
$credSenha  = $env:UMU_PASS
$ServidorDC = {dc_p}
$BaseDN     = "DC=umuarama,DC=local"
$SAM        = "{sam_esc}"
$Acao       = "{acao}"
$NovaSenha  = "{senha_esc}"
$ForcarTroca = {troca}
$NovaOU      = "{ou_esc}"

function New-Entry([string]$path) {{
    $escaped = $path -replace '/', '\\/'
    if ($ServidorDC) {{ $p = "LDAP://$ServidorDC/$escaped" }} else {{ $p = "LDAP://$escaped" }}
    return New-Object System.DirectoryServices.DirectoryEntry($p, $credUser, $credSenha)
}}

$ldap = if ($ServidorDC) {{ "LDAP://$ServidorDC/$BaseDN" }} else {{ "LDAP://$BaseDN" }}
try {{
    $root = New-Object System.DirectoryServices.DirectoryEntry($ldap, $credUser, $credSenha)
    if (-not $root.Guid) {{ throw "Falha de autenticacao com o DC." }}
    $s = New-Object System.DirectoryServices.DirectorySearcher($root)
    $s.Filter = "(&(objectClass=user)(sAMAccountName=$SAM))"
    $s.SearchScope = "Subtree"
    $res = $s.FindOne()
    if ($res -eq $null) {{ throw "Usuario '$SAM' nao encontrado no AD." }}
    $u = $res.GetDirectoryEntry()

    if ($Acao -eq "alterar_senha") {{
        $u.Invoke("SetPassword", $NovaSenha)
        if ($ForcarTroca) {{ $u.Properties["pwdLastSet"].Value = 0 }}
        $u.CommitChanges()
        Write-Output "ACAO_OK|$SAM|Senha alterada com sucesso."

    }} elseif ($Acao -eq "desativar") {{
        $currUac = [int]$u.Properties["userAccountControl"].Value
        $newUac  = $currUac -bor 2
        $u.Properties["userAccountControl"].Value = $newUac
        $u.CommitChanges()
        Write-Output "ACAO_OK|$SAM|Conta desativada."

    }} elseif ($Acao -eq "reativar") {{
        $currUac = [int]$u.Properties["userAccountControl"].Value
        $newUac  = ($currUac -band -bnot 2) -bor 512
        $u.Properties["userAccountControl"].Value = $newUac
        if ($NovaSenha) {{
            $u.Invoke("SetPassword", $NovaSenha)
            if ($ForcarTroca) {{ $u.Properties["pwdLastSet"].Value = 0 }}
        }}
        $u.CommitChanges()

        # Move para nova OU se informada
        if ($NovaOU) {{
            try {{
                $ouEntry = New-Entry $NovaOU
                if (-not $ouEntry.Guid) {{ throw "OU nao encontrada" }}
                $u.MoveTo($ouEntry)
                $ouEntry.Dispose()
                Write-Output "ACAO_OK|$SAM|Conta reativada e movida para nova OU."
            }} catch {{
                Write-Output "ACAO_OK|$SAM|Conta reativada (falha ao mover OU: $($_.Exception.Message))."
            }}
        }} else {{
            Write-Output "ACAO_OK|$SAM|Conta reativada com nova senha."
        }}
    }} else {{
        throw "Acao desconhecida: $Acao"
    }}
    $u.Dispose()
    $root.Dispose()
}} catch {{
    Write-Output "ACAO_ERRO|$SAM|$($_.Exception.Message)"
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
            env=env, text=True, encoding="utf-8", errors="replace",
            creationflags=subprocess.CREATE_NO_WINDOW
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
# Dialogo de Login — exibido antes da janela principal
# ──────────────────────────────────────────────────────────────────────────────
class DialogLogin(tk.Tk):
    """Janela de login modal que valida as credenciais contra o AD via bind LDAP.
    Atributos publicos apos sucesso:
        self.usuario  — login informado (ex: umuarama\\admin)
        self.senha    — senha validada (mantida apenas em memoria)
        self.dc       — servidor DC resolvido
    Encerra o processo se o login for cancelado ou falhar.
    """

    COR_BG   = "#1e1e2e"
    COR_CARD = "#2a2a3e"
    COR_ACC  = "#7c6af7"
    COR_TXT  = "#cdd6f4"
    COR_SUB  = "#6c7086"
    COR_ERR  = "#f38ba8"

    def __init__(self):
        super().__init__()
        self.title("Importar Usuarios AD - Autenticacao")
        self.resizable(False, False)
        self.configure(bg=self.COR_BG)
        self._autenticado = False
        self.usuario = ""
        self.senha   = ""
        self.dc      = ""
        self._build()
        # Centraliza na tela
        self.update_idletasks()
        larg, alt = 420, 420
        x = (self.winfo_screenwidth()  - larg) // 2
        y = (self.winfo_screenheight() - alt)  // 2
        self.geometry(f"{larg}x{alt}+{x}+{y}")
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

    def _build(self):
        C = self  # alias para as cores
        pad = dict(padx=32, pady=0)

        tk.Label(self, text="🔐  Importar Usuarios AD",
                 bg=C.COR_BG, fg=C.COR_ACC,
                 font=("Segoe UI", 15, "bold")).pack(pady=(32, 4))
        tk.Label(self, text="Grupo Umuarama  ·  Autenticação requerida",
                 bg=C.COR_BG, fg=C.COR_SUB,
                 font=("Segoe UI", 9)).pack(pady=(0, 20))

        # Campos
        tk.Label(self, text="Usuário AD  (ex: umuarama\\admin):",
                 bg=C.COR_BG, fg=C.COR_TXT,
                 font=("Segoe UI", 9), anchor="w").pack(fill="x", **pad)
        self._var_user = tk.StringVar()
        self._ent_user = tk.Entry(self, textvariable=self._var_user,
                 bg=C.COR_CARD, fg=C.COR_TXT, insertbackground=C.COR_TXT,
                 font=("Segoe UI", 11), relief="flat", bd=8)
        self._ent_user.pack(fill="x", padx=32, pady=(2, 10))

        tk.Label(self, text="Senha:",
                 bg=C.COR_BG, fg=C.COR_TXT,
                 font=("Segoe UI", 9), anchor="w").pack(fill="x", **pad)
        self._var_pass = tk.StringVar()
        self._ent_pass = tk.Entry(
            self, textvariable=self._var_pass, show="*",
            bg=C.COR_CARD, fg=C.COR_TXT, insertbackground=C.COR_TXT,
            font=("Segoe UI", 11), relief="flat", bd=8)
        self._ent_pass.pack(fill="x", padx=32, pady=(2, 4))

        # Servidor DC
        tk.Label(self, text=f"Servidor DC  (padrão: {DC_DEFAULT_LABEL}):",
                 bg=C.COR_BG, fg=C.COR_TXT,
                 font=("Segoe UI", 9), anchor="w").pack(fill="x", **pad)
        self._var_dc = tk.StringVar(value=DC_DEFAULT_LABEL)
        tk.Entry(self, textvariable=self._var_dc,
                 bg=C.COR_CARD, fg=C.COR_TXT, insertbackground=C.COR_TXT,
                 font=("Segoe UI", 10), relief="flat", bd=6
                 ).pack(fill="x", padx=32, pady=(2, 14))

        # Mensagem de erro
        self._var_msg = tk.StringVar(value="")
        self._lbl_msg = tk.Label(self, textvariable=self._var_msg,
                                 bg=C.COR_BG, fg=C.COR_ERR,
                                 font=("Segoe UI", 8), wraplength=360)
        self._lbl_msg.pack(pady=(0, 6))

        # Botoes
        bf = tk.Frame(self, bg=C.COR_BG)
        bf.pack(pady=6)
        self._btn_entrar = tk.Button(
            bf, text="  Entrar  ",
            bg=C.COR_ACC, fg="white",
            activebackground="#6a58e0", activeforeground="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            padx=18, pady=7, cursor="hand2",
            command=self._tentar_login)
        self._btn_entrar.pack(side="left", padx=8)
        tk.Button(bf, text="Cancelar",
                  bg=C.COR_CARD, fg=C.COR_SUB,
                  activebackground="#3a3a5e",
                  font=("Segoe UI", 10), relief="flat",
                  padx=14, pady=7, cursor="hand2",
                  command=self._cancelar).pack(side="left", padx=8)

        self.bind("<Return>",   lambda _e: self._tentar_login())
        self.bind("<KP_Enter>", lambda _e: self._tentar_login())
        self.bind("<Escape>",   lambda _e: self._cancelar())
        # Foca no campo de usuario ao abrir
        self.after(80, lambda: self._ent_user.focus_set())

    def _resolver_dc(self):
        val = self._var_dc.get().strip()
        return DC_DEFAULT_IP if (not val or val == DC_DEFAULT_LABEL) else val

    def _tentar_login(self):
        user  = self._var_user.get().strip()
        senha = self._var_pass.get()
        if not user or not senha:
            self._var_msg.set("Preencha o usuário e a senha.")
            return

        dc = self._resolver_dc()
        self._var_msg.set("Validando credenciais...")
        self._btn_entrar.configure(state="disabled", text="Aguardando...")
        self.update_idletasks()

        # Validacao via bind LDAP em thread separada para nao travar a UI
        import threading as _threading
        def _validar():
            ok, msg = _autenticar_ad(user, senha, dc)
            self.after(0, lambda: self._resultado_login(ok, msg, user, senha, dc))

        _threading.Thread(target=_validar, daemon=True).start()

    def _resultado_login(self, ok, msg, user, senha, dc):
        if ok:
            self.usuario = user
            self.senha   = senha
            self.dc      = dc
            self._autenticado = True
            self.destroy()
        else:
            self._var_msg.set(f"✘  {msg}")
            self._btn_entrar.configure(state="normal", text="  Entrar  ")

    def _cancelar(self):
        self.destroy()
        sys.exit(0)

    @property
    def autenticado(self):
        return self._autenticado


_GRUPO_RESTRITO = "Admins. do domínio"   # SamAccountName do grupo autorizado

def _autenticar_ad(usuario, senha, dc):
    """Valida as credenciais fazendo bind LDAP contra o AD e verifica se o
    usuario pertence ao grupo '_GRUPO_RESTRITO'.
    Retorna (True, '') em caso de sucesso ou (False, mensagem) em caso de falha.
    Nao armazena as credenciais em disco.
    """
    try:
        import subprocess as _sp, tempfile as _tf, os as _os
        grupo_esc = _GRUPO_RESTRITO.replace('"', "'")
        script = f"""
$credUser  = \"{usuario.replace('"', "'")}\"
$credSenha = \"{senha.replace('"', "'")}\"
$dc        = \"{dc}\"
$base      = \"DC=umuarama,DC=local\"
$ldap      = if ($dc) {{ \"LDAP://$dc/$base\" }} else {{ \"LDAP://$base\" }}

# 1. Bind: valida as credenciais
try {{
    $root = New-Object System.DirectoryServices.DirectoryEntry($ldap, $credUser, $credSenha)
    if ($root.Guid -eq $null -or $root.Guid -eq [guid]::Empty) {{
        Write-Output \"FALHA_CRED|Credenciais invalidas ou DC inacessivel.\"
        exit
    }}
}} catch {{
    Write-Output \"FALHA_CRED|$($_.Exception.Message)\"
    exit
}}

# 2. Localiza o objeto do usuario no AD para obter memberOf recursivo
try {{
    $s = New-Object System.DirectoryServices.DirectorySearcher($root)
    # Extrai apenas o sAMAccountName (sem dominio ex: umuarama\\admin -> admin)
    $sam = (\"{usuario.replace('"', "'")}\" -split '[\\\\\\\\]')[-1]
    $s.Filter = \"(&(objectClass=user)(sAMAccountName=$sam))\"
    $s.SearchScope = \"Subtree\"
    $s.PropertiesToLoad.Add(\"memberOf\") | Out-Null
    $res = $s.FindOne()
    if ($res -eq $null) {{
        Write-Output \"FALHA_GRUPO|Usuario nao encontrado no diretorio.\"
        exit
    }}

    # 3. Verifica memberOf (recursivo via LDAP_MATCHING_RULE_IN_CHAIN)
    $s2 = New-Object System.DirectoryServices.DirectorySearcher($root)
    $s2.Filter = \"(&(objectClass=group)(sAMAccountName={grupo_esc}))\"
    $s2.SearchScope = \"Subtree\"
    $s2.PropertiesToLoad.Add(\"distinguishedName\") | Out-Null
    $grp = $s2.FindOne()
    if ($grp -eq $null) {{
        Write-Output \"FALHA_GRUPO|Grupo restrito '{grupo_esc}' nao encontrado no AD.\"
        exit
    }}
    $grupoDN = $grp.Properties[\"distinguishedName\"][0]

    # Busca o usuario verificando pertencimento recursivo ao grupo
    $s3 = New-Object System.DirectoryServices.DirectorySearcher($root)
    $s3.Filter = \"(&(objectClass=user)(sAMAccountName=$sam)(memberOf:1.2.840.113556.1.4.1941:=$grupoDN))\"
    $s3.SearchScope = \"Subtree\"
    $membro = $s3.FindOne()
    if ($membro -ne $null) {{
        Write-Output \"OK\"
    }} else {{
        Write-Output \"FALHA_GRUPO|Acesso negado. Voce nao pertence ao grupo '{grupo_esc}'.\"
    }}
}} catch {{
    Write-Output \"FALHA_GRUPO|$($_.Exception.Message)\"
}} finally {{
    $root.Dispose()
}}
"""
        tmp = _tf.NamedTemporaryFile(suffix=".ps1", delete=False,
                                    mode="w", encoding="utf-8-sig")
        tmp.write(script); tmp.close()
        proc = _sp.Popen(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", tmp.name],
            stdout=_sp.PIPE, stderr=_sp.STDOUT,
            text=True, encoding="utf-8", errors="replace",
            creationflags=_sp.CREATE_NO_WINDOW)
        saida, _ = proc.communicate(timeout=15)
        _os.unlink(tmp.name)
        for linha in saida.splitlines():
            linha = linha.strip()
            if linha == "OK":
                return True, ""
            if linha.startswith("FALHA_CRED|"):
                return False, linha[11:]
            if linha.startswith("FALHA_GRUPO|"):
                return False, linha[12:]
            if linha.startswith("FALHA|"):   # fallback legado
                return False, linha[6:]
        return False, "Resposta inesperada do AD."
    except Exception as ex:
        return False, str(ex)


# ──────────────────────────────────────────────────────────────────────────────
# Interface Grafica  —  proporcao 16:9  (1120 × 630)
# ──────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self, csv_preload="", sessao_usuario="", sessao_senha="", sessao_dc=""):
        super().__init__()
        self.title("Importar Usuarios AD - Grupo Umuarama")
        self.geometry("1120x630")
        self.minsize(1120, 630)
        self.resizable(True, True)
        self.configure(bg="#1e1e2e")
        self._preview_realizado = False
        # ── Sessao autenticada (credenciais em memoria, sem gravar em disco) ──
        self._sess_usuario = sessao_usuario
        self._sess_senha   = sessao_senha
        self._sess_dc      = sessao_dc
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

        tab3 = ttk.Frame(self._nb, style="TFrame")
        self._nb.add(tab3, text="  CONSULTA & GESTÃO AD  ")
        self._build_tab_gestao(tab3, COR_BG, COR_CARD, COR_ACC, COR_TXT, COR_SUB)

        tab4 = ttk.Frame(self._nb, style="TFrame")
        self._nb.add(tab4, text="  CRIAR USUÁRIO (PJ)  ")
        self._build_tab_individual(tab4, COR_BG, COR_CARD, COR_ACC, COR_TXT, COR_SUB)

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

        # Sessao autenticada: exibe o usuario logado (somente leitura)
        ttk.Label(card, text="Sessão autenticada:", background=COR_CARD).grid(
            row=1, column=0, sticky="w", pady=5)
        lbl_sessao = tk.Label(
            card,
            text=f"✔  {self._sess_usuario}  ·  DC: {self._sess_dc or DC_DEFAULT_LABEL}",
            bg=COR_CARD, fg="#a6e3a1",
            font=("Segoe UI", 9))
        lbl_sessao.grid(row=1, column=1, sticky="w", padx=8)

        # Forcar troca de senha
        self.var_troca = tk.BooleanVar(value=True)
        ttk.Checkbutton(card, text="Forcar troca de senha no 1 login",
                        variable=self.var_troca).grid(
            row=2, column=1, sticky="w", padx=8, pady=6)
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
        Usa o DC da sessao autenticada no login; 'umuarama.local' e campo
        vazio resolvem para DC_DEFAULT_IP (10.56.24.10).
        """
        val = self._sess_dc.strip() if self._sess_dc else ""
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
        # Usa credenciais da sessao autenticada no login
        user  = self._sess_usuario
        senha = self._sess_senha
        dc    = self._resolver_dc()
        if not path or not os.path.exists(path):
            messagebox.showerror("Erro", "Selecione um arquivo CSV valido."); return

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
                env=env, text=True, encoding="utf-8", errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW
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
                env=env, text=True, encoding="utf-8", errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW
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
                    elif status == "OK_READMITIDO":
                        caminho = ou_para_caminho(detalhe)
                        v_stat = vpn_status.get(sam, "NAO")
                        self._log(f"     ✔  READMITIDO (CPF)  →  {caminho}", "ok")
                        c["ok"] += 1
                        if r_item:
                            r_item.update({"status": "READMITIDO", "ou": detalhe, "vpn": v_stat,
                                           "detalhe": "Conta reativada e atualizada via CPF"})
                        if v_stat == "SIM":
                            self._log(f"     → Adicionado ao grupo VPN (UsuariosVPN)", "ok")
                        elif v_stat.startswith("ERRO:"):
                            self._log(f"     ⚠ Falha VPN: {v_stat[5:].strip()}", "aviso")
                    elif status == "CONFLITO_CPF":
                        # CPF ja existe em conta ATIVA com outro login
                        partes_cpf = detalhe.split("|", 1)
                        sam_existente = partes_cpf[0] if partes_cpf else "?"
                        nome_existente = partes_cpf[1] if len(partes_cpf) > 1 else "?"
                        self._log(
                            f"     ⚠  CPF JA EXISTE em conta ATIVA: '{sam_existente}' ({nome_existente})"
                            f" — usuario '{sam}' nao importado.", "aviso")
                        c["existe"] += 1
                        if r_item:
                            r_item.update({"status": "CONFLITO CPF",
                                           "detalhe": f"CPF ja consta na conta ativa: {sam_existente}"})
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

    # ── Aba 3: CONSULTA & GESTÃO AD ───────────────────────────────────────────
    def _build_tab_gestao(self, parent, COR_BG, COR_CARD, COR_ACC, COR_TXT, COR_SUB):
        """Constroi a aba de consulta e gestao de contas no AD."""
        COR_OK    = "#a6e3a1"
        COR_ERR   = "#f38ba8"

        # ── Secao de busca ────────────────────────────────────────────────────
        card_busca = ttk.Frame(parent, style="Card.TFrame", padding=14)
        card_busca.pack(fill="x", padx=0, pady=6)

        ttk.Label(card_busca, text="Buscar colaborador:", background=COR_CARD).grid(
            row=0, column=0, sticky="w", pady=5)
        self.var_gestao_termo = tk.StringVar()
        ent_busca = ttk.Entry(card_busca, textvariable=self.var_gestao_termo, width=46)
        ent_busca.grid(row=0, column=1, padx=8, sticky="ew")
        ttk.Label(card_busca,
                  text="(CPF, login ou nome completo)",
                  background=COR_CARD, foreground=COR_SUB,
                  font=("Segoe UI", 8)).grid(row=0, column=2, sticky="w")
        card_busca.columnconfigure(1, weight=1)

        bf_busca = ttk.Frame(parent, style="TFrame")
        bf_busca.pack(pady=(0, 4))
        self._btn_consultar = ttk.Button(bf_busca, text="Consultar",
                                         style="Accent.TButton",
                                         command=self._consultar_usuario)
        self._btn_consultar.pack(side="left", padx=6)
        ttk.Button(bf_busca, text="Limpar",
                   command=self._limpar_card_usuario).pack(side="left", padx=6)
        ttk.Button(bf_busca, text="+ Criar Usuário (PJ)",
                   command=self._ir_para_criacao_individual).pack(side="left", padx=6)

        ent_busca.bind("<Return>",    lambda _e: self._consultar_usuario())
        ent_busca.bind("<KP_Enter>",  lambda _e: self._consultar_usuario())

        # ── Card de resultado ─────────────────────────────────────────────────
        self._card_resultado = ttk.Frame(parent, style="Card.TFrame", padding=14)
        self._card_resultado.pack(fill="x", padx=0, pady=2)

        # Badge de status
        self._lbl_status_badge = tk.Label(
            self._card_resultado, text="", width=16,
            font=("Segoe UI", 9, "bold"), relief="flat", bd=0)
        self._lbl_status_badge.grid(row=0, column=2, sticky="e", padx=(0, 4), pady=4)

        # Campos informativos
        _labels = [
            ("Nome completo:", "var_g_nome"),
            ("Login (SAM):",   "var_g_login"),
            ("E-mail:",        "var_g_email"),
            ("CPF:",           "var_g_cpf"),
            ("Localização:",   "var_g_ou"),
        ]
        for i, (lbl, var_name) in enumerate(_labels):
            setattr(self, var_name, tk.StringVar(value="—"))
            ttk.Label(self._card_resultado, text=lbl,
                      background=COR_CARD, font=("Segoe UI", 9, "bold")).grid(
                row=i, column=0, sticky="w", pady=2, padx=(0, 8))
            ttk.Label(self._card_resultado, textvariable=getattr(self, var_name),
                      background=COR_CARD, wraplength=520, justify="left").grid(
                row=i, column=1, sticky="w", pady=2)

        self._card_resultado.columnconfigure(1, weight=1)

        # Armazena sam e dn do usuario atual para uso nas acoes
        self._gestao_sam = ""
        self._gestao_dn  = ""
        self._gestao_uac = 0

        # ── Painel de acoes ───────────────────────────────────────────────────
        card_acoes = ttk.Frame(parent, style="Card.TFrame", padding=12)
        card_acoes.pack(fill="x", padx=0, pady=(2, 4))

        tk.Label(card_acoes, text="Ações na conta:", bg=COR_CARD, fg=COR_SUB,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 6))

        bf_acoes = tk.Frame(card_acoes, bg=COR_CARD)
        bf_acoes.pack(fill="x")

        self._btn_alterar_senha = tk.Button(
            bf_acoes, text="🔑  Alterar Senha",
            bg="#313244", fg=COR_TXT, activebackground="#45475a",
            font=("Segoe UI", 10), relief="flat", padx=14, pady=5,
            cursor="hand2", state="disabled",
            command=self._acao_alterar_senha)
        self._btn_alterar_senha.pack(side="left", padx=(0, 8))

        self._btn_desativar = tk.Button(
            bf_acoes, text="🚫  Desativar",
            bg="#313244", fg=COR_ERR, activebackground="#45475a",
            font=("Segoe UI", 10), relief="flat", padx=14, pady=5,
            cursor="hand2", state="disabled",
            command=self._acao_desativar)
        self._btn_desativar.pack(side="left", padx=(0, 8))

        self._btn_reativar = tk.Button(
            bf_acoes, text="🔄  Reativar / Readmissão",
            bg="#313244", fg=COR_OK, activebackground="#45475a",
            font=("Segoe UI", 10), relief="flat", padx=14, pady=5,
            cursor="hand2", state="disabled",
            command=self._acao_reativar)
        self._btn_reativar.pack(side="left", padx=(0, 8))

        # Guarda referencia as cores para reuso nos metodos de acao
        self._G = {
            "COR_BG": COR_BG, "COR_CARD": COR_CARD, "COR_ACC": COR_ACC,
            "COR_TXT": COR_TXT, "COR_SUB": COR_SUB,
            "COR_OK": COR_OK,   "COR_ERR": COR_ERR,
        }

    # ── Consultar colaborador no AD ───────────────────────────────────────────
    def _consultar_usuario(self):
        termo = self.var_gestao_termo.get().strip()
        if not termo:
            messagebox.showwarning("Atenção", "Informe CPF, login ou nome para buscar.")
            return
        # Usa credenciais da sessao autenticada
        user  = self._sess_usuario
        senha = self._sess_senha
        dc    = self._resolver_dc()
        self._btn_consultar.configure(state="disabled")
        self.var_status.set("Consultando AD...")
        self._limpar_card_usuario()
        threading.Thread(
            target=self._executar_consulta,
            args=(termo, user, senha, dc), daemon=True).start()

    def _executar_consulta(self, termo, user, senha, dc):
        script = gerar_ps1_consultar(termo, dc)
        tmp = tempfile.NamedTemporaryFile(suffix=".ps1", delete=False, mode="w", encoding="utf-8")
        tmp.write(script); tmp.close()
        env = os.environ.copy()
        env["UMU_USER"] = user
        env["UMU_PASS"] = senha
        resultados = []
        erro_msg = ""
        try:
            proc = subprocess.Popen(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmp.name],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                env=env, text=True, encoding="utf-8", errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            for linha in proc.stdout:
                linha = linha.rstrip()
                if linha.startswith("ENCONTRADO|"):
                    partes = linha.split("|", 6)
                    # ENCONTRADO|sam|nome|email|cpf|uac|dn
                    if len(partes) >= 7:
                        resultados.append({
                            "sam": partes[1], "nome": partes[2],
                            "email": partes[3], "cpf": partes[4],
                            "uac": int(partes[5]) if partes[5].lstrip("-").isdigit() else 512,
                            "dn": partes[6],
                        })
                elif linha.startswith("ERRO|"):
                    erro_msg = linha[5:]
                elif linha == "FIM":
                    break
            proc.wait()
        finally:
            try:
                os.unlink(tmp.name)
            except OSError:
                pass

        if erro_msg:
            self.after(0, lambda: self._exibir_resultado_consulta({"erro": erro_msg}))
        elif not resultados:
            self.after(0, lambda: self._exibir_resultado_consulta(None))
        elif len(resultados) == 1:
            self.after(0, lambda: self._exibir_resultado_consulta(resultados[0]))
        else:
            # Multiplos resultados: verifica se ha match exato em sam ou cpf
            termo_lower = termo.strip().lower()
            termo_dig = re.sub(r'\D', '', termo_lower)
            exato = None
            for r in resultados:
                if r["sam"].lower() == termo_lower:
                    exato = r
                    break
                cpf_limpo = re.sub(r'\D', '', r.get("cpf", ""))
                if termo_dig and len(termo_dig) >= 8 and (cpf_limpo == termo_dig or cpf_limpo.lstrip('0') == termo_dig.lstrip('0')):
                    exato = r
                    break
            if exato:
                self.after(0, lambda: self._exibir_resultado_consulta(exato))
            else:
                self.after(0, lambda: self._abrir_modal_selecao_usuarios(resultados, termo))

    def _exibir_resultado_consulta(self, resultado):
        self._btn_consultar.configure(state="normal")
        G = self._G
        if resultado is None:
            self.var_status.set("Nenhum colaborador encontrado.")
            self._log("  Nenhum colaborador encontrado para o termo informado.", "aviso")
            resp = messagebox.askyesno(
                "Não encontrado",
                "Nenhuma conta localizada no AD com esse CPF, login ou nome.\n\n"
                "Deseja abrir o formulário para cadastrar como Novo Usuário (PJ / Individual) agora?",
                parent=self
            )
            if resp:
                self._ir_para_criacao_individual()
            return
        if "erro" in resultado:
            self.var_status.set("Erro na consulta.")
            self._log(f"  Erro ao consultar AD: {resultado['erro']}", "erro")
            messagebox.showerror("Erro de consulta", resultado["erro"])
            return

        sam  = resultado["sam"]
        nome = resultado["nome"]
        uac  = resultado["uac"]
        dn   = resultado["dn"]
        is_disabled = (uac & 2) == 2

        # Preenche o card
        self.var_g_nome.set(nome or "—")
        self.var_g_login.set(sam or "—")
        self.var_g_email.set(resultado["email"] or "—")
        self.var_g_cpf.set(resultado["cpf"] or "—")
        # OU legivel: extrai da DN
        ou_dn = re.sub(r"^CN=[^,]+,\s*", "", dn)
        self.var_g_ou.set(ou_para_caminho(ou_dn) if ou_dn else "—")

        # Badge de status
        if is_disabled:
            self._lbl_status_badge.configure(
                text="  DESATIVADO  ", bg="#f38ba8", fg="#1e1e2e")
        else:
            self._lbl_status_badge.configure(
                text="  ATIVO  ", bg="#a6e3a1", fg="#1e1e2e")

        # Habilita botoes de acao de acordo com o estado
        self._gestao_sam = sam
        self._gestao_dn  = dn
        self._gestao_uac = uac

        self._btn_alterar_senha.configure(state="normal")
        if is_disabled:
            self._btn_desativar.configure(state="disabled")
            self._btn_reativar.configure(state="normal")
        else:
            self._btn_desativar.configure(state="normal")
            self._btn_reativar.configure(state="disabled")

        self.var_status.set(f"Encontrado: {nome}  ({'DESATIVADO' if is_disabled else 'ATIVO'})")
        self._log(f"  Consulta: {nome} ({sam}) — {'DESATIVADO' if is_disabled else 'ATIVO'}", "ok")

    def _limpar_card_usuario(self):
        """Reseta o card de resultado e desabilita os botoes de acao."""
        for v in ("var_g_nome", "var_g_login", "var_g_email", "var_g_cpf", "var_g_ou"):
            if hasattr(self, v):
                getattr(self, v).set("—")
        if hasattr(self, "_lbl_status_badge"):
            self._lbl_status_badge.configure(text="", bg=self._G.get("COR_CARD", "#2a2a3e"))
        self._gestao_sam = ""
        self._gestao_dn  = ""
        self._gestao_uac = 0
        for btn in ("_btn_alterar_senha", "_btn_desativar", "_btn_reativar"):
            if hasattr(self, btn):
                getattr(self, btn).configure(state="disabled")

    def _abrir_modal_selecao_usuarios(self, lista, termo):
        """Abre modal para o operador escolher quando a busca retorna multiplas contas."""
        self._btn_consultar.configure(state="normal")
        self.var_status.set(f"{len(lista)} colaboradores encontrados para '{termo}'.")
        self._log(f"  Consulta: {len(lista)} contas encontradas para '{termo}'. Selecione uma.", "aviso")

        G = self._G
        dlg = tk.Toplevel(self)
        dlg.title("Selecionar Colaborador")
        dlg.configure(bg=G["COR_BG"])
        dlg.geometry("740x420")
        dlg.minsize(620, 320)
        dlg.grab_set()

        # Centraliza na janela pai
        self.update_idletasks()
        px = self.winfo_x() + max(0, (self.winfo_width() - 740) // 2)
        py = self.winfo_y() + max(0, (self.winfo_height() - 420) // 2)
        dlg.geometry(f"740x420+{px}+{py}")

        header = tk.Frame(dlg, bg=G["COR_BG"])
        header.pack(fill="x", padx=16, pady=(12, 6))
        tk.Label(header, text="🔍  Várias contas localizadas", bg=G["COR_BG"], fg=G["COR_ACC"],
                 font=("Segoe UI", 12, "bold")).pack(anchor="w")
        tk.Label(header, text=f"Foram encontradas {len(lista)} contas para '{termo}'. Selecione a conta desejada:",
                 bg=G["COR_BG"], fg=G["COR_TXT"], font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 0))

        # Treeview de resultados
        frame_tree = tk.Frame(dlg, bg=G["COR_BG"])
        frame_tree.pack(fill="both", expand=True, padx=16, pady=6)

        cols = ("nome", "sam", "cpf", "status", "ou")
        tree = ttk.Treeview(frame_tree, columns=cols, show="headings", selectmode="browse")
        tree.heading("nome", text="Nome Completo")
        tree.heading("sam", text="Login")
        tree.heading("cpf", text="CPF")
        tree.heading("status", text="Status")
        tree.heading("ou", text="Localização")

        tree.column("nome", width=220, minwidth=140)
        tree.column("sam", width=130, minwidth=90)
        tree.column("cpf", width=110, minwidth=80)
        tree.column("status", width=90, minwidth=70, anchor="center")
        tree.column("ou", width=170, minwidth=120)

        sb = ttk.Scrollbar(frame_tree, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        for idx, u in enumerate(lista):
            is_dis = (u["uac"] & 2) == 2
            st_text = "DESATIVADO" if is_dis else "ATIVO"
            ou_dn = re.sub(r"^CN=[^,]+,\s*", "", u["dn"])
            caminho_ou = ou_para_caminho(ou_dn) if ou_dn else "—"
            tree.insert("", "end", iid=str(idx), values=(
                u["nome"] or "—",
                u["sam"] or "—",
                u["cpf"] or "—",
                st_text,
                caminho_ou
            ))

        def _confirmar():
            sel = tree.selection()
            if not sel:
                return
            idx = int(sel[0])
            dlg.destroy()
            self._exibir_resultado_consulta(lista[idx])

        tree.bind("<Double-1>", lambda _e: _confirmar())
        tree.bind("<Return>", lambda _e: _confirmar())

        bf = tk.Frame(dlg, bg=G["COR_BG"])
        bf.pack(fill="x", padx=16, pady=(6, 12))
        tk.Button(bf, text="Selecionar", bg=G["COR_ACC"], fg="white",
                  activebackground="#6a58e0", font=("Segoe UI", 9, "bold"),
                  relief="flat", padx=14, pady=5, cursor="hand2",
                  command=_confirmar).pack(side="left", padx=(0, 8))
        tk.Button(bf, text="Cancelar", bg=G["COR_CARD"], fg=G["COR_SUB"],
                  activebackground="#3a3a5e", font=("Segoe UI", 9),
                  relief="flat", padx=14, pady=5, cursor="hand2",
                  command=dlg.destroy).pack(side="left")

        if lista:
            tree.selection_set("0")
            tree.focus("0")

    # ── Acao: Alterar Senha ───────────────────────────────────────────────────
    def _acao_alterar_senha(self):
        if not self._gestao_sam:
            return
        G = self._G
        dlg = tk.Toplevel(self)
        dlg.title("Alterar Senha")
        dlg.configure(bg=G["COR_BG"])
        dlg.resizable(False, False)
        self.update_idletasks()
        larg, alt = 440, 280
        px = self.winfo_x() + max(0, (self.winfo_width() - larg) // 2)
        py = self.winfo_y() + max(0, (self.winfo_height() - alt) // 2)
        dlg.geometry(f"{larg}x{alt}+{px}+{py}")
        dlg.grab_set()

        cont = tk.Frame(dlg, bg=G["COR_BG"])
        cont.pack(fill="both", expand=True, padx=20, pady=16)

        tk.Label(cont, text="🔑  Alterar Senha", bg=G["COR_BG"], fg=G["COR_ACC"],
                 font=("Segoe UI", 12, "bold")).pack(pady=(0, 4))
        tk.Label(cont, text=f"Conta: {self._gestao_sam}",
                 bg=G["COR_BG"], fg=G["COR_TXT"], font=("Segoe UI", 9)).pack(pady=(0, 10))

        tk.Label(cont, text="Nova senha:", bg=G["COR_BG"], fg=G["COR_TXT"],
                 font=("Segoe UI", 9)).pack(anchor="w")
        var_senha = tk.StringVar(value=self._gerar_senha_padrao(self.var_g_nome.get()))
        row_s = tk.Frame(cont, bg=G["COR_BG"])
        row_s.pack(fill="x", pady=(2, 6))
        ent_senha = tk.Entry(row_s, textvariable=var_senha, bg=G["COR_CARD"], fg=G["COR_TXT"],
                             insertbackground=G["COR_TXT"], font=("Segoe UI", 11),
                             relief="flat", bd=6)
        ent_senha.pack(side="left", fill="x", expand=True)
        tk.Button(row_s, text="⟳", bg=G["COR_CARD"], fg=G["COR_SUB"],
                  relief="flat", font=("Segoe UI", 11), cursor="hand2",
                  command=lambda: var_senha.set(
                      self._gerar_senha_padrao(self.var_g_nome.get()))
                  ).pack(side="left", padx=(4, 0))

        var_troca = tk.BooleanVar(value=True)
        tk.Checkbutton(cont, text="Forçar troca no próximo logon",
                       variable=var_troca, bg=G["COR_BG"], fg=G["COR_TXT"],
                       selectcolor=G["COR_CARD"], activebackground=G["COR_BG"],
                       font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 10))

        def _confirmar():
            senha = var_senha.get().strip()
            if not senha:
                messagebox.showwarning("Atenção", "Informe a nova senha.", parent=dlg)
                return
            dlg.destroy()
            self._executar_acao_ad("alterar_senha", self._gestao_sam,
                                   nova_senha=senha, forcar_troca=var_troca.get())

        bf = tk.Frame(cont, bg=G["COR_BG"])
        bf.pack(pady=4)
        tk.Button(bf, text="Confirmar", bg=G["COR_ACC"], fg="white",
                  activebackground="#6a58e0", font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=14, pady=5, cursor="hand2",
                  command=_confirmar).pack(side="left", padx=8)
        tk.Button(bf, text="Cancelar", bg=G["COR_CARD"], fg=G["COR_SUB"],
                  activebackground="#3a3a5e", font=("Segoe UI", 10),
                  relief="flat", padx=14, pady=5, cursor="hand2",
                  command=dlg.destroy).pack(side="left", padx=8)
        dlg.bind("<Return>",   lambda _e: _confirmar())
        dlg.bind("<Escape>",   lambda _e: dlg.destroy())
        ent_senha.focus_set()

    # ── Acao: Desativar ───────────────────────────────────────────────────────
    def _acao_desativar(self):
        if not self._gestao_sam:
            return
        sam  = self._gestao_sam
        nome = self.var_g_nome.get()
        if not messagebox.askyesno(
                "Confirmar desativação",
                f"Desativar a conta '{sam}' ({nome}) no Active Directory?\n\n"
                f"O colaborador perderá o acesso imediatamente."):
            return
        self._executar_acao_ad("desativar", sam)

    # ── Acao: Reativar / Readmissao ───────────────────────────────────────────
    def _acao_reativar(self):
        if not self._gestao_sam:
            return
        G    = self._G
        sam  = self._gestao_sam
        nome = self.var_g_nome.get()

        dlg = tk.Toplevel(self)
        dlg.title("Reativar / Readmissão")
        dlg.configure(bg=G["COR_BG"])
        dlg.resizable(False, False)
        self.update_idletasks()
        larg, alt = 500, 360
        px = self.winfo_x() + max(0, (self.winfo_width() - larg) // 2)
        py = self.winfo_y() + max(0, (self.winfo_height() - alt) // 2)
        dlg.geometry(f"{larg}x{alt}+{px}+{py}")
        dlg.grab_set()

        cont = tk.Frame(dlg, bg=G["COR_BG"])
        cont.pack(fill="both", expand=True, padx=20, pady=16)

        tk.Label(cont, text="🔄  Reativar / Readmissão", bg=G["COR_BG"], fg=G["COR_OK"],
                 font=("Segoe UI", 12, "bold")).pack(pady=(0, 2))
        tk.Label(cont, text=f"Conta: {sam}  ({nome})",
                 bg=G["COR_BG"], fg=G["COR_TXT"], font=("Segoe UI", 9)).pack(pady=(0, 10))

        # Nova senha
        tk.Label(cont, text="Nova senha temporária:", bg=G["COR_BG"], fg=G["COR_TXT"],
                 font=("Segoe UI", 9)).pack(anchor="w")
        var_senha = tk.StringVar(value=self._gerar_senha_padrao(nome))
        row_s = tk.Frame(cont, bg=G["COR_BG"])
        row_s.pack(fill="x", pady=(2, 6))
        ent_senha = tk.Entry(row_s, textvariable=var_senha, bg=G["COR_CARD"], fg=G["COR_TXT"],
                             insertbackground=G["COR_TXT"], font=("Segoe UI", 11),
                             relief="flat", bd=6)
        ent_senha.pack(side="left", fill="x", expand=True)
        tk.Button(row_s, text="⟳", bg=G["COR_CARD"], fg=G["COR_SUB"],
                  relief="flat", font=("Segoe UI", 11), cursor="hand2",
                  command=lambda: var_senha.set(self._gerar_senha_padrao(nome))
                  ).pack(side="left", padx=(4, 0))

        var_troca = tk.BooleanVar(value=True)
        tk.Checkbutton(cont, text="Forçar troca no próximo logon",
                       variable=var_troca, bg=G["COR_BG"], fg=G["COR_TXT"],
                       selectcolor=G["COR_CARD"], activebackground=G["COR_BG"],
                       font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))

        # Nova OU (opcional)
        tk.Label(cont, text="Nova OU (opcional — para readmissão em outra unidade):",
                 bg=G["COR_BG"], fg=G["COR_TXT"], font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(cont,
                 text="Deixe em branco para manter na OU atual.",
                 bg=G["COR_BG"], fg=G["COR_SUB"], font=("Segoe UI", 8)).pack(anchor="w")
        var_ou = tk.StringVar()
        ttk.Entry(cont, textvariable=var_ou).pack(fill="x", pady=(2, 10))

        def _confirmar():
            senha = var_senha.get().strip()
            if not senha:
                messagebox.showwarning("Atenção", "Informe a nova senha temporária.", parent=dlg)
                return
            dlg.destroy()
            self._executar_acao_ad("reativar", sam,
                                   nova_senha=senha,
                                   forcar_troca=var_troca.get(),
                                   nova_ou=var_ou.get().strip())

        bf = tk.Frame(cont, bg=G["COR_BG"])
        bf.pack(pady=4)
        tk.Button(bf, text="Reativar", bg=G["COR_OK"], fg="#1e1e2e",
                  activebackground="#89d8a1", font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=14, pady=5, cursor="hand2",
                  command=_confirmar).pack(side="left", padx=8)
        tk.Button(bf, text="Cancelar", bg=G["COR_CARD"], fg=G["COR_SUB"],
                  activebackground="#3a3a5e", font=("Segoe UI", 10),
                  relief="flat", padx=14, pady=5, cursor="hand2",
                  command=dlg.destroy).pack(side="left", padx=8)
        dlg.bind("<Return>", lambda _e: _confirmar())
        dlg.bind("<Escape>", lambda _e: dlg.destroy())
        ent_senha.focus_set()

    # ── Executa acao administrativa via PowerShell ────────────────────────────
    def _executar_acao_ad(self, acao, sam, nova_senha="", forcar_troca=True, nova_ou=""):
        # Usa credenciais da sessao autenticada
        user  = self._sess_usuario
        senha = self._sess_senha
        dc    = self._resolver_dc()

        nomes_acao = {
            "alterar_senha": "Alterando senha",
            "desativar":     "Desativando conta",
            "reativar":      "Reativando conta",
        }
        self.var_status.set(f"{nomes_acao.get(acao, acao)} de {sam}...")

        def _run():
            script = gerar_ps1_acao(acao, sam, dc,
                                    nova_senha=nova_senha,
                                    forcar_troca=forcar_troca,
                                    nova_ou=nova_ou)
            tmp = tempfile.NamedTemporaryFile(
                suffix=".ps1", delete=False, mode="w", encoding="utf-8")
            tmp.write(script); tmp.close()
            env = os.environ.copy()
            env["UMU_USER"] = user
            env["UMU_PASS"] = senha
            msg_ok   = ""
            msg_erro = ""
            try:
                proc = subprocess.Popen(
                    ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                     "-File", tmp.name],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    env=env, text=True, encoding="utf-8", errors="replace",
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                for linha in proc.stdout:
                    linha = linha.rstrip()
                    if linha.startswith("ACAO_OK|"):
                        partes = linha.split("|", 2)
                        msg_ok = partes[2] if len(partes) > 2 else "Concluido."
                    elif linha.startswith("ACAO_ERRO|"):
                        partes = linha.split("|", 2)
                        msg_erro = partes[2] if len(partes) > 2 else "Erro desconhecido."
                    elif linha == "FIM":
                        break
                proc.wait()
            finally:
                try:
                    os.unlink(tmp.name)
                except OSError:
                    pass

            def _atualizar():
                if msg_ok:
                    self._log(f"  ✔  [{sam}] {msg_ok}", "ok")
                    self.var_status.set(f"OK: {sam} — {msg_ok}")
                    # Atualiza o badge de status se desativou ou reativou
                    if acao == "desativar":
                        self._lbl_status_badge.configure(
                            text="  DESATIVADO  ", bg="#f38ba8", fg="#1e1e2e")
                        self._btn_desativar.configure(state="disabled")
                        self._btn_reativar.configure(state="normal")
                    elif acao == "reativar":
                        self._lbl_status_badge.configure(
                            text="  ATIVO  ", bg="#a6e3a1", fg="#1e1e2e")
                        self._btn_desativar.configure(state="normal")
                        self._btn_reativar.configure(state="disabled")
                        # Re-consulta para atualizar OU no card apos possivel movimentacao
                        self.after(500, self._consultar_usuario)
                elif msg_erro:
                    self._log(f"  ✘  [{sam}] {msg_erro}", "erro")
                    self.var_status.set(f"Erro em {sam}: {msg_erro}")
                    messagebox.showerror("Erro na ação", msg_erro)

            self.after(0, _atualizar)

        threading.Thread(target=_run, daemon=True).start()

    # ── Gerador de senha padrao ───────────────────────────────────────────────
    @staticmethod
    def _gerar_senha_padrao(nome_completo):
        """Gera a senha padrao @PrimeiroNome2026 a partir do nome completo."""
        primeiro = nome_completo.strip().split()[0] if nome_completo.strip() else "Colaborador"
        primeiro = unicodedata.normalize("NFD", primeiro).encode("ascii", "ignore").decode()
        if len(primeiro) >= 2:
            return "@" + primeiro[0].upper() + primeiro[1:].lower() + "2026"
        return "@" + primeiro.upper() + "2026"


# ──────────────────────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────────────────────

    # ── Aba 4: CRIAR USUÁRIO INDIVIDUAL (PJ / PRESTADOR) ─────────────────────
    def _build_tab_individual(self, parent, COR_BG, COR_CARD, COR_ACC, COR_TXT, COR_SUB):
        """Constrói a aba de cadastro individual de usuários (ex: PJ ou terceirizados)."""
        card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        card.pack(fill="x", padx=0, pady=6)

        # Variáveis do formulário
        self.var_indiv_nome  = tk.StringVar()
        self.var_indiv_login = tk.StringVar()
        self.var_indiv_email = tk.StringVar()
        self.var_indiv_cpf   = tk.StringVar()
        self.var_indiv_ou    = tk.StringVar()
        self.var_indiv_ou_dn = tk.StringVar(value="Selecione uma localização na lista acima...")
        self.var_indiv_senha = tk.StringVar()
        self.var_indiv_troca = tk.BooleanVar(value=True)
        self.var_indiv_vpn   = tk.BooleanVar(value=False)

        self._indiv_login_manual = False
        self._indiv_email_manual = False
        self._indiv_senha_manual = False
        self._bloquear_auto_indiv = False

        # Prepara OUs legíveis
        self._carregar_lista_ous_indiv()

        # ── Grid do formulário ──
        # Linha 0: Nome Completo
        ttk.Label(card, text="Nome completo:*", background=COR_CARD, font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, sticky="w", pady=4)
        self.ent_indiv_nome = ttk.Entry(card, textvariable=self.var_indiv_nome, width=38)
        self.ent_indiv_nome.grid(row=0, column=1, padx=(6, 12), sticky="ew", pady=4)
        ttk.Label(card, text="(Nome e sobrenome — sugere Login, E-mail e Senha)",
                  background=COR_CARD, foreground=COR_SUB, font=("Segoe UI", 8)).grid(
            row=0, column=2, columnspan=2, sticky="w", pady=4)

        # Linha 1: Login (SAM) e E-mail
        ttk.Label(card, text="Login (SAM):*", background=COR_CARD, font=("Segoe UI", 9, "bold")).grid(
            row=1, column=0, sticky="w", pady=4)
        self.ent_indiv_login = ttk.Entry(card, textvariable=self.var_indiv_login, width=28)
        self.ent_indiv_login.grid(row=1, column=1, padx=(6, 12), sticky="w", pady=4)

        ttk.Label(card, text="E-mail:*", background=COR_CARD, font=("Segoe UI", 9, "bold")).grid(
            row=1, column=2, sticky="w", pady=4, padx=(10, 0))
        self.ent_indiv_email = ttk.Entry(card, textvariable=self.var_indiv_email, width=36)
        self.ent_indiv_email.grid(row=1, column=3, padx=(6, 0), sticky="ew", pady=4)

        # Linha 2: CPF e Senha Inicial
        ttk.Label(card, text="CPF:", background=COR_CARD, font=("Segoe UI", 9, "bold")).grid(
            row=2, column=0, sticky="w", pady=4)
        self.ent_indiv_cpf = ttk.Entry(card, textvariable=self.var_indiv_cpf, width=28)
        self.ent_indiv_cpf.grid(row=2, column=1, padx=(6, 12), sticky="w", pady=4)

        ttk.Label(card, text="Senha inicial:*", background=COR_CARD, font=("Segoe UI", 9, "bold")).grid(
            row=2, column=2, sticky="w", pady=4, padx=(10, 0))

        frame_senha = tk.Frame(card, bg=COR_CARD)
        frame_senha.grid(row=2, column=3, sticky="w", padx=(6, 0), pady=4)

        self.ent_indiv_senha = tk.Entry(frame_senha, textvariable=self.var_indiv_senha,
                                        bg="#313244", fg=COR_TXT, insertbackground=COR_TXT,
                                        relief="flat", bd=4, font=("Segoe UI", 9), width=20)
        self.ent_indiv_senha.pack(side="left", padx=(0, 6))

        tk.Button(frame_senha, text="🔄 Regerar", bg="#313244", fg=COR_TXT,
                  activebackground="#45475a", font=("Segoe UI", 8),
                  relief="flat", padx=6, pady=2, cursor="hand2",
                  command=self._regerar_senha_individual).pack(side="left")

        # Linha 3: Localização (OU)
        ttk.Label(card, text="Localização (OU):*", background=COR_CARD, font=("Segoe UI", 9, "bold")).grid(
            row=3, column=0, sticky="w", pady=4)

        self.cb_indiv_ou = ttk.Combobox(card, textvariable=self.var_indiv_ou,
                                        values=self._lista_ous_display,
                                        width=56)
        self.cb_indiv_ou.grid(row=3, column=1, columnspan=3, padx=(6, 0), sticky="ew", pady=4)
        self.cb_indiv_ou.bind("<<ComboboxSelected>>", self._on_indiv_ou_selected)
        self.cb_indiv_ou.bind("<KeyRelease>", self._filtrar_ous_combobox)

        # Linha 4: DN correspondente
        ttk.Label(card, text="DN no AD:", background=COR_CARD,
                  font=("Segoe UI", 8), foreground=COR_SUB).grid(row=4, column=0, sticky="w", pady=(0, 4))
        self.lbl_indiv_dn = tk.Label(card, textvariable=self.var_indiv_ou_dn,
                                     bg=COR_CARD, fg="#89b4fa", font=("Segoe UI", 8),
                                     anchor="w", wraplength=700, justify="left")
        self.lbl_indiv_dn.grid(row=4, column=1, columnspan=3, sticky="w", padx=(6, 0), pady=(0, 4))

        # Linha 5: Opções
        frame_opts = tk.Frame(card, bg=COR_CARD)
        frame_opts.grid(row=5, column=1, columnspan=3, sticky="w", padx=(6, 0), pady=(4, 6))

        tk.Checkbutton(frame_opts, text="Forçar troca de senha no 1º logon",
                       variable=self.var_indiv_troca, bg=COR_CARD, fg=COR_TXT,
                       selectcolor="#1e1e2e", activebackground=COR_CARD,
                       font=("Segoe UI", 9)).pack(side="left", padx=(0, 20))

        tk.Checkbutton(frame_opts, text="Habilitar acesso VPN (Grupo UsuariosVPN)",
                       variable=self.var_indiv_vpn, bg=COR_CARD, fg=COR_TXT,
                       selectcolor="#1e1e2e", activebackground=COR_CARD,
                       font=("Segoe UI", 9)).pack(side="left")

        card.columnconfigure(1, weight=1)
        card.columnconfigure(3, weight=1)

        # Botões de Ação
        bf = ttk.Frame(parent, style="TFrame")
        bf.pack(pady=6)

        self._btn_indiv_criar = ttk.Button(bf, text="Criar Usuário no AD",
                                           style="Accent.TButton",
                                           command=self._executar_criacao_individual)
        self._btn_indiv_criar.pack(side="left", padx=6)

        ttk.Button(bf, text="Limpar Campos",
                   command=self._limpar_campos_individual).pack(side="left", padx=6)

        # Triggers de automação
        self.var_indiv_nome.trace_add("write", self._on_indiv_nome_change)
        self.ent_indiv_login.bind("<Key>", self._marcar_login_manual)
        self.ent_indiv_email.bind("<Key>", self._marcar_email_manual)
        self.ent_indiv_senha.bind("<Key>", self._marcar_senha_manual)
        self.var_indiv_cpf.trace_add("write", self._formatar_cpf_individual)

    def _marcar_login_manual(self, event=None):
        if event and event.keysym not in ("Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock"):
            self._indiv_login_manual = True

    def _marcar_email_manual(self, event=None):
        if event and event.keysym not in ("Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock"):
            self._indiv_email_manual = True

    def _marcar_senha_manual(self, event=None):
        if event and event.keysym not in ("Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock"):
            self._indiv_senha_manual = True

    def _carregar_lista_ous_indiv(self):
        """Carrega e ordena as OUs mapeadas no ou_map.json."""
        self._mapa_ous_legiveis = {}
        for (marca, cidade), dn in OU_MAP.items():
            if dn:
                caminho = ou_para_caminho(dn)
                self._mapa_ous_legiveis[caminho] = dn
        self._lista_ous_display = sorted(self._mapa_ous_legiveis.keys())

    def _on_indiv_ou_selected(self, _event=None):
        sel = self.cb_indiv_ou.get().strip()
        dn = self._mapa_ous_legiveis.get(sel, "")
        if dn:
            self.var_indiv_ou_dn.set(dn)
        else:
            self.var_indiv_ou_dn.set("OU personalizada / manual")

    def _filtrar_ous_combobox(self, event=None):
        if event and event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return
        digitado = self.cb_indiv_ou.get().strip().lower()
        if not digitado:
            self.cb_indiv_ou["values"] = self._lista_ous_display
            return
        filtrados = [ou for ou in self._lista_ous_display if digitado in ou.lower()]
        self.cb_indiv_ou["values"] = filtrados if filtrados else [self.cb_indiv_ou.get()]
        sel = self.cb_indiv_ou.get().strip()
        dn = self._mapa_ous_legiveis.get(sel, "")
        if dn:
            self.var_indiv_ou_dn.set(dn)

    def _on_indiv_nome_change(self, *args):
        if self._bloquear_auto_indiv:
            return
        nome = self.var_indiv_nome.get()
        if not self._indiv_login_manual:
            login = gerar_login(nome)
            self.var_indiv_login.set(login)
            if not self._indiv_email_manual:
                self.var_indiv_email.set(f"{login}@umuarama.local" if login else "")
        if not self._indiv_senha_manual:
            if nome.strip():
                self.var_indiv_senha.set(self._gerar_senha_padrao(nome))
            else:
                self.var_indiv_senha.set("")

    def _formatar_cpf_individual(self, *args):
        val = self.var_indiv_cpf.get()
        nums = re.sub(r'[^0-9]', '', val)
        if len(nums) == 11 and "." not in val and "-" not in val:
            fmt = f"{nums[:3]}.{nums[3:6]}.{nums[6:9]}-{nums[9:]}"
            self.var_indiv_cpf.set(fmt)

    def _regerar_senha_individual(self):
        nome = self.var_indiv_nome.get().strip()
        self.var_indiv_senha.set(self._gerar_senha_padrao(nome) if nome else "@Umuarama2026")
        self._indiv_senha_manual = False

    def _limpar_campos_individual(self):
        self._bloquear_auto_indiv = True
        self.var_indiv_nome.set("")
        self.var_indiv_login.set("")
        self.var_indiv_email.set("")
        self.var_indiv_cpf.set("")
        self.var_indiv_ou.set("")
        self.var_indiv_ou_dn.set("Selecione uma localização na lista acima...")
        self.var_indiv_senha.set("")
        self.var_indiv_troca.set(True)
        self.var_indiv_vpn.set(False)
        self._indiv_login_manual = False
        self._indiv_email_manual = False
        self._indiv_senha_manual = False
        self._bloquear_auto_indiv = False
        self.cb_indiv_ou["values"] = self._lista_ous_display

    def _ir_para_criacao_individual(self):
        termo = self.var_gestao_termo.get().strip()
        self._nb.select(3)
        if termo:
            termo_dig = re.sub(r'[^0-9]', '', termo)
            if len(termo_dig) == 11 or (len(termo_dig) >= 9 and "." in termo):
                self.var_indiv_cpf.set(termo)
            elif "." in termo and " " not in termo:
                self.var_indiv_login.set(termo)
                self._indiv_login_manual = True
            else:
                self.var_indiv_nome.set(termo)
                self.ent_indiv_nome.focus_set()

    def _executar_criacao_individual(self):
        nome_completo = self.var_indiv_nome.get().strip()
        login = self.var_indiv_login.get().strip().lower()
        email = self.var_indiv_email.get().strip()
        cpf = self.var_indiv_cpf.get().strip()
        ou_display = self.cb_indiv_ou.get().strip()
        senha = self.var_indiv_senha.get().strip()
        forcar_troca = self.var_indiv_troca.get()
        vpn = self.var_indiv_vpn.get()

        partes = nome_completo.split()
        if len(partes) < 2:
            messagebox.showwarning("Atenção", "Informe o nome completo do colaborador (nome e sobrenome).", parent=self)
            self.ent_indiv_nome.focus_set()
            return

        if not login:
            messagebox.showwarning("Atenção", "Informe o login (SAM) para a conta.", parent=self)
            self.ent_indiv_login.focus_set()
            return

        if len(login) > 20:
            messagebox.showwarning("Atenção", "O login não pode ter mais de 20 caracteres.", parent=self)
            self.ent_indiv_login.focus_set()
            return

        if not re.match(r'^[a-zA-Z0-9._\-]+$', login):
            messagebox.showwarning("Atenção", "O login contém caracteres inválidos. Use apenas letras, números, ponto, hífen ou underline.", parent=self)
            self.ent_indiv_login.focus_set()
            return

        if not email:
            email = f"{login}@umuarama.local"
            self.var_indiv_email.set(email)

        ou_dn = self._mapa_ous_legiveis.get(ou_display)
        if not ou_dn:
            if ou_display.upper().startswith("OU=") and "DC=" in ou_display.upper():
                ou_dn = ou_display
            else:
                messagebox.showwarning("Atenção", "Selecione uma Localização (OU) válida da lista.", parent=self)
                self.cb_indiv_ou.focus_set()
                return

        if not senha:
            senha = self._gerar_senha_padrao(nome_completo)
            self.var_indiv_senha.set(senha)

        msg_confirm = (
            f"Confirma a criação do usuário no Active Directory?\n\n"
            f"• Nome: {nome_completo}\n"
            f"• Login (SAM): {login}\n"
            f"• E-mail: {email}\n"
            f"• CPF: {cpf or '(não informado)'}\n"
            f"• Localização: {ou_display}\n"
            f"• Senha inicial: {senha}\n"
            f"• Forçar troca de senha: {'Sim' if forcar_troca else 'Não'}\n"
            f"• Acesso VPN: {'Sim' if vpn else 'Não'}"
        )
        if not messagebox.askyesno("Confirmar criação de usuário", msg_confirm, parent=self):
            return

        self._btn_indiv_criar.configure(state="disabled")
        self.var_status.set(f"Criando usuário {login} no Active Directory...")
        self._log(f"\n[PJ / INDIVIDUAL] Criando conta de '{nome_completo}' ({login})...", "titulo")

        primeiro = partes[0]
        sobrenome = " ".join(partes[1:])
        u_data = {
            "nome": primeiro,
            "sobrenome": sobrenome,
            "login": login,
            "email": email,
            "ou": ou_dn,
            "cpf": cpf,
            "vpn": vpn,
            "senha": senha,
        }

        user = self._sess_usuario
        user_pass = self._sess_senha
        dc = self._resolver_dc()

        def _run():
            script = gerar_ps1([u_data], dc, forcar_troca)
            tmp = tempfile.NamedTemporaryFile(suffix=".ps1", delete=False, mode="w", encoding="utf-8")
            tmp.write(script); tmp.close()
            env = os.environ.copy()
            env["UMU_USER"] = user
            env["UMU_PASS"] = user_pass

            resultado_status = "ERRO"
            resultado_detalhe = "Sem resposta do PowerShell."
            vpn_status = "NAO"

            try:
                proc = subprocess.Popen(
                    ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmp.name],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    env=env, text=True, encoding="utf-8", errors="replace",
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                for linha in proc.stdout:
                    linha = linha.rstrip()
                    if linha.startswith("RESULTADO|"):
                        partes_res = linha.split("|", 3)
                        if len(partes_res) >= 4:
                            _, sam_res, status_res, det_res = partes_res
                            resultado_status = status_res
                            resultado_detalhe = det_res
                    elif linha.startswith("VPN_OK|"):
                        vpn_status = "SIM"
                    elif linha.startswith("VPN_ERRO|"):
                        partes_vpn = linha.split("|", 2)
                        vpn_status = f"ERRO: {partes_vpn[2] if len(partes_vpn) > 2 else ''}"
                    elif linha == "FIM":
                        break
                proc.wait()
            finally:
                try:
                    os.unlink(tmp.name)
                except OSError:
                    pass

            def _finalizar():
                self._btn_indiv_criar.configure(state="normal")
                caminho_legivel = ou_para_caminho(ou_dn)

                if resultado_status in ("OK", "OK_ATIVADO", "OK_EXISTE_ATIVO", "OK_READMITIDO"):
                    if resultado_status == "OK":
                        msg_sucesso = f"Usuário '{login}' criado com sucesso no Active Directory!"
                        self._log(f"  ✔ [CRIADO] {nome_completo} ({login}) → {caminho_legivel}", "ok")
                    elif resultado_status == "OK_READMITIDO":
                        msg_sucesso = f"Colaborador com CPF '{cpf}' readmitido! Conta reativada, movida e senha redefinida."
                        self._log(f"  ✔ [READMITIDO] CPF {cpf} reativado para {login} → {caminho_legivel}", "ok")
                    elif resultado_status == "OK_ATIVADO":
                        msg_sucesso = f"A conta '{login}' já existia desativada e foi reativada com sucesso!"
                        self._log(f"  ✔ [REATIVADO] Conta {login} reativada → {caminho_legivel}", "ok")
                    else:
                        msg_sucesso = f"A conta '{login}' já existe e está ativa no AD."
                        self._log(f"  ℹ [JÁ ATIVO] Conta {login} já existe ativa no AD.", "info")

                    if vpn:
                        if vpn_status == "SIM":
                            self._log(f"  ✔ [VPN] Adicionado com sucesso ao grupo 'UsuariosVPN'.", "ok")
                        elif vpn_status.startswith("ERRO:"):
                            self._log(f"  ⚠ [VPN] Falha ao adicionar ao grupo VPN: {vpn_status[5:]}", "aviso")

                    self.var_status.set(f"OK: Conta {login} processada com sucesso.")
                    messagebox.showinfo("Sucesso", f"{msg_sucesso}\n\n"
                                                  f"• Login: {login}\n"
                                                  f"• Nome: {nome_completo}\n"
                                                  f"• E-mail: {email}\n"
                                                  f"• Senha inicial: {senha}\n"
                                                  f"• Localização: {caminho_legivel}", parent=self)
                elif resultado_status in ("JA_EXISTE", "CONFLITO_NOME"):
                    self._log(f"  ⚠ [CONFLITO LOGIN] O login '{login}' já existe no AD para outro usuário.", "aviso")
                    self.var_status.set(f"Conflito: login {login} já em uso.")
                    messagebox.showwarning("Login em uso",
                        f"O login '{login}' já está em uso por outro usuário no Active Directory.\n\n"
                        f"Por favor, altere o campo Login (SAM) para outro identificador.", parent=self)
                elif resultado_status == "CONFLITO_CPF":
                    self._log(f"  ⚠ [CONFLITO CPF] O CPF '{cpf}' já está associado a outra conta ativa no AD.", "aviso")
                    self.var_status.set(f"Conflito: CPF {cpf} já cadastrado em conta ativa.")
                    messagebox.showwarning("CPF já cadastrado",
                        f"O CPF '{cpf}' já está cadastrado em outra conta ativa no Active Directory.\n"
                        f"Consulte o colaborador na aba 'CONSULTA & GESTÃO AD'.", parent=self)
                else:
                    self._log(f"  ✘ [ERRO] Falha ao criar usuário '{login}': {resultado_detalhe}", "erro")
                    self.var_status.set(f"Erro ao criar usuário {login}.")
                    messagebox.showerror("Erro na criação",
                        f"Ocorreu um erro ao criar o usuário no Active Directory:\n\n{resultado_detalhe}", parent=self)

            self.after(0, _finalizar)

        threading.Thread(target=_run, daemon=True).start()

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
        # ── 1. Tela de login — valida contra o AD antes de abrir a janela principal
        login_dlg = DialogLogin()
        login_dlg.mainloop()

        if not login_dlg.autenticado:
            # Operador fechou ou cancelou o login — encerra sem abrir a app
            sys.exit(0)

        # ── 2. Sessao validada — abre a janela principal com as credenciais em memoria
        csv_preload = ""
        if "--csv-preload" in sys.argv:
            idx = sys.argv.index("--csv-preload")
            if idx + 1 < len(sys.argv):
                csv_preload = sys.argv[idx + 1]

        app = App(
            csv_preload=csv_preload,
            sessao_usuario=login_dlg.usuario,
            sessao_senha=login_dlg.senha,
            sessao_dc=login_dlg.dc,
        )
        app.mainloop()