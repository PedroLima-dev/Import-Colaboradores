"""
ImportarUsuariosUmuarama.py
Modos de uso:
  - Interface grafica (duplo clique ou via ExecutarImport.bat)
  - Linha de comando (chamado pelo EXECUTAR_IMPORT_COMPLETO.bat):
      python ImportarUsuariosUmuarama.py --modo-bat --csv <path> --usuario <user> --senha <pass> [--dc <dc>] [--sem-troca]
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import subprocess, threading, sys, os, csv, io, tempfile, unicodedata, chardet, argparse

# ── Importacao opcional do modulo de geracao de CSV ───────────────────────────
_DIR_SCRIPTS = os.path.dirname(os.path.abspath(__file__))
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
    ("HOLDING UAC", "RECURSOS HUMANOS"):            f"OU=RECURSOS HUMANOS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "DP"):                        f"OU=DP,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "FINANCEIRO"):                f"OU=FINANCEIRO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "FINANCIAMENTO"):             f"OU=FINANCIAMENTO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "MARKETING"):                 f"OU=MARKETING,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "PLANEJAMENTO"):              f"OU=PLANEJAMENTO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "COMPRAS"):                   f"OU=COMPRAS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "COMPRAS - OBRAS E INFRAESTRUTURA"): f"OU=COMPRAS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "DIRETORIA"):                 f"OU=DIRETORIA,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "SEGURANCA DO TRABALHO"):     f"OU=SEGURANCA DO TRABALHO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "REPASSE / SEMINOVOS"):       f"OU=REPASSE / SEMINOVOS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "VEICULOS USADOS"):           f"OU=REPASSE / SEMINOVOS,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "VENDAS DIRETAS"):            f"OU=VENDAS DIRETAS / MOBILIDADE / LICITACAO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "CENTRAL DE RELACIONAMENTO"): f"OU=CENTRAL DE RELACIONAMENTO,OU=HOLDING UAC,{BASE}",
    ("HOLDING UAC", "ADMINISTRACAO"):             f"OU=HOLDING UAC,{BASE}",
    ("HOLDING UAP", "FINANCEIRO"):    f"OU=FINANCEIRO,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "CONTABIL"):      f"OU=CONTABIL,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "CONTROLADORIA"): f"OU=CONTROLADORIA,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "FISCAL"):        f"OU=FISCAL,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "AUDITORIA"):     f"OU=AUDITORIA,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "MEIO AMBIENTE"): f"OU=MEIO AMBIENTE,OU=HOLDING UAP,{BASE}",
    ("HOLDING UAP", "ADMINISTRACAO"): f"OU=HOLDING UAP,{BASE}",
}

ALIAS_MARCA = {
    "VOLKS": "VOLKSWAGEN", "VW": "VOLKSWAGEN", "VOLKSWAGEN": "VOLKSWAGEN",
    "TOYOTA": "TOYOTA", "FIAT": "FIAT", "KIA": "KIA",
    "HARLEY": "HARLEY", "HARLEY-DAVIDSON": "HARLEY", "HD": "HD",
    "KTM": "KTM", "TRIUMPH": "TRIUMPH",
    "CITROEN": "CITROEN-PEUGEOT", "PEUGEOT": "CITROEN-PEUGEOT",
    "CITROEN-PEUGEOT": "CITROEN-PEUGEOT", "CRT": "CRT",
    "SEMINOVOS": "SEMINOVOS", "HOLDING UAC": "HOLDING UAC",
    "HOLDING UAP": "HOLDING UAP", "AGRO": "AGRO",
    "CORRETORA DE SEGUROS": "CORRETORA DE SEGUROS",
    "CORRETORA": "CORRETORA DE SEGUROS",
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
    if u_norm in ("HOLDING UAC", "HOLDING UAP"):
        return True
    d_norm = norm(depto)
    for termo in ("POS VENDAS", "POS-VENDAS", "VEICULOS NOVOS", "ADMINISTRACAO"):
        if d_norm.startswith(termo):
            return True
    return False

def parse_unidade(unidade_raw, departamento_raw):
    u = unidade_raw.strip()
    d = departamento_raw.strip()
    u_norm = norm(u)
    if u_norm in ("HOLDING UAC", "HOLDING UAP"):
        depto_norm = norm(d)
        if (u_norm, depto_norm) in OU_MAP:
            return u_norm, depto_norm
        depto_curto = norm(d.split("-")[0].split("–")[0].strip())
        return u_norm, depto_curto
    if " - " in u:
        partes = u.split(" - ", 1)
        marca_raw, cidade_raw = partes[0].strip(), partes[1].strip()
    else:
        marca_raw, cidade_raw = u, ""
    marca  = ALIAS_MARCA.get(norm(marca_raw), norm(marca_raw))
    cidade = ALIAS_CIDADE.get(norm(cidade_raw), norm(cidade_raw))
    return marca, cidade

def resolver_ou(marca, cidade):
    return OU_MAP.get((marca, cidade)) or OU_MAP.get((marca, ""))

def ou_para_caminho(ou_dn):
    """Converte DN de OU para formato legivel 'GRUPO UMUARAMA > CONCESSIONARIAS > TOYOTA > ITUMBIARA'."""
    partes = []
    for seg in ou_dn.split(","):
        seg = seg.strip()
        if seg.upper().startswith("OU="):
            partes.append(seg[3:].strip())
    partes.reverse()
    return " > ".join(partes) if partes else ou_dn

def gerar_login(nome_completo):
    partes = nome_completo.strip().split()
    if len(partes) < 2:
        return strip_accents(partes[0]).lower()[:15]
    login = f"{strip_accents(partes[0]).lower()}.{strip_accents(partes[-1]).lower()}"
    return login[:15]

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
        ou = resolver_ou(marca, cidade)
        if not ou:
            erros.append(
                f"Linha {i} ({nome_completo}): OU nao mapeada "
                f"[unidade='{unidade}' -> marca='{marca}' cidade='{cidade}']"
            )
            continue

        partes = nome_completo.strip().split()
        login     = gerar_login(nome_completo)
        primeiro  = partes[0]
        sobrenome = " ".join(partes[1:]) if len(partes) > 1 else ""
        if not email:
            email = f"{login}@umuarama.local"

        usuarios.append({
            "nome": primeiro, "sobrenome": sobrenome,
            "login": login, "email": email,
            "ou": ou, "cpf": cpf, "unidade": unidade,
            "vpn": pertence_grupo_vpn(unidade, departamento),
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
                print(f"  -> {nome} ({sam})", end="")
            elif linha.startswith("RESULTADO|"):
                partes = linha.split("|")
                sam = partes[1]
                status = partes[2]
                if status == "OK":
                    print(f"  [CRIADO]  →  {ou_para_caminho(partes[3])}"); c["ok"] += 1
                elif status == "OK_EXISTE_ATIVO":
                    print(f"  [JA EXISTE E ATIVO (PULADO)]  →  {ou_para_caminho(partes[3])}"); c["ok"] += 1
                elif status == "OK_ATIVADO":
                    print(f"  [ATIVADO]  →  {ou_para_caminho(partes[3])}"); c["ok"] += 1
                elif status in ("JA_EXISTE", "CONFLITO_NOME"):
                    print(f"  [JA EXISTE]"); c["existe"] += 1
                else:
                    detalhe = partes[3] if len(partes) > 3 else ""
                    print(f"  [ERRO] {detalhe}"); c["erro"] += 1
            elif linha.startswith("VPN_OK|"):
                print(" [VPN OK]", end="")
            elif linha.startswith("VPN_ERRO|"):
                _, sam, erro_msg = linha.split("|", 2)
                print(f" [VPN ERRO: {erro_msg}]", end="")
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
                if hasattr(os, "startfile"):
                    os.startfile(resultado)
                else:
                    import subprocess
                    subprocess.run(["open", resultado] if sys.platform == "darwin" else ["xdg-open", resultado])
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
        for u in usuarios:
            caminho = ou_para_caminho(u['ou'])
            self._log(f"  OK  {u['nome']} {u['sobrenome']:<25}  login: {u['login']:<18}  {caminho}", "ok")
        if erros:
            self._log("")
            for e in erros:
                self._log(f"  XX  {e}", "erro")
        self._preview_realizado = True

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
    def _dialog_novo_login(self, u_data, user, senha, dc, c, criados_log=None):
        """Abre dialogo modal, valida o login em tempo real e tenta criar o usuario
        inline (sem fechar o dialogo). Rexibe mensagem de erro se o novo login
        tambem existir no AD. So fecha quando: OK criado, operador pulou (0 / X).

        Retorna o status final: 'OK', 'PULADO' ou 'ERRO'.
        """
        import re as _re

        _SAM_MAX   = 20
        _SAM_REGEX = _re.compile(r'^[a-zA-Z0-9._\-]+$')

        def _validar_fmt(login):
            if login == "0":
                return None
            if not login:
                return "O login nao pode estar vazio."
            if len(login) > _SAM_MAX:
                return f"Login muito longo: {len(login)}/{_SAM_MAX} caracteres (max {_SAM_MAX})."
            if not _SAM_REGEX.match(login):
                return "Caracteres invalidos. Use apenas letras, numeros, ponto, hifen ou underline."
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
            dlg.title("Usuario duplicado")
            dlg.configure(bg=COR_BG)
            dlg.resizable(False, False)
            dlg.grab_set()
            dlg.focus_force()

            nome_completo   = f"{u_data['nome']} {u_data['sobrenome']}"

            tk.Label(dlg, text="\u26a0  Usuario duplicado", bg=COR_BG, fg=COR_AVS,
                     font=("Segoe UI", 12, "bold")).pack(padx=24, pady=(18, 4))
            tk.Label(dlg, text=nome_completo, bg=COR_BG, fg=COR_TXT,
                     font=("Segoe UI", 11, "bold")).pack(padx=24)
            tk.Label(dlg, text=f"Login em conflito:  {u_data['login']}",
                     bg=COR_BG, fg=COR_SUB, font=("Segoe UI", 9)).pack(padx=24, pady=(2, 10))

            conflict_name   = u_data.get("conflict_name")
            conflict_ou     = u_data.get("conflict_ou")
            conflict_status = u_data.get("conflict_status")
            if conflict_name:
                txt_conflict = f"{conflict_name} ({conflict_ou}) status: {conflict_status}"
                tk.Label(dlg, text=txt_conflict, bg=COR_BG, fg="#f38ba8",
                         font=("Segoe UI", 9, "bold"), wraplength=400).pack(padx=24, pady=(2, 10))

            tk.Label(dlg, text="Informe um novo login  ou  0  para pular:",
                     bg=COR_BG, fg=COR_TXT, font=("Segoe UI", 9)).pack(padx=24, pady=(0, 6))

            # ── Campo + contador ───────────────────────────────────────────────
            row_entry = tk.Frame(dlg, bg=COR_BG)
            row_entry.pack(padx=24, pady=(0, 2))
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
            var_mcor = [COR_SUB]
            lbl_msg  = tk.Label(dlg, textvariable=var_msg,
                                bg=COR_BG, font=("Segoe UI", 8),
                                wraplength=320, justify="left")
            lbl_msg.pack(padx=24, pady=(0, 6))

            def _set_msg(txt, cor):
                var_msg.set(txt)
                lbl_msg.configure(fg=cor)

            # ── Botoes ─────────────────────────────────────────────────────────
            bf = tk.Frame(dlg, bg=COR_BG)
            bf.pack(pady=10)
            btn_ok = tk.Button(bf, text="Confirmar", bg=COR_ACC, fg="white",
                               font=("Segoe UI", 10, "bold"), relief="flat", padx=12, pady=5)
            btn_ok.pack(side="left", padx=8)
            btn_pular = tk.Button(bf, text="Pular (0)", bg=COR_CARD, fg=COR_SUB,
                                  font=("Segoe UI", 10), relief="flat", padx=12, pady=5)
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
                    btn_ok.configure(state="disabled", bg="#44415a", text="Confirmar")
                else:
                    _set_msg("", COR_SUB)
                    btn_ok.configure(state="normal", bg=COR_ACC, text="Confirmar")

            var_entrada.trace_add("write", _on_change)
            _on_change()

            # ── Confirmar: tenta criar inline sem fechar o dialogo ────────────
            def _tentar_criar(_tk_ev=None):
                val = var_entrada.get().strip()
                if val == "0":
                    resultado[0] = "PULADO"
                    dlg.destroy()
                    return
                if not val or _validar_fmt(val):
                    return   # formato invalido — nao age

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
                    status = self._criar_usuario_individual(u_retry, user, senha, dc, c, criados_log)

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
                            btn_ok.configure(state="disabled", bg="#44415a", text="Confirmar")
                            entry.focus_set()
                        else:  # ERRO
                            resultado[0] = "ERRO"
                            dlg.destroy()

                    self.after(0, _atualizar)

                threading.Thread(target=_executar_ps1, daemon=True).start()

            entry.bind("<Return>", _tentar_criar)
            btn_ok.configure(command=_tentar_criar)
            entry.focus_set()

            def _ao_fechar():
                resultado[0] = "PULADO"
                self._log(f"     Usuario pulado pelo operador: {nome_completo} ({u_data['login']})", "aviso")
                dlg.destroy()

            btn_pular.configure(command=lambda: (
                resultado.__setitem__(0, "PULADO"),
                self._log(f"     Usuario pulado pelo operador: {nome_completo} ({u_data['login']})", "aviso"),
                dlg.destroy()
            ))

            dlg.protocol("WM_DELETE_WINDOW", _ao_fechar)

            def _aguardar():
                if dlg.winfo_exists():
                    self.after(50, _aguardar)
                else:
                    event.set()

            self.after(50, _aguardar)

        self.after(0, _mostrar)
        event.wait()
        return resultado[0]

    def _criar_usuario_individual(self, u_data, user, senha, dc, c, criados_log=None):
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
                        if criados_log is not None:
                            nome_exib = f"{u_data.get('nome','')} {u_data.get('sobrenome','')}".strip()
                            criados_log.append({
                                "nome": nome_exib,
                                "login": sam,
                                "ou": detalhe,
                                "via": "conflito",
                                "vpn": vpn_info
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

    # ── Log de importacao em arquivo .txt ─────────────────────────────────────
    @staticmethod
    def _salvar_log_txt(dc, criados, pulados, c, csv_origem):
        """Gera um arquivo .txt em logs/ com o resumo da importacao."""
        import datetime
        pasta_logs = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(pasta_logs, exist_ok=True)
        agora = datetime.datetime.now()
        nome_arq = agora.strftime("importacao-%d%m%Y-%H%M%S.txt")
        caminho_arq = os.path.join(pasta_logs, nome_arq)

        sep  = "=" * 70
        sep2 = "-" * 70

        linhas = [
            sep,
            f"  IMPORTACAO AD - GRUPO UMUARAMA",
            f"  Data/Hora : {agora.strftime('%d/%m/%Y %H:%M:%S')}",
            f"  Servidor  : {dc}",
            f"  Origem CSV: {csv_origem}",
            sep,
            "",
        ]

        if criados:
            linhas += [
                f"  USUARIOS CRIADOS ({len(criados)})",
                sep2,
                f"  {'NOME COMPLETO':<40}  {'LOGIN':<22}  {'VPN':<12}  LOCALIZACAO",
                sep2,
            ]
            for u in criados:
                caminho_ou = ou_para_caminho(u['ou'])
                via = "  [via conflito]" if u.get('via') == 'conflito' else ""
                vpn_val = u.get('vpn', 'NAO')
                linhas.append(f"  {u['nome']:<40}  {u['login']:<22}  {vpn_val:<12}  {caminho_ou}{via}")
            linhas.append("")

        if pulados:
            linhas += [
                f"  USUARIOS PULADOS / JA EXISTIAM ({len(pulados)})",
                sep2,
            ]
            for u in pulados:
                linhas.append(f"  {u['nome']:<40}  {u['login']}")
            linhas.append("")

        linhas += [
            sep,
            f"  RESUMO:  {c['ok']} criados   {c['existe']} ja existiam   {c['erro']} erros",
            sep,
            "",
        ]

        with open(caminho_arq, "w", encoding="utf-8") as f:
            f.write("\n".join(linhas))

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
        criados_log    = []    # lista de usuarios criados para o arquivo de log
        pulados_log    = []    # lista de usuarios pulados/ja existiam
        vpn_status     = {}    # mapeia login -> status da vpn ("SIM", "NAO", "ERRO (msg)")
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
                    if status in ("OK", "OK_EXISTE_ATIVO", "OK_ATIVADO"):
                        caminho = ou_para_caminho(detalhe)
                        if status == "OK_EXISTE_ATIVO":
                            self._log(f"     ✔  JA EXISTE E ATIVO (CRIACAO PULADA)  \u2192  {caminho}", "ok")
                        elif status == "OK_ATIVADO":
                            self._log(f"     ✔  ATIVADO E ATIVO (CRIACAO PULADA)  \u2192  {caminho}", "ok")
                        else:
                            self._log(f"     OK CRIADO  \u2192  {caminho}", "ok")
                        c["ok"] += 1
                        
                        v_stat = vpn_status.get(sam, "NAO")
                        if v_stat == "SIM":
                            self._log(f"     → Adicionado ao grupo VPN (UsuariosVPN)", "ok")
                        elif v_stat.startswith("ERRO:"):
                            self._log(f"     ⚠ Falha ao adicionar ao grupo VPN: {v_stat[5:].strip()}", "aviso")
                        else:
                            self._log(f"     → Sem VPN (Não elegível)", "info")

                        nome_exib = f"{u_atual['nome']} {u_atual['sobrenome']}" if u_atual else sam
                        criados_log.append({
                            "nome": nome_exib,
                            "login": sam,
                            "ou": detalhe,
                            "via": "ja_existia" if status != "OK" else "",
                            "vpn": v_stat
                        })
                    elif status == "JA_EXISTE":
                        u_conflict = next((u for u in usuarios if u["login"] == sam), None)
                        if u_conflict:
                            ja_existe_users.append(u_conflict)
                            pulados_log.append({"nome": f"{u_conflict['nome']} {u_conflict['sobrenome']}",
                                                "login": sam})
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
                            pulados_log.append({"nome": f"{u_conflict['nome']} {u_conflict['sobrenome']}",
                                                "login": sam})
                        self._log(f"     CONFLITO DE LOGIN (Nome diferente no AD) - sera tratado apos o batch...", "aviso")
                        c["existe"] += 1
                    else:
                        self._log(f"     ERRO - {detalhe}", "erro"); c["erro"] += 1
                elif linha.startswith("VPN_OK|"):
                    _, sam = linha.split("|", 1)
                    vpn_status[sam] = "SIM"
                elif linha.startswith("VPN_ERRO|"):
                    _, sam, erro_msg = linha.split("|", 2)
                    vpn_status[sam] = f"ERRO: {erro_msg}"
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
                # O dialogo gerencia o retry internamente; passa criados_log para registrar OK
                res = self._dialog_novo_login(u_conflict, user, senha, dc, c, criados_log)
                # Se o conflito foi resolvido, remove da lista de pulados
                if res == "OK":
                    pulados_log = [p for p in pulados_log if p["login"] != u_conflict["login"]]

        self._log("")
        self._log("=" * 72, "titulo")
        self._log(f"  CONCLUIDO  OK {c['ok']} criados   {c['existe']} ja existiam   XX {c['erro']} erros", "titulo")
        self._log("=" * 72, "titulo")
        self.var_status.set(f"Concluido - {c['ok']} criados, {c['existe']} ja existiam, {c['erro']} erros.")

        # ── Gera arquivo de log .txt ───────────────────────────────────────────
        if c["ok"] > 0 or c["existe"] > 0 or c["erro"] > 0:
            try:
                arq_log = self._salvar_log_txt(dc, criados_log, pulados_log, c, csv_origem)
                self._log("")
                self._log(f"  Log salvo em: {arq_log}", "ok")
            except Exception as e_log:
                self._log(f"  Aviso: nao foi possivel salvar o log - {e_log}", "aviso")

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