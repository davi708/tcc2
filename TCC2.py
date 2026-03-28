import io
import json
import math
import secrets
import ssl
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import certifi
except ModuleNotFoundError:
    certifi = None

try:
    import tomllib
except ModuleNotFoundError:
    tomllib = None

st.set_page_config(page_title="Dimensionamento Eletrico Residencial", layout="wide")

CATEGORIAS_DEMANDA = {
    "b": "Chuveiro, torneira eletrica, aquecedor de passagem, ferro eletrico",
    "c": "Boiler / aquecedor central",
    "d": "Secadora, forno eletrico, lava-loucas, micro-ondas",
    "e": "Fogao eletrico",
    "f": "Ar-condicionado tipo janela",
    "g": "Motor / maquina de solda a motor",
    "h": "Equipamento especial",
    "i": "Hidromassagem",
    "x": "Sem categoria GED-13 (considerar FD = 1)",
}

PADRAO_ENTRADA_TABELA_1A = {
    "A1": {
        "tipo_caixa": "II",
        "disjuntor_a": 32,
        "medida_eletroduto": "32 (1)",
    },
    "A2": {
        "tipo_caixa": "II",
        "disjuntor_a": 63,
        "medida_eletroduto": "32 (1)",
    },
    "B1": {
        "tipo_caixa": "II",
        "disjuntor_a": 63,
        "medida_eletroduto": "40 (1 1/4)",
    },
    "B2": {
        "tipo_caixa": "II",
        "disjuntor_a": 80,
        "medida_eletroduto": "40 (1 1/4)",
    },
    "C1": {
        "tipo_caixa": "III",
        "disjuntor_a": 63,
        "medida_eletroduto": "40 (1 1/4)",
    },
    "C2": {
        "tipo_caixa": "III",
        "disjuntor_a": 80,
        "medida_eletroduto": "40 (1 1/4)",
    },
    "C3": {
        "tipo_caixa": "III",
        "disjuntor_a": 100,
        "medida_eletroduto": "40 (1 1/4)",
    },
    "C4": {
        "tipo_caixa": "III",
        "disjuntor_a": 125,
        "medida_eletroduto": "50 (1 1/2)",
    },
    "C5": {
        "tipo_caixa": "H",
        "disjuntor_a": 150,
        "medida_eletroduto": "50 (1 1/2)",
    },
    "C6": {
        "tipo_caixa": "H",
        "disjuntor_a": 200,
        "medida_eletroduto": "60 (2)",
    },
}


# -------------------------------
# FORMATACAO
# -------------------------------
def formatar_numero_br(valor: float, casas: int = 1) -> str:
    texto = f"{valor:,.{casas}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


# -------------------------------
# NORMALIZACAO
# -------------------------------
def normalizar_ambiente(ambiente: str) -> str:
    amb = ambiente.strip().lower()
    trocas = {
        "a": ["a", "á", "à", "ã", "â"],
        "e": ["e", "é", "ê"],
        "i": ["i", "í"],
        "o": ["o", "ó", "ô", "õ"],
        "u": ["u", "ú"],
        "c": ["c", "ç"],
    }
    for destino, origens in trocas.items():
        for origem in origens[1:]:
            amb = amb.replace(origem, destino)

    aliases = {
        "suite": "suite",
        "suíte": "suite",
        "dormitorio": "quarto",
        "dormitório": "quarto",
        "sala de estar": "sala",
        "sala de jantar": "sala",
        "area de servico": "area_servico",
        "área de servico": "area_servico",
        "area de serviço": "area_servico",
        "área de serviço": "area_servico",
        "copa-cozinha": "copa_cozinha",
        "copa cozinha": "copa_cozinha",
        "bwc": "banheiro",
        "wc": "banheiro",
        "hall de escadaria": "hall",
        "hall escadaria": "hall",
        "casa de maquinas": "casa_maquinas",
        "casa de máquinas": "casa_maquinas",
        "sala de bombas": "sala_bombas",
    }
    return aliases.get(amb, amb)


# -------------------------------
# NBR 5410 - ILUMINACAO
# -------------------------------
def calcular_iluminacao(area: float) -> Dict[str, int]:
    pontos_minimos = 1

    if area <= 6:
        potencia_va = 100
    else:
        acrescimos = math.floor((area - 6) / 4)
        potencia_va = 100 + acrescimos * 60

    return {
        "pontos_minimos": pontos_minimos,
        "potencia_va": potencia_va,
    }


# -------------------------------
# NBR 5410 - TUG
# -------------------------------
def calcular_tug(
    area: float,
    perimetro: float,
    ambiente: str,
    bancadas_validas: int = 0,
) -> Dict[str, object]:
    amb = normalizar_ambiente(ambiente)

    if amb == "banheiro":
        pontos = 1
        potencias = [600]

    elif amb in {"cozinha", "copa", "copa_cozinha", "area_servico", "lavanderia"}:
        pontos_perimetro = math.ceil(perimetro / 3.5)
        pontos = max(pontos_perimetro, bancadas_validas)
        potencias = [600] * min(pontos, 3) + [100] * max(0, pontos - 3)

    elif amb in {"garagem", "varanda", "sotao", "subsolo", "hall", "casa_maquinas", "sala_bombas", "barrilete"}:
        pontos = 1
        potencias = [100]

    else:
        if area <= 6:
            pontos = 1
        else:
            pontos = math.ceil(perimetro / 5)
        potencias = [100] * pontos

    return {
        "pontos": pontos,
        "potencias_va": potencias,
        "potencia_total_va": sum(potencias),
    }


# -------------------------------
# NBR 5410 - TUE
# -------------------------------
def calcular_tue(equipamentos: List[Dict[str, object]]) -> Dict[str, object]:
    total_w = 0.0
    descricoes = []

    for eq in equipamentos:
        nome = str(eq.get("nome", "")).strip() or "Equipamento"
        potencia = float(eq.get("potencia_w", 0) or 0)
        total_w += potencia
        descricoes.append(f"{nome} ({potencia:.0f} W)")

    return {
        "descricao": " / ".join(descricoes) if descricoes else "-",
        "potencia_total_w": total_w,
    }


# -------------------------------
# TUG - FORMATACAO DO PONTO
# -------------------------------
def formatar_sponto_tug(potencias: List[int]) -> str:
    if not potencias:
        return "-"

    contagem = Counter(potencias)
    partes = []
    for va in sorted(contagem.keys(), reverse=True):
        partes.append(f"{contagem[va]} de {va}")
    return " e ".join(partes)


# -------------------------------
# CPFL / GED-13 - DEMANDA (versao simplificada didatica)
# -------------------------------
def fd_iluminacao_tug(carga_kw: float) -> float:
    if carga_kw <= 1:
        return 0.86
    if carga_kw <= 2:
        return 0.75
    if carga_kw <= 3:
        return 0.66
    if carga_kw <= 4:
        return 0.59
    if carga_kw <= 5:
        return 0.52
    if carga_kw <= 6:
        return 0.45
    if carga_kw <= 7:
        return 0.40
    if carga_kw <= 8:
        return 0.35
    if carga_kw <= 9:
        return 0.31
    if carga_kw <= 10:
        return 0.27
    return 0.24


FD_TABELA_B = {
    1: 1.00,
    2: 1.00,
    3: 0.84,
    4: 0.76,
    5: 0.70,
    6: 0.65,
    7: 0.60,
    8: 0.57,
    9: 0.54,
    10: 0.52,
    11: 0.49,
    12: 0.48,
    13: 0.46,
    14: 0.45,
    15: 0.44,
    16: 0.43,
    17: 0.42,
    18: 0.41,
    19: 0.40,
    20: 0.40,
    21: 0.39,
    22: 0.39,
    23: 0.39,
    24: 0.38,
    25: 0.38,
}


def fd_categoria_b(qtd: int) -> float:
    return FD_TABELA_B.get(qtd, 0.38)



def fd_categoria_c(qtd: int) -> float:
    if qtd <= 1:
        return 1.00
    if qtd == 2:
        return 0.72
    return 0.62



def fd_categoria_d(qtd: int) -> float:
    if qtd <= 1:
        return 1.00
    if 2 <= qtd <= 4:
        return 0.70
    if 5 <= qtd <= 6:
        return 0.60
    return 0.50



def fd_categoria_e(qtd: int) -> float:
    tabela = {
        1: 1.00,
        2: 0.60,
        3: 0.48,
        4: 0.40,
        5: 0.37,
        6: 0.35,
        7: 0.33,
        8: 0.32,
        9: 0.31,
    }
    if qtd in tabela:
        return tabela[qtd]
    if 10 <= qtd <= 11:
        return 0.30
    if 12 <= qtd <= 15:
        return 0.28
    return 0.26



def fd_categoria_f_residencial(qtd: int) -> float:
    return 1.00



def demanda_maiores_primeiro(potencias_w: List[float], fatores_maiores: List[float], fator_demais: float) -> float:
    if not potencias_w:
        return 0.0

    potencias = sorted((float(p) for p in potencias_w), reverse=True)
    demanda = 0.0

    for i, potencia in enumerate(potencias):
        if i < len(fatores_maiores):
            demanda += potencia * fatores_maiores[i]
        else:
            demanda += potencia * fator_demais

    return demanda



def calcular_demanda_cpfl_simplificada(
    carga_iluminacao_va: float,
    carga_tug_va: float,
    equipamentos_tue: List[Dict[str, object]],
) -> Tuple[pd.DataFrame, float]:
    linhas = []

    carga_a_kw = (carga_iluminacao_va + carga_tug_va) / 1000
    fd_a = fd_iluminacao_tug(carga_a_kw)
    demanda_a_w = (carga_iluminacao_va + carga_tug_va) * fd_a
    linhas.append(
        {
            "Categoria": "a",
            "Descricao": "Iluminacao + TUG",
            "Carga instalada": f"{carga_iluminacao_va + carga_tug_va:.0f} VA",
            "FD": fd_a,
            "Demanda": f"{demanda_a_w:.0f} W",
            "Equipamentos": "-",
        }
    )

    grupos = defaultdict(list)
    for eq in equipamentos_tue:
        categoria = str(eq.get("categoria_demanda", "x"))
        grupos[categoria].append(eq)

    for categoria, itens in grupos.items():
        potencias = [float(item.get("potencia_w", 0) or 0) for item in itens]
        nomes = " / ".join(str(item.get("nome", "Equipamento")) for item in itens)
        qtd = len(itens)

        if categoria == "b":
            fd = fd_categoria_b(qtd)
            demanda_w = sum(potencias) * fd
            desc = "Chuveiro / torneira / aquecedor de passagem / ferro"
        elif categoria == "c":
            fd = fd_categoria_c(qtd)
            demanda_w = sum(potencias) * fd
            desc = "Boiler / aquecedor central"
        elif categoria == "d":
            fd = fd_categoria_d(qtd)
            demanda_w = sum(potencias) * fd
            desc = "Secadora / forno eletrico / lava-loucas / micro-ondas"
        elif categoria == "e":
            fd = fd_categoria_e(qtd)
            demanda_w = sum(potencias) * fd
            desc = "Fogao eletrico"
        elif categoria == "f":
            fd = fd_categoria_f_residencial(qtd)
            demanda_w = sum(potencias) * fd
            desc = "Ar-condicionado tipo janela (uso residencial)"
        elif categoria == "g":
            fd = "maiores"
            demanda_w = demanda_maiores_primeiro(potencias, [1.00, 0.90, 0.80, 0.80, 0.80], 0.70)
            desc = "Motores / solda a motor"
        elif categoria == "h":
            fd = 1.00
            demanda_w = sum(potencias)
            desc = "Equipamentos especiais (simplificado)"
        elif categoria == "i":
            fd = "maiores"
            demanda_w = demanda_maiores_primeiro(potencias, [1.00, 0.90, 0.80, 0.80, 0.80], 0.70)
            desc = "Hidromassagem"
        else:
            fd = 1.00
            demanda_w = sum(potencias)
            desc = "Sem categoria GED-13"

        linhas.append(
            {
                "Categoria": categoria,
                "Descricao": desc,
                "Carga instalada": f"{sum(potencias):.0f} W",
                "FD": fd,
                "Demanda": f"{demanda_w:.0f} W",
                "Equipamentos": nomes,
            }
        )

    df = pd.DataFrame(linhas)
    total_demanda_w = float(demanda_a_w)
    for _, row in df.iloc[1:].iterrows():
        total_demanda_w += float(str(row["Demanda"]).replace(" W", ""))

    return df, total_demanda_w


# -------------------------------
# PADRAO DE ENTRADA - TABELA 1A GED-13
# -------------------------------
def determinar_categoria_padrao_entrada(carga_instalada_w: float, demanda_total_w: float) -> Tuple[str, str]:
    carga_instalada_kw = carga_instalada_w / 1000
    demanda_total_kva = demanda_total_w / 1000

    if carga_instalada_kw <= 6:
        return "A1", "carga instalada"
    if carga_instalada_kw <= 12:
        return "A2", "carga instalada"
    if carga_instalada_kw <= 18:
        return "B1", "carga instalada"
    if carga_instalada_kw <= 25:
        return "B2", "carga instalada"

    if demanda_total_kva <= 23:
        return "C1", "demanda"
    if demanda_total_kva <= 30:
        return "C2", "demanda"
    if demanda_total_kva <= 38:
        return "C3", "demanda"
    if demanda_total_kva <= 47:
        return "C4", "demanda"
    if demanda_total_kva <= 57:
        return "C5", "demanda"
    if demanda_total_kva <= 76:
        return "C6", "demanda"

    return "CONSULTAR GED-13", "consultar GED-13"



def resolver_fase_padrao(categoria: str, fase_escolhida: str) -> str:
    if fase_escolhida != "Automatico":
        return fase_escolhida

    if categoria in {"A1", "A2", "B1", "B2"}:
        return "Monofasico"
    if categoria in {"C1", "C2", "C3", "C4", "C5", "C6"}:
        return "Trifasico"
    return "Consultar GED-13"



def calcular_padrao_entrada(
    carga_instalada_w: float,
    demanda_total_w: float,
    fase_escolhida: str,
) -> Dict[str, str]:
    categoria, criterio = determinar_categoria_padrao_entrada(carga_instalada_w, demanda_total_w)
    fase = resolver_fase_padrao(categoria, fase_escolhida)
    caracteristicas = PADRAO_ENTRADA_TABELA_1A.get(categoria, {})

    return {
        "Fase": fase,
        "Categoria": categoria,
        "Carga Instalada": f"{formatar_numero_br(carga_instalada_w, 1)} W",
        "Tipo de Caixa": caracteristicas.get("tipo_caixa", "Consultar GED-13"),
        "Disjuntor": str(caracteristicas.get("disjuntor_a", "Consultar GED-13")),
        "Medida do Eletroduto": caracteristicas.get("medida_eletroduto", "Consultar GED-13"),
        "Criterio da Categoria": criterio,
        "Demanda Total Simplificada": f"{formatar_numero_br(demanda_total_w, 1)} W",
    }


# -------------------------------
# EXPORTACAO XLSX
# -------------------------------
def ajustar_largura_colunas(worksheet) -> None:
    for coluna in worksheet.columns:
        maximo = 0
        indice_coluna = coluna[0].column
        letra = get_column_letter(indice_coluna)
        for celula in coluna:
            valor = "" if celula.value is None else str(celula.value)
            if len(valor) > maximo:
                maximo = len(valor)
        worksheet.column_dimensions[letra].width = min(maximo + 3, 60)



def aplicar_estilo_planilha(writer: pd.ExcelWriter) -> None:
    workbook = writer.book
    cabecalho_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    cabecalho_font = Font(color="FFFFFF", bold=True)

    for worksheet in workbook.worksheets:
        if worksheet.max_row >= 1:
            for celula in worksheet[1]:
                celula.fill = cabecalho_fill
                celula.font = cabecalho_font
        worksheet.freeze_panes = "A2"
        ajustar_largura_colunas(worksheet)



def gerar_excel_bytes(
    df_resultados: pd.DataFrame,
    df_demanda: pd.DataFrame,
    df_padrao: pd.DataFrame,
    df_resumo: pd.DataFrame,
) -> bytes:
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_resultados.to_excel(writer, sheet_name="Tabela de Cargas", index=False)
        df_demanda.to_excel(writer, sheet_name="Demanda CPFL", index=False)
        df_padrao.to_excel(writer, sheet_name="Padrao de Entrada", index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
        aplicar_estilo_planilha(writer)

    buffer.seek(0)
    return buffer.getvalue()


ARQUIVO_SECRETS = Path(__file__).with_name("secrets.toml")
ARQUIVO_OAUTH_ESTADO = Path(__file__).with_name("oauth_state.json")


# -------------------------------
# AUTENTICACAO
# -------------------------------
def carregar_config_autenticacao() -> Dict[str, object]:
    config: Dict[str, object] = {}

    try:
        config.update(dict(st.secrets))
    except Exception:
        pass

    if ARQUIVO_SECRETS.exists() and tomllib is not None:
        with ARQUIVO_SECRETS.open("rb") as arquivo:
            config_arquivo = tomllib.load(arquivo)
        for chave, valor in config_arquivo.items():
            config[chave] = valor

    return config


def obter_query_params() -> Dict[str, str]:
    if hasattr(st, "query_params"):
        params = {}
        for chave in st.query_params.keys():
            valor = st.query_params[chave]
            if isinstance(valor, list):
                params[chave] = valor[0] if valor else ""
            else:
                params[chave] = valor
        return params

    params_legacy = st.experimental_get_query_params()
    return {chave: valores[0] for chave, valores in params_legacy.items() if valores}


def limpar_query_params() -> None:
    if hasattr(st, "query_params"):
        for chave in list(st.query_params.keys()):
            del st.query_params[chave]
    else:
        st.experimental_set_query_params()


def salvar_estado_oauth(state: str, provider: str) -> None:
    dados = {"state": state, "provider": provider}
    ARQUIVO_OAUTH_ESTADO.write_text(json.dumps(dados), encoding="utf-8")


def carregar_estado_oauth() -> Dict[str, str]:
    if not ARQUIVO_OAUTH_ESTADO.exists():
        return {}

    try:
        conteudo = ARQUIVO_OAUTH_ESTADO.read_text(encoding="utf-8")
        dados = json.loads(conteudo)
        if isinstance(dados, dict):
            return {str(chave): str(valor) for chave, valor in dados.items()}
    except Exception:
        return {}

    return {}


def limpar_estado_oauth() -> None:
    try:
        if ARQUIVO_OAUTH_ESTADO.exists():
            ARQUIVO_OAUTH_ESTADO.unlink()
    except Exception:
        pass


def carregar_usuarios_locais(config: Dict[str, object]) -> List[Dict[str, str]]:
    usuarios: List[Dict[str, str]] = []

    json_usuarios = config.get("LOCAL_USERS_JSON")
    if json_usuarios:
        try:
            dados = json.loads(str(json_usuarios))
            if isinstance(dados, list):
                for item in dados:
                    if isinstance(item, dict):
                        usuarios.append({str(k): str(v) for k, v in item.items()})
        except json.JSONDecodeError:
            pass

    email = str(config.get("LOCAL_LOGIN_EMAIL", "")).strip()
    senha = str(config.get("LOCAL_LOGIN_PASSWORD", "")).strip()
    nome = str(config.get("LOCAL_LOGIN_NAME", "Usuario local")).strip() or "Usuario local"
    if email and senha:
        usuarios.append({
            "email": email,
            "password": senha,
            "name": nome,
        })

    return usuarios


def autenticar_login_local(email: str, senha: str, usuarios: List[Dict[str, str]]) -> Dict[str, str] | None:
    email_normalizado = email.strip().lower()

    for usuario in usuarios:
        email_usuario = str(usuario.get("email", "")).strip().lower()
        senha_usuario = str(usuario.get("password", "")).strip()
        if email_usuario == email_normalizado and senha_usuario == senha:
            return {
                "provider": "local",
                "email": email_usuario,
                "name": str(usuario.get("name", email_usuario)).strip() or email_usuario,
            }

    return None


def obter_config_oauth(provider: str, config: Dict[str, object]) -> Dict[str, object] | None:
    if provider == "google":
        client_id = str(config.get("GOOGLE_CLIENT_ID", "")).strip()
        client_secret = str(config.get("GOOGLE_CLIENT_SECRET", "")).strip()
        redirect_uri = str(config.get("GOOGLE_REDIRECT_URI", "")).strip()
        if not client_id or not client_secret or not redirect_uri or "..." in redirect_uri:
            return None
        return {
            "provider": "google",
            "label": "Google",
            "client_id": client_id,
            "client_secret": client_secret,
            "redirect_uri": redirect_uri,
            "authorize_url": "https://accounts.google.com/o/oauth2/v2/auth",
            "token_url": "https://oauth2.googleapis.com/token",
            "userinfo_url": "https://openidconnect.googleapis.com/v1/userinfo",
            "scope": "openid email profile",
            "extra_params": {"prompt": "select_account"},
        }


    return None


def montar_url_autorizacao(provider: str, config: Dict[str, object]) -> str | None:
    oauth = obter_config_oauth(provider, config)
    if oauth is None:
        return None

    state = f"{provider}:{secrets.token_urlsafe(24)}"
    st.session_state["oauth_state"] = state
    salvar_estado_oauth(state, provider)

    params = {
        "client_id": oauth["client_id"],
        "redirect_uri": oauth["redirect_uri"],
        "response_type": "code",
        "scope": oauth["scope"],
        "state": state,
    }
    params.update(oauth.get("extra_params", {}))

    return f"{oauth['authorize_url']}?{urllib.parse.urlencode(params)}"


def criar_contexto_ssl() -> ssl.SSLContext:
    if certifi is not None:
        return ssl.create_default_context(cafile=certifi.where())
    return ssl.create_default_context()


def requisicao_json(url: str, data: Dict[str, str] | None = None, headers: Dict[str, str] | None = None) -> Dict[str, object]:
    headers = headers or {}
    contexto_ssl = criar_contexto_ssl()

    if data is None:
        request = urllib.request.Request(url, headers=headers)
    else:
        body = urllib.parse.urlencode(data).encode("utf-8")
        headers = {"Content-Type": "application/x-www-form-urlencoded", **headers}
        request = urllib.request.Request(url, data=body, headers=headers)

    with urllib.request.urlopen(request, timeout=20, context=contexto_ssl) as resposta:
        return json.loads(resposta.read().decode("utf-8"))


def trocar_code_por_token(provider: str, code: str, config: Dict[str, object]) -> Dict[str, object]:
    oauth = obter_config_oauth(provider, config)
    if oauth is None:
        raise ValueError(f"Configuracao de {provider} nao encontrada.")

    payload = {
        "code": code,
        "client_id": str(oauth["client_id"]),
        "client_secret": str(oauth["client_secret"]),
        "redirect_uri": str(oauth["redirect_uri"]),
        "grant_type": "authorization_code",
    }
    return requisicao_json(str(oauth["token_url"]), data=payload)


def decodificar_id_token_sem_validacao(id_token: str) -> Dict[str, object]:
    partes = id_token.split(".")
    if len(partes) < 2:
        return {}

    payload = partes[1]
    padding = "=" * (-len(payload) % 4)
    try:
        conteudo = urllib.parse.unquote_to_bytes(payload + padding)
    except Exception:
        conteudo = (payload + padding).encode("utf-8")

    try:
        import base64
        bruto = base64.urlsafe_b64decode(payload + padding)
        return json.loads(bruto.decode("utf-8"))
    except Exception:
        try:
            return json.loads(conteudo.decode("utf-8"))
        except Exception:
            return {}


def obter_usuario_oauth(provider: str, token: Dict[str, object], config: Dict[str, object]) -> Dict[str, str]:
    oauth = obter_config_oauth(provider, config)
    if oauth is None:
        raise ValueError(f"Configuracao de {provider} nao encontrada.")

    access_token = str(token.get("access_token", "")).strip()
    if access_token:
        dados = requisicao_json(
            str(oauth["userinfo_url"]),
            headers={"Authorization": f"Bearer {access_token}"},
        )
        return {
            "provider": provider,
            "email": str(dados.get("email") or dados.get("preferred_username") or "").strip(),
            "name": str(dados.get("name") or dados.get("given_name") or dados.get("preferred_username") or "Usuario").strip(),
        }

    id_token = str(token.get("id_token", "")).strip()
    dados = decodificar_id_token_sem_validacao(id_token)
    return {
        "provider": provider,
        "email": str(dados.get("email") or dados.get("preferred_username") or "").strip(),
        "name": str(dados.get("name") or dados.get("given_name") or dados.get("preferred_username") or "Usuario").strip(),
    }


def processar_callback_oauth(config: Dict[str, object]) -> None:
    params = obter_query_params()
    code = params.get("code", "")
    state = params.get("state", "")
    erro = params.get("error", "")

    if erro:
        st.error(f"Falha no login SSO: {erro}.")
        limpar_query_params()
        return

    if not code or not state:
        return

    estado_salvo = carregar_estado_oauth()
    state_esperado = st.session_state.get("oauth_state", "") or estado_salvo.get("state", "")
    if not state_esperado or state != state_esperado:
        st.error("Nao foi possivel validar a resposta do provedor de login.")
        limpar_query_params()
        limpar_estado_oauth()
        return

    provider = estado_salvo.get("provider", "") or state.split(":", 1)[0]

    try:
        token = trocar_code_por_token(provider, code, config)
        usuario = obter_usuario_oauth(provider, token, config)
    except Exception as exc:
        st.error(f"Falha ao concluir o login com {provider.title()}: {exc}")
        limpar_query_params()
        limpar_estado_oauth()
        return

    if not usuario.get("email"):
        st.error("O provedor retornou o login, mas sem um e-mail utilizavel.")
        limpar_query_params()
        limpar_estado_oauth()
        return

    st.session_state["auth_user"] = usuario
    limpar_query_params()
    limpar_estado_oauth()
    st.rerun()


def usuario_autenticado() -> bool:
    return bool(st.session_state.get("auth_user"))


def sair() -> None:
    st.session_state.pop("auth_user", None)
    st.session_state.pop("oauth_state", None)
    limpar_query_params()
    limpar_estado_oauth()
    st.rerun()


def renderizar_botao_oauth(provider: str, config: Dict[str, object]) -> None:
    oauth = obter_config_oauth(provider, config)
    if oauth is None:
        st.warning(f"Login com {provider.title()} ainda nao esta configurado no secrets.toml.")
        return

    url = montar_url_autorizacao(provider, config)
    if not url:
        st.warning(f"Nao foi possivel montar a URL de login do {provider.title()}.")
        return

    label = f"Entrar com {oauth['label']}"
    if hasattr(st, "link_button"):
        st.link_button(label, url, use_container_width=True)
    else:
        st.markdown(f"[{label}]({url})")

    st.caption(f"Redirect URI configurada: {oauth['redirect_uri']}")


def renderizar_tela_login(config: Dict[str, object]) -> None:
    st.title("Bem-vindo ao Projeto TCC")
    st.caption("Escolha como deseja entrar antes de acessar o sistema de dimensionamento.")

    metodo = st.radio(
        "Como deseja entrar?",
        options=["Login local", "SSO Google"],
        horizontal=True,
    )

    if metodo == "Login local":
        usuarios = carregar_usuarios_locais(config)
        with st.form("login_local"):
            email = st.text_input("E-mail")
            senha = st.text_input("Senha", type="password")
            enviar = st.form_submit_button("Entrar", use_container_width=True)

        if enviar:
            usuario = autenticar_login_local(email, senha, usuarios)
            if usuario is None:
                st.error("E-mail ou senha invalidos.")
            else:
                st.session_state["auth_user"] = usuario
                st.rerun()

        if not usuarios:
            st.info("Configure LOCAL_LOGIN_EMAIL e LOCAL_LOGIN_PASSWORD no secrets.toml para habilitar o login local.")

    else:
        st.write("Use sua conta Google para entrar no sistema.")
        renderizar_botao_oauth("google", config)

    st.stop()


def renderizar_aplicacao_principal() -> None:
    usuario = st.session_state.get("auth_user", {})

    with st.sidebar:
        st.success(f"Conectado como: {usuario.get('name', 'Usuario')}")
        st.caption(f"E-mail: {usuario.get('email', '-')}")
        st.caption(f"Acesso: {str(usuario.get('provider', 'local')).title()}")
        if st.button("Sair", use_container_width=True):
            sair()

    st.title("Sistema de Dimensionamento Eletrico Residencial")
    st.caption(
        "Versao corrigida para previsao minima de cargas pela NBR 5410, resumo simplificado de demanda CPFL/GED-13 "
        "e caracteristicas do padrao de entrada conforme a tabela 1A usada nas aulas."
    )

    with st.sidebar:
        st.header("Projeto")
        nome_projeto = st.text_input("Nome do projeto", value="Meu Projeto")
        responsavel = st.text_input("Responsavel", value="")
        numero_comodos = st.number_input("Quantidade de comodos", min_value=1, max_value=30, value=5, step=1)

        st.header("Padrao de Entrada")
        fase_padrao = st.selectbox(
            "Fase do padrao de entrada",
            options=["Automatico", "Monofasico", "Bifasico", "Trifasico"],
            index=0,
            help="Se ficar em Automatico, o app sugere a fase com base na categoria. Se voce quiser, pode fixar manualmente.",
        )

    st.subheader("1) Cadastro dos comodos")

    comodos = []
    equipamentos_gerais_demanda = []

    for i in range(int(numero_comodos)):
        with st.expander(f"Comodo {i + 1}", expanded=(i == 0)):
            col1, col2, col3 = st.columns(3)
            with col1:
                nome_comodo = st.text_input(f"Nome do comodo {i + 1}", key=f"nome_{i}", value=f"Comodo {i + 1}")
            with col2:
                tipo_comodo = st.selectbox(
                    f"Tipo do comodo {i + 1}",
                    options=[
                        "quarto",
                        "suite",
                        "sala",
                        "cozinha",
                        "banheiro",
                        "area_servico",
                        "lavanderia",
                        "copa",
                        "copa_cozinha",
                        "escritorio",
                        "circulacao",
                        "corredor",
                        "hall",
                        "lavabo",
                        "closet",
                        "varanda",
                        "garagem",
                        "sotao",
                        "subsolo",
                        "casa_maquinas",
                        "sala_bombas",
                        "barrilete",
                        "outro",
                    ],
                    key=f"tipo_{i}",
                )
            with col3:
                area = st.number_input(f"Area (m2) - {i + 1}", min_value=0.01, value=10.0, step=0.1, key=f"area_{i}")

            col4, col5 = st.columns(2)
            with col4:
                perimetro = st.number_input(f"Perimetro (m) - {i + 1}", min_value=0.01, value=12.0, step=0.1, key=f"per_{i}")
            with col5:
                bancadas_validas = 0
                if normalizar_ambiente(tipo_comodo) in {"cozinha", "copa", "copa_cozinha", "area_servico", "lavanderia"}:
                    bancadas_validas = st.number_input(
                        f"Qtde. de bancadas >= 0,30 m - {i + 1}",
                        min_value=0,
                        value=0,
                        step=1,
                        key=f"bancadas_{i}",
                    )
                else:
                    st.markdown("**Bancadas:** nao se aplica")

            st.markdown("**TUEs do comodo**")
            qtd_tues = st.number_input(
                f"Quantidade de TUEs - {i + 1}",
                min_value=0,
                max_value=10,
                value=0,
                step=1,
                key=f"qtd_tue_{i}",
            )
            tues = []

            for j in range(int(qtd_tues)):
                c1, c2, c3 = st.columns([2, 1, 2])
                with c1:
                    nome_eq = st.text_input(f"Equipamento {j + 1}", key=f"eq_nome_{i}_{j}", value=f"Equipamento {j + 1}")
                with c2:
                    potencia_w = st.number_input(
                        f"Potencia (W) {j + 1}",
                        min_value=0.0,
                        value=1000.0,
                        step=100.0,
                        key=f"eq_pot_{i}_{j}",
                    )
                with c3:
                    categoria_demanda = st.selectbox(
                        f"Categoria demanda {j + 1}",
                        options=list(CATEGORIAS_DEMANDA.keys()),
                        format_func=lambda x: f"{x}) {CATEGORIAS_DEMANDA[x]}",
                        key=f"eq_cat_{i}_{j}",
                    )

                registro_eq = {
                    "nome": nome_eq,
                    "potencia_w": potencia_w,
                    "categoria_demanda": categoria_demanda,
                    "comodo": nome_comodo,
                }
                tues.append(registro_eq)
                equipamentos_gerais_demanda.append(registro_eq)

            comodos.append(
                {
                    "nome": nome_comodo,
                    "tipo": tipo_comodo,
                    "area": area,
                    "perimetro": perimetro,
                    "bancadas_validas": bancadas_validas,
                    "tues": tues,
                }
            )

    if st.button("Calcular projeto", type="primary"):
        resultados = []
        total_ilum_va = 0.0
        total_tug_va = 0.0
        total_tue_w = 0.0

        for comodo in comodos:
            ilum = calcular_iluminacao(comodo["area"])
            tug = calcular_tug(
                area=comodo["area"],
                perimetro=comodo["perimetro"],
                ambiente=comodo["tipo"],
                bancadas_validas=int(comodo["bancadas_validas"]),
            )
            tue = calcular_tue(comodo["tues"])

            total_ilum_va += ilum["potencia_va"]
            total_tug_va += tug["potencia_total_va"]
            total_tue_w += tue["potencia_total_w"]

            resultados.append(
                {
                    "Comodo": comodo["nome"],
                    "Tipo": comodo["tipo"],
                    "Area (m2)": round(comodo["area"], 2),
                    "Perimetro (m)": round(comodo["perimetro"], 2),
                    "Iluminacao - Pontos min.": ilum["pontos_minimos"],
                    "Iluminacao - Stotal (VA)": ilum["potencia_va"],
                    "TUG - Pontos": tug["pontos"],
                    "TUG - Sponto (VA)": formatar_sponto_tug(tug["potencias_va"]),
                    "TUG - Stotal (VA)": tug["potencia_total_va"],
                    "TUE - Equipamentos": tue["descricao"],
                    "TUE - Ptotal (W)": round(tue["potencia_total_w"], 1),
                }
            )

        df_resultados = pd.DataFrame(resultados)

        st.subheader("2) Tabela de cargas por comodo")
        st.dataframe(df_resultados, use_container_width=True)

        potencia_instalada_total_w = total_ilum_va + total_tug_va + total_tue_w

        st.subheader("3) Totais instalados")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Iluminacao total", f"{formatar_numero_br(total_ilum_va, 1)} VA")
        c2.metric("TUG total", f"{formatar_numero_br(total_tug_va, 1)} VA")
        c3.metric("TUE total", f"{formatar_numero_br(total_tue_w, 1)} W")
        c4.metric("Carga instalada total", f"{formatar_numero_br(potencia_instalada_total_w, 1)} W")

        st.info(
            "Observacao: a NBR 5410 define a carga minima de iluminacao em VA para dimensionamento; "
            "isso nao determina automaticamente a quantidade real de lampadas/luminarias do projeto."
        )

        st.subheader("4) Resumo simplificado de demanda CPFL / GED-13")
        df_demanda, total_demanda_w = calcular_demanda_cpfl_simplificada(
            carga_iluminacao_va=total_ilum_va,
            carga_tug_va=total_tug_va,
            equipamentos_tue=equipamentos_gerais_demanda,
        )
        st.dataframe(df_demanda, use_container_width=True)
        st.metric("Demanda total simplificada", f"{formatar_numero_br(total_demanda_w, 1)} W")

        st.subheader("5) Caracteristicas do Padrao de Entrada")
        padrao_entrada = calcular_padrao_entrada(
            carga_instalada_w=potencia_instalada_total_w,
            demanda_total_w=total_demanda_w,
            fase_escolhida=fase_padrao,
        )

        df_padrao = pd.DataFrame(
            {
                "Caracteristica": [
                    "Fase",
                    "Categoria",
                    "Carga Instalada",
                    "Tipo de Caixa",
                    "Disjuntor",
                    "Medida do Eletroduto",
                ],
                "Valor": [
                    padrao_entrada["Fase"],
                    padrao_entrada["Categoria"],
                    padrao_entrada["Carga Instalada"],
                    padrao_entrada["Tipo de Caixa"],
                    padrao_entrada["Disjuntor"],
                    padrao_entrada["Medida do Eletroduto"],
                ],
            }
        )
        st.table(df_padrao)

        st.caption(
            "Categoria sugerida conforme a tabela 1A da GED-13 usada nas aulas. "
            "Se necessario, a fase pode ser ajustada manualmente no menu lateral."
        )

        st.subheader("6) Exportacao")
        df_resumo = pd.DataFrame(
            {
                "Indicador": [
                    "Projeto",
                    "Responsavel",
                    "Quantidade de comodos",
                    "Iluminacao total (VA)",
                    "TUG total (VA)",
                    "TUE total (W)",
                    "Carga instalada total (W)",
                    "Demanda total simplificada (W)",
                    "Fase padrao",
                    "Categoria padrao",
                    "Tipo de caixa",
                    "Disjuntor (A)",
                    "Medida do eletroduto",
                ],
                "Valor": [
                    nome_projeto,
                    responsavel or "-",
                    len(comodos),
                    total_ilum_va,
                    total_tug_va,
                    total_tue_w,
                    potencia_instalada_total_w,
                    total_demanda_w,
                    padrao_entrada["Fase"],
                    padrao_entrada["Categoria"],
                    padrao_entrada["Tipo de Caixa"],
                    padrao_entrada["Disjuntor"],
                    padrao_entrada["Medida do Eletroduto"],
                ],
            }
        )

        excel_bytes = gerar_excel_bytes(
            df_resultados=df_resultados,
            df_demanda=df_demanda,
            df_padrao=df_padrao,
            df_resumo=df_resumo,
        )

        st.download_button(
            label="Baixar relatorio em XLSX",
            data=excel_bytes,
            file_name=f"relatorio_eletrico_{nome_projeto.replace(' ', '_').lower()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("7) Conferencia rapida")
        st.markdown(
            f"""
            **Projeto:** {nome_projeto}  
            **Responsavel:** {responsavel or '-'}  
            **Quantidade de comodos:** {len(comodos)}  
            **Soma iluminacao:** {formatar_numero_br(total_ilum_va, 1)} VA  
            **Soma TUG:** {formatar_numero_br(total_tug_va, 1)} VA  
            **Soma TUE:** {formatar_numero_br(total_tue_w, 1)} W  
            **Carga instalada total:** {formatar_numero_br(potencia_instalada_total_w, 1)} W  
            **Demanda simplificada:** {formatar_numero_br(total_demanda_w, 1)} W  
            **Padrao sugerido:** {padrao_entrada['Categoria']} / {padrao_entrada['Fase']} / Caixa {padrao_entrada['Tipo de Caixa']} / Disj. {padrao_entrada['Disjuntor']} A / Eletroduto {padrao_entrada['Medida do Eletroduto']}
            """
        )


def main() -> None:
    config = carregar_config_autenticacao()
    processar_callback_oauth(config)

    if not usuario_autenticado():
        renderizar_tela_login(config)

    renderizar_aplicacao_principal()


main()
