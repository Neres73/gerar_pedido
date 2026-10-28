"""
Macro: geração de requisições no Coupa a partir da nota do fornecedor ou de uma planilha.

Entrada (perguntada ao rodar):
  [1] Planilha "geração pedidos coupa.xlsx" (aba Linhas: PLU, qtd, valor, CNPJ, CC, G = conta, H = TR)
  [2] Nota: caminho/conteúdo do XML da NF-e (leitura exata) ou texto copiado do PDF da DANFE.
      Para cada item, a macro escolhe o PLU (De-Para, confiáveis, parecidos) e confere no Coupa se ele
      está livre (sem fornecedor, sem valor, com o CC exigido).

Fluxo no Coupa (em loop, uma requisição após a outra):
  1. Abre o carrinho -> "Revisar carrinho"; apaga linhas que tenham sobrado.
  2. Cabeçalho: perfil pré-definido (perfis.json) ou manual.
  3. Cada linha: PLU -> fornecedor (Outro + CNPJ) -> qtd/preço -> data -> TR -> empresa + CC -> Salvar
     -> lupa "Escolher uma conta" (CC da loja + K, ou nº da coluna G + A).
  4. Confere todas as linhas com a nota (valor e CC) e grava os relatórios em Geração_pedidos_coupa/.

Teclas: F8 pausa/continua · Ctrl+Alt+F8 para com segurança (o progresso fica salvo para continuar).
"""

import builtins
import json
import math
import random
import re
import sys
import threading
import time
import traceback
import unicodedata
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import urljoin

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

# ================== CONFIGURAÇÃO ==================
# Os dados do ambiente (caminhos, URL, códigos internos) ficam no arquivo .env - veja o .env.example.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass
import os


def _env(nome, padrao=""):
    return os.getenv(nome, padrao).strip().strip('"')


CAMINHO_PLANILHA = _env("CAMINHO_PLANILHA", "geração pedidos coupa.xlsx")
CAMINHO_REGIOES = _env("CAMINHO_REGIOES", "regioes.xlsb")   # CNPJ da loja -> centro de custo (arquivo ou pasta)
URL_COUPA = _env("COUPA_URL")
PERFIL_EDGE = _env("PERFIL_EDGE", "./perfil_edge")  # pasta do perfil do navegador com a sessão salva
COD_EMPRESA = _env("COD_EMPRESA")                  # o que é pesquisado na caixa antes do centro de custo
CNPJ_RAIZ_EMPRESA = _env("CNPJ_RAIZ_EMPRESA")      # 8 primeiros dígitos do CNPJ da própria empresa (nunca é o fornecedor)
PERFIS_JSON = _env("PERFIS_JSON", "perfis.json")  # perfis de cabeçalho (veja perfis.example.json)
TIMEOUT = 20000

# Modo nota: a macro sugere PLUs da aba PLUs para cada item e pergunta se pode usar.
MAX_SUGESTOES = 5             # quantas sugestões por item antes da genérica
COMPLETAR_COM_BARATOS = True  # se houver menos sugestões parecidas, completa com PLUs de valor baixo
POOL_BARATOS = 40             # sorteia entre os N PLUs mais baratos (usados 3+ vezes no último ano)
PLU_GENERICO = _env("PLU_GENERICO") or None  # PLU "coringa" oferecido por último (vazio = não oferecer)
QTD_1_NO_GENERICO = True      # com o PLU genérico, lança quantidade 1 e o valor total da linha
AUTO_CONFIAVEL = True         # PLU marcado como confiável é usado sem perguntar
CONFIAVEL_SO_PARECIDOS = False  # True = automático só se o confiável for parecido com o item (ou do De-Para)

# XML da NF-e (leitura exata, sem depender do layout do PDF)
PASTA_XML = None              # pasta onde ficam os XMLs (ex.: r"\\servidor\xml_nfe"); None = não procura

VALOR_DAS_LINHAS = "nota"     # "nota" = as linhas somam o VALOR TOTAL DA NOTA (produtos + ST + frete + IPI
                              #          + seguro + outras despesas - desconto) - é o que bate no agendamento;
                              # "produto" = só o valor dos produtos
QTD_1_SE_NAO_DIVIDIR = True   # se valor ÷ quantidade não der um preço exato em centavos, lança 1 un com o
                              # valor total da linha (senão o total da linha no Coupa sai com centavos de diferença)

# Verificação automática do PLU no Coupa (fornecedor, valor e centro de custo)
VERIFICAR_NO_COUPA = True     # False = volta a perguntar "pode ser utilizado?"
CC_EXIGIDO = _env("CC_EXIGIDO")  # centro de custo que precisa estar na lista do PLU (vazio = não exige)
DIAS_BLOQUEIO = None          # None = PLU reprovado fica bloqueado para sempre (ou um nº de dias)
ESPERA_RESULTADO = 8000       # ms esperando o PLU aparecer na busca antes de considerar "não encontrado"
MAX_VERIFICACOES_ITEM = 15    # se nenhuma sugestão passar, continua sorteando PLUs baratos até esse limite    # ignorado ao procurar o CNPJ do emitente no texto

# Onde a macro guarda os relatórios (None = pasta "Geração_pedidos_coupa" ao lado da planilha)
PASTA_GERACAO = None
#   Geração_pedidos_coupa/
#     2026-09-30_1415 REQ 300001 - NF 12345/   <- uma pasta por requisição: resumo, conferência, log, nota, prints
#     _logs/                                   <- log completo de cada execução
#     _controle/                               <- arquivos internos (progresso, cache das regiões, pendências)

# Controle da execução
TECLA_PAUSA = "<f8>"                 # pausa / continua (funciona com qualquer janela em foco)
TECLA_PARAR = "<ctrl>+<alt>+<f8>"    # para com segurança no próximo ponto seguro (salva o progresso)
TENTATIVAS_AUTOMATICAS = 1           # quantas vezes a macro tenta de novo sozinha antes de perguntar
APAGAR_LINHAS_EXISTENTES = True      # requisição NOVA com linhas no carrinho: apaga antes de lançar
CONFIRMAR_ANTES_DE_APAGAR = False    # True = pergunta antes de apagar

# Seletores confirmados (HTML que você mandou)
SEL_CARRINHO = 'path[d^="M16.9508 5.07516"]'
SEL_LUPA_CONTA = 'img.searchImg[title="Escolher uma conta"], img[alt="Escolher uma conta"]'
TIPO_CONTA = "K"                                   # "K - Centro de custo" na janela da conta (padrão)

# Perfis de cabeçalho (o 1º é o padrão da pergunta), lidos de perfis.json - veja perfis.example.json.
# "conta_*" = o que vai na lupa "Escolher uma conta" de cada linha: com K vai o centro de custo da loja;
# com A vai o nº da COLUNA G da aba Linhas. Sem o arquivo, o cabeçalho é preenchido à mão.
def _carregar_perfis():
    import json
    try:
        return json.loads(Path(PERFIS_JSON).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except Exception as e:
        print(f"ℹ Não consegui ler {PERFIS_JSON} ({e}) - o cabeçalho será preenchido à mão.")
        return {}


PERFIS = _carregar_perfis()
SEL_SELECIONAR_TODOS = "#select_all_checkbox"
SEL_TR = "#requisition_line_custom_field_10"       # campo "TR" da linha (perfis com campo_tr = coluna H)
SEL_REVISAR_CARRINHO = 'a.coupaCartPopover__reviewCartBtn, a[title="Revisar carrinho"]'
SEL_ADICIONAR_LINHA = 'span:text-is("Adicionar linha")'
SEL_PLU = "#requisition_line_description"
SEL_SELECT_FORNECEDOR = "#requisition_line_supplier_picker"   # Nenhum / Indefinido / Outro
SEL_FORNECEDOR = "input.supplierSearch"                        # campo que aparece ao escolher "Outro"
SEL_PRECO = "#requisition_line_unit_price_amount"
SEL_CC_LINHA = "#requisition_line_custom_field_3_id_chosen"   # caixa "Selecionar" da empresa/CC da LINHA
SEL_SALVAR = 'span:text-is("Salvar")'

# a barra "Do que você precisa?" muda de lugar conforme a página (topo ou meio da tela):
# a macro usa a que estiver VISÍVEL, em qualquer posição
SEL_BUSCA_BTN = ('button.s-searchBarBtn, button.searchBar__btn, '
                 'button[aria-label^="Pesquise itens"]')
SEL_RESULTADO_PLU = ('[data-testid="hosted-title-link"][title^="{plu} |"], '
                     '.item__title__link[title^="{plu} |"], '
                     'a.s-requestHostedTitle[title^="{plu} |"]')
SEL_FORNECEDOR_ITEM = "#supplier_item_picker"

SEL_QUANTIDADE = "#requisition_line_quantity"
SEL_DATA = "#requisition_line_need_by_date"
# ==================================================


# ================================================================ ROBUSTEZ
# ---------------------------------------------------------------- pausa / parada por tecla
PAUSA = threading.Event()
PARAR = threading.Event()


class ParadaSolicitada(Exception):
    """Levantada nos pontos seguros quando você aperta a tecla de parar."""


def iniciar_teclas():
    """Liga as teclas de pausa/parada (biblioteca pynput). Sem ela, a macro roda sem esse recurso."""
    import importlib.util
    if importlib.util.find_spec("pynput") is None:
        print("ℹ Pausa por tecla desligada: rode  pip install pynput  para ativar "
              f"({TECLA_PAUSA.strip('<>').upper()} pausa/continua).")
        return None
    try:
        from pynput import keyboard
    except Exception as e:
        print(f"ℹ Não consegui ligar as teclas de pausa ({e}). A macro roda normalmente sem elas.")
        return None

    def alternar():
        if PAUSA.is_set():
            PAUSA.clear()
            print("\n▶ Continuando...", flush=True)
        else:
            PAUSA.set()
            print(f"\n⏸ Pausa pedida — a macro para no próximo ponto seguro. "
                  f"{TECLA_PAUSA.strip('<>').upper()} de novo para continuar.", flush=True)

    def parar():
        PARAR.set()
        PAUSA.clear()
        print("\n⏹ Parada pedida — a macro termina o passo atual, salva o progresso e sai.", flush=True)

    try:
        ouvinte = keyboard.GlobalHotKeys({TECLA_PAUSA: alternar, TECLA_PARAR: parar})
        ouvinte.daemon = True
        ouvinte.start()
        print(f"⌨ {TECLA_PAUSA.strip('<>').upper()} = pausar/continuar · "
              f"{TECLA_PARAR.replace('<', '').replace('>', '').upper()} = parar com segurança")
        return ouvinte
    except Exception as e:
        print(f"ℹ Não consegui ligar as teclas de pausa ({e}). A macro roda normalmente sem elas.")
        return None


def ponto_de_pausa(page=None, onde=""):
    """Ponto seguro: aqui a macro obedece a pausa e a parada. Enquanto pausada, continua atendendo o
    navegador (avisos do site etc.) - dá para mexer no Coupa à mão e depois continuar."""
    if PARAR.is_set():
        raise ParadaSolicitada(onde)
    if not PAUSA.is_set():
        return
    print(f"  ⏸ Pausado{(' (' + onde + ')') if onde else ''}. Pode mexer no Coupa; "
          f"{TECLA_PAUSA.strip('<>').upper()} para continuar.", flush=True)
    while PAUSA.is_set():
        if PARAR.is_set():
            raise ParadaSolicitada(onde)
        try:
            page.wait_for_timeout(300) if page else time.sleep(0.3)
        except Exception:
            time.sleep(0.3)
    if PARAR.is_set():
        raise ParadaSolicitada(onde)


# ---------------------------------------------------------------- pastas de relatórios
def pasta_base():
    base = Path(PASTA_GERACAO) if PASTA_GERACAO else Path(CAMINHO_PLANILHA).parent / "Geração_pedidos_coupa"
    base.mkdir(parents=True, exist_ok=True)
    return base


def pasta_controle():
    p = pasta_base() / "_controle"
    p.mkdir(exist_ok=True)
    return p


PASTA_REQ = {"dir": None}
LOG_REQ = {"arq": None}


def _nome_seguro(texto):
    return re.sub(r'[<>:"/\\|?*]+', "-", texto).strip(" .")[:120]


def iniciar_pasta_requisicao(descricao, existente=None):
    """Cria (ou reabre, na retomada) a pasta desta requisição e passa a gravar o log dela lá também."""
    fechar_pasta_requisicao()
    if existente and Path(existente).is_dir():
        pasta = Path(existente)
    else:
        pasta = pasta_base() / _nome_seguro(f"{datetime.now():%Y-%m-%d_%H%M} {descricao}")
        k = 2
        while pasta.exists():
            pasta = pasta.with_name(f"{pasta.name} ({k})")
            k += 1
        pasta.mkdir(parents=True)
    PASTA_REQ["dir"] = pasta
    try:
        LOG_REQ["arq"] = open(pasta / "log.txt", "a", encoding="utf-8")
    except Exception:
        LOG_REQ["arq"] = None
    return pasta


def nomear_pasta_requisicao(num_req):
    """Quando o nº da requisição aparece, põe ele no nome da pasta ('... REQ 300001 - NF 12345')."""
    pasta = PASTA_REQ["dir"]
    if not pasta or not num_req or f"REQ {num_req}" in pasta.name:
        return pasta
    partes = pasta.name.split(" ", 1)
    novo = pasta.with_name(_nome_seguro(f"{partes[0]} REQ {num_req}" + (f" - {partes[1]}" if len(partes) > 1 else "")))
    arq_log = LOG_REQ["arq"]
    try:
        if arq_log:
            arq_log.close()
        pasta.rename(novo)
        pasta = novo
    except Exception:
        pass                                        # se o Windows não deixar renomear, fica o nome antigo
    PASTA_REQ["dir"] = pasta
    try:
        LOG_REQ["arq"] = open(pasta / "log.txt", "a", encoding="utf-8")
    except Exception:
        LOG_REQ["arq"] = None
    return pasta


def fechar_pasta_requisicao():
    if LOG_REQ["arq"]:
        try:
            LOG_REQ["arq"].close()
        except Exception:
            pass
    LOG_REQ["arq"] = None
    PASTA_REQ["dir"] = None


def pasta_atual():
    """Pasta da requisição em andamento (ou a base, se ainda não houver requisição)."""
    return PASTA_REQ["dir"] or pasta_base()


def gravar_na_pasta(nome, conteudo):
    try:
        arq = pasta_atual() / nome
        if isinstance(conteudo, bytes):
            arq.write_bytes(conteudo)
        else:
            arq.write_text(conteudo, encoding="utf-8")
        return arq
    except Exception:
        return None


NOTA_ATUAL = {}


def descricao_requisicao(modo):
    """Texto do nome da pasta: 'NF 12345 - fornecedor 01.234.567-0001-89' ou 'planilha'."""
    if modo == "1":
        return "planilha"
    chave = NOTA_ATUAL.get("chave") or ""
    partes = []
    if len(chave) == 44:
        partes.append(f"NF {int(chave[25:34])}")
    elif NOTA_ATUAL.get("numero_nf"):
        partes.append(f"NFS-e {NOTA_ATUAL['numero_nf']}")
    if NOTA_ATUAL.get("cnpj_forn"):
        partes.append(f"fornecedor {fmt_cnpj(NOTA_ATUAL['cnpj_forn']).replace('/', '-')}")
    return " - ".join(partes) or "nota"


def guardar_nota_na_pasta():
    tipo, conteudo = NOTA_ATUAL.get("tipo"), NOTA_ATUAL.get("conteudo")
    if not conteudo:
        return
    if tipo == "xml":
        gravar_na_pasta("nota.xml", conteudo if isinstance(conteudo, bytes) else conteudo.encode("utf-8"))
    else:
        gravar_na_pasta("nota_ocr.txt" if NOTA_ATUAL.get("ocr") else "nota_texto.txt", conteudo)
    if NOTA_ATUAL.get("pdf"):
        gravar_na_pasta("nota.pdf", NOTA_ATUAL["pdf"])
    if NOTA_ATUAL.get("imagem"):
        ext, dados = NOTA_ATUAL["imagem"]
        gravar_na_pasta(f"nota{ext}", dados)


def gravar_resumo(estado, divergencias=None, status=""):
    linhas = estado.get("linhas", [])
    txt = [f"Requisição nº {_num_requisicao(estado.get('url')) or '?'}",
           f"Endereço: {estado.get('url') or '-'}",
           f"Iniciada em: {estado.get('inicio', '-')}  |  gravado em: {datetime.now():%d/%m/%Y %H:%M}",
           f"Perfil: {estado.get('perfil_nome') or 'manual'}",
           f"Status: {status}",
           f"Linhas: {len(estado.get('salvas', []))} de {len(linhas)} salvas · "
           f"{len(estado.get('contas', []))} contas ajustadas",
           f"Total das linhas: R$ {brl(round(sum(float(l['total']) for l in linhas), 2))}"
           + (f"  |  total da nota: R$ {brl(estado['total_nota'])}" if estado.get("total_nota") else ""),
           "", "Linhas:"]
    for n, l in enumerate(linhas, start=1):
        txt.append(f"  {n:>3}. PLU {l['plu']:<9} qtd {l.get('qtd', 1):g}  R$ {brl(float(l['total'])):>12}  "
                   f"CC {l.get('cc', '')}  {l.get('origem', '')}")
    if divergencias is not None:
        txt += ["", "Conferência: " + ("tudo confere" if not divergencias else f"{len(divergencias)} divergência(s)")]
        txt += [f"  - {d}" for d in divergencias]
    gravar_na_pasta("resumo.txt", "\n".join(txt) + "\n")


def migrar_arquivos_antigos():
    """Leva para a pasta nova os arquivos internos que ficavam soltos ao lado da planilha (versões antigas)."""
    antiga = Path(CAMINHO_PLANILHA).parent
    for nome in ("progresso_macro.json", "aprendizado_pendente.json", "regioes_cache.json"):
        velho, novo = antiga / nome, pasta_controle() / nome
        if velho.exists() and not novo.exists():
            try:
                velho.replace(novo)
            except Exception:
                pass


# ---------------------------------------------------------------- log da execução
class _Duplicador:
    """Escreve no terminal E no arquivo de log."""
    def __init__(self, original, arquivo):
        self.original, self.arquivo = original, arquivo

    def write(self, texto):
        self.original.write(texto)
        for arq in (self.arquivo, LOG_REQ["arq"]):
            if arq is None:
                continue
            try:
                arq.write(texto)
                arq.flush()
            except Exception:
                pass

    def flush(self):
        self.original.flush()


def iniciar_log():
    """Tudo que aparece no terminal (e as suas respostas) vai também para logs_macro/AAAA-MM-DD_HHMM.log."""
    try:
        pasta = pasta_base() / "_logs"
        pasta.mkdir(exist_ok=True)
        arq = open(pasta / f"{datetime.now():%Y-%m-%d_%H%M%S}.log", "a", encoding="utf-8")
        sys.stdout = _Duplicador(sys.stdout, arq)
        sys.stderr = _Duplicador(sys.stderr, arq)
        input_original = builtins.input

        def input_com_log(msg=""):
            r = input_original(msg)
            for a in (arq, LOG_REQ["arq"]):
                if a is None:
                    continue
                try:
                    a.write(f"{r}\n")
                    a.flush()
                except Exception:
                    pass
            return r
        builtins.input = input_com_log
        return arq.name
    except Exception:
        return None


# ---------------------------------------------------------------- progresso (retomar de onde parou)
def _arq_progresso():
    return pasta_controle() / "progresso_macro.json"


def salvar_progresso(estado):
    """Grava o progresso a cada linha (arquivo temporário + troca, para nunca ficar corrompido)."""
    try:
        arq = _arq_progresso()
        tmp = arq.with_suffix(".tmp")
        tmp.write_text(json.dumps(estado, ensure_ascii=False, indent=1), encoding="utf-8")
        tmp.replace(arq)
    except Exception as e:
        print(f"  ⚠ Não consegui salvar o progresso ({e})")


def carregar_progresso():
    arq = _arq_progresso()
    if not arq.exists():
        return None
    try:
        return json.loads(arq.read_text(encoding="utf-8"))
    except Exception:
        return None


def encerrar_progresso():
    """Execução concluída: guarda o arquivo com a data (histórico) e tira do caminho."""
    arq = _arq_progresso()
    if arq.exists():
        try:
            destino = PASTA_REQ["dir"] or pasta_controle()
            arq.replace(destino / ("progresso.json" if PASTA_REQ["dir"]
                                   else f"progresso_concluido_{datetime.now():%Y%m%d_%H%M%S}.json"))
        except Exception:
            pass


def resumo_progresso(estado):
    total = len(estado.get("linhas", []))
    salvas = len(estado.get("salvas", []))
    contas = len(estado.get("contas", []))
    return (f"iniciada em {estado.get('inicio', '?')} · perfil {estado.get('perfil_nome') or 'manual'} · "
            f"{salvas} de {total} linhas salvas · {contas} contas ajustadas"
            + (f"\n    requisição: {estado['url']}" if estado.get("url") else " · requisição ainda não aberta"))


# ---------------------------------------------------------------- validação antes de lançar
def ajustar_quantidades(linhas):
    """Com o total da nota rateado (ST, frete...), valor ÷ quantidade quase nunca dá preço exato em
    centavos - e o total da linha no Coupa sairia com diferença. Nesses casos lança 1 un com o valor
    total da linha (a quantidade da nota fica registrada na origem da linha)."""
    if not QTD_1_SE_NAO_DIVIDIR:
        return
    ajustadas = []
    for n, l in enumerate(linhas, start=1):
        q, t = float(l.get("qtd") or 1), round(float(l["total"]), 2)
        if q != 1 and round(round(t / q, 2) * q, 2) != t:
            l["origem"] = f"{l.get('origem', '')} (nota: {q:g} un)".strip()
            l["qtd"] = 1.0
            ajustadas.append(n)
    if ajustadas:
        print(f"  ℹ {len(ajustadas)} linha(s) vão como 1 un com o valor total, porque valor ÷ quantidade não "
              f"dá preço exato em centavos: {', '.join(map(str, ajustadas))}")


def validar_linhas(linhas):
    """Confere as linhas antes de abrir o Coupa. Devolve a lista de problemas (vazia = ok)."""
    problemas = []
    for n, l in enumerate(linhas, start=1):
        onde = f"linha {n} ({l.get('origem', '')})"
        if not str(l.get("plu", "")).strip():
            problemas.append(f"{onde}: sem PLU")
        if not l.get("qtd") or l["qtd"] <= 0:
            problemas.append(f"{onde}: quantidade inválida ({l.get('qtd')})")
        if not l.get("total") or l["total"] <= 0:
            problemas.append(f"{onde}: valor inválido ({l.get('total')})")
        if not cnpj_valido(l.get("cnpj", "")):
            problemas.append(f"{onde}: CNPJ do fornecedor inválido ({l.get('cnpj')})")
        if not str(l.get("cc", "")).strip():
            problemas.append(f"{onde}: sem centro de custo")
        if "tr" in l and not l["tr"]:
            problemas.append(f"{onde}: sem texto do TR (coluna H vazia)")
    return problemas


# ---------------------------------------------------------------- aprendizado que não coube na planilha
def _arq_pendente():
    return pasta_controle() / "aprendizado_pendente.json"


def gravar_pendente(novos_de_para, novos_confiaveis, novos_bloqueados):
    arq = _arq_pendente()
    try:
        atual = json.loads(arq.read_text(encoding="utf-8")) if arq.exists() else {"de_para": [], "conf": [], "bloq": {}}
    except Exception:
        atual = {"de_para": [], "conf": [], "bloq": {}}
    atual["de_para"] += list(novos_de_para)
    atual["conf"] += [p for p in novos_confiaveis if p not in atual["conf"]]
    atual["bloq"].update(novos_bloqueados or {})
    arq.write_text(json.dumps(atual, ensure_ascii=False, indent=1), encoding="utf-8")


def aplicar_pendente(plus=None):
    """Na execução seguinte, tenta gravar na planilha o que ficou pendente."""
    arq = _arq_pendente()
    if not arq.exists():
        return
    try:
        dados = json.loads(arq.read_text(encoding="utf-8"))
    except Exception:
        return
    print("Gravando na planilha o aprendizado pendente da execução anterior...")
    arq.unlink()
    salvar_aprendizado(dados.get("de_para", []), dados.get("conf", []), plus or {}, dados.get("bloq", {}))


# ---------------------------------------------------------------- utilidades
def br(valor, casas=2):
    """1234.5 -> '1234,50'"""
    return f"{valor:.{casas}f}".replace(".", ",")


def num_br(s):
    """'1.527,5100' -> 1527.51"""
    return float(s.replace(".", "").replace(",", "."))


def padrao_codigo(codigo):
    """Casa a opção cujo CÓDIGO é esse, não qualquer texto que contenha o número.
    '1220' casa '1220 - LJ CENTRO (1220)' e '1000 > 1220 - LJ CENTRO', mas NÃO casa
    '1047 - HUB NORTE EX LJ 1220 (1047)' nem '3019 - HUB SUL - 1220 (3019)'."""
    c = re.escape(str(codigo).strip())
    return re.compile(rf"(?:^|[>|]\s*){c}\s*-|\(\s*{c}\s*\)\s*$")


def normalizar_cnpj(valor):
    """CNPJ só com números e SEMPRE com 14 dígitos: o Excel transforma 01234567000189 em 1234567000189
    (número), e aí o CNPJ parece inválido. Completa os zeros à esquerda."""
    d = so_digitos(texto_celula(valor) if not isinstance(valor, str) else valor)
    return d.zfill(14) if 11 <= len(d) < 14 else d


def formas_de_busca_cnpj(cnpj):
    """Jeitos de digitar o CNPJ na busca de fornecedor do Coupa, do mais provável ao menos.
    O Coupa guarda alguns CNPJs SEM o zero da frente (01234567000189 só aparece como 1234567000189)."""
    d = normalizar_cnpj(cnpj)
    sem_zero = d.lstrip("0")
    formas = [sem_zero, d] if sem_zero != d else [d]
    formas += [fmt_cnpj(d)]
    if FORMA_CNPJ_QUE_FUNCIONOU.get(d) in formas:           # o que deu certo antes vai primeiro
        formas.remove(FORMA_CNPJ_QUE_FUNCIONOU[d])
        formas.insert(0, FORMA_CNPJ_QUE_FUNCIONOU[d])
    return list(dict.fromkeys(formas))


FORMA_CNPJ_QUE_FUNCIONOU = {}


def so_digitos(s):
    return re.sub(r"\D", "", str(s or ""))


# ---------------------------------------------------------------- planilha
def codigo_plu(valor):
    """Só o número do PLU. Aceita célula com '1234567 | PECA EXEMPLO 100', '1234567' ou 1234567.0."""
    t = texto_celula(valor)
    m = re.match(r"\s*(\d{4,})", t)
    return m.group(1) if m else t


def ler_planilha(wb, permitir_repetido=False):
    ws = wb["Linhas"]
    linhas = []
    for n, (plu, qtd, total, cnpj, cc, _consulta, col_g, col_h) in enumerate(
        ws.iter_rows(min_row=2, max_col=8, values_only=True), start=2
    ):
        if plu is None:
            continue
        linhas.append({
            "origem": f"linha {n} da planilha",
            "plu": codigo_plu(plu),                  # "1234567 | PECA EXEMPLO 100" -> "1234567"
            "qtd": float(qtd or 1),
            "total": float(total),
            "cnpj": normalizar_cnpj(cnpj),
            "cc": str(cc).strip(),
            "g": texto_celula(col_g),
            "h": texto_celula(col_h),
        })
    repetidos = {l["plu"] for l in linhas if sum(x["plu"] == l["plu"] for x in linhas) > 1}
    if repetidos and not permitir_repetido:
        print("✖ PLU repetido na aba Linhas (cada PLU só pode ir em uma linha):")
        for plu in sorted(repetidos):
            print(f"   PLU {plu}: " + ", ".join(l["origem"] for l in linhas if l["plu"] == plu))
        return []
    return linhas


def carregar_plus(wb):
    """{PLU: {desc, um, pedidos, preco}} da aba PLUs."""
    if "PLUs" not in wb.sheetnames:
        return {}
    plus = {}
    for plu, desc, um, _grp, pedidos, data, preco in wb["PLUs"].iter_rows(min_row=2, max_col=7, values_only=True):
        if plu and isinstance(pedidos, (int, float)):
            plus[str(plu).strip()] = {"desc": str(desc or ""), "um": um or "",
                                      "pedidos": int(pedidos), "data": data,
                                      "preco": preco if isinstance(preco, (int, float)) else None}
    return plus


def carregar_de_para(wb):
    """{(cnpj, código do fornecedor): PLU} da aba De-Para."""
    if "De-Para" not in wb.sheetnames:
        return {}
    mapa = {}
    for cnpj, cod, _desc, plu in wb["De-Para"].iter_rows(min_row=2, max_col=4, values_only=True):
        if cod and plu:
            mapa[(normalizar_cnpj(cnpj), str(cod).strip())] = str(plu).strip()
    return mapa


def carregar_confiaveis(wb):
    """{PLU: data} da aba Confiáveis - PLUs que você já aprovou e podem ser usados sem checar."""
    if "Confiáveis" not in wb.sheetnames:
        return {}
    return {str(r[0]).strip(): r[2] for r in wb["Confiáveis"].iter_rows(min_row=2, max_col=3, values_only=True) if r[0]}


def carregar_bloqueados(wb):
    """{PLU: motivo} da aba Bloqueados (todos, ou só os dos últimos DIAS_BLOQUEIO dias)."""
    if "Bloqueados" not in wb.sheetnames:
        return {}
    limite = date.today() - timedelta(days=DIAS_BLOQUEIO) if DIAS_BLOQUEIO else date.min
    bloq = {}
    for plu, motivo, quando in wb["Bloqueados"].iter_rows(min_row=2, max_col=3, values_only=True):
        if not plu:
            continue
        try:
            d = quando.date() if isinstance(quando, datetime) else datetime.strptime(str(quando), "%d/%m/%Y").date()
        except ValueError:
            d = date.today()
        if d >= limite:
            bloq[str(plu).strip()] = motivo
    return bloq


def salvar_aprendizado(novos_de_para, novos_confiaveis, plus, novos_bloqueados=None):
    """Grava De-Para, Confiáveis e Bloqueados na planilha (cria as abas se não existirem)."""
    novos_bloqueados = novos_bloqueados or {}
    if not novos_de_para and not novos_confiaveis and not novos_bloqueados:
        return
    for tentativa in range(3):
        try:
            wb = load_workbook(CAMINHO_PLANILHA)
            if "De-Para" not in wb.sheetnames:
                wb.create_sheet("De-Para").append(
                    ["CNPJ do fornecedor", "Código do fornecedor (nota)", "Descrição na nota", "PLU"])
            if "Confiáveis" not in wb.sheetnames:
                wb.create_sheet("Confiáveis").append(["PLU", "Descrição", "Aprovado em"])
            if "Bloqueados" not in wb.sheetnames:
                wb.create_sheet("Bloqueados").append(["PLU", "Motivo", "Verificado em"])
            for linha in novos_de_para:
                wb["De-Para"].append(linha)
            hoje = date.today().strftime("%d/%m/%Y")
            for plu in novos_confiaveis:
                wb["Confiáveis"].append([plu, plus.get(plu, {}).get("desc", ""), hoje])
            for plu, motivo in novos_bloqueados.items():
                wb["Bloqueados"].append([plu, motivo, hoje])
            wb.save(CAMINHO_PLANILHA)
            if novos_de_para:
                print(f"  {len(novos_de_para)} código(s) salvos na aba De-Para.")
            if novos_confiaveis:
                print(f"  {len(novos_confiaveis)} PLU(s) marcados como confiáveis: {', '.join(novos_confiaveis)}")
            if novos_bloqueados:
                print(f"  {len(novos_bloqueados)} PLU(s) reprovados gravados na aba Bloqueados.")
            return
        except PermissionError:
            if tentativa < 2:
                input("  ⚠ A planilha está aberta no Excel. Feche-a e aperte Enter para eu gravar o aprendizado...")
                continue
        except Exception as e:
            print(f"  ⚠ Erro ao gravar na planilha ({e}).")
            break
    # não deu: guarda num arquivo à parte e grava na próxima execução - nada se perde
    try:
        gravar_pendente(novos_de_para, novos_confiaveis, novos_bloqueados)
        print("  ℹ Aprendizado guardado em aprendizado_pendente.json; vai para a planilha na próxima execução.")
    except Exception as e:
        print(f"  ⚠ Não consegui guardar o aprendizado ({e}). O lançamento continua normalmente.")


# ---------------------------------------------------------------- nota fiscal
NUM4 = r"\d{1,3}(?:\.\d{3})*,\d{2,4}"      # quantidade / valor unitário: 2 a 4 casas (algumas notas cortam em 3)
NUM2 = r"\d{1,3}(?:\.\d{3})*,\d{2}"        # valor total: 2 casas
ITEM_NOTA = re.compile(
    # a descrição pode quebrar em até 3 linhas, mas NUNCA engole uma linha que começa outro item
    # o código pode vir na mesma linha da descrição (texto copiado) ou sozinho na linha de cima
    # (texto extraído do arquivo PDF, que põe cada coluna numa linha)
    rf"^(?P<cod>\d{{4,10}})[ \t]*\n?[ \t]*(?P<desc>[A-Za-zÀ-ú][^\n]*?(?:\n(?!\d{{4,10}}\s+[A-Za-zÀ-ú])[^\n]*?){{0,3}}?)\s+"
    rf"(?P<ncm>\d{{8}})\s+(?P<cst>\d{{3}})\s+(?P<cfop>\d{{4}})\s+(?P<un>[A-Z]{{1,4}})\s*"
    rf"(?P<qtd>{NUM4})\s*(?P<unit>{NUM4})\s*(?P<total>{NUM2})",
    re.MULTILINE | re.DOTALL,
)


def ratear_ate(itens, alvo):
    """Ajusta o 'total' dos itens para somarem exatamente 'alvo', distribuindo a diferença
    proporcionalmente ao valor de cada item (os centavos que sobram vão para o maior item).
    Devolve a diferença distribuída."""
    soma = round(sum(i["total"] for i in itens), 2)
    dif = round(alvo - soma, 2)
    if abs(dif) < 0.005 or soma <= 0:
        return 0.0
    base = [i["total"] for i in itens]
    for i, b in zip(itens, base):
        parte = round(dif * b / soma, 2)
        i["total"] = round(i["total"] + parte, 2)
        i["acrescimos"] = round(i.get("acrescimos", 0) + parte, 2)
    resto = round(alvo - sum(i["total"] for i in itens), 2)
    if resto:
        maior = max(itens, key=lambda i: i["total"])
        maior["total"] = round(maior["total"] + resto, 2)
        maior["acrescimos"] = round(maior.get("acrescimos", 0) + resto, 2)
    return dif


def total_geral_da_nota(texto):
    """'VALOR TOTAL DA NOTA' da DANFE (mesma regra do total dos produtos: rótulo e valor sozinhos)."""
    m = re.search(r"^[ \t]*VALOR TOTAL DA NOTA[ \t]*(?:\n[ \t]*)?([\d.]+,\d{2})[ \t]*$",
                  texto.replace("\r", ""), re.I | re.M)
    return num_br(m.group(1)) if m else None


def total_produtos_da_nota(texto):
    """'VALOR TOTAL DOS PRODUTOS' da DANFE (para conferir a soma dos itens lidos)."""
    # só quando o rótulo está sozinho na linha e o valor vem sozinho (na mesma linha ou na seguinte);
    # em layouts com vários rótulos numa linha e vários valores na outra, não dá para saber qual é - pula
    m = re.search(r"^[ \t]*VALOR TOTAL DOS PRODUTOS[ \t]*(?:\n[ \t]*)?([\d.]+,\d{2})[ \t]*$",
                  texto.replace("\r", ""), re.I | re.M)
    return num_br(m.group(1)) if m else None


def ler_area_transferencia():
    try:
        import tkinter as tk
        r = tk.Tk()
        r.withdraw()
        texto = r.clipboard_get()
        r.destroy()
        return texto
    except Exception:
        return ""


def obter_texto_nota():
    digitado = input("\nCopie a nota e aperte Enter aqui. Pode ser: o texto do PDF (Ctrl+A, Ctrl+C), o conteúdo do XML,\n"
                     "ou o caminho do arquivo .xml ou .pdf (Shift + clique direito no arquivo > 'Copiar como caminho').\n"
                     "Também pode COLAR aqui mesmo o caminho do .xml ou do .pdf da nota e dar Enter: ")
    if digitado.strip():                    # colou/digitou algo no terminal (ex.: o caminho do .xml)
        return digitado
    texto = ler_area_transferencia()
    if texto.strip():
        return texto
    print("Não consegui ler a área de transferência. Cole o texto abaixo e digite FIM numa linha sozinha:")
    linhas = []
    while (l := input()).strip().upper() != "FIM":
        linhas.append(l)
    return "\n".join(linhas)


# ---------------------------------------------------------------- XML da NF-e
def _xml_sem_ns(conteudo):
    """Lê o XML e tira os namespaces dos nomes DEPOIS de ler (assim 'det', 'prod'... funcionam direto).
    Aceita texto ou bytes; com bytes respeita a codificação declarada no arquivo (UTF-8, ISO-8859-1...)."""
    import xml.etree.ElementTree as ET
    if isinstance(conteudo, str):
        texto = conteudo.lstrip("\ufeff").strip()
        texto = re.sub(r"^<\?xml[^>]*\?>", "", texto)          # a declaração pode citar outra codificação
        dados = texto.encode("utf-8")
    else:
        dados = conteudo.lstrip(b"\xef\xbb\xbf").strip()
    # alguns sistemas geram o XML com "&" solto no texto (ex.: nas informações ao fisco "&|101940|0.00|&"),
    # o que é XML inválido - escapa esses "&" antes de ler (não mexe em &amp; &lt; &#233; etc.)
    dados = re.sub(rb"&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)", b"&amp;", dados)
    raiz = ET.fromstring(dados)
    for el in raiz.iter():
        if isinstance(el.tag, str) and "}" in el.tag:
            el.tag = el.tag.split("}", 1)[1]
        if ":" in str(el.tag):
            el.tag = el.tag.split(":", 1)[1]
    return raiz


def diagnosticar_xml(conteudo):
    """Explica por que um arquivo .xml não serve como nota (em vez de só dizer 'inválido')."""
    t = conteudo.decode("utf-8", "replace") if isinstance(conteudo, bytes) else conteudo
    inicio = t.lstrip("\ufeff \r\n\t")[:500].lower()
    if not inicio:
        return "o arquivo está vazio"
    if inicio.startswith("<!doctype html") or "<html" in inicio:
        return ("o arquivo é uma PÁGINA DA INTERNET salva com final .xml, não o XML da nota "
                "(o download não deu certo - baixe de novo ou peça o XML ao fornecedor)")
    if "<resnfe" in t.lower():
        return ("é só o RESUMO da nota (resNFe), sem os itens - acontece sem certificado. Peça o XML completo ao fornecedor ou à área fiscal")
    if "<proceventonfe" in t.lower() or "<evento" in t.lower() and "<infnfe" not in t.lower():
        return ("é um XML de EVENTO (cancelamento, carta de correção, manifestação), não a nota em si")
    if "<nfse" in t.lower() or "<compnfse" in t.lower():
        return "é uma NFS-e (nota de SERVIÇO), não uma NF-e de produtos - esse formato a macro ainda não lê"
    if "<infnfe" not in t.lower():
        return "o arquivo não tem a parte <infNFe> - não parece ser uma NF-e"
    return None


def _num(el, tag):
    v = el.findtext(tag) if el is not None else None
    try:
        return float(v) if v not in (None, "") else 0.0
    except ValueError:
        return 0.0


def eh_xml_nfe(texto):
    return bool(texto) and any(x in texto for x in ("<nfeProc", "<NFe", "<infNFe", "<resNFe", "<procEventoNFe",
                                                   ":nfeProc", ":NFe", ":infNFe"))


def ler_nota_xml(texto):
    """Itens, CNPJs e totais direto do XML (padrão nacional - igual para todo fornecedor).
    'texto' pode ser o conteúdo (str) ou os bytes do arquivo."""
    motivo = diagnosticar_xml(texto)
    if motivo:
        raise ValueError(motivo)
    try:
        raiz = _xml_sem_ns(texto)
    except Exception as e:
        raise ValueError(f"o XML está com a estrutura quebrada ({e}) - o arquivo pode ter sido cortado "
                         "ou alterado; baixe de novo")
    inf = raiz if raiz.tag == "infNFe" else raiz.find(".//infNFe")
    if inf is None:
        raise ValueError("o XML não tem a parte infNFe (não parece ser uma NF-e)")
    chave = so_digitos(inf.get("Id") or "") or so_digitos(raiz.findtext(".//chNFe") or "")
    emit = so_digitos(inf.findtext("emit/CNPJ") or "")
    dest = so_digitos(inf.findtext("dest/CNPJ") or "")
    itens = []
    for n, det in enumerate(inf.findall("det"), start=1):
        prod = det.find("prod")
        if prod is None:
            continue
        v_prod = _num(prod, "vProd")
        imp = det.find("imposto")
        acrescimos = (_num(prod, "vFrete") + _num(prod, "vSeg") + _num(prod, "vOutro") - _num(prod, "vDesc")
                      + (_num(imp.find(".//IPITrib"), "vIPI") if imp is not None else 0)
                      + (sum(_num(x, "vICMSST") + _num(x, "vFCPST")
                             for x in imp.iter() if x.find("vICMSST") is not None or x.find("vFCPST") is not None)
                         if imp is not None else 0))
        total = v_prod if VALOR_DAS_LINHAS == "produto" else round(v_prod + acrescimos, 2)
        itens.append({
            "cod": (prod.findtext("cProd") or "").strip() or str(n),
            "desc": " ".join((prod.findtext("xProd") or "").split()),
            "un": (prod.findtext("uCom") or "").strip(),
            "qtd": _num(prod, "qCom") or 1.0,
            "unit": _num(prod, "vUnCom"),
            "total": round(total, 2),
            "vprod": round(v_prod, 2),
            "acrescimos": round(acrescimos, 2),
        })
    tot = inf.find("total/ICMSTot")
    totais = {"produtos": _num(tot, "vProd"), "frete": _num(tot, "vFrete"), "st": _num(tot, "vST"),
              "ipi": _num(tot, "vIPI"), "nota": _num(tot, "vNF")}
    if VALOR_DAS_LINHAS != "produto" and totais["nota"]:
        ratear_ate(itens, totais["nota"])        # sobra de centavos/outros componentes: fecha no vNF
    return itens, emit, dest, chave, totais


def achar_xml_na_pasta(chave):
    """Procura na PASTA_XML um arquivo .xml com a chave de acesso no nome."""
    if not (PASTA_XML and chave):
        return None
    pasta = Path(PASTA_XML)
    if not pasta.is_dir():
        print(f"  ⚠ PASTA_XML não encontrada: {pasta}")
        return None
    for arq in pasta.rglob(f"*{chave}*.xml"):
        return arq
    return None


def chave_do_texto(texto):
    for m in re.finditer(r"(?<!\d)(?:\d{4}[ .]?){10}\d{4}(?!\d)", texto or ""):
        chave = so_digitos(m.group())
        if len(chave) == 44:
            return chave
    return None


NOTA_PDF = {}


def extrair_texto_pdf(arq):
    """Texto do PDF da DANFE (o mesmo que sairia no Ctrl+A / Ctrl+C). Devolve (texto, motivo_se_falhar)."""
    try:
        from pypdf import PdfReader
    except ImportError:
        return "", ("para ler o PDF direto falta a biblioteca pypdf - no terminal do PyCharm rode:  "
                    "pip install pypdf   (ou copie o texto do PDF com Ctrl+A / Ctrl+C)")
    try:
        leitor = PdfReader(str(arq))
        if leitor.is_encrypted:
            try:
                leitor.decrypt("")
            except Exception:
                return "", "o PDF está protegido por senha - copie o texto dele com Ctrl+A / Ctrl+C"
        texto = "\n".join((pg.extract_text() or "") for pg in leitor.pages)
    except Exception as e:
        return "", f"não consegui abrir o PDF ({e})"
    if len(re.sub(r"\s", "", texto)) < 50:
        return "", ("o PDF não tem texto (é uma imagem/escaneado) - peça o XML ao fornecedor "
                    "ou um PDF gerado pelo sistema dele")
    return texto, None


# ---------------------------------------------------------------- OCR local (PDF/imagem sem texto)
OCR_ESCALA = 3              # quanto ampliar a página antes do OCR (texto pequeno lê melhor ampliado)


def _imagens_do_arquivo(arq):
    """Páginas do PDF (ou a própria imagem) como imagens PIL, ampliadas para o OCR."""
    try:
        from PIL import Image
    except ImportError:
        raise RuntimeError("falta a biblioteca Pillow (imagens) - no terminal do PyCharm rode:  pip install pillow")
    if arq.suffix.lower() == ".pdf":
        try:
            import pypdfium2 as pdfium
        except ImportError:
            raise RuntimeError("falta a biblioteca pypdfium2 - rode:  pip install pypdfium2")
        pdf = pdfium.PdfDocument(str(arq))
        return [pdf[i].render(scale=OCR_ESCALA).to_pil().convert("RGB") for i in range(len(pdf))]
    img = Image.open(arq).convert("RGB")
    return [img.resize((img.width * 2, img.height * 2))] if max(img.size) < 1500 else [img]


PACOTES_OCR_WINDOWS = ("winrt-runtime winrt-Windows.Foundation winrt-Windows.Foundation.Collections "
                       "winrt-Windows.Media.Ocr winrt-Windows.Globalization winrt-Windows.Graphics.Imaging "
                       "winrt-Windows.Storage winrt-Windows.Storage.Streams")


def _ocr_windows(imagens):
    """OCR do próprio Windows (Windows.Media.Ocr) - roda no computador, nada é enviado para fora."""
    import asyncio
    import tempfile
    from concurrent.futures import ThreadPoolExecutor
    try:                                    # pacote "winsdk" (Python até 3.12)
        from winsdk.windows.media.ocr import OcrEngine
        from winsdk.windows.globalization import Language
        from winsdk.windows.graphics.imaging import BitmapDecoder
        from winsdk.windows.storage import StorageFile, FileAccessMode
    except ImportError:                     # pacotes "winrt-..." (Python 3.13/3.14)
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.globalization import Language
        from winrt.windows.graphics.imaging import BitmapDecoder
        from winrt.windows.storage import StorageFile, FileAccessMode

    motor = OcrEngine.try_create_from_language(Language("pt-BR"))
    if motor is None:
        motor = OcrEngine.try_create_from_user_profile_languages()
    if motor is None:
        raise RuntimeError("o Windows não tem idioma de OCR instalado (Configurações > Hora e idioma > "
                           "Idioma > Português (Brasil) > Opções > Reconhecimento óptico de caracteres)")
    try:
        limite = int(OcrEngine.max_image_dimension)
    except Exception:
        limite = 2600

    async def ler(caminho):
        arq = await StorageFile.get_file_from_path_async(caminho)
        fluxo = await arq.open_async(FileAccessMode.READ)
        decod = await BitmapDecoder.create_async(fluxo)
        bmp = await decod.get_software_bitmap_async()
        res = await motor.recognize_async(bmp)
        return "\n".join(linha.text for linha in res.lines)

    textos = []
    with tempfile.TemporaryDirectory() as pasta:
        for n, img in enumerate(imagens):
            if max(img.size) > limite:                    # o OCR do Windows tem tamanho máximo de imagem
                f = limite / max(img.size)
                img = img.resize((int(img.width * f), int(img.height * f)))
            caminho = str(Path(pasta) / f"pagina{n}.png")
            img.save(caminho)
            # o Playwright já usa um "event loop" nesta thread: o OCR roda numa thread separada
            with ThreadPoolExecutor(1) as ex:
                textos.append(ex.submit(lambda c=caminho: asyncio.run(ler(c))).result())
    return "\n".join(textos)


def ocr_arquivo(arq):
    """Texto de um PDF-imagem ou de uma foto da nota. Tenta o OCR do Windows; se não tiver, o Tesseract.
    Devolve (texto, motivo_se_falhar)."""
    try:
        imagens = _imagens_do_arquivo(arq)
    except Exception as e:
        return "", str(e)
    erros = []
    try:
        texto = _ocr_windows(imagens)
        if len(re.sub(r"\s", "", texto)) >= 50:
            try:
                (pasta_controle() / "ultimo_ocr.txt").write_text(texto, encoding="utf-8")
            except Exception:
                pass
            return texto, None
        erros.append("o OCR do Windows não reconheceu texto")
    except ImportError as e:
        erros.append(f"faltam as bibliotecas do OCR do Windows ({e}) - rode no terminal:  "
                     f'"{sys.executable}" -m pip install ' + PACOTES_OCR_WINDOWS)
    except Exception as e:
        erros.append(f"OCR do Windows: {e}")
    try:                                                   # reserva: Tesseract, se estiver instalado
        import pytesseract
        texto = "\n".join(pytesseract.image_to_string(i, lang="por+eng", config="--psm 4") for i in imagens)
        if len(re.sub(r"\s", "", texto)) >= 50:
            return texto, None
    except Exception:
        pass
    return "", "não consegui ler a imagem da nota: " + "; ".join(erros)


# ---------------------------------------------------------------- NFS-e (nota de serviço)
RE_CNPJ_OCR = re.compile(r"(?<!\d)\d{2}\s?[.,]?\s?\d{3}\s?[.,]?\s?\d{3}\s?/?\s?\d{4}\s?-?\s?\d{2}(?!\d)")


def eh_nfse(texto):
    return bool(re.search(r"NFS-?e|NOTA FISCAL (?:ELETR[ÔO]NICA )?DE SERVI|PRESTADOR DE SERVI|TOMADOR DE SERVI",
                          texto or "", re.I))


def _valor_ocr(txt):
    """'10.908,12' / '10,908.12' / '10908,12' -> 10908.12 (o OCR às vezes troca ponto e vírgula)."""
    t = re.sub(r"[^\d.,]", "", txt)
    if not t:
        return None
    dec = max(t.rfind(","), t.rfind("."))
    if dec >= 0 and len(t) - dec - 1 == 2:
        inteiro, cent = re.sub(r"[.,]", "", t[:dec]), t[dec + 1:]
        return float(f"{inteiro}.{cent}") if inteiro else None
    return float(re.sub(r"[.,]", "", t)) if t.replace(".", "").replace(",", "").isdigit() else None


def ler_nfse(texto, lojas):
    """Dados de uma NFS-e (texto copiado ou lido por OCR). Devolve dict com itens (1 linha), CNPJs e total."""
    t = texto.replace("\r", "")
    # CNPJs válidos, na ordem do texto (o dígito verificador filtra erros do OCR)
    cnpjs = []
    for m in RE_CNPJ_OCR.finditer(t):
        d = so_digitos(m.group())
        if len(d) == 14 and cnpj_valido(d) and d not in [c for c, _ in cnpjs]:
            cnpjs.append((d, m.start()))
    pos_tomador = (re.search(r"TOMADOR DE SERVI", t, re.I) or re.search(r"$", t)).start()
    loja = next((c for c, _ in cnpjs if c in lojas), None)
    prestador = next((c for c, p in cnpjs if c != loja and p < pos_tomador and not (CNPJ_RAIZ_EMPRESA and c.startswith(CNPJ_RAIZ_EMPRESA))),
                     None) or next((c for c, _ in cnpjs if c != loja and not (CNPJ_RAIZ_EMPRESA and c.startswith(CNPJ_RAIZ_EMPRESA))), None)

    # valor: rótulo e número podem vir na mesma linha ou em linhas separadas, e o OCR às vezes põe
    # espaços dentro do número ("10. 908, 12")
    NUM = r"(\d{1,3}(?:\s?[.,]\s?\d{3})*\s?[.,]\s?\d{2})(?!\d)"
    total = None
    for rot in (r"VALOR\s*TOTAL\s*DO\s*SERVI\w*", r"VALOR\s*TOTAL\s*DA\s*(?:NOTA|NFS-?e)",
                r"VALOR\s*TOTAL\s*COBRADO", r"VALOR\s*DOS\s*SERVI\w*", r"VALOR\s*L[ÍI]QUIDO",
                r"BASE\s*DE\s*C[ÁA]LCULO"):
        m = re.search(rot + r"[^\d]{0,40}?" + NUM, t, re.I)
        if m:
            total = _valor_ocr(m.group(1))
            if total:
                break
    if not total:
        # reserva: o valor que mais se repete na nota (total, base de cálculo e total cobrado costumam
        # ser o mesmo número); empate -> o maior. Sempre passa pela sua confirmação.
        valores = [v for v in (_valor_ocr(x) for x in re.findall(NUM, t)) if v and v >= 10]
        if valores:
            cont = Counter(valores)
            total = max(cont, key=lambda v: (cont[v], v))

    desc = ""
    linhas_txt = [l.strip() for l in t.split("\n") if l.strip()]
    m = re.search(r"DISCRIMINA\w*", t, re.I)
    depois = t[m.end():] if m else t
    cands = [l.strip() for l in depois.split("\n")
             if re.match(r"\s*SERVI[ÇC]O", l, re.I) and len(l.strip()) > 12]
    if cands:
        desc = cands[0]
    if not desc:                            # "SERVIÇO DE ..." no meio de uma linha
        m = re.search(r"SERVI[ÇC]O\s+DE\s+[^\n]{8,90}", depois, re.I)
        if m:
            desc = re.split(r"\s{2,}|\*\*|ATEN[ÇC][ÃA]O", m.group(0))[0].strip()
    if not desc:
        for k, l in enumerate(linhas_txt):
            if re.search(r"C[ÓO]DIGO\s+DO\s+SERVI", l, re.I):
                resto = re.sub(r".*C[ÓO]DIGO\s+DO\s+SERVI\w*\s*", "", l, flags=re.I)
                desc = resto if len(resto) > 10 else next((x for x in linhas_txt[k + 1:k + 3] if len(x) > 10), "")
                break
    if not desc:
        desc = "SERVIÇO (NFS-e)"
    numero = ""
    m = (re.search(RE_CNPJ_OCR.pattern + r"[\s|]+(0\d{5,9})\b", t)            # "CNPJ  00000121  código"
         or re.search(r"N[ÚU]MERO\s*DA\s*NOTA\D{0,60}?\b(0\d{5,9})\b", t, re.I)
         or re.search(r"N[ÚU]MERO DA NOTA\D{0,40}?\b(0*\d{3,9})\b", t, re.I)
         or re.search(r"N[ºo°]\s*(?:DA\s*)?(?:NOTA|NFS-?e)\s*:?\s*(\d{1,9})\b", t, re.I))
    if m:
        numero = str(int(m.group(1)))
    pedido = (re.search(r"PEDIDO\s*:?\s*(\d{6,12})", t, re.I) or [None, ""])[1]
    itens = []
    if total:
        itens = [{"cod": "NFSE", "desc": " ".join(desc.split())[:80], "un": "UN", "qtd": 1.0,
                  "unit": total, "total": total, "vprod": total, "acrescimos": 0.0}]
    return {"itens": itens, "prestador": prestador, "loja": loja, "total": total,
            "numero": numero, "pedido": pedido}


def obter_nota():
    """O que foi copiado pode ser: o texto da DANFE, o conteúdo do XML ou o CAMINHO de um .xml
    (no Explorador: Shift + clique direito no arquivo -> 'Copiar como caminho').
    Devolve ("xml", conteúdo) ou ("texto", texto)."""
    bruto = obter_texto_nota()
    caminho = bruto.strip().strip('"').strip("'").strip()
    if caminho.lower().startswith("file:///"):
        from urllib.parse import unquote
        caminho = unquote(caminho[8:])
    if caminho.lower().endswith(".xml") and "\n" not in caminho:
        arq = Path(caminho)
        if arq.is_file():
            print(f"Lendo o XML {arq.name}")
            return "xml", arq.read_bytes()              # bytes: respeita a codificação do arquivo
        print(f"  ⚠ Arquivo não encontrado: {arq}\n    Confira o caminho (a pasta e o nome, inclusive o '(1)').")
        return "texto", ""
    if re.search(r"\.(png|jpe?g|tiff?|bmp)$", caminho, re.I) and "\n" not in caminho:
        arq = Path(caminho)
        if not arq.is_file():
            print(f"  ⚠ Arquivo não encontrado: {arq}")
            return "texto", ""
        print(f"Lendo a imagem {arq.name} com OCR no próprio computador...")
        texto, motivo = ocr_arquivo(arq)
        if motivo:
            print(f"  ✖ {motivo}")
            return "falha", ""
        NOTA_PDF["ocr"] = True
        NOTA_PDF["imagem"] = (arq.suffix.lower(), arq.read_bytes())
        return "texto", texto
    if caminho.lower().endswith(".pdf") and "\n" not in caminho:
        arq = Path(caminho)
        if not arq.is_file():
            print(f"  ⚠ Arquivo não encontrado: {arq}\n    Confira o caminho (a pasta e o nome, inclusive o '(1)').")
            return "texto", ""
        print(f"Lendo o PDF {arq.name}")
        texto, motivo = extrair_texto_pdf(arq)
        if motivo and "não tem texto" in motivo:
            print("  O PDF é uma imagem - lendo com OCR no próprio computador (pode levar alguns segundos)...")
            texto, motivo = ocr_arquivo(arq)
            if not motivo:
                NOTA_PDF["ocr"] = True
        if motivo:
            print(f"  ✖ {motivo}")
            return "falha", ""
        NOTA_PDF["bytes"] = arq.read_bytes()             # guarda uma cópia na pasta da requisição
        bruto = texto                                    # daqui em diante é igual ao texto copiado
    if eh_xml_nfe(bruto):
        return "xml", bruto
    arq = achar_xml_na_pasta(chave_do_texto(bruto))
    if arq:
        print(f"XML desta nota encontrado na pasta: {arq.name} — usando o XML")
        return "xml", arq.read_bytes()
    return "texto", bruto


def ler_itens_nota(texto):
    texto = texto.replace("\r", "")
    itens = []
    for m in ITEM_NOTA.finditer(texto):
        itens.append({
            "cod": m["cod"],
            "desc": m["desc"].split("\n")[0].strip(),
            "un": m["un"],
            "qtd": num_br(m["qtd"]),
            "unit": num_br(m["unit"]),
            "total": num_br(m["total"]),
        })
    # linhas que parecem item mas não foram reconhecidas
    lidos = {i["cod"] for i in itens}
    spans = [m.span() for m in ITEM_NOTA.finditer(texto)]
    area = texto[spans[0][0]:spans[-1][1]] if spans else ""
    suspeitos = [m.group(0).replace("\n", " ").strip() for m in re.finditer(r"(?m)^(\d{4,10})[ \t]*\n?[ \t]*[A-Z][^\n]{5,}", area)
                 if m.group(1) not in lidos]
    return itens, suspeitos


def cnpjs_da_nota(texto, lojas):
    """Acha na nota inteira o CNPJ do fornecedor (emitente) e o da loja (destinatário).
    - emitente: vem da chave de acesso (44 dígitos, posições 7 a 20); se não houver chave,
      é o primeiro CNPJ que NÃO é da própria empresa e aparece antes do bloco da transportadora
    - loja: o CNPJ do texto que existe na planilha de regiões"""
    texto = texto.replace("\r", "")
    emitente = None
    for m in re.finditer(r"(?<!\d)(?:\d{4}[ .]?){10}\d{4}(?!\d)", texto):
        chave = so_digitos(m.group())
        if len(chave) == 44 and cnpj_valido(chave[6:20]):
            emitente = chave[6:20]
            break
    fim_transp = re.search(r"TRANSPORTADOR", texto, re.I)
    loja = None
    for m in re.finditer(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}", texto):
        d = so_digitos(m.group())
        if len(d) != 14 or not cnpj_valido(d):
            continue
        if d in lojas:
            loja = loja or d
        elif not emitente and not (fim_transp and m.start() > fim_transp.start()):
            emitente = d
    return emitente, loja


# ---------------------------------------------------------------- verificação no Coupa
JS_CENTROS = r"""() => {
    const re = /^\s*(?:\d{4}|ZZZZ)(?:\s*,\s*(?:\d{4}|ZZZZ)){5,}\s*$/;
    for (const el of document.querySelectorAll('p, span, div, td, li')) {
        if (el.children.length === 0 && re.test(el.textContent))
            return el.textContent.split(',').map(s => s.trim());
    }
    return null;
}"""


def busca_visivel(page):
    """Barra de busca que está visível na página, esteja ela no topo ou mais abaixo."""
    return page.locator(f"{SEL_BUSCA_BTN} >> visible=true").first


class VerificadorCoupa:
    """Abre o PLU no Coupa (numa aba separada) e confere:
       1) fornecedor = Indefinido   2) valor = 0,00   3) CC_EXIGIDO na lista de centros de custo."""

    def __init__(self, ctx):
        self.ctx = ctx
        self.page = ctx.new_page()
        ir_para_coupa(self.page)
        self.cache = {}

    def _ir_para_inicio(self):
        """Volta para a tela inicial e espera a barra de busca aparecer (até 2 tentativas)."""
        p = self.page
        for _ in range(2):
            ir_para_coupa(p)
            try:
                busca_visivel(p).wait_for(state="visible", timeout=TIMEOUT)
                return True
            except Exception:
                continue
        return False

    def _buscar(self, plu):
        """Volta ao início, abre a busca, APAGA o que estiver no campo, digita o PLU e dá Enter."""
        p = self.page
        for _ in range(2):
            if not self._ir_para_inicio():
                continue
            botao = busca_visivel(p)
            botao.scroll_into_view_if_needed()
            botao.click()
            try:
                p.wait_for_function(
                    "document.activeElement && ['INPUT','TEXTAREA'].includes(document.activeElement.tagName)",
                    timeout=5000,
                )
            except Exception:
                continue          # a busca não abriu: volta ao início e tenta de novo
            p.keyboard.press("Control+a")
            p.keyboard.press("Backspace")
            p.keyboard.type(plu, delay=40)
            digitado = p.evaluate("document.activeElement ? document.activeElement.value : ''")
            if digitado.strip() != plu:
                continue          # não ficou só o PLU no campo: volta ao início e tenta de novo
            p.keyboard.press("Enter")
            return                # pesquisa feita com o PLU certo no campo
        # não conseguiu: salva um print da tela para diagnóstico
        arquivo = pasta_atual() / f"falha_busca_{plu}.png"
        try:
            p.screenshot(path=str(arquivo), full_page=True)
        except Exception:
            pass
        raise RuntimeError(f"não consegui abrir a busca do Coupa (print salvo em {arquivo.name})")

    def _abrir_resultado(self, plu):
        """Clica no resultado 'PLU | descrição' e devolve a página do item (ou None se não achou)."""
        p = self.page
        link = p.locator(SEL_RESULTADO_PLU.format(plu=plu)).first
        try:
            link.wait_for(timeout=ESPERA_RESULTADO)
        except Exception:
            return None
        abas_antes = len(self.ctx.pages)
        link.click()
        try:
            p.locator(SEL_FORNECEDOR_ITEM).first.wait_for(state="attached", timeout=TIMEOUT)
            return p
        except Exception:
            if len(self.ctx.pages) > abas_antes:      # abriu em outra aba
                nova = self.ctx.pages[-1]
                nova.locator(SEL_FORNECEDOR_ITEM).first.wait_for(state="attached", timeout=TIMEOUT)
                return nova
            raise

    def verificar(self, plu):
        """Retorna (True, "ok") | (False, motivo) | (None, erro técnico)."""
        if plu in self.cache:
            return self.cache[plu]
        try:
            # se a pesquisa foi feita (PLU certo no campo + Enter) e não apareceu, não existe
            self._buscar(plu)
            p = self._abrir_resultado(plu)
            if p is None:
                return self._guardar(plu, (False, "PLU não encontrado na busca do Coupa"))

            sel = p.locator(SEL_FORNECEDOR_ITEM).first
            dados = json.loads(sel.get_attribute("data-ui") or "{}")

            motivos = []
            # 1) fornecedor
            itens_forn = dados.get("supplier_items") or []
            fornecedores = [str(si.get("supplier") or "").strip() for si in itens_forn]
            if not fornecedores:
                fornecedores = [sel.locator("option[selected]").first.inner_text().strip()]
            atrelados = [f for f in fornecedores if f and f.lower() != "indefinido"]
            if atrelados:
                motivos.append("fornecedor atrelado: " + ", ".join(sorted(set(atrelados))))

            # 2) valor
            valores = [float((dados.get("price") or {}).get("amount") or 0)]
            for si in itens_forn:
                for u in si.get("uoms") or []:
                    try:
                        valores.append(num_br(str(u.get("unit_price") or "0,00")))
                    except ValueError:
                        pass
            if max(valores) > 0:
                motivos.append(f"valor atrelado: R$ {br(max(valores))}")

            # 3) centro de custo
            try:
                p.wait_for_function(JS_CENTROS, timeout=5000)
            except Exception:
                pass
            centros = p.evaluate(JS_CENTROS)
            if centros is None:
                return self._guardar(plu, (None, "não achei a lista de centros de custo na página"))
            if CC_EXIGIDO and CC_EXIGIDO not in centros:
                motivos.append(f"sem o centro de custo {CC_EXIGIDO}")

            if p is not self.page:
                p.close()
            return self._guardar(plu, (False, "; ".join(motivos)) if motivos else (True, "ok"))
        except Exception as e:
            return None, f"erro na verificação: {str(e).splitlines()[0]}"

    def _guardar(self, plu, resultado):
        self.cache[plu] = resultado
        return resultado

    def fechar(self):
        try:
            self.page.close()
        except Exception:
            pass


# ---------------------------------------------------------------- loja e fornecedor
def cnpj_valido(cnpj):
    """Confere os dígitos verificadores do CNPJ (pega erro de digitação)."""
    d = so_digitos(cnpj)
    if len(d) != 14 or d == d[0] * 14:
        return False
    for tam in (12, 13):
        pesos = list(range(tam - 7, 1, -1)) + list(range(9, 1, -1))
        soma = sum(int(n) * p for n, p in zip(d[:tam], pesos))
        dv = 0 if soma % 11 < 2 else 11 - soma % 11
        if dv != int(d[tam]):
            return False
    return True


def fmt_cnpj(d):
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"


def arquivo_regioes():
    """CAMINHO_REGIOES pode ser o arquivo .xlsb ou a PASTA onde ele fica. Se for pasta, usa o
    .xlsb de regiões mais recente dela (assim a troca de versão _9 -> _10 não exige mexer no código)."""
    caminho = Path(CAMINHO_REGIOES)
    if caminho.is_dir():
        candidatos = [f for f in caminho.glob("*.xlsb")
                      if "regi" in f.name.lower() and not f.name.startswith("~$")]
        if not candidatos:
            raise FileNotFoundError(f"nenhum arquivo .xlsb de regiões na pasta {caminho}")
        return max(candidatos, key=lambda f: f.stat().st_mtime)
    return caminho


def carregar_lojas():
    """{CNPJ: [(cc, nome da loja, status, segmento), ...]} da aba Regiões da planilha de regiões.
    Ler o .xlsb demora (~25 s), então o resultado fica guardado em regioes_cache.json
    e só é relido quando a planilha de regiões for alterada."""
    cache = pasta_controle() / "regioes_cache.json"
    arquivo = arquivo_regioes()
    mtime = arquivo.stat().st_mtime
    if cache.exists():
        try:
            dados = json.loads(cache.read_text(encoding="utf-8"))
            if dados.get("mtime") == mtime and dados.get("arquivo") == str(arquivo):
                return {k: [tuple(x) for x in v] for k, v in dados["lojas"].items()}
        except Exception:
            pass

    print(f"Lendo a planilha de regiões {arquivo.name} (só na primeira vez ou quando ela mudar)...")
    from pyxlsb import open_workbook
    lojas = {}
    with open_workbook(str(arquivo)) as wb:
        with wb.get_sheet("Regiões") as sh:
            linhas = sh.rows()
            cab = [str(c.v or "").strip() for c in next(linhas)]
            i_cc, i_loja = cab.index("CC Texto"), cab.index("Loja")
            i_status, i_cnpj = cab.index("Ultimo Status da Loja"), cab.index("CNPJ")
            i_seg = cab.index("DESC_SEGMENTO") if "DESC_SEGMENTO" in cab else None
            for row in linhas:
                v = [c.v for c in row]
                if len(v) <= i_cnpj or not v[i_cc]:
                    continue
                cc = v[i_cc]
                cc = str(int(cc)).zfill(4) if isinstance(cc, float) else str(cc).strip().zfill(4)
                cnpj = so_digitos(v[i_cnpj])
                seg = str(v[i_seg] or "").strip() if i_seg is not None and len(v) > i_seg else ""
                if len(cnpj) == 14:
                    lojas.setdefault(cnpj, []).append(
                        (cc, str(v[i_loja] or "").strip(), str(v[i_status] or "").strip(), seg))
    try:
        cache.write_text(json.dumps({"mtime": mtime, "arquivo": str(arquivo), "lojas": lojas}, ensure_ascii=False),
                         encoding="utf-8")
    except Exception:
        pass
    return lojas


def perguntar_cnpj(msg):
    while True:
        d = normalizar_cnpj(input(msg))
        if cnpj_valido(d):
            return d
        print("  ✖ CNPJ inválido (confira os números).")


def carregar_lojas_ou_sair():
    try:
        return carregar_lojas()
    except ModuleNotFoundError:
        raise SystemExit("\n✖ Falta a biblioteca que lê a planilha de regiões (.xlsb).\n"
                         "  No terminal do PyCharm, com o .venv ativo, rode:  pip install pyxlsb\n"
                         "  e execute a macro de novo.")
    except FileNotFoundError:
        raise SystemExit(f"\n✖ Planilha de regiões não encontrada em:\n  {CAMINHO_REGIOES}\n"
                         "  Coloque o arquivo nesse caminho ou ajuste CAMINHO_REGIOES no .env.")
    except Exception as e:
        raise SystemExit(f"\n✖ Não consegui ler a planilha de regiões: {e}")


def escolher_cc(lojas, cnpj_loja, mostrar_lista=False):
    """Centro de custo de um CNPJ da base de regiões. Retorna (cc, nome) ou (None, None).
    Com um só CC (ou um só de segmento LOJA) escolhe sozinho; senão mostra a lista."""
    opcoes = lojas.get(cnpj_loja, [])
    if not opcoes:
        return None, None
    abertas = [o for o in opcoes if o[2].lower() == "loja aberta"]
    lista = abertas or opcoes
    so_loja = [o for o in lista if o[3].upper() == "LOJA"]
    if not mostrar_lista and (len(lista) == 1 or len(so_loja) == 1):
        c, nome, _, _ = lista[0] if len(lista) == 1 else so_loja[0]
        return c, nome
    if not abertas:
        print("  ⚠ Nenhum centro de custo ABERTO para esse CNPJ; mostrando os fechados.")
    print(f"  Esse CNPJ tem {len(lista)} centro(s) de custo:")
    for k, (c, nome, status, seg) in enumerate(lista, start=1):
        print(f"   [{k}] {nome} ({status}{' · ' + seg if seg else ''})")
    while True:
        r = input("  Qual deles? (número): ").strip()
        if r.isdigit() and 1 <= int(r) <= len(lista):
            c, nome, _, _ = lista[int(r) - 1]
            return c, nome


def perguntar_loja(lojas):
    while True:
        cnpj_loja = perguntar_cnpj("CNPJ da loja (destinatário da nota): ")
        cc, nome = escolher_cc(lojas, cnpj_loja, mostrar_lista=len(lojas.get(cnpj_loja, [])) > 1)
        if cc:
            return cnpj_loja, cc, nome
        print(f"  ✖ CNPJ {fmt_cnpj(cnpj_loja)} não está na planilha de regiões. Confira e digite de novo.")


def mostrar_itens(itens, suspeitos):
    print(f"\n{len(itens)} item(ns) lidos | total R$ {br(sum(i['total'] for i in itens))}")
    for i in itens:
        ok = "" if abs(i["qtd"] * i["unit"] - i.get("vprod", i["total"])) < 0.01 else "  ⚠ qtd × unitário ≠ total"
        print(f"  {i['cod']:>8}  {i['desc'][:32]:<32} {i['qtd']:>7g} {i['un']:<3} R$ {br(i['total']):>10}{ok}")
    if suspeitos:
        print("\n⚠ Estas linhas parecem itens mas NÃO foram lidas — confira a nota:")
        for s in suspeitos:
            print("   ", s)


def ler_nota_completa(lojas):
    """Lê a nota copiada, acha itens, fornecedor e loja, e pede só a confirmação.
    Retorna (itens, cnpj_fornecedor, cc) ou None."""
    tipo, texto = obter_nota()
    if tipo == "falha":                     # o motivo já foi mostrado (ex.: biblioteca faltando)
        return None
    NOTA_ATUAL.clear()
    NOTA_ATUAL.update({"tipo": tipo, "conteudo": texto, "chave": None, "pdf": NOTA_PDF.pop("bytes", None),
                       "ocr": NOTA_PDF.pop("ocr", False), "imagem": NOTA_PDF.pop("imagem", None)})
    if tipo == "xml":
        try:
            itens, cnpj_forn, cnpj_loja, chave, totais = ler_nota_xml(texto)
            NOTA_ATUAL["chave"] = chave
            total_ref = totais["produtos"] if VALOR_DAS_LINHAS == "produto" else totais["nota"]
        except Exception as e:
            print(f"✖ Esse XML não serve: {e}")
            return None
        if not itens:
            print("O XML não tem itens.")
            return None
        mostrar_itens(itens, [])
        soma_prod = round(sum(i["vprod"] for i in itens), 2)
        print(f"  (XML) produtos R$ {brl(totais['produtos'])} · ST R$ {brl(totais['st'])} · IPI R$ {brl(totais['ipi'])} · "
              f"frete R$ {brl(totais['frete'])} · TOTAL DA NOTA R$ {brl(totais['nota'])}")
        if VALOR_DAS_LINHAS != "produto":
            print(f"  ✔ As linhas vão somar o total da nota: R$ {brl(round(sum(i['total'] for i in itens), 2))}")
        if abs(soma_prod - totais["produtos"]) >= 0.01:
            print(f"  ⚠ soma dos itens (R$ {br(soma_prod)}) ≠ total dos produtos do XML — confira")
        extras = round(sum(i["acrescimos"] for i in itens), 2)
        if extras and VALOR_DAS_LINHAS == "produto":
            print(f"  ℹ Os itens têm R$ {br(extras)} de frete/impostos/outros que NÃO entram nas linhas "
                  "(VALOR_DAS_LINHAS = \"produto\").")
        if cnpj_loja not in lojas:
            cnpj_loja = None
    elif eh_nfse(texto):
        lido_ocr = NOTA_ATUAL.get("ocr")
        nf = ler_nfse(texto, lojas)
        print("\n📄 Nota de SERVIÇO (NFS-e)" + (" — lida por OCR, confira os números" if lido_ocr else ""))
        print(f"  Nº {nf['numero'] or '?'}  ·  serviço: {nf['itens'][0]['desc'] if nf['itens'] else '?'}"
              + (f"  ·  pedido citado: {nf['pedido']}" if nf["pedido"] else ""))
        valor = nf["total"]
        r = input(f"  Valor total do serviço: R$ {brl(valor) if valor else '?'} — Enter confirma ou digite o valor certo: "
                  ).strip()
        if r:
            valor = _valor_ocr(r)
        while not valor:
            valor = _valor_ocr(input("  Digite o valor total do serviço (ex.: 10.908,12): "))
        itens = [{"cod": "NFSE", "desc": (nf["itens"][0]["desc"] if nf["itens"] else "SERVIÇO (NFS-e)"),
                  "un": "UN", "qtd": 1.0, "unit": valor, "total": valor, "vprod": valor, "acrescimos": 0.0}]
        total_ref = valor
        cnpj_forn, cnpj_loja = nf["prestador"], nf["loja"]
        NOTA_ATUAL["numero_nf"] = nf["numero"]
    else:
        itens, suspeitos = ler_itens_nota(texto)
        if not itens:
            print("Nenhum item reconhecido no texto. Confira se copiou a nota com a tabela de produtos "
                  "(ou copie o XML / o caminho do arquivo .xml).")
            return None
        mostrar_itens(itens, suspeitos)
        total_nota = total_produtos_da_nota(texto)
        total_ref = total_nota
        soma = round(sum(i["total"] for i in itens), 2)
        if total_nota is not None:
            if abs(total_nota - soma) < 0.01:
                print(f"  ✔ soma dos itens confere com o total dos produtos da nota (R$ {br(total_nota)})")
            else:
                print(f"\n  ⚠ ATENÇÃO: soma dos itens lidos R$ {br(soma)} ≠ total dos produtos da nota "
                      f"R$ {br(total_nota)} (diferença R$ {br(total_nota - soma)}). Algum item não foi lido certo.")
                if input("  Continuar mesmo assim? [s/n]: ").strip().lower() != "s":
                    return None
        if VALOR_DAS_LINHAS != "produto":
            total_geral = total_geral_da_nota(texto)
            if total_geral is None:
                print("  ⚠ Não achei o 'VALOR TOTAL DA NOTA' no texto: as linhas vão só com o valor dos produtos.\n"
                      "    Confira o total com a nota (ou use o XML, que traz tudo certinho).")
            elif total_geral - soma >= 0.01:
                if total_nota is None or abs(total_nota - soma) < 0.01:
                    dif = ratear_ate(itens, total_geral)
                    print(f"  ℹ Diferença de R$ {brl(dif)} entre o total da nota (R$ {brl(total_geral)}) e os produtos "
                          "(ST/frete/IPI/outros) rateada entre as linhas, proporcional ao valor de cada item.")
                    total_ref = total_geral
                else:
                    print("  ⚠ Os itens não bateram com o total dos produtos - não vou ratear a diferença até o "
                          "total da nota. Confira os valores.")
        cnpj_forn, cnpj_loja = cnpjs_da_nota(texto, lojas)
        NOTA_ATUAL["chave"] = chave_do_texto(texto)

    cc, nome_loja = escolher_cc(lojas, cnpj_loja) if cnpj_loja else (None, None)

    while True:
        print("\n  Fornecedor: " + (fmt_cnpj(cnpj_forn) if cnpj_forn else "✖ não encontrado na nota"))
        print("  Loja:       " + (f"{fmt_cnpj(cnpj_loja)} → CC {nome_loja}" if cc else "✖ não encontrada na nota"))
        if cnpj_forn and cc:
            r = input("\nAs informações estão corretas? [s] sim | [f] corrigir fornecedor | "
                      "[l] corrigir loja | [c] cancelar: ").strip().lower()
            if r == "s":
                NOTA_ATUAL.update({"cnpj_forn": cnpj_forn, "cc": cc, "loja": nome_loja})
                return itens, cnpj_forn, cc, total_ref
            if r == "c":
                return None
        else:
            r = "f" if not cnpj_forn else "l"
        if r == "f":
            cnpj_forn = perguntar_cnpj("CNPJ do fornecedor (emitente da nota): ")
        elif r == "l":
            cnpj_loja, cc, nome_loja = perguntar_loja(lojas)

# ---------------------------------------------------------------- sugestão de PLU
STOP = {"DE", "DA", "DO", "DAS", "DOS", "E", "C", "P", "X", "COM", "PARA", "CJ", "CONJ", "CJTO", "JG",
        "KIT", "UN", "PC", "PECA", "MM", "COD", "CODIGO", "R", "F", "S", "D", "POL", "ORIG"}


def _norm(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper()


def _tokens(s):
    return [t for t in re.findall(r"[A-Z0-9]+", _norm(s).replace("'", "")) if t not in STOP and len(t) >= 2]


def _palavra(t):
    return t.isalpha() and len(t) >= 3


def _casa(t, ts):
    if t in ts:
        return 1.0
    if len(t) >= 5 and t.isalpha() and any(len(u) >= 5 and u[:5] == t[:5] for u in ts):
        return 0.8          # VALVULAS ~ VALVULA
    return 0.0


class Sugeridor:
    """Compara a descrição da nota com as descrições da aba PLUs.
    A primeira palavra (BIELA, JUNTA, FLANGE...) precisa bater; códigos e números pesam menos."""

    def __init__(self, plus):
        self.plus = plus
        self.toks = {p: set(_tokens(d["desc"])) for p, d in plus.items()}
        self.primeira = {p: next(iter(_tokens(d["desc"])), "") for p, d in plus.items()}
        freq = Counter(t for ts in self.toks.values() for t in ts)
        self.idf = {t: math.log(len(plus) / c) for t, c in freq.items()}

    def sugerir(self, desc, n=MAX_SUGESTOES):
        alvo = set(_tokens(desc))
        palavras = [t for t in _tokens(desc) if _palavra(t)]
        nucleo = palavras[0] if palavras else None
        res = []
        for p, ts in self.toks.items():
            if nucleo and not _casa(nucleo, ts):
                continue
            s = sum(_casa(t, ts) * (1.0 if _palavra(t) else 0.3) * self.idf.get(t, 2.0) for t in alvo)
            if nucleo and self.primeira[p] == nucleo:
                s *= 1.5    # descrição do PLU começa com a mesma palavra (OLEO..., JUNTA...)
            if s > 0:
                res.append((s, self.plus[p]["pedidos"], p))
        res.sort(reverse=True)
        return [p for _, _, p in res[:n]]

    def baratos(self, excluir, n):
        """Sorteia n PLUs entre os de menor valor (usados 3+ vezes no último ano), do mais barato ao mais caro."""
        limite = datetime.now() - timedelta(days=365)
        pool = [p for p, d in self.plus.items()
                if p not in excluir and d["preco"] is not None and d["preco"] >= 1
                and not d["desc"].startswith("(sem descrição")
                and d["pedidos"] >= 3 and isinstance(d["data"], datetime) and d["data"] >= limite]
        pool.sort(key=lambda p: self.plus[p]["preco"])
        escolhidos = random.sample(pool[:POOL_BARATOS], min(n, len(pool[:POOL_BARATOS])))
        return sorted(escolhidos, key=lambda p: self.plus[p]["preco"])


def por_preco(plus, lista):
    """Ordena do menor para o maior último preço (sem preço vai para o fim)."""
    return sorted(lista, key=lambda p: (plus.get(p, {}).get("preco") is None,
                                        plus.get(p, {}).get("preco") or 0))


def info_plu(plus, plu):
    d = plus.get(plu)
    if not d:
        return f"PLU {plu} (não está na aba PLUs)"
    preco = f" · último R$ {br(d['preco'])}" if d["preco"] is not None else ""
    return f"PLU {plu} {d['desc']} ({d['um']} · {d['pedidos']} pedidos{preco})"


def _perguntar_plu(plus, usados, verificador, novos_bloq, msg):
    """Pede um PLU digitado; se houver verificador, confere antes de aceitar."""
    while True:
        r = input(msg).strip().lower()
        if r == "p":
            return None
        if not r.isdigit():
            continue
        if r in usados:
            print(f"  ✖ PLU {r} já está no item {usados[r]} desta nota — não pode repetir.")
            continue
        print(f"  → {info_plu(plus, r)}")
        if verificador:
            ok, motivo = verificador.verificar(r)
            if ok is False:
                novos_bloq[r] = motivo
                print(f"  ✖ Reprovado no Coupa: {motivo}")
                if input("  Usar mesmo assim? [s/n]: ").strip().lower() != "s":
                    continue
            elif ok is None:
                print(f"  ⚠ Não consegui verificar ({motivo}) — usando porque você digitou.")
        return r


def escolher_plu(item, n_item, total_itens, plus, sugeridor, plu_anterior, usados, confiaveis,
                 bloqueados, verificador, novos_bloq):
    """Escolhe o PLU de um item da nota.
    Retorna (plu, generico, como) - como = "auto" (confiável), "verificado", "sim" ou "digitado";
    ou (None, False, None) se o item for pulado.
    usados = {PLU: código do item} já escolhidos nesta nota - cada PLU só pode ir em uma linha."""
    print(f"\n── Item {n_item}/{total_itens} ── {item['cod']} {item['desc']} | "
          f"{item['qtd']:g} {item['un']} | R$ {br(item['total'])}")

    fora = set(usados) | set(bloqueados) | set(novos_bloq)
    anterior = None
    if plu_anterior and plu_anterior in usados:
        print(f"  (PLU {plu_anterior}, usado antes para este código, já está no item {usados[plu_anterior]} desta nota)")
    elif plu_anterior and plu_anterior not in fora:
        anterior = plu_anterior

    parecidos = [p for p in sugeridor.sugerir(item["desc"], MAX_SUGESTOES + len(fora))
                 if p not in (plu_anterior, PLU_GENERICO) and p not in fora][:MAX_SUGESTOES]
    parecidos = por_preco(plus, parecidos)
    outros_confiaveis = por_preco(plus, [p for p in confiaveis
                                         if p not in fora and p not in parecidos
                                         and p not in (anterior, PLU_GENERICO)])

    # 1) confiável: usa direto, sem verificar
    if AUTO_CONFIAVEL:
        ordem = ([anterior] if anterior else []) + parecidos
        if not CONFIAVEL_SO_PARECIDOS:
            ordem += outros_confiaveis
        for p in ordem:
            if p in confiaveis:
                relacao = "" if (p == anterior or p in parecidos) else " · sem relação com o item"
                print(f"  ✔ {info_plu(plus, p)}  [confiável{relacao} — usado automaticamente]")
                return p, p == PLU_GENERICO, "auto"

    # 2) fila de candidatos
    marca = lambda p: " · ✔ confiável" if p in confiaveis else ""
    fila = []
    if anterior:
        fila.append((anterior, "usado antes para este código" + marca(anterior)))
    fila += [(p, marca(p).lstrip(" ·")) for p in parecidos]
    if not parecidos:
        print("  Nenhum PLU parecido na aba PLUs.")
    if not AUTO_CONFIAVEL:
        fila += [(p, "✔ confiável · sem relação com o item") for p in outros_confiaveis[:MAX_SUGESTOES]]
    faltam = MAX_SUGESTOES - len(parecidos)
    if COMPLETAR_COM_BARATOS and faltam > 0:
        excluir = {p for p, _ in fila} | {PLU_GENERICO} | fora
        fila += [(p, "sem relação com o item · valor baixo") for p in sugeridor.baratos(excluir, faltam)]
    if PLU_GENERICO and PLU_GENERICO not in fora:
        extra = " → 1 un com o valor total" if QTD_1_NO_GENERICO else ""
        fila.append((PLU_GENERICO, f"genérico{extra}{marca(PLU_GENERICO)}"))

    # 3a) verificando no Coupa: o primeiro que passar é usado
    if verificador:
        def candidatos():
            vistos = set()
            for plu, obs in fila:
                vistos.add(plu)
                yield plu, obs
            # nenhuma passou: segue sorteando PLUs de valor baixo até o limite
            while len(vistos) < MAX_VERIFICACOES_ITEM:
                extra = sugeridor.baratos(vistos | fora | {PLU_GENERICO} | set(novos_bloq), 1)
                if not extra:
                    return
                vistos.add(extra[0])
                yield extra[0], "sem relação com o item · valor baixo"

        for plu, obs in candidatos():
            ponto_de_pausa(verificador.page, "verificação de PLU")
            obs = f"  [{obs}]" if obs else ""
            print(f"  🔎 {info_plu(plus, plu)}{obs}")
            ok, motivo = verificador.verificar(plu)
            if ok:
                print("     ✔ liberado (sem fornecedor, sem valor" + (f", com CC {CC_EXIGIDO}" if CC_EXIGIDO else "") + ")")
                return plu, plu == PLU_GENERICO, "verificado"
            if ok is False:
                novos_bloq[plu] = motivo
                print(f"     ✖ {motivo}")
                continue
            # erro técnico: pergunta para você
            print(f"     ⚠ {motivo}")
            r = input("     Esse PLU pode ser utilizado? [s] sim | [n] próximo | [p] pular item: ").strip().lower()
            if r == "s":
                return plu, plu == PLU_GENERICO, "sim"
            if r == "p":
                return None, False, None
        plu = _perguntar_plu(plus, usados, verificador, novos_bloq,
                             "  Nenhuma sugestão passou. Digite um PLU ou [p] para pular o item: ")
        return (plu, plu == PLU_GENERICO, "digitado") if plu else (None, False, None)

    # 3b) sem verificação: pergunta um por um
    for k, (plu, obs) in enumerate(fila, start=1):
        obs = f"  [{obs}]" if obs else ""
        print(f"  Sugestão {k}/{len(fila)}: {info_plu(plus, plu)}{obs}")
        while True:
            r = input("  Esse PLU pode ser utilizado? [s] sim | [n] próxima | [número] outro PLU | [p] pular item: ").strip().lower()
            if r.isdigit() and r in usados:
                print(f"  ✖ PLU {r} já está no item {usados[r]} desta nota — não pode repetir.")
                continue
            break
        if r == "s":
            return plu, plu == PLU_GENERICO, "sim"
        if r == "p":
            return None, False, None
        if r.isdigit():
            print(f"  → {info_plu(plus, r)}")
            return r, r == PLU_GENERICO, "digitado"

    plu = _perguntar_plu(plus, usados, None, novos_bloq, "  Sem mais sugestões. Digite o PLU ou [p] para pular o item: ")
    return (plu, plu == PLU_GENERICO, "digitado") if plu else (None, False, None)


def montar_linhas_da_nota(wb, itens, cnpj, cc, verificador=None, permitir_repetido=False):
    # CNPJ do fornecedor e centro de custo já vieram do início (iguais para todas as linhas)

    # PLU de cada item: a macro sugere e você confirma
    plus = carregar_plus(wb)
    sugeridor = Sugeridor(plus)
    de_para = carregar_de_para(wb)
    confiaveis = carregar_confiaveis(wb)
    bloqueados = carregar_bloqueados(wb)
    novos_bloq = {}
    if confiaveis:
        print(f"\n{len(confiaveis)} PLU(s) confiáveis na planilha"
              + (" — serão usados automaticamente quando couberem." if AUTO_CONFIAVEL else "."))
    novos, novos_conf, linhas, usados = [], [], [], {}
    for n, i in enumerate(itens, start=1):
        anterior = de_para.get((cnpj, i["cod"]))
        # com PLU repetido liberado (perfil), nenhum PLU fica "ocupado" por outro item
        plu, generico, como = escolher_plu(i, n, len(itens), plus, sugeridor, anterior,
                                           {} if permitir_repetido else usados, confiaveis,
                                           bloqueados, verificador, novos_bloq)
        if not plu:
            print("  (item pulado)")
            continue
        usados[plu] = i["cod"]
        if como in ("sim", "digitado", "verificado") and plu not in confiaveis and plu not in novos_conf:
            novos_conf.append(plu)
        if plu != anterior:
            novos.append([cnpj, i["cod"], i["desc"], plu])
        usar_qtd_1 = generico and QTD_1_NO_GENERICO
        linhas.append({
            "origem": f"item {i['cod']} da nota",
            "plu": plu,
            "qtd": 1.0 if usar_qtd_1 else i["qtd"],
            "total": i["total"],
            "cnpj": cnpj,
            "cc": cc,
            "auto": como in ("auto", "verificado"),
        })

    salvar_aprendizado(novos, novos_conf, plus, novos_bloq)

    print("\nResumo do que será lançado:")
    for l in linhas:
        auto = "  ✔ auto" if l.get("auto") else ""
        print(f"  PLU {l['plu']:<9} {plus.get(l['plu'], {}).get('desc', '')[:30]:<30} qtd {l['qtd']:g}  R$ {br(l['total'])}{auto}")
    print(f"  CNPJ {cnpj} | CC {cc} | total R$ {br(sum(l['total'] for l in linhas))}")
    if input("\nConfirma? [s/n]: ").strip().lower() != "s":
        return []
    return linhas


# ---------------------------------------------------------------- perfil de cabeçalho
def texto_celula(v):
    """Valor de célula como texto (número inteiro sem ',0')."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def escolher_perfil():
    """Pergunta o perfil do cabeçalho. Retorna (nome, dados) ou (None, None) para manual."""
    nomes = list(PERFIS)
    if not nomes:
        print(f"\nCabeçalho da requisição: sem {PERFIS_JSON}, preenchimento manual.")
        return None, None
    print("\nCabeçalho da requisição:")
    for k, nome in enumerate(nomes, start=1):
        print(f"  [{k}] {nome}" + ("  (padrão)" if k == 1 else ""))
    print("  [m] Preencher manualmente")
    while True:
        r = input("Escolha (Enter = 1): ").strip().lower()
        if r == "":
            r = "1"
        if r == "m":
            return None, None
        if r.isdigit() and 1 <= int(r) <= len(nomes):
            nome = nomes[int(r) - 1]
            return nome, PERFIS[nome]


def valor_coluna_g(wb):
    """1º valor preenchido na coluna G da aba Linhas (nº usado na conta do perfil)."""
    for (g,) in wb["Linhas"].iter_rows(min_row=2, min_col=7, max_col=7, values_only=True):
        if texto_celula(g):
            return texto_celula(g)
    return ""


def valor_coluna_h(wb):
    """1º texto preenchido na coluna H da aba Linhas (campo TR do perfil)."""
    for (h,) in wb["Linhas"].iter_rows(min_row=2, min_col=8, max_col=8, values_only=True):
        if texto_celula(h):
            return texto_celula(h)
    return ""


def confirmar_valor_conta(perfil, sugestao):
    """Confirma (ou pede) o nº que vai na conta das linhas no lugar do centro de custo."""
    tipo = perfil.get("conta_tipo", TIPO_CONTA)
    if sugestao:
        r = input(f"Nº para a conta ({tipo}) — coluna G da planilha: {sugestao}. Enter confirma ou digite outro: ").strip()
        return r or sugestao
    while True:
        r = input(f"A coluna G da planilha está vazia. Digite o nº para a conta ({tipo}): ").strip()
        if r:
            return r


def aplicar_perfil_nas_linhas(linhas, perfil, valor_padrao, tr_padrao=""):
    if not perfil:
        return
    for l in linhas:
        if perfil.get("campo_tr") == "coluna H":
            l["tr"] = (l.get("h") or tr_padrao or "").strip()
        l["conta_tipo"] = perfil.get("conta_tipo", TIPO_CONTA)
        l["conta_busca"] = perfil.get("conta_busca", l["conta_tipo"])
        if perfil.get("conta_valor") == "coluna G":
            l["conta_valor"] = l.get("g") or valor_padrao


def _selecionar_opcao(page, id_select, texto):
    """Escolhe uma opção de <select> pelo texto (funciona também em select múltiplo/escondido)."""
    loc = page.locator(f"#{id_select}")
    loc.wait_for(state="attached", timeout=TIMEOUT)
    try:
        loc.select_option(label=texto, timeout=5000)
    except Exception:
        ok = loc.evaluate("""(sel, texto) => {
            const norm = t => (t || '').replace(/\\s+/g, ' ').trim().toUpperCase();
            let achou = false;
            for (const o of sel.options) {
                const igual = norm(o.text) === norm(texto) || norm(o.value) === norm(texto);
                if (igual) achou = true;
                o.selected = igual;
            }
            sel.dispatchEvent(new Event('change', {bubbles: true}));
            if (window.jQuery) window.jQuery(sel).trigger('chosen:updated');
            return achou;
        }""", texto)
        if not ok:
            raise RuntimeError(f"opção '{texto}' não existe em #{id_select}")
    marcadas = loc.evaluate("s => [...s.selectedOptions].map(o => o.text.trim().toUpperCase())")
    if texto.strip().upper() not in marcadas:
        raise RuntimeError(f"'{texto}' não ficou selecionado em #{id_select}")


def _escolher_endereco(page, codigo):
    """Lupa 'Escolher um endereço' -> pesquisa o código -> escolhe a linha dele."""
    lupa = page.locator('img[title="Escolher um endereço"], img[alt="Escolher um endereço"]').first
    lupa.scroll_into_view_if_needed()
    lupa.click()
    page.wait_for_timeout(1500)
    janela = _janela_conta(page)
    busca = janela.locator("input[type='text'], input[type='search'] >> visible=true")
    if busca.count():
        busca.first.fill(codigo)
        busca.first.press("Enter")
        page.wait_for_timeout(2000)
    padrao = re.compile(rf"(^|\D){re.escape(codigo)}(\D|$)")
    alvo = janela.locator("a >> visible=true").filter(has_text=padrao_codigo(codigo))
    if not alvo.count():
        alvo = janela.locator("a >> visible=true").filter(has_text=padrao)
    if not alvo.count():
        alvo = janela.locator("tr, li, td >> visible=true").filter(has_text=padrao)
    if not alvo.count():
        raise RuntimeError(f"endereço {codigo} não apareceu na janela de endereços")
    alvo.first.click()
    page.wait_for_timeout(1200)
    botao = janela.locator("button, input[type='submit'], input[type='button'], a.button").filter(
        has_text=re.compile(r"^\s*(Escolher|Selecionar|OK|Aplicar|Confirmar)\s*$", re.I))
    if botao.count() and botao.first.is_visible():
        botao.first.click()
        page.wait_for_timeout(1200)


def preencher_cabecalho(page, nome, perfil):
    """Preenche o cabeçalho com o perfil. O que falhar fica listado para você completar à mão."""
    print(f"\nPreenchendo o cabeçalho (perfil {nome})...")
    pendentes = []

    def etapa(descricao, funcao):
        ponto_de_pausa(page, "cabeçalho")
        try:
            funcao()
            print(f"  ✔ {descricao}")
        except ParadaSolicitada:
            raise
        except Exception as e:
            if conexao_caiu(e):
                raise
            print(f"  ✖ {descricao}: {str(e).splitlines()[0]}")
            pendentes.append(descricao)
            page.keyboard.press("Escape")

    # justificativa
    def justificativa():
        campo = page.locator("#requisition_header_justification")
        campo.scroll_into_view_if_needed()
        campo.fill(perfil["justificativa"])
    etapa(f"Justificativa: {perfil['justificativa']}", justificativa)

    # tipo de requisição (caixa custom_field_1)
    def tipo_requisicao():
        caixa = page.locator("#requisition_header_custom_field_1_id_chosen")
        caixa.scroll_into_view_if_needed()
        alvo = perfil["tipo_requisicao"]
        busca = alvo.split(" - ")[-1][:5]                     # "Padrã"
        if not tentar_selecionar(page, caixa, busca, re.compile(re.escape(alvo), re.I)):
            raise RuntimeError(f"'{alvo}' não apareceu nas opções")
    etapa(f"Tipo: {perfil['tipo_requisicao']}", tipo_requisicao)

    # categoria em 2 níveis (a caixa é achada pelas opções: a que mostrar o 1º nível)
    def categoria():
        nivel1, nivel2 = perfil["categoria"]
        todas = page.locator('[id^="requisition_header_"][id$="_chosen"]')
        ids = [todas.nth(i).get_attribute("id") for i in range(todas.count())]
        ids = [i for i in ids if i and "custom_field_1_id" not in i]
        busca1 = nivel1.split(",")[0].split(" - ")[-1].strip()      # "MRO"
        busca2 = nivel2.split(",")[0].split(" - ")[-1].strip()      # 2º nível da categoria
        for id_cx in ids:
            caixa = page.locator(f"#{id_cx}")
            if not caixa.is_visible():
                continue
            caixa.scroll_into_view_if_needed()
            if tentar_selecionar(page, caixa, busca1, re.compile(re.escape(nivel1), re.I)):
                page.wait_for_timeout(800)
                if not tentar_selecionar(page, caixa, busca2, re.compile(re.escape(nivel2), re.I)):
                    raise RuntimeError(f"escolhi '{nivel1}', mas '{nivel2}' não apareceu no 2º nível")
                return
        raise RuntimeError(f"nenhuma caixa do cabeçalho mostrou '{nivel1}' nas opções")
    etapa("Categoria: " + " | ".join(perfil["categoria"]), categoria)

    etapa(f"Unidade de negócio: {perfil['unidade_negocio']}",
          lambda: _selecionar_opcao(page, "requisition_header_custom_field_4", perfil["unidade_negocio"]))
    etapa(f"Motivo: {perfil['motivo']}",
          lambda: _selecionar_opcao(page, "requisition_header_custom_field_5", perfil["motivo"]))

    def data():
        campo = page.locator("#requisition_header_custom_field_17")
        campo.scroll_into_view_if_needed()
        campo.click()
        campo.press("Control+a")
        campo.press("Backspace")
        campo.press_sequentially((date.today() + timedelta(days=1)).strftime("%d/%m/%Y"), delay=60)
        campo.press("Tab")
    etapa(f"Necessário até: {(date.today() + timedelta(days=1)).strftime('%d/%m/%Y')}", data)

    etapa(f"Endereço: {perfil['endereco']}", lambda: _escolher_endereco(page, perfil["endereco"]))

    if pendentes:
        try:
            page.screenshot(path=str(pasta_atual() / "falha_cabecalho.png"), full_page=True)
        except Exception:
            pass
        print("\n⚠ Complete à mão no Coupa: " + "; ".join(pendentes) + "  (print em falha_cabecalho.png)")
    input("\nConfira o cabeçalho no Coupa e aperte Enter aqui para lançar as linhas...")


# ---------------------------------------------------------------- Coupa
def visivel(page, seletor):
    return page.locator(f"{seletor} >> visible=true").last


def digitar(campo, texto):
    campo.click()
    campo.press("Control+a")
    campo.press("Backspace")
    campo.press_sequentially(texto, delay=80)


def selecionar_plu(page, plu):
    """Digita o PLU no campo Item e clica na sugestão do autocomplete (jQuery UI).
    Se a lista não abrir, "cutuca" o campo para disparar a busca de novo.
    Digita SÓ o número: se fosse "1234567 | PECA EXEMPLO 100", depois de escolher a sugestão o campo
    ficaria com o mesmo texto e a macro não teria como saber se selecionou."""
    plu = codigo_plu(plu)
    campo = visivel(page, SEL_PLU)
    menu_id = campo.get_attribute("aria-controls")          # ex.: ui-id-3 (lista ligada a este campo)
    sel_opcoes = (f"#{menu_id} li" if menu_id else "ul.ui-autocomplete li.ui-menu-item") + " >> visible=true"
    opcoes = page.locator(sel_opcoes)

    for tentativa in range(3):
        if tentativa == 0:
            digitar(campo, plu)
        else:                                               # redispara a busca do autocomplete
            campo.click()
            campo.press("End")
            campo.press("Backspace")
            campo.press_sequentially(plu[-1], delay=150)
        try:
            opcoes.first.wait_for(state="visible", timeout=6000)
            break
        except Exception:
            continue
    else:
        raise RuntimeError(f"a lista de sugestões do PLU {plu} não abriu no campo Item")

    page.wait_for_timeout(300)                              # deixa a lista terminar de carregar
    textos = [" ".join(t.split()) for t in opcoes.all_inner_texts()]
    # a sugestão CERTA é a que começa com o número do PLU ("1234567 | ...") - não "11234567 | ..."
    idx = next((k for k, t in enumerate(textos) if re.match(rf"\s*{re.escape(plu)}(?!\d)", t)), None)
    if idx is None:
        idx = next((k for k, t in enumerate(textos) if re.search(rf"(?<!\d){re.escape(plu)}(?!\d)", t)), None)
    if idx is None:
        idx = 0
        print(f"  ⚠ Nenhuma sugestão com o número {plu}; escolhendo a primeira: {' '.join(textos[0].split())[:60]}")

    # escolhe pelo TECLADO (seta até a sugestão + Enter), sem clicar: algumas sugestões do Coupa trazem
    # o nome do item como LINK para a página do PLU, e o clique abria essa página em vez de selecionar
    url_antes = page.url
    for _ in range(idx + 1):
        campo.press("ArrowDown")
    campo.press("Enter")
    page.wait_for_timeout(700)
    _voltar_se_saiu(page, url_antes)

    # confirma que selecionou (o campo deixa de ter só o número digitado)
    if campo.input_value().strip() == plu:
        # reserva: clica no canto da sugestão (longe do link do nome)
        if opcoes.count() > idx:
            opcoes.nth(idx).click(position={"x": 4, "y": 4})
            page.wait_for_timeout(700)
            _voltar_se_saiu(page, url_antes)
        if campo.input_value().strip() == plu:
            raise RuntimeError(f"digitei o PLU {plu}, mas a sugestão não foi selecionada")


def _voltar_se_saiu(page, url_antes):
    """Se a seleção abriu outra página (ex.: a página do PLU), volta para a requisição e avisa."""
    if page.url.split("#")[0] == url_antes.split("#")[0]:
        return
    print(f"  ⚠ A seleção do PLU abriu outra página ({page.url[:70]}) - voltando para a requisição")
    try:
        page.go_back(wait_until="domcontentloaded")
    except Exception:
        pass
    if page.url.split("#")[0] != url_antes.split("#")[0]:
        ir_para_coupa(page, url_antes)
    raise RuntimeError("a sugestão do PLU abriu a página do item em vez de selecionar")


def linha_em_edicao(page):
    """Bloco da LINHA que está sendo editada (nunca o cabeçalho da requisição).
    Parte do campo Item da linha (#requisition_line_description) e sobe até o formulário dela."""
    campo = visivel(page, SEL_PLU)
    for xp in ("xpath=ancestor::*[.//span[normalize-space()='Salvar']][1]",      # ancestral mais próximo com Salvar
               "xpath=ancestor::*[contains(@class,'line-wrapper') or contains(@class,'editing')][1]"):
        bloco = campo.locator(xp)
        if bloco.count() and bloco.first.locator(SEL_SALVAR).count():
            return bloco.first
    raise RuntimeError("não consegui identificar o bloco da linha em edição (para não mexer no cabeçalho)")


def preencher_linha(page, item):
    ponto_de_pausa(page, "antes do PLU")
    selecionar_plu(page, item["plu"])
    ponto_de_pausa(page, "depois do PLU")
    linha = linha_em_edicao(page)      # daqui para baixo, tudo só dentro da linha

    # Fornecedor: no select escolhe "Outro", depois digita o CNPJ e clica na sugestão
    select_forn = visivel(linha, SEL_SELECT_FORNECEDOR)
    select_forn.select_option(value="other")
    campo_forn = visivel(linha, SEL_FORNECEDOR)       # o id muda a cada linha, por isso busca pela classe
    campo_forn.wait_for(state="visible", timeout=TIMEOUT)

    escolhido = False
    for texto_cnpj in formas_de_busca_cnpj(item["cnpj"]):   # sem o zero da frente, com o zero, formatado
        digitar(campo_forn, texto_cnpj)
        lista_id = campo_forn.get_attribute("aria-owns")
        opcao_forn = page.locator(f"#{lista_id} [role='option'], #{lista_id} li").first
        try:
            opcao_forn.wait_for(state="visible", timeout=8000)
        except Exception:
            continue
        opcao_forn.click()
        escolhido = True
        FORMA_CNPJ_QUE_FUNCIONOU[normalizar_cnpj(item["cnpj"])] = texto_cnpj   # próximas linhas já vão direto
        break
    if not escolhido:
        raise RuntimeError(f"fornecedor {fmt_cnpj(item['cnpj'])} não apareceu na busca do campo Fornecedor")
    page.wait_for_timeout(800)   # o Coupa atualiza a linha depois de escolher o fornecedor

    ponto_de_pausa(page, "depois do fornecedor")
    # Quantidade e preço unitário (total / quantidade)
    unitario = item["total"] / item["qtd"]
    if round(round(unitario, 2) * item["qtd"], 2) != round(item["total"], 2):
        print(f"  ⚠ {item['origem']}: {br(item['total'])} / {item['qtd']:g} "
              f"não divide exato. Confira o total no Coupa.")
    digitar(visivel(linha, SEL_QUANTIDADE), f"{item['qtd']:g}".replace(".", ","))
    digitar(visivel(linha, SEL_PRECO), br(unitario))

    # Data de necessidade = amanhã
    amanha = (date.today() + timedelta(days=1)).strftime("%d/%m/%Y")
    campo_data = visivel(linha, SEL_DATA)
    digitar(campo_data, amanha)
    campo_data.press("Tab")

    # Campo TR (perfis com campo_tr = coluna H: texto da coluna H)
    if item.get("tr"):
        campo_tr = linha.locator(f"{SEL_TR} >> visible=true")
        if not campo_tr.count():
            raise RuntimeError("não achei o campo TR da linha (#requisition_line_custom_field_10)")
        texto_tr = item["tr"][:255]
        if len(item["tr"]) > 255:
            print(f"  ⚠ texto do TR tem {len(item['tr'])} caracteres; o campo aceita 255 - cortei o resto")
        campo_tr.first.fill(texto_tr)
        if campo_tr.first.input_value() != texto_tr:
            raise RuntimeError("o campo TR não ficou com o texto da coluna H")

    ponto_de_pausa(page, "antes do centro de custo")
    # Empresa -> centro de custo, na caixa "Selecionar" da linha:
    # pesquisa o código da empresa + Enter, depois pesquisa o centro de custo da loja + Enter
    selecionar_empresa_cc(page, item["cc"])


JS_RESULTADOS_NA_TELA = r"""() => {
    // opções de lista (chosen) que estão REALMENTE na tela, em qualquer lugar da página
    const out = [];
    for (const li of document.querySelectorAll('.chosen-results li')) {
        if (li.classList.contains('no-results')) continue;
        if (/no results|nenhum resultado|sem resultado/i.test(li.innerText)) continue;
        const r = li.getBoundingClientRect();
        if (r.width < 2 || r.height < 2) continue;
        const el = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
        if (el && (el === li || li.contains(el))) out.push(li.innerText.trim());
    }
    return out;
}"""


def _resultados_na_tela(page):
    try:
        return page.evaluate(JS_RESULTADOS_NA_TELA)
    except Exception:
        return []


def _caixa_aberta(page, caixa):
    """Aberta = a lista de opções está de fato na tela (ou o chosen marcou chosen-with-drop).
    Fechada, o campo de pesquisa continua no HTML, só escondido - por isso não basta 'existir'."""
    try:
        if "chosen-with-drop" in (caixa.get_attribute("class") or ""):
            return True
    except Exception:
        pass
    return bool(_resultados_na_tela(page))


def _campo_pesquisa(page, caixa):
    """Campo de pesquisa da caixa aberta (ou None se não houver campo - aí digita no foco)."""
    loc = caixa.locator(".chosen-search input, input.chosen-search-input")
    return loc.first if loc.count() else None


def _abrir_caixa(page, caixa):
    """Abre a caixa 'Selecionar'. Tenta clique, seta para baixo e espaço
    (o próprio Coupa avisa: 'Use space or down arrow key to open the combobox'). Retorna True/False."""
    try:
        aberta_mesmo = "chosen-with-drop" in (caixa.get_attribute("class") or "")
    except Exception:
        aberta_mesmo = False
    if not aberta_mesmo and _resultados_na_tela(page):   # a lista aberta é de OUTRA caixa: fecha antes
        page.keyboard.press("Escape")
        page.wait_for_timeout(400)
    if _caixa_aberta(page, caixa):
        return True
    botao = caixa.locator("a.chosen-single").first
    botao.scroll_into_view_if_needed()
    for acao in ("click", "ArrowDown", "Space", "click_ul"):
        if acao == "click":
            botao.click()
        elif acao == "click_ul":
            caixa.locator("ul.chosen-scopes").first.click()
        else:
            botao.focus()
            page.keyboard.press(acao)
        page.wait_for_timeout(800)
        if _caixa_aberta(page, caixa):
            return True
    return False


def _diagnostico_cc(page, caixa, etapa):
    """Salva print e HTML da caixa para descobrir por que não selecionou."""
    pasta = pasta_atual()
    try:
        page.screenshot(path=str(pasta / "falha_cc.png"), full_page=False)
        html = caixa.evaluate("e => e.outerHTML")
        foco = page.evaluate("() => document.activeElement ? document.activeElement.outerHTML.slice(0, 500) : ''")
        drops = page.evaluate("() => [...document.querySelectorAll('.chosen-drop')].map(d => d.outerHTML.slice(0, 3000)).join('\\n\\n')")
        (pasta / "falha_cc.html").write_text(
            f"<!-- etapa: {etapa} -->\n<!-- foco: {foco} -->\n<!-- na tela: {_resultados_na_tela(page)} -->\n"
            f"{html}\n\n<!-- todas as listas chosen da página -->\n{drops}", encoding="utf-8")
    except Exception:
        pass
    return f"(print e HTML salvos em falha_cc.png / falha_cc.html na pasta {pasta.name})"


def pesquisar_e_enter(page, caixa, texto):
    """Abre a caixa (se preciso), pesquisa o texto, espera o resultado e aperta Enter."""
    if not _abrir_caixa(page, caixa):
        raise RuntimeError(f"a caixa de centro de custo não abriu para digitar {texto} "
                           + _diagnostico_cc(page, caixa, f"abrir para {texto}"))
    campo = _campo_pesquisa(page, caixa)
    if campo:
        campo.focus()                   # focus, não click: o campo fica "por baixo" de outros elementos
        campo.press("Control+a")
        campo.press("Backspace")
        campo.press_sequentially(texto, delay=120)
    else:
        page.keyboard.type(texto, delay=120)

    padrao = padrao_codigo(texto) if str(texto).strip().isdigit() else re.compile(re.escape(texto), re.I)
    for _ in range(20):                 # até ~10 s esperando o resultado (a lista pode vir do servidor)
        if any(padrao.search(t) for t in _resultados_na_tela(page)):
            break
        page.wait_for_timeout(500)
    # o Enter pega a opção destacada (a 1ª); se a exata não for a 1ª (ex.: 12416 antes de 2416),
    # desce com a seta até ela
    textos = [t for t in _resultados_na_tela(page) if t]
    exatos = [k for k, t in enumerate(textos) if padrao.search(t)]
    for _ in range(exatos[0] if exatos else 0):
        page.keyboard.press("ArrowDown")
    page.keyboard.press("Enter")
    page.wait_for_timeout(1200)


def selecionar_empresa_cc(page, cc):
    caixa = page.locator(f"{SEL_CC_LINHA} >> visible=true")
    if not caixa.count():
        raise RuntimeError(f"não achei a caixa de centro de custo da linha ({SEL_CC_LINHA})")
    caixa = caixa.last

    texto_caixa = lambda: " ".join(caixa.locator("a.chosen-single").first.inner_text().split())

    pesquisar_e_enter(page, caixa, COD_EMPRESA)            # 1) código da empresa + Enter
    if COD_EMPRESA not in texto_caixa() and not _caixa_aberta(page, caixa):
        raise RuntimeError(f"pesquisei {COD_EMPRESA} e dei Enter, mas não selecionou "
                           + _diagnostico_cc(page, caixa, f"depois do Enter no {COD_EMPRESA}"))

    pesquisar_e_enter(page, caixa, cc)                     # 2) centro de custo da loja + Enter
    escolhido = texto_caixa()
    if not padrao_codigo(cc).search(escolhido):
        raise RuntimeError(f"a caixa de centro de custo ficou com \"{escolhido[:60]}\" (esperado {cc}) "
                           + _diagnostico_cc(page, caixa, "depois do Enter no CC"))


def salvar_linha(page):
    visivel(linha_em_edicao(page), SEL_SALVAR).click()
    # salvamento é AJAX: espera o formulário de edição sumir
    page.wait_for_function(
        "document.querySelectorAll('#requisition_line_description').length === 0 || "
        "[...document.querySelectorAll('#requisition_line_description')]"
        ".every(e => e.offsetParent === null)",
        timeout=TIMEOUT,
    )
    page.wait_for_timeout(1000)


def verificar_planilha(linhas, wb, verificador):
    """Modo planilha: confere no Coupa os PLUs que ainda não são confiáveis."""
    confiaveis = carregar_confiaveis(wb)
    problemas, novos_conf, novos_bloq = [], [], {}
    for l in linhas:
        if l["plu"] in confiaveis:
            continue
        ok, motivo = verificador.verificar(l["plu"])
        print(f"  🔎 PLU {l['plu']}: " + ("✔ liberado" if ok else f"✖ {motivo}"))
        if ok:
            novos_conf.append(l["plu"])
        else:
            problemas.append(f"{l['origem']} (PLU {l['plu']}): {motivo}")
            if ok is False:
                novos_bloq[l["plu"]] = motivo
    salvar_aprendizado([], novos_conf, carregar_plus(wb), novos_bloq)
    if problemas:
        print("\n✖ PLUs com problema:")
        for pr in problemas:
            print("   ", pr)
        return input("Lançar mesmo assim? [s/n]: ").strip().lower() == "s"
    return True


# ---------------------------------------------------------------- conta da linha (lupa)
JS_JANELA = r"""() => {
    // janela/popup visível que apareceu depois de clicar na lupa
    const cands = [...document.querySelectorAll('.ui-dialog, [role="dialog"], .modal, .popup, .lightbox, .coupa-dialog')];
    const vis = cands.filter(e => { const r = e.getBoundingClientRect(); return r.width > 50 && r.height > 50; });
    return vis.length ? vis.length - 1 : -1;
}"""


def _janela_conta(page):
    idx = page.evaluate(JS_JANELA)
    if idx < 0:
        return page
    return page.locator('.ui-dialog, [role="dialog"], .modal, .popup, .lightbox, .coupa-dialog').nth(idx)


def _diagnostico_conta(page, janela):
    pasta = pasta_atual()
    try:
        page.screenshot(path=str(pasta / "falha_conta.png"), full_page=False)
        html = janela.evaluate("e => e.outerHTML") if janela is not page else page.content()
        (pasta / "falha_conta.html").write_text(html[:300000], encoding="utf-8")
    except Exception:
        pass
    return "print e HTML salvos em falha_conta.png / falha_conta.html"


def tentar_selecionar(page, caixa, texto, padrao, espera_ms=6000):
    """Abre a caixa, pesquisa o texto e, SE aparecer uma opção que casa com o padrão, dá Enter nela.
    Se nenhuma opção casar, fecha a caixa (Esc) e devolve False - é assim que a macro descobre
    qual das caixas iguais é a certa: pela lista de opções que cada uma mostra."""
    if not _abrir_caixa(page, caixa):
        return False
    campo = _campo_pesquisa(page, caixa)
    if campo:
        campo.focus()
        campo.press("Control+a")
        campo.press("Backspace")
        campo.press_sequentially(texto, delay=120)
    else:
        page.keyboard.type(texto, delay=120)

    achou = False
    for _ in range(max(1, espera_ms // 500)):
        if any(padrao.search(t) for t in _resultados_na_tela(page)):
            achou = True
            break
        page.wait_for_timeout(500)
    if not achou:
        page.keyboard.press("Escape")
        page.wait_for_timeout(400)
        return False

    textos = [t for t in _resultados_na_tela(page) if t]
    exatos = [k for k, t in enumerate(textos) if padrao.search(t)]
    for _ in range(exatos[0] if exatos else 0):
        page.keyboard.press("ArrowDown")
    page.keyboard.press("Enter")
    page.wait_for_timeout(1200)
    rotulo = caixa.locator("a.chosen-single").first if caixa.locator("a.chosen-single").count() else caixa
    return bool(padrao.search(" ".join(rotulo.inner_text().split())))


SEL_PAGINAS_LINHAS = ('a.requisition_line_wait, a[href*="search_requisition_line_table"], '
                      '.pagination a, .pageLinks a, .paginate a')

JS_LUPA_DO_PLU = r"""([seletor, plu]) => {
    const _visivel = e => e.getClientRects().length > 0 && getComputedStyle(e).visibility !== 'hidden';
    const _blocoDaLupa = (l, seletor) => {
        // a LINHA inteira: o contêiner de linha do Coupa (ou a linha da tabela); cada linha pode ter
        // mais de uma lupa (ex.: uma no resumo e outra no formulário de edição escondido)
        for (const s of ['.line-wrapper', '.lineSummary', 'tr']) {
            const c = l.closest(s);
            if (c) return c;
        }
        let b = l;       // sem contêiner conhecido: maior bloco com UMA lupa visível
        while (b.parentElement &&
               [...b.parentElement.querySelectorAll(seletor)].filter(_visivel).length <= 1) b = b.parentElement;
        return b;
    };
    const _linhas = (seletor) => {
        // [{bloco, i}] - um item por LINHA (sem repetir), só linhas com alguma lupa visível;
        // i = índice de uma lupa visível dessa linha na lista completa de lupas
        const todas = [...document.querySelectorAll(seletor)];
        const vistos = new Map();
        todas.forEach((l, i) => {
            if (!_visivel(l)) return;
            const b = _blocoDaLupa(l, seletor);
            if (!vistos.has(b)) vistos.set(b, i);
        });
        return [...vistos.entries()].map(([bloco, i]) => ({bloco, i}));
    };
    const re = new RegExp('(^|\\D)' + plu + '(\\D|$)');
    const linhas = _linhas(seletor);
    for (let k = linhas.length - 1; k >= 0; k--)          // de baixo para cima: a mais recente primeiro
        if (re.test(linhas[k].bloco.innerText || '')) return linhas[k].i;
    return -1;
}"""


def _paginas_disponiveis(page):
    """{nº da página: link} dos links de paginação da lista de linhas."""
    paginas = {}
    for a in page.locator(f"{SEL_PAGINAS_LINHAS} >> visible=true").all():
        try:
            info = (a.get_attribute("href") or "") + " " + (a.get_attribute("data-params") or "")
            m = re.search(r"page=(\d+)", info)
            n = int(m.group(1)) if m else (int(a.inner_text().strip()) if a.inner_text().strip().isdigit() else None)
        except Exception:
            n = None
        if n and n not in paginas:
            paginas[n] = a
    return paginas


def _texto_das_linhas(page):
    try:
        return page.evaluate(f"() => [...document.querySelectorAll('{SEL_LUPA_CONTA}')]"
                             ".map(l => (l.closest('tr, li, .line-wrapper, div') || l).innerText).join('|')")
    except Exception:
        return ""


def _pagina_atual_e_ultima(page):
    """(página atual, última página) da lista de linhas. A atual é o número que aparece no
    paginador sem ser link."""
    links = _paginas_disponiveis(page)
    if not links:
        return 1, 1
    try:
        texto = next(iter(links.values())).evaluate(
            "a => (a.closest('.pagination, .pageLinks, .paginate, nav, div, p') || a.parentElement).innerText")
    except Exception:
        texto = ""
    numeros = {int(n) for n in re.findall(r"\b\d{1,3}\b", texto)}
    atual_cands = sorted(numeros - set(links))
    atual = atual_cands[0] if atual_cands else 1
    return atual, max(set(links) | {atual})


def localizar_lupa_da_linha(page, plu, ultima_ocorrencia=False):
    """Acha a lupa 'Escolher uma conta' da linha do PLU, passando pelas páginas da lista de linhas
    (a partir de umas 5 linhas o Coupa cria uma 2ª página). Começa pela ÚLTIMA página, que é onde a
    linha recém-salva costuma ficar. Com PLU repetido (perfil que permite), vai direto à última
    página e pega a última ocorrência - senão acabaria mexendo numa linha anterior com o mesmo PLU.
    Devolve o locator da lupa ou None se não achar."""
    def procurar():
        i = page.evaluate(JS_LUPA_DO_PLU, [SEL_LUPA_CONTA, plu])
        return page.locator(SEL_LUPA_CONTA).nth(i) if i >= 0 else None

    def ir_para(n, paginas):
        antes = _texto_das_linhas(page)
        paginas[n].scroll_into_view_if_needed()
        paginas[n].click()
        for _ in range(16):                      # espera a lista trocar (AJAX), até ~8 s
            page.wait_for_timeout(500)
            if _texto_das_linhas(page) != antes:
                break

    visitadas = set()
    if ultima_ocorrencia:
        atual, ultima = _pagina_atual_e_ultima(page)
        visitadas.add(atual)
        if ultima != atual:
            paginas = _paginas_disponiveis(page)
            if ultima in paginas:
                ir_para(ultima, paginas)
                visitadas.add(ultima)

    achada = procurar()
    if achada:
        return achada
    for _ in range(20):
        paginas = _paginas_disponiveis(page)
        pendentes = sorted((n for n in paginas if n not in visitadas), reverse=True)   # da última p/ a 1ª
        if not pendentes:
            return None
        n = pendentes[0]
        visitadas.add(n)
        ir_para(n, paginas)
        achada = procurar()
        if achada:
            print(f"  (linha do PLU {plu} encontrada na página {n} das linhas)")
            return achada
    return None


def escolher_da_lista(page, caixa, padrao):
    """Abre a caixa e escolhe a opção que casa com o padrão SEM digitar nada: desce com a seta até
    ela e dá Enter (se não pegar, clica na opção). Para caixas que são só de seleção."""
    if not _abrir_caixa(page, caixa):
        return False
    page.wait_for_timeout(400)
    textos = [t for t in _resultados_na_tela(page) if t]
    alvo = next((k for k, t in enumerate(textos) if padrao.search(t)), None)
    if alvo is None:
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)
        return False
    rotulo = caixa.locator("a.chosen-single").first if caixa.locator("a.chosen-single").count() else caixa
    # 1ª tentativa: teclado (a 1ª opção já vem destacada)
    for _ in range(alvo):
        page.keyboard.press("ArrowDown")
    page.keyboard.press("Enter")
    page.wait_for_timeout(1000)
    if padrao.search(" ".join(rotulo.inner_text().split())):
        return True
    # 2ª tentativa: clique na opção
    if not _abrir_caixa(page, caixa):
        return False
    opcao = page.locator(".chosen-results li").filter(has_text=padrao)
    for k in range(opcao.count()):
        if opcao.nth(k).is_visible():
            opcao.nth(k).click()
            break
    page.wait_for_timeout(1000)
    return bool(padrao.search(" ".join(rotulo.inner_text().split())))


def _select_escondido(janela, padrao):
    """Escolhe direto no <select> (mesmo escondido atrás da caixa) a opção que casa com o padrão,
    avisando a página da mudança. Devolve True se encontrou e marcou."""
    for sel in janela.locator("select").all():
        try:
            ok = sel.evaluate("""(s, fonte) => {
                const re = new RegExp(fonte, 'i');
                const o = [...s.options].find(o => re.test(o.text.trim()));
                if (!o) return false;
                s.value = o.value;
                o.selected = true;
                s.dispatchEvent(new Event('change', {bubbles: true}));
                if (window.jQuery) window.jQuery(s).trigger('change').trigger('chosen:updated');
                return true;
            }""", padrao.pattern)
        except Exception:
            ok = False
        if ok:
            return True
    return False


def ajustar_conta_da_linha(page, item):
    """Depois de salvar a linha: lupa 'Escolher uma conta' -> troca o centro de custo (vem 7065 - PCM)
    pelo da loja -> escolhe K - Centro de Custo -> confirma.
    As caixas da janela são idênticas (mesmo HTML), então a macro identifica cada uma pelas OPÇÕES:
    pesquisa o valor e só confirma na caixa em que ele aparece na lista."""
    cc = item["cc"]
    # a lupa certa é a da linha deste PLU: com muitas linhas a lista pagina e a linha pode estar
    # em outra página - a macro procura o PLU página por página e NUNCA mexe na conta de outra linha
    page.locator(SEL_LUPA_CONTA).first.wait_for(state="attached", timeout=TIMEOUT)
    lupa = localizar_lupa_da_linha(page, item["plu"], ultima_ocorrencia=item.get("plu_repetido", False))
    if lupa is None:
        raise RuntimeError(f"não achei a linha do PLU {item['plu']} em nenhuma página da lista de linhas")
    lupa.scroll_into_view_if_needed()                # rola até ela aparecer
    lupa.wait_for(state="visible", timeout=TIMEOUT)
    lupa.click()
    page.wait_for_timeout(1500)
    janela = _janela_conta(page)

    todas = janela.locator(".chosen-container")
    caixas = [todas.nth(i) for i in range(todas.count())]
    if not caixas:
        raise RuntimeError("não achei as caixas de seleção na janela da conta")

    def texto_da(cx):
        try:
            return " ".join(cx.inner_text().split())
        except Exception:
            return ""

    tipo = item.get("conta_tipo", TIPO_CONTA)              # K (centro de custo) ou A (ativo imobilizado)
    valor = str(item.get("conta_valor") or cc)              # CC da loja (K) ou nº da coluna G (A)
    padrao_tipo = re.compile(rf"^\s*{re.escape(tipo)}\s*-", re.I)
    padrao_valor = padrao_codigo(valor)          # só a opção com ESSE código (não "LJ 1220" de outra)
    eh_valor_agora = re.compile(r"\d{3,}\s*-\s*.+\(\d{3,}\)")        # "7065 - PCM (7065)"
    eh_tipo_agora = re.compile(r"[A-Z]\s*-\s*.+\([A-Z]\)", re.I)       # "K - Centro de Custo (K)"

    def escolher_tipo(excluir):
        """O tipo (K / A) é caixa SÓ DE SELEÇÃO: escolhe da lista, sem digitar."""
        restantes = [i for i in range(len(caixas)) if i not in excluir]
        restantes.sort(key=lambda i: 0 if eh_tipo_agora.search(texto_da(caixas[i])) else 1)
        for i in restantes:                                  # já está no tipo certo?
            if padrao_tipo.search(texto_da(caixas[i])):
                return i
        for i in restantes:                                  # abre e escolhe da lista (sem digitar)
            if escolher_da_lista(page, caixas[i], padrao_tipo):
                return i
        for sel in janela.locator("select >> visible=true").all():          # <select> comum visível
            alvo = next((r for r in sel.locator("option").all_inner_texts() if padrao_tipo.search(r)), None)
            if alvo:
                sel.select_option(label=alvo)
                return -1
        if _select_escondido(janela, padrao_tipo):          # <select> por trás da caixa
            page.wait_for_timeout(1000)
            for i in restantes:
                if padrao_tipo.search(texto_da(caixas[i])):
                    return i
            return -1
        raise RuntimeError(f"não achei onde escolher '{tipo} - ...' na janela da conta")

    def escolher_valor(excluir):
        ordem = [i for i in range(len(caixas)) if i not in excluir]
        ordem.sort(key=lambda i: 0 if eh_valor_agora.search(texto_da(caixas[i])) else 1)
        for i in ordem:
            if tentar_selecionar(page, caixas[i], valor, padrao_valor):
                return i
        raise RuntimeError(f"nenhuma caixa da janela mostrou {valor} nas opções")

    if tipo.upper() == "K":
        # K já costuma vir marcado: primeiro o centro de custo, depois confere o K
        i_valor = escolher_valor(set())
        escolher_tipo({i_valor})
    else:
        # outro tipo (ex.: A - Ativo imobilizado): troca o tipo ANTES, porque ele muda o campo do valor
        i_tipo = escolher_tipo(set())
        page.wait_for_timeout(1000)
        todas = janela.locator(".chosen-container")
        caixas[:] = [todas.nth(i) for i in range(todas.count())]
        escolher_valor({i_tipo} if i_tipo >= 0 else set())

    # 3) confirma a janela
    botao = janela.locator("button, input[type='submit'], input[type='button'], a.button").filter(
        has_text=re.compile(r"^\s*(Escolher|Selecionar|Salvar|OK|Aplicar|Confirmar)\s*$", re.I))
    if not botao.count():
        botao = janela.locator("input[value*='Escolher'], input[value*='Salvar'], input[value='OK'], "
                               "input[value*='Aplicar'] >> visible=true")
    if botao.count():
        botao.first.click()
        page.wait_for_timeout(1500)


def ajustar_conta_seguro(page, item):
    """Tenta ajustar a conta; se não conseguir, salva diagnóstico e deixa você terminar à mão
    (a linha já está salva - não pode ser refeita)."""
    try:
        ponto_de_pausa(page, "antes da conta")
        ajustar_conta_da_linha(page, item)
        print(f"  ✔ conta ajustada: {item.get('conta_valor') or item['cc']} · {item.get('conta_tipo', TIPO_CONTA)}")
        return True
    except ParadaSolicitada:
        raise
    except Exception as e:
        if conexao_caiu(e):
            raise
        diag = _diagnostico_conta(page, _janela_conta(page))
        print(f"  ⚠ Não consegui ajustar a conta da linha: {str(e).splitlines()[0]} ({diag})")
        input(f"  Ajuste à mão: tipo '{item.get('conta_tipo', TIPO_CONTA)}' e "
              f"'{item.get('conta_valor') or item['cc']}'. "
              "Depois aperte Enter aqui para seguir...")
        return True


AVISOS_DO_SITE = []


ACEITAR_PROXIMO_AVISO = {"ativo": False}


def tratar_aviso(dialog):
    """O Coupa às vezes abre uma janela de aviso do navegador (alert/confirm/"sair desta página?").
    Sem tratador, o Playwright tenta fechá-la sozinho e, se ela já sumiu, derruba a conexão.
    Normalmente a macro registra o texto e CANCELA (fica na página, não perde a requisição).
    Exceção: logo depois de a própria macro clicar em Excluir, o "Tem certeza?" é CONFIRMADO (OK)."""
    texto = " ".join((dialog.message or "").split())
    AVISOS_DO_SITE.append(texto)
    aceitar = ACEITAR_PROXIMO_AVISO["ativo"] and dialog.type in ("confirm", "alert")
    ACEITAR_PROXIMO_AVISO["ativo"] = False
    print(f"\n  ⚠ Aviso do site ({dialog.type}): {texto or '(sem texto)'} → "
          + ("OK (confirmado)" if aceitar else "cancelado"))
    try:
        dialog.accept() if aceitar else dialog.dismiss()
    except Exception:
        pass            # a janela já tinha fechado sozinha


def vigiar_avisos(ctx):
    for pg in ctx.pages:
        pg.on("dialog", tratar_aviso)
    ctx.on("page", lambda pg: pg.on("dialog", tratar_aviso))


def conexao_caiu(erro):
    t = str(erro)
    return any(x in t for x in ("Connection closed", "Target closed", "Browser has been closed",
                                "Target page, context or browser has been closed"))


def abrir_requisicao(page):
    """Carrinho -> 'Revisar carrinho' (abre a requisição). Devolve o endereço da requisição."""
    if not page.locator(SEL_CARRINHO).first.is_visible():
        ir_para_coupa(page)
        garantir_login(page)
    page.locator(SEL_CARRINHO).first.locator(
        "xpath=ancestor::*[self::a or self::button][1]"
    ).click()
    revisar = page.locator(f"{SEL_REVISAR_CARRINHO} >> visible=true").first
    revisar.wait_for(state="visible", timeout=TIMEOUT)
    href = revisar.get_attribute("href")
    if href:
        ir_para_coupa(page, urljoin(page.url, href))          # /requisition_headers/<nº>/edit
    else:
        revisar.click()
    page.wait_for_load_state("domcontentloaded")
    return page.url


def _na_requisicao(page, url):
    return bool(url) and url.split("?")[0].split("#")[0] in page.url


def garantir_na_requisicao(page, url):
    """Antes de cada linha: se a sessão caiu ou a página saiu da requisição, volta para ela."""
    if not url or _na_requisicao(page, url):
        return
    print("  ↻ A página saiu da requisição (sessão expirou ou navegação). Voltando para ela...")
    ir_para_coupa(page, url)
    if not _na_requisicao(page, url):
        garantir_login(page)
        ir_para_coupa(page, url)
    page.wait_for_load_state("domcontentloaded")


def _cancelar_edicao_aberta(page):
    """Fecha SEM SALVAR a linha que estiver aberta para edição (botão Cancelar da própria linha)."""
    linha = linha_em_edicao(page)
    botao = linha.locator('a:text-is("Cancelar"), button:text-is("Cancelar"), '
                          'span:text-is("Cancelar") >> visible=true')
    if not botao.count():
        raise RuntimeError("uma linha já salva ficou aberta para edição e não achei o Cancelar dela - "
                           "clique em Cancelar nessa linha no Coupa")
    botao.first.click()
    for _ in range(16):
        page.wait_for_timeout(500)
        if not page.locator(f"{SEL_PLU} >> visible=true").count():
            return
    raise RuntimeError("cliquei em Cancelar, mas a linha continuou aberta para edição")


def _abrir_ou_reusar_linha(page, marca):
    """Abre a linha onde o item 'marca' vai ser lançado.
    - Se a linha aberta é a que a PRÓPRIA macro abriu para este item (tentativa anterior), reaproveita.
    - Se é uma linha nova em branco, reaproveita.
    - Se é uma linha JÁ SALVA que ficou aberta para edição (clicar no resumo de uma linha abre ela),
      fecha sem salvar - senão o próximo item seria escrito POR CIMA dela.
    Depois, se preciso, clica em 'Adicionar linha' e marca o campo Item com 'marca'."""
    aberto = page.locator(f"{SEL_PLU} >> visible=true")
    if aberto.count():
        campo = aberto.last
        dono = campo.get_attribute("data-macro")
        if dono == marca:
            return
        if dono is None and not campo.input_value().strip():
            campo.evaluate("(e, m) => e.setAttribute('data-macro', m)", marca)
            return
        print("  ⚠ Havia uma linha já salva aberta para edição - fechando sem alterar antes de seguir.")
        _cancelar_edicao_aberta(page)
    visivel(page, SEL_ADICIONAR_LINHA).click()
    campo = visivel(page, SEL_PLU)
    campo.wait_for(timeout=TIMEOUT)
    campo.evaluate("(e, m) => e.setAttribute('data-macro', m)", marca)


def brl(valor):
    """1234.5 -> '1.234,50'"""
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


JS_BLOCOS_LINHAS = r"""(seletor) => {
    const _visivel = e => e.getClientRects().length > 0 && getComputedStyle(e).visibility !== 'hidden';
    const _blocoDaLupa = (l, seletor) => {
        // a LINHA inteira: o contêiner de linha do Coupa (ou a linha da tabela); cada linha pode ter
        // mais de uma lupa (ex.: uma no resumo e outra no formulário de edição escondido)
        for (const s of ['.line-wrapper', '.lineSummary', 'tr']) {
            const c = l.closest(s);
            if (c) return c;
        }
        let b = l;       // sem contêiner conhecido: maior bloco com UMA lupa visível
        while (b.parentElement &&
               [...b.parentElement.querySelectorAll(seletor)].filter(_visivel).length <= 1) b = b.parentElement;
        return b;
    };
    const _linhas = (seletor) => {
        // [{bloco, i}] - um item por LINHA (sem repetir), só linhas com alguma lupa visível;
        // i = índice de uma lupa visível dessa linha na lista completa de lupas
        const todas = [...document.querySelectorAll(seletor)];
        const vistos = new Map();
        todas.forEach((l, i) => {
            if (!_visivel(l)) return;
            const b = _blocoDaLupa(l, seletor);
            if (!vistos.has(b)) vistos.set(b, i);
        });
        return [...vistos.entries()].map(([bloco, i]) => ({bloco, i}));
    };
    return _linhas(seletor).map(x => (x.bloco.innerText || '').replace(/\s+/g, ' ').trim());
}"""


JS_CONTAR_LINHAS = r"""(seletor) => {
    const _visivel = e => e.getClientRects().length > 0 && getComputedStyle(e).visibility !== 'hidden';
    const _blocoDaLupa = (l, seletor) => {
        // a LINHA inteira: o contêiner de linha do Coupa (ou a linha da tabela); cada linha pode ter
        // mais de uma lupa (ex.: uma no resumo e outra no formulário de edição escondido)
        for (const s of ['.line-wrapper', '.lineSummary', 'tr']) {
            const c = l.closest(s);
            if (c) return c;
        }
        let b = l;       // sem contêiner conhecido: maior bloco com UMA lupa visível
        while (b.parentElement &&
               [...b.parentElement.querySelectorAll(seletor)].filter(_visivel).length <= 1) b = b.parentElement;
        return b;
    };
    const _linhas = (seletor) => {
        // [{bloco, i}] - um item por LINHA (sem repetir), só linhas com alguma lupa visível;
        // i = índice de uma lupa visível dessa linha na lista completa de lupas
        const todas = [...document.querySelectorAll(seletor)];
        const vistos = new Map();
        todas.forEach((l, i) => {
            if (!_visivel(l)) return;
            const b = _blocoDaLupa(l, seletor);
            if (!vistos.has(b)) vistos.set(b, i);
        });
        return [...vistos.entries()].map(([bloco, i]) => ({bloco, i}));
    };
    return _linhas(seletor).length;
}"""


def _ler_linhas_da_pagina(page):
    try:
        return page.evaluate(JS_BLOCOS_LINHAS, SEL_LUPA_CONTA)
    except Exception:
        return []


def ler_linhas_da_requisicao(page):
    """Texto de TODAS as linhas da requisição, passando por todas as páginas da lista (da 1ª à última)."""
    def trocar_para(n, paginas):
        antes = _texto_das_linhas(page)
        paginas[n].scroll_into_view_if_needed()
        paginas[n].click()
        for _ in range(16):
            page.wait_for_timeout(500)
            if _texto_das_linhas(page) != antes:
                break

    paginas = _paginas_disponiveis(page)
    if paginas:
        atual, _ = _pagina_atual_e_ultima(page)
        if atual != 1 and 1 in paginas:
            trocar_para(1, paginas)
    textos, visitadas = list(_ler_linhas_da_pagina(page)), {1}
    for _ in range(30):
        paginas = _paginas_disponiveis(page)
        proximas = sorted(n for n in paginas if n not in visitadas and n > max(visitadas))
        if not proximas:
            break
        n = proximas[0]
        visitadas.add(n)
        trocar_para(n, paginas)
        textos += _ler_linhas_da_pagina(page)
    return textos


def _valores(texto):
    return [num_br(v) for v in re.findall(r"(?<![\d.,])\d{1,3}(?:\.\d{3})*,\d{2}(?!\d)", texto)]


def _total_da_linha(texto, qtd):
    """Valor TOTAL da linha no texto do Coupa (e não o preço unitário ou a quantidade 'x,00'):
    procura um par unitário × quantidade = total; sem par, usa o último valor da linha."""
    vals = _valores(texto)
    if not vals:
        return None
    q = float(qtd or 1)
    for j, t in enumerate(vals):
        for k, u in enumerate(vals):
            if j != k and abs(u * q - t) < 0.02 and not (abs(u - q) < 0.001 and abs(t - q * q) < 0.02):
                return t
    return vals[-1]


def _tem_codigo(texto, codigo):
    return bool(codigo) and re.search(rf"(?<![\d.,]){re.escape(str(codigo))}(?![\d.,])", texto) is not None


RE_ITEM_NA_LINHA = re.compile(r"(?<![\d.,])(\d{5,10})\s*[|\-–]\s*[A-Za-zÀ-ú0-9]")


def filtrar_linhas_por_plu(blocos, linhas):
    """Fica só com os blocos que são LINHAS de item (têm 'PLU | descrição') e conta cada PLU uma vez:
    no Coupa a mesma linha às vezes aparece em mais de um bloco (resumo + detalhe), e há blocos que nem
    são linhas (ex.: a lupa da conta no cabeçalho). Se o PLU aparece N vezes na nota (perfil que permite
    repetir), aceita até N blocos diferentes dele."""
    esperado = Counter(str(l["plu"]) for l in linhas)
    com_item = []
    for t in blocos:
        m = RE_ITEM_NA_LINHA.search(t)
        if m:
            com_item.append((m.group(1), t))
    if not com_item:                       # o Coupa mostra o item de outro jeito: usa os PLUs da nota
        com_item = [(next((p for p in esperado if _tem_codigo(t, p)), None), t) for t in blocos]
        com_item = [(p, t) for p, t in com_item if p]
    if not com_item:
        return blocos                      # sem como identificar: mantém como estava
    por_plu = {}
    for plu, t in com_item:
        por_plu.setdefault(plu, [])
        if t not in por_plu[plu]:          # textos idênticos = o mesmo bloco repetido
            por_plu[plu].append(t)
    filtrados = []
    for plu, textos in por_plu.items():
        n = max(1, esperado.get(plu, 1))
        com_valor = [t for t in textos if _total_da_linha(t, 1) is not None]
        detalhes = " ".join(t for t in textos if t not in com_valor)
        if not com_valor:                   # nenhum bloco com valor: junta tudo numa linha só
            filtrados.append(" ".join(textos))
            continue
        # o resumo (com valor) + os blocos de detalhe do mesmo PLU (conta, centro de custo...)
        filtrados += [(t + " " + detalhes).strip() for t in com_valor[:n]]
        filtrados += com_valor[n:]          # o mesmo PLU lançado mais vezes do que na nota: aparece como sobra
    return filtrados


def conferir_requisicao(page, linhas, total_nota=None):
    """Lê as linhas lançadas no Coupa e confere com o que veio da nota: valor e centro de custo de cada
    linha, quantidade de linhas e soma. Mostra uma tabela e devolve a lista de divergências."""
    print("\n🔎 Conferindo a requisição com a nota...")
    blocos = ler_linhas_da_requisicao(page)
    if not blocos:
        print("  ⚠ Não consegui ler as linhas da requisição para conferir - confira à mão.")
        return ["conferência automática não conseguiu ler as linhas"]
    blocos = filtrar_linhas_por_plu(blocos, linhas)        # 1 PLU = 1 linha

    usados, divergencias, soma_coupa = set(), [], 0.0
    print(f"  {'#':>3}  {'PLU':<9} {'valor da nota':>14}  {'no Coupa':<16} {'centro de custo':<18} conta")
    for i, l in enumerate(linhas, start=1):
        plu, esperado, cc = str(l["plu"]), round(float(l["total"]), 2), str(l.get("cc", ""))
        conta = str(l.get("conta_valor") or "")
        candidatos = [k for k, t in enumerate(blocos) if k not in usados and _tem_codigo(t, plu)]
        # prefere a linha do mesmo PLU que tem o mesmo valor (importa quando o PLU se repete)
        k = next((k for k in candidatos
                  if abs((_total_da_linha(blocos[k], l.get("qtd")) or -1) - esperado) < 0.01),
                 candidatos[0] if candidatos else None)
        if k is None:
            print(f"  {i:>3}  {plu:<9} {brl(esperado):>14}  ✖ linha não encontrada")
            divergencias.append(f"linha {i} (PLU {plu}): não encontrada no Coupa")
            continue
        usados.add(k)
        texto = blocos[k]
        total_coupa = _total_da_linha(texto, l.get("qtd"))
        valor_ok = total_coupa is not None and abs(total_coupa - esperado) < 0.01
        soma_coupa += total_coupa or 0
        cc_ok = _tem_codigo(texto, cc)
        conta_ok = _tem_codigo(texto, conta) if conta else None
        col_valor = (f"✔ {brl(total_coupa)}" if valor_ok
                     else f"✖ {brl(total_coupa)}" if total_coupa is not None else "✖ sem valor")
        col_cc = f"{'✔' if cc_ok else '✖'} {cc}"
        col_conta = "" if conta_ok is None else f"{'✔' if conta_ok else '✖'} {conta}"
        print(f"  {i:>3}  {plu:<9} {brl(esperado):>14}  {col_valor:<16} {col_cc:<18} {col_conta}")
        if not valor_ok:
            achado = f"R$ {brl(total_coupa)}" if total_coupa is not None else "nenhum valor"
            divergencias.append(f"linha {i} (PLU {plu}): valor da nota R$ {brl(esperado)}, no Coupa {achado}")
        if not cc_ok:
            divergencias.append(f"linha {i} (PLU {plu}): centro de custo {cc} não aparece na linha")
        if conta_ok is False:
            divergencias.append(f"linha {i} (PLU {plu}): nº da conta {conta} não aparece na linha")

    sobra = [k for k in range(len(blocos)) if k not in usados]
    if sobra:
        plus_sobra = []
        for k in sobra:
            m = RE_ITEM_NA_LINHA.search(blocos[k])
            plus_sobra.append(m.group(1) if m else "?")
        divergencias.append(f"{len(sobra)} linha(s) no Coupa que não são desta nota "
                            f"(PLU {', '.join(plus_sobra)}) - confira se não foram lançadas a mais")
    soma_nota = round(sum(float(l["total"]) for l in linhas), 2)
    print(f"\n  Linhas: {len(linhas)} da nota · {len(blocos)} no Coupa · "
          f"soma no Coupa R$ {brl(soma_coupa)} · soma da nota R$ {brl(soma_nota)}"
          + (f" (total dos produtos da nota: R$ {brl(total_nota)})" if total_nota else ""))
    if total_nota and abs(total_nota - soma_nota) >= 0.01:
        divergencias.append(f"soma das linhas R$ {brl(soma_nota)} ≠ total dos produtos da nota R$ {brl(total_nota)}")
    if divergencias:
        print(f"\n  ⚠ {len(divergencias)} divergência(s):")
        for d in divergencias:
            print("   -", d)
    else:
        print("  ✔ Tudo confere: valores e centros de custo de todas as linhas batem com a nota.")
    return divergencias


def apagar_linhas_existentes(page):
    """Carrinho de requisição NOVA que já tem linhas (sobra de outra tentativa): marca 'Selecionar todos',
    clica em Excluir e confirma o 'Tem certeza?'. Repete se houver mais de uma página de linhas.
    Devolve quantas linhas apagou."""
    apagadas = 0
    for _ in range(10):
        n = page.evaluate(JS_CONTAR_LINHAS, SEL_LUPA_CONTA)   # linhas (uma linha pode ter 2 lupas)
        if n == 0:
            break
        if apagadas == 0:
            print("  🗑 O carrinho já tem linha(s) de antes - apagando para começar a requisição do zero...")
            if CONFIRMAR_ANTES_DE_APAGAR and input("  Apagar as linhas existentes? [Enter] sim | [n] não: "
                                                   ).strip().lower() == "n":
                return 0
        todos = page.locator(SEL_SELECIONAR_TODOS).first
        todos.wait_for(state="attached", timeout=TIMEOUT)
        todos.scroll_into_view_if_needed()
        if not todos.is_checked():
            todos.check()
        page.wait_for_timeout(500)
        excluir = page.locator('span:text-is("Excluir") >> visible=true')
        # o "Excluir" de dentro de uma linha em edição não serve - quer o da lista (carrinho)
        botao = None
        for k in range(excluir.count()):
            e = excluir.nth(k)
            if not e.evaluate("el => !!el.closest('.line_bottom_inside, .editing')"):
                botao = e
                break
        if botao is None:
            raise RuntimeError("marquei 'Selecionar todos', mas não achei o botão Excluir")
        primeira = page.locator(SEL_LUPA_CONTA).first.element_handle()   # para saber quando a lista trocou
        ACEITAR_PROXIMO_AVISO["ativo"] = True               # o próximo "Tem certeza?" é para confirmar
        botao.click()
        sumiu = False
        for _ in range(30):                                  # espera as linhas sumirem (até ~15 s)
            page.wait_for_timeout(500)
            try:
                sumiu = not primeira.evaluate("e => e.isConnected")
            except Exception:
                sumiu = True                                 # o elemento nem existe mais
            if sumiu:
                break
        ACEITAR_PROXIMO_AVISO["ativo"] = False
        if not sumiu:
            raise RuntimeError("cliquei em Excluir e confirmei, mas as linhas continuaram no carrinho")
        apagadas += n
        page.wait_for_timeout(800)                           # a próxima página de linhas pode aparecer
    if apagadas:
        print(f"  ✔ {apagadas} linha(s) antiga(s) apagada(s)")
    return apagadas


def _num_requisicao(url):
    m = re.search(r"requisition_headers/(\d+)", url or "")
    return m.group(1) if m else None


def lancar_no_coupa(page, estado, perfil_nome=None, perfil=None, url_anterior=None):
    """Lança (ou continua lançando) as linhas do 'estado', gravando o progresso a cada passo.
    estado = {"linhas": [...], "salvas": [nº...], "contas": [nº...], "url": ..., "cabecalho": bool}"""
    linhas = estado["linhas"]
    page.bring_to_front()

    # 1. Requisição: abre pelo carrinho, ou volta na que já estava sendo lançada
    if estado.get("url"):
        print(f"\nVoltando para a requisição em andamento: {estado['url']}")
        ir_para_coupa(page, estado["url"])
        garantir_na_requisicao(page, estado["url"])
    else:
        estado["url"] = abrir_requisicao(page)
        # o carrinho ainda é a requisição da nota anterior? (não foi submetida) - não mistura as duas
        for _ in range(3):
            if not url_anterior or _num_requisicao(estado["url"]) != _num_requisicao(url_anterior):
                break
            input(f"\n⚠ O carrinho ainda é a requisição anterior (nº {_num_requisicao(url_anterior)}). "
                  "Submeta (ou esvazie) ela no Coupa e aperte Enter aqui...")
            ir_para_coupa(page)
            garantir_login(page)
            estado["url"] = abrir_requisicao(page)
        print(f"  Requisição nº {_num_requisicao(estado['url']) or '?'}")
        pasta = nomear_pasta_requisicao(_num_requisicao(estado["url"]))
        if pasta:
            estado["pasta"] = str(pasta)
            print(f"  📁 Relatórios desta requisição: {pasta}")
        if APAGAR_LINHAS_EXISTENTES:
            apagar_linhas_existentes(page)
        salvar_progresso(estado)

    # 2. Cabeçalho (só se ainda não foi feito)
    if not estado.get("cabecalho"):
        if perfil:
            preencher_cabecalho(page, perfil_nome, perfil)
        else:
            input("\nPreencha o cabeçalho da requisição e aperte Enter aqui para continuar...")
        estado["cabecalho"] = True
        salvar_progresso(estado)

    # 3. Linhas
    vistos = set()
    for item in linhas:
        item["plu_repetido"] = item["plu"] in vistos
        vistos.add(item["plu"])
    salvas, contas = set(estado.get("salvas", [])), set(estado.get("contas", []))
    if salvas:
        print(f"  {len(salvas)} linha(s) já salvas antes - continuando a partir das que faltam.")

    for i, item in enumerate(linhas, start=1):
        if i in salvas and i in contas:
            continue
        print(f"[{i}/{len(linhas)}] PLU {item['plu']} | CC {item['cc']} | R$ {br(item['total'])}")
        ponto_de_pausa(page, f"antes da linha {i}")
        garantir_na_requisicao(page, estado["url"])

        if i not in salvas:
            falhas = 0
            while True:
                try:
                    _abrir_ou_reusar_linha(page, f"linha-{i}")
                    preencher_linha(page, item)
                    salvar_linha(page)
                    print("  ✔ salva")
                    salvas.add(i)
                    estado["salvas"] = sorted(salvas)
                    salvar_progresso(estado)
                    break
                except ParadaSolicitada:
                    raise
                except Exception as e:
                    if conexao_caiu(e):
                        raise
                    falhas += 1
                    msg = str(e).splitlines()[0] if str(e) else type(e).__name__
                    if falhas <= TENTATIVAS_AUTOMATICAS:
                        print(f"  ⚠ {msg} — tentando de novo sozinho ({falhas}/{TENTATIVAS_AUTOMATICAS})...")
                        try:
                            page.keyboard.press("Escape")
                            page.wait_for_timeout(1500)
                        except Exception:
                            pass
                        garantir_na_requisicao(page, estado["url"])
                        continue
                    print(f"  ✖ Erro em {item['origem']}: {msg}")
                    resp = input("  Corrija/cancele a linha no Coupa. "
                                 "[Enter] tentar de novo | [p] pular | [s] sair (o progresso fica salvo): "
                                 ).strip().lower()
                    if resp == "p":
                        break
                    if resp == "s":
                        raise ParadaSolicitada("você escolheu sair")
                    falhas = 0

        if i in salvas and i not in contas:
            if ajustar_conta_seguro(page, item):
                contas.add(i)
                estado["contas"] = sorted(contas)
                salvar_progresso(estado)

    # 4. Conferência automática com a nota + revisão final
    faltando = [i for i in range(1, len(linhas) + 1) if i not in salvas]
    if faltando:
        print(f"\n⚠ Linhas NÃO lançadas (puladas): {', '.join(map(str, faltando))}")
    try:
        ponto_de_pausa(page, "conferência")
        garantir_na_requisicao(page, estado["url"])
        import io
        import contextlib

        _SAIDA = sys.stdout

        class _Tee(io.StringIO):                 # mostra no terminal E guarda para o conferencia.txt
            def write(self, t):
                _SAIDA.write(t)
                return super().write(t)
        buf = _Tee()
        with contextlib.redirect_stdout(buf):
            divergencias = conferir_requisicao(page, linhas, estado.get("total_nota"))
        gravar_na_pasta("conferencia.txt", buf.getvalue())
    except ParadaSolicitada:
        raise
    except Exception as e:
        if conexao_caiu(e):
            raise
        print(f"  ⚠ A conferência automática falhou ({str(e).splitlines()[0]}) - confira à mão.")
        divergencias = ["conferência não concluída"]
    gravar_resumo(estado, divergencias,
                  "concluída" if not faltando else f"linhas puladas: {', '.join(map(str, faltando))}")
    if divergencias:
        input("\nCorrija as divergências no Coupa, revise e SUBMETA a requisição; depois aperte Enter aqui...")
    else:
        input("\nRevise e SUBMETA a requisição no Coupa; depois aperte Enter aqui...")
    return not faltando


def ir_para_coupa(page, url=None):
    """page.goto que aguenta o redirecionamento do login da Microsoft: quando a sessão expirou,
    o Coupa manda para login.microsoftonline.com no meio da navegação e o goto normal dá erro
    ("interrupted by another navigation"). Aqui isso não é erro - só significa que precisa logar."""
    try:
        page.goto(url or URL_COUPA, wait_until="domcontentloaded")
    except Exception as e:
        t = str(e)
        if not any(x in t for x in ("interrupted by another navigation", "ERR_ABORTED", "net::ERR_")):
            raise
        page.wait_for_timeout(1500)


def garantir_login(page):
    """Espera o Coupa abrir logado (barra de busca visível). Se cair na tela da Microsoft,
    pede para a pessoa fazer o login e tenta de novo."""
    for _ in range(5):
        try:
            busca_visivel(page).wait_for(state="visible", timeout=15000)
            return
        except Exception:
            pass
        if "coupahost.com" in page.url and "login" not in page.url.lower():
            ir_para_coupa(page)                     # está no Coupa, mas em outra tela: volta ao início
            try:
                busca_visivel(page).wait_for(state="visible", timeout=10000)
                return
            except Exception:
                pass
        input("\nFaça o login no Coupa na janela do Edge (conta Microsoft) e aperte Enter aqui "
              "quando a página inicial do Coupa abrir...")
        ir_para_coupa(page)
    raise SystemExit("\n✖ Não consegui chegar à página inicial do Coupa logado. Confira o login e rode de novo.")


def abrir_navegador(p):
    """Abre o Edge com o perfil da macro numa aba NOVA e vai para o Coupa.
    Depois de uma execução que caiu, o Edge pode abrir restaurando abas e fechar a primeira -
    por isso não usa a aba que já vem aberta e tenta de novo se ela fechar."""
    try:
        ctx = p.chromium.launch_persistent_context(
            PERFIL_EDGE, channel="msedge", headless=False,
            args=["--hide-crash-restore-bubble", "--no-first-run", "--no-default-browser-check"],
        )
    except Exception as e:
        raise SystemExit(
            f"\n✖ Não consegui abrir o Edge com o perfil {PERFIL_EDGE}: {str(e).splitlines()[0]}\n"
            "  Provavelmente ficou um Edge aberto da execução anterior usando esse perfil.\n"
            "  Feche as janelas do Edge abertas pela macro (ou finalize 'Microsoft Edge' no Gerenciador\n"
            "  de Tarefas) e rode de novo.")
    vigiar_avisos(ctx)
    for tentativa in range(3):
        try:
            page = ctx.new_page()
            for extra in ctx.pages:                  # fecha abas restauradas/em branco
                if extra is not page:
                    try:
                        extra.close()
                    except Exception:
                        pass
            ir_para_coupa(page)
            return ctx, page
        except Exception as e:
            if "has been closed" in str(e) and tentativa < 2:
                try:
                    ctx.pages                        # contexto ainda vivo? tenta outra aba
                    continue
                except Exception:
                    break
            raise SystemExit(
                f"\n✖ O Edge fechou logo ao abrir ({str(e).splitlines()[0]}).\n"
                "  Feche todas as janelas do Edge abertas pela macro (ou finalize 'Microsoft Edge'\n"
                "  no Gerenciador de Tarefas) e rode de novo.")
    raise SystemExit("\n✖ O Edge fechou logo ao abrir. Feche as janelas do Edge da macro e rode de novo.")


def _descartar_progresso():
    arq = _arq_progresso()
    if arq.exists():
        try:
            arq.replace(pasta_controle() / f"progresso_descartado_{datetime.now():%Y%m%d_%H%M%S}.json")
        except Exception:
            pass


def _print_erro(page, erro):
    """Erro inesperado: registra tudo no log e tira um print, sem derrubar a macro com traceback."""
    print(f"\n✖ Erro inesperado: {type(erro).__name__}: {str(erro).splitlines()[0] if str(erro) else ''}")
    traceback.print_exc(file=sys.stdout)            # vai para o log (e aparece no terminal)
    if page is not None:
        try:
            arq = pasta_atual() / f"erro_{datetime.now():%Y%m%d_%H%M%S}.png"
            page.screenshot(path=str(arq), full_page=False)
            print(f"  Print da tela salvo em {arq.name}")
        except Exception:
            pass


_LOJAS = None


def lojas_em_cache():
    global _LOJAS
    if _LOJAS is None:
        _LOJAS = carregar_lojas_ou_sair()
    return _LOJAS


def perguntar_configuracao(wb, anterior=None):
    """Modo, perfil e nº da conta. Com 'anterior', repete modo e perfil (só reconfirma a coluna G)."""
    if anterior:
        modo, perfil_nome, perfil = anterior["modo"], anterior["perfil_nome"], anterior["perfil"]
    else:
        modo = input("\nEntrada: [1] planilha  [2] nota (texto do PDF ou XML) (padrão): ").strip()
        modo = "1" if modo == "1" else "2"
        perfil_nome, perfil = escolher_perfil()
    valor_conta = ""
    if perfil and perfil.get("conta_valor") == "coluna G":
        valor_conta = confirmar_valor_conta(perfil, valor_coluna_g(wb))
    tr = ""
    if perfil and perfil.get("campo_tr") == "coluna H":
        tr = valor_coluna_h(wb)
        if modo == "2":           # na nota, um texto só para todas as linhas
            r = input(f"Texto do campo TR (coluna H): {tr or '(vazio)'}. Enter confirma ou digite outro: ").strip()
            tr = r or tr
    return {"modo": modo, "perfil_nome": perfil_nome, "perfil": perfil, "valor_conta": valor_conta, "tr": tr}


def _descrever(config):
    if not config:
        return ""
    return (("planilha" if config["modo"] == "1" else "nota") + " · perfil "
            + (config["perfil_nome"] or "manual"))


def main():
    faltando = [n for n, v in (("COUPA_URL", URL_COUPA), ("COD_EMPRESA", COD_EMPRESA)) if not v]
    if faltando:
        print(f"Configure {', '.join(faltando)} no arquivo .env (veja o .env.example).")
        return
    log = iniciar_log()
    if log:
        print(f"📝 Log desta execução: {log}")
    if not Path(CAMINHO_PLANILHA).exists():
        print(f"Planilha não encontrada em:\n  {CAMINHO_PLANILHA}\n"
              "Confira se o arquivo está nesse caminho ou ajuste CAMINHO_PLANILHA no .env.")
        return
    migrar_arquivos_antigos()
    print(f"📁 Relatórios em: {pasta_base()}")
    aplicar_pendente(carregar_plus(load_workbook(CAMINHO_PLANILHA, data_only=True)))
    iniciar_teclas()

    with sync_playwright() as p:
        nav = {"ctx": None, "page": None}

        def navegador():
            """Abre o Edge na 1ª vez e reabre se ele tiver caído - o login fica salvo entre as notas."""
            vivo = False
            if nav["page"] is not None:
                try:
                    _ = nav["page"].url
                    nav["page"].evaluate("1")
                    vivo = True
                except Exception:
                    vivo = False
            if not vivo:
                if nav["ctx"] is not None:
                    try:
                        nav["ctx"].close()
                    except Exception:
                        pass
                nav["ctx"], nav["page"] = abrir_navegador(p)
                garantir_login(nav["page"])
            return nav["ctx"], nav["page"]

        config_anterior, ultima_url, rodada = None, None, 0
        try:
            while True:
                rodada += 1
                wb = load_workbook(CAMINHO_PLANILHA, data_only=True)    # relê: a planilha pode ter mudado
                estado, retomar = carregar_progresso(), False
                if estado and estado.get("linhas"):
                    print("\n⚠ Há uma requisição que não terminou:\n    " + resumo_progresso(estado))
                    r = input("[c] continuar de onde parou | [n] começar uma nova (descarta a anterior) | [s] sair: "
                              ).strip().lower()
                    if r == "s":
                        break
                    if r == "c":
                        retomar = True
                    else:
                        _descartar_progresso()
                        estado = None

                if retomar:
                    perfil_nome = estado.get("perfil_nome")
                    perfil = PERFIS.get(perfil_nome) if perfil_nome else None
                    config = None
                    iniciar_pasta_requisicao("retomada", existente=estado.get("pasta"))
                else:
                    if rodada > 1:
                        print("\n" + "=" * 60 + "\nNOVA REQUISIÇÃO\n" + "=" * 60)
                    config = perguntar_configuracao(wb, config_anterior)
                    perfil_nome, perfil = config["perfil_nome"], config["perfil"]
                    if config["modo"] == "2":
                        dados = ler_nota_completa(lojas_em_cache())
                        if not dados:
                            print("Nada para lançar.")
                            config = None
                    if config:
                        pasta = iniciar_pasta_requisicao(descricao_requisicao(config["modo"]))
                        if config["modo"] == "2":
                            guardar_nota_na_pasta()
                        print(f"📁 Pasta desta requisição: {pasta}")
                    estado = None

                try:
                    if retomar or config:
                        ctx, page = navegador()
                        if not retomar:
                            modo = config["modo"]
                            permitir = bool(perfil and perfil.get("permitir_plu_repetido"))
                            pular = modo == "1" and perfil and perfil.get("verificar_plu_planilha") is False
                            verificador = VerificadorCoupa(ctx) if VERIFICAR_NO_COUPA and not pular else None
                            if modo == "1":
                                linhas = ler_planilha(wb, permitir_repetido=permitir)
                                if linhas and verificador and not verificar_planilha(linhas, wb, verificador):
                                    linhas = []
                            else:
                                itens, cnpj_forn, cc, total_ref = dados
                                linhas = montar_linhas_da_nota(wb, itens, cnpj_forn, cc, verificador,
                                                               permitir_repetido=permitir)
                            if verificador:
                                verificador.fechar()
                            if linhas:
                                aplicar_perfil_nas_linhas(linhas, perfil, config["valor_conta"], config.get("tr", ""))
                                ajustar_quantidades(linhas)
                                problemas = validar_linhas(linhas)
                                if problemas:
                                    print("\n⚠ Problemas encontrados antes de lançar:")
                                    for pr in problemas:
                                        print("   -", pr)
                                    if input("Lançar mesmo assim? [s/n]: ").strip().lower() != "s":
                                        linhas = []
                            if linhas:
                                estado = {"inicio": f"{datetime.now():%d/%m/%Y %H:%M}", "perfil_nome": perfil_nome,
                                          "linhas": linhas, "salvas": [], "contas": [], "url": None,
                                          "cabecalho": False,
                                          "total_nota": (total_ref if config["modo"] == "2" else None),
                                          "pasta": str(pasta_atual())}
                                salvar_progresso(estado)
                            else:
                                print("Nada para lançar.")

                        if estado:
                            completo = lancar_no_coupa(page, estado, perfil_nome, perfil, url_anterior=ultima_url)
                            ultima_url = estado.get("url") or ultima_url
                            if completo:
                                encerrar_progresso()
                                print(f"✔ Requisição nº {_num_requisicao(ultima_url) or '?'} concluída.")
                            else:
                                print("  As linhas puladas continuam no progresso (opção [c] na próxima nota).")

                except (ParadaSolicitada, KeyboardInterrupt) as e:
                    print(f"\n⏹ Execução parada{(' (' + str(e) + ')') if str(e) else ''}.")
                    if estado:
                        print(f"  Progresso salvo: {len(estado.get('salvas', []))} de {len(estado['linhas'])} "
                              "linhas. Rode de novo e escolha [c] para continuar de onde parou.")
                    break
                except SystemExit:
                    raise
                except Exception as e:
                    if conexao_caiu(e):
                        print("\n✖ A conexão com o navegador caiu (ele será reaberto na próxima nota).")
                        if AVISOS_DO_SITE:
                            print(f"  Último aviso do site: {AVISOS_DO_SITE[-1]}")
                    else:
                        _print_erro(nav["page"], e)
                    if estado:
                        print(f"  Progresso salvo: {len(estado.get('salvas', []))} de {len(estado['linhas'])} "
                              "linhas - escolha [c] para continuar essa requisição.")

                if estado and carregar_progresso():
                    gravar_resumo(estado, status="interrompida - rode de novo e escolha [c] para continuar")
                fechar_pasta_requisicao()
                # ---- próxima nota?
                if config:
                    config_anterior = config
                mesmo = _descrever(config_anterior)
                r = input("\nO que fazer agora?"
                          + (f"  [Enter] outra nota ({mesmo})" if mesmo else "  [Enter] outra nota")
                          + "  |  [t] outra nota trocando modo/perfil  |  [f] finalizar: ").strip().lower()
                if r == "f":
                    break
                if r == "t":
                    config_anterior = None
                PAUSA.clear()
        finally:
            if nav["ctx"] is not None:
                try:
                    nav["ctx"].close()
                except Exception:
                    pass
    print("\nMacro finalizada.")


if __name__ == "__main__":
    main()