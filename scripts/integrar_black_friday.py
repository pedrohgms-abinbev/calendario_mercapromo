from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

# Permite reutilizar as funções de busca/incorporação de imagens do gerador atual.
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

try:
    import gerar_html as gerador
except ImportError as exc:  # pragma: no cover - mensagem operacional
    raise ImportError(
        "Não consegui importar scripts/gerar_html.py. "
        "Salve este arquivo em scripts/integrar_black_friday.py."
    ) from exc


ROOT = Path(__file__).resolve().parents[1]
ARQUIVO_HTML = ROOT / "site" / "calendario_mercapromo.html"
ARQUIVO_DADOS = ROOT / "data" / "processed" / "black_friday_cerveja_enriquecido.json"
ARQUIVO_RELATORIO = ROOT / "data" / "processed" / "relatorio_black_friday_html.txt"

UFS = {
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA",
    "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN",
    "RS", "RO", "RR", "SC", "SP", "SE", "TO",
}


# ============================================================
# FORMATAÇÃO E CLASSIFICAÇÃO
# ============================================================

def texto(valor: Any) -> str:
    if valor is None:
        return ""
    return str(valor).strip()


def numero_pt(valor: Any) -> str:
    if valor is None or valor == "":
        return ""
    try:
        numero = float(str(valor).replace(",", "."))
        if numero.is_integer():
            return str(int(numero))
        return str(numero).replace(".", ",")
    except (TypeError, ValueError):
        return texto(valor)


def divisao_e_uf(divisao: str) -> bool:
    return texto(divisao).upper() in UFS


def periodo_da_campanha(ofertas: list[dict[str, Any]]) -> str:
    inicios = [texto(o.get("INICIO")) for o in ofertas if texto(o.get("INICIO"))]
    fins = [texto(o.get("FIM")) for o in ofertas if texto(o.get("FIM"))]
    if inicios and fins:
        # A planilha já foi validada para o período único da campanha.
        return f"{inicios[0]} a {fins[0]}"
    return "23/11/2026 a 29/11/2026"


def carregar_ofertas() -> list[dict[str, Any]]:
    if not ARQUIVO_DADOS.is_file():
        raise FileNotFoundError(
            f"Base enriquecida da Black Friday não encontrada: {ARQUIVO_DADOS}\n"
            "Execute primeiro scripts/aplicar_depara_black_friday.py."
        )

    with ARQUIVO_DADOS.open("r", encoding="utf-8") as arquivo:
        dados = json.load(arquivo)

    if not isinstance(dados, list):
        raise ValueError("O JSON enriquecido precisa conter uma lista de ofertas.")

    return [item for item in dados if isinstance(item, dict)]


# ============================================================
# CARD BLACK FRIDAY
# ============================================================

def render_card(oferta: dict[str, Any], mostrar_divisao: bool) -> tuple[str, bool]:
    produto = texto(oferta.get("PRODUTO")) or "Produto"
    sku = texto(oferta.get("SKU")) or produto
    nome_depara = texto(oferta.get("DEPARA_NOME")) or produto
    divisao = texto(oferta.get("DIVISAO"))
    config = oferta.get("CONFIGURACAO") or {}
    if not isinstance(config, dict):
        config = {}

    caminho_imagem = texto(oferta.get("CAMINHO_IMAGEM"))
    nome_img = texto(oferta.get("NOME_IMG"))
    imagem_uri = ""
    if caminho_imagem:
        imagem_uri = gerador.image_data_uri(caminho_imagem)
    if not imagem_uri and nome_img:
        imagem_uri = gerador.image_data_uri(nome_img)
    if not imagem_uri:
        imagem_html = '<div class="bf-image-placeholder">Imagem indisponível</div>'
        imagem_encontrada = False
    else:
        imagem_html = (
            f'<img src="{imagem_uri}" alt="{gerador.esc(sku)}" loading="lazy">'
        )
        imagem_encontrada = True

    tipo = texto(config.get("tipo"))
    maximo = numero_pt(config.get("max_por_cpf"))

    if tipo == "leve_pague":
        leve = numero_pt(config.get("leve"))
        pague = numero_pt(config.get("pague"))
        promocao_html = (
            '<div class="bf-mechanic-label">LEVE E PAGUE</div>'
            '<div class="bf-take-pay">'
            f'<span><small>LEVE</small><strong>{gerador.esc(leve or "—")}</strong></span>'
            '<i>·</i>'
            f'<span><small>PAGUE</small><strong>{gerador.esc(pague or "—")}</strong></span>'
            '</div>'
        )
        limite_unidade = "combos"

    elif tipo == "desconto":
        desconto = numero_pt(config.get("desconto_percentual"))
        minimo = numero_pt(config.get("min_un_ativacao"))
        promocao_html = (
            '<div class="bf-mechanic-label">DESCONTO</div>'
            f'<div class="bf-discount-value">{gerador.esc(desconto or "—")}<span>% OFF</span></div>'
            + (
                '<div class="bf-threshold">'
                f'A partir de {gerador.esc(minimo)} '
                f'{"unidade" if minimo in {"1", "1,0"} else "unidades"}'
                '</div>'
                if minimo else ""
            )
        )
        limite_unidade = "unidades"

    else:
        promocao_html = (
            '<div class="bf-mechanic-label">MECÂNICA</div>'
            '<div class="bf-threshold">Consulte a mecânica da oferta</div>'
        )
        limite_unidade = "unidades"

    limite_html = ""
    if maximo and maximo not in {"0", "0,0", "-"}:
        limite_html = (
            '<div class="bf-limit">'
            f'Máx.: {gerador.esc(maximo)} {limite_unidade} por CPF'
            '</div>'
        )

    divisao_html = ""
    if mostrar_divisao and divisao:
        divisao_html = f'<span class="bf-division">{gerador.esc(divisao)}</span>'

    card = (
        '<article class="bf-product-card">'
        f'<div class="bf-product-image">{imagem_html}</div>'
        '<div class="bf-product-info">'
        f'<div class="bf-card-topline">{divisao_html}'
        '<span class="bf-card-type">OFERTA BLACK FRIDAY</span></div>'
        f'<h3 title="{gerador.esc(nome_depara)}">{gerador.esc(sku)}</h3>'
        f'<div class="bf-product-name">{gerador.esc(produto)}</div>'
        f'<div class="bf-promo-box">{promocao_html}{limite_html}</div>'
        '</div></article>'
    )
    return card, imagem_encontrada


# ============================================================
# MONTAGEM DA ABA BLACK FRIDAY
# ============================================================

def render_grid(ofertas: list[dict[str, Any]], mostrar_divisao: bool) -> tuple[str, int, int]:
    ordenadas = sorted(
        ofertas,
        key=lambda o: (
            texto(o.get("DIVISAO")).casefold(),
            texto(o.get("SKU")).casefold(),
            texto(o.get("PRODUTO")).casefold(),
        ),
    )

    cards: list[str] = []
    imagens_ok = 0
    for oferta in ordenadas:
        card, imagem_ok = render_card(oferta, mostrar_divisao)
        cards.append(card)
        imagens_ok += int(imagem_ok)

    if not cards:
        return '<div class="bf-empty">Não há ofertas cadastradas nesta visualização.</div>', 0, 0

    count = len(cards)
    columns = 1 if count == 1 else 2 if count <= 4 else 3 if count <= 9 else 4
    return (
        f'<div class="bf-offers-grid" style="--bf-columns:{columns}">'
        + "".join(cards)
        + "</div>",
        count,
        imagens_ok,
    )


def montar_black_friday(ofertas: list[dict[str, Any]]) -> tuple[str, dict[str, int]]:
    """Renderiza todas as ofertas como personalizadas, agrupadas por DIVISÃO.

    MT e MS aparecem primeiro, cada uma em sua própria aba. As demais
    divisões (redes/varejistas) também ganham uma aba individual.
    Não existe mais painel de ofertas gerais.
    """
    personalizadas: dict[str, list[dict[str, Any]]] = defaultdict(list)
    nomes_divisao: dict[str, str] = {}

    for oferta in ofertas:
        divisao_original = texto(oferta.get("DIVISAO"))
        divisao = divisao_original or "DIVISÃO NÃO INFORMADA"
        chave = divisao.casefold()
        nomes_divisao.setdefault(chave, divisao)
        personalizadas[chave].append(oferta)

    # MT e MS ficam sempre no início, cada uma em uma aba própria.
    ordem_prioritaria = {"mt": 0, "ms": 1}

    chaves_ordenadas = sorted(
        personalizadas,
        key=lambda chave: (
            ordem_prioritaria.get(nomes_divisao[chave].strip().casefold(), 2),
            nomes_divisao[chave].casefold(),
        ),
    )

    periodo = periodo_da_campanha(ofertas)
    botoes: list[str] = []
    paineis: list[str] = []
    total_cards = 0
    total_imagens = 0
    total_ufs = 0
    total_redes = 0

    for indice, chave in enumerate(chaves_ordenadas, start=1):
        divisao = nomes_divisao[chave]
        itens = personalizadas[chave]
        view_id = f"bf-personalizado-{indice}"
        ativa = indice == 1
        divisao_e_uf_atual = divisao.strip().upper() in UFS

        if divisao_e_uf_atual:
            total_ufs += 1
        else:
            total_redes += 1

        botoes.append(
            f'<button type="button" class="bf-filter-button{" active" if ativa else ""}" '
            f'data-bf-view-target="{view_id}" aria-pressed="{"true" if ativa else "false"}">'
            f'{gerador.esc(divisao)}</button>'
        )

        grid, quantidade, imagens = render_grid(
            itens,
            mostrar_divisao=False,
        )
        total_cards += quantidade
        total_imagens += imagens

        tipo_divisao = "UF" if divisao_e_uf_atual else "REDE"
        atributo_hidden = "" if ativa else " hidden"

        paineis.append(
            f'<section class="bf-view-panel" data-bf-view-panel="{view_id}"{atributo_hidden}>'
            '<div class="bf-section-heading"><div>'
            f'<span class="bf-eyebrow">OFERTAS PERSONALIZADAS · {tipo_divisao}</span>'
            f'<h3>{gerador.esc(divisao)}</h3></div>'
            f'<span class="bf-count">{quantidade} ofertas</span></div>'
            f'<div class="bf-network-period">Válido de {gerador.esc(periodo)}</div>'
            f'{grid}</section>'
        )

    botoes_html = "".join(botoes)
    paineis_html = "".join(paineis)
    markup = (
        '<section id="black-friday-campaign" class="bf-campaign" hidden>'
        '<header class="bf-hero">'
        '<div class="bf-hero-copy"><span class="bf-eyebrow">CAMPANHA ESPECIAL</span>'
        '<h2>BLACK <strong>FRIDAY</strong></h2>'
        '<p>Ofertas personalizadas por UF e rede</p></div>'
        f'<div class="bf-date-stamp"><span>PERÍODO DA CAMPANHA</span><strong>{gerador.esc(periodo)}</strong></div>'
        '</header>'
        '<nav class="bf-filter-shell" aria-label="Ofertas personalizadas Black Friday">'
        '<div class="bf-filter-row">'
        '<span class="bf-filter-caption">PERSONALIZADOS</span>'
        f'{botoes_html}'
        '</div></nav>'
        '<div class="bf-views">'
        f'{paineis_html}'
        '</div>'
        '<footer class="bf-footer">BLACK FRIDAY 2026 <span>·</span> Ofertas personalizadas válidas no período indicado</footer>'
        '</section>'
    )

    resumo = {
        "ofertas": len(ofertas),
        "ofertas_personalizadas": len(ofertas),
        "abas_personalizadas": len(personalizadas),
        "abas_uf": total_ufs,
        "abas_rede": total_redes,
        "cards_renderizados": total_cards,
        "imagens_localizadas": total_imagens,
    }
    return markup, resumo


# ============================================================
# INTEGRAÇÃO NO HTML EXISTENTE
# ============================================================

CAMPAIGN_SWITCH = r'''
<nav class="campaign-switch" aria-label="Campanhas">
  <button type="button" class="campaign-tab active" data-campaign-target="mercapromo" aria-pressed="true">Mercapromo</button>
  <button type="button" class="campaign-tab campaign-tab-bf" data-campaign-target="black-friday" aria-pressed="false">Black Friday <span>2026</span></button>
</nav>
'''.strip()

EXTRA_CSS = r'''
/* Navegação entre campanhas: mantém o Mercapromo original e destaca a Black Friday com discrição. */
.campaign-switch{display:flex;gap:8px;align-items:center;margin:0 0 12px;padding:7px;background:#f3f4f1;border:1px solid #c1c6c2;border-radius:12px;width:max-content;max-width:100%}
.campaign-tab{appearance:none;border:1px solid #aeb8bf;background:#fff;color:#20272d;border-radius:8px;padding:10px 16px;min-height:38px;font:900 12px/1 Arial,Helvetica,sans-serif;cursor:pointer;transition:all .16s ease}
.campaign-tab.active{background:#07519a;border-color:#07519a;color:#fff;box-shadow:0 2px 4px #07519a2b}
.campaign-tab-bf{border-color:#b69b5b;background:#faf8f1;color:#25231d}
.campaign-tab-bf span{margin-left:5px;color:#92732c;font-size:10px}
.campaign-tab-bf.active{background:#171717;border-color:#171717;color:#e8cc83;box-shadow:0 2px 5px #00000025}
.campaign-tab-bf.active span{color:#e8cc83}
.bf-campaign{margin:2px 0 12px}
.bf-hero{display:flex;align-items:center;justify-content:space-between;gap:18px;margin:0 0 12px;padding:18px 21px;background:linear-gradient(112deg,#171717 0%,#292720 100%);border:1px solid #8d7746;border-left:5px solid #c4a45b;border-radius:15px;color:#fff;box-shadow:0 4px 9px #00000016}
.bf-eyebrow{display:block;color:#b69a5d;font-size:10px;font-weight:1000;letter-spacing:1.8px;text-transform:uppercase;margin-bottom:5px}
.bf-hero h2{margin:0;color:#f8f6ef;font-size:clamp(27px,3vw,42px);line-height:1;font-weight:1000;letter-spacing:1px}
.bf-hero h2 strong{color:#dfc27b}
.bf-hero-copy p{margin:7px 0 0;color:#d4d0c5;font-size:12px;font-weight:700}
.bf-date-stamp{display:flex;flex-direction:column;gap:5px;text-align:right;padding:10px 13px;border:1px solid #71603c;border-radius:9px;background:#ffffff0b;min-width:180px}
.bf-date-stamp span{font-size:9px;font-weight:900;letter-spacing:1.2px;color:#c8b47e}
.bf-date-stamp strong{font-size:15px;color:#fff;font-weight:1000;white-space:nowrap}
.bf-filter-shell{margin:0 0 12px;padding:10px 12px;background:#f2f1eb;border:1px solid #c8c1ad;border-radius:12px}
.bf-filter-row{display:flex;align-items:center;gap:7px;flex-wrap:wrap}
.bf-filter-caption{font-size:9px;letter-spacing:1.1px;font-weight:1000;color:#766a4d;margin:0 3px}
.bf-filter-button{appearance:none;border:1px solid #c3b58f;border-radius:8px;background:#fff;color:#302c22;padding:9px 13px;min-height:35px;font:900 11px/1 Arial,Helvetica,sans-serif;cursor:pointer;transition:all .15s ease}
.bf-filter-button:hover{border-color:#9e803e}
.bf-filter-button.active{background:#24231f;border-color:#24231f;color:#e7ce8e;box-shadow:0 2px 4px #0000001c}
.bf-views{min-width:0}
.bf-view-panel[hidden]{display:none!important}
.bf-section-heading{display:flex;align-items:center;justify-content:space-between;gap:10px;margin:0 2px 8px;padding:0 2px}
.bf-section-heading h3{margin:0;color:#20201e;font-size:clamp(19px,1.6vw,25px);font-weight:1000;letter-spacing:-.3px;text-transform:uppercase}
.bf-count{font-size:10px;font-weight:900;color:#6b6048;background:#f3eddd;border:1px solid #d7c89f;border-radius:20px;padding:5px 9px;white-space:nowrap}
.bf-section-note,.bf-network-period{margin:-3px 4px 9px;color:#6e706e;font-size:10px;font-weight:700}
.bf-network-period{margin-top:-2px}
.bf-offers-grid{display:grid;grid-template-columns:repeat(var(--bf-columns,4),minmax(0,1fr));gap:9px;align-items:stretch}
.bf-product-card{display:grid;grid-template-columns:minmax(76px,30%) minmax(0,1fr);align-items:center;gap:9px;min-width:0;min-height:142px;padding:9px;background:linear-gradient(135deg,#f7f7f3 0%,#eeefea 100%);border:1px solid #c7c9c2;border-top:3px solid #b79a55;border-radius:12px;box-shadow:0 3px 6px #0000000d}
.bf-product-image{display:flex;align-items:center;justify-content:center;min-width:0;min-height:112px}
.bf-product-image img{display:block;max-width:100%;max-height:112px;object-fit:contain;filter:drop-shadow(0 2px 2px #00000013)}
.bf-image-placeholder{width:90%;height:70px;display:flex;align-items:center;justify-content:center;text-align:center;border:1px dashed #c4c7c2;border-radius:8px;color:#7a7e7a;font-size:9px}
.bf-product-info{display:flex;flex-direction:column;align-items:stretch;gap:5px;min-width:0}
.bf-card-topline{display:flex;align-items:center;justify-content:space-between;gap:5px;min-height:17px}
.bf-division{display:inline-flex;padding:3px 6px;border-radius:5px;background:#e9f0f5;border:1px solid #c1d1dc;color:#07519a;font-size:9px;font-weight:1000}
.bf-card-type{color:#8d7746;font-size:8px;font-weight:1000;letter-spacing:.7px;text-align:right}
.bf-product-info h3{margin:0;color:#161819;font-size:clamp(14px,1.05vw,18px);font-weight:1000;line-height:1.05;text-transform:uppercase;overflow-wrap:anywhere}
.bf-product-name{color:#697075;font-size:9px;font-weight:700;line-height:1.15;overflow-wrap:anywhere}
.bf-promo-box{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4px;min-height:66px;padding:7px 6px;text-align:center;background:#f0e5cb;border:1px solid #d2bd8b;border-radius:9px;color:#1e1d19}
.bf-mechanic-label{font-size:8px;line-height:1;font-weight:1000;letter-spacing:1.1px;color:#766039}
.bf-take-pay{display:flex;align-items:center;justify-content:center;gap:10px;line-height:1}
.bf-take-pay span{display:flex;align-items:center;gap:4px}
.bf-take-pay small{font-size:9px;font-weight:1000;color:#555044}
.bf-take-pay strong{font-size:clamp(23px,1.8vw,31px);font-weight:1000;color:#111}
.bf-take-pay i{font-style:normal;color:#9d8247;font-size:17px;font-weight:1000}
.bf-discount-value{font-size:clamp(25px,2vw,34px);line-height:1;font-weight:1000;color:#111;white-space:nowrap}
.bf-discount-value span{font-size:.42em;margin-left:4px;letter-spacing:.3px}
.bf-threshold{font-size:10px;font-weight:900;color:#48453c;line-height:1.15}
.bf-limit{font-size:9px;font-weight:900;line-height:1.15;color:#615b4d}
.bf-empty{padding:18px;text-align:center;background:#f7f7f4;border:1px dashed #c6c2b6;border-radius:10px;color:#74736d;font-size:11px;font-weight:700}
.bf-footer{margin:14px 3px 2px;padding-top:10px;border-top:1px solid #d0c8b3;color:#81765c;font-size:9px;font-weight:900;letter-spacing:1px;text-align:right}
.bf-footer span{padding:0 4px;color:#b29755}
@media(max-width:1100px){.bf-offers-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:760px){.campaign-switch{width:100%;}.campaign-tab{flex:1}.bf-hero{align-items:flex-start;flex-direction:column;padding:15px}.bf-date-stamp{min-width:0;width:100%;text-align:left}.bf-offers-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.bf-product-card{grid-template-columns:minmax(62px,29%) minmax(0,1fr);gap:6px;padding:7px;min-height:132px}.bf-product-image{min-height:95px}.bf-product-image img{max-height:95px}.bf-take-pay{gap:5px}.bf-take-pay strong{font-size:23px}.bf-product-info h3{font-size:13px}}
@media(max-width:470px){.bf-offers-grid{grid-template-columns:1fr}.bf-product-card{grid-template-columns:82px minmax(0,1fr);min-height:126px}.bf-product-image{min-height:94px}.bf-product-image img{max-height:90px}.bf-hero h2{font-size:30px}}
@media print{.campaign-switch{display:none}.bf-campaign[hidden]{display:none!important}.bf-product-card{break-inside:avoid;box-shadow:none}}
'''

EXTRA_JS = r'''
<script>
(function(){
  const campaignButtons=[...document.querySelectorAll('[data-campaign-target]')];
  const mercapromoElements=[
    document.querySelector('.month-stamp'),
    document.querySelector('.filter-shell'),
    document.querySelector('#calendar-views'),
    document.querySelector('.page-footer')
  ].filter(Boolean);
  const bfSection=document.querySelector('#black-friday-campaign');
  const brand=document.querySelector('.brand-title h1');
  const subtitle=document.querySelector('.brand-subtitle');
  const originalBrand=brand?brand.textContent:'MERCAPROMO';
  const originalSubtitle=subtitle?subtitle.textContent:'CALENDÁRIO';

  function activateCampaign(name){
    const isBlackFriday=name==='black-friday';
    mercapromoElements.forEach(el=>{el.hidden=isBlackFriday;});
    if(bfSection) bfSection.hidden=!isBlackFriday;
    campaignButtons.forEach(button=>{
      const active=button.dataset.campaignTarget===name;
      button.classList.toggle('active',active);
      button.setAttribute('aria-pressed',active?'true':'false');
    });
    if(brand) brand.textContent=isBlackFriday?'BLACK FRIDAY':originalBrand;
    if(subtitle) subtitle.textContent=isBlackFriday?'OFERTAS':'CALENDÁRIO';
    window.scrollTo({top:0,behavior:'smooth'});
  }

  campaignButtons.forEach(button=>button.addEventListener('click',()=>activateCampaign(button.dataset.campaignTarget)));

  const bfButtons=[...document.querySelectorAll('[data-bf-view-target]')];
  const bfPanels=[...document.querySelectorAll('[data-bf-view-panel]')];
  function activateBlackFridayView(id){
    bfPanels.forEach(panel=>{panel.hidden=panel.dataset.bfViewPanel!==id;});
    bfButtons.forEach(button=>{
      const active=button.dataset.bfViewTarget===id;
      button.classList.toggle('active',active);
      button.setAttribute('aria-pressed',active?'true':'false');
    });
  }
  bfButtons.forEach(button=>button.addEventListener('click',()=>activateBlackFridayView(button.dataset.bfViewTarget)));
  const firstBfButton=bfButtons[0];
  if(firstBfButton) activateBlackFridayView(firstBfButton.dataset.bfViewTarget);
})();
</script>
'''


def integrar() -> dict[str, int]:
    if not ARQUIVO_HTML.is_file():
        raise FileNotFoundError(
            f"HTML do Mercapromo não encontrado: {ARQUIVO_HTML}. "
            "Execute primeiro scripts/gerar_html.py."
        )

    if not ARQUIVO_DADOS.is_file():
        print(
            "AVISO: base Black Friday enriquecida não encontrada. "
            "O HTML do Mercapromo foi mantido sem a nova aba."
        )
        return {"integrado": 0}

    html_atual = ARQUIVO_HTML.read_text(encoding="utf-8")
    if 'id="black-friday-campaign"' in html_atual:
        raise RuntimeError(
            "O HTML já contém a Black Friday. Gere novamente o HTML original "
            "com scripts/gerar_html.py antes de integrar para evitar duplicações."
        )

    ofertas = carregar_ofertas()
    markup, resumo = montar_black_friday(ofertas)

    # Insere a navegação logo abaixo do cabeçalho original.
    anchor_switch = '</header><div class="month-stamp">'
    if anchor_switch not in html_atual:
        raise ValueError(
            "Não encontrei o ponto de inserção do seletor de campanhas no HTML. "
            "Confira se o arquivo foi gerado pelo gerar_html.py informado."
        )
    html_atual = html_atual.replace(
        anchor_switch,
        '</header>' + CAMPAIGN_SWITCH + '<div class="month-stamp">',
        1,
    )

    # Insere a campanha antes do rodapé compartilhado.
    anchor_bf = '<footer class="page-footer">'
    if anchor_bf not in html_atual:
        raise ValueError("Não encontrei o rodapé do HTML para inserir a Black Friday.")
    html_atual = html_atual.replace(anchor_bf, markup + anchor_bf, 1)

    # Acrescenta CSS sem substituir o CSS já existente.
    if "/* Navegação entre campanhas" not in html_atual:
        if "</style>" not in html_atual:
            raise ValueError("Não encontrei o fechamento do bloco CSS no HTML.")
        html_atual = html_atual.replace("</style>", EXTRA_CSS + "\n</style>", 1)

    # Acrescenta o JavaScript de alternância antes do fechamento do body.
    if "data-campaign-target" in html_atual and "function activateCampaign(name)" not in html_atual:
        if "</body>" not in html_atual:
            raise ValueError("Não encontrei o fechamento do body no HTML.")
        html_atual = html_atual.replace("</body>", EXTRA_JS + "\n</body>", 1)

    ARQUIVO_HTML.write_text(html_atual, encoding="utf-8")

    linhas = [
        "RELATÓRIO DE INTEGRAÇÃO BLACK FRIDAY",
        "",
        f"Ofertas na base: {resumo['ofertas']}",
        f"Ofertas personalizadas: {resumo['ofertas_personalizadas']}",
        f"Abas personalizadas: {resumo['abas_personalizadas']}",
        f"Abas de UF: {resumo['abas_uf']}",
        f"Abas de rede: {resumo['abas_rede']}",
        f"Cards renderizados: {resumo['cards_renderizados']}",
        f"Imagens localizadas durante a geração: {resumo['imagens_localizadas']}",
        f"HTML: {ARQUIVO_HTML}",
    ]
    ARQUIVO_RELATORIO.parent.mkdir(parents=True, exist_ok=True)
    ARQUIVO_RELATORIO.write_text("\n".join(linhas) + "\n", encoding="utf-8")

    print("\n" + "=" * 60)
    print("BLACK FRIDAY INTEGRADA AO HTML")
    print("=" * 60)
    for chave, valor in resumo.items():
        print(f"{chave.replace('_', ' ').capitalize()}: {valor}")
    print(f"\nHTML atualizado: {ARQUIVO_HTML}")
    print(f"Relatório: {ARQUIVO_RELATORIO}")
    return resumo


if __name__ == "__main__":
    integrar()
