# Calendário Mercapromo — foco nas mecânicas

Este projeto gera um HTML estático no estilo do material de referência do Mercapromo. O foco dos cards é a mecânica promocional — por exemplo, **Leve 6 / Pague 5** ou **Desconto 20% / Mínimo de 3 itens** — e não preços. Não há abas por UF: ofertas iguais são consolidadas e diferenças de regra recebem um sinalizador.

## Estrutura do projeto

```text
mercapromo_calendario/
├── scripts/
│   ├── explorar_excel.py
│   ├── tratamento.py
│   ├── gerar_html.py
│   └── atualizar.py
├── data/
│   ├── unprocessed/
│   │   └── Calendario Mercapromo.xlsx
│   └── processed/
├── cadastro/
│   └── de_para_mercapromo.xlsx
├── images/
│   ├── logo.png
│   └── img/
├── site/
└── requirements.txt
```

## Fonte e colunas

A fonte de ofertas é `data/unprocessed/Calendario Mercapromo.xlsx`. O script lê as abas `CERVEJA`, `NAB`, `MATCH` e `PERSONALIZADO`. **As colunas são lidas pela posição física/letra, não pelos nomes dos cabeçalhos.** Isso permite que H:K tenham um cabeçalho mesclado.

### CERVEJA, NAB e MATCH

| Coluna física | Uso |
|---|---|
| B | Divisão/abrangência |
| C | Início |
| D | Fim |
| F | Produto |
| G | Mecânica/descritivo |
| H | Rótulo do tipo: `Leve:` ou `Desconto:` |
| I | Quantidade a levar ou porcentagem |
| J | Rótulo de ativação: `Pague:` ou mínimo de unidades |
| K | Quantidade a pagar ou mínimo para ativar |
| Q | Máximo de combos por CPF (rodapé do painel) |

Os cabeçalhos principais estão na linha 1 e os dados começam na linha 2.

### PERSONALIZADO

A linha 1 contém uma instrução; os cabeçalhos estão na linha 2 e os dados começam na linha 3. O script usa a aba chamada exatamente `PERSONALIZADO` e ignora `PERSONALIZADO (2)`, que é uma aba separada/alternativa.

| Coluna física | Uso |
|---|---|
| A | GEO/abrangência |
| C | Varejista/rede |
| D | Início |
| E | Fim |
| G | Produto |
| H | Mecânica/descritivo |
| I | Rótulo: `Leve:` ou `Desconto:` |
| J | Quantidade a levar ou porcentagem |
| K | Rótulo de ativação: `Pague:` ou mínimo de unidades |
| L | Quantidade a pagar ou mínimo para ativar |
| N | Máximo por cliente/CPF |

A coluna F (tipo de oferta) e os campos auxiliares à direita não são necessários para montar a mecânica principal do card.

## Navegação e layout do HTML

A página tem um filtro principal **Calendário principal** e um filtro por rede para PERSONALIZADOS. O calendário principal mostra as duas quinzenas em sequência vertical, sem filtro para alternar entre elas. Cada quinzena reúne CERVEJA, NAB e MATCH nos três painéis coloridos do modelo de referência, com CERVEJA mais largo e NAB/MATCH ao lado. Ao selecionar uma rede, o calendário principal é ocultado e aparece somente a rede selecionada; clicar em Calendário principal volta à visualização principal. Assim, as ofertas de PERSONALIZADOS não ficam aparecendo no final da visualização principal.

- **PERSONALIZADOS · POR REDE**: há um filtro por rede-base, por exemplo, Comper e Fort. Ao selecionar uma rede, suas duas quinzenas aparecem em sequência vertical; em cada uma, as ofertas ficam divididas em CERVEJA, NAB e MATCH, usando o tipo cadastrado no De/Para (`CERV`, `NAB` ou `MATCH`). Itens sem tipo reconhecido ficam numa pequena seção `REVISAR CADASTRO` para não desaparecerem.
- O asterisco de abrangência é usado tanto no calendário principal quanto nos personalizados: `*DF e GO`, `*MT e MS`, etc. Registros `TOTAL`/`Todas as divisões` não recebem esse selo. A rede Fort é exibida simplesmente como **Fort**; como toda a rede é de atacado, o texto/sinalizador `Atacado` é omitido.
- Quando uma oferta idêntica vale em mais de uma região, os escopos aparecem reunidos, por exemplo, `*DF e GO`. Se a mecânica muda por divisão/região, os cards mantêm os selos respectivos e recebem um aviso de condição variável.
- Os painéis ficam alinhados dentro de cada quinzena, com espaçamento compacto e sem altura mínima artificial de tela cheia. Se houver muitas ofertas, o conteúdo continua acessível por rolagem vertical.
- Em telas menores, CERVEJA, NAB e MATCH empilham para preservar a legibilidade.

## Mecânicas exibidas

- `H=Leve:`, `I=6`, `J=Pague:`, `K=5` (ou `I=6`, `J=Pague:`, `K=5` no personalizado) → **Leve 6 / Pague 5**.
- `H=Desconto:`, `I=20`, `J=Min. Un. Ativação:`, `K=3` (equivalentes I:L na aba PERSONALIZADO) → **Desconto 20% / Mínimo de 3 itens**.
- Não são exibidos preços, mesmo quando há colunas de preço auxiliares no Excel.
- O limite de combos por CPF, quando preenchido em Q (CERVEJA/NAB/MATCH) ou N (PERSONALIZADO), aparece no rodapé do painel.

## De/Para e imagens

O arquivo opcional/recomendado `cadastro/de_para_mercapromo.xlsx` deve conter `Tipo`, `Promax`, `EG`, `SKU`, `Nome IMG`, `Caminho` e `Nome`.

- `Nome` é a descrição completa do produto e é a primeira referência para relacionar o nome na coluna `PRODUTO` das ofertas.
- `SKU` é o nome abreviado exibido no card (ex.: `AP 269`); `Nome IMG` e `Caminho` indicam o nome/arquivo da imagem.
- O tratamento primeiro compara `PRODUTO` com `Nome` (permitindo que o nome da promoção seja uma parte do nome completo do De/Para), sempre respeitando a categoria/tipo. Se não encontra uma correspondência segura, usa a lógica de marca, tamanho, formato (LN/LT), variante e multipack como fallback. Exemplo: “Antarctica Pilsen LT 269ml” → `ANTARCTICA PILSEN LT 269ML SH TRANSP C/15 NPAL` → `AP 269` → `AP269.png`. O card exibe `AP 269`, não o nome comercial longo.
- **O nome do arquivo no De/Para não contém a imagem em si.** Os arquivos de imagem reais precisam existir em alguma pasta acessível. O primeiro local recomendado é `images/img/` dentro deste projeto.
- O gerador procura `Caminho`, `Nome IMG` e o SKU abreviado (ex.: `AP 269` pode encontrar `AP269.png`), aceita subpastas, diferenças entre maiúsculas/minúsculas e caminhos com barras do Windows. As imagens encontradas são incorporadas ao HTML como `data:image/...;base64`, tornando a página independente dos arquivos depois de gerada.
- O gerador tenta também pastas `images/img` de repositórios irmãos dentro de `/workspaces`, mas isso depende de o repositório antigo existir no mesmo Codespace. A opção `--images-dir` é o modo mais confiável quando você sabe onde estão os arquivos; ela também está disponível em `atualizar.py`.
- A busca percorre pastas de imagens do projeto atual e dos repositórios presentes no Codespaces, sem depender de nomes exatos de diretório. Ao terminar, gera `data/processed/relatorio_imagens.txt`, com a ligação de cada alias ao caminho físico localizado e os aliases não encontrados.

Se as imagens ainda não aparecerem, encontre primeiro a pasta que contém os arquivos reais:

```bash
find /workspaces -type d -path '*/images/img' -print
```

Depois passe o caminho exato que apareceu no resultado. Por exemplo, se a pasta estiver no repositório antigo:

```bash
python scripts/atualizar.py --images-dir "/workspaces/-rvore-de-precos/images/img"
```

Como alternativa, copie os arquivos para a pasta local do novo projeto e regenere o HTML:

```bash
mkdir -p images/img
cp -r "/CAMINHO/REAL/DA/PASTA/images/img/." images/img/
python scripts/atualizar.py
```

Se um produto não casar com o De/Para, ele continua aparecendo, com um sinalizador para revisar o cadastro.

### Oferta especial fixa em MATCH

O tratamento acrescenta o card especial **BEATS - FF**, com as linhas **Leve 2 Beats / Ganhe 1 FF**, na segunda janela de datas de MATCH disponível no mês vigente. Para a composição visual, tenta localizar a imagem de Beats Sabores LT 269ml e Flying Fish LT 350ml usando o mesmo De/Para e a pasta de imagens. Essa oferta é adicionada mesmo que não exista como linha na base.

## Consolidação e diferenças

- Não são criadas abas por UF.
- Ofertas idênticas de produto, período e mecânica são consolidadas num card.
- Se a mecânica variar por divisão, as alternativas permanecem visíveis com o sinalizador **Condição varia por divisão**.
- Na aba `PERSONALIZADO`, as ofertas são agrupadas por varejista/rede; diferenças por GEO/rede recebem um sinalizador.
- Registros de abrangência `TOTAL` são eliminados apenas quando há uma linha individual de divisão com a mesma condição, evitando duplicação de um mesmo card.

## Mês exibido

Se existir uma aba `CONFIG` com `MesAtual` e `AnoAtual`, esses valores são respeitados. No arquivo atual, como não há `CONFIG`, o script infere o mês de campanha a partir das datas das linhas válidas das quatro abas ativas. Por isso, por exemplo, se as abas ativas contiverem promoções de novembro, ele seleciona novembro mesmo que o sistema ainda esteja em outubro. A aba `PERSONALIZADO (2)` não é usada para inferir o mês.

## Instalação e execução no Codespaces

Execute na raiz do repositório:

```bash
pip install -r requirements.txt
python scripts/atualizar.py
```

Saídas:

```text
data/processed/calendario_mercapromo.csv
data/processed/calendario_mercapromo_meta.json
data/processed/relatorio_tratamento.txt
site/calendario_mercapromo.html
```

Para indicar os arquivos manualmente:

```bash
python scripts/atualizar.py \
  --ofertas "data/unprocessed/Calendario Mercapromo.xlsx" \
  --de-para "cadastro/de_para_mercapromo.xlsx"
```

O relatório mostra o mês selecionado, quantas ofertas foram lidas por categoria, quais produtos não encontraram correspondência no De/Para e os avisos de dados.

## Privacidade

As planilhas podem conter campanhas internas. Não publique dados reais num repositório público sem autorização.


### Conferência de imagens

O tratamento gera `data/processed/mapeamento_produtos_imagens.csv`, que mostra a ligação entre o nome da promoção, o campo completo `Nome` do De/Para e o SKU abreviado, incluindo `Nome IMG` e `Caminho`. O gerador HTML gera `data/processed/relatorio_imagens.txt` com o vínculo de cada card ao arquivo físico efetivamente encontrado. Isso permite identificar se a falha é no relacionamento produto→SKU ou na localização do arquivo de imagem.
