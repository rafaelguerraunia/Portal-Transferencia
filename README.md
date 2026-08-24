# Portal de Transferência de Material — Smart Hub

Aplicação Google Apps Script (Web App + Google Sheets) para o planejamento das
**transferências entre plantas** (STOs e requisições), com confirmação de data e
quantidade pelo Planejamento e visão consolidada de estoque, pedidos em aberto e
situação física da mercadoria.

- **Database:** [`1Beq9Gwq3l1o19r1t-yfX-_IVMTDBi8FPUsALGc58YtM`](https://docs.google.com/spreadsheets/d/1Beq9Gwq3l1o19r1t-yfX-_IVMTDBi8FPUsALGc58YtM)
- **Implantação:** `https://script.google.com/macros/s/AKfycbzuDRcraiJgfukV5m74_Q2YlU8MSWsIayx6O_-WS3gVHuuSAQVjqHyajbFWhlhpOU1G/exec`

## Estrutura do repositório

| Arquivo | Descrição |
| --- | --- |
| `src/STO-Backend.gs` | Backend do portal: leitura da `Pagina Transferência`, cálculo de status, gravação das confirmações e emissão de tokens de acesso. |
| `src/Sync.gs` | Sincronização das bases exportadas do SAP (XLSX no Drive → abas), com o **store de confirmações** que sobrevive ao ciclo destrutivo. |
| `src/STO-Frontend.html` | Portal do Planejamento (Bootstrap 5) servido por `HtmlService`. |
| `src/Tema_SmartHub.html` | Aparência (CSS/layout) comum aos portais Smart Hub, incluída pelo `STO-Frontend.html` via `include('Tema_SmartHub')`. |
| `src/Raster.gs` | Requisição à API da Raster (rastreamento de transporte) usada para alimentar o Status do Fluxo. |
| `src/Formulas.gs` | **Mapa de fórmulas** da `Pagina Transferência`: varre a aba, classifica cada coluna por custo e escreve a aba `Mapa_Formulas`. |
| `src/Calculo.gs` | **Refaz em JS**, com índice, as 19 colunas calculadas da página — o equivalente ao `firmarColunasCalculadasResumo` do Portal de Pedidos. |
| `src/Firmar.gs` | Caminho genérico: congela o resultado da própria fórmula, para coluna que ainda não tem conta em JS. Guarda também o botão de desfazer. |
| `src/appsscript.json` | Manifesto do projeto Apps Script (Drive v3 para converter os XLSX de origem). |

> ⚠️ O nome do arquivo `STO-Frontend.html` **não é livre**: `STO-Backend.gs` o carrega por
> `HtmlService.createTemplateFromFile('STO-Frontend')`. Renomear o arquivo quebra o `doGet`.
> O mesmo vale para `Tema_SmartHub.html`, incluído por nome via `include('Tema_SmartHub')`.

> ⚠️ **Segurança:** `src/Raster.gs` contém usuário e senha da API da Raster em texto puro,
> exatamente como veio da última exportação. Considere mover essas credenciais para
> `PropertiesService.getScriptProperties()` antes de tratar este repositório como compartilhável.

## Abas da planilha

| Aba | Papel |
| --- | --- |
| `Pagina Transferência` | Visão que o portal lê (montada por fórmulas a partir das demais). |
| `ME2W` | STOs exportadas do SAP. **Única base com dado insubstituível** — carrega as 5 colunas Smarthub. |
| `ME5A` / `ME2N` / `RESB` / `Stock Control BR14 BR10 BR12` | Bases de análise, reconstruíveis a cada sync. |
| `Confirmacoes_Store` | Fonte da verdade das confirmações manuais. Nunca remove linhas. |
| `Tokens_Link` | Tokens de acesso ao Web App (`?token=`). |
| `Mapa_Formulas` | Inventário das fórmulas da `Pagina Transferência` e **configuração do que é firmado**. Gerada por `mapearFormulasPaginaTransferencia()`. |

### As 5 colunas Smarthub

A `Pagina Transferência` lê a `ME2W` por `VLOOKUP` com **índice fixo**, então estas
colunas precisam ser as **últimas 5**, exatamente nesta ordem:

```
Confirmação Smarthub
Deliv Date - Confirmação Planejamento (SMART HUB)
Qtd - Confirmação Planejamento (SMART HUB)
Prioridade Smarthub
Causa de Desvio
```

Mudar a largura da `ME2W` quebra todas as fórmulas de uma vez.

## Por que existe o store de confirmações

O sync reescreve a `ME2W` inteira (`clearContents` + `setValues`) a cada execução.
Sem um lugar fora desse ciclo, toda confirmação manual do Planejamento seria perdida
quando a ordem sumisse do export do SAP.

O `Confirmacoes_Store` guarda, por chave `Purchasing Document | Item | Schedule Line`:

- as 5 colunas manuais;
- a **fotografia** da `Delivery Date` e da `Order Quantity` no ato da confirmação;
- o status da ordem: `ATIVA`, `AUSENTE` (sumiu do export) ou `REAPARECEU`.

Quando uma ordem ausente reaparece **com data ou quantidade diferentes**, a linha volta
marcada com `⚠️ Reapareceu no SAP com … diferente — revalidar` na Causa de Desvio, em vez
de ser restaurada em silêncio.

### Proteções do sync

- A `ME2W` **aborta** (sem escrever) se o export vier sem as colunas-chave, sem linhas de
  dados, ou com queda de mais de 50% no volume — um export ruim apagaria confirmações de
  forma irreversível.
- A escrita da `ME2W` e os saves do portal disputam o **mesmo `LockService`**, para o sync
  não passar por cima de um clique em andamento (e vice-versa).
- As bases de análise não têm esse risco: cada uma falha por conta própria sem derrubar as
  demais, e o e-mail de alerta consolida as falhas.
- O sync roda de segunda a sábado, da 1h em diante, e **pula cada base cujo XLSX não mudou** (carimbo por arquivo, não um carimbo único dos cinco).

### Como cada base é copiada

As duas rotas seguem o mesmo caminho do `Sincronizacao_Backend` do Portal de Pedidos —
**converte o XLSX, lê o temporário de uma vez, grava na aba de uma vez**:

```
Drive.Files.create (XLSX → Sheets)  →  getValues()  →  clearContents() + setValues()
```

São **duas chamadas de planilha por base**. A rota crítica (`ME2W`) só acrescenta
validação, `LockService` e restauro do store por cima disso.

Antes, as quatro bases de análise iam por uma cópia construída sobre a **Sheets API**:
metadados da planilha alvo, metadados do temporário, um `Values.get` e um `Values.update`
por bloco de 400 mil células, `Values.clear`, mais um `get` de formato numérico e um
`batchUpdate` de `repeatCell` por coluna — 8 a 12 idas à API por base. O laço de blocos
ainda era guiado pelo `rowCount` do **grid** do temporário, e um XLSX convertido vem com
folga de linhas vazias no fim: lia faixa sem dado até um bloco curto denunciar o fim.

O que aquela rota comprava era transportar data como **número de série**, sem depender de
locale — e por isso precisava replicar o formato numérico coluna a coluna, senão a data
aparecia como `45123` na tela. Lendo e escrevendo por `SpreadsheetApp`, a data vai e volta
como `Date` de verdade: some o serial, some a replicação de formato e some a dependência
do serviço avançado.

> As abas ainda não ressincronizadas desde a troca continuam com número de série nas
> colunas de data. O `ptDia_` do `Calculo.gs` aceita as duas formas, então não há
> necessidade de `forcarRessincronizacao()` — a próxima passada de cada base normaliza.

**A grade só cresce, nunca encolhe.** `clearContents()` limpa o conteúdo sem mexer na
grade, e `garantirGradeAba` só acrescenta linha/coluna quando o export não cabe: apagar
linha ou coluna de uma aba que a `Pagina Transferência` referencia por intervalo produz
`#REF!` irreversível.

## Instalação no projeto Apps Script

O projeto depende de **um serviço avançado**. Ele está declarado no
`src/appsscript.json`, mas a declaração só vale se o **manifesto chegar ao projeto** —
colar apenas os `.gs` no editor deixa o serviço desligado:

| Serviço | Identificador | Versão | Para quê |
| --- | --- | --- | --- |
| Google Drive API | `Drive` | v3 | Converter os XLSX de origem |

Habilitar pelo editor: **Serviços** (o `+` na barra lateral) → escolher a API →
conferir o identificador → **Adicionar**. Depois rode `sincronizarNovasBases()` uma vez
na mão para reautorizar.

> A **Google Sheets API** (`Sheets`, v4) não é mais necessária: era ela que sustentava a
> cópia em blocos das bases de análise. Sem ela, `RESB`, `ME5A`, `ME2N` e `Stock Control`
> nunca sincronizavam — só a `ME2W` atualizava, e a `Pagina Transferência` cruzava STO
> nova com estoque e pedidos velhos. Esse modo de falha deixou de existir.

## As fórmulas da `Pagina Transferência`

A página é montada por fórmula de ponta a ponta. Cada recálculo da planilha reavalia
`VLOOKUP` na `ME2W`, cruzamento com `RESB` / `ME2N` / `Stock` e o que mais estiver ali —
e o recálculo **dispara sozinho**: o sync reescreve as bases, `TODAY()` vira à
meia-noite, um `IMPORTRANGE` se atualiza por conta própria. Enquanto isso corre, o
`getTransferData()` espera: é a planilha em recálculo permanente que faz a tela demorar
a abrir.

É o mesmo problema que a `Resumo P.O` tinha no Portal de Pedidos, e a saída é a mesma —
calcular uma vez por sincronização e gravar valor. A diferença é o caminho:

| | Portal de Pedidos | Portal de Transferência |
| --- | --- | --- |
| Rotina | `firmarColunasCalculadasResumo()` | `refirmarPaginaTransferencia()` |
| Quem faz a conta | reimplementada em JS, com índices em memória | a própria planilha, uma vez por ciclo |
| Precisa conhecer a fórmula? | sim, coluna a coluna | não |

### 1. Levantar o que existe

```
mapearFormulasPaginaTransferencia()
```

Varre a aba inteira e escreve a aba **`Mapa_Formulas`**, uma linha por coluna de
fórmula, **ordenada por custo estimado** — o topo da lista é o que mais paga para virar
valor. Também grava um `.txt` no Drive com todos os padrões distintos (a aba mostra um
exemplo por coluna) e repete tudo no log.

| Coluna do mapa | O que diz |
| --- | --- |
| `Tipo` | `array` (uma âncora que derrama) ou `preenchida` (padrão repetido linha a linha) — decide como a fórmula volta no restauro |
| `Uniforme?` | `não` = há mais de um padrão na mesma coluna, alguém editou célula avulsa |
| `Volátil?` | `sim` = recalcula sozinha (`TODAY`, `NOW`, `INDIRECT`, `IMPORTRANGE`…) |
| `Peso estimado` | células de fórmula × custo das funções × 2 se volátil |
| `Firmar?` | **o único campo editável à mão** |

### 2. Escolher e firmar

Marque `SIM` na coluna `Firmar?` das colunas que quer congelar. **O default de toda
coluna é `NÃO`** — nada é firmado por conta própria, e uma sincronização nunca apaga
fórmula que ninguém autorizou. A escolha sobrevive a regerar o mapa.

**As colunas de `A` a `X` são a exceção**: elas não aceitam `SIM` e o mapa já as
escreve bloqueadas. São o derrame da `A2`, que monta a página — firmar ali esvazia
`B:X`. Veja *Proteções da firmação*.

A partir daí, a sincronização usa **um de dois caminhos** por coluna:

| Caminho | Quando | O que faz |
| --- | --- | --- |
| `Calculo.gs` | a coluna está em `PT_COLUNAS_CALCULADAS` | refaz a conta em JS, com índice, e grava |
| `Firmar.gs` | qualquer outra coluna marcada | repõe a fórmula, espera calcular, grava o resultado |

Os dois **nunca tocam a mesma coluna**: o Sync passa ao caminho genérico a lista do
que o `Calculo.gs` conhece (`ptLetrasCalculadas_()`), e ele pula essas. Sem isso o
genérico repunha a fórmula das colunas que o `Calculo.gs` tinha acabado de gravar —
`AB` e `AJ` voltavam aos 365 `SUMIFS` por linha, a espera pelo recálculo estourava o
tempo do gatilho e a página terminava o ciclo em fórmula.

### O fluxo do gatilho

Cada passo só paga o custo se o anterior deu o que fazer:

```
1. VERIFICAR      carimbo do .xlsx  ≠  guardado em SYNC_<aba>?
                          │ não → base pulada, nem converte
                          ↓ sim
2. SINCRONIZAR    converte → lê o temporário → grava na aba
                  (ME2W pela rota crítica; as outras quatro pela direta)
                          ↓  escreveu ≥ 1 base?
3. FIRMAR         base nova OU o dia virou?   ── não → página já está certa, pula
                          ↓ sim
                  Calculo.gs   → Y..AQ, conta em JS, grava valor
                  Firmar.gs    → só as colunas marcadas que o Calculo.gs não conhece
```

O passo 3 nunca vem antes do 2: as contas da página precisam enxergar a base nova
antes de virar valor. E o gate do passo 3 (`fmPrecisaRefirmar_`) é o que separa
firmar uma ou duas vezes por dia de firmar a cada 15 minutos sem motivo — só duas
coisas envelhecem a página, base nova escrita e a virada do dia (`TODAY()`).

Quanto ao registro do que está firmado (`PT_COLUNAS_FIRMADAS`): os dois caminhos
escrevem nele, então ele é **união**, nunca substituição, e uma restauração parcial
remove só as colunas que voltaram a ser fórmula. Substituir fazia um estouro de
tempo no caminho genérico apagar o registro inteiro — a página ficava firmada e
sem registro, e o gate lia "nunca firmou" e mandava refirmar a cada gatilho.

### 2b. As 19 colunas que já têm conta em JS

`Y`…`AQ`: Req Semana, os quatro saldos de estoque, os dois "Dias Disponíveis", as
listas de pedidos em aberto, Pallet Order, Pré Agendado?, Separado? e Status
Transporte.

**Por que não bastava congelar o resultado da fórmula.** `AB` e `AJ` não são caras,
são inviáveis — têm um `MAP` dentro de outro `MAP`:

```
MAP( <linhas>, LAMBDA(w, n, i,
  MATCH(TRUE, MAP(SEQUENCE(365,1,0), LAMBDA(dias,
    SUMIFS(RESB!D:D, RESB!A:A,n, RESB!B:B,i,
           RESB!C:C,">="&TODAY(), RESB!C:C,"<="&TODAY()+dias) >= w )), 0) - 1 ))
```

São **365 SUMIFS por linha**, cada um varrendo `RESB!A:D` inteiro. Com as 265 linhas
do export real isso dá **193.450 SUMIFS por recálculo** só nessas duas colunas — e o
`MAP` ainda itera 89.373 vezes por coluna, porque a faixa de entrada é `Z2:Z89374`
(o `IF(n="","")` corta antes do `SUMIFS` nas ~89.100 linhas vazias).

As duas respondem por **97,9% do custo de recálculo da página**. Firmar só `AB` e `AJ`
já entrega quase todo o ganho — é o primeiro passo recomendado.

"Em quantos dias o estoque acaba" é uma série de consumo ordenada por data com soma
acumulada, percorrida uma vez. O índice de `RESB` é montado uma vez por execução e
serve as ~89 mil linhas: de `O(linhas × 365 × RESB)` para `O(linhas + RESB)`.

Os três `IMPORTRANGE` também somem — as planilhas de Pré-Agendamento e
Plano_Transporte passam a ser abertas por `openById`, direto. Mesma decisão do
`rpFontesStos_` do lado de Pedidos, pelo mesmo motivo: `IMPORTRANGE` tem cache
próprio e exibe o retrato de horas atrás sem um `#N/A` sequer para denunciar.

Entre duas sincronizações a página é estática: o portal lê e vai embora.

### Proteções da firmação

- **A:X nunca é firmada.** A `A2` não é uma coluna de fórmula como as outras: o
  `HSTACK` dela derrama na **horizontal**, e tudo de `A` até `X` é o resultado de uma
  fórmula só. Firmar a coluna `A` grava valor na âncora, mata a fórmula e leva junto o
  derrame de `B` até `X` — sem `Material`, `Planta` e `Documento/Item/Schedule Line`,
  as colunas `Y`…`AQ` passam a ler célula vazia e a página inteira para. O
  `fmLerCatalogo_` recusa qualquer coluna dentro de `A:X` mesmo marcada `SIM`
  (`FM_ULTIMA_COLUNA_DERRAME`), e o `Mapa_Formulas` já as escreve como
  `NÃO (derrame da A2 — quebra a página)`. Não há o que ganhar do outro lado: `A:X`
  é dado cru copiado da `ME5A` e da `ME2W`; o custo do recálculo mora em `Y`…`AQ`.
- **Nada é firmado no escuro.** A assinatura das colunas precisa vir igual em duas
  leituras seguidas; se ainda estiver mudando depois do tempo limite, a fórmula fica.
- **Erro aborta.** Qualquer `#REF!` / `#N/A` / `Loading...` nas colunas marcadas cancela
  a passada inteira — firmar cedo demais congela dado pela metade até o próximo ciclo,
  que é bem pior do que a fórmula lenta que ao menos se corrige sozinha.
- **Só as colunas marcadas.** A gravação vai em corridas de colunas adjacentes, nunca
  num retângulo do mínimo ao máximo: as colunas não marcadas no meio continuam fórmula
  (ou preenchimento manual de operação).
- **Mesmo `LockService` do sync e dos saves do portal**, para não gravar a página por
  cima de uma confirmação que acabou de ser salva.
- **Dá para desfazer.** A fórmula fica arquivada na própria `Mapa_Formulas`;
  `restaurarFormulasPaginaTransferencia()` devolve todas de uma vez, sem depender de
  backup.
- **Só refirma quando precisa**: base nova escrita, ou o dia virou (`TODAY()` passou a
  responder outra coisa). Fora disso a página firmada já está certa, e refirmar pagaria
  um recálculo inteiro para chegar no mesmo resultado.

> **Efeito colateral aceito**, o mesmo da `Resumo P.O`: entre duas sincronizações, uma
> linha que apareça na página fica com as colunas firmadas vazias até a próxima passada.
> A janela é o intervalo do gatilho.

## Divergências encontradas nas fórmulas

Levantadas pelo mapa e **conferidas contra um export real da página** (265 linhas,
19/08/2026). Duas previsões feitas só pela leitura da fórmula não se confirmaram, e
estão registradas como não-defeitos para ninguém "corrigir" o que funciona.

### Confirmadas

| # | Onde | O quê | Medido |
| --- | --- | --- | --- |
| 1 | `AO` `AP` `AQ` | A fórmula está em **uma célula só** e nunca foi arrastada. A coluna responde pela linha 2 e fica vazia no resto. `Status Transporte` não tem valor em nenhuma linha, então `statusFluxo1` do `STO-Backend` é sempre `""` e o Status do Fluxo nunca mostra Pré-Agendado / Separado / transporte. | 1 de 265 linhas com valor em `AO` e `AP`; 0 de 265 em `AQ` |
| 2 | `AM` Planta Pedido | Tem filtro **próprio**, mais restrito que o de `AF`/`AG`/`AH`/`AI`, e produz uma lista mais curta. O `STO-Backend` percorre as quatro em paralelo por índice (`arrPlanta[j]` com `arrDeliv[j]` e `arrQtd[j]`) — com comprimentos diferentes o zip desalinha e casa pedido com a planta errada. | 27 de 162 linhas com pedido em aberto (**17%**) |
| 3 | `AB` `AJ` | O `IFERROR` rotula `"Superior a 30 dias"`, mas o `SEQUENCE` procura **365**. Rótulo mantido: o `STO-Backend` testa por `.includes("Superior")` para virar prioridade `999`. | 37 linhas com o rótulo em `AJ` |

### Latentes — não doem hoje, doem quando a página crescer

| Onde | O quê |
| --- | --- |
| `Y` Req Semana | `MAP(N2:N9374, I2:I89374, …)` — os dois arrays têm tamanhos diferentes. Com 265 linhas isso passa sem erro; a divergência de 80 mil linhas entre as duas faixas é uma bomba armada para o dia em que a página passar de 9.373 linhas. |
| `AF` `AG` `AH` | Usam `N2:N9374` enquanto as vizinhas usam `N2:N89374`. Da linha 9.375 em diante as três listas ficariam vazias. |

### Previsões que a base real desmentiu

| O que eu previ | O que o dado mostrou |
| --- | --- |
| `Y` estaria em `#N/A` pelo `MAP` de tamanhos diferentes, e o Status Estoque cairia sempre em `✅ Suficiente` | **Não.** Zero erros em 265 linhas — o Sheets tolera a diferença nesse tamanho. Continua sendo dívida (acima), não defeito ativo. |
| `AH` sairia como número de série (`"46253"`), e o `new Date()` do `STO-Backend` daria `Invalid Date` | **Não.** Sai como `8/25/2026`, que o `new Date()` parseia certo. |

> ⚠️ A segunda quase virou um defeito *introduzido por esta mudança*: a primeira
> versão do `Calculo.gs` gravava `AH` em ISO (`2026-08-25`). Em `America/Sao_Paulo`,
> `new Date("2026-08-25")` cai no **dia 24 às 21h**, e toda entrega do próprio dia
> passaria a contar como atrasada. O `ptDataTexto_` escreve `M/D/YYYY`, igual ao que
> a fórmula já produz.

## Status apresentados no portal

| Coluna | O que responde |
| --- | --- |
| **Status Planejamento** | Situação da linha: `Firme`, `Solicitar ajuste`, `Aguardando Confirmação`, `Pendente Criação STO`, `Revisão Urgente`. |
| **Status Estoque (+7D)** | Cruza estoque físico da planta destino, estoque BR14 e pedidos em aberto que chegam na janela de 7 dias. |
| **Status do Fluxo** | Situação física e documental: combina Pré-Agendado / Separado / Status Transporte com as quantidades `Issued` / `Delivered` / `Received`. |

`Prioridade` e `Causa de Desvio` são **calculadas** — o portal não pede que o usuário digite.

## Operação

| Rotina | Quando rodar |
| --- | --- |
| `sincronizarNovasBases()` | Por gatilho de tempo (~15 min). Ignora execução fora da janela, pula cada base cujo XLSX não mudou e firma a página no fim. |
| `sincronizarHistoricos()` | Por gatilho próprio (~1 h). Acumula `ME2W-Historico` e `RESB-Historico` a partir das abas já sincronizadas. |
| `instalarGatilhos()` | Uma vez, para criar os dois gatilhos acima (recria do zero, não duplica). |
| `forcarRessincronizacao()` | Depois de mexer manualmente numa aba de base: limpa os carimbos e o próximo gatilho reimporta tudo. |
| `mapearFormulasPaginaTransferencia()` | Sempre que as fórmulas da página mudarem — regera a `Mapa_Formulas` preservando a coluna `Firmar?`. |
| `statusFirmacaoPaginaTransferencia()` | Diagnóstico: o que está marcado, o que está firmado e desde quando. |
| `firmarColunasCalculadasTransferencia({todas:true})` | Roda o cálculo em JS à mão, ignorando o `Firmar?` — para conferir o resultado antes de autorizar. |
| `restaurarFormulasPaginaTransferencia()` | Botão de desfazer: devolve a fórmula a todas as colunas marcadas. |
| `restaurarDerramePaginaTransferencia()` | Socorro, uma vez só: repõe a fórmula da `A2` e limpa o retângulo `A:X` depois de uma firmação que congelou o derrame e deixou `B:X` vazias. |
| `getOrCreateToken(nome)` | Uma vez por usuário/planta, para gerar o link de acesso. |
