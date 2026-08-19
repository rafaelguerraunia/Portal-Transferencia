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
| `src/appsscript.json` | Manifesto do projeto Apps Script (Drive v3 para converter XLSX, **Sheets v4 para a cópia rápida**). |

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

## Instalação no projeto Apps Script

O projeto depende de **dois serviços avançados**. Eles estão declarados no
`src/appsscript.json`, mas a declaração só vale se o **manifesto chegar ao projeto** —
colar apenas os `.gs` no editor deixa os serviços desligados:

| Serviço | Identificador | Versão | Para quê |
| --- | --- | --- | --- |
| Google Drive API | `Drive` | v3 | Converter os XLSX de origem |
| Google Sheets API | `Sheets` | v4 | Cópia rápida das bases de análise |

Habilitar pelo editor: **Serviços** (o `+` na barra lateral) → escolher a API →
conferir o identificador → **Adicionar**. Depois rode `sincronizarNovasBases()` uma vez
na mão para reautorizar.

> ⚠️ **Sem o `Sheets`, quatro das cinco bases nunca são sincronizadas.** `RESB`, `ME5A`,
> `ME2N` e `Stock Control` vão pela rota rápida, que é toda construída sobre a Sheets API;
> sem o serviço o sync as pula sem tocar nas abas, registra uma única falha explicando o
> passo que falta e manda o e-mail de alerta. Só a `ME2W` continua atualizando — e a
> `Pagina Transferência` passa a cruzar dado novo de STO com estoque e pedidos velhos.

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

A partir daí, a sincronização usa **um de dois caminhos** por coluna:

| Caminho | Quando | O que faz |
| --- | --- | --- |
| `Calculo.gs` | a coluna está em `PT_COLUNAS_CALCULADAS` | refaz a conta em JS, com índice, e grava |
| `Firmar.gs` | qualquer outra coluna marcada | repõe a fórmula, espera calcular, grava o resultado |

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

São **365 SUMIFS por linha**, cada um varrendo `RESB!A:D` inteiro. Com ~89 mil
linhas isso dá **32,6 milhões de SUMIFS por coluna** — 65 milhões somando as duas, a
cada recálculo. Nenhuma espera de estabilização resolve isso.

"Em quantos dias o estoque acaba" é uma série de consumo ordenada por data com soma
acumulada, percorrida uma vez. O índice de `RESB` é montado uma vez por execução e
serve as ~89 mil linhas: de `O(linhas × 365 × RESB)` para `O(linhas + RESB)`.

Os três `IMPORTRANGE` também somem — as planilhas de Pré-Agendamento e
Plano_Transporte passam a ser abertas por `openById`, direto. Mesma decisão do
`rpFontesStos_` do lado de Pedidos, pelo mesmo motivo: `IMPORTRANGE` tem cache
próprio e exibe o retrato de horas atrás sem um `#N/A` sequer para denunciar.

Entre duas sincronizações a página é estática: o portal lê e vai embora.

### Proteções da firmação

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

O mapeamento expôs cinco defeitos que já estavam em produção. Os quatro primeiros a
reimplementação em JS corrige por construção; o quinto é uma mudança deliberada.

| # | Onde | O quê |
| --- | --- | --- |
| 1 | `Y` Req Semana | `MAP(N2:N9374, I2:I89374, …)` — os dois arrays têm **tamanhos diferentes** (9.373 × 89.373 linhas). O `MAP` do Sheets exige dimensões iguais, então a coluna inteira volta em erro. O `STO-Backend` lê isso como `0`, e com `req = 0` o Status Estoque cai sempre em `✅ Suficiente (Físico)`. |
| 2 | `AF` `AG` `AH` | `MAP(N2:N9374, …)` enquanto as vizinhas usam `N2:N89374` — as três listas **param na linha 9.374**. Da 9.375 em diante, "Lista de Pedidos", "Fornecedor Pedido" e "Deliver Date Pedido" ficam vazias. |
| 3 | `AM` Planta Pedido | Tem filtro **próprio**, mais restrito que o de `AF`/`AG`/`AH`/`AI`, e por isso produz uma lista mais curta. O `STO-Backend` percorre as quatro em paralelo por índice (`arrPlanta[j]` com `arrDeliv[j]` e `arrQtd[j]`) — com comprimentos diferentes, o zip desalinha e casa pedido com a planta errada. |
| 4 | `AB` `AJ` | O `IFERROR` rotula `"Superior a 30 dias"`, mas o `SEQUENCE` procura **365**. O rótulo foi mantido como está: o `STO-Backend` testa por `.includes("Superior")` para virar prioridade `999`. |
| 5 | `AH` Deliver Date Pedido | O `TEXTJOIN` de uma data entrega o **número de série** (`"46253"`), e o `new Date("46253")` do `STO-Backend` devolve `Invalid Date` — as comparações de atraso (`hasAtrasado`, `hasMaior7d`) nunca disparavam. Passa a ser gravada como `yyyy-MM-dd`. |

Sobre o #3: em JS as cinco colunas passam a percorrer **a mesma lista, na mesma
ordem** — a regra de planta vira o *valor* de cada posição em vez de filtrar a lista.
Os comprimentos passam a bater e o zip do `STO-Backend` volta a alinhar. É uma
mudança de conteúdo da `AM`: onde a fórmula omitia a posição, agora há uma string
vazia entre os separadores.

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
| `getOrCreateToken(nome)` | Uma vez por usuário/planta, para gerar o link de acesso. |
