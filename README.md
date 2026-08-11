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
| `src/STOBackend.gs` | Backend do portal: leitura da `Pagina Transferência`, cálculo de status, gravação das confirmações e emissão de tokens de acesso. |
| `src/Sync.gs` | Sincronização das bases exportadas do SAP (XLSX no Drive → abas), com o **store de confirmações** que sobrevive ao ciclo destrutivo. |
| `src/Sto-Frontend.html` | Portal do Planejamento (Bootstrap 5) servido por `HtmlService`. |
| `src/appsscript.json` | Manifesto do projeto Apps Script (Drive v3 para converter XLSX, Sheets v4 para a cópia rápida). |

> ⚠️ O nome do arquivo `Sto-Frontend.html` **não é livre**: `STOBackend.gs` o carrega por
> `HtmlService.createTemplateFromFile('Sto-Frontend')`. Renomear o arquivo quebra o `doGet`.

## Abas da planilha

| Aba | Papel |
| --- | --- |
| `Pagina Transferência` | Visão que o portal lê (montada por fórmulas a partir das demais). |
| `ME2W` | STOs exportadas do SAP. **Única base com dado insubstituível** — carrega as 5 colunas Smarthub. |
| `ME5A` / `ME2N` / `RESB` / `Stock Control BR14 BR10 BR12` | Bases de análise, reconstruíveis a cada sync. |
| `Confirmacoes_Store` | Fonte da verdade das confirmações manuais. Nunca remove linhas. |
| `Tokens_Link` | Tokens de acesso ao Web App (`?token=`). |

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

## As duas rotas do sync

O que cada base tem a perder decide por onde ela passa. O catálogo fica em `BASES`,
no topo do `Sync.gs`.

| Rota | Bases | O que acontece |
| --- | --- | --- |
| **`critica`** | `ME2W` | Valida o export, pega o `LockService`, restaura o store, escreve e regrava o store. |
| **`rapida`** | `RESB`, `ME5A`, `ME2N`, `Stock Control BR14 BR10 BR12` | Cópia pura via Sheets API, em blocos: lê, limpa, escreve. Não percorre linha a linha nem monta chave. |

As bases da rota rápida são insumos reconstruídos a cada sync — não têm dado manual,
o portal não escreve nelas e um export ruim custa só o próximo gatilho. **Não há o que
verificar nelas**, então não se verifica.

### Por que a cópia rápida transporta data como número de série

O caminho rápido lê com `UNFORMATTED_VALUE` + `SERIAL_NUMBER` e escreve com `RAW`:
serial não depende de locale para ser lido nem escrito, então nenhuma data é
reinterpretada no meio do caminho. O preço é que a coluna apareceria como `45123` —
por isso `replicarFormatoNumerico` copia o formato numérico da origem junto.

A grade da aba de destino **só cresce, nunca encolhe**: apagar linha ou coluna de uma
aba que a `Pagina Transferência` referencia produz `#REF!` irreversível.

### Proteções do sync

- A `ME2W` **aborta** (sem escrever) se o export vier sem as colunas-chave, sem linhas de
  dados, ou com queda de mais de 50% no volume — um export ruim apagaria confirmações de
  forma irreversível.
- A escrita da `ME2W` e os saves do portal disputam o **mesmo `LockService`**, para o sync
  não passar por cima de um clique em andamento (e vice-versa).
- As bases de análise não têm esse risco: cada uma falha por conta própria sem derrubar as
  demais, e o e-mail de alerta consolida as falhas.
- Na rota rápida, a aba só é limpa **depois** que o primeiro bloco de dados chega: um
  export vazio aborta a base sem zerar o que a `Pagina Transferência` lê.
- O sync roda de segunda a sábado, da 1h em diante, e pula cada base cujo XLSX não mudou.

### Como o sync cabe nos 6 minutos

O gatilho do Apps Script morre em 6 min. Três decisões mantêm a sincronização dentro disso:

1. **Carimbo por arquivo** (`SYNC_<aba>` no `ScriptProperties`). Um carimbo único dos cinco
   fazia um export novo de `ME2W` arrastar `RESB`, `ME2N` e `Stock` inteiros junto, sem
   nada ter mudado neles.
2. **Uma base por vez**: converte, escreve, apaga o temporário e só então carimba. Converter
   os cinco XLSX antes de escrever gastava o orçamento inteiro antes da primeira linha
   entrar na planilha — e um estouro no meio jogava fora as cinco conversões.
3. **Orçamento de 4min30**: o sync para por conta própria antes do corte. Como cada base
   concluída já está carimbada, o próximo gatilho continua de onde parou em vez de refazer
   tudo do zero e estourar de novo.

O acúmulo nas abas `-Historico` saiu do sync justamente por isso: o custo do dedup cresce
com o histórico, e crescia dentro do mesmo orçamento da cópia das bases.

> ⚠️ A linha nova do histórico é gravada como `[timestamp] + linha da origem`, na ordem de
> colunas do export. Se o SAP mudar a ordem das colunas de uma base, as linhas antigas do
> histórico ficam desalinhadas em relação às novas (o dedup continua correto — ele resolve o
> índice pelo cabeçalho de cada lado). Comportamento herdado, mantido como estava.

## Status apresentados no portal

| Coluna | O que responde |
| --- | --- |
| **Status Planejamento** | Situação da linha: `Firme`, `Solicitar ajuste`, `Aguardando Confirmação`, `Pendente Criação STO`, `Revisão Urgente`. |
| **Status Estoque (+7D)** | Cruza estoque físico da planta destino, estoque BR14 e pedidos em aberto que chegam na janela de 7 dias. |
| **Status do Fluxo** | Situação física e documental: combina Pré-Agendado / Separado / Status Transporte com as quantidades `Issued` / `Delivered` / `Received`. |

`Prioridade` e `Causa de Desvio` são **calculadas** — o portal não pede que o usuário digite.

## Instalação no projeto Apps Script

O projeto depende de **dois serviços avançados**. Eles estão declarados no
`src/appsscript.json`, mas a declaração só vale se o **manifesto chegar ao projeto** —
colar apenas o `Sync.gs` no editor deixa os serviços desligados:

| Serviço | Identificador | Versão | Para quê |
| --- | --- | --- | --- |
| Google Drive API | `Drive` | v3 | Converter os XLSX de origem |
| Google Sheets API | `Sheets` | v4 | Cópia rápida das bases sem verificação |

Habilitar pelo editor: **Serviços** (o `+` na barra lateral) → escolher a API →
conferir o identificador → **Adicionar**. Depois rode `sincronizarNovasBases()` uma vez
na mão para reautorizar.

Sem o `Sheets`, o sync não quebra as abas: ele pula as bases de cópia rápida sem tocar
nelas, registra uma única falha explicando o passo que falta e manda o e-mail de alerta.
Como o carimbo só é gravado depois do sucesso, as bases puladas voltam no próximo gatilho.

## Operação

| Rotina | Quando rodar |
| --- | --- |
| `sincronizarNovasBases()` | Por gatilho de tempo (~15 min). Ignora execução fora da janela e pula cada base cujo XLSX não mudou. |
| `sincronizarHistoricos()` | Por gatilho de tempo próprio (~1 h). Acumula `ME2W-Historico` e `RESB-Historico` a partir das abas já sincronizadas. Só trabalha se o sync marcou pendência. |
| `instalarGatilhos()` | Uma vez, para criar os dois gatilhos acima (recria do zero, não duplica). |
| `forcarRessincronizacao()` | Depois de mexer manualmente numa aba de base: limpa os carimbos e o próximo gatilho reimporta tudo. |
| `getOrCreateToken(nome)` | Uma vez por usuário/planta, para gerar o link de acesso. |

> O `sincronizarHistoricos()` lê a aba já sincronizada na planilha alvo — **não** reconverte
> o XLSX. Se ele atrasar, o histórico fica para trás, mas o sync das bases não para.
