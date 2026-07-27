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
| `src/appsscript.json` | Manifesto do projeto Apps Script (Drive v3 habilitado — o Sync converte XLSX). |

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

### Proteções do sync

- A `ME2W` **aborta** (sem escrever) se o export vier sem as colunas-chave, sem linhas de
  dados, ou com queda de mais de 50% no volume — um export ruim apagaria confirmações de
  forma irreversível.
- A escrita da `ME2W` e os saves do portal disputam o **mesmo `LockService`**, para o sync
  não passar por cima de um clique em andamento (e vice-versa).
- As bases de análise não têm esse risco: cada uma falha por conta própria sem derrubar as
  demais, e o e-mail de alerta consolida as falhas.
- O sync só roda em dia útil, das 8h às 18h, e sai cedo se nenhum XLSX de origem mudou.

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
| `sincronizarNovasBases()` | Por gatilho de tempo. Ignora execução fora da janela e quando nada mudou. |
| `getOrCreateToken(nome)` | Uma vez por usuário/planta, para gerar o link de acesso. |
