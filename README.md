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

## Por que o portal serve de cache

A `Pagina Transferência` é montada **inteira por fórmula** — `VLOOKUP` nas bases do SAP e
em duas planilhas externas. O `SpreadsheetApp.openById()` só devolve o controle depois que
o servidor termina de recalcular esse grafo: **123,5 s** no diagnóstico de 24/08/2026,
contra menos de 4 s para tudo o que vem depois (`getValues`, `getDisplayValues` e o
`getTransferData` inteiro somados).

Como cada requisição do Web App é uma execução nova, essa espera era paga **duas vezes** por
abertura: uma no `doGet`, só para validar o token, e outra no `getTransferData`. Mais de
quatro minutos de tela branca — o portal simplesmente não abria.

O cache (`CacheService`) tira as duas esperas da frente do usuário:

| Chave | Conteúdo | Validade |
| --- | --- | --- |
| `PORTAL_TOKENS_V1` | Lista de tokens válidos. | 30 min |
| `PORTAL_PAYLOAD_V1` | Payload pronto do portal, em fatias de 32 K (o limite é 100 KB por chave e o payload passa de 300 KB). | 6 h, servido só até 45 min de idade |

Quem paga o recalculo é o **`atualizarCachePortal()`**, chamado no fim do
`sincronizarNovasBases()` — é o único momento em que a planilha já está quente, porque as
fórmulas acabaram de ser recalculadas para gravar as bases. O gatilho `aquecerCachePortal()`, de meia em meia hora, é só rede de segurança para os ciclos em que
nenhum export mudou — o intervalo tem de caber nos 45 min de idade servível do payload.

As gravações do portal **não derrubam o cache**: `saveMultipleConfirmations()` e
`clearConfirmation()` aplicam a confirmação nas linhas tocadas do payload guardado. Derrubar
devolveria a espera de dois minutos para o próximo que abrisse; não mexer faria o próprio
autor da confirmação recarregar a página e não ver o que acabou de salvar.

> O cabeçalho do portal mostra **"Dados de dd/MM HH:mm"** — a hora em que o cache foi gerado,
> não a hora atual. É a idade real do que está na tela.

### Consequências que valem saber

- **Revogar um token não é imediato:** ele continua aceito por até 30 min. Rode
  `limparCachePortal()` para cortar na hora. Emitir token por `getOrCreateToken()` já
  invalida o cache sozinho.
- **Token desconhecido só relê a aba `Tokens_Link` uma vez por minuto.** O Web App é
  `ANYONE_ANONYMOUS`: sem esse teto, um link errado — ou um robô batendo no `/exec` —
  dispararia uma abertura de planilha de dois minutos *por requisição*, e a cota diária de
  execução acabaria sozinha. Na prática, um token recém-emitido pode levar até um minuto
  para ser aceito, a menos que tenha saído do próprio `getOrCreateToken()`.
- **A raiz do problema continua na planilha.** O cache esconde o recalculo, não o elimina.
  A firmação das colunas calculadas (`refirmarPaginaTransferencia`) nunca rodou — o
  diagnóstico registra *"nenhum registro de firmação — a página está inteira em fórmula"*,
  com 18 colunas marcadas. Rodá-la derruba o `openById` de dois minutos para poucos segundos
  e é o que torna o cache uma otimização, e não uma muleta.

## Operação

| Rotina | Quando rodar |
| --- | --- |
| `sincronizarNovasBases()` | Por gatilho de tempo. Ignora execução fora da janela e quando nada mudou. Atualiza o cache do portal ao terminar. |
| `instalarGatilhos()` | Uma vez, após publicar. Instala sync (15 min), histórico (1 h) e aquecimento do portal (30 min). |
| `aquecerCachePortal()` | Por gatilho, a cada 30 min. Pula domingo e o período fora de 5 h–21 h, para não gastar cota à toa. |
| `atualizarCachePortal()` | Na mão, para forçar a releitura da planilha agora. |
| `limparCachePortal()` | Depois de revogar um token, ou para descartar tudo o que está guardado. |
| `getOrCreateToken(nome)` | Uma vez por usuário/planta, para gerar o link de acesso. |
