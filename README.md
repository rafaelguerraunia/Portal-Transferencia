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
| `src/Calculo.gs` | As 19 colunas calculadas (`Y..AQ`) refeitas em JS com índice, gravadas como valor. É o caminho normal de "tirar as fórmulas". |
| `src/Firmar.gs` | Caminho genérico de firmação: repõe a fórmula, espera estabilizar e congela o resultado. Só para coluna que o `Calculo.gs` não sabe calcular. |
| `src/Formulas.gs` | Gera a aba `Mapa_Formulas` — o catálogo das fórmulas da página e o **arquivo** de cada uma, que é o que torna a firmação reversível. |
| `src/Debug.gs` | `diagnosticarPortal()` — percorre, na ordem em que o portal depende delas, cada peça que precisa estar de pé e diz onde para. |
| `src/STO-Frontend.html` | Portal do Planejamento (Bootstrap 5) servido por `HtmlService`. |
| `src/Tema_SmartHub.html` | Aparência e componentes comuns aos portais Smart Hub (KPIs, filtros, relógio, data da base). |
| `src/appsscript.json` | Manifesto do projeto Apps Script (Drive v3 habilitado — o Sync converte XLSX). |

> ⚠️ Dois nomes de arquivo **não são livres**, e errar qualquer um derruba o `doGet`
> inteiro (a tela nem chega a pedir dado):
> `STO-Backend.gs` carrega o portal por `HtmlService.createTemplateFromFile('STO-Frontend')`,
> e o `STO-Frontend.html` puxa o tema por `include('Tema_SmartHub')`.
> Os arquivos no editor do Apps Script têm de se chamar exatamente `STO-Frontend` e
> `Tema_SmartHub`. A etapa 1b do `diagnosticarPortal()` testa os dois.

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
| **Estoque & Compras** | O veredito de cobertura (+7D) e o selo **`Já possui STO, total X`**. O extrato inteiro abre no card, ao passar o mouse no código do material. |

`Prioridade` e `Causa de Desvio` são **calculadas** — o portal não pede que o usuário digite.

### O selo `Já possui STO`

A tela lista **linhas** (documento / item / schedule line), mas quem planeja decide por
**item** — material numa planta destino. O selo soma, por material + planta, a quantidade
das STOs daquele item que estão em **`Firme`**:

- só entra `Firme`, que no backend é confirmação existente **e** batendo com o SAP em data
  e quantidade. `Aguardando Confirmação` é quantidade que o Planejamento ainda não assumiu,
  `Solicitar ajuste` é quantidade em disputa com o SAP e `Revisão Urgente` é ordem completa
  ou cancelada — nenhuma delas está de pé;
- a unidade só aparece quando todas as STOs firmes do item falam a mesma.

É o que responde, sem filtrar a tabela pelo material, se uma requisição está pedindo algo
que já foi mandado buscar. Salvar uma linha refaz a conta: o índice é remontado a cada
render, então uma STO que acabou de virar `Firme` entra no total na hora.

### A coluna opcional `Status de Est.`

O botão **Habilitar Status de Est.**, abaixo do `Exportar`, liga uma coluna com o mesmo
extrato do card de Estoque & Compras — Consumo +7D, estoque na planta, dias de cobertura,
BR14 e SubContractor, pedidos em aberto, Insp. Qualidade — um dado abaixo do outro. Não
traz número novo e **não tem funil**: é leitura corrida, para comparar várias linhas sem
perseguir o mouse. Nasce desligada e o estado sobrevive ao F5 (`localStorage`), como o tema
e a densidade.

Com ela ligada, `Status Planejamento`, `Estoque +7D` e `Status do Fluxo` encolhem e o
conteúdo passa a quebrar em várias linhas (`Aguardando` / `Confirmação`) — é a largura que
paga a coluna nova, e não custa altura, porque a altura da linha já é ditada por ela. Com a
coluna desligada as três voltam ao tamanho de sempre, numa linha só.

## Operação

| Rotina | Quando rodar |
| --- | --- |
| `sincronizarNovasBases()` | Por gatilho de tempo. Ignora execução fora da janela e quando nada mudou. Termina firmando a página (passo 3). |
| `getOrCreateToken(nome)` | Uma vez por usuário/planta, para gerar o link de acesso. |
| `diagnosticarPortal()` | Quando a tela não abre ou fica em "Carregando dados...". Só lê; a última linha impressa é a resposta. |
| `mapearFormulasPaginaTransferencia()` | Para (re)gerar a aba `Mapa_Formulas` e revisar a coluna `Firmar?`. |
| `firmarColunasCalculadasTransferencia({todas:true})` | Para recalcular `Y..AQ` na mão e conferir o resultado. |
| `statusFirmacaoPaginaTransferencia()` | O que está firmado, desde quando, e o que o mapa diz que deveria estar. |

## Firmar a `Pagina Transferência` — e o que **não** fazer

A página é montada por fórmula de ponta a ponta, e o `getTransferData()` do portal
espera a planilha terminar de recalcular em toda abertura de tela. Firmar é trocar
fórmula por valor uma vez por sincronização, para a leitura do portal não pagar conta
nenhuma.

Há **dois caminhos**, e eles leem a mesma coluna `Firmar?` ao contrário um do outro:

| Caminho | O que "SIM" significa |
| --- | --- |
| `Calculo.gs` | refaz a conta em JS e grava **valor**. A fórmula sai. |
| `Firmar.gs` | **repõe** a fórmula, espera o recálculo e congela o resultado. A fórmula volta antes de sair. |

Para `AB` (*Dias Disponíveis em Estoque*) e `AJ` (*Dias Disponíveis na BR14*) o segundo
caminho não é mais caro — é **inviável**: são `MAP` dentro de `MAP` sobre `SEQUENCE(365)`,
365 `SUMIFS` por linha, cada um varrendo a `RESB` inteira. Repor essa fórmula deixa a
planilha em recálculo permanente, e aí o portal para de abrir.

Por isso as colunas que o `Calculo.gs` conhece ficam **fora do caminho genérico por
padrão**, venha a chamada de onde vier. Para desfazer mesmo, e assumindo a espera:

```js
restaurarFormulasPaginaTransferencia({ incluirCalculadas: true })
```

Duas coisas que **quebram a página** e não têm volta fácil:

- **Marcar `SIM` em qualquer coluna de `A` até `X`.** Elas não são colunas — são o
  derrame da `A2`, uma fórmula só que monta a página inteira. Gravar valor na âncora
  esvazia `B:X` e o portal abre em branco. O `Mapa_Formulas` já sai marcando essas
  células como bloqueadas; se acontecer, `restaurarDerramePaginaTransferencia()`.
- **Regerar o mapa e depois firmar sem conferir.** Coluna já firmada não tem mais
  fórmula para a varredura enxergar; o mapa agora **preserva** a fórmula arquivada dela
  (com `Nº de fórmulas` em 0), que é o que mantém a firmação reversível.
