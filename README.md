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
| `src/EmailEmergencial.gs` | Aviso por e-mail das **Transferências entre Plantas Emergenciais** — STOs firmadas depois do corte de D-1 10h (ver abaixo). |
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
| `Confirmacoes_Store` | Fonte da verdade das confirmações manuais **e** aba de rastreio das STOs: uma linha por STO viva, mais as confirmadas que saíram do export. |
| `Tokens_Link` | Tokens de acesso ao Web App (`?token=`). |
| `Log_Email_Emergencial` | Uma linha por STO enviada no aviso emergencial. Só cresce: é o que impede o reenvio e de onde sai o acumulado do dia. |

### As 5 colunas Smarthub

A `Pagina Transferência` lê a `ME2W` por `VLOOKUP` com **índice fixo**, então estas
colunas vêm **logo depois das colunas do SAP**, exatamente nesta ordem:

```
Confirmação Smarthub
Deliv Date - Confirmação Planejamento (SMART HUB)
Qtd - Confirmação Planejamento (SMART HUB)
Prioridade Smarthub
Causa de Desvio
```

Coluna a mais **antes** delas (no export do SAP, por exemplo) desloca as 5 e quebra todas as
fórmulas de uma vez. **Depois** delas vêm as 7 colunas do rastreio da STO (ver "Rastreio da
STO" abaixo), que não mexem na posição das 5. Se a página reclamar da largura nova,
`RASTREIO_NA_ME2W = false` no `Sync.gs` e a próxima sincronização da ME2W volta a gravá-la
sem elas.

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

### Rastreio da STO (out/2026)

O mesmo `Confirmacoes_Store` é a aba de histórico das STOs — e a `ME2W` leva uma **cópia**
das 7 colunas, depois das 5 Smarthub (`RASTREIO_NA_ME2W`). O store é a fonte: a `ME2W` é
reescrita a cada sincronização e recebe a cópia de lá; o clique de confirmação grava o
`Firmado em/por` nas duas na hora. Ele guarda só a **última**
ocorrência de cada coisa, nas 7 colunas do fim, com os mesmos nomes do histórico da ME2N do
Portal de Pedidos:

| Coluna | O que guarda | Quem grava |
| --- | --- | --- |
| `Entrou no sistema em` | Hora da sincronização em que a chave apareceu no export pela primeira vez. | Sync |
| `Firmado em` / `Firmado por` | O clique de confirmação do Planejamento. Fica até o próximo clique: o **Limpar não apaga** (ele fica em `Atualizado em/por`). | Portal |
| `Ult.Alteração Order Qty` / `Dt.Qnd.Alt.OrderQty` | A Order Quantity **anterior** e a hora da sincronização que viu a mudança. | Sync |
| `Ult.Alteração Deliver Date` / `Dt.Qnd.Alt.DD` | A Delivery Date **anterior** e a hora da sincronização que viu a mudança. | Sync |

Como a sincronização decide:

- Antes de sobrescrever a `ME2W`, ela lê a lista que está na aba (a **lista anterior**) e
  compara com o export novo pela chave. O `ME2W-Historico` não entra: cresceu demais para
  ser lido a cada passada.
- **Entrou**: a chave está no export e não estava nem na lista anterior nem no store.
- **Alteração**: Order Quantity ou Delivery Date diferentes entre a lista anterior e o
  export. Só conta com valor dos dois lados; data de um lado e texto do outro fica fora (aviso
  no registro de execução).
- A hora é a da **sincronização**, não a do SAP: o export não traz quando a ordem mudou, e
  duas alterações entre um export e outro aparecem como uma.
- A STO que sai do export e volta não é comparada (não estava na lista anterior); continua
  valendo o aviso de "Reapareceu" acima.

Para o store não crescer para sempre (a planilha tem teto de 10 milhões de células), a
STO **sem nada confirmado** que sai do export é tirada do store na mesma passada. As que
têm confirmação ficam como sempre ficaram (`AUSENTE`). A sem confirmação que voltar entra
de novo, com novo `Entrou no sistema em`.

Na primeira sincronização depois de publicar, as STOs que já estavam na `ME2W` ficam com
`Entrou no sistema em` vazio (não há como saber desde quando estão lá), e o `Firmado em`
das confirmadas sai do `Atualizado em` — pela regra acima, é o mesmo clique.

As **colunas** (P a V) são conferidas em toda passada do `sincronizarNovasBases`, mesmo
quando nenhum arquivo mudou — pelo cabeçalho, e não pela largura da aba (conteúdo solto à
direita não engana mais a conferência). `prepararRastreioSto()` faz o mesmo pelo editor. As **linhas** de cada STO viva só chegam quando a ME2W é
processada, ou seja, com um `STO-ME2W.xlsx` novo. Para ter tudo já, sem reimportar as outras
bases: `forcarRessincronizacaoMe2w()` e depois `sincronizarNovasBases()`. Com o mesmo arquivo
dos dois lados, nada é acusado como entrada nem como alteração.

O portal não lê o store inteiro a cada Salvar: lê só a coluna `Chave` para achar a linha
(`localizarLinhasStore`). As 7 colunas saem no fim do **Exportar** (`Transferencias_Export.xlsx`),
lidas só no clique (`getRastreioSto`); a abertura da tela não paga essa leitura.

### Proteções do sync

- A `ME2W` **aborta** (sem escrever) se o export vier sem as colunas-chave, sem linhas de
  dados, ou com queda de mais de 50% no volume — um export ruim apagaria confirmações de
  forma irreversível.
- A escrita da `ME2W` e os saves do portal disputam o **mesmo `LockService`**, para o sync
  não passar por cima de um clique em andamento (e vice-versa).
- As bases de análise não têm esse risco: cada uma falha por conta própria sem derrubar as
  demais, e o e-mail de alerta consolida as falhas.
- O sync só roda em dia útil, das 8h às 18h, e sai cedo se nenhum XLSX de origem mudou.

## Aviso de Transferências Emergenciais

A cada 15 min (`instalarAvisoEmergencial`) as STOs firmadas **depois do corte** que ainda não
foram avisadas vão num e-mail só para `armazembr14@ftatransportes.com.br` e
`bp_fernanda_frazao@colpal.com`, divididas pela `Plant` da ME2W (o destino). Sem nada novo,
nada é enviado.

**A regra**, por STO:

- `Status Planejamento` = **Firme**, o mesmo da tela (confirmação de pé e igual ao SAP em data e
  quantidade). Completa ou cancelada no SAP não é Firme, e por isso a STO já atendida não entra;
- `Firmado em` (do `Confirmacoes_Store`) a partir do **corte = 10:00 do dia anterior à entrega,
  em dias corridos** — a entrega de segunda tem corte no domingo às 10h. A firmada no próprio
  dia da entrega (D0) entra a qualquer hora, inclusive antes das 10h;
- a entrega não tinha passado no dia do firme (firmar STO vencida não é emergência);
- firme das últimas 24 h e posterior à instalação do aviso — ligar não despeja o histórico.

| Entrega | Firmado em | Corte | Resultado |
| --- | --- | --- | --- |
| Qui 08/Out | Qua 07/Out 14:32 | Qua 07/Out 10:00 | Emergencial |
| Qui 08/Out | Qua 07/Out 09:40 | Qua 07/Out 10:00 | Normal |
| Qua 07/Out | Qua 07/Out 08:20 | Ter 06/Out 10:00 | Emergencial (D0) |
| Sex 09/Out | Qua 07/Out 15:00 | Qui 08/Out 10:00 | Normal |
| Seg 12/Out | Sex 09/Out 16:00 | Dom 11/Out 10:00 | Normal (dias corridos) |

**Uma vez por firme.** O `Log_Email_Emergencial` guarda a chave e o `Firmado em` de cada linha
enviada, e o log é gravado **depois** do envio: falha no e-mail fica para o próximo gatilho,
nunca se perde. A STO firmada de novo com outra data ou quantidade volta como **Atualizada**,
com o valor anterior ao lado; o mesmo firme repetido sem mudança não volta. A STO avisada que
depois é limpa ou reprogramada para fora da regra **não** gera aviso.

**Os números do topo** contam só as linhas do próprio e-mail — cada um se refaz contando a
tabela abaixo dele:

| Número | Conta |
| --- | --- |
| STOs neste aviso | As linhas do e-mail (novas + atualizadas). |
| Ainda p/ hoje | As de entrega hoje, Firme e em aberto **no momento do envio**. As já atendidas e as avisadas antes não entram — não é o total de entregas do dia. |
| Para amanhã | As de entrega amanhã. |
| Paletes (aprox.) | Soma do `Pallet Order` das linhas; as sem conversão ficam fora e são contadas. |
| Acumulado de hoje | Linhas do log com envio hoje, este incluído — a soma dos "STOs neste aviso" do dia. |

Como nenhum firme sai em dois avisos, somar os avisos do dia dá o total do dia sem contar nada
duas vezes. Cada aviso leva o seu número no dia (`Aviso nº`): um número pulado é um e-mail que
não chegou. Qualquer aviso antigo se reconstrói filtrando o log por data e `Aviso nº`.

**Para ligar:** `testarAvisoEmergencial()` primeiro — monta o e-mail com os firmes emergenciais
dos últimos 7 dias e manda **só** para o `EMAIL_ALERTA`, sem gravar nada. Depois
`instalarAvisoEmergencial()`. O e-mail sai da conta dona do gatilho, que precisa autorizar o
escopo `script.scriptapp` (novo no manifesto) na primeira execução.

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
| `prepararRastreioSto()` | Logo depois de publicar: cria as colunas de rastreio (P a V) do `Confirmacoes_Store` e mostra no registro o que encontrou. Se ela não aparece no seletor, o `Sync.gs` do projeto é o antigo. |
| `forcarRessincronizacaoMe2w()` | Para reprocessar só a ME2W no próximo `sincronizarNovasBases`, sem arquivo novo (ex.: preencher o rastreio logo depois de publicar). |
| `testarAvisoEmergencial()` | Antes de ligar o aviso: manda o e-mail emergencial dos últimos 7 dias só para o `EMAIL_ALERTA`. Não grava log. |
| `instalarAvisoEmergencial()` | Liga o aviso emergencial (gatilho de 15 min). Só entram firmes feitos depois da primeira instalação. |
| `desligarAvisoEmergencial()` | Tira o gatilho. O log e o marco de instalação ficam. |
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
