# Controle de Separação

Web App em Google Apps Script que usa uma planilha do Google Sheets como banco de dados e consulta o Amber (Aton ERP) para trazer os itens de cada lote de separação.

> **O Amber é usado só para consulta.** A API do Aton faz consultas via `POST`, então a proteção é uma **lista fixa de dois endpoints de consulta** (`ENDPOINTS_PERMITIDOS` no `Code.gs`): `pedidosvenda/consulta` e `produtos/listagemid`. Os endpoints que alteram dados (os `PUT`: atualizar pedidos, estoques, baixas etc.) não são chamados em lugar nenhum, e `requisicaoAmber_()` recusa qualquer endereço fora da lista. O sistema não cria, altera nem apaga nada no Amber.

## Arquivos

| Arquivo | Para quê |
|---|---|
| `Code.gs` | Back-end: setup, regras, integração com o Amber, painel |
| `Index.html` | Front-end responsivo (Lançar / Painel / Produtos) |
| `appsscript.json` | Manifesto (fuso de São Paulo, escopos, publicação) |
| `.clasp.json.example` | Modelo para o clasp |

## Passo a passo

### 1. Criar a planilha e o projeto (com clasp)

Requer Node.js 18+.

```bash
npm install
npx clasp login                       # abre o navegador para autorizar
npx clasp create --type sheets --title "Controle de Separação" --rootDir .
```

Isso cria a planilha e o script vinculado a ela, e gera o `.clasp.json`.
Se o clasp reclamar que a *Apps Script API* está desligada, ative em https://script.google.com/home/usersettings e rode de novo.
O `clasp create` pode sobrescrever o `appsscript.json`. Se isso acontecer, restaure o do repositório (`git checkout appsscript.json`).

Sem clasp: abra uma planilha nova, **Extensões → Apps Script**, e cole o conteúdo de `Code.gs`, crie um arquivo HTML chamado `Index` com o conteúdo de `Index.html` e, em **Configurações do projeto**, marque "Mostrar arquivo de manifesto appsscript.json" e cole o manifesto.

### 2. Enviar o código

```bash
npx clasp push
npx clasp open                        # abre o editor do Apps Script
```

### 3. Criar as abas (uma vez)

No editor, selecione a função **`setup`** e clique em **Executar**. O Google vai pedir autorização (veja o passo 6).
O `setup` cria as abas **Registros**, **Itens_Lote**, **Produtos** e **Colaboradores**, com cabeçalhos, validações (dificuldade só aceita 1, 2 ou 3; Ativo só aceita SIM/NÃO) e colunas de data/hora/SKU em formato texto.

Depois, na aba **Colaboradores**, cadastre os separadores (coluna A = nome, coluna B = `SIM`).
Apague a linha de exemplo.

### 4. Testar sem a API (MODO_TESTE)

`CONFIG.MODO_TESTE` já vem `true`: o sistema gera itens fictícios para qualquer lote. Faça a implantação (passo 7), lance alguns lotes, classifique produtos e veja o painel. O app mostra a etiqueta "MODO TESTE" no topo.

### 5. Conectar ao Amber

1. **Token:** Editor → ⚙️ **Configurações do projeto** → **Propriedades do script** → adicione `AMBER_TOKEN` = sua chave. Nunca coloque o token no código nem em conversas.
2. **Rode `setup` de novo** (é seguro repetir): ele cria a aba **Indice_Amber**, usada pelo índice de lotes.
3. **Endereço e cabeçalhos:** já estão em `CONFIG` (`AMBER_BASE_URL`, `Authorization` = token, `Integrador: ATONAPI`).
4. Troque `CONFIG.MODO_TESTE` para `false` e faça `clasp push` (ou cole o código de novo) e uma nova versão da implantação.

**Como funciona:** a API do Aton não tem endpoint "itens do lote". Cada pedido de venda traz o campo **`lotesep`** (lote de separação) e a lista de itens. Como um pedido agendado de até 30 dias atrás pode entrar num lote de hoje, o sistema lê os pedidos dos últimos `AMBER_DIAS_INDICE` dias (30), guarda na aba **Indice_Amber** os que estão em algum lote e, a partir dela, preenche os itens de cada lançamento. A descrição dos produtos novos vem de `produtos/listagemid`. Pedidos **cancelados** depois de gerar o lote continuam contando (a separação foi feita).

**Rotina do dia:**
1. Gere os lotes no Amber normalmente.
2. Ao longo do dia, lance cada separação no app (aba **Lançar**): data, número do lote, separador e, se for medir o tempo, início e término (opcionais). **Nada é consultado no Amber nesse momento**: o lançamento fica com o ícone ⏳ (aguardando itens).
3. **Uma vez ao dia, no fim da tarde**, com todos os lotes já gerados, clique em **Atualizar itens do dia**. A barra de progresso mostra a leitura do Amber (alguns minutos; deixe a página aberta) e, no fim, os itens são gravados em todos os lançamentos pendentes. Se der erro ou você fechar a página, clique de novo: continua de onde parou.
4. Só depois disso o painel tem os itens e pontos desses lotes (até lá ele avisa quais lotes ainda estão sem itens).
5. Lançamentos feitos depois da atualização ficam pendentes até a atualização do dia seguinte (que também preenche pendências antigas).

**Resultado da atualização:** o app informa quantos lançamentos foram preenchidos e lista (a) **lotes não encontrados** no Amber, geralmente número digitado errado: exclua o lançamento (🗑 na lista) e lance de novo; (b) **lotes gerados durante a leitura**: clique em "Atualizar itens do dia" de novo. Não gere lotes nem mexa em pedidos enquanto a atualização roda.

**Regra de segurança:** a atualização lê primeiro os dias mais recentes e anota o maior número de lote visto, e no fim relê o dia mais recente. Se nenhum lote novo surgiu nesse intervalo, todos os lotes lidos são confirmados. Se surgiu, só valem os lotes até o maior número visto no início; os demais ficam pendentes até a próxima atualização.

**Conferir a integração:** com `MODO_TESTE = false`, a função **`testarAmber`** (só consulta) mostra no Registro de execução quantos pedidos vêm por posição, quantos pedidos há por `lotesep`, se pedidos cancelados mantêm o lote, a distribuição das datas e os itens do lote `CONFIG.TESTE_LOTE`. **`testarPedido`** mostra a posição, as datas e o lote de um pedido (`CONFIG.TESTE_PEDIDO`), para investigar um item que não apareceu.

### 6. Autorizar

Na primeira execução aparece "Autorização necessária":
**Revisar permissões** → escolha sua conta → *O Google não verificou este app* → **Avançado** → **Acessar … (não seguro)** → **Permitir**.
É normal para scripts próprios. As permissões são: planilhas (ler/gravar) e acesso externo (chamar o Amber).

### 7. Publicar o Web App

Editor → **Implantar → Nova implantação → ⚙️ Tipo: App da Web**
- **Executar como:** Eu (você)
- **Quem pode acessar:** conforme abaixo
- **Implantar** e copie a URL `https://script.google.com/macros/s/.../exec`.

Pelo clasp: `npx clasp deploy`. O acesso vem do `appsscript.json` (padrão `MYSELF`, só você).

**Quem pode acessar:** como o app roda *como você*, quem abrir o link grava na sua planilha. Prefira **"Qualquer pessoa da sua organização"** (Google Workspace). Com Gmail pessoal, use "Qualquer pessoa" só se aceitar que o link vira a "senha" do sistema, e não o divulgue. No celular, abra a URL e use "Adicionar à tela inicial".

**Atualizações:** depois de `clasp push`, vá em **Implantar → Gerenciar implantações → ✏️ → Versão: Nova versão → Implantar**. A URL continua a mesma.

## Como funciona

- **Lançamento:** valida os dados e grava o registro (aba Registros). Um lote só pode ser lançado uma vez, e o salvamento usa `LockService`. Os itens do lote (agrupados por SKU) são gravados na atualização do fim do dia; no `MODO_TESTE` são gravados na hora.
- **Duração:** término − início em minutos. Se o término for menor, considera virada da meia-noite. Término igual ao início é recusado.
- **Lote sem horário:** início e término são opcionais, mas é **os dois ou nenhum**. Um lote sem horário conta em lotes, itens e pontos, mas fica **fora** das horas, de itens/h e de pontos/h (para não inflar a produtividade). O painel mostra quantos lotes estão sem horário e, na lista, aparece "sem horário".
- **Dificuldade:** SKU inédito entra em **Produtos** com dificuldade em branco (pendente). O peso é lido da planilha na hora de gerar o painel, então classificar um produto vale também para lotes antigos. Pendente = peso 1, com aviso no painel.
- **Painel:** Pontos = quantidade × dificuldade. Pontos/hora da equipe = pontos totais ÷ horas totais trabalhadas. Semana = segunda a domingo. O ranking é ordenado por pontos/hora. Itens sem classificação aparecem em "Simples" até serem classificados.
- **Excluir um lançamento:** na aba **Lançar**, em **Lançamentos**, escolha a data e clique em 🗑 no lançamento. O app pede uma confirmação e remove o registro e os itens dele; o lote fica livre para ser lançado de novo. Se excluir por engano, use o histórico de versões do Google Sheets (Arquivo → Histórico de versões).

## Limites e pontos de atenção

- **Janela do índice:** só cobre `AMBER_DIAS_INDICE` dias (30). Pedidos mais antigos que isso, dentro de um lote, não entram. Aumente o valor se precisar (a atualização fica mais longa).
- **Lote digitado errado:** só é descoberto na atualização dos itens (lote não encontrado). Exclua o lançamento (🗑) e lance de novo.
- **Lote 0 / com letras:** no Amber `lotesep = 0` significa "sem lote". Com o Amber ligado, o app só aceita lotes numéricos maiores que zero (no `MODO_TESTE` aceita qualquer um).
- **Carga no Amber:** cada atualização completa faz centenas de consultas (leitura). Faça quando precisar, não de hora em hora. Cada consulta tem até 3 tentativas para erros temporários do servidor.
- **Cota do Google:** cada atualização completa consome parte do limite diário de consultas externas e de tempo de execução do Apps Script; confira os limites atuais da sua conta (Gmail ou Workspace) na documentação do Google.
