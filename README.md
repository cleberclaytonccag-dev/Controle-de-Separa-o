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
2. **Endereço da API:** já está em `CONFIG.AMBER_BASE_URL` (`https://api.ambarxcall.com.br/AtonSNIsapi.dll/atonerp`), conforme a documentação. Os cabeçalhos `Authorization` (token) e `Integrador: ATONAPI` também já estão configurados.
3. **Como o lote é encontrado:** a API não tem um endpoint "itens do lote". Cada pedido de venda traz o campo **`lotesep`** (lote de separação) e a lista de itens. O sistema consulta os pedidos de venda de `AMBER_DIAS_ANTES` dias até a data do lançamento, nas posições de `AMBER_POSICOES`, filtra os pedidos com `lotesep` igual ao lote, e soma os itens por SKU (`codigo_interno`). A descrição vem de `produtos/listagemid` (só para SKUs que ainda não estão na aba Produtos).
4. Com `MODO_TESTE = false`, rode **`testarAmber`**. O Registro de execução mostra: quantos pedidos vieram por posição, o modo de paginação detectado, quantos pedidos há por `lotesep`, um pedido de exemplo (sem etiquetas) e os itens do lote `CONFIG.TESTE_LOTE`. **Confira se o `lotesep` bate com os lotes que você conhece**, e ajuste `AMBER_POSICOES` e `AMBER_DIAS_ANTES` ao que o seu fluxo precisa (menos posições e menos dias = consulta mais rápida).
5. Troque `CONFIG.MODO_TESTE` para `false` e rode `npx clasp push`.

**Tempo de consulta:** a primeira consulta varre os pedidos do período (50 por chamada). Com 3 dias e ~2.000 pedidos foram 44 chamadas e cerca de 1 minuto (a janela padrão agora é de 7 dias, então a primeira consulta leva mais). Como a varredura traz **todos os lotes** do período, o resultado fica em cache por 30 minutos (`AMBER_CACHE_SEGUNDOS`) e os lançamentos seguintes saem na hora. Se um lote não estiver no cache, o sistema varre de novo. O limite do Apps Script é de 6 minutos por execução. Para acelerar, reduza `AMBER_DIAS_ANTES`.

**Lote incompleto na janela:** se um lote tiver pedidos feitos antes do início da janela (`AMBER_DIAS_ANTES`), esses pedidos não entram. Teste com um lote antigo e compare com o Amber; se faltar item, aumente `AMBER_DIAS_ANTES`.

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

- **Lançamento:** valida os dados, consulta o Amber pelo lote e grava o registro e os itens agrupados por SKU. Um lote só pode ser lançado uma vez, e o salvamento usa `LockService`. Se o Amber falhar, nada é gravado.
- **Duração:** término − início em minutos. Se o término for menor, considera virada da meia-noite. Término igual ao início é recusado.
- **Dificuldade:** SKU inédito entra em **Produtos** com dificuldade em branco (pendente). O peso é lido da planilha na hora de gerar o painel, então classificar um produto vale também para lotes antigos. Pendente = peso 1, com aviso no painel.
- **Painel:** Pontos = quantidade × dificuldade. Pontos/hora da equipe = pontos totais ÷ horas totais trabalhadas. Semana = segunda a domingo. O ranking é ordenado por pontos/hora. Itens sem classificação aparecem em "Simples" até serem classificados.
- **Corrigir um lançamento errado:** apague a linha do lote em **Registros** e as linhas com o mesmo `ID_Registro` em **Itens_Lote**. Depois lance de novo.

## Limites e pontos de atenção

- **Lote em outra posição ou data:** o sistema só encontra pedidos cujas posições estão em `AMBER_POSICOES` e cuja data de pedido está dentro da janela. Se um lote retornar "nenhum pedido encontrado", confira o número e amplie a janela.
- **Lote sem pedidos no Amber:** `lotesep = 0` significa "sem lote". O app não aceita lote 0 nem lotes com letras ao consultar o Amber (no `MODO_TESTE` aceita).
- **Paginação:** a documentação não diz se `offset` é o número da página ou do registro. Com `AMBER_OFFSET_MODO: 'auto'` o sistema detecta sozinho; depois de rodar `testarAmber` você pode fixar `'pagina'` ou `'registro'`.
