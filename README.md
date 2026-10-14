# Controle de Separação

Web App em Google Apps Script que usa uma planilha do Google Sheets como banco de dados e consulta o Amber (Aton ERP) para trazer os itens de cada lote de separação.

> **O Amber é usado só para consulta.** O código faz apenas `GET`, sem corpo, em um único ponto (`consultarAmber_`). Qualquer método diferente de GET é recusado. O sistema não cria, altera nem apaga nada no Amber.

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

1. **Token:** Editor → ⚙️ **Configurações do projeto** → **Propriedades do script** → adicione `AMBER_TOKEN` = sua chave. Nunca coloque o token no código.
2. **Endpoint:** em `Code.gs`, ajuste `CONFIG.AMBER_URL` com `{lote}` no lugar do número do lote. Veja a seção "Endpoint pendente" abaixo.
3. **Cabeçalhos:** já vão `Authorization: <token>` e `Integrador: ATONAPI`. Se a API exigir um prefixo (ex.: `Bearer `), use `CONFIG.AMBER_PREFIXO_TOKEN`.
4. Rode **`testarAmber`** (com `MODO_TESTE = false`). O **Registro de execução** mostra a resposta crua do lote de `CONFIG.TESTE_LOTE` e os itens mapeados.
   Se os nomes dos campos forem outros, ajuste as listas dentro de `mapearItensAmber_` (é a única função a mexer).
5. Troque `CONFIG.MODO_TESTE` para `false` e rode `npx clasp push`.

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

## Endpoint pendente

Não consegui abrir a documentação do Postman (https://documenter.getpostman.com/view/12902127/VV4tUJaC) a partir do ambiente de desenvolvimento, porque a rede bloqueia o domínio. Por isso **`CONFIG.AMBER_URL` ainda é um placeholder**. Preciso do endpoint que retorna os itens de um lote de separação (método GET, com o número do lote) para fechar essa parte.
