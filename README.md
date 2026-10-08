# Tratador de Planilhas para WhatsApp

App **100% no navegador** para tratar planilhas Excel e padronizar números de telefone no formato WhatsApp (com DDI `55` no início, somente dígitos). Nada é enviado para servidor — a planilha é processada localmente.

> Pensado para fluxos de cobrança/inadimplência, mas funciona com qualquer planilha que tenha uma coluna de telefone.

## Funcionalidades

- **Uma tela só**: soltou a planilha, o resultado já aparece. Mapeamento à esquerda, arquivo final ao vivo à direita
- Upload por drag & drop em qualquer lugar da página ou por clique (`.xlsx`, `.xls`, `.csv`); soltar outro arquivo troca a planilha
- **Exportação principal: Modelo de Disparo** (`<arquivo>_disparo.xlsx`), idêntico ao `Modelo de Disparo.xlsx`:
  - Aba `Página1`, colunas `telefone`, `nome`, `email`, `cpfcnpj`, `genero`, `estado`, `cidade`, `referencia`, `aniversario`, `endereco`, (coluna K vazia), `atualizar`, `carteira`, `whatsapp`, `tag`, `status`
  - **O operador escolhe onde cada dado vai**: para cada coluna do modelo, uma coluna da planilha, um valor fixo ou vazio
  - `telefone` recebe o número já tratado (com DDI, gravado como texto)
  - `atualizar` e `whatsapp` começam como valor fixo (`1` e `56`), editáveis
- Mapeamento automático pelo nome do cabeçalho (nome, e-mail, CPF/CNPJ, cidade, UF, nascimento...) e detecção da coluna de telefone pelos dígitos
- **Lembra as escolhas do último uso** (mapeamento, valores fixos e regras) no próprio navegador; na próxima planilha com os mesmos cabeçalhos é só baixar
- Suporte a planilhas **com ou sem cabeçalho** e seleção da aba quando há várias
- Limpeza configurável (recolhida por padrão, com resumo visível):
  - Remove tudo que não é dígito (espaços, parênteses, hífens, pontos, `+`, letras)
  - Adiciona DDI (padrão `55`) no início
  - **Prefixo inteligente**: não duplica o DDI em números com 12-13 dígitos que já começam com ele
  - Descarta números abaixo do mínimo de dígitos (padrão: 10)
  - Remove repetidos (mantém a primeira ocorrência)
- Prévia do arquivo final (até 200 linhas) e aba **Descartados** com linha, número original e motivo (repetido, curto demais, sem número)
- Exportação secundária: planilha original com a coluna `whatsapp` no início (`<arquivo>_whatsapp.xlsx`)

## Como usar

1. Abra o app (link do GitHub Pages ou rode localmente, veja abaixo)
2. Solte a planilha na página
3. Confira o mapeamento em **Onde cada dado vai** (na maioria das vezes já vem certo)
4. Clique em **Baixar Modelo de Disparo**

## Rodar localmente

Como é um app estático, qualquer servidor HTTP serve.

### Opção 1 — Python (já instalado na maioria das máquinas)

```powershell
python -m http.server 8080
```

Abra http://localhost:8080

### Opção 2 — Node

```powershell
npx serve .
```

### Opção 3 — Só abrir o `index.html`

Funciona, mas alguns navegadores restringem leitura de arquivos via `file://`. Prefira as opções acima.

## Deploy no GitHub Pages

1. Crie um repositório no GitHub e suba estes arquivos (`index.html`, `styles.css`, `app.js`, `README.md`).
2. No repositório: **Settings → Pages**.
3. Em **Source**, escolha **Deploy from a branch**.
4. Branch: `main` · Folder: `/ (root)` · Save.
5. Aguarde ~1 min — o app fica em `https://<seu-usuario>.github.io/<nome-do-repo>/`.

Comandos:

```powershell
git init
git add .
git commit -m "feat: tratador de planilhas para whatsapp"
git branch -M main
git remote add origin https://github.com/<seu-usuario>/<nome-do-repo>.git
git push -u origin main
```

## Como a limpeza funciona (exemplos)

| Entrada                         | Saída         | Observação                                            |
| ------------------------------- | ------------- | ----------------------------------------------------- |
| `(32) 99978-5390`               | `5532999785390` | Remove `( )`, espaço, `-`, e adiciona `55`            |
| `32999785390`                   | `5532999785390` | Adiciona o `55`                                        |
| `5532999785390`                 | `5532999785390` | Já tem 13 dígitos começando com `55` → não duplica    |
| `55319605231`                   | `5555319605231` | 11 dígitos: o `55` inicial é DDD, então adiciona DDI  |
| ` 011 9 8403-0561 `             | `5511984030561` | Espaços/pontuação fora, `55` adicionado               |
| `abc123`                        | (inválido)    | Apenas 3 dígitos — descartado se filtro ligado        |

> O "prefixo inteligente" assume que apenas números com **12 ou 13 dígitos** começando com o DDI já estão prefixados. Isso evita falsos positivos como o caso da DDD `55` (Rio Grande do Sul).

## Estrutura do projeto

```
corretor-planilhas/
├── index.html      # UI
├── styles.css      # Visual claro padrão Prisme Chatbot
├── app.js          # Lógica (leitura/escrita xlsx, limpeza)
├── PRODUCT.md      # Contexto do produto (usuários, princípios)
├── README.md
└── .gitignore
```

## Stack

- HTML + CSS + JavaScript puro (sem build, sem dependências de npm)
- [SheetJS](https://sheetjs.com/) (`xlsx`) via CDN — para ler/escrever Excel no navegador

## Privacidade

A planilha **nunca sai do navegador**. Toda a leitura, processamento e geração do arquivo final acontecem do lado do cliente. Importante para conformidade com LGPD ao lidar com dados de clientes.

## Licença

MIT
