# LINK10 · Tratador de Planilhas para WhatsApp

App **100% no navegador** para tratar planilhas Excel e padronizar números de telefone no formato WhatsApp (com DDI `55` no início, somente dígitos). Nada é enviado para servidor — a planilha é processada localmente.

> Pensado para o fluxo de inadimplência da LINK10, mas funciona com qualquer planilha que tenha uma coluna de telefone.

## Funcionalidades

- Upload por drag & drop ou clique (suporta `.xlsx`, `.xls`, `.csv`)
- Detecção automática de qual coluna parece ser a do telefone (auto-pick)
- Suporte a planilhas **sem cabeçalho** (caso do `Inad_1905.xlsx`) ou **com cabeçalho**
- Seleção da aba (sheet) quando o arquivo tem várias
- Limpeza configurável:
  - Remove tudo que não é dígito (espaços, parênteses, hífens, pontos, `+`, letras)
  - Adiciona DDI (padrão `55`) no início
  - **Prefixo inteligente**: detecta se o número já está prefixado (12-13 dígitos começando com o DDI) e não duplica
  - Filtra números inválidos por tamanho mínimo (padrão: 10 dígitos)
  - Remove duplicatas (mantém a primeira ocorrência)
- Preview lado a lado das primeiras 10 linhas antes de exportar
- Estatísticas: total, tratados, duplicatas removidas, inválidos descartados
- Exporta novo `.xlsx` com a coluna `whatsapp` adicionada no início + todas as colunas originais
- A coluna `whatsapp` é gravada como texto, garantindo que o Excel não corte o `55` ou converta para notação científica

## Como usar

1. Abra o app (link do GitHub Pages ou rode localmente, veja abaixo)
2. Arraste a planilha para a área de upload
3. Escolha a aba e a coluna do telefone (geralmente já vem auto-selecionada)
4. Ajuste as opções (DDI, mínimo de dígitos, duplicatas etc.)
5. Clique em **Processar planilha**
6. Confira o preview e os totais
7. Clique em **Baixar planilha tratada (.xlsx)**

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
LINK10APP/
├── index.html      # UI
├── styles.css      # Estilo dark, inspirado no WhatsApp
├── app.js          # Lógica (leitura/escrita xlsx, limpeza)
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
