# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

Operador de cobrança/inadimplência que usa a ferramenta várias vezes ao dia, sempre na mesma rotina: recebe uma planilha (geralmente com formato parecido), padroniza os telefones e baixa o arquivo para disparar mensagens no WhatsApp. Prioridade: velocidade e não ter que reconfigurar tudo a cada uso.

## Product Purpose

Tratar planilhas Excel/CSV no navegador e padronizar números de telefone no formato WhatsApp (somente dígitos, DDI `55` no início), removendo duplicatas e números inválidos. Sucesso = soltar a planilha e baixar o arquivo certo com o mínimo de cliques.

## Positioning

100% local: a planilha nunca sai do navegador (relevante para LGPD ao lidar com dados de clientes). Sem login, sem servidor, sem instalação.

## Operating Context

- Planilhas `.xlsx`, `.xls` ou `.csv`, com ou sem cabeçalho, às vezes com várias abas.
- Exportação principal: planilha tratada padrão (`<arquivo>_whatsapp.xlsx`, coluna `whatsapp` no início + colunas originais).
- Exportação secundária: formato COBRANCA (`<arquivo>_cobranca.xlsx`, aba `Página1`, layout fixo do modelo, valores fixos em `atualizar` e `whatsapp`).
- Publicado via GitHub Pages a partir da branch `main`.

## Capabilities and Constraints

- HTML + CSS + JS puro, sem build e sem npm; SheetJS (`xlsx@0.18.5`) via CDN.
- Detecção automática da coluna de telefone e de nome.
- Regras: DDI configurável, prefixo inteligente (12/13 dígitos começando com o DDI), mínimo de dígitos, remover duplicatas, descartar inválidos.
- Coluna de telefone gravada como texto no Excel.

## Brand Commitments

- Feita pela Prisme. Padrão visual claro do Prisme Chatbot (o mesmo das páginas Benfica, DVI e Boa Vista): Nunito, primária `#00ba70`, fundo `#fafafa`, cards brancos. Confirmado pelo usuário em 2026-10-08 como identidade a manter.

## Evidence on Hand

Nenhum depoimento, métrica ou cliente público. Não inventar.

## Product Principles

1. Soltar o arquivo é a única ação obrigatória; o resto deve vir pronto.
2. Lembrar as escolhas do operador entre usos.
3. Mostrar o resultado (totais e prévia) antes do download, sem rolagem longa.
4. Nada sai do navegador.
