<p align="center"><img src="assets/logo-horizontal.svg" alt="Concilia.CFT" width="420"></p>

<p align="center"><b>Conciliação bancária assistida: cruza o extrato do banco com o controle interno e mostra o que bate e o que falta dos dois lados.</b></p>

<p align="center">Python · Streamlit · pandas</p>

<p align="center">
  <img src="https://img.shields.io/badge/status-em%20testes-F2B233?style=flat-square" alt="Status">
  <img src="https://img.shields.io/badge/python-3.11-0B2545?style=flat-square&logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/streamlit-1.50%2B-12A37A?style=flat-square&logo=streamlit&logoColor=white" alt="Streamlit">
  <img src="https://img.shields.io/badge/uso-privado-13315C?style=flat-square" alt="Uso privado">
</p>

---

## Sumário

1. [Visão geral](#1-visão-geral)
2. [Como funciona](#2-como-funciona)
3. [Formatos aceitos](#3-formatos-aceitos)
4. [Como rodar (sem instalar nada)](#4-como-rodar-sem-instalar-nada)
5. [Estado do projeto](#5-estado-do-projeto)
6. [Identidade visual](#6-identidade-visual)
7. [Estrutura do repositório](#7-estrutura-do-repositório)
8. [Segurança e dados](#8-segurança-e-dados)
9. [Pendências](#9-pendências)

---

## 1. Visão geral

| Item | Descrição |
|---|---|
| Nome | **Concilia.CFT** (antes Concilia.CRT / Consolida.CRT) |
| Para quem | Setor financeiro do conselho, na conferência mensal entre banco e controle interno |
| O que entrega | Sugestões de pares com nota de confiança, listas do que está **só no banco** e **só no sistema**, ajustes manuais, justificativas e relatório em Excel com histórico |
| Arquivo principal | `concilia_cft.py` |

## 2. Como funciona

| Etapa | O que acontece |
|---|---|
| 1. Importar | Envie o extrato bancário e o controle interno e informe o responsável pela conciliação |
| 2. Conferir colunas | O app detecta a linha de cabeçalho e as colunas de **Data**, **Descrição** e **Valor** (ou **Crédito/Débito**). Tudo pode ser ajustado antes de começar. Linhas de saldo e de total são ignoradas |
| 3. Pareamento | Para cada valor igual, o app considera a data (dentro da tolerância definida) e a semelhança da descrição. Cada par recebe a confiança **Alta**, **Média** ou **Baixa**. Quando há vários lançamentos com o mesmo valor e a mesma data, o par aparece marcado como ambíguo |
| 4. Revisar | Concilie ou recuse cada sugestão. As ações em massa fora da confiança alta pedem confirmação. Uma sugestão recusada devolve os dois lançamentos para as pendências e pode ser restaurada |
| 5. Ajustes manuais | Pareie manualmente o que sobrou, inclusive com valores diferentes (com confirmação), ou justifique tarifas, estornos e rendimentos |
| 6. Relatório | Planilha `.xlsx` com as abas Resumo, Conciliados, Só no banco, Só no sistema, Justificados, A revisar, Recusados e Histórico (quem fez o quê e quando) |

## 3. Formatos aceitos

| Formato | Observações |
|---|---|
| **OFX** | Extrato exportado pelo internet banking. Já vem estruturado e é o formato recomendado |
| **CSV / TXT** | Separador `;`, `,`, tabulação ou `|`, detectado automaticamente. Codificação UTF-8 ou Windows (Latin-1) |
| **XLSX / XLS** | Todas as abas. A aba usada é escolhida na tela |
| Valores | `1.500,00` · `1,500.00` · `1500` · `R$ -150,00` · `(150,00)` · `150,00-` · `1.500,00 D` / `C` |
| Datas | Datas do Excel, `02/10/2026`, `2/10/26`, `2026-10-02` e o formato do OFX |

A pasta `exemplos/` tem arquivos **fictícios** para testar o app.

## 4. Como rodar (sem instalar nada)

| Caminho | Passo a passo |
|---|---|
| **GitHub Codespaces** (para testar) | No repositório: **Code → Codespaces → Create codespace on main**. O ambiente instala tudo e abre o app sozinho na porta 8501 |
| **Streamlit Community Cloud** (para deixar no ar) | Em **share.streamlit.io → Create app → Deploy a public app from GitHub**: repositório `bkservicoss/concilia.cft`, branch `main`, arquivo `concilia_cft.py` → **Deploy** |

As dependências estão em `requirements.txt`, e o tema e o limite de upload (50 MB) ficam em `.streamlit/config.toml`.

## 5. Estado do projeto

| Área | Situação |
|---|---|
| Leitura de OFX, CSV, XLS e XLSX | ✅ Pronta |
| Detecção de cabeçalho e mapeamento de colunas | ✅ Pronta |
| Pareamento por valor + data + descrição, com confiança | ✅ Pronto |
| Listas "só no banco" e "só no sistema" | ✅ Prontas |
| Conciliação manual, justificativas, desfazer e restaurar | ✅ Prontos |
| Relatório Excel com trilha de auditoria | ✅ Pronto |
| Salvar conciliações entre sessões | ⏳ Pendente (hoje tudo fica só na memória da sessão) |
| Login de usuários | ⏳ Pendente |

## 6. Identidade visual

<p align="center"><img src="assets/logo.svg" alt="Ícone" width="72"></p>

O ícone mostra duas colunas (banco e controle interno) ligadas por um sinal de conferido.

| Uso | Cor | Hex |
|---|---|---|
| Principal (fundo da marca, títulos) | Azul-marinho | `#0B2545` |
| Apoio | Azul | `#13315C` |
| Conciliado / ação principal | Verde | `#12A37A` |
| Destaque / atenção | Âmbar | `#F2B233` |
| Saídas / erro | Vermelho | `#D64545` |
| Texto secundário | Cinza | `#5B6B7F` |
| Fundo | Cinza-claro | `#F5F7FA` |

| Item | Definição |
|---|---|
| Fonte | Inter (com Segoe UI e Arial como alternativas). No app, a fonte sem serifa padrão do Streamlit |
| Logo horizontal | `assets/logo-horizontal.svg` |
| Ícone | `assets/logo.svg` (vetor) e `assets/icone.png` (256 px, ícone da aba do navegador) |
| Escrita do nome | **Concilia**, com **.CFT** em verde |

## 7. Estrutura do repositório

```
concilia.cft/
├── CLAUDE.md                   # instruções do projeto para o Claude
├── concilia_cft.py             # aplicativo (leitura, pareamento, interface e relatório)
├── requirements.txt            # dependências
├── assets/
│   ├── logo.svg                # ícone
│   ├── logo-horizontal.svg     # logo com nome
│   └── icone.png               # ícone da aba do navegador
├── exemplos/                   # arquivos fictícios para teste
│   ├── extrato_exemplo.csv
│   ├── extrato_exemplo.ofx
│   └── controle_exemplo.xlsx
├── .streamlit/config.toml      # tema com as cores da marca
├── .devcontainer/              # configuração do Codespaces
└── .gitignore
```

## 8. Segurança e dados

| Regra | Detalhe |
|---|---|
| Arquivos enviados | São processados só na memória da sessão. Nada é gravado em disco nem enviado para fora |
| Planilhas reais | **Nunca** devem ir para o repositório. A pasta `dados/` está no `.gitignore` para uso local |
| Segredos | Se um dia houver chaves (banco de dados, login), elas ficam nos Secrets do Streamlit Cloud ou do Codespaces, nunca no código |
| Visibilidade | Repositório público hoje. Para dados do conselho, prefira tornar o repositório privado antes de publicar o app |

## 9. Pendências

- [ ] Tornar o repositório privado antes do uso com dados reais
- [ ] Publicar no Streamlit Community Cloud
- [ ] Testar com extratos reais do banco usado pelo conselho
- [ ] Salvar conciliações entre sessões (ex.: Supabase)
- [ ] Login de usuários

<p align="center"><sub>Projeto privado · BK Serviços, Produção &amp; Comunicação</sub></p>
