# CLAUDE.md — Concilia.CFT

O Concilia.CFT é um app de conciliação bancária assistida feito em Streamlit. Ele cruza o extrato do banco (OFX, CSV, XLS ou XLSX) com o controle interno, sugere pares com nota de confiança, mostra o que está só no banco e só no sistema e gera um relatório Excel com trilha de auditoria.

Projeto privado da BK Serviços, Produção & Comunicação.

## Como o dono trabalha
- **Não instala nada no computador.** Tudo roda na nuvem: GitHub Codespaces para testar e Streamlit Community Cloud para deixar no ar. Nunca oriente a instalar Python, pacotes ou programas no PC. Quando algo precisar rodar, ajuste o repositório (`.devcontainer/`, `requirements.txt`) e explique o caminho pelo navegador.
- Fala **português do Brasil**: respostas, textos da interface, comentários, README e mensagens de commit.
- Prefere passo a passo curto, com o caminho exato dos menus (ex.: **Code → Codespaces → Create codespace on main**).
- Quando houver decisões a tomar, pergunte em checklist com opções clicáveis.

## Mapa
| Caminho | O que é |
|---|---|
| `concilia_cft.py` | App inteiro, dividido em seções numeradas: 1 normalização, 2 leitura, 3 mapeamento de colunas, 4 motor de pareamento, 5 estado e ações (auditoria), 6 relatório Excel, 7 interface |
| `requirements.txt` | Dependências (streamlit ≥ 1.50, pandas, openpyxl, xlrd, tzdata) |
| `assets/` | `logo.svg`, `logo-horizontal.svg`, `icone.png` |
| `exemplos/` | Arquivos **fictícios** para teste (CSV de banco, OFX, controle em XLSX) |
| `.streamlit/config.toml` | Tema com as cores da marca e limite de upload de 50 MB |
| `.devcontainer/devcontainer.json` | Codespaces: instala as dependências e abre o app na porta 8501 |

## Como rodar
- **Codespaces:** Code → Codespaces → Create codespace on main. O app abre sozinho (`streamlit run concilia_cft.py`).
- **Streamlit Community Cloud:** share.streamlit.io → Create app → repositório `bkservicoss/concilia.cft`, branch `main`, arquivo `concilia_cft.py`.
- Não há workflows do GitHub Actions. Nenhum caminho dispara compilação.

## Regras que não podem quebrar
- Valores são guardados em **centavos (int)**. Nada de float em comparação de dinheiro.
- O padrão brasileiro vence em caso de dúvida: `1.500` é mil e quinhentos, e `D`/`C` no fim do valor definem o sinal.
- Saldo nunca é lançamento: linhas cuja descrição começa com "saldo" ou "total" são descartadas, e a coluna de saldo nunca é sugerida como valor.
- O pareamento escolhe pela **maior nota global** (valor igual → data na tolerância → descrição parecida), sem reaproveitar lançamentos. Empates de valor e data marcam o par como ambíguo e rebaixam a confiança.
- As listas "só no banco" e "só no sistema" sempre ficam visíveis. Recusar uma sugestão devolve os dois lançamentos às pendências.
- Ações em massa fora da confiança **Alta** exigem confirmação.
- Toda ação do usuário passa por `registrar()` (histórico com responsável e hora de Brasília) e invalida o relatório em cache.
- Texto vindo das planilhas entra no HTML sempre com `html.escape`.
- Use `width="stretch"` (não `use_container_width`) nos componentes do Streamlit.

## Segurança
- Nunca versionar tokens, chaves, senhas, `.env` ou `.streamlit/secrets.toml` (já estão no `.gitignore`).
- Planilhas reais de extrato e controle **não entram no repositório**. Use a pasta `dados/`, que é ignorada.
- Se o dono colar um token ou senha no chat: não usar, não repetir e orientar a gerar outro na hora e salvar direto nos Secrets (Streamlit Cloud ou Codespaces).
- O repositório está público hoje. Recomende torná-lo privado antes do uso com dados reais.

## Identidade visual
| Uso | Cor |
|---|---|
| Azul-marinho (principal) | `#0B2545` |
| Azul (apoio) | `#13315C` |
| Verde (conciliado, ação principal) | `#12A37A` |
| Âmbar (destaque) | `#F2B233` |
| Vermelho (saídas, erro) | `#D64545` |
| Cinza (texto secundário) | `#5B6B7F` |
| Fundo | `#F5F7FA` |

- Fonte: Inter (com Segoe UI e Arial como alternativas). O nome se escreve **Concilia** com **.CFT** em verde.
- Ícone: duas colunas (banco e controle interno) ligadas por um sinal de conferido.

## README
- Siga o padrão da BK Serviços: logo centralizada, frase de apresentação em negrito, selos shields.io nas cores da marca, `---`, Sumário numerado, seções com tabelas, estrutura do repositório em árvore, pendências com `- [ ]` e o rodapé `Projeto privado · BK Serviços, Produção & Comunicação`.
- Toda mudança no repositório entra no README no mesmo trabalho.

## Repositórios novos
Sempre que o dono criar um repositório novo, pergunte se ele quer criar o `CLAUDE.md`.
