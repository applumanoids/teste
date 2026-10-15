# Central de Automações (ZDEF e ZT31)
100% estático (HTML/CSS/JS), roda no navegador. Bibliotecas locais em `libs/` (ExcelJS 4.4.0, JSZip 3.10.1). Sem CDN, banco, cookies ou localStorage.

## Publicar no GitHub Pages
1. Crie o repositório (preferencialmente privado/organização). 2. Envie o CONTEÚDO desta pasta. **Nunca envie planilhas com dados reais** (a pasta `modelos/` só tem dados fictícios). 3. Settings > Pages > branch `main`, pasta `/ (root)`.

## Estrutura
`index.html` · `js/core.js` (leitura, validações, Excel, .eml, destinatários) · `js/app.js` (telas/etapas) · `js/automacoes.js` (registro dos cards) · `js/paginas.js` (Manual e Privacidade) · `config/regras.json` · `modelos/Modelo_Base_Destinatarios.xlsx` (também baixável na Etapa 5).

## Fluxo (7 etapas)
Base → Dashboard → Bandeira → Clientes/preenchimento → Destinatários → E-mail (painel lateral) → Gerar/Baixar. "Voltar/Continuar" preservam o que foi preenchido. Vários rascunhos: "Adicionar outra bandeira ao lote".

## Colunas obrigatórias
- Carteira (ZDEF): Bandeira, Código Cliente, Condição de Pagamento. Carteira (ZT31): Bandeira, Código Cliente. Grupo de Preço é opcional (vazio vira `*`). Rede, UF e Nome Gerente Regional só alimentam o dashboard (se faltarem, aparecem como "(vazio)").
- Base de Destinatários: Automação (ZDEF/ZT31/AMBAS), Bandeira, Nome, Email, Tipo (PARA/CC), Ativo (SIM/NÃO).

## Cruzamentos
- Carteira.Bandeira = bandeira escolhida → todos os clientes entram (sem seleção individual).
- Destinatários: Ativo = SIM, Bandeira igual e Automação = escolhida **ou** AMBAS. PARA → campo Para; CC → Cc; endereços repetidos são eliminados (PARA prevalece). Pode editar/adicionar/remover antes de gerar. Sem PARA, a geração é bloqueada. Nenhum endereço fixo no código.

## Regras de negócio PENDENTES de aprovação
Divisão 8000 e Canal 12 (vêm dos exemplos; a Carteira não tem Canal); Grupo de Clientes e Setor = `*`; unidade do Montante (gravado como digitado, 2,5 → 2.5); datas opcionais no ZT31 (só alerta). Assinatura a preencher em `regras.json`.

## Validações implementadas
Extensão ≠ .xlsx, arquivo vazio, corrompido/com senha, aba CARTEIRA ausente, cabeçalho não encontrado, coluna obrigatória ausente (lista as encontradas), cabeçalho duplicado, linhas vazias ignoradas, bandeira inexistente, código vazio (indica linha), condição de pagamento vazia (ZDEF), dias inválidos, datas ausentes/invertidas, Montante inválido, códigos duplicados (alerta, sem remoção), zeros à esquerda preservados, e-mail inválido (com nome), e-mail sem endereço, sem destinatário PARA, falha ao gerar Excel/rascunho/ZIP, navegador incompatível, colunas da Base de Destinatários.

## Adicionar nova automação
1. `js/automacoes.js`: troque um card "Em breve" por `{id, titulo, desc, icone, ativo:true}`. 2. `js/core.js`: defina as colunas em `COLS`, a linha em `linhaReport` e as regras em `validar`. 3. `config/regras.json`: modelo de e-mail em `email.<id>`. 4. `js/app.js`: campos manuais em `MF()`. (Modular, mas ainda exige editar esses pontos.)

## Testes
Veja `TESTES.md`.
