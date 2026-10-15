# Testes
## Executados (Node, com a Carteira real de 1.421 registros / 265 bandeiras e o modelo fictício)
- Leitura da Carteira com campos opcionais (Rede, UF, Gerente); geração dos Excel ZDEF/ZT31 e do .eml (cabeçalhos, anexo base64).
- Validações: código vazio, condição vazia, dias inválidos, datas invertidas, Montante vazio.
- Base de Destinatários: modelo fictício gerado (11 linhas, e-mail de teste) e relido; filtro por automação/AMBAS/bandeira/ativo; remoção de duplicados (PARA prevalece); erro detalhado de colunas ausentes.
- Varredura estática do código: nenhum XMLHttpRequest, WebSocket, sendBeacon, localStorage, sessionStorage, IndexedDB, cookie ou URL externa; única chamada de rede: `fetch('config/regras.json')` (mesmo site). CSP no index.html: `connect-src 'self'`, `script-src 'self'`.
- **Navegador real (Chromium headless, via http e file://):** fluxo completo ZDEF/BRAIR com a Carteira real (carregar, dashboard, bandeira, preenchimento, destinatários com modelo baixado, painel de e-mail, ZIP com xlsx + eml + resumo). Resultado: 0 requisições externas e 0 erros de JavaScript. Também verificado: troca de bandeira sem apagar o campo, anexo da Base de Destinatários e tabela anexada no painel lateral.
## NÃO executados (a fazer antes do uso)
- Teste em Edge/Chrome corporativos e em outras resoluções; fluxo ZT31 completo no navegador.
- Abertura dos .eml no Outlook corporativo (editar, anexo, enviar). Se falhar, usar alternativa (ex.: baixar o Excel e montar o e-mail com o texto exibido no painel).
- Prova de rede: F12 > Network durante o processamento (esperado: nenhuma requisição).
- Validação da Segurança da Informação.
