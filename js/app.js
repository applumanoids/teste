/* Central de Automações — navegação e fluxo em etapas. Nada é salvo no navegador nem enviado pela rede. */
(function () {
  const $ = s => document.querySelector(s), C = Core;
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const STEPS = ['Base de dados', 'Dashboard', 'Bandeira', 'Clientes', 'Destinatários', 'E-mail', 'Gerar'];
  const S = { tipo: null, step: 1, carteira: null, arq: null, dest: null, band: '', rows: [], rec: [], recFor: '', jobs: [], base: null,
    cfg: { divisao: '', canal: '', grupoClientes: '*', setor: '*' }, mail: { assunto: '', saud: '', corpo: '', assinatura: '' } };
  const msg = (t, k = 'err') => { const d = document.createElement('div'); d.className = k; d.textContent = t; $('#msg').appendChild(d); d.scrollIntoView({ block: 'nearest' }); };
  const lista = (itens, k) => { const d = document.createElement('div'); d.className = k; d.innerHTML = `<b>${itens.t}</b><ul class="lst">${itens.l.map(x => `<li>${esc(x)}</li>`).join('')}</ul>`; return d; };
  const clr = () => $('#msg').innerHTML = '';
  const view = v => { clr(); document.querySelectorAll('[data-v]').forEach(b => b.classList.toggle('act', b.dataset.v === v || (v === 'auto' && b.dataset.v === 'home'))); ['home', 'auto', 'manual', 'priv'].forEach(x => $('#v-' + x).hidden = x !== v); if (PAGINAS[v]) $('#v-' + v).innerHTML = PAGINAS[v]; };
  document.querySelectorAll('[data-v]').forEach(b => b.onclick = () => view(b.dataset.v));
  $('#reiniciar').onclick = () => { if (confirm('Limpar todos os dados da sessão e recomeçar?')) location.reload(); };
  if (!window.ExcelJS || !window.JSZip) msg('Este navegador não é compatível. Use o Microsoft Edge ou o Google Chrome atualizados.');
  fetch('config/regras.json').then(r => r.json()).then(r => S.base = r).catch(() => { S.base = { divisaoNegocio: '', canal: '', grupoClientes: '*', setorAtividade: '*', email: { assinatura: '', ZDEF: { assunto: '', saudacao: '', corpo: '' }, ZT31: { assunto: '', saudacao: '', corpo: '' } } }; msg('Não foi possível ler config/regras.json. Preencha as regras na Etapa 4.', 'warn'); });

  $('#cards').innerHTML = AUTOS.map(a => `<button class="card${a.ativo ? '' : ' off'}" data-id="${a.id}" ${a.ativo ? '' : 'disabled'}><i>${a.icone}</i><b>${a.titulo}</b><span>${a.ativo ? a.desc : '🔜 Em breve'}</span></button>`).join('');
  $('#cards').onclick = e => { const b = e.target.closest('.card'); if (b && !b.disabled) abrir(b.dataset.id); };

  function abrir(id) {
    const b = S.base, t = b.email[id];
    Object.assign(S, { tipo: id, step: 1, band: '', rows: [], rec: [], recFor: '', jobs: [], cfg: { divisao: b.divisaoNegocio, canal: b.canal, grupoClientes: b.grupoClientes, setor: b.setorAtividade }, mail: { assunto: t.assunto, saud: t.saudacao, corpo: t.corpo, assinatura: b.email.assinatura } });
    if (S.carteira && S.carteira.tipo !== id) { S.carteira = null; }
    view('auto'); render();
  }

  async function lerWb(f) {
    if (!/\.xlsx$/i.test(f.name)) throw new Error(`O arquivo "${f.name}" não tem extensão .xlsx.\nSalve a planilha como "Pasta de Trabalho do Excel (.xlsx)" e carregue novamente.`);
    if (!f.size) throw new Error(`O arquivo "${f.name}" está vazio (0 KB). Selecione o arquivo correto.`);
    const wb = new ExcelJS.Workbook();
    try { await wb.xlsx.load(await f.arrayBuffer()); } catch (e) { throw new Error(/password|encrypt|CFB|Central Directory|zip/i.test(String(e.message)) ? `Não foi possível abrir "${f.name}". O arquivo pode estar protegido por senha ou corrompido.\nRemova a senha no Excel (Arquivo > Informações > Proteger Pasta de Trabalho), salve e carregue novamente.` : `Não foi possível abrir "${f.name}": o arquivo parece corrompido ou não é um .xlsx válido.`); }
    return wb;
  }
  const dz = (id, fn) => { const d = $('#' + id), i = d.querySelector('input'); d.onclick = () => i.click(); i.onchange = () => i.files[0] && fn(i.files[0]); d.ondragover = e => { e.preventDefault(); d.classList.add('over'); }; d.ondragleave = () => d.classList.remove('over'); d.ondrop = e => { e.preventDefault(); d.classList.remove('over'); e.dataTransfer.files[0] && fn(e.dataTransfer.files[0]); }; };
  const dropHtml = (id, ic, t1, t2, bt) => `<label id="${id}" class="drop" tabindex="0"><span class="ic">${ic}</span><b>${t1}</b><small>${t2}</small><span class="bt">${bt}</span><input type="file" accept=".xlsx" hidden></label>`;

  async function carregarBase(f) {
    clr();
    try {
      const p = C.parseCarteira(await lerWb(f), S.tipo); if (!p.ok) throw new Error(p.erro);
      S.carteira = { ...p, tipo: S.tipo }; S.arq = { nome: f.name, kb: (f.size / 1024).toFixed(0), em: new Date().toLocaleString('pt-BR') }; S.band = ''; S.rows = []; S.jobs = [];
      render(); msg(`Base carregada com sucesso. Foram encontrados ${p.info.registros.toLocaleString('pt-BR')} registros e ${p.info.bandeiras.toLocaleString('pt-BR')} bandeiras.`, 'ok');
    } catch (e) { msg(e.message); }
  }
  async function carregarDest(f) {
    clr();
    try { const p = C.parseDest(await lerWb(f)); if (!p.ok) throw new Error(p.erro); S.dest = p.rows; S.destArq = { nome: f.name, n: p.rows.length }; S.recFor = ''; fillRec(); render(); msg(`Base de Destinatários carregada: ${p.rows.length} linha(s).`, 'ok'); } catch (e) { msg(e.message); }
  }
  function fillRec() { const k = S.tipo + '|' + S.band; if (S.dest && S.recFor !== k) { S.rec = C.destinatarios(S.dest, S.tipo, S.band); S.recFor = k; } }

  const mailStr = t => S.rec.filter(r => r.tipo === t && r.email.trim()).map(r => r.email.trim()).join(';');
  const curJob = () => ({ tipo: S.tipo, bandeira: S.band, rows: S.rows.map(r => ({ ...r })), cfg: { ...S.cfg }, rec: S.rec.map(r => ({ ...r })), mail: { ...S.mail, para: mailStr('PARA'), cc: mailStr('CC') } });
  const emailOk = e => /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/.test(e.trim());

  function gate(n) { // retorna {t,l} de erros ao tentar sair da etapa n
    if (n === 1 && !S.carteira) return { t: 'Carregue a base Carteira para continuar.', l: [] };
    if (n === 3 && !S.band) return { t: 'Selecione uma bandeira da lista para continuar.', l: [] };
    if (n === 4) { const v = C.validar({ ...curJob(), mail: undefined }); if (v.erros.length) return { t: `${v.erros.length} erro(s) impedem a geração. Seus dados continuam na tela; corrija e continue:`, l: v.erros }; }
    if (n === 5) {
      const l = []; if (!S.dest) l.push('Carregue a Base de Destinatários (ou use o botão para baixar o modelo).');
      S.rec.forEach(r => { if (!r.email.trim()) l.push(`O destinatário "${r.nome || '(sem nome)'}" está sem endereço de e-mail.`); else if (!emailOk(r.email)) l.push(`O endereço "${r.email}" de ${r.nome || '(sem nome)'} é inválido. Corrija o endereço (exemplo: nome@empresa.com).`); });
      if (!mailStr('PARA')) l.push(`Não há destinatário principal (PARA) para ${S.tipo} / ${S.band}. O e-mail não pode ser gerado sem destinatário: cadastre na Base de Destinatários ou adicione manualmente.`);
      if (l.length) return { t: 'Revise os destinatários:', l };
    }
    return null;
  }
  function go(n) {
    clr(); if (n > S.step) { const g = gate(S.step); if (g) { $('#msg').appendChild(lista(g, 'err')); return; } }
    if (n === 5) fillRec(); S.step = Math.max(1, Math.min(7, n)); render(); window.scrollTo(0, 0);
  }

  function render() {
    $('#prog').innerHTML = `<div class="prog">${STEPS.map((s, i) => `<div class="${i + 1 === S.step ? 'on' : i + 1 < S.step ? 'done' : ''}">${i + 1}. ${s}</div>`).join('')}</div>`;
    $('.lay').className = 'lay ' + (S.step < 4 ? 'noaside' : S.step === 6 ? 'w2' : 'w1');
    const a = AUTOS.find(x => x.id === S.tipo);
    $('#body').innerHTML = `<h2>${a.icone} ${a.titulo}</h2>` + [null, s1, s2, s3, s4, s5, s6, s7][S.step]();
    const nv = `<button class="sec bk" ${S.step === 1 ? 'disabled' : ''}>← Voltar</button>${S.step < 7 ? '<button class="fw">Continuar →</button>' : ''}`;
    $('#nav').innerHTML = $('#navtop').innerHTML = nv;
    document.querySelectorAll('.navb .bk').forEach(b => b.onclick = () => go(S.step - 1)); document.querySelectorAll('.navb .fw').forEach(b => b.onclick = () => go(S.step + 1));
    [null, b1, b2, b3, b4, b5, b6, b7][S.step](); preview(); document.querySelectorAll('[data-w]').forEach(e => requestAnimationFrame(() => e.style.width = e.dataset.w + '%')); document.querySelectorAll('[data-bg]').forEach(e => e.style.background = e.dataset.bg);
  }

  /* Etapa 1 */
  function s1() {
    const z = S.tipo === 'ZDEF', cols = [['Head', 0], ['Rede', 0], ['Bandeira', 1], ['Diretoria', 0], ['Desc. Diretoria', 0], ['Setor de Atividade', 0], ['Código Cliente', 1], ['Razão Social Cliente', 0], ['Grupo de Preço', 0], ['Condição de Pagamento', z ? 1 : 0], ['Grupo de Cliente', 0], ['Descrição Grupo de Cliente', 0], ['UF', 0], ['Nome Executivo de Vendas', 0], ['Nome Gerente Regional', 0]];
    const info = S.carteira ? `<div class="ok">✅ <b>${esc(S.arq.nome)}</b> carregada (${S.arq.kb} KB, ${S.arq.em})<br>Abas: ${esc(S.carteira.info.abas.join(', '))} &nbsp;|&nbsp; <b>${S.carteira.info.registros.toLocaleString('pt-BR')}</b> registros &nbsp;|&nbsp; <b>${S.carteira.info.bandeiras}</b> bandeiras<br>Estrutura validada. Clique em <b>Continuar</b> para ver o dashboard. Para trocar de arquivo, carregue outro no quadro acima.</div>` : '';
    return `<div class="hero"><div class="stp"><span>1</span><div><b>Carregue aqui a planilha da Carteira</b><small>Arraste o arquivo Excel para o quadro abaixo ou clique no botão roxo.</small></div></div>${dropHtml('dzb', '📥', 'Arraste a planilha Excel para cá', 'Formato aceito: .xlsx (Excel)', '📂 Escolher arquivo no computador')}${info}</div>
    <div class="how"><h3>📋 Como a planilha precisa ser</h3><p>O Excel precisa ter uma aba chamada <b>CARTEIRA</b> e, na <b>primeira linha</b>, os títulos das colunas mostrados abaixo. Pode ter outras abas — o sistema acha a CARTEIRA sozinho. Use sempre a versão mais atual.</p>
    <div class="legend"><span><i class="lg req"></i> Coluna <b>obrigatória</b> para esta automação</span><span><i class="lg opt"></i> Opcional (alimenta o dashboard)</span></div>
    <div class="tw"><table class="ex"><tr>${cols.map(c => `<th class="${c[1] ? 'req' : 'opt'}">${c[0]}${c[1] ? ' ★' : ''}</th>`).join('')}</tr><tr>${['NOME FICTÍCIO', 'REDE EXEMPLO', 'BANDEIRA EXEMPLO', '99', 'DIRETORIA EXEMPLO', '60', '100001', 'CLIENTE EXEMPLO LTDA', '81', 'C45', 'AA', 'GRUPO EXEMPLO', 'SP', 'EXECUTIVO EXEMPLO', 'GERENTE EXEMPLO'].map(x => `<td>${x}</td>`).join('')}</tr></table></div><p class="dica">Exemplo totalmente fictício, só para mostrar os títulos esperados.</p></div>`;
  }
  function b1() { dz('dzb', carregarBase); }

  /* Etapa 2 — dashboard só com dados da planilha */
  const cnt = (rows, f) => { const m = {}; rows.forEach(r => { const k = f(r) || '(vazio)'; m[k] = (m[k] || 0) + 1; }); return Object.entries(m).sort((a, b) => b[1] - a[1]); };
  const bars = (l, c, n = 10) => { const t = l.slice(0, n), mx = Math.max(1, ...t.map(x => x[1])); return t.map(x => `<div class="bar ${c}"><span title="${esc(x[0])}">${esc(x[0])}</span><div data-w="${Math.max(3, x[1] / mx * 100).toFixed(1)}"></div><b>${x[1]}</b></div>`).join(''); };
  function s2() {
    const R = S.carteira.rows, d = new Set(R.map(r => r.cliente).filter(Boolean)), ds = f => new Set(R.map(f).filter(Boolean)).size;
    const inc = R.filter(r => !r.cliente || !r.bandeira || (S.tipo === 'ZDEF' && !r.cond)).length, dup = cnt(R.filter(r => r.cliente), r => r.cliente).filter(x => x[1] > 1);
    const k = [['Registros', R.length, '📄'], ['Clientes distintos', d.size, '👤'], ['Bandeiras', S.carteira.bandeiras.length, '🏷️'], ['Redes', ds(r => r.rede), '🔗'], ['Gerentes regionais', ds(r => r.gerente), '🧑‍💼'], ['Registros incompletos', inc, '⚠️']];
    const q = [['Código Cliente vazio', R.filter(r => !r.cliente).length], ['Bandeira vazia', R.filter(r => !r.bandeira).length], ['Condição de Pagamento vazia', R.filter(r => !r.cond).length], ['Grupo de Preço vazio (permitido)', R.filter(r => !r.gp).length], ['Códigos de cliente duplicados', dup.length]];
    const pct = Math.round((R.length - inc) / R.length * 100), ufs = [...new Set(R.map(r => r.uf).filter(Boolean))].sort();
    return `<p class="dica">Resumo do que existe na planilha carregada. Confira antes de continuar.</p><div class="kpis">${k.map((x, i) => `<div class="kpi k${i + 1}"><i>${x[2]}</i><b>${x[1].toLocaleString('pt-BR')}</b>${x[0]}</div>`).join('')}</div><div class="charts">
    <div class="ch"><h4>🏷️ Clientes por bandeira (top 10)</h4>${bars(cnt(R, r => r.bandeira), 'c1')}</div><div class="ch"><h4>📍 Clientes por estado (top 10)</h4>${bars(cnt(R, r => r.uf), 'c2')}</div>
    <div class="ch"><h4>🧑‍💼 Clientes por gerente regional (top 10)</h4>${bars(cnt(R, r => r.gerente), 'c3')}</div>
    <div class="ch"><h4>✅ Qualidade da base</h4><div class="donut" data-bg="conic-gradient(#059669 ${pct}%, #fecaca 0)"><b>${pct}%</b></div><p class="dica" style="text-align:center">dos registros completos</p><table>${q.map(x => `<tr><td>${x[0]}</td><td><b>${x[1]}</b></td></tr>`).join('')}</table>${dup.length ? `<p class="dica">Duplicados (primeiros): ${esc(dup.slice(0, 8).map(x => x[0]).join(', '))}</p>` : ''}</div></div>
    <h3>🔎 Consultar registros</h3><div class="row"><input id="q" placeholder="Pesquisar por bandeira, código, razão social…"><select id="fuf"><option value="">Todos os estados</option>${ufs.map(u => `<option>${esc(u)}</option>`).join('')}</select></div><div class="tw" id="tb"></div>`.replace(' style="text-align:center"', '');
  }
  function b2() {
    const f = () => { const q = C.N($('#q').value), u = $('#fuf').value, L = S.carteira.rows.filter(r => (!u || r.uf === u) && (!q || C.N([r.bandeira, r.cliente, r.razao, r.rede, r.gerente, r.uf].join(' ')).includes(q)));
      $('#tb').innerHTML = `<table><tr><th>Bandeira</th><th>Código</th><th>Razão Social</th><th>UF</th><th>Cond. Pgto</th><th>Grupo Preço</th><th>Gerente</th></tr>${L.slice(0, 200).map(r => `<tr><td>${esc(r.bandeira)}</td><td>${esc(r.cliente)}</td><td>${esc(r.razao)}</td><td>${esc(r.uf)}</td><td>${esc(r.cond)}</td><td>${esc(r.gp)}</td><td>${esc(r.gerente)}</td></tr>`).join('')}</table><p class="dica" style="padding:0 12px">${L.length} registro(s); exibindo até 200.</p>`.replace(' style="padding:0 12px"', ''); };
    $('#q').oninput = f; $('#fuf').onchange = f; f();
  }

  /* Etapa 3 */
  function s3() { return `<div class="stp"><span>3</span><div><b>Escolha a bandeira</b><small>Clique no campo para ver todas as bandeiras ou digite para pesquisar. Para trocar, é só clicar de novo.</small></div></div><div class="combo"><input id="band" value="${esc(S.band)}" placeholder="Clique aqui ou digite o nome da bandeira…" autocomplete="off"><button class="arr" id="arr" type="button">▾</button><div class="dd" id="bl" hidden></div></div><div id="res"></div>`; }
  function resumo() {
    if (!S.band) return $('#res').innerHTML = '';
    const inc = S.rows.filter(r => !r.cliente || (S.tipo === 'ZDEF' && !r.cond)).length;
    $('#res').innerHTML = `<h3>🏷️ ${esc(S.band)} — ${S.rows.length} registro(s) encontrados</h3>${inc ? `<div class="warn">${inc} registro(s) incompleto(s) (sem código ou sem condição de pagamento). Eles serão apontados como erro na Etapa 4.</div>` : ''}<div class="tw"><table><tr><th>Código</th><th>Razão Social</th><th>Canal</th><th>Cond. Pgto</th><th>Grupo Preço</th></tr>${S.rows.map(r => `<tr><td>${esc(r.cliente || '—')}</td><td>${esc(r.razao)}</td><td>${esc(S.cfg.canal)}</td><td>${esc(r.cond || '—')}</td><td>${esc(r.gp || '—')}</td></tr>`).join('')}</table></div>`;
  }
  function b3() {
    resumo(); const inp = $('#band'), box = $('#bl');
    const lista_ = all => { const q = all ? '' : C.N(inp.value), l = S.carteira.bandeiras.filter(b => !q || C.N(b).includes(q)).slice(0, 100); box.innerHTML = l.map(b => `<div data-b="${esc(b)}">${esc(b)}</div>`).join('') || '<div class="none">Nenhuma bandeira encontrada com esse texto</div>'; box.hidden = false; };
    const pick = b => { if (b !== S.band) { S.band = b; S.rows = S.carteira.rows.filter(r => r.bandeira === b).map(r => ({ ...r, dias: '', desde: '', ate: '', montante: '' })); S.recFor = ''; } inp.value = b; box.hidden = true; clr(); resumo(); preview(); };
    inp.onfocus = () => { inp.select(); lista_(true); }; inp.onclick = () => lista_(true); $('#arr').onclick = () => { inp.focus(); lista_(true); };
    inp.oninput = () => { lista_(false); const b = S.carteira.bandeiras.find(x => C.N(x) === C.N(inp.value)); if (b) pick(b); else { S.band = ''; S.rows = []; $('#res').innerHTML = inp.value ? '<div class="warn">Bandeira não encontrada na base carregada. Verifique a escrita ou escolha uma opção da lista.</div>' : ''; } };
    box.onmousedown = e => { const d = e.target.closest('[data-b]'); if (d) { e.preventDefault(); pick(d.dataset.b); } };
    inp.onblur = () => setTimeout(() => box.hidden = true, 150);
  }

  /* Etapa 4 */
  const MF = () => S.tipo === 'ZDEF' ? { f: ['dias', 'desde', 'ate'], h: ['Dias adicionais', 'Válido desde', 'Válido até'], t: ['number', 'date', 'date'] } : { f: ['montante', 'desde', 'ate'], h: ['Montante (desconto)', 'Válido desde', 'Válido até'], t: ['text', 'date', 'date'] };
  function s4() {
    const m = MF();
    return `<p class="dica">${S.tipo === 'ZDEF' ? '' : 'Informe o desconto como número (ex.: 2,5); é gravado exatamente como digitado. Confirme a unidade esperada pelo sistema de destino. '}Use "Aplicar para todos" e depois ajuste linhas individuais.</p>
    <div class="grid">${m.f.map((k, n) => `<label>${m.h[n]} (todos)<input type="${m.t[n]}" id="all_${k}"></label>`).join('')}<label>&nbsp;<button class="sec" id="ap">Aplicar para todos</button></label></div>
    <div class="tw"><table><tr><th>Cliente</th><th>Razão Social</th>${S.tipo === 'ZDEF' ? '<th>Cond.</th>' : ''}<th>Grupo Preço</th>${m.h.map(x => `<th>${x}</th>`).join('')}</tr>${S.rows.map((r, i) => `<tr><td>${esc(r.cliente)}</td><td>${esc(r.razao)}</td>${S.tipo === 'ZDEF' ? `<td>${esc(r.cond)}</td>` : ''}<td>${esc(r.gp || '*')}</td>${m.f.map((k, n) => `<td><input type="${m.t[n]}" data-i="${i}" data-f="${k}" value="${esc(r[k])}"></td>`).join('')}</tr>`).join('')}</table></div>
    <details><summary>Regras de preenchimento (pendentes de aprovação do negócio)</summary><div class="grid">${[['divisao', 'Divisão de Negócio'], ['canal', 'Canal'], ['grupoClientes', 'Grupo de Clientes'], ['setor', 'Setor de Atividade']].map(x => `<label>${x[1]}<input data-c="${x[0]}" value="${esc(S.cfg[x[0]])}"></label>`).join('')}</div></details><button id="vl" class="sec">Validar</button>`;
  }
  function b4() {
    const m = MF();
    $('#body').oninput = e => { const t = e.target; if (t.dataset.f) S.rows[t.dataset.i][t.dataset.f] = t.value; if (t.dataset.c) S.cfg[t.dataset.c] = t.value; };
    $('#ap').onclick = () => { m.f.forEach(k => { const v = $('#all_' + k).value; if (v !== '') S.rows.forEach(r => r[k] = v); }); render(); };
    $('#vl').onclick = () => { clr(); const v = C.validar({ ...curJob(), mail: undefined }); if (v.erros.length) $('#msg').appendChild(lista({ t: `${v.erros.length} erro(s) impedem a geração:`, l: v.erros }, 'err')); else msg('Nenhum erro encontrado.', 'ok'); if (v.avisos.length) $('#msg').appendChild(lista({ t: `${v.avisos.length} alerta(s) (permitem continuar):`, l: v.avisos }, 'warn')); };
  }

  /* Etapa 5 */
  function s5() {
    const st = S.dest ? `<div class="ok">✅ <b>${esc(S.destArq.nome)}</b> anexada (${S.destArq.n} linhas). Se precisar, anexe outra para substituir.</div>` : '';
    return `<div class="stp"><span>5</span><div><b>Quem vai receber o e-mail?</b><small>Siga os 3 passos abaixo. O sistema usa só as pessoas ativas da automação <b>${S.tipo}</b> (ou AMBAS) e da bandeira <b>${esc(S.band)}</b>.</small></div></div>
    <div class="s3"><div><div class="stp"><span>1</span><b>Baixe o modelo</b></div><p class="dica">Planilha de exemplo com as colunas certas.</p><button class="sec" id="mod">⬇ Baixar modelo (Excel)</button></div>
    <div><div class="stp"><span>2</span><b>Preencha</b></div><p class="dica">Uma pessoa por linha: Automação (ZDEF, ZT31 ou AMBAS), Bandeira, Nome, Email, Tipo (PARA ou CC) e Ativo (SIM ou NÃO). Salve no computador.</p></div>
    <div><div class="stp"><span>3</span><b>Anexe aqui</b></div><p class="dica">Volte e envie a planilha preenchida no quadro abaixo.</p></div></div>
    ${dropHtml('dzd', '📎', 'Anexar a planilha de destinatários', 'Arraste o arquivo preenchido para cá (.xlsx)', '📂 Escolher planilha de e-mails')}${st}
    <h3>👥 Quem vai receber</h3>${S.dest && !S.rec.length ? '<div class="warn">Nenhum destinatário ativo encontrado para esta automação e bandeira. Adicione manualmente abaixo ou corrija a planilha e anexe de novo.</div>' : ''}
    <div class="tw"><table id="rt"><tr><th>Nome</th><th>E-mail</th><th>Tipo</th><th></th></tr>${S.rec.map((r, i) => `<tr><td><input data-r="${i}" data-k="nome" value="${esc(r.nome)}"></td><td><input data-r="${i}" data-k="email" value="${esc(r.email)}"></td><td><select data-r="${i}" data-k="tipo"><option${r.tipo === 'PARA' ? ' selected' : ''}>PARA</option><option${r.tipo === 'CC' ? ' selected' : ''}>CC</option></select></td><td><button class="sec" data-del="${i}">Remover</button></td></tr>`).join('') || '<tr><td colspan="4" class="dica">Nenhum destinatário ainda.</td></tr>'}</table></div><button class="sec" id="addr">+ Adicionar destinatário manualmente</button>`;
  }
  function b5() {
    dz('dzd', carregarDest);
    $('#mod').onclick = async () => dl(new Blob([await C.modeloDest()]), 'Modelo_Base_Destinatarios.xlsx');
    $('#body').oninput = e => { const t = e.target; if (t.dataset.k) { S.rec[t.dataset.r][t.dataset.k] = t.value; preview(); } };
    $('#body').onchange = e => { if (e.target.dataset.k === 'tipo') preview(); };
    $('#body').onclick = e => { if (e.target.dataset.del !== undefined) { S.rec.splice(+e.target.dataset.del, 1); render(); } };
    $('#addr').onclick = () => { S.rec.push({ nome: '', email: '', tipo: S.rec.some(r => r.tipo === 'PARA') ? 'CC' : 'PARA' }); render(); };
  }

  /* Etapa 6 */
  function s6() {
    return `<p>Revise e edite o e-mail. A pré-visualização à direita é atualizada ao digitar. Use {BANDEIRA}${S.tipo === 'ZT31' ? ' e {PERCENTUAL}' : ''} no texto.</p><label>Assunto<input data-m="assunto" value="${esc(S.mail.assunto)}"></label><label>Saudação<input data-m="saud" value="${esc(S.mail.saud)}"></label><label>Corpo da mensagem<textarea data-m="corpo" rows="7">${esc(S.mail.corpo)}</textarea></label><label>Assinatura<textarea data-m="assinatura" rows="5">${esc(S.mail.assinatura)}</textarea></label>`;
  }
  function b6() { $('#body').oninput = e => { if (e.target.dataset.m) { S.mail[e.target.dataset.m] = e.target.value; preview(); } }; }

  /* Painel lateral de e-mail */
  function preview() {
    if (S.step < 4 || !S.band) return;
    const j = curJob(), m = C.montarEmail(j), ch = t => S.rec.filter(r => r.tipo === t && r.email).map(r => `<span class="chip">${esc(r.email)}</span>`).join('') || '<i>(vazio)</i>';
    const cols = C.COLS[S.tipo].map(c => c[0]), lin = j.rows.map(r => C.linhaReport(S.tipo, j, r));
    $('#aside').innerHTML = `<h3>✉️ Pré-visualização do e-mail</h3><div class="ol"><div class="bar0">Nova mensagem — ${S.tipo} / ${esc(S.band)}</div>
    <div class="f"><span>Para</span><b>${ch('PARA')}</b></div><div class="f"><span>Cc</span><b>${ch('CC')}</b></div><div class="f"><span>Assunto</span><b>${esc(m.assunto)}</b></div><div class="tx">${esc(m.corpo)}</div>
    <div class="at"><p><span class="fl">📎 ${C.nomeArq(j)}.xlsx</span></p><p class="dica"><b>Esta é a tabela que está anexada ao e-mail</b> (${lin.length} linha(s)):</p><div class="tw"><table><tr>${cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr>${lin.map(r => `<tr>${r.map(v => `<td>${esc(v == null ? '' : v)}</td>`).join('')}</tr>`).join('')}</table></div></div></div>`;
  }

  /* Etapa 7 */
  function s7() {
    return `<p>Tudo pronto para <b>${S.tipo} / ${esc(S.band)}</b> (${S.rows.length} cliente(s)).</p>${S.jobs.length ? `<h4>Já no lote</h4>${S.jobs.map((j, i) => `<details><summary>${i + 1}. ${j.tipo} — ${esc(j.bandeira)} (${j.rows.length} clientes)</summary><div class="mail"><b>Para:</b> ${esc(j.mail.para)}\n<b>Cc:</b> ${esc(j.mail.cc)}\n\n${esc(C.montarEmail(j).corpo)}</div></details>`).join('')}` : ''}
    <div class="row"><button class="sec" id="more">+ Adicionar outra bandeira ao lote</button><button id="gen">Gerar relatórios e rascunhos</button></div><div id="dlz"></div>`;
  }
  const addJob = () => { const j = curJob(); S.jobs = S.jobs.filter(x => !(x.tipo === j.tipo && x.bandeira === j.bandeira)); S.jobs.push(j); };
  function b7() {
    $('#more').onclick = () => { addJob(); S.band = ''; S.rows = []; S.rec = []; S.recFor = ''; go(3); msg('Bandeira adicionada ao lote. Escolha a próxima.', 'ok'); };
    $('#gen').onclick = gerar;
  }
  function dl(blob, nome) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 5000); }
  async function gerar() {
    clr(); addJob();
    if (!confirm(`Gerar ${S.jobs.length} relatório(s) e rascunho(s)?\nNenhum e-mail será enviado automaticamente.`)) return;
    const zip = new JSZip(), R = zip.folder('RELATORIOS'), D = zip.folder('RASCUNHOS'), log = [], falhas = []; let total = 0;
    for (const j of S.jobs) {
      const n = C.nomeArq(j);
      try { const x = await C.gerarExcel(j); R.file(n + '.xlsx', x); total += j.rows.length; log.push(`Relatório ${n}.xlsx: ${j.rows.length} cliente(s)`);
        try { D.file('EMAIL_' + n + '.eml', C.gerarEml(j, x)); log.push(`Rascunho EMAIL_${n}.eml (Para: ${j.mail.para}${j.mail.cc ? ' | Cc: ' + j.mail.cc : ''})`); } catch (e) { falhas.push(`Não foi possível preparar o rascunho da bandeira ${j.bandeira}.\nO relatório foi gerado, mas o rascunho não foi concluído. Verifique o endereço do destinatário e tente novamente.`); }
      } catch (e) { falhas.push(`Falha ao gerar o Excel da bandeira ${j.bandeira} (${j.tipo}): ${e.message}`); }
    }
    falhas.forEach(f => msg(f));
    zip.file('RESUMO.txt', `Processamento em ${new Date().toLocaleString('pt-BR')}\nRegistros lidos na Carteira: ${S.carteira.info.registros}\nBandeiras: ${[...new Set(S.jobs.map(j => j.bandeira))].join(', ')}\nClientes incluídos: ${total}\n\n${log.join('\n')}\n\nAlertas/falhas: ${falhas.length ? '\n' + falhas.join('\n') : 'nenhum'}\nNenhum e-mail foi enviado automaticamente.\n`);
    try { const blob = await zip.generateAsync({ type: 'blob' }); $('#dlz').innerHTML = '<button id="dl2">⬇ Baixar todos os arquivos</button>'; $('#dl2').onclick = () => dl(blob, 'RESULTADO_AUTOMACAO.zip'); msg(`Arquivos gerados: ${S.jobs.length} relatório(s) e rascunho(s). Clique em "Baixar todos os arquivos".`, 'ok'); }
    catch (e) { msg('Não foi possível gerar o ZIP (possível falta de memória). Processe menos bandeiras por vez.'); }
  }
  $('#boot').hidden = true; view('home');
})();
