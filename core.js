/* Núcleo: leitura da Carteira, validações, geração do Excel e do .eml. Tudo local no navegador. */
(function (root) {
  const N = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
  const REQ = {
    bandeira: ['bandeira'], cliente: ['codigo cliente', 'cod cliente', 'codigo do cliente'],
    gp: ['grupo de preco'], cond: ['condicao de pagamento', 'condicao pagamento', 'cond pagamento']
  };
  const NOME = { bandeira: 'Bandeira', cliente: 'Código Cliente', gp: 'Grupo de Preço', cond: 'Condição de Pagamento' };

  function val(c) {
    let v = c.value;
    if (v instanceof Date) return v;
    if (v && typeof v === 'object') {
      if ('result' in v) v = v.result;
      else if (v.richText) v = v.richText.map(t => t.text).join('');
      else if (v.text) v = v.text; else v = '';
    }
    return v == null ? '' : v;
  }
  const txt = v => (v instanceof Date ? '' : String(v)).trim();

  /* Lê a aba CARTEIRA. Retorna {ok, erro, rows, info}. */
  function parseCarteira(wb, tipo) {
    const nomes = wb.worksheets.map(w => w.name);
    const ws = wb.worksheets.find(w => N(w.name) === 'carteira');
    if (!ws) return { ok: false, erro: `A aba CARTEIRA não foi encontrada neste arquivo.\n\nAbas encontradas: ${nomes.join(', ') || '(nenhuma)'}.\n\nConfira se a aba existe e se o nome está correto, ou carregue a Carteira em um arquivo separado.` };
    let hr = 0, map = {}, hdrs = [];
    for (let r = 1; r <= Math.min(5, ws.rowCount); r++) {
      const names = []; ws.getRow(r).eachCell({ includeEmpty: false }, (c, i) => names[i] = N(val(c)));
      if (names.includes('bandeira')) { hr = r; hdrs = names; break; }
    }
    if (!hr) return { ok: false, erro: `Não foi possível localizar a linha de cabeçalho da aba CARTEIRA (a coluna "Bandeira" não foi encontrada nas primeiras linhas).\n\nVerifique se os títulos das colunas estão na primeira linha da aba.` };
    const achadas = ws.getRow(hr).values.filter(v => v != null).map(v => txt(val({ value: v })));
    for (const k in REQ) {
      const idx = hdrs.map((h, i) => REQ[k].includes(h) ? i : -1).filter(i => i > 0);
      if (!idx.length && (k === 'gp' || (k === 'cond' && tipo !== 'ZDEF'))) { map[k] = 0; continue; } // campo não usado/opcional
      if (!idx.length) return { ok: false, erro: `Não foi possível continuar porque a coluna "${NOME[k]}" não foi localizada na aba CARTEIRA.\n\nVerifique se essa coluna existe e se o nome está correto.\n\nColunas identificadas: ${achadas.join(', ')}.\n\nCorrija a planilha e carregue novamente.` };
      if (idx.length > 1) return { ok: false, erro: `A coluna "${NOME[k]}" aparece mais de uma vez na aba CARTEIRA (colunas ${idx.join(', ')}).\n\nRemova a coluna duplicada e carregue novamente, para evitar o uso de dados errados.` };
      map[k] = idx[0];
    }
    const fi = n => Math.max(0, hdrs.findIndex(h => h === n)), rzIdx = hdrs.findIndex(h => h === 'razao social cliente'), rdI = fi('rede'), ufI = fi('uf'), gtI = fi('nome gerente regional');
    const rows = [];
    for (let r = hr + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const g = i => i > 0 ? val(row.getCell(i)) : '';
      const o = { linha: r, bandeira: txt(g(map.bandeira)), cliente: txt(g(map.cliente)), gp: txt(g(map.gp)), cond: txt(g(map.cond)), razao: txt(g(rzIdx)), rede: txt(g(rdI)), uf: txt(g(ufI)), gerente: txt(g(gtI)) };
      if (!o.bandeira && !o.cliente && !o.gp && !o.cond && !o.razao && !o.uf) continue; // linha vazia
      rows.push(o);
    }
    if (!rows.length) return { ok: false, erro: 'A aba CARTEIRA está vazia (nenhum registro abaixo do cabeçalho).' };
    const band = [...new Set(rows.map(r => r.bandeira).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    return { ok: true, rows, bandeiras: band, info: { abas: nomes, registros: rows.length, bandeiras: band.length } };
  }

  /* Detecta se a primeira aba é modelo ZDEF ou ZT31 */
  function detectModel(wb) {
    const ws = wb.worksheets.find(w => N(w.name) !== 'carteira' && N(w.name) !== 'email');
    if (!ws) return null;
    for (let r = 1; r <= 3; r++) {
      const h = []; ws.getRow(r).eachCell(c => h.push(N(val(c))));
      if (h.includes('componente politica')) return 'ZT31';
      if (h.includes('cond/pgto/')) return 'ZDEF';
    }
    return null;
  }
  function readEmailSheet(wb) {
    const ws = wb.worksheets.find(w => N(w.name) === 'email'); const o = {};
    if (!ws) return o;
    ws.eachRow(row => row.eachCell((c, i) => { const k = N(val(c)); if (k === 'para' || k === 'cc') o[k] = txt(val(row.getCell(i + 1))); }));
    return o;
  }

  const COLS = {
    ZDEF: [['Divisao de Negocio', 21.3], ['Canal', 7.4], ['Grupo de Clientes', 19.6], ['Grupo de Preço', 16.7], ['Cliente', 11.1], ['Setor de Atividade', 20.1], ['Sub/Agr/S&OP', 15.3], ['CategComercial', 18.3], ['Material', 18.9], ['Cond/pgto/', 12.6], ['Dias valor adicions/', 21.4], ['Data efet/fixa', 15.1], ['Valido Desde', 14.7], ['Valido Até', 12.9], ['Bandeira', 31.9]],
    ZT31: [['Divisao de Negocio', 21.3], ['Componente Politica', 23.4], ['Canal', 7.4], ['Grupo de Clientes', 19.6], ['Grupo de Preço', 16.7], ['Cliente', 10.3], ['Setor de Atividade', 20.1], ['Sub.Agr.S&OP', 18.6], ['Material', 9.9], ['Montante', 11.1], ['Valido Desde', 14.7], ['Valido Até', 12], ['Bandeira', 18.9]]
  };
  const num = s => /^[1-9]\d*$/.test(String(s)) ? Number(s) : String(s); // preserva zeros à esquerda
  const fmtData = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? `${m[3]}.${m[2]}.${m[1]}` : ''; };
  const pnum = s => { const t = String(s).trim().replace(/\s/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'); return t !== '' && /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : NaN; };

  function linhaReport(tipo, j, r) {
    const c = j.cfg, gp = r.gp || '*';
    if (tipo === 'ZDEF') return [num(c.divisao), num(c.canal), c.grupoClientes, gp, num(r.cliente), c.setor, '*', '*', '*', r.cond, pnum(r.dias), '*', fmtData(r.desde), fmtData(r.ate), j.bandeira];
    return [num(c.divisao), 'ZT31', num(c.canal), c.grupoClientes, gp, num(r.cliente), c.setor, '*', '*', r.montante === '' ? null : pnum(r.montante), fmtData(r.desde) || null, fmtData(r.ate) || null, j.bandeira];
  }

  /* Validação: retorna {erros:[], avisos:[]} */
  function validar(j) {
    const E = [], A = [], T = j.tipo, seen = {};
    if (!j.rows.length) E.push('Nenhum cliente foi selecionado. Marque pelo menos um cliente na etapa 3.');
    for (const k of ['divisao', 'canal', 'grupoClientes', 'setor']) if (!String(j.cfg[k]).trim()) E.push(`A regra "${k}" está vazia. Informe um valor (ou * quando não se aplica).`);
    j.rows.forEach(r => {
      const id = r.cliente || `(sem código, linha ${r.linha} da CARTEIRA)`;
      if (!r.cliente) E.push(`A linha ${r.linha} da aba CARTEIRA está vinculada à bandeira ${j.bandeira}, mas não possui Código Cliente.\nPreencha o código ou corrija a base antes de continuar.`);
      if (seen[r.cliente]) A.push(`O código ${id} aparece mais de uma vez na bandeira ${j.bandeira}. Confirme se é duplicidade real; nada foi removido automaticamente.`);
      seen[r.cliente] = 1;
      if (T === 'ZDEF') {
        if (!r.cond) E.push(`O cliente de código ${id} não possui Condição de Pagamento cadastrada na Carteira.\nEsse campo é necessário para gerar o relatório ZDEF. Corrija a base e tente novamente.`);
        const d = pnum(r.dias);
        if (r.dias === '' || isNaN(d) || d <= 0 || !Number.isInteger(d)) E.push(`Dias adicionais do cliente ${id}: informe um número inteiro maior que zero (valor encontrado: "${r.dias || 'vazio'}").`);
        if (!r.desde) E.push(`Preencha a data "Válido desde" do cliente ${id}.`);
        if (!r.ate) E.push(`Preencha a data "Válido até" do cliente ${id}.`);
      } else {
        if (isNaN(pnum(r.montante)) || pnum(r.montante) <= 0) E.push(`O desconto informado para o cliente ${id} não está no formato esperado (valor encontrado: "${r.montante || 'vazio'}").\nUse apenas números, por exemplo 2,5. Confira a regra de preenchimento do campo Montante.`);
        if (!r.desde || !r.ate) A.push(`O cliente ${id} está sem "Válido desde" e/ou "Válido até". O modelo permite deixar em branco; confirme se está de acordo com a regra aprovada.`);
      }
      if (r.desde && r.ate && r.ate < r.desde) E.push(`A data "Válido até" do cliente ${id} é anterior à data "Válido desde". Revise as datas informadas.`);
      if (!r.gp) A.push(`O cliente ${id} não tem Grupo de Preço na Carteira; será preenchido com *.`);
    });
    const m = j.mail; if (!m) return { erros: E, avisos: A };
    const bad = l => String(l).split(/[;,]/).map(s => s.trim()).filter(Boolean).filter(e => !/^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/.test(e));
    if (!String(m.para).trim()) E.push(`Preencha o destinatário (Para) do e-mail da bandeira ${j.bandeira}.`);
    [['Para', m.para], ['Cc', m.cc]].forEach(([n, v]) => { const b = bad(v); if (b.length) E.push(`Endereço(s) inválido(s) no campo ${n}: ${b.join(', ')}.\nVerifique o endereço do destinatário e tente novamente.`); });
    return { erros: E, avisos: A };
  }

  function percentual(j) {
    const v = [...new Set(j.rows.map(r => pnum(r.montante)))];
    return v.length === 1 ? `de ${String(v[0]).replace('.', ',')}% ` : '(os percentuais estão detalhados por cliente na planilha anexa) ';
  }
  function montarEmail(j) {
    const sub = j.mail.assunto.replace(/\{BANDEIRA\}/g, j.bandeira);
    let corpo = j.mail.corpo.replace(/\{BANDEIRA\}/g, j.bandeira).replace(/\{PERCENTUAL\}/g, j.tipo === 'ZT31' ? percentual(j) : '');
    return { assunto: sub, corpo: [j.mail.saud, corpo, j.mail.assinatura].filter(x => String(x || '').trim()).join('\n\n') };
  }
  const nomeArq = j => `${j.tipo}_${j.bandeira.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9]+/g, '_').replace(/^_|_$/g, '')}`;

  async function gerarExcel(j) {
    const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Sheet1'), cols = COLS[j.tipo];
    ws.columns = cols.map(c => ({ width: c[1] }));
    ws.getRow(1).values = cols.slice(0, j.tipo === 'ZDEF' ? 14 : 12).map((_, i) => i + 1);
    ws.getRow(2).values = cols.map(c => c[0]);
    ws.getRow(2).eachCell(c => { c.font = { bold: true }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } }; c.alignment = { horizontal: 'center' }; });
    j.rows.forEach((r, i) => {
      const row = ws.getRow(3 + i); row.values = linhaReport(j.tipo, j, r);
      if (j.tipo === 'ZT31') { row.getCell(3).numFmt = '00'; row.getCell(10).numFmt = '0.000'; }
    });
    return new Uint8Array(await wb.xlsx.writeBuffer());
  }

  function b64(u8) {
    if (typeof Buffer !== 'undefined') return Buffer.from(u8).toString('base64');
    let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s);
  }
  const enc = s => b64(new TextEncoder().encode(s));
  const wrap = s => s.replace(/(.{76})/g, '$1\r\n');
  function gerarEml(j, xlsx) {
    const m = montarEmail(j), bd = '----=_Part_' + Date.now().toString(36), fn = nomeArq(j) + '.xlsx';
    const h = [`X-Unsent: 1`, `To: ${j.mail.para.replace(/,/g, ';')}`];
    if (j.mail.cc.trim()) h.push(`Cc: ${j.mail.cc.replace(/,/g, ';')}`);
    h.push(`Subject: =?UTF-8?B?${enc(m.assunto)}?=`, 'MIME-Version: 1.0', `Content-Type: multipart/mixed; boundary="${bd}"`, '');
    h.push(`--${bd}`, 'Content-Type: text/plain; charset=utf-8', 'Content-Transfer-Encoding: base64', '', wrap(enc(m.corpo.replace(/\n/g, '\r\n'))));
    h.push(`--${bd}`, `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name="${fn}"`, 'Content-Transfer-Encoding: base64', `Content-Disposition: attachment; filename="${fn}"`, '', wrap(b64(xlsx)), `--${bd}--`, '');
    return h.join('\r\n');
  }

  /* Base de Destinatários */
  const DCOLS = ['Automação', 'Bandeira', 'Nome', 'Email', 'Tipo', 'Ativo'];
  function parseDest(wb) {
    const ws = wb.worksheets[0]; if (!ws) return { ok: false, erro: 'O arquivo não possui nenhuma aba.' };
    const h = []; ws.getRow(1).eachCell((c, i) => h[i] = N(val(c)));
    const need = DCOLS.map(c => [c, h.indexOf(N(c))]), miss = need.filter(x => x[1] < 1).map(x => x[0]);
    if (miss.length) return { ok: false, erro: `Não foi possível continuar porque a(s) coluna(s) "${miss.join('", "')}" não foi(ram) localizada(s) na aba "${ws.name}" da Base de Destinatários.\n\nColunas encontradas: ${h.filter(Boolean).join(', ') || '(nenhuma)'}.\nColunas esperadas: ${DCOLS.join(', ')}.\n\nBaixe o modelo na tela, ajuste a planilha e carregue novamente.` };
    const ix = Object.fromEntries(need), rows = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const g = c => txt(val(ws.getRow(r).getCell(ix[c])));
      const o = { linha: r, aut: N(g('Automação')), band: g('Bandeira'), nome: g('Nome'), email: g('Email'), tipo: N(g('Tipo')), ativo: N(g('Ativo')) };
      if (o.band || o.email || o.nome) rows.push(o);
    }
    return { ok: true, rows };
  }
  function destinatarios(rows, tipo, band) {
    const seen = new Set(), out = [];
    const f = rows.filter(r => r.ativo === 'sim' && N(r.band) === N(band) && (r.aut === N(tipo) || r.aut === 'ambas'));
    for (const t of ['para', 'cc']) f.filter(r => r.tipo === t).forEach(r => { const k = r.email.toLowerCase(); if (!k || seen.has(k)) return; seen.add(k); out.push({ nome: r.nome, email: r.email, tipo: t.toUpperCase() }); });
    return out;
  }
  async function modeloDest() {
    const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Destinatarios');
    ws.columns = DCOLS.map((c, i) => ({ header: c, width: [12, 24, 22, 34, 8, 8][i] }));
    ws.getRow(1).eachCell(c => { c.font = { bold: true }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } }; });
    const e = 'luma.sampaio.dados@gmail.com', d = [['ZDEF', 'BRAIR', 'Pessoa Teste 1', 'PARA'], ['ZDEF', 'BRAIR', 'Pessoa Teste 2', 'CC'], ['ZT31', 'BRAIR', 'Pessoa Teste 1', 'PARA'], ['AMBAS', 'BANDEIRA EXEMPLO A', 'Pessoa Teste 3', 'PARA'], ['AMBAS', 'BANDEIRA EXEMPLO A', 'Pessoa Teste 4', 'CC'], ['ZDEF', 'BANDEIRA EXEMPLO B', 'Pessoa Teste 5', 'PARA'], ['ZT31', 'BANDEIRA EXEMPLO B', 'Pessoa Teste 6', 'PARA'], ['ZT31', 'BANDEIRA EXEMPLO B', 'Pessoa Teste 7', 'CC'], ['AMBAS', 'BANDEIRA EXEMPLO C', 'Pessoa Teste 8', 'PARA'], ['AMBAS', 'BANDEIRA EXEMPLO C', 'Pessoa Teste 9', 'CC'], ['ZDEF', 'BANDEIRA EXEMPLO C', 'Pessoa Inativa', 'PARA']];
    d.forEach((r, i) => ws.addRow([r[0], r[1], r[2], e, r[3], i === d.length - 1 ? 'NÃO' : 'SIM']));
    return new Uint8Array(await wb.xlsx.writeBuffer());
  }
  const api = { COLS, linhaReport, parseDest, destinatarios, modeloDest, N, parseCarteira, detectModel, readEmailSheet, validar, montarEmail, gerarExcel, gerarEml, nomeArq, pnum };
  if (typeof module !== 'undefined') module.exports = api; else root.Core = api;
})(this);
