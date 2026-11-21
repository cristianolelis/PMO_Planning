/* =====================================================================
 * Análise de Documentação de Referência → EAP (boas práticas PMBOK)
 * ---------------------------------------------------------------------
 * Lê propostas técnicas, contratos, termos de referência e editais
 * (PDF, DOCX, TXT/MD), extrai o escopo e monta uma EAP orientada a
 * entregas, com dicionário, marcos, durações e predecessoras.
 *
 * Saídas: MS Project XML (.xml — abre direto no MS Project), .MPP via
 * helper local, Excel (.xlsx), CSV e JSON (projeto salvo).
 *
 * Dois modos de análise:
 *   1. Automática (offline, por regras) — nenhum dado sai do computador.
 *   2. Com IA (Claude) — opcional, usa a chave de API do próprio usuário,
 *      ou o modo "copiar prompt / colar JSON" sem chave.
 * ===================================================================== */
(function () {
  'use strict';

  const LS_PROJ = 'GLC_AE_PROJETO_V1';
  const LS_KEY = 'GLC_AE_APIKEY';
  const LS_MODEL = 'GLC_AE_MODEL';
  const MAX_AI_CHARS = 180000;

  // ---------------------------------------------------------------
  // Estado
  // ---------------------------------------------------------------
  let AE = novoProjeto();
  let uidSeq = 1;

  function novoProjeto() {
    return {
      meta: { nome: '', inicio: proxDiaUtilISO(new Date()), tipo: '', prazoDias: null, modo: '', fontes: [], geradoEm: '' },
      escopo: { objeto: '', prazos: [], marcos: [], premissas: [], restricoes: [], exclusoes: [], criterios: [], riscos: [] },
      nodes: [],
      texto: ''
    };
  }

  // ---------------------------------------------------------------
  // Utilitários
  // ---------------------------------------------------------------
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const norm = (s) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const say = (m) => (typeof toast === 'function' ? toast(m) : console.log(m));
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

  function capitalizar(s) {
    s = String(s || '').trim();
    return s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
  }
  function ehDiaUtil(d) { const w = d.getDay(); return w !== 0 && w !== 6; }
  function proxDiaUtil(d) { const x = new Date(d); while (!ehDiaUtil(x)) x.setDate(x.getDate() + 1); return x; }
  function proxDiaUtilISO(d) { return isoDate(proxDiaUtil(d)); }
  function isoDate(d) { const z = (n) => String(n).padStart(2, '0'); return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate()); }
  function brDate(d) { const z = (n) => String(n).padStart(2, '0'); return z(d.getDate()) + '/' + z(d.getMonth() + 1) + '/' + d.getFullYear(); }
  function parseISO(s) { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : proxDiaUtil(new Date()); }
  function addDiasUteis(d, n) { const x = new Date(d); let k = n; while (k > 0) { x.setDate(x.getDate() + 1); if (ehDiaUtil(x)) k--; } return x; }
  function baixar(nome, conteudo, tipo) {
    const blob = conteudo instanceof Blob ? conteudo : new Blob([conteudo], { type: tipo || 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = nome; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }
  function nomeArquivoBase() {
    return (AE.meta.nome || 'EAP_Projeto').replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_').slice(0, 80);
  }
  function novoNo(lvl, name, type, extra) {
    return Object.assign({ uid: uidSeq++, lvl, name, type: type || 'wp', dur: type === 'ms' ? 0 : 5, preds: [], desc: '', crit: '', fonte: '' }, extra || {});
  }

  // ---------------------------------------------------------------
  // Leitura de arquivos
  // ---------------------------------------------------------------
  async function lerArquivo(file) {
    const nome = file.name;
    const ext = (nome.split('.').pop() || '').toLowerCase();
    const buf = await file.arrayBuffer();
    if (ext === 'pdf') return lerPDF(buf);
    if (ext === 'docx') {
      if (!window.mammoth) throw new Error('Leitor de DOCX (mammoth) não carregado.');
      const r = await window.mammoth.extractRawText({ arrayBuffer: buf });
      return r.value || '';
    }
    if (['txt', 'md', 'csv', 'json', 'xml', 'html', 'htm'].includes(ext)) {
      let t = new TextDecoder('utf-8').decode(buf);
      if (t.includes('�')) t = new TextDecoder('windows-1252').decode(buf);
      if (ext === 'html' || ext === 'htm') { const div = document.createElement('div'); div.innerHTML = t; t = div.innerText; }
      return t;
    }
    if (ext === 'doc') throw new Error('Arquivos .doc (Word 97-2003) não são suportados. Salve como .docx ou PDF.');
    throw new Error('Formato não suportado: .' + ext);
  }

  async function lerPDF(buf) {
    const lib = window.pdfjsLib;
    if (!lib) throw new Error('Leitor de PDF (pdf.js) não carregado.');
    const pdf = await lib.getDocument({ data: new Uint8Array(buf) }).promise;
    const partes = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const tc = await page.getTextContent();
      let linha = '', ultimoY = null;
      const linhas = [];
      tc.items.forEach((it) => {
        const y = it.transform ? Math.round(it.transform[5]) : null;
        if (ultimoY !== null && y !== null && Math.abs(y - ultimoY) > 2) { linhas.push(linha); linha = ''; }
        linha += it.str;
        if (it.hasEOL) { linhas.push(linha); linha = ''; ultimoY = null; return; }
        ultimoY = y;
      });
      if (linha) linhas.push(linha);
      partes.push(linhas.map((l) => l.replace(/\s+/g, ' ').trim()).filter(Boolean).join('\n'));
    }
    const texto = partes.join('\n\n');
    if (texto.replace(/\s/g, '').length < 50) {
      throw new Error('O PDF parece ser digitalizado (imagem), sem texto selecionável. Use um PDF pesquisável (OCR) ou cole o texto.');
    }
    return texto;
  }

  async function aoSelecionarArquivos(files) {
    const lista = Array.from(files || []);
    if (!lista.length) return;
    const ta = $('ae-texto');
    const status = $('ae-arqs');
    for (const f of lista) {
      try {
        status.innerHTML += `<div>Lendo <b>${esc(f.name)}</b>…</div>`;
        const t = await lerArquivo(f);
        ta.value += (ta.value ? '\n\n' : '') + `===== ARQUIVO: ${f.name} =====\n` + t.trim() + '\n';
        if (!AE.meta.fontes.includes(f.name)) AE.meta.fontes.push(f.name);
        status.lastElementChild.innerHTML = `<b>${esc(f.name)}</b> — ${t.length.toLocaleString('pt-BR')} caracteres`;
      } catch (e) {
        status.lastElementChild.innerHTML = `<b>${esc(f.name)}</b> — ${esc(e.message)}`;
      }
    }
    atualizarContador();
  }

  function atualizarContador() {
    const n = ($('ae-texto').value || '').length;
    $('ae-chars').textContent = n.toLocaleString('pt-BR') + ' caracteres';
  }

  // ---------------------------------------------------------------
  // Análise automática (regras)
  // ---------------------------------------------------------------
  const RE_ENTREGAVEL = /\b(relatorio|estudo|projeto (basico|executivo|conceitual|funcional|preliminar|legal|de )|anteprojeto|plano (de|do|da|diretor)|memorial|especificac|desenho|levantamento|diagnostico|cronograma|orcamento|manual|caderno|parecer|laudo|modelagem|modelo (bim|digital|3d)|\bbim\b|as[ -]?built|treinamento|capacitac|licenc|\beia\b|\brima\b|sondage|topograf|inventario|viabilidade|termo de referencia|planilha|matriz|dossie|sistema|software|aplicativo|prototipo|piloto|documentac|procedimento|programa (de|do|da)|cadastro|mapeamento|analise (de|do|da)|avaliacao|auditoria|inspec|ensaio|testes?\b|comissionamento|implantac|instalac|fornecimento|montagem|obras?\b|execucao (de|do|da|das|dos)|construc|supervisao|fiscalizac|gerenciamento|apoio tecnico|consultoria|workshop|produto \d|entregave)/;
  const RE_NAO_ESCOPO = /(multa|pagament|rescis|\bforo\b|garantia contratual|reajust|penalidad|dotacao|cnpj|assinatura|testemunha|clausula penal|faturament|nota fiscal|medicao financeira|preco global|valor global|r\$\s?\d|seguro garantia|sancao|sancoes|recurso orcamentario|lei n|decreto)/;
  const RE_VERBO = /^(elaborar|elaboracao (de|do|da|dos|das)|desenvolver|desenvolvimento (de|do|da)|realizar|realizacao (de|do|da)|executar|preparar|preparacao (de|do|da)|emitir|emissao (de|do|da)|apresentar|apresentacao (de|do|da)|fornecer|entregar|entrega (de|do|da)|efetuar|fazer|confeccionar|confeccao (de|do|da)|produzir|producao (de|do|da)|prestar|prestacao (de|do|da)|conduzir|promover|providenciar)\s+/;
  const RE_SEC_BLOQ = /(premissa|pressupost|exclus|restric|criterio|condic(ao|oes) (comerc|gera|de pagamento)|pagamento|penalidad|sancoes|obrigac(ao|oes) (da|do|das|dos) contratante|garantia|apresentacao|introducao|historico|sobre a empresa|qualificac|experiencia|equipe tecnica|curriculo|atestado|valor|preco|proposta comercial|vigencia|foro|rescisao|disposic(ao|oes) (gera|fina))/;
  const RE_LOE = /(relatorio(s)? (mensa|de progresso|de andamento|semanal|gerencia)|reunio(es|ao) (de acompanhamento|mensa|semana|periodic)|monitoramento|acompanhamento (do|da) (projeto|contrato)|gestao do contrato)/;
  const KW_ESCOPO = /(escopo|objeto|servicos|atividades|produtos|entregaveis|entregas|etapas|fases|metodologia|plano de trabalho|especificac|termo de referencia|descricao dos servicos|descricao do servico|abrangencia|itens de servico|composicao dos servicos)/;

  const FASES = [
    { k: 'mob', ordem: 1, nome: 'Mobilização e Levantamentos', re: /(mobiliza|levantamento|cadastro|topograf|sondage|inventario|diagnostico|coleta de dados|visita tecnica|vistoria|dados existentes|investigac)/ },
    { k: 'est', ordem: 2, nome: 'Estudos e Concepção', re: /(estudo|viabilidade|alternativa|conceitual|concepcao|anteprojeto|preliminar|plano diretor|masterplan|modelagem)/ },
    { k: 'pb', ordem: 3, nome: 'Projeto Básico', re: /(projeto basico|\bbasico\b)/ },
    { k: 'pe', ordem: 4, nome: 'Projeto Executivo', re: /(projeto executivo|\bexecutivo\b|detalhamento)/ },
    { k: 'lic', ordem: 5, nome: 'Licenciamento e Aprovações', re: /(licenc|\beia\b|\brima\b|outorga|ambiental|aprovac(ao|oes) (junto|nos|nas|em)|autorizac)/ },
    { k: 'orc', ordem: 6, nome: 'Orçamento e Suprimentos', re: /(orcament|planilha|quantitativ|composic(ao|oes) de custo|suprimento|aquisic|compras?\b|licitac|edital|termo de referencia|especificac(ao|oes) de compra)/ },
    { k: 'exe', ordem: 7, nome: 'Execução e Implantação', re: /(\bobras?\b|execucao|construc|implantac|instalac|montagem|fornecimento|mobiliario|equipamento)/ },
    { k: 'sup', ordem: 8, nome: 'Supervisão e Fiscalização', re: /(supervis|fiscaliz|acompanhamento (de|da|das) obra|apoio tecnico (a|na) obra|assistencia tecnica)/ },
    { k: 'tst', ordem: 9, nome: 'Testes e Comissionamento', re: /(teste|ensaio|comission|inspec|homolog|startup|start-up|partida assistida)/ },
    { k: 'trn', ordem: 10, nome: 'Treinamento e Transição', re: /(treinamento|capacitac|manual|operacao assistida|as[ -]?built|transic|documentacao final|data book|databook|entrega final)/ },
    { k: 'out', ordem: 11, nome: 'Outros Entregáveis Técnicos', re: /$^/ }
  ];
  const RE_GER = /(plano de gerenciamento|relatorio(s)? (mensa|de progresso|de andamento|gerencia|de status|semanal)|reunio(es|ao) (de|mensal|semanal|de partida|de acompanhamento)|kick ?off|gestao do contrato|gerenciamento do projeto|cronograma (fisico|geral|de linha de base|detalhado)|eap\b|estrutura analitica)/;

  function limparNome(s) {
    let t = String(s || '')
      .replace(/^\s*([•●▪■◦\-\*–—>·o]|\(?[a-z]\)|\(?\d{1,2}\)|\d{1,2}(\.\d{1,2}){0,4}[\.\)]?|[ivxl]{1,6}[\.\)\-–—])\s+/i, '')
      .replace(/\s+/g, ' ')
      .replace(/[;:,\.]+$/, '')
      .trim();
    let n = norm(t);
    let guard = 0;
    while (RE_VERBO.test(n) && guard++ < 2) {
      const m = RE_VERBO.exec(n);
      t = t.slice(m[0].length); n = norm(t);
      t = t.replace(/^(o|a|os|as|um|uma|de|do|da|dos|das)\s+/i, '');
      n = norm(t);
    }
    if (t.length > 110) {
      const corte = t.slice(0, 110);
      const i = Math.max(corte.lastIndexOf(','), corte.lastIndexOf(' - '), corte.lastIndexOf(';'));
      t = (i > 40 ? corte.slice(0, i) : corte.replace(/\s+\S*$/, '')) + '…';
    }
    if (t === t.toUpperCase() && t.length > 6) t = t.toLowerCase();
    return capitalizar(t);
  }

  function classificarLinhas(texto) {
    const linhas = texto.replace(/\r/g, '').split('\n').map((l) => l.replace(/\s+/g, ' ').trim());
    const out = [];
    const reNum = /^(\d{1,2}(?:\.\d{1,2}){0,4})[\.\)]?\s*[-–—]?\s+(\S.{2,140})$/;
    const reClausula = /^(CL[ÁA]USULA|ANEXO|CAP[ÍI]TULO|SE[ÇC][ÃA]O|ETAPA|FASE|PRODUTO|M[ÓO]DULO|LOTE|META|ITEM)\s+([\w\dºª°\.\-]{1,12})\s*[-–—:.]?\s*(.{0,140})$/i;
    const reRomano = /^([IVXL]{1,6})\s*[\.\)\-–—]\s+(\S.{2,140})$/;
    const reBullet = /^([•●▪■◦\-\*–—·>]|\(?[a-z]\)|\(?\d{1,2}\))\s+(\S.{2,220})$/;
    linhas.forEach((l, i) => {
      if (!l) return;
      let m;
      if (/^=====\s*ARQUIVO:/.test(l)) { out.push({ i, t: 'arq', txt: l.replace(/^=+\s*ARQUIVO:\s*|\s*=+$/g, ''), d: 0 }); return; }
      if ((m = reClausula.exec(l))) { out.push({ i, t: 'h', d: 1, num: m[1] + ' ' + m[2], txt: (m[3] || (m[1] + ' ' + m[2])).trim(), tipoH: norm(m[1]) }); return; }
      if ((m = reNum.exec(l)) && !/[\.;]$/.test(l) && l.length <= 150 && !/^\d+\s*(dias|meses|%|m2|m²|km|un)/i.test(m[2])) {
        out.push({ i, t: 'h', d: m[1].split('.').length, num: m[1], txt: m[2].trim() }); return;
      }
      if ((m = reRomano.exec(l)) && l.length <= 150) { out.push({ i, t: 'h', d: 1, num: m[1], txt: m[2].trim() }); return; }
      if ((m = reBullet.exec(l))) { out.push({ i, t: 'b', txt: m[2].trim() }); return; }
      const letras = l.replace(/[^A-Za-zÀ-ú]/g, '');
      if (letras.length >= 4 && l.length <= 90 && letras === letras.toUpperCase() && !/[\.;,]$/.test(l)) {
        out.push({ i, t: 'h', d: 1, num: '', txt: l, caps: true }); return;
      }
      out.push({ i, t: 'p', txt: l });
    });
    return out;
  }

  function extrairPrazoDiasUteis(nt) {
    const re = /prazo[^\n.]{0,90}?(\d{1,4})\s*(?:\([^)]{0,40}\))?\s*(dias? uteis|dias? corridos|dias?|meses|mes|semanas)/g;
    let m, best = null;
    while ((m = re.exec(nt))) {
      const v = parseInt(m[1], 10);
      const u = m[2];
      let d;
      if (/uteis/.test(u)) d = v;
      else if (/mes/.test(u)) d = Math.round(v * 21);
      else if (/semana/.test(u)) d = v * 5;
      else d = Math.round(v * 5 / 7);
      if (d >= 5 && d <= 3000 && (best === null || d > best)) best = d;
    }
    return best;
  }

  function coletarFrases(linhasOrig, re, max) {
    const res = [];
    const vistos = new Set();
    linhasOrig.forEach((l) => {
      const n = norm(l);
      const limpo = l.replace(/^\s*([•●▪■◦\-\*–—·>]|\(?[a-z]\)|\(?\d{1,2}\))\s+/, '').trim();
      const ehTitulo = limpo.length < 60 && limpo === limpo.toUpperCase();
      if (re.test(n) && !ehTitulo && limpo.length >= 15 && limpo.length <= 400) {
        const k = norm(limpo).slice(0, 80);
        if (!vistos.has(k)) { vistos.add(k); res.push(limpo); }
      }
    });
    return res.slice(0, max || 15);
  }

  function extrairEscopo(texto) {
    const linhas = texto.replace(/\r/g, '').split(/\n|(?<=[.;])\s+(?=[A-ZÁÉÍÓÚ])/).map((l) => l.replace(/\s+/g, ' ').trim()).filter(Boolean);
    const nt = norm(texto);
    const esc0 = { objeto: '', prazos: [], marcos: [], premissas: [], restricoes: [], exclusoes: [], criterios: [], riscos: [] };
    const iObj = linhas.findIndex((l) => /^(\d+(\.\d+)*[\.\)]?\s*)?(do\s+)?objeto\b/i.test(l) || /\bobjeto (do presente|deste|desta|da presente|do contrato|da proposta)/i.test(l));
    if (iObj >= 0) {
      let o = linhas[iObj].replace(/^(\d+(\.\d+)*[\.\)]?\s*)?(do\s+)?objeto\s*[:\-–—]?\s*/i, '');
      if (o.length < 40 && linhas[iObj + 1]) o = (o + ' ' + linhas[iObj + 1]).trim();
      if (o.length < 60 && linhas[iObj + 2]) o = (o + ' ' + linhas[iObj + 2]).trim();
      esc0.objeto = o.slice(0, 600);
    }
    esc0.prazos = coletarFrases(linhas, /prazo[^.]{0,80}\d+\s*\(?[^)]{0,30}\)?\s*(dias|meses|semanas)/, 10);
    esc0.marcos = coletarFrases(linhas, /(marco|milestone|entrega (do|da|dos|das) produto|data limite|ate o dia|\b\d{2}\/\d{2}\/\d{4}\b.*(entrega|conclus|inicio))/, 12);
    esc0.premissas = coletarFrases(linhas, /(premissa|pressupost|considera-se que|sera fornecid[oa] pel[oa] (contratante|cliente)|caber[aá] (a|ao) (contratante|cliente)|disponibiliza(ra|do) pel[oa])/, 12);
    esc0.restricoes = coletarFrases(linhas, /(restric|limitac|nao podera|vedad[oa]|obrigatoriamente|devera obedecer|norma(s)? (abnt|nbr|tecnica)|\bnbr\b|horario de trabalho)/, 12);
    esc0.exclusoes = coletarFrases(linhas, /(nao (esta|estao|estara|estarao) (incluid|contemplad|previst)|exclui-se|excluem-se|nao faz(em)? parte|fora do escopo|nao contempla|nao inclu|exclusoes|excluido)/, 12);
    esc0.criterios = coletarFrases(linhas, /(criterio(s)? de aceita|aceite|aceitacao|aprovac(ao|oes) (formal|do contratante|da fiscalizacao|pela)|sera(o)? considerad[oa]s? (aceit|aprovad|concluid))/, 12);
    esc0.riscos = coletarFrases(linhas, /(multa|penalidad|risco|atraso|garantia|retencao|glosa|rescis)/, 12);
    esc0.premissas = esc0.premissas.concat(coletarFrases(linhas, /^(a |o )?(contratante|cliente) (disponibilizara|fornecera|providenciara|devera fornecer|sera responsavel)/, 8));
    // itens listados sob títulos de categoria (ex.: "5. PREMISSAS")
    const cats = [['premissas', /premissa|pressupost/], ['exclusoes', /exclus|nao incluso|fora do escopo/], ['restricoes', /restric|limitac/], ['criterios', /criterio(s)? de aceit|aceitacao/], ['riscos', /riscos?\b|penalidad/]];
    let cat = null, dCat = 0;
    classificarLinhas(texto).forEach((x) => {
      if (x.t === 'h') { const nh = norm(x.txt); const c = cats.find(([, re]) => re.test(nh)); if (c) { cat = c[0]; dCat = x.d; return; } if (cat && x.d <= dCat) cat = null; return; }
      if (!cat || x.t === 'arq') return;
      const v = x.txt.trim();
      if (v.length >= 12 && !esc0[cat].some((y) => norm(y).slice(0, 60) === norm(v).slice(0, 60))) esc0[cat].push(v);
    });
    Object.keys(esc0).forEach((k) => { if (Array.isArray(esc0[k])) esc0[k] = esc0[k].slice(0, 15); });
    return { esc0, prazo: extrairPrazoDiasUteis(nt) };
  }

  function detectarTipo(nt) {
    const pontos = [
      ['Contrato', /(contratante|contratada|clausula|celebram o presente|termo de contrato)/g],
      ['Proposta Técnica', /(proposta tecnica|nossa proposta|proponente|metodologia de execucao|plano de trabalho)/g],
      ['Termo de Referência', /(termo de referencia|projeto basico de contratacao|especificacoes tecnicas)/g],
      ['Edital', /(edital|licitante|pregao|concorrencia|habilitacao)/g]
    ];
    let melhor = ['Documento técnico', 0];
    pontos.forEach(([t, re]) => { const c = (nt.match(re) || []).length; if (c > melhor[1]) melhor = [t, c]; });
    return melhor[0];
  }

  function nomeProjetoSugerido(texto, objeto) {
    if (objeto) {
      let o = objeto.replace(/^(a |o )?(contratacao de |prestacao de |prestação de |contratação de |servicos de |serviços de )/i, '');
      o = limparNome(o.split(/[,;]|\. | visando | com o objetivo/i)[0]).replace(/…$/, '');
      if (o.length > 100) o = o.slice(0, 100).replace(/\s+\S*$/, '');
      if (o.length >= 10) return o;
    }
    const l = texto.split('\n').map((x) => x.trim()).find((x) => x.length > 15 && x.length < 120 && !/^=+/.test(x));
    return l ? capitalizar(l.slice(0, 90)) : 'Projeto';
  }

  function analisarRegras(texto) {
    const t = String(texto || '');
    const nt = norm(t);
    const itens = classificarLinhas(t);
    const { esc0, prazo } = extrairEscopo(t);

    // 1) Seções de escopo (cabeçalhos com palavras-chave)
    const heads = itens.filter((x) => x.t === 'h');
    const escopoIdx = new Set();
    heads.forEach((h, k) => {
      if (!KW_ESCOPO.test(norm(h.txt))) return;
      const fim = (heads.slice(k + 1).find((o) => o.d <= h.d && !o.caps) || { i: Infinity }).i;
      itens.forEach((x) => { if (x.i > h.i && x.i < fim) escopoIdx.add(x.i); });
    });

    // 2) Candidatos a entregáveis
    let secAtual = '';
    let faseDoc = null;
    let bloqueio = null; // profundidade da seção não-escopo em vigor
    const cands = [];
    const fasesDoc = [];
    itens.forEach((x, k) => {
      if (x.t === 'arq') { secAtual = ''; faseDoc = null; bloqueio = null; return; }
      const n = norm(x.txt);
      if (x.t === 'h') {
        if (bloqueio !== null && (x.d <= bloqueio || x.caps)) bloqueio = null;
        if (RE_SEC_BLOQ.test(n) && !RE_ENTREGAVEL.test(n.replace(/criterio.*$/, ''))) { bloqueio = x.d; faseDoc = null; secAtual = x.txt; return; }
        if (KW_ESCOPO.test(n) && !RE_ENTREGAVEL.test(n)) { secAtual = x.txt; faseDoc = null; return; }
        secAtual = x.txt;
        const ehFase = (x.tipoH && /^(etapa|fase|produto|modulo|lote|meta)$/.test(x.tipoH)) || /^(etapa|fase|produto|modulo|meta)\s+[\w\d]{1,6}\b/.test(n);
        // título com itens logo abaixo (marcadores/subtítulos) dentro do escopo = grupo (fase/entrega principal)
        const prox = itens[k + 1];
        const temFilhosAbaixo = prox && (prox.t === 'b' || (prox.t === 'h' && prox.d > x.d));
        if ((ehFase && (escopoIdx.has(x.i) || RE_ENTREGAVEL.test(n))) || (escopoIdx.has(x.i) && temFilhosAbaixo && !x.caps)) {
          faseDoc = { nome: limparNome(x.num && x.tipoH ? (x.num + (x.txt && x.txt !== x.num ? ' – ' + x.txt : '')) : x.txt), itens: [], i: x.i, sec: x.txt };
          fasesDoc.push(faseDoc);
          return;
        }
      }
      if (bloqueio !== null) return;
      if (RE_NAO_ESCOPO.test(n)) return;
      if (/^(a |o )?(contratante|contratada|cliente)\b|^considera-se|^(nao|não) (esta|estao|faz|fazem|inclui|contempla)/.test(n)) return;
      const noEscopo = escopoIdx.has(x.i);
      const pareceEntregavel = RE_ENTREGAVEL.test(n);
      const curto = x.txt.length <= 220;
      if (!curto) return;
      if ((x.t === 'b' || x.t === 'h') && (noEscopo || pareceEntregavel)) {
        cands.push({ txt: x.txt, sec: secAtual, fase: faseDoc, i: x.i, forte: pareceEntregavel });
      } else if (x.t === 'p' && pareceEntregavel && noEscopo && x.txt.length <= 160) {
        cands.push({ txt: x.txt, sec: secAtual, fase: faseDoc, i: x.i, forte: true });
      }
    });

    // 3) Deduplicação e limpeza
    const vistos = new Set();
    const limpos = [];
    cands.forEach((c) => {
      const nome = limparNome(c.txt);
      const k = norm(nome).replace(/[^a-z0-9]+/g, ' ').trim().slice(0, 50);
      if (nome.length < 4 || k.length < 4 || vistos.has(k)) return;
      if (/^(objeto|escopo|servicos|atividades|produtos|entregaveis|metodologia|introducao|apresentacao|sumario|indice|anexos?)$/.test(k)) return;
      if (!c.forte && !c.fase) return;
      vistos.add(k);
      limpos.push(Object.assign({}, c, { nome }));
    });

    // 4) Montagem da EAP
    uidSeq = 1;
    const nodes = [];
    const nomeProj = AE.meta.nome || nomeProjetoSugerido(t, esc0.objeto);
    const root = novoNo(1, nomeProj, 'summary');
    nodes.push(root);

    const ger = gerarGerenciamento(nodes);
    const gerItens = limpos.filter((c) => RE_GER.test(norm(c.nome)));
    const tecnicos = limpos.filter((c) => !RE_GER.test(norm(c.nome)));

    // Grupos (fases): estrutura do documento se houver, senão fases canônicas
    let grupos = [];
    const fasesComItens = fasesDoc.filter((f) => tecnicos.some((c) => c.fase === f));
    const fasesSoGerenciamento = fasesDoc.filter((f) => !fasesComItens.includes(f) && gerItens.some((c) => c.fase === f));
    if (fasesComItens.length >= 2) {
      grupos = fasesComItens.map((f) => ({ nome: f.nome, itens: tecnicos.filter((c) => c.fase === f) }));
      const soltos = tecnicos.filter((c) => !c.fase || !fasesComItens.includes(c.fase));
      fasesDoc.filter((f) => !fasesComItens.includes(f) && !fasesSoGerenciamento.includes(f) && RE_ENTREGAVEL.test(norm(f.nome)) && !RE_GER.test(norm(f.nome)))
        .forEach((f) => soltos.push({ nome: f.nome, txt: f.sec, sec: f.sec, i: f.i, forte: true }));
      if (soltos.length) grupos.push({ nome: 'Outros Entregáveis Técnicos', itens: soltos });
    } else {
      const porFase = {};
      tecnicos.forEach((c) => {
        const alvo = norm(c.nome) + ' ' + norm(c.sec);
        const f = FASES.find((F) => F.re.test(norm(c.nome))) || FASES.find((F) => F.re.test(alvo)) || FASES[FASES.length - 1];
        (porFase[f.k] = porFase[f.k] || { f, itens: [] }).itens.push(c);
      });
      grupos = Object.values(porFase).sort((a, b) => a.f.ordem - b.f.ordem).map((g) => ({ nome: g.f.nome, itens: g.itens }));
    }

    let avisoGenerico = false;
    if (!grupos.length) {
      avisoGenerico = true;
      grupos = [
        { nome: 'Levantamentos e Diagnóstico', itens: [{ nome: 'Relatório de diagnóstico' }] },
        { nome: 'Desenvolvimento Técnico', itens: [{ nome: 'Entregável técnico principal (revisar)' }, { nome: 'Memorial descritivo' }] },
        { nome: 'Entrega e Aceite', itens: [{ nome: 'Documentação final' }] }
      ];
    }

    // Limite de pacotes por fase para manter legibilidade
    grupos.forEach((g) => { if (g.itens.length > 25) g.itens = g.itens.slice(0, 25); });

    // Durações
    const total = prazo || null;
    const somaPesos = grupos.reduce((a, g) => a + Math.max(1, Math.sqrt(g.itens.length)), 0);
    const disponivel = total ? Math.max(grupos.length * 5, total - 12) : null;

    let anteriorMarco = ger.plano;
    grupos.forEach((g, gi) => {
      const durFase = disponivel ? Math.max(5, Math.round(disponivel * Math.max(1, Math.sqrt(g.itens.length)) / somaPesos)) : 20;
      const fase = novoNo(2, g.nome, 'summary');
      nodes.push(fase);
      g.itens.forEach((c) => {
        const wp = novoNo(3, c.nome, 'wp', {
          dur: durFase,
          preds: [anteriorMarco.uid],
          desc: c.txt && c.txt !== c.nome ? c.txt : '',
          fonte: c.sec ? 'Seção: ' + c.sec : (c.i != null ? 'Linha ' + (c.i + 1) : ''),
          crit: 'Aprovação formal do contratante'
        });
        nodes.push(wp);
      });
      const pacotes = nodes.filter((n) => n.lvl === 3 && nodes.indexOf(n) > nodes.indexOf(fase));
      const marco = novoNo(3, 'Aprovação – ' + g.nome, 'ms', { preds: pacotes.map((p) => p.uid), crit: 'Documento de aceite / aprovação da fase' });
      nodes.push(marco);
      anteriorMarco = marco;
      if (gi === 0 && !g.itens.length) marco.preds = [ger.plano.uid];
    });

    // Itens de gerenciamento encontrados no documento
    const loe = [ger.relatorios];
    gerItens.forEach((c) => {
      const nc = norm(c.nome);
      const fonte = c.sec ? 'Seção: ' + c.sec : '';
      if (/plano de gerenciamento/.test(nc)) { ger.plano.fonte = fonte; ger.plano.desc = c.txt || ''; return; }
      if (/relatorio(s)? (mensa|de progresso|de andamento|gerencia)/.test(nc)) { ger.relatorios.name = c.nome; ger.relatorios.fonte = fonte; ger.relatorios.desc = c.txt || ''; return; }
      if (/^(eap|estrutura analitica)|cronograma/.test(nc)) { ger.linhaBase.fonte = fonte; return; }
      if (nodes.some((n) => norm(n.name) === nc)) return;
      const pos = nodes.indexOf(ger.mcFim);
      const novo = novoNo(4, c.nome, 'wp', { dur: 5, preds: [ger.plano.uid], desc: c.txt, fonte, crit: RE_LOE.test(nc) ? 'Registros/atas emitidos no período' : 'Aprovação do contratante' });
      nodes.splice(pos + 1, 0, novo);
      ger.mcFim = novo;
      if (RE_LOE.test(nc)) loe.push(novo);
    });

    // Atividades de esforço contínuo (relatórios, reuniões) acompanham o prazo total
    const duracaoTotal = estimarCaminho(nodes);
    loe.forEach((n) => { n.dur = Math.max(5, (total || duracaoTotal) - 10); n.loe = true; });

    // Encerramento depende do último marco técnico e de toda atividade sem sucessora (rede fechada)
    const comSucessor = new Set();
    nodes.forEach((n) => (n.preds || []).forEach((u) => comSucessor.add(u)));
    const soltas = nodes.filter((n, i) => !temFilhos(i, nodes) && !comSucessor.has(n.uid) && n !== ger.aceite && n !== ger.licoes && n !== ger.fim && n !== anteriorMarco);
    ger.aceite.preds = [anteriorMarco.uid].concat(soltas.map((n) => n.uid));

    AE.nodes = nodes;
    AE.escopo = esc0;
    AE.meta.nome = nomeProj;
    AE.meta.tipo = detectarTipo(nt);
    AE.meta.prazoDias = total;
    AE.meta.modo = 'Automática (regras)';
    AE.meta.geradoEm = new Date().toISOString();
    AE.texto = t;
    return { avisoGenerico, nCands: limpos.length };
  }

  function gerarGerenciamento(nodes) {
    const g = novoNo(2, 'Gerenciamento do Projeto', 'summary');
    nodes.push(g);
    nodes.push(novoNo(3, 'Iniciação', 'summary'));
    const kick = novoNo(4, 'Reunião de partida (kick-off)', 'ms', { crit: 'Ata da reunião assinada' });
    nodes.push(kick);
    const tap = novoNo(4, 'Termo de Abertura do Projeto', 'wp', { dur: 2, preds: [kick.uid], crit: 'TAP aprovado pelo patrocinador' });
    nodes.push(tap);
    nodes.push(novoNo(3, 'Planejamento', 'summary'));
    const plano = novoNo(4, 'Plano de Gerenciamento do Projeto', 'wp', { dur: 5, preds: [tap.uid], crit: 'Plano aprovado (escopo, cronograma, custos, qualidade, riscos, comunicações)' });
    nodes.push(plano);
    const linhaBase = novoNo(4, 'EAP, Dicionário e Cronograma de Linha de Base', 'wp', { dur: 3, preds: [tap.uid], crit: 'Linha de base aprovada' });
    nodes.push(linhaBase);
    nodes.push(novoNo(3, 'Monitoramento e Controle', 'summary'));
    const relatorios = novoNo(4, 'Relatórios de Progresso', 'wp', { dur: 20, preds: [plano.uid], crit: 'Relatórios periódicos emitidos' });
    nodes.push(relatorios);
    const mudancas = novoNo(4, 'Registro de Mudanças, Riscos e Questões', 'wp', { dur: 5, preds: [plano.uid], crit: 'Registros atualizados' });
    nodes.push(mudancas);
    nodes.push(novoNo(3, 'Encerramento', 'summary'));
    const aceite = novoNo(4, 'Termo de Aceite Final', 'wp', { dur: 3, preds: [], crit: 'Termo de aceite assinado pelo contratante' });
    nodes.push(aceite);
    const licoes = novoNo(4, 'Lições Aprendidas', 'wp', { dur: 2, preds: [aceite.uid], crit: 'Registro de lições aprendidas arquivado' });
    nodes.push(licoes);
    const fim = novoNo(4, 'Projeto Encerrado', 'ms', { preds: [licoes.uid] });
    nodes.push(fim);
    return { plano, relatorios, aceite, licoes, fim, linhaBase, mcFim: mudancas };
  }

  function estimarCaminho(nodes) {
    const sch = calcularDatas(nodes, new Date(2026, 0, 5));
    return sch[nodes[0].uid] ? sch[nodes[0].uid].fimPt : 0;
  }

  // ---------------------------------------------------------------
  // Estrutura (árvore em lista plana)
  // ---------------------------------------------------------------
  function fimSubarvore(i) {
    const lvl = AE.nodes[i].lvl;
    let j = i + 1;
    while (j < AE.nodes.length && AE.nodes[j].lvl > lvl) j++;
    return j; // exclusivo
  }
  function temFilhos(i, arr) {
    const a = arr || AE.nodes;
    return i + 1 < a.length && a[i + 1].lvl > a[i].lvl;
  }
  function tipoEfetivo(i, arr) {
    const a = arr || AE.nodes;
    return temFilhos(i, a) ? 'summary' : (a[i].type === 'ms' ? 'ms' : 'wp');
  }
  function calcularEDT(arr) {
    const a = arr || AE.nodes;
    const cont = [];
    const res = [];
    a.forEach((n) => {
      cont.length = n.lvl;
      cont[n.lvl - 1] = (cont[n.lvl - 1] || 0) + 1;
      for (let k = 0; k < n.lvl - 1; k++) if (!cont[k]) cont[k] = 1;
      res.push(cont.slice(0, n.lvl).join('.'));
    });
    return res;
  }
  function normalizarNiveis() {
    const a = AE.nodes;
    if (!a.length) return;
    a[0].lvl = 1;
    for (let i = 1; i < a.length; i++) {
      if (a[i].lvl < 2) a[i].lvl = 2;
      if (a[i].lvl > a[i - 1].lvl + 1) a[i].lvl = a[i - 1].lvl + 1;
    }
    const uids = new Set(a.map((n) => n.uid));
    a.forEach((n, i) => {
      n.preds = (n.preds || []).filter((u) => uids.has(u) && u !== n.uid);
      if (temFilhos(i)) n.type = n.type === 'ms' ? 'wp' : n.type;
    });
  }

  // ---------------------------------------------------------------
  // Datas (dias úteis, seg-sex) — prévia; o MS Project recalcula
  // ---------------------------------------------------------------
  function calcularDatas(nodes, inicio) {
    // Modelo em "pontos" (dias úteis decorridos desde o início):
    // pacote ocupa [es, es+dur); marco é um ponto no tempo.
    const a = nodes;
    const ini = proxDiaUtil(inicio);
    const sch = {};
    const idxPorUid = {};
    a.forEach((n, i) => { idxPorUid[n.uid] = i; });
    const folha = (i) => !temFilhos(i, a);
    const fimPt = (uid) => {
      const i = idxPorUid[uid];
      if (i == null) return 0;
      if (folha(i)) return sch[uid] ? sch[uid].fimPt : null;
      let j = i + 1, m = 0;
      while (j < a.length && a[j].lvl > a[i].lvl) {
        if (folha(j)) { const s = sch[a[j].uid]; if (!s) return null; m = Math.max(m, s.fimPt); }
        j++;
      }
      return m;
    };
    let pendentes = a.map((n, i) => i).filter(folha);
    let ciclo = false;
    for (let guarda = 0; pendentes.length && guarda < a.length + 5; guarda++) {
      const prox = [];
      pendentes.forEach((i) => {
        const n = a[i];
        let es = 0, ok = true;
        (n.preds || []).forEach((u) => { const f = fimPt(u); if (f === null) ok = false; else es = Math.max(es, f); });
        if (!ok) { prox.push(i); return; }
        const ms = n.type === 'ms';
        const d = ms ? 0 : Math.max(1, Math.round(+n.dur || 1));
        sch[n.uid] = { iniPt: es, fimPt: es + d, ms };
      });
      if (prox.length === pendentes.length) { ciclo = true; prox.forEach((i) => { sch[a[i].uid] = { iniPt: 0, fimPt: a[i].type === 'ms' ? 0 : Math.max(1, +a[i].dur || 1), ms: a[i].type === 'ms', ciclo: true }; }); break; }
      pendentes = prox;
    }
    for (let i = a.length - 1; i >= 0; i--) {
      if (folha(i)) continue;
      let j = i + 1, mi = Infinity, mf = 0;
      while (j < a.length && a[j].lvl > a[i].lvl) { const s = sch[a[j].uid]; if (s) { mi = Math.min(mi, s.iniPt); mf = Math.max(mf, s.fimPt); } j++; }
      sch[a[i].uid] = { iniPt: mi === Infinity ? 0 : mi, fimPt: mf, resumo: true };
    }
    Object.values(sch).forEach((s) => {
      s.iniIdx = s.ms ? Math.max(0, s.fimPt - 1) : s.iniPt;
      if (s.ms && s.fimPt === 0) s.iniIdx = 0;
      s.fimIdx = s.ms ? s.iniIdx : Math.max(s.iniPt, s.fimPt - 1);
      s.ini = addDiasUteis(ini, s.iniIdx);
      s.fim = addDiasUteis(ini, s.fimIdx);
      s.dur = s.ms ? 0 : (s.fimPt - s.iniPt);
    });
    if (ciclo) sch.__ciclo = true;
    return sch;
  }

  // ---------------------------------------------------------------
  // Validação (boas práticas PMBOK / Practice Standard for WBS)
  // ---------------------------------------------------------------
  function validar() {
    const a = AE.nodes;
    const edt = calcularEDT();
    const avisos = [];
    const add = (nivel, msg, i) => avisos.push({ nivel, msg, edt: i != null ? edt[i] : '' });
    if (!a.length) return avisos;
    const pacotes = a.map((n, i) => i).filter((i) => tipoEfetivo(i) === 'wp');
    const marcos = a.map((n, i) => i).filter((i) => tipoEfetivo(i) === 'ms');
    if (!pacotes.length) add('erro', 'Nenhum pacote de trabalho (nível mais baixo) definido.');
    if (!a.some((n) => n.lvl === 2 && norm(n.name).includes('gerenciamento'))) add('aviso', 'Inclua um ramo de "Gerenciamento do Projeto" no nível 2 (o trabalho de gestão também faz parte dos 100% do escopo).');
    const maxNivel = Math.max(...a.map((n) => n.lvl));
    if (maxNivel > 6) add('aviso', `A EAP tem ${maxNivel} níveis. Acima de 5–6 níveis costuma indicar detalhamento excessivo.`);
    a.forEach((n, i) => {
      const tp = tipoEfetivo(i);
      if (tp === 'summary') {
        const filhos = [];
        let j = i + 1;
        while (j < a.length && a[j].lvl > n.lvl) { if (a[j].lvl === n.lvl + 1) filhos.push(j); j++; }
        if (filhos.length === 1) add('aviso', `"${n.name}" tem só 1 filho: decomponha em 2+ elementos ou elimine o nível (regra dos 100%).`, i);
        const nomes = filhos.map((k) => norm(a[k].name));
        const dup = nomes.filter((x, k) => nomes.indexOf(x) !== k);
        if (dup.length) add('aviso', `"${n.name}" tem elementos com nome repetido.`, i);
      }
      if (tp === 'wp') {
        if (RE_VERBO.test(norm(n.name))) add('aviso', `"${n.name}": prefira substantivos (entregas), não verbos — ex.: "Relatório X" em vez de "Elaborar relatório X".`, i);
        if ((+n.dur || 0) <= 0) add('erro', `"${n.name}": pacote de trabalho sem duração.`, i);
        if ((+n.dur || 0) > 60 && !n.loe && !RE_LOE.test(norm(n.name))) add('aviso', `"${n.name}": ${n.dur} dias úteis. Pacotes muito longos dificultam o controle — avalie decompor.`, i);
        if (!(n.crit || '').trim()) add('info', `"${n.name}": sem critério de aceitação no dicionário da EAP.`, i);
      }
      if (tp === 'ms' && (+n.dur || 0) > 0) add('aviso', `"${n.name}": marco deve ter duração zero.`, i);
      if (n.name.trim().length < 3) add('erro', 'Elemento sem nome.', i);
    });
    // Dependências: atividades soltas
    const temSucessor = new Set();
    a.forEach((n) => (n.preds || []).forEach((u) => temSucessor.add(u)));
    const folhas = pacotes.concat(marcos);
    const semPred = folhas.filter((i) => !(a[i].preds || []).length);
    if (semPred.length > 1) add('aviso', `${semPred.length} atividades sem predecessora (início em aberto). Ideal: apenas o marco de início.`);
    const semSuc = folhas.filter((i) => !temSucessor.has(a[i].uid));
    if (semSuc.length > 1) add('info', `${semSuc.length} atividades sem sucessora (término em aberto). Ideal: apenas o marco de fim.`);
    const sch = calcularDatas(a, parseISO(AE.meta.inicio));
    if (sch.__ciclo) add('erro', 'Há dependência circular nas predecessoras.');
    if (AE.meta.prazoDias) {
      const fim = sch[a[0].uid] ? sch[a[0].uid].fimPt : 0;
      if (fim > AE.meta.prazoDias * 1.05) add('aviso', `Duração estimada (${fim} dias úteis) excede o prazo identificado no documento (${AE.meta.prazoDias} dias úteis).`);
    }
    return avisos;
  }

  // ---------------------------------------------------------------
  // Renderização
  // ---------------------------------------------------------------
  let renderAgendado = false;
  function agendarRender() {
    if (renderAgendado) return;
    renderAgendado = true;
    setTimeout(() => { renderAgendado = false; render(); }, 0);
  }

  function render() {
    normalizarNiveis();
    const a = AE.nodes;
    $('ae-nome').value = AE.meta.nome || '';
    $('ae-inicio').value = AE.meta.inicio || '';
    $('ae-resultado').style.display = a.length ? 'block' : 'none';
    if (!a.length) return;
    const edt = calcularEDT();
    const uid2edt = {};
    a.forEach((n, i) => { uid2edt[n.uid] = edt[i]; });
    const sch = calcularDatas(a, parseISO(AE.meta.inicio));

    // Resumo
    const nPac = a.filter((n, i) => tipoEfetivo(i) === 'wp').length;
    const nMs = a.filter((n, i) => tipoEfetivo(i) === 'ms').length;
    const sRoot = sch[a[0].uid];
    const durTot = sRoot ? sRoot.dur : 0;
    $('ae-stats').innerHTML = [
      ['Elementos', a.length, 'var(--accent)'],
      ['Pacotes de trabalho', nPac, 'var(--green)'],
      ['Marcos', nMs, 'var(--yellow)'],
      ['Níveis', Math.max(...a.map((n) => n.lvl)), 'var(--purple)'],
      ['Duração (dias úteis)', durTot, 'var(--text)']
    ].map(([l, v, c]) => `<div class="st"><div class="n" style="color:${c}">${v}</div><div class="l">${l}</div></div>`).join('');
    $('ae-meta-info').innerHTML =
      `<b>Documento:</b> ${esc(AE.meta.tipo || '—')} &nbsp;|&nbsp; <b>Análise:</b> ${esc(AE.meta.modo || '—')} &nbsp;|&nbsp; ` +
      `<b>Prazo identificado:</b> ${AE.meta.prazoDias ? AE.meta.prazoDias + ' dias úteis' : 'não identificado'} &nbsp;|&nbsp; ` +
      `<b>Término previsto:</b> ${sRoot ? brDate(sRoot.fim) : '—'}` +
      (AE.meta.fontes.length ? ` &nbsp;|&nbsp; <b>Fontes:</b> ${AE.meta.fontes.map(esc).join(', ')}` : '');

    // Tabela
    let h = '<table class="ae-tbl"><thead><tr><th>EDT</th><th style="min-width:340px">Elemento da EAP</th><th>Tipo</th><th>Dur. (d.u.)</th><th>Predecessoras (EDT)</th><th>Início</th><th>Término</th><th>Ações</th></tr></thead><tbody>';
    a.forEach((n, i) => {
      const tp = tipoEfetivo(i);
      const s = sch[n.uid] || {};
      const cor = n.lvl === 1 ? 'var(--accent)' : n.lvl === 2 ? 'var(--green)' : tp === 'ms' ? 'var(--yellow)' : 'var(--text)';
      const peso = tp === 'summary' ? 'font-weight:700;' : '';
      const predTxt = (n.preds || []).map((u) => uid2edt[u]).filter(Boolean).join(';');
      h += `<tr data-i="${i}" class="${tp === 'summary' ? 'ae-sum' : ''}">` +
        `<td class="ae-edt">${edt[i]}</td>` +
        `<td><div style="padding-left:${(n.lvl - 1) * 18}px;display:flex;align-items:center;gap:6px">` +
        `<span class="ae-ico">${tp === 'summary' ? '▾' : tp === 'ms' ? '◆' : '▪'}</span>` +
        `<input class="ae-in" data-k="name" value="${esc(n.name)}" style="color:${cor};${peso}"></div></td>` +
        `<td>${i === 0 ? 'Projeto' : tp === 'summary' ? 'Resumo' : `<select class="ae-in ae-sel" data-k="type"><option value="wp"${tp === 'wp' ? ' selected' : ''}>Pacote</option><option value="ms"${tp === 'ms' ? ' selected' : ''}>Marco</option></select>`}</td>` +
        `<td>${tp === 'wp' ? `<input class="ae-in ae-num" data-k="dur" type="number" min="1" value="${+n.dur || 1}">` : `<span class="ae-calc">${tp === 'ms' ? 0 : (s.dur || '')}</span>`}</td>` +
        `<td>${tp === 'summary' ? '' : `<input class="ae-in ae-pred" data-k="preds" value="${esc(predTxt)}" placeholder="ex.: 1.2.1;1.2.2">`}</td>` +
        `<td class="ae-dt">${s.ini ? brDate(s.ini) : ''}</td><td class="ae-dt">${s.fim ? brDate(s.fim) : ''}</td>` +
        `<td class="ae-acts">${i === 0 ? `<button data-a="add" title="Adicionar filho">＋</button>` :
          `<button data-a="up" title="Subir">↑</button><button data-a="down" title="Descer">↓</button>` +
          `<button data-a="out" title="Promover (nível acima)">⇤</button><button data-a="in" title="Rebaixar (nível abaixo)">⇥</button>` +
          `<button data-a="add" title="Adicionar filho">＋</button><button data-a="sib" title="Adicionar irmão abaixo">⤓</button>` +
          `<button data-a="dic" title="Dicionário da EAP" aria-label="Dicionário da EAP">≡</button><button data-a="del" class="ae-del" title="Excluir (com filhos)">✕</button>`}</td>` +
        '</tr>';
    });
    h += '</tbody></table>';
    $('ae-arvore').innerHTML = h;

    // Validação
    const av = validar();
    const cont = { erro: 0, aviso: 0, info: 0 };
    av.forEach((x) => cont[x.nivel]++);
    ['ae-val-badge','ae-val-badge2'].forEach((id) => { const b = $(id); if (!b) return; b.textContent = av.length ? `${cont.erro} erro(s) · ${cont.aviso} aviso(s)` : 'Boas práticas: OK'; b.className = 'bdg ' + (cont.erro ? 'bdg-err' : cont.aviso ? 'bdg-y' : 'bdg-g'); });
    $('ae-val').innerHTML = av.length
      ? '<ul class="ae-list">' + av.map((x) => `<li class="ae-${x.nivel}"><b>${x.nivel.toUpperCase()}</b>${x.edt ? ' [' + esc(x.edt) + ']' : ''} — ${esc(x.msg)}</li>`).join('') + '</ul>'
      : '<p style="color:var(--green)">Nenhum problema encontrado nas verificações de boas práticas.</p>';

    // Declaração de escopo
    const E = AE.escopo || {};
    const bloco = (tit, arr) => `<details class="fs" ${arr && arr.length ? 'open' : ''}><summary>${tit} <span class="bdg">${(arr || []).length}</span></summary>` +
      ((arr && arr.length) ? '<ul class="ae-list">' + arr.map((x) => `<li>${esc(typeof x === 'string' ? x : (x.nome ? x.nome + (x.prazo ? ' — ' + x.prazo : '') : JSON.stringify(x)))}</li>`).join('') + '</ul>' : '<p class="ae-muted">Nada identificado — complete manualmente se necessário.</p>') + '</details>';
    $('ae-escopo').innerHTML =
      `<p><b>Objeto:</b> ${esc(E.objeto || 'não identificado')}</p>` +
      bloco('Prazos', E.prazos) + bloco('Marcos contratuais', E.marcos) + bloco('Critérios de aceitação', E.criterios) +
      bloco('Premissas', E.premissas) + bloco('Restrições', E.restricoes) + bloco('Exclusões do escopo', E.exclusoes) + bloco('Riscos e penalidades', E.riscos);

    // Dicionário
    let d = '<table class="ae-tbl"><thead><tr><th>EDT</th><th>Pacote / Marco</th><th>Descrição</th><th>Critério de aceitação</th><th>Fonte</th></tr></thead><tbody>';
    a.forEach((n, i) => {
      const tp = tipoEfetivo(i);
      if (tp === 'summary') return;
      d += `<tr data-i="${i}"><td class="ae-edt">${edt[i]}</td><td>${tp === 'ms' ? '◆ ' : ''}${esc(n.name)}</td>` +
        `<td><textarea class="ae-in ae-ta" data-k="desc" rows="2">${esc(n.desc)}</textarea></td>` +
        `<td><textarea class="ae-in ae-ta" data-k="crit" rows="2">${esc(n.crit)}</textarea></td>` +
        `<td><input class="ae-in" data-k="fonte" value="${esc(n.fonte)}"></td></tr>`;
    });
    d += '</tbody></table>';
    $('ae-dic').innerHTML = d;
    salvarLocal();
  }

  // ---------------------------------------------------------------
  // Edição
  // ---------------------------------------------------------------
  function aplicarCampo(i, k, v) {
    const n = AE.nodes[i];
    if (!n) return;
    if (k === 'name') n.name = v;
    else if (k === 'dur') n.dur = Math.max(1, parseInt(v, 10) || 1);
    else if (k === 'type') { n.type = v; if (v === 'ms') n.dur = 0; else if (!n.dur) n.dur = 5; }
    else if (k === 'preds') {
      const edt = calcularEDT();
      const mapa = {};
      edt.forEach((e, j) => { mapa[e] = AE.nodes[j].uid; });
      const invalidos = [];
      n.preds = String(v).split(/[;,\s]+/).map((x) => x.trim()).filter(Boolean).map((e) => { if (!mapa[e]) invalidos.push(e); return mapa[e]; }).filter((u) => u && u !== n.uid);
      if (invalidos.length) say('EDT não encontrada: ' + invalidos.join(', '));
    } else n[k] = v;
  }

  function acao(i, a) {
    const arr = AE.nodes;
    const n = arr[i];
    const fim = fimSubarvore(i);
    if (a === 'add') {
      arr.splice(fim, 0, novoNo(n.lvl + 1, 'Novo pacote de trabalho', 'wp'));
    } else if (a === 'sib') {
      arr.splice(fim, 0, novoNo(n.lvl, 'Novo elemento', 'wp'));
    } else if (a === 'del') {
      const qtd = fim - i;
      if (!confirm(`Excluir "${n.name}"${qtd > 1 ? ` e seus ${qtd - 1} elemento(s) filho(s)` : ''}?`)) return;
      arr.splice(i, qtd);
    } else if (a === 'in') {
      if (i <= 1 || arr[i - 1].lvl < n.lvl) return say('Não é possível rebaixar: não há elemento irmão acima.');
      for (let k = i; k < fim; k++) arr[k].lvl++;
    } else if (a === 'out') {
      if (n.lvl <= 2) return say('O nível 2 é o mais alto abaixo do projeto.');
      for (let k = i; k < fim; k++) arr[k].lvl--;
    } else if (a === 'up') {
      let j = i - 1;
      while (j > 0 && arr[j].lvl > n.lvl) j--;
      if (j <= 0 || arr[j].lvl !== n.lvl) return;
      const bloco = arr.splice(i, fim - i);
      arr.splice(j, 0, ...bloco);
    } else if (a === 'down') {
      if (fim >= arr.length || arr[fim].lvl !== n.lvl) return;
      const fim2 = fimSubarvore(fim);
      const bloco = arr.splice(fim, fim2 - fim);
      arr.splice(i, 0, ...bloco);
    } else if (a === 'dic') {
      document.querySelector('#ae-dic-wrap').open = true;
      setTimeout(() => { const tr = document.querySelector(`#ae-dic tr[data-i="${i}"]`); if (tr) { tr.scrollIntoView({ behavior: 'smooth', block: 'center' }); tr.classList.add('ae-flash'); } }, 50);
      return;
    }
    render();
  }

  // ---------------------------------------------------------------
  // IA (Claude) — opcional
  // ---------------------------------------------------------------
  function montarPrompt(texto) {
    const t = texto.length > MAX_AI_CHARS ? texto.slice(0, MAX_AI_CHARS) + '\n[...texto truncado...]' : texto;
    return `Você é um gerente de projetos sênior (PMP) especializado em projetos de engenharia e consultoria.
Analise a documentação de referência abaixo (proposta técnica, contrato, termo de referência ou edital) e monte a Estrutura Analítica do Projeto (EAP) segundo o PMBOK Guide e o Practice Standard for Work Breakdown Structures do PMI.

REGRAS OBRIGATÓRIAS
1. EAP orientada a ENTREGAS: nomes de elementos são substantivos (ex.: "Relatório de Diagnóstico", "Projeto Executivo de Drenagem"), nunca verbos.
2. Regra dos 100%: a EAP cobre todo o escopo do documento — e nada além dele. Não invente escopo; cite a fonte (cláusula/seção/página) de cada pacote.
3. Nível 1 = o projeto. Nível 2 = "Gerenciamento do Projeto" (iniciação, planejamento, monitoramento e controle, encerramento) + fases/entregas principais.
4. Todo elemento-resumo tem pelo menos 2 filhos. Pacotes de trabalho ficam no nível mais baixo. Máximo 5 níveis.
5. Inclua marcos (duração 0) para aprovações/aceites de cada fase e para início e fim do projeto.
6. Durações em dias úteis, coerentes com o prazo contratual identificado. Predecessoras apenas entre pacotes e marcos (tipo término-início), referenciadas pela EDT.
7. Para cada pacote e marco preencha o dicionário: descrição e critério de aceitação.
8. Responda em português do Brasil e SOMENTE com um JSON válido, sem comentários e sem markdown, no formato:

{
  "projeto": "nome curto do projeto",
  "tipo_documento": "Contrato | Proposta Técnica | Termo de Referência | Edital | Outro",
  "objeto": "objeto do contrato/proposta em 1-3 frases",
  "prazo_total_dias_uteis": 0,
  "prazos": ["..."],
  "marcos_contratuais": ["..."],
  "premissas": ["..."],
  "restricoes": ["..."],
  "exclusoes": ["..."],
  "criterios_aceitacao": ["..."],
  "riscos": ["..."],
  "eap": [
    {"edt": "1", "nome": "...", "tipo": "resumo"},
    {"edt": "1.1", "nome": "Gerenciamento do Projeto", "tipo": "resumo"},
    {"edt": "1.1.1", "nome": "Reunião de partida", "tipo": "marco", "duracao_dias_uteis": 0, "predecessoras": [], "descricao": "...", "criterio_aceitacao": "...", "fonte": "..."},
    {"edt": "1.2.1", "nome": "...", "tipo": "pacote", "duracao_dias_uteis": 10, "predecessoras": ["1.1.1"], "descricao": "...", "criterio_aceitacao": "...", "fonte": "Cláusula 3.2"}
  ]
}

DOCUMENTAÇÃO DE REFERÊNCIA
<<<
${t}
>>>`;
  }

  function importarJSONIA(obj) {
    if (typeof obj === 'string') {
      const s = obj.trim().replace(/^```(?:json)?/i, '').replace(/```$/, '');
      const i = s.indexOf('{'), j = s.lastIndexOf('}');
      if (i < 0 || j < 0) throw new Error('Não encontrei um JSON na resposta.');
      obj = JSON.parse(s.slice(i, j + 1));
    }
    if (!obj || !Array.isArray(obj.eap) || !obj.eap.length) throw new Error('JSON sem a lista "eap".');
    const linhas = obj.eap
      .filter((e) => e && e.edt && e.nome)
      .map((e) => Object.assign({}, e, { edt: String(e.edt).trim().replace(/\.$/, '') }))
      .sort((x, y) => {
        const a = x.edt.split('.').map(Number), b = y.edt.split('.').map(Number);
        for (let k = 0; k < Math.max(a.length, b.length); k++) { const d = (a[k] ?? -1) - (b[k] ?? -1); if (d) return d; }
        return 0;
      });
    uidSeq = 1;
    const porEdt = {};
    const nodes = linhas.map((e) => {
      const tp = /marco|milestone/i.test(e.tipo || '') ? 'ms' : /resumo|summary|fase/i.test(e.tipo || '') ? 'summary' : 'wp';
      const n = novoNo(Math.max(1, e.edt.split('.').length), String(e.nome).trim(), tp, {
        dur: tp === 'ms' ? 0 : Math.max(1, Math.round(+e.duracao_dias_uteis || +e.duracao || 5)),
        desc: e.descricao || '', crit: e.criterio_aceitacao || '', fonte: e.fonte || ''
      });
      n._edt = e.edt; n._preds = Array.isArray(e.predecessoras) ? e.predecessoras : String(e.predecessoras || '').split(/[;,\s]+/);
      porEdt[e.edt] = n;
      return n;
    });
    if (nodes[0] && nodes[0].lvl !== 1) nodes.unshift(novoNo(1, obj.projeto || 'Projeto', 'summary'));
    nodes.forEach((n) => { n.preds = (n._preds || []).map((x) => porEdt[String(x).trim()]).filter(Boolean).map((p) => p.uid); delete n._preds; delete n._edt; });
    AE.nodes = nodes;
    AE.meta.nome = obj.projeto || AE.meta.nome || nodes[0].name;
    AE.meta.tipo = obj.tipo_documento || AE.meta.tipo;
    AE.meta.prazoDias = +obj.prazo_total_dias_uteis || AE.meta.prazoDias || null;
    AE.meta.geradoEm = new Date().toISOString();
    AE.escopo = {
      objeto: obj.objeto || '', prazos: obj.prazos || [], marcos: obj.marcos_contratuais || [],
      premissas: obj.premissas || [], restricoes: obj.restricoes || [], exclusoes: obj.exclusoes || [],
      criterios: obj.criterios_aceitacao || [], riscos: obj.riscos || []
    };
    normalizarNiveis();
  }

  async function analisarIA() {
    const texto = $('ae-texto').value.trim();
    if (texto.length < 200) return say('Carregue ou cole a documentação primeiro (mín. 200 caracteres).');
    const chave = $('ae-apikey').value.trim();
    const modelo = $('ae-modelo').value.trim() || 'claude-sonnet-5-5';
    if (!chave) return say('Informe sua chave de API da Anthropic, ou use "Copiar prompt" + "Importar JSON".');
    if ($('ae-lembrar').checked) { try { localStorage.setItem(LS_KEY, chave); } catch (e) { /* ignora */ } } else { try { localStorage.removeItem(LS_KEY); } catch (e) { /* ignora */ } }
    try { localStorage.setItem(LS_MODEL, modelo); } catch (e) { /* ignora */ }
    const btn = $('ae-btn-ia');
    btn.disabled = true; btn.textContent = 'Analisando com IA…';
    try {
      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'content-type': 'application/json',
          'x-api-key': chave,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true'
        },
        body: JSON.stringify({ model: modelo, max_tokens: 16000, messages: [{ role: 'user', content: montarPrompt(texto) }] })
      });
      const data = await res.json();
      if (!res.ok) throw new Error((data && data.error && data.error.message) || ('HTTP ' + res.status));
      const out = (data.content || []).filter((c) => c.type === 'text').map((c) => c.text).join('\n');
      AE.texto = texto;
      AE.meta.nome = $('ae-nome').value.trim();
      importarJSONIA(out);
      AE.meta.modo = 'IA (' + modelo + ')';
      render();
      say('EAP gerada pela IA. Revise a validação e o dicionário.');
      $('ae-resultado').scrollIntoView({ behavior: 'smooth' });
    } catch (e) {
      say('Falha na análise com IA: ' + e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Analisar com IA (Claude)';
    }
  }

  // ---------------------------------------------------------------
  // Exportações
  // ---------------------------------------------------------------
  function xmlEsc(s) { return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ''); }

  function gerarMSPDI() {
    const a = AE.nodes;
    const edt = calcularEDT();
    const inicio = parseISO(AE.meta.inicio);
    const ini = isoDate(proxDiaUtil(inicio)) + 'T08:00:00';
    const agora = new Date().toISOString().slice(0, 19);
    const dias = [1, 2, 3, 4, 5, 6, 7].map((d) => {
      const util = d >= 2 && d <= 6;
      return `<WeekDay><DayType>${d}</DayType><DayWorking>${util ? 1 : 0}</DayWorking>` +
        (util ? '<WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime><WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes>' : '') +
        '</WeekDay>';
    }).join('');
    let x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Project xmlns="http://schemas.microsoft.com/project">\n';
    x += `<SaveVersion>14</SaveVersion><Name>${xmlEsc(nomeArquivoBase())}.xml</Name><Title>${xmlEsc(AE.meta.nome)}</Title><Author>Gerador EAP</Author>`;
    x += `<CreationDate>${agora}</CreationDate><ScheduleFromStart>1</ScheduleFromStart><StartDate>${ini}</StartDate>`;
    x += '<CalendarUID>1</CalendarUID><DefaultStartTime>08:00:00</DefaultStartTime><DefaultFinishTime>17:00:00</DefaultFinishTime>';
    x += '<MinutesPerDay>480</MinutesPerDay><MinutesPerWeek>2400</MinutesPerWeek><DaysPerMonth>20</DaysPerMonth><DurationFormat>7</DurationFormat>\n';
    x += `<Calendars><Calendar><UID>1</UID><Name>Padrão (seg-sex, 8h)</Name><IsBaseCalendar>1</IsBaseCalendar><IsBaselineCalendar>0</IsBaselineCalendar><BaseCalendarUID>-1</BaseCalendarUID><WeekDays>${dias}</WeekDays></Calendar></Calendars>\n<Tasks>\n`;
    a.forEach((n, i) => {
      const tp = tipoEfetivo(i);
      const uid = i + 1;
      const notas = [n.desc && 'Descrição: ' + n.desc, n.crit && 'Critério de aceitação: ' + n.crit, n.fonte && 'Fonte: ' + n.fonte].filter(Boolean).join('\n');
      x += `<Task><UID>${uid}</UID><ID>${uid}</ID><Name>${xmlEsc(n.name)}</Name><Active>1</Active><Manual>0</Manual><Type>0</Type><IsNull>0</IsNull>`;
      x += `<WBS>${edt[i]}</WBS><OutlineNumber>${edt[i]}</OutlineNumber><OutlineLevel>${n.lvl}</OutlineLevel><Priority>500</Priority><Start>${ini}</Start>`;
      if (tp !== 'summary') {
        const h = tp === 'ms' ? 0 : Math.max(1, +n.dur || 1) * 8;
        x += `<Duration>PT${h}H0M0S</Duration><DurationFormat>7</DurationFormat><Estimated>0</Estimated>`;
      }
      x += `<Milestone>${tp === 'ms' ? 1 : 0}</Milestone><Summary>${tp === 'summary' ? 1 : 0}</Summary><ConstraintType>0</ConstraintType><CalendarUID>-1</CalendarUID>`;
      if (notas) x += `<Notes>${xmlEsc(notas)}</Notes>`;
      if (tp !== 'summary') {
        (n.preds || []).forEach((u) => {
          const j = a.findIndex((m) => m.uid === u);
          if (j >= 0) x += `<PredecessorLink><PredecessorUID>${j + 1}</PredecessorUID><Type>1</Type><CrossProject>0</CrossProject><LinkLag>0</LinkLag><LagFormat>7</LagFormat></PredecessorLink>`;
        });
      }
      x += '</Task>\n';
    });
    x += '</Tasks>\n</Project>\n';
    return x;
  }

  function exportarXML() {
    if (!AE.nodes.length) return say('Gere a EAP primeiro.');
    const erros = validar().filter((v) => v.nivel === 'erro');
    if (erros.length && !confirm(`Há ${erros.length} erro(s) na validação. Exportar mesmo assim?`)) return;
    baixar(nomeArquivoBase() + '_MSProject.xml', gerarMSPDI(), 'application/xml');
    say('XML do MS Project exportado. Abra no MS Project: Arquivo › Abrir › tipo "XML".');
  }

  function linhasTabela() {
    const edt = calcularEDT();
    const uid2edt = {};
    AE.nodes.forEach((n, i) => { uid2edt[n.uid] = edt[i]; });
    const sch = calcularDatas(AE.nodes, parseISO(AE.meta.inicio));
    return AE.nodes.map((n, i) => {
      const tp = tipoEfetivo(i);
      const s = sch[n.uid] || {};
      return {
        ID: i + 1, EDT: edt[i], Nivel: n.lvl, Nome: n.name,
        Tipo: i === 0 ? 'Projeto' : tp === 'summary' ? 'Resumo' : tp === 'ms' ? 'Marco' : 'Pacote de trabalho',
        Duracao_dias_uteis: tp === 'summary' ? (s.dur || '') : tp === 'ms' ? 0 : (+n.dur || 1),
        Predecessoras_EDT: (n.preds || []).map((u) => uid2edt[u]).filter(Boolean).join(';'),
        Inicio: s.ini ? brDate(s.ini) : '', Termino: s.fim ? brDate(s.fim) : '',
        Descricao: n.desc || '', Criterio_aceitacao: n.crit || '', Fonte: n.fonte || ''
      };
    });
  }

  function exportarCSV() {
    if (!AE.nodes.length) return say('Gere a EAP primeiro.');
    const rows = linhasTabela();
    const cab = Object.keys(rows[0]);
    const csv = '﻿' + [cab].concat(rows.map((r) => cab.map((k) => r[k]))).map((r) => r.map((c) => `"${String(c ?? '').replace(/"/g, '""')}"`).join(';')).join('\n');
    baixar(nomeArquivoBase() + '_EAP.csv', csv, 'text/csv;charset=utf-8');
    say('CSV exportado.');
  }

  function exportarXLSX() {
    if (!AE.nodes.length) return say('Gere a EAP primeiro.');
    if (!window.XLSX) return say('Biblioteca XLSX não carregada.');
    const wb = XLSX.utils.book_new();
    const rows = linhasTabela();
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [6, 10, 6, 60, 16, 10, 18, 11, 11, 50, 40, 25].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, 'EAP');
    const dic = rows.filter((r) => r.Tipo !== 'Resumo' && r.Tipo !== 'Projeto').map((r) => ({ EDT: r.EDT, Elemento: r.Nome, Tipo: r.Tipo, Descricao: r.Descricao, Criterio_aceitacao: r.Criterio_aceitacao, Fonte: r.Fonte, Duracao_dias_uteis: r.Duracao_dias_uteis }));
    const wsd = XLSX.utils.json_to_sheet(dic);
    wsd['!cols'] = [10, 50, 16, 60, 45, 25, 10].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsd, 'Dicionário da EAP');
    const E = AE.escopo || {};
    const esc1 = [['Projeto', AE.meta.nome], ['Tipo de documento', AE.meta.tipo], ['Objeto', E.objeto], ['Prazo (dias úteis)', AE.meta.prazoDias || ''], ['Início previsto', brDate(parseISO(AE.meta.inicio))], ['Modo de análise', AE.meta.modo], []];
    [['Prazos', E.prazos], ['Marcos contratuais', E.marcos], ['Critérios de aceitação', E.criterios], ['Premissas', E.premissas], ['Restrições', E.restricoes], ['Exclusões', E.exclusoes], ['Riscos e penalidades', E.riscos]]
      .forEach(([t, arr]) => { esc1.push([t]); (arr || []).forEach((x) => esc1.push(['', typeof x === 'string' ? x : JSON.stringify(x)])); esc1.push([]); });
    const wse = XLSX.utils.aoa_to_sheet(esc1);
    wse['!cols'] = [{ wch: 26 }, { wch: 120 }];
    XLSX.utils.book_append_sheet(wb, wse, 'Declaração de Escopo');
    XLSX.writeFile(wb, nomeArquivoBase() + '_EAP.xlsx');
    say('Excel exportado (EAP, Dicionário e Declaração de Escopo).');
  }

  function exportarJSON() {
    if (!AE.nodes.length) return say('Nada para salvar.');
    const obj = Object.assign({}, AE, { versao: 'GLC-AE-1', texto: undefined });
    baixar(nomeArquivoBase() + '_projeto_EAP.json', JSON.stringify(obj, null, 2), 'application/json');
    say('Projeto salvo em JSON.');
  }

  function abrirJSON(file) {
    file.text().then((t) => {
      const obj = JSON.parse(t);
      if (obj.versao === 'GLC-AE-1' && Array.isArray(obj.nodes)) {
        AE = Object.assign(novoProjeto(), obj);
        uidSeq = Math.max(0, ...AE.nodes.map((n) => n.uid)) + 1;
      } else {
        importarJSONIA(obj);
        AE.meta.modo = 'Importado (JSON)';
      }
      render();
      say('Projeto carregado.');
    }).catch((e) => say('JSON inválido: ' + e.message));
  }

  async function exportarMPP() {
    if (!AE.nodes.length) return say('Gere a EAP primeiro.');
    const base = typeof HELPER_BASE_URL !== 'undefined' ? HELPER_BASE_URL : 'http://127.0.0.1:8765';
    try {
      const h = await fetch(base + '/health').then((r) => r.json());
      if (!h.ok || !h.has_com) throw new Error('Helper online, mas sem MS Project/pywin32.');
    } catch (e) {
      return say('Helper local indisponível (' + e.message + '). Use o botão "MS Project (.xml)" — não precisa de helper.');
    }
    const edt = calcularEDT();
    const records = AE.nodes.map((n, i) => {
      const tp = tipoEfetivo(i);
      return {
        id: n.uid, nivel: n.lvl, edt: edt[i], nome: n.name,
        duracao: tp === 'summary' ? '' : (tp === 'ms' ? '0d' : (Math.max(1, +n.dur || 1) + 'd')),
        marco: tp === 'ms', predecessoras: tp === 'summary' ? '' : (n.preds || []).join(';'),
        notas: [n.desc && 'Descrição: ' + n.desc, n.crit && 'Critério de aceitação: ' + n.crit, n.fonte && 'Fonte: ' + n.fonte].filter(Boolean).join('\n'),
        text1: edt[i]
      };
    });
    try {
      say('Gerando .MPP no MS Project…');
      const r = await fetch(base + '/export/mpp', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ tipo: 'eap-pmbok', project_name: 'EAP - ' + (AE.meta.nome || 'Projeto'), inicio: AE.meta.inicio, records }) });
      const data = await r.json();
      if (!r.ok || !data.ok) throw new Error(data.error || 'falha no helper');
      const f = await fetch(data.download_url);
      if (!f.ok) throw new Error('arquivo não pôde ser baixado');
      baixar(data.file_name, await f.blob());
      say('.MPP gerado: ' + data.file_name);
    } catch (e) {
      say('Erro ao gerar .MPP: ' + e.message);
    }
  }

  // ---------------------------------------------------------------
  // Persistência local
  // ---------------------------------------------------------------
  function salvarLocal() {
    try { localStorage.setItem(LS_PROJ, JSON.stringify(Object.assign({}, AE, { texto: (AE.texto || '').slice(0, 400000) }))); } catch (e) { /* cota excedida: ignora */ }
  }
  function carregarLocal() {
    try {
      const raw = localStorage.getItem(LS_PROJ);
      if (!raw) return;
      const obj = JSON.parse(raw);
      if (obj && Array.isArray(obj.nodes)) {
        AE = Object.assign(novoProjeto(), obj);
        uidSeq = Math.max(0, ...AE.nodes.map((n) => n.uid)) + 1;
        if (AE.texto) $('ae-texto').value = AE.texto;
      }
    } catch (e) { /* ignora */ }
  }

  // ---------------------------------------------------------------
  // Inicialização / eventos
  // ---------------------------------------------------------------
  function init() {
    if (!$('tab-analise')) return;
    if (window.pdfjsLib && window.pdfjsLib.GlobalWorkerOptions) {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
    }
    try {
      const k = localStorage.getItem(LS_KEY);
      if (k) { $('ae-apikey').value = k; $('ae-lembrar').checked = true; }
      const m = localStorage.getItem(LS_MODEL);
      if (m) $('ae-modelo').value = m;
    } catch (e) { /* ignora */ }
    carregarLocal();
    atualizarContador();

    const drop = $('ae-drop');
    const inp = $('ae-file');
    drop.addEventListener('click', () => inp.click());
    inp.addEventListener('change', () => { aoSelecionarArquivos(inp.files); inp.value = ''; });
    ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', (e) => aoSelecionarArquivos(e.dataTransfer.files));
    $('ae-texto').addEventListener('input', atualizarContador);

    $('ae-nome').addEventListener('change', (e) => { AE.meta.nome = e.target.value.trim(); if (AE.nodes[0]) AE.nodes[0].name = AE.meta.nome || AE.nodes[0].name; render(); });
    $('ae-inicio').addEventListener('change', (e) => { AE.meta.inicio = e.target.value || AE.meta.inicio; render(); });

    $('ae-btn-regras').addEventListener('click', () => {
      const t = $('ae-texto').value;
      if (t.trim().length < 200) return say('Carregue ou cole a documentação primeiro (mín. 200 caracteres).');
      AE.meta.nome = $('ae-nome').value.trim();
      AE.meta.inicio = $('ae-inicio').value || AE.meta.inicio;
      const r = analisarRegras(t);
      render();
      say(r.avisoGenerico ? 'Escopo técnico não identificado com clareza: EAP genérica criada — revise.' : `EAP gerada: ${r.nCands} entregáveis identificados. Revise e ajuste.`);
      $('ae-resultado').scrollIntoView({ behavior: 'smooth' });
    });
    $('ae-btn-ia').addEventListener('click', analisarIA);
    $('ae-btn-prompt').addEventListener('click', () => {
      const t = $('ae-texto').value.trim();
      if (t.length < 200) return say('Carregue ou cole a documentação primeiro.');
      const p = montarPrompt(t);
      navigator.clipboard.writeText(p).then(() => say('Prompt copiado. Cole no Claude, copie o JSON da resposta e use "Importar JSON".'), () => baixar('prompt_EAP.txt', p, 'text/plain'));
    });
    $('ae-btn-importar').addEventListener('click', () => {
      const s = prompt('Cole aqui o JSON retornado pela IA:');
      if (!s) return;
      try { AE.texto = $('ae-texto').value; importarJSONIA(s); AE.meta.modo = 'IA (JSON importado)'; render(); say('EAP importada.'); } catch (e) { say('' + e.message); }
    });
    $('ae-btn-limpar').addEventListener('click', () => {
      if (!confirm('Limpar texto, EAP e análise desta aba?')) return;
      AE = novoProjeto(); $('ae-texto').value = ''; $('ae-arqs').innerHTML = ''; atualizarContador();
      try { localStorage.removeItem(LS_PROJ); } catch (e) { /* ignora */ }
      render();
    });

    $('ae-arvore').addEventListener('change', (e) => {
      const el = e.target.closest('.ae-in'); if (!el) return;
      const i = +el.closest('tr').dataset.i;
      aplicarCampo(i, el.dataset.k, el.value);
      if (i === 0 && el.dataset.k === 'name') AE.meta.nome = el.value;
      agendarRender();
    });
    $('ae-arvore').addEventListener('click', (e) => {
      const b = e.target.closest('button[data-a]'); if (!b) return;
      const tr = b.closest('tr');
      setTimeout(() => acao(+tr.dataset.i, b.dataset.a), 0);
    });
    $('ae-dic').addEventListener('change', (e) => {
      const el = e.target.closest('.ae-in'); if (!el) return;
      aplicarCampo(+el.closest('tr').dataset.i, el.dataset.k, el.value);
      salvarLocal();
      $('ae-val').dataset.dirty = '1';
    });

    $('ae-x-xml').addEventListener('click', exportarXML);
    $('ae-x-mpp').addEventListener('click', exportarMPP);
    $('ae-x-xlsx').addEventListener('click', exportarXLSX);
    $('ae-x-csv').addEventListener('click', exportarCSV);
    $('ae-x-json').addEventListener('click', exportarJSON);
    $('ae-x-abrir').addEventListener('click', () => $('ae-json-file').click());
    $('ae-json-file').addEventListener('change', (e) => { if (e.target.files[0]) abrirJSON(e.target.files[0]); e.target.value = ''; });

    render();
  }

  // API pública (testes e integração)
  window.AnaliseEAP = { analisarRegras, importarJSONIA, gerarMSPDI, validar, calcularEDT, montarPrompt, get estado() { return AE; }, set estado(v) { AE = v; }, render };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
