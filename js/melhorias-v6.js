/* =====================================================================
 * Melhorias v6.0 do Gerador EAP
 *  - Backup da base de códigos (Dados Unificados): exportar/importar JSON
 *  - Restaurar a base padrão
 * ===================================================================== */
(function () {
  'use strict';

  function baixarJSON(nome, obj) {
    const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = nome; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  function limparFlags(obj) {
    // remove marcadores internos de edição (__editing)
    return JSON.parse(JSON.stringify(obj, (k, v) => (k === '__editing' ? undefined : v)));
  }

  function reaplicarTela() {
    if (typeof rDados === 'function') rDados();
    const cur = document.getElementById('iseg') && document.getElementById('iseg').value;
    if (cur && typeof rSel === 'function') { rSel(cur); if (typeof uDrops === 'function') uDrops(); }
  }

  window.segExportarJSON = function () {
    const data = new Date().toISOString().slice(0, 10);
    baixarJSON('base_codigos_LD_' + data + '.json', { versao: 'GLC-SEG-1', exportadoEm: new Date().toISOString(), segmentos: limparFlags(SEG) });
    toast('Base de códigos exportada.');
  };

  window.segImportarJSON = function (input) {
    const f = input.files && input.files[0];
    input.value = '';
    if (!f) return;
    f.text().then((t) => {
      const obj = JSON.parse(t);
      const segs = obj && obj.segmentos ? obj.segmentos : obj;
      if (!segs || typeof segs !== 'object') throw new Error('estrutura inesperada');
      const campos = ['FASE', 'AREA', 'SUBAREA', 'LOCAL', 'DISC', 'SIST', 'TIPO'];
      let n = 0;
      Object.keys(segs).forEach((seg) => {
        if (!segs[seg] || typeof segs[seg] !== 'object') return;
        SEG[seg] = SEG[seg] || {};
        campos.forEach((c) => {
          if (Array.isArray(segs[seg][c])) {
            SEG[seg][c] = segs[seg][c].filter((x) => x && (x.code != null || x.desc != null)).map((x) => ({ code: String(x.code ?? '').trim(), desc: String(x.desc ?? '').trim() }));
            n++;
          }
        });
      });
      if (!n) throw new Error('nenhum campo reconhecido (FASE, AREA, SUBAREA, LOCAL, DISC, SIST, TIPO)');
      saveSEG();
      reaplicarTela();
      toast('Base importada (' + n + ' listas atualizadas).');
    }).catch((e) => toast('Falha ao importar: ' + e.message));
  };

  window.segRestaurarPadrao = function () {
    if (!confirm('Restaurar a base de códigos original? Todas as edições feitas em "Dados Unificados" serão perdidas.\n\nDica: exporte um backup antes.')) return;
    Object.keys(SEG).forEach((k) => delete SEG[k]);
    Object.assign(SEG, JSON.parse(JSON.stringify(SEG_PADRAO)));
    try { localStorage.removeItem(LS_KEY); } catch (e) { /* ignora */ }
    reaplicarTela();
    toast('Base padrão restaurada.');
  };

})();
