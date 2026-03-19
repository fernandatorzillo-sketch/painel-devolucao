(function(){
  const payload = window.__PAINEL_DEVOLUCOES__ || {};
  const meta = payload.meta || {};
  const MONTHS = [
    { value:'01', label:'Jan' },
    { value:'02', label:'Fev' },
    { value:'03', label:'Mar' },
    { value:'04', label:'Abr' },
    { value:'05', label:'Mai' },
    { value:'06', label:'Jun' },
    { value:'07', label:'Jul' },
    { value:'08', label:'Ago' },
    { value:'09', label:'Set' },
    { value:'10', label:'Out' },
    { value:'11', label:'Nov' },
    { value:'12', label:'Dez' }
  ];
  const MONTH_LABEL = MONTHS.reduce(function(acc, item){ acc[item.value] = item.label; return acc; }, {});
  const SIZE_MAP = {
    'PP':'PP/XS','PP/XS':'PP/XS','P':'P/S','P/S':'P/S','M':'M/M','M/M':'M/M',
    'G':'G/L','G/L':'G/L','GG':'GG/XL','GG/XL':'GG/XL','XG':'XG/XXL','EG':'XG/XXL',
    'EXG':'XG/XXL','XG/XXL':'XG/XXL','U':'U/ONE SIZE','UN':'U/ONE SIZE',
    'UNICO':'U/ONE SIZE','ONE SIZE':'U/ONE SIZE','U/ONE SIZE':'U/ONE SIZE'
  };
  const moneyFmt = new Intl.NumberFormat('pt-BR', { style:'currency', currency:'BRL' });
  const intFmt = new Intl.NumberFormat('pt-BR', { maximumFractionDigits:0 });
  const qtyFmt = new Intl.NumberFormat('pt-BR', { minimumFractionDigits:0, maximumFractionDigits:3 });
  const pctFmt = new Intl.NumberFormat('pt-BR', { minimumFractionDigits:2, maximumFractionDigits:2 });
  function rawText(value){ return String(value == null ? '' : value).trim(); }
  function normalizeText(value, fallback){ const text = rawText(value); return text || fallback; }
  function buildCategoryLabel(group, subgroup){
    const mainGroup = normalizeText(group, 'SEM_GRUPO');
    const mainSubgroup = normalizeText(subgroup, 'SEM_SUBGRUPO');
    if(mainGroup === 'SEM_GRUPO' && mainSubgroup === 'SEM_SUBGRUPO'){ return 'SEM_GRUPO / SEM_SUBGRUPO'; }
    if(mainGroup === 'SEM_GRUPO'){ return mainSubgroup; }
    if(mainSubgroup === 'SEM_SUBGRUPO'){ return mainGroup; }
    return mainGroup + ' / ' + mainSubgroup;
  }
  function safeNum(value){ const num = Number(value); return Number.isFinite(num) ? num : 0; }
  function round(value, digits){ const factor = Math.pow(10, digits || 0); return Math.round((safeNum(value) + Number.EPSILON) * factor) / factor; }
  function truthyFlag(value){ if(value === true || value === 1 || value === '1'){ return true; } const text = rawText(value).toUpperCase(); return text === 'TRUE' || text === 'VERDADEIRO'; }
  function normalizeSize(value){ const text = normalizeText(value, 'SEM_TAMANHO').toUpperCase(); return SIZE_MAP[text] || text; }
  function rateOrNull(numerator, denominator){ return denominator > 0 ? round((numerator / denominator) * 100, 2) : null; }
  function formatMoney(value){ return moneyFmt.format(safeNum(value)); }
  function formatQty(value){ const num = safeNum(value); return Math.abs(num - Math.round(num)) < 0.0001 ? intFmt.format(Math.round(num)) : qtyFmt.format(num); }
  function formatPercent(value){ if(value === null || value === undefined || value === ''){ return 'Base insuficiente'; } const num = Number(value); return Number.isFinite(num) ? pctFmt.format(num) + '%' : 'Base insuficiente'; }
  function formatDate(iso){ const text = rawText(iso); if(!text){ return ''; } const parts = text.split('-'); return parts.length === 3 ? (parts[2] + '/' + parts[1] + '/' + parts[0]) : text; }
  function clipText(value, max){ const text = rawText(value); return text.length <= max ? text : text.slice(0, Math.max(0, max - 1)) + '...'; }
  function escapeHtml(value){
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/\"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function compareLabel(a, b){
    const left = rawText(a);
    const right = rawText(b);
    const leftSem = /^SEM_/.test(left);
    const rightSem = /^SEM_/.test(right);
    if(leftSem && !rightSem){ return 1; }
    if(!leftSem && rightSem){ return -1; }
    return left.localeCompare(right, 'pt-BR');
  }
  function splitIso(iso){
    const parts = String(iso || '').split('-').map(function(part){ return Number(part); });
    if(parts.length !== 3 || parts.some(function(part){ return !Number.isFinite(part); })){ return null; }
    return { year:parts[0], month:parts[1], day:parts[2] };
  }
  function daysInMonth(year, month){ return new Date(Date.UTC(year, month, 0)).getUTCDate(); }
  function toIsoDate(date){ return [String(date.getUTCFullYear()).padStart(4, '0'), String(date.getUTCMonth() + 1).padStart(2, '0'), String(date.getUTCDate()).padStart(2, '0')].join('-'); }
  function toUtcDate(iso){ const parts = splitIso(iso); return parts ? new Date(Date.UTC(parts.year, parts.month - 1, parts.day)) : null; }
  function addDaysIso(iso, delta){ const date = toUtcDate(iso); if(!date){ return iso; } date.setUTCDate(date.getUTCDate() + delta); return toIsoDate(date); }
  function addMonthsIso(iso, delta){ const parts = splitIso(iso); if(!parts){ return iso; } const date = new Date(Date.UTC(parts.year, parts.month - 1 + delta, 1)); date.setUTCDate(Math.min(parts.day, daysInMonth(date.getUTCFullYear(), date.getUTCMonth() + 1))); return toIsoDate(date); }
  function shiftYearIso(iso, delta){ const parts = splitIso(iso); if(!parts){ return iso; } const year = parts.year + delta; const day = Math.min(parts.day, daysInMonth(year, parts.month)); return [String(year).padStart(4, '0'), String(parts.month).padStart(2, '0'), String(day).padStart(2, '0')].join('-'); }
  function inRange(iso, start, end){ return rawText(iso) >= rawText(start) && rawText(iso) <= rawText(end); }
  function deltaMaybe(current, prior){ return current !== null && prior !== null ? round(current - prior, 2) : null; }
  function shareText(count, total){ const base = safeNum(total); return base ? formatPercent((safeNum(count) / base) * 100) : '0%'; }
  function buildPair(pedido, sku){ const left = rawText(pedido); const right = rawText(sku); return left && right ? (left + '|' + right) : ''; }
  function emptyHtml(text){ return '<div class="empty">' + escapeHtml(text || 'Sem dados para o filtro atual.') + '</div>'; }
  function tagHtml(label, tone){ return label ? '<span class="tag ' + escapeHtml(tone || 'slate') + '">' + escapeHtml(label) + '</span>' : ''; }
  function toneHex(tone, fallback){
    if(tone === 'red'){ return '#b42318'; }
    if(tone === 'amber'){ return '#b45309'; }
    if(tone === 'green'){ return '#15803d'; }
    if(tone === 'blue'){ return '#2563eb'; }
    if(tone === 'slate'){ return '#64748b'; }
    return fallback || '#2563eb';
  }
  function criticalityColor(row, fallback){
    if(!row){ return toneHex('slate', fallback); }
    if(row.Alerta === 'Base insuficiente' || row.BaseStatus === 'BASE_BAIXA'){ return toneHex('slate', fallback); }
    return toneHex(row.AlertTone || 'green', fallback);
  }
  function criticalityTagHtml(row){
    if(!row){ return '-'; }
    if(row.Alerta === 'Base insuficiente'){ return tagHtml('BASE INSUFICIENTE', 'slate'); }
    if(row.BaseStatus === 'BASE_BAIXA'){ return tagHtml('BASE BAIXA', 'slate'); }
    return tagHtml(row.ClassificacaoCriticidade || 'SAUDAVEL', row.AlertTone || 'green');
  }
  function sortDesc(fieldA, fieldB){
    return function(left, right){
      const primary = safeNum(right[fieldA]) - safeNum(left[fieldA]);
      if(primary !== 0){ return primary; }
      if(fieldB){
        const secondary = safeNum(right[fieldB]) - safeNum(left[fieldB]);
        if(secondary !== 0){ return secondary; }
      }
      return compareLabel(left.Label || left.Produto || left.Cliente || left.Key, right.Label || right.Produto || right.Cliente || right.Key);
    };
  }
  function sortByCriticality(left, right){
    const primary = safeNum(right.MaxTaxaPct) - safeNum(left.MaxTaxaPct);
    if(primary !== 0){ return primary; }
    const secondary = safeNum(right.ValorDevolvido) - safeNum(left.ValorDevolvido);
    if(secondary !== 0){ return secondary; }
    const tertiary = safeNum(right.QuantidadeDevolvida) - safeNum(left.QuantidadeDevolvida);
    if(tertiary !== 0){ return tertiary; }
    return compareLabel(left.Label || left.Produto || left.Cliente || left.Key, right.Label || right.Produto || right.Cliente || right.Key);
  }
  function uniqueSorted(values){ return Array.from(new Set(values.filter(function(value){ return rawText(value); }))).sort(compareLabel); }
  function setSelectOptions(select, options, includeAll){
    const rows = [];
    if(includeAll !== false){ rows.push('<option value="TODOS">Todos</option>'); }
    options.forEach(function(option){ rows.push('<option value="' + escapeHtml(option.value) + '">' + escapeHtml(option.label) + '</option>'); });
    select.innerHTML = rows.join('');
  }
  function tableHtml(columns, rows, emptyText){
    if(!rows || !rows.length){ return emptyHtml(emptyText || 'Sem dados para o filtro atual.'); }
    const head = '<tr>' + columns.map(function(column){ return '<th>' + escapeHtml(column.label) + '</th>'; }).join('') + '</tr>';
    const body = rows.map(function(row){
      return '<tr>' + columns.map(function(column){
        const value = column.render ? column.render(row) : row[column.key];
        return column.html ? '<td>' + value + '</td>' : '<td>' + escapeHtml(value == null ? '' : value) + '</td>';
      }).join('') + '</tr>';
    }).join('');
    return '<table><thead>' + head + '</thead><tbody>' + body + '</tbody></table>';
  }
  function renderTable(targetId, columns, rows, emptyText){ const el = document.getElementById(targetId); if(el){ el.innerHTML = tableHtml(columns, rows, emptyText); } }
  function renderBars(targetId, items, options){
    const el = document.getElementById(targetId);
    if(!el){ return; }
    if(!items || !items.length){ el.innerHTML = emptyHtml(options && options.emptyText ? options.emptyText : 'Sem dados para o filtro atual.'); return; }
    const valueFn = options && options.value ? options.value : function(item){ return item.Value; };
    const labelFn = options && options.label ? options.label : function(item){ return item.Label; };
    const metaFn = options && options.meta ? options.meta : function(){ return ''; };
    const colorFn = options && options.color ? options.color : function(){ return '#2563eb'; };
    const maxValue = Math.max.apply(null, items.map(function(item){ return safeNum(valueFn(item)); }).concat([1]));
    const html = items.map(function(item){
      const value = safeNum(valueFn(item));
      const width = Math.max(3, Math.round((value / maxValue) * 100));
      const barValue = options && options.format ? options.format(value, item) : formatQty(value);
      return '<div class="bar-row">'
        + '<div class="bar-label"><strong>' + escapeHtml(labelFn(item)) + '</strong>' + (metaFn(item) ? '<span class="bar-meta">' + escapeHtml(metaFn(item)) + '</span>' : '') + '</div>'
        + '<div class="bar-track"><div class="bar-fill" style="width:' + width + '%;background:' + colorFn(item) + '"></div></div>'
        + '<div class="bar-value">' + escapeHtml(barValue) + '</div>'
      + '</div>';
    }).join('');
    el.innerHTML = '<div class="bars">' + html + '</div>';
  }
  function kpiHtml(title, value, sub, tone){
    return '<div class="kpi ' + escapeHtml(tone || 'blue') + '"><div class="kpi-title">' + escapeHtml(title) + '</div><div class="kpi-value">' + escapeHtml(value) + '</div><div class="kpi-sub">' + escapeHtml(sub || '') + '</div></div>';
  }
  function buildFinanceSummary(rows){
    const summary = { ValorVendido:0, ValorDevolvido:0, QuantidadeVendida:0, QuantidadeDevolvida:0 };
    rows.forEach(function(row){
      summary.ValorVendido += safeNum(row.ValorVendido);
      summary.ValorDevolvido += safeNum(row.ValorDevolvido);
      summary.QuantidadeVendida += safeNum(row.QuantidadeVendida);
      summary.QuantidadeDevolvida += safeNum(row.QuantidadeDevolvida);
    });
    summary.ValorVendido = round(summary.ValorVendido, 2);
    summary.ValorDevolvido = round(summary.ValorDevolvido, 2);
    summary.QuantidadeVendida = round(summary.QuantidadeVendida, 3);
    summary.QuantidadeDevolvida = round(summary.QuantidadeDevolvida, 3);
    summary.TaxaValorPct = rateOrNull(summary.ValorDevolvido, summary.ValorVendido);
    summary.TaxaQuantidadePct = rateOrNull(summary.QuantidadeDevolvida, summary.QuantidadeVendida);
    return summary;
  }
  function addFinanceMetrics(bucket, row){
    bucket.QuantidadeVendida += safeNum(row.QuantidadeVendida);
    bucket.ValorVendido += safeNum(row.ValorVendido);
    bucket.QuantidadeDevolvida += safeNum(row.QuantidadeDevolvida);
    bucket.ValorDevolvido += safeNum(row.ValorDevolvido);
  }
  function buildFinanceGroups(rows, keyFn, seedFn){
    const map = new Map();
    rows.forEach(function(row){
      const key = keyFn(row);
      if(!map.has(key)){ map.set(key, seedFn(row, key)); }
      addFinanceMetrics(map.get(key), row);
    });
    return Array.from(map.values()).map(function(row){
      row.QuantidadeVendida = round(row.QuantidadeVendida, 3);
      row.ValorVendido = round(row.ValorVendido, 2);
      row.QuantidadeDevolvida = round(row.QuantidadeDevolvida, 3);
      row.ValorDevolvido = round(row.ValorDevolvido, 2);
      row.TaxaValorPct = rateOrNull(row.ValorDevolvido, row.ValorVendido);
      row.TaxaQuantidadePct = rateOrNull(row.QuantidadeDevolvida, row.QuantidadeVendida);
      return row;
    });
  }
  function blankMetrics(base){ return { Key:base.Key, QuantidadeVendida:0, ValorVendido:0, QuantidadeDevolvida:0, ValorDevolvido:0, TaxaValorPct:null, TaxaQuantidadePct:null }; }
  function hasSalesBase(row){ return safeNum(row.QuantidadeVendida) > 0 || safeNum(row.ValorVendido) > 0; }
  function criticalShareText(criticalCount, totalCount){
    const total = safeNum(totalCount);
    return total ? (formatPercent((safeNum(criticalCount) / total) * 100) + ' do total.') : 'Base insuficiente';
  }
  function meetsBaseMinima(row, dimensionType){
    if(dimensionType === 'cliente'){ return safeNum(row.QuantidadeVendida) >= 2 || safeNum(row.ValorVendido) >= 200; }
    if(dimensionType === 'produto'){ return safeNum(row.QuantidadeVendida) >= 5; }
    if(dimensionType === 'tamanho'){ return safeNum(row.QuantidadeVendida) >= 10; }
    return true;
  }
  function classifyAlert(row, rollingFull, dimensionType){
    const maxRate = Math.max(safeNum(row.TaxaValorPct), safeNum(row.TaxaQuantidadePct));
    const baseMinima = meetsBaseMinima(row, dimensionType);
    const classificacao = maxRate >= 80 ? 'CRITICO' : (maxRate >= 50 ? 'ATENCAO' : 'SAUDAVEL');
    row.BaseMinima = baseMinima;
    row.MaxTaxaPct = maxRate;
    row.ClassificacaoCriticidade = classificacao;
    row.DimensaoCriticidade = dimensionType || '';
    row.BaseStatus = baseMinima ? 'BASE_OK' : 'BASE_BAIXA';
    if(!rollingFull){ row.Alerta = 'Base insuficiente'; row.AlertTone = 'slate'; return row; }
    if(!baseMinima){ row.Alerta = 'Base baixa'; row.AlertTone = 'slate'; return row; }
    row.Alerta = classificacao;
    row.AlertTone = classificacao === 'CRITICO' ? 'red' : (classificacao === 'ATENCAO' ? 'amber' : 'green');
    return row;
  }
  function buildMergedDimension(rollingRows, periodRows, priorRows, keyFn, seedFn, rollingFull, dimensionType){
    const rolling = buildFinanceGroups(rollingRows, keyFn, seedFn);
    const period = buildFinanceGroups(periodRows, keyFn, seedFn);
    const prior = buildFinanceGroups(priorRows, keyFn, seedFn);
    const rollingMap = new Map(rolling.map(function(row){ return [row.Key, row]; }));
    const periodMap = new Map(period.map(function(row){ return [row.Key, row]; }));
    const priorMap = new Map(prior.map(function(row){ return [row.Key, row]; }));
    const keys = Array.from(new Set([].concat(Array.from(rollingMap.keys()), Array.from(periodMap.keys()), Array.from(priorMap.keys()))));
    return keys.map(function(key){
      const base = rollingMap.get(key) || periodMap.get(key) || priorMap.get(key);
      const roll = rollingMap.get(key) || blankMetrics(base);
      const current = periodMap.get(key) || blankMetrics(base);
      const previous = priorMap.get(key) || blankMetrics(base);
      const merged = Object.assign({}, base, {
        QuantidadeVendida:roll.QuantidadeVendida,
        ValorVendido:roll.ValorVendido,
        QuantidadeDevolvida:roll.QuantidadeDevolvida,
        ValorDevolvido:roll.ValorDevolvido,
        TaxaValorPct:roll.TaxaValorPct,
        TaxaQuantidadePct:roll.TaxaQuantidadePct,
        PeriodQuantidadeVendida:current.QuantidadeVendida,
        PeriodValorVendido:current.ValorVendido,
        PeriodQuantidadeDevolvida:current.QuantidadeDevolvida,
        PeriodValorDevolvido:current.ValorDevolvido,
        PeriodTaxaValorPct:current.TaxaValorPct,
        PeriodTaxaQuantidadePct:current.TaxaQuantidadePct,
        PriorQuantidadeVendida:previous.QuantidadeVendida,
        PriorValorVendido:previous.ValorVendido,
        PriorQuantidadeDevolvida:previous.QuantidadeDevolvida,
        PriorValorDevolvido:previous.ValorDevolvido,
        PriorTaxaValorPct:previous.TaxaValorPct,
        PriorTaxaQuantidadePct:previous.TaxaQuantidadePct,
        DeltaPeriodValorVendido:round(current.ValorVendido - previous.ValorVendido, 2),
        DeltaPeriodValorDevolvido:round(current.ValorDevolvido - previous.ValorDevolvido, 2),
        DeltaPeriodTaxaValorPct:deltaMaybe(current.TaxaValorPct, previous.TaxaValorPct),
        DeltaPeriodTaxaQuantidadePct:deltaMaybe(current.TaxaQuantidadePct, previous.TaxaQuantidadePct)
      });
      return classifyAlert(merged, rollingFull, dimensionType);
    }).filter(function(row){
      return safeNum(row.QuantidadeVendida) > 0 || safeNum(row.ValorVendido) > 0 || safeNum(row.QuantidadeDevolvida) > 0 || safeNum(row.ValorDevolvido) > 0 || safeNum(row.PeriodValorDevolvido) > 0 || safeNum(row.PeriodQuantidadeDevolvida) > 0;
    });
  }
  function buildOperationalCounts(rows, keyFn, seedFn){
    const map = new Map();
    rows.forEach(function(row){
      const key = keyFn(row);
      if(!map.has(key)){ map.set(key, seedFn(row, key)); }
      const bucket = map.get(key);
      bucket.Count += 1;
      bucket.QuantidadeSolicitada += safeNum(row.QuantidadeSolicitada);
    });
    return Array.from(map.values()).map(function(row){ row.QuantidadeSolicitada = round(row.QuantidadeSolicitada, 3); return row; });
  }

  const financeData = (payload.finance || []).map(function(row){
    const pedidoKey = rawText(row.Pedido);
    const skuKey = rawText(row.Sku);
    const flagSale = truthyFlag(row.FlagSale);
    return {
      AnoBase:Number(row.AnoBase || 0),
      Fonte:normalizeText(row.Fonte, 'SEM_FONTE'),
      Pedido:pedidoKey,
      PedidoKey:pedidoKey,
      ClienteId:normalizeText(row.ClienteId, 'CLIENTE_SEM_ID'),
      ClienteLabel:normalizeText(row.ClienteLabel, normalizeText(row.ClienteId, 'CLIENTE_SEM_ID')),
      Data:rawText(row.Data),
      Sku:normalizeText(row.Sku, 'SEM_SKU'),
      SkuKey:skuKey,
      Produto:normalizeText(row.Produto, normalizeText(row.Sku, 'SEM_PRODUTO')),
      Departamento:normalizeText(row.Departamento, 'SEM_DEPARTAMENTO'),
      Grupo:normalizeText(row.Grupo, 'SEM_GRUPO'),
      Subgrupo:normalizeText(row.Subgrupo, 'SEM_SUBGRUPO'),
      Categoria:normalizeText(row.Categoria, 'SEM_CATEGORIA'),
      CategoriaPainel:buildCategoryLabel(row.Grupo, row.Subgrupo),
      Genero:normalizeText(row.Genero, 'SEM_GENERO'),
      Tamanho:normalizeSize(row.Tamanho),
      Cor:normalizeText(row.Cor, 'SEM_COR'),
      Colecao:normalizeText(row.Colecao, 'SEM_COLECAO'),
      CodEmpresa:normalizeText(row.CodEmpresa, ''),
      Canal:normalizeText(row.Canal, 'SEM_CANAL'),
      Loja:normalizeText(row.Loja || row.Canal, 'SEM_LOJA'),
      CodOperacao:normalizeText(row.CodOperacao, 'SEM_CODIGO'),
      TipoOperacao:normalizeText(row.TipoOperacao, 'SEM_TIPO'),
      QuantidadeVendida:round(row.QuantidadeVendida || 0, 3),
      ValorVendido:round(row.ValorVendido || 0, 2),
      QuantidadeDevolvida:round(row.QuantidadeDevolvida || 0, 3),
      ValorDevolvido:round(row.ValorDevolvido || 0, 2),
      FlagSale:flagSale,
      StatusSale:normalizeText(row.StatusSale, flagSale ? 'SALE' : 'NAO_SALE'),
      SaleBucket:flagSale ? 'SALE' : 'NAO_SALE',
      Year:rawText(row.Data).slice(0, 4),
      Month:rawText(row.Data).slice(5, 7),
      Day:rawText(row.Data).slice(8, 10),
      PedidoSku:buildPair(pedidoKey, skuKey)
    };
  }).filter(function(row){ return rawText(row.Data); });

  const operationalData = (payload.operational || []).map(function(row){
    const pedidoKey = rawText(row.Pedido);
    const skuKey = rawText(row.Sku);
    const flagSale = truthyFlag(row.FlagSale);
    return {
      Plataforma:normalizeText(row.Plataforma, 'SEM_PLATAFORMA'),
      IdSolicitacao:normalizeText(row.IdSolicitacao, 'SEM_ID'),
      Pedido:pedidoKey,
      PedidoKey:pedidoKey,
      Cliente:normalizeText(row.Cliente, 'CLIENTE_SEM_NOME'),
      Documento:normalizeText(row.Documento, ''),
      Data:rawText(row.Data),
      Sku:normalizeText(row.Sku, 'SEM_SKU'),
      SkuKey:skuKey,
      Produto:normalizeText(row.Produto, normalizeText(row.Sku, 'SEM_PRODUTO')),
      Departamento:normalizeText(row.Departamento, 'SEM_DEPARTAMENTO'),
      Grupo:normalizeText(row.Grupo, 'SEM_GRUPO'),
      Subgrupo:normalizeText(row.Subgrupo, 'SEM_SUBGRUPO'),
      Categoria:normalizeText(row.Categoria, 'SEM_CATEGORIA'),
      CategoriaPainel:buildCategoryLabel(row.Grupo, row.Subgrupo),
      Genero:normalizeText(row.Genero, 'SEM_GENERO'),
      Tamanho:normalizeSize(row.Tamanho),
      Cor:normalizeText(row.Cor, 'SEM_COR'),
      Colecao:normalizeText(row.Colecao, 'SEM_COLECAO'),
      Loja:normalizeText(row.Loja, 'SEM_LOJA'),
      QuantidadeSolicitada:round(row.QuantidadeSolicitada || 0, 3),
      Motivo:normalizeText(row.Motivo, 'SEM_MOTIVO'),
      CategoriaComentario:normalizeText(row.CategoriaComentario, 'SEM_COMENTARIO'),
      StatusSolicitacao:normalizeText(row.StatusSolicitacao, 'SEM_STATUS'),
      Estagio:normalizeText(row.Estagio, 'SEM_ESTAGIO'),
      FlagSale:flagSale,
      StatusSale:normalizeText(row.StatusSale, flagSale ? 'SALE' : 'NAO_SALE'),
      SaleBucket:flagSale ? 'SALE' : 'NAO_SALE',
      Comentario:normalizeText(row.Comentario, ''),
      Year:rawText(row.Data).slice(0, 4),
      Month:rawText(row.Data).slice(5, 7),
      Day:rawText(row.Data).slice(8, 10),
      PedidoSku:buildPair(pedidoKey, skuKey)
    };
  }).filter(function(row){ return rawText(row.Data); });

  const inferredMinDate = [
    financeData.reduce(function(acc, row){ return !acc || row.Data < acc ? row.Data : acc; }, ''),
    operationalData.reduce(function(acc, row){ return !acc || row.Data < acc ? row.Data : acc; }, '')
  ].filter(Boolean).sort()[0] || '';
  const inferredMaxDate = [
    financeData.reduce(function(acc, row){ return !acc || row.Data > acc ? row.Data : acc; }, ''),
    operationalData.reduce(function(acc, row){ return !acc || row.Data > acc ? row.Data : acc; }, '')
  ].filter(Boolean).sort().slice(-1)[0] || '';
  const dataMin = rawText(meta.minStart) || inferredMinDate;
  const dataMax = rawText(meta.maxEnd) || inferredMaxDate;
  const defaultStart = rawText(meta.defaultStart) || dataMin;
  const defaultEnd = rawText(meta.defaultEnd) || dataMax;
  const currentYear = Number(meta.currentYear || (splitIso(dataMax) ? splitIso(dataMax).year : new Date().getFullYear()));
  const priorYear = Number(meta.priorYear || (currentYear - 1));
  const earliestFullRollingEnd = dataMin ? addDaysIso(addMonthsIso(dataMin, 12), -1) : dataMax;
  const fixedPeriodStart = rawText(meta.fixedPeriodStart) || '2025-01-01';
  const fixedPeriodEndSource = rawText(meta.fixedPeriodEnd) || dataMax;
  const fixedPeriodEnd = dataMax && fixedPeriodEndSource > dataMax ? dataMax : fixedPeriodEndSource;
  const fixedPeriodLabel = rawText(meta.fixedPeriodLabel) || '01/2025 ate atual';
  const fixedPeriodSubtitle = rawText(meta.fixedPeriodSubtitle) || ('Periodo: ' + formatDate(fixedPeriodStart) + ' ate ' + formatDate(fixedPeriodEnd));
  const fixedPeriodHistoryComplete = meta.fixedPeriodHistoryComplete === false ? false : (!dataMin || dataMin <= fixedPeriodStart);
  const filterMin = fixedPeriodStart;
  const filterMax = fixedPeriodEnd;

  function buildFilterOptions(){
    const years = uniqueSorted([].concat(financeData.map(function(row){ return row.Year; }), operationalData.map(function(row){ return row.Year; }))).map(function(year){ return { value:year, label:year }; });
    const collections = uniqueSorted([].concat(financeData.map(function(row){ return row.Colecao; }), operationalData.map(function(row){ return row.Colecao; }))).map(function(item){ return { value:item, label:item }; });
    const categories = uniqueSorted([].concat(financeData.map(function(row){ return row.CategoriaPainel; }), operationalData.map(function(row){ return row.CategoriaPainel; }))).map(function(item){ return { value:item, label:item }; });
    const stores = uniqueSorted([].concat(financeData.map(function(row){ return row.Loja; }), operationalData.map(function(row){ return row.Loja; }))).map(function(item){ return { value:item, label:item }; });
    const sizes = uniqueSorted([].concat(financeData.map(function(row){ return row.Tamanho; }), operationalData.map(function(row){ return row.Tamanho; }))).map(function(item){ return { value:item, label:item }; });
    const motives = uniqueSorted(operationalData.map(function(row){ return row.Motivo; })).map(function(item){ return { value:item, label:item }; });
    const statuses = uniqueSorted(operationalData.map(function(row){ return row.StatusSolicitacao; })).map(function(item){ return { value:item, label:item }; });
    const platforms = uniqueSorted(operationalData.map(function(row){ return row.Plataforma; })).map(function(item){ return { value:item, label:item }; });
    const commentCategories = uniqueSorted(operationalData.map(function(row){ return row.CategoriaComentario; })).map(function(item){ return { value:item, label:item }; });
    return { years:years, months:MONTHS, collections:collections, categories:categories, stores:stores, sizes:sizes, motives:motives, statuses:statuses, platforms:platforms, commentCategories:commentCategories };
  }

  function ensureStyle(){
    if(document.getElementById('painel-devolucoes-style')){ return; }
    const style = document.createElement('style');
    style.id = 'painel-devolucoes-style';
    style.textContent = ""
      + ":root{--bg:#f4f4ef;--ink:#142126;--muted:#5d6d73;--line:#d6ddd7;--card:#fffdf8;--blue:#0f766e;--red:#b42318;--amber:#b45309;--green:#166534;--slate:#475569;--sky:#eff8ff}"
      + "*{box-sizing:border-box}"
      + "body{margin:0;background:radial-gradient(circle at top left,#fffdf6,#f3f4ee 38%,#f4f4ef 100%);color:var(--ink);font-family:Segoe UI,Tahoma,sans-serif}"
      + ".wrap{padding:24px 24px 40px}.head{display:grid;grid-template-columns:minmax(280px,1.3fr) minmax(320px,1fr);gap:18px;align-items:start}"
      + ".title{font-size:30px;font-weight:800;letter-spacing:-.02em}.sub{font-size:13px;color:var(--muted);margin-top:6px;line-height:1.5}"
      + ".pill-row{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}.pill{padding:7px 12px;border:1px solid var(--line);border-radius:999px;background:#fff;font-size:12px;font-weight:700;color:#314249}"
      + ".filter-panel{background:rgba(255,253,248,.96);border:1px solid var(--line);border-radius:20px;padding:16px;box-shadow:0 16px 42px rgba(20,33,38,.06)}"
      + ".filter-title{font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:#52646b;font-weight:800}.filter-grid{display:grid;grid-template-columns:repeat(4,minmax(130px,1fr));gap:12px;margin-top:12px}"
      + ".filter-field{display:flex;flex-direction:column;gap:6px}.filter-field span{font-size:12px;color:#42565d;font-weight:700}.filter-field input,.filter-field select{border:1px solid var(--line);border-radius:12px;padding:10px 12px;font:inherit;background:#fff;color:var(--ink)}"
      + ".filter-actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px}.filter-actions button{border:1px solid var(--line);border-radius:999px;padding:9px 14px;background:#fff;color:#314249;font-weight:800;cursor:pointer}.filter-actions .primary{background:#142126;color:#fff;border-color:#142126}"
      + ".filter-status{margin-top:10px;font-size:12px;color:#42565d;line-height:1.6}.alert{margin-top:10px;padding:10px 12px;border-radius:12px;border:1px solid #fecaca;background:#fff1f2;color:#991b1b;font-size:12px;display:none}"
      + ".nav{display:flex;gap:8px;flex-wrap:wrap;margin-top:18px}.nav button{border:1px solid var(--line);background:#fff;border-radius:999px;padding:9px 14px;font-weight:800;color:#314249;cursor:pointer}.nav button.active{background:#142126;color:#fff;border-color:#142126}"
      + ".page{display:none;margin-top:18px}.page.active{display:block}.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}"
      + ".kpi{background:var(--card);border:1px solid var(--line);border-left:6px solid #94a3b8;border-radius:18px;padding:14px 15px;box-shadow:0 10px 30px rgba(20,33,38,.05)}.kpi.blue{background:#ecfdf5;border-left-color:var(--blue)}.kpi.red{background:#fff5f5;border-left-color:var(--red)}.kpi.amber{background:#fffaf0;border-left-color:var(--amber)}.kpi.green{background:#f0fdf4;border-left-color:var(--green)}.kpi.slate{background:#f8fafc;border-left-color:var(--slate)}.kpi.sky{background:#eff8ff;border-left-color:#0284c7}"
      + ".kpi-title{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#52646b;font-weight:800}.kpi-value{font-size:28px;font-weight:800;line-height:1.15;margin-top:4px}.kpi-sub{font-size:12px;color:var(--muted);margin-top:6px}"
      + ".callout{margin-top:16px;padding:14px 16px;border:1px solid #c7d8d3;border-radius:18px;background:linear-gradient(135deg,#f4fbf8,#fffdf8)}.callout strong{display:block;margin-bottom:4px}"
      + ".grid-2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}.grid-3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px}.pad-top{margin-top:16px}"
      + ".card{background:var(--card);border:1px solid var(--line);border-radius:20px;padding:16px;box-shadow:0 10px 28px rgba(20,33,38,.04)}.card h2{margin:0 0 6px;font-size:18px;letter-spacing:-.01em}.muted{font-size:12px;color:var(--muted);line-height:1.5}.table-wrap{overflow:auto}"
      + "table{border-collapse:collapse;width:100%;margin-top:10px}th,td{border:1px solid #e4e8e3;padding:8px 9px;text-align:left;font-size:12px;vertical-align:top}th{background:#fafaf7;color:#314249}"
      + ".bars{display:flex;flex-direction:column;gap:10px;margin-top:10px}.bar-row{display:grid;grid-template-columns:220px 1fr 120px;gap:10px;align-items:center}.bar-label strong{display:block;font-size:12px;color:#22343b;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.bar-meta{display:block;font-size:11px;color:var(--muted);margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.bar-track{height:16px;border-radius:999px;background:#edf1ec;overflow:hidden}.bar-fill{height:100%;border-radius:999px}.bar-value{font-size:12px;text-align:right;color:#22343b;font-variant-numeric:tabular-nums}"
      + ".tag{display:inline-flex;align-items:center;border-radius:999px;padding:4px 9px;font-size:11px;font-weight:800;border:1px solid transparent}.tag.red{background:#fff1f2;color:#9f1239;border-color:#fecdd3}.tag.amber{background:#fff7ed;color:#9a3412;border-color:#fed7aa}.tag.green{background:#f0fdf4;color:#166534;border-color:#bbf7d0}.tag.blue{background:#eff6ff;color:#1d4ed8;border-color:#bfdbfe}.tag.slate{background:#f8fafc;color:#475569;border-color:#e2e8f0}"
      + ".empty{padding:16px 0;font-size:12px;color:var(--muted)}.note-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin-top:10px}.note-box{padding:12px;border:1px solid var(--line);border-radius:14px;background:#fafaf7}.note-box strong{display:block;font-size:12px;margin-bottom:4px}.chart-wrap{margin-top:12px;border:1px solid var(--line);border-radius:16px;padding:14px;background:#fff}.legend{display:flex;gap:12px;flex-wrap:wrap;font-size:12px;color:#314249;margin-bottom:10px}.legend span{display:inline-flex;align-items:center;gap:6px}.dot{width:10px;height:10px;border-radius:50%}.axis-label{font-size:11px;fill:#64748b}.footer-note{margin-top:18px;font-size:12px;color:var(--muted);line-height:1.6}"
      + "@media (max-width:1280px){.filter-grid{grid-template-columns:repeat(3,minmax(130px,1fr))}}@media (max-width:1120px){.head,.grid-2,.grid-3{grid-template-columns:1fr}.filter-grid{grid-template-columns:repeat(2,minmax(130px,1fr))}}@media (max-width:760px){.wrap{padding:18px 16px 32px}.title{font-size:24px}.filter-grid{grid-template-columns:1fr}.bar-row{grid-template-columns:1fr;gap:6px}.bar-value{text-align:left}}";
    document.head.appendChild(style);
  }

  function mountShell(){
    document.title = ('Painel devolucoes ' + rawText(meta.tag)).trim();
    document.body.innerHTML = ''
      + '<div class="wrap"><div class="head"><div><div class="title">Painel analitico de devolucoes</div><div id="hero-sub-1" class="sub"></div><div class="sub">Valores financeiros sao calculados apenas com as bases oficiais. Genius e Troque Facil entram somente para motivo, comentario, status e contexto operacional.</div><div class="pill-row"><span class="pill">Base fixa 01/2025 ate atual para taxas analiticas</span><span class="pill">Filtro global de data em todas as paginas</span><span class="pill">Comparativo mensal fixo 2025 vs 2026</span></div></div>'
      + '<div class="filter-panel"><div class="filter-title">Filtros globais</div><div class="filter-grid">'
      + '<label class="filter-field"><span>Data inicial</span><input id="f-date-start" type="date"></label><label class="filter-field"><span>Data final</span><input id="f-date-end" type="date"></label><label class="filter-field"><span>Ano</span><select id="f-year"></select></label><label class="filter-field"><span>Mes</span><select id="f-month"></select></label>'
      + '<label class="filter-field"><span>Colecao</span><select id="f-collection"></select></label><label class="filter-field"><span>Categoria (Grupo/Subgrupo PA)</span><select id="f-category"></select></label><label class="filter-field"><span>SALE</span><select id="f-sale"></select></label><label class="filter-field"><span>Loja</span><select id="f-store"></select></label>'
      + '<label class="filter-field"><span>Cliente</span><input id="f-client" type="text" placeholder="Nome, documento ou codigo"></label><label class="filter-field"><span>Produto / SKU</span><input id="f-product" type="text" placeholder="Descricao, referencia ou SKU"></label><label class="filter-field"><span>Tamanho</span><select id="f-size"></select></label><label class="filter-field"><span>Motivo</span><select id="f-reason"></select></label>'
      + '<label class="filter-field"><span>Status</span><select id="f-status"></select></label><label class="filter-field"><span>Plataforma</span><select id="f-platform"></select></label><label class="filter-field"><span>Comentario classificado</span><select id="f-comment-category"></select></label>'
      + '</div><div class="filter-actions"><button id="f-apply" class="primary" type="button">Aplicar filtros</button><button id="f-reset" type="button">Voltar ao padrao</button></div><div id="filter-status" class="filter-status"></div><div id="filter-alert" class="alert"></div></div></div>'
      + '<div class="nav"><button class="active" data-page="exec">Visao executiva</button><button data-page="origem">Origem do aumento</button><button data-page="produtos">Produtos</button><button data-page="sale">SALE</button><button data-page="colecoes">Colecoes</button><button data-page="categorias">Categorias</button><button data-page="lojas">Lojas</button><button data-page="clientes">Clientes</button><button data-page="tamanhos">Tamanhos</button><button data-page="motivos">Motivos e comentarios</button><button data-page="operacao">Operacao</button><button data-page="detalhes">Detalhamento</button></div>'
      + '<section id="page-exec" class="page active"><div id="exec-kpis" class="kpi-grid"></div><div id="exec-callout" class="callout"></div><div class="grid-2 pad-top"><div class="card"><h2>Comparativo mensal 2025 vs 2026</h2><div class="muted">A linha mostra a taxa de devolucao sobre vendas mes a mes. A tabela abaixo traz vendido, devolvido e percentual por ano.</div><div id="monthly-chart"></div><div id="monthly-table" class="table-wrap"></div></div><div class="card"><h2>Alertas e destaques</h2><div class="muted">Criticos so entram quando ha base minima e base fixa completa no periodo.</div><div id="exec-highlights" class="table-wrap"></div></div></div><div class="grid-2 pad-top"><div class="card"><h2>Produtos com maior impacto absoluto</h2><div class="muted">Ranking por valor devolvido na base fixa do periodo.</div><div id="exec-products-bars"></div></div><div class="card"><h2>Tamanhos e clientes sob risco</h2><div class="muted">Leitura rapida dos maiores percentuais e maiores bases de devolucao.</div><div id="exec-sizes-bars"></div><div id="exec-clients-bars"></div></div></div></section>'
      + '<section id="page-origem" class="page"><div class="grid-2"><div class="card"><h2>Origem do aumento por colecao</h2><div class="muted">Delta do periodo filtrado contra o mesmo recorte do ano anterior.</div><div id="origem-colecao" class="table-wrap"></div></div><div class="card"><h2>Origem do aumento por categoria</h2><div class="muted">Mostra onde o valor devolvido cresceu mais no recorte atual.</div><div id="origem-categoria" class="table-wrap"></div></div></div><div class="grid-2 pad-top"><div class="card"><h2>Origem do aumento por loja</h2><div id="origem-loja" class="table-wrap"></div></div><div class="card"><h2>Origem do aumento por SALE</h2><div id="origem-sale" class="table-wrap"></div></div></div></section>'
      + '<section id="page-produtos" class="page"><div class="grid-2"><div class="card"><h2>Produtos criticos</h2><div class="muted">Itens com taxa >= 80% e base minima relevante no periodo base.</div><div id="products-critical" class="table-wrap"></div></div><div class="card"><h2>Produtos com maior impacto absoluto</h2><div class="muted">Maior valor devolvido, independente da taxa.</div><div id="products-impact" class="table-wrap"></div></div></div><div class="grid-2 pad-top"><div class="card"><h2>Produtos para monitorar</h2><div class="muted">Alta taxa com base curta ou deterioracao forte no periodo atual.</div><div id="products-monitor" class="table-wrap"></div></div><div class="card"><h2>Produtos mais devolvidos por quantidade</h2><div id="products-qty" class="table-wrap"></div></div></div></section>'
      + '<section id="page-sale" class="page"><div class="grid-2"><div class="card"><h2>SALE vs NAO_SALE</h2><div class="muted">Comparativo do periodo base com valor vendido, devolvido, quantidade e taxa.</div><div id="sale-table" class="table-wrap"></div></div><div class="card"><h2>Ranking visual</h2><div class="muted">A barra principal usa valor devolvido no periodo; o subtitulo mostra a taxa.</div><div id="sale-bars"></div></div></div></section>'
      + '<section id="page-colecoes" class="page"><div class="grid-2"><div class="card"><h2>Ranking de colecoes</h2><div class="muted">Valor devolvido, taxa do periodo e delta do periodo atual.</div><div id="colecoes-bars"></div></div><div class="card"><h2>Tabela por colecao</h2><div id="colecoes-table" class="table-wrap"></div></div></div></section>'
      + '<section id="page-categorias" class="page"><div class="grid-2"><div class="card"><h2>Ranking de categorias</h2><div class="muted">Categoria calculada por Grupo PA + Subgrupo PA na base do periodo.</div><div id="categorias-bars"></div></div><div class="card"><h2>Tabela por categoria</h2><div id="categorias-table" class="table-wrap"></div></div></div></section>'
      + '<section id="page-lojas" class="page"><div class="grid-2"><div class="card"><h2>Ranking de lojas</h2><div class="muted">Comparativo entre lojas/origens com base financeira oficial.</div><div id="lojas-bars"></div></div><div class="card"><h2>Tabela por loja</h2><div id="lojas-table" class="table-wrap"></div></div></div></section>'
      + '<section id="page-clientes" class="page"><div class="grid-2"><div class="card"><h2>Clientes criticos</h2><div class="muted">Clientes com mais de 80% de devolucao por valor ou quantidade, com base minima.</div><div id="clients-critical" class="table-wrap"></div></div><div class="card"><h2>Top clientes por valor devolvido</h2><div id="clients-value" class="table-wrap"></div></div></div><div class="grid-2 pad-top"><div class="card"><h2>Top clientes por taxa do periodo</h2><div id="clients-rate" class="table-wrap"></div></div><div class="card"><h2>Top clientes por quantidade devolvida</h2><div id="clients-qty" class="table-wrap"></div></div></div></section>'
      + '<section id="page-tamanhos" class="page"><div class="grid-2"><div class="card"><h2>Tamanhos com maior taxa de devolucao</h2><div class="muted">Leitura comparando quantidade e valor devolvido contra vendido.</div><div id="sizes-table" class="table-wrap"></div></div><div class="card"><h2>Resumo critico de tamanhos</h2><div id="sizes-bars"></div><div id="sizes-note" class="callout"></div></div></div></section>'
      + '<section id="page-motivos" class="page"><div class="grid-2"><div class="card"><h2>Motivos operacionais</h2><div class="muted">Contagens puras de processos. Nenhum valor financeiro e somado aqui.</div><div id="motivos-bars"></div><div id="motivos-table" class="table-wrap"></div></div><div class="card"><h2>Comentarios classificados</h2><div class="muted">Inclui a nova classe RECEBIMENTO_ERRADO.</div><div id="themes-bars"></div><div id="themes-table" class="table-wrap"></div></div></div><div class="card pad-top"><h2>Recebimento errado</h2><div class="muted">Termos como produto errado, item divergente, pedido trocado e tamanho/cor errados foram consolidados na categoria RECEBIMENTO_ERRADO.</div><div id="wrong-kpis" class="kpi-grid"></div></div><div class="grid-2 pad-top"><div class="card"><h2>Recebimento errado por produto</h2><div id="wrong-product" class="table-wrap"></div></div><div class="card"><h2>Recebimento errado por loja</h2><div id="wrong-store" class="table-wrap"></div></div></div><div class="grid-2 pad-top"><div class="card"><h2>Recebimento errado por categoria</h2><div id="wrong-category" class="table-wrap"></div></div><div class="card"><h2>Recebimento errado por colecao</h2><div id="wrong-collection" class="table-wrap"></div></div></div><div class="card pad-top"><h2>Comentarios recentes</h2><div id="comments-table" class="table-wrap"></div></div></section>'
      + '<section id="page-operacao" class="page"><div class="grid-3"><div class="card"><h2>Status das solicitacoes</h2><div id="status-bars"></div></div><div class="card"><h2>Estagios</h2><div id="stage-table" class="table-wrap"></div></div><div class="card"><h2>Plataformas</h2><div id="platform-table" class="table-wrap"></div></div></div><div class="grid-2 pad-top"><div class="card"><h2>Qualidade e cobertura</h2><div class="muted">Mostra historico disponivel, campos mapeados e quando a ponte operacional esta sendo usada para filtrar o financeiro.</div><div id="coverage-table" class="table-wrap"></div></div><div class="card"><h2>Artefatos exportados</h2><div id="exports-table" class="table-wrap"></div></div></div></section>'
      + '<section id="page-detalhes" class="page"><div class="grid-2"><div class="card"><h2>Detalhamento financeiro</h2><div class="muted">Valores e quantidades continuam 100% vindos da base financeira oficial.</div><div id="detail-finance" class="table-wrap"></div></div><div class="card"><h2>Detalhamento operacional</h2><div class="muted">Motivo, comentario, status e plataforma para investigacao operacional.</div><div id="detail-operational" class="table-wrap"></div></div></div></section>'
      + '<div id="footer-note" class="footer-note"></div></div>';
  }

  ensureStyle();
  mountShell();

  const controls = {
    dateStart:document.getElementById('f-date-start'), dateEnd:document.getElementById('f-date-end'),
    year:document.getElementById('f-year'), month:document.getElementById('f-month'),
    collection:document.getElementById('f-collection'), category:document.getElementById('f-category'),
    sale:document.getElementById('f-sale'), store:document.getElementById('f-store'),
    client:document.getElementById('f-client'), product:document.getElementById('f-product'),
    size:document.getElementById('f-size'), reason:document.getElementById('f-reason'),
    status:document.getElementById('f-status'), platform:document.getElementById('f-platform'),
    commentCategory:document.getElementById('f-comment-category'),
    apply:document.getElementById('f-apply'), reset:document.getElementById('f-reset'),
    statusText:document.getElementById('filter-status'), alertBox:document.getElementById('filter-alert')
  };

  const filterOptions = buildFilterOptions();
  const defaultState = {
    dateStart:defaultStart, dateEnd:defaultEnd, year:'TODOS', month:'TODOS', collection:'TODOS',
    category:'TODOS', sale:'TODOS', store:'TODOS', client:'', product:'', size:'TODOS',
    reason:'TODOS', status:'TODOS', platform:'TODOS', commentCategory:'TODOS'
  };
  let state = Object.assign({}, defaultState);

  function populateControls(){
    controls.dateStart.min = filterMin; controls.dateStart.max = filterMax; controls.dateEnd.min = filterMin; controls.dateEnd.max = filterMax;
    setSelectOptions(controls.year, filterOptions.years, true);
    setSelectOptions(controls.month, filterOptions.months, true);
    setSelectOptions(controls.collection, filterOptions.collections, true);
    setSelectOptions(controls.category, filterOptions.categories, true);
    setSelectOptions(controls.store, filterOptions.stores, true);
    setSelectOptions(controls.size, filterOptions.sizes, true);
    setSelectOptions(controls.reason, filterOptions.motives, true);
    setSelectOptions(controls.status, filterOptions.statuses, true);
    setSelectOptions(controls.platform, filterOptions.platforms, true);
    setSelectOptions(controls.commentCategory, filterOptions.commentCategories, true);
    setSelectOptions(controls.sale, [{ value:'SALE', label:'SALE' }, { value:'NAO_SALE', label:'NAO_SALE' }], true);
    resetControls();
  }
  function resetControls(){
    controls.dateStart.value = defaultState.dateStart; controls.dateEnd.value = defaultState.dateEnd; controls.year.value = defaultState.year; controls.month.value = defaultState.month;
    controls.collection.value = defaultState.collection; controls.category.value = defaultState.category; controls.sale.value = defaultState.sale; controls.store.value = defaultState.store;
    controls.client.value = defaultState.client; controls.product.value = defaultState.product; controls.size.value = defaultState.size; controls.reason.value = defaultState.reason;
    controls.status.value = defaultState.status; controls.platform.value = defaultState.platform; controls.commentCategory.value = defaultState.commentCategory;
  }
  function readStateFromControls(){
    return {
      dateStart:rawText(controls.dateStart.value) || defaultStart, dateEnd:rawText(controls.dateEnd.value) || defaultEnd,
      year:controls.year.value || 'TODOS', month:controls.month.value || 'TODOS', collection:controls.collection.value || 'TODOS',
      category:controls.category.value || 'TODOS', sale:controls.sale.value || 'TODOS', store:controls.store.value || 'TODOS',
      client:rawText(controls.client.value), product:rawText(controls.product.value), size:controls.size.value || 'TODOS',
      reason:controls.reason.value || 'TODOS', status:controls.status.value || 'TODOS', platform:controls.platform.value || 'TODOS', commentCategory:controls.commentCategory.value || 'TODOS'
    };
  }
  function validateState(nextState){
    const out = Object.assign({}, nextState);
    if(out.dateStart < filterMin){ out.dateStart = filterMin; }
    if(out.dateStart > filterMax){ out.dateStart = filterMax; }
    if(out.dateEnd < filterMin){ out.dateEnd = filterMin; }
    if(out.dateEnd > filterMax){ out.dateEnd = filterMax; }
    if(out.dateStart > out.dateEnd){ return { ok:false, message:'A data inicial nao pode ser maior que a data final.' }; }
    return { ok:true, value:out };
  }
  function hasOperationalFilters(currentState){ return currentState.reason !== 'TODOS' || currentState.status !== 'TODOS' || currentState.platform !== 'TODOS' || currentState.commentCategory !== 'TODOS'; }
  function matchesYearMonth(iso, currentState, shiftYears){
    const year = rawText(iso).slice(0, 4);
    const month = rawText(iso).slice(5, 7);
    if(currentState.year !== 'TODOS' && year !== String(Number(currentState.year) + safeNum(shiftYears))){ return false; }
    if(currentState.month !== 'TODOS' && month !== currentState.month){ return false; }
    return true;
  }
  function matchesCommonFilters(row, currentState, scope){
    if(currentState.collection !== 'TODOS' && row.Colecao !== currentState.collection){ return false; }
    if(currentState.category !== 'TODOS' && row.CategoriaPainel !== currentState.category){ return false; }
    if(currentState.sale !== 'TODOS' && row.SaleBucket !== currentState.sale){ return false; }
    if(currentState.store !== 'TODOS' && row.Loja !== currentState.store){ return false; }
    if(currentState.size !== 'TODOS' && row.Tamanho !== currentState.size){ return false; }
    if(currentState.client){
      const needle = currentState.client.toLowerCase();
      const hay = scope === 'finance' ? (row.ClienteId + ' ' + row.ClienteLabel).toLowerCase() : (row.Cliente + ' ' + row.Documento).toLowerCase();
      if(hay.indexOf(needle) === -1){ return false; }
    }
    if(currentState.product){
      const hayProduct = (row.Sku + ' ' + row.Produto).toLowerCase();
      if(hayProduct.indexOf(currentState.product.toLowerCase()) === -1){ return false; }
    }
    return true;
  }
  function matchesOperationalFilters(row, currentState){
    if(currentState.reason !== 'TODOS' && row.Motivo !== currentState.reason){ return false; }
    if(currentState.status !== 'TODOS' && row.StatusSolicitacao !== currentState.status){ return false; }
    if(currentState.platform !== 'TODOS' && row.Plataforma !== currentState.platform){ return false; }
    if(currentState.commentCategory !== 'TODOS' && row.CategoriaComentario !== currentState.commentCategory){ return false; }
    return true;
  }
  function buildBridge(start, end, currentState, yearShift, useYearMonth){
    if(!hasOperationalFilters(currentState)){ return null; }
    const bridge = { pedidoSku:new Set(), pedido:new Set(), sku:new Set(), count:0 };
    operationalData.forEach(function(row){
      if(!inRange(row.Data, start, end)){ return; }
      if(useYearMonth && !matchesYearMonth(row.Data, currentState, yearShift)){ return; }
      if(!matchesCommonFilters(row, currentState, 'operational')){ return; }
      if(!matchesOperationalFilters(row, currentState)){ return; }
      bridge.count += 1;
      if(row.PedidoSku){ bridge.pedidoSku.add(row.PedidoSku); }
      if(row.PedidoKey){ bridge.pedido.add(row.PedidoKey); }
      if(row.SkuKey){ bridge.sku.add(row.SkuKey); }
    });
    return bridge;
  }
  function financeMatchesBridge(row, bridge){
    if(!bridge){ return true; }
    if(row.PedidoSku && bridge.pedidoSku.has(row.PedidoSku)){ return true; }
    if(row.PedidoKey && bridge.pedido.has(row.PedidoKey)){ return true; }
    return !row.PedidoKey && row.SkuKey && bridge.sku.has(row.SkuKey);
  }
  function buildMonthlyComparison(currentState){
    const selected = splitIso(currentState.dateEnd) || splitIso(dataMax);
    const available = splitIso(dataMax);
    if(!selected || !available){ return { rows:[], maxRate:0 }; }
    const finalMonth = Math.min(selected.month, available.month);
    const finalDay = finalMonth === selected.month ? selected.day : available.day;
    const currentBridge = buildBridge(currentYear + '-01-01', dataMax, currentState, 0, false);
    const priorBridge = buildBridge(priorYear + '-01-01', shiftYearIso(dataMax, -1), currentState, 0, false);
    function aggregateForYear(year, bridge){
      const buckets = {};
      for(let index = 1; index <= finalMonth; index += 1){ buckets[String(index).padStart(2, '0')] = { ValorVendido:0, ValorDevolvido:0, QuantidadeVendida:0, QuantidadeDevolvida:0 }; }
      financeData.forEach(function(row){
        if(row.Year !== String(year)){ return; }
        if(!matchesCommonFilters(row, currentState, 'finance')){ return; }
        if(!financeMatchesBridge(row, bridge)){ return; }
        const month = Number(row.Month);
        const day = Number(row.Day);
        if(month < 1 || month > finalMonth || (month === finalMonth && day > finalDay)){ return; }
        addFinanceMetrics(buckets[row.Month], row);
      });
      return buckets;
    }
    const currentBuckets = aggregateForYear(currentYear, currentBridge);
    const priorBuckets = aggregateForYear(priorYear, priorBridge);
    const rows = [];
    let maxRate = 0;
    for(let index = 1; index <= finalMonth; index += 1){
      const monthKey = String(index).padStart(2, '0');
      const currentBucket = currentBuckets[monthKey] || { ValorVendido:0, ValorDevolvido:0, QuantidadeVendida:0, QuantidadeDevolvida:0 };
      const priorBucket = priorBuckets[monthKey] || { ValorVendido:0, ValorDevolvido:0, QuantidadeVendida:0, QuantidadeDevolvida:0 };
      const currentRate = rateOrNull(currentBucket.ValorDevolvido, currentBucket.ValorVendido);
      const priorRate = rateOrNull(priorBucket.ValorDevolvido, priorBucket.ValorVendido);
      maxRate = Math.max(maxRate, safeNum(currentRate), safeNum(priorRate));
      rows.push({
        Mes:MONTH_LABEL[monthKey], MesNumero:index,
        PriorValorVendido:round(priorBucket.ValorVendido, 2), PriorValorDevolvido:round(priorBucket.ValorDevolvido, 2), PriorTaxaValorPct:priorRate,
        CurrentValorVendido:round(currentBucket.ValorVendido, 2), CurrentValorDevolvido:round(currentBucket.ValorDevolvido, 2), CurrentTaxaValorPct:currentRate,
        DeltaTaxaValorPct:deltaMaybe(currentRate, priorRate)
      });
    }
    return { rows:rows, maxRate:maxRate };
  }
  function buildContext(currentState){
    const periodStart = currentState.dateStart;
    const periodEnd = currentState.dateEnd;
    const priorStart = shiftYearIso(periodStart, -1);
    const priorEnd = shiftYearIso(periodEnd, -1);
    const rollingStart = fixedPeriodStart;
    let rollingEnd = periodEnd;
    if(rollingEnd > fixedPeriodEnd){ rollingEnd = fixedPeriodEnd; }
    if(rollingEnd < rollingStart){ rollingEnd = rollingStart; }
    const rollingFull = fixedPeriodHistoryComplete && rollingEnd >= rollingStart;
    const periodBridge = buildBridge(periodStart, periodEnd, currentState, 0, true);
    const priorBridge = buildBridge(priorStart, priorEnd, currentState, -1, true);
    const rollingBridge = buildBridge(rollingStart, rollingEnd, currentState, 0, false);
    const periodFinance = financeData.filter(function(row){ return inRange(row.Data, periodStart, periodEnd) && matchesYearMonth(row.Data, currentState, 0) && matchesCommonFilters(row, currentState, 'finance') && financeMatchesBridge(row, periodBridge); });
    const priorFinance = financeData.filter(function(row){ return inRange(row.Data, priorStart, priorEnd) && matchesYearMonth(row.Data, currentState, -1) && matchesCommonFilters(row, currentState, 'finance') && financeMatchesBridge(row, priorBridge); });
    const rollingFinance = financeData.filter(function(row){ return inRange(row.Data, rollingStart, rollingEnd) && matchesCommonFilters(row, currentState, 'finance') && financeMatchesBridge(row, rollingBridge); });
    const periodOperational = operationalData.filter(function(row){ return inRange(row.Data, periodStart, periodEnd) && matchesYearMonth(row.Data, currentState, 0) && matchesCommonFilters(row, currentState, 'operational') && matchesOperationalFilters(row, currentState); });
    return {
      periodStart:periodStart, periodEnd:periodEnd, priorStart:priorStart, priorEnd:priorEnd, rollingStart:rollingStart, rollingEnd:rollingEnd, rollingFull:rollingFull,
      periodBridgeCount:periodBridge ? periodBridge.count : 0, priorBridgeCount:priorBridge ? priorBridge.count : 0, rollingBridgeCount:rollingBridge ? rollingBridge.count : 0,
      operationalFiltersApplied:hasOperationalFilters(currentState),
      periodFinance:periodFinance, priorFinance:priorFinance, rollingFinance:rollingFinance, periodOperational:periodOperational,
      periodSummary:buildFinanceSummary(periodFinance), priorSummary:buildFinanceSummary(priorFinance), rollingSummary:buildFinanceSummary(rollingFinance), monthly:buildMonthlyComparison(currentState)
    };
  }
  function buildViews(ctx){
    const clients = buildMergedDimension(ctx.rollingFinance, ctx.periodFinance, ctx.priorFinance, function(row){ return row.ClienteId; }, function(row, key){ return { Key:key, ClienteId:key, Cliente:row.ClienteLabel, QuantidadeVendida:0, ValorVendido:0, QuantidadeDevolvida:0, ValorDevolvido:0 }; }, ctx.rollingFull, 'cliente');
    const products = buildMergedDimension(ctx.rollingFinance, ctx.periodFinance, ctx.priorFinance, function(row){ return row.Sku; }, function(row, key){ return { Key:key, Sku:key, Produto:row.Produto, Colecao:row.Colecao, Categoria:row.CategoriaPainel, Tamanho:row.Tamanho, Loja:row.Loja, SaleBucket:row.SaleBucket, QuantidadeVendida:0, ValorVendido:0, QuantidadeDevolvida:0, ValorDevolvido:0 }; }, ctx.rollingFull, 'produto');
    const sizes = buildMergedDimension(ctx.rollingFinance, ctx.periodFinance, ctx.priorFinance, function(row){ return row.Tamanho; }, function(row, key){ return { Key:key, Label:key, Tamanho:key, QuantidadeVendida:0, ValorVendido:0, QuantidadeDevolvida:0, ValorDevolvido:0 }; }, ctx.rollingFull, 'tamanho');
    const collections = buildMergedDimension(ctx.rollingFinance, ctx.periodFinance, ctx.priorFinance, function(row){ return row.Colecao; }, function(row, key){ return { Key:key, Label:key, QuantidadeVendida:0, ValorVendido:0, QuantidadeDevolvida:0, ValorDevolvido:0 }; }, ctx.rollingFull, 'colecao');
    const categories = buildMergedDimension(ctx.rollingFinance, ctx.periodFinance, ctx.priorFinance, function(row){ return row.CategoriaPainel; }, function(row, key){ return { Key:key, Label:key, QuantidadeVendida:0, ValorVendido:0, QuantidadeDevolvida:0, ValorDevolvido:0 }; }, ctx.rollingFull, 'categoria');
    const stores = buildMergedDimension(ctx.rollingFinance, ctx.periodFinance, ctx.priorFinance, function(row){ return row.Loja; }, function(row, key){ return { Key:key, Label:key, QuantidadeVendida:0, ValorVendido:0, QuantidadeDevolvida:0, ValorDevolvido:0 }; }, ctx.rollingFull, 'loja');
    const sale = buildMergedDimension(ctx.rollingFinance, ctx.periodFinance, ctx.priorFinance, function(row){ return row.SaleBucket; }, function(row, key){ return { Key:key, Label:key, QuantidadeVendida:0, ValorVendido:0, QuantidadeDevolvida:0, ValorDevolvido:0 }; }, ctx.rollingFull, 'sale');
    const motives = buildOperationalCounts(ctx.periodOperational, function(row){ return row.Motivo; }, function(row, key){ return { Key:key, Label:key, Count:0, QuantidadeSolicitada:0 }; }).sort(sortDesc('Count', 'QuantidadeSolicitada'));
    const themes = buildOperationalCounts(ctx.periodOperational, function(row){ return row.CategoriaComentario; }, function(row, key){ return { Key:key, Label:key, Count:0, QuantidadeSolicitada:0 }; }).sort(sortDesc('Count', 'QuantidadeSolicitada'));
    const status = buildOperationalCounts(ctx.periodOperational, function(row){ return row.StatusSolicitacao; }, function(row, key){ return { Key:key, Label:key, Count:0, QuantidadeSolicitada:0 }; }).sort(sortDesc('Count', 'QuantidadeSolicitada'));
    const stage = buildOperationalCounts(ctx.periodOperational, function(row){ return row.Estagio; }, function(row, key){ return { Key:key, Label:key, Count:0, QuantidadeSolicitada:0 }; }).sort(sortDesc('Count', 'QuantidadeSolicitada'));
    const platforms = buildOperationalCounts(ctx.periodOperational, function(row){ return row.Plataforma; }, function(row, key){ return { Key:key, Label:key, Count:0, QuantidadeSolicitada:0 }; }).sort(sortDesc('Count', 'QuantidadeSolicitada'));
    const wrongRows = ctx.periodOperational.filter(function(row){ return row.CategoriaComentario === 'RECEBIMENTO_ERRADO'; });
    const wrongByProduct = buildOperationalCounts(wrongRows, function(row){ return row.Sku; }, function(row, key){ return { Key:key, Sku:key, Produto:row.Produto, Count:0, QuantidadeSolicitada:0 }; }).sort(sortDesc('Count', 'QuantidadeSolicitada'));
    const wrongByStore = buildOperationalCounts(wrongRows, function(row){ return row.Loja; }, function(row, key){ return { Key:key, Label:key, Count:0, QuantidadeSolicitada:0 }; }).sort(sortDesc('Count', 'QuantidadeSolicitada'));
    const wrongByCategory = buildOperationalCounts(wrongRows, function(row){ return row.CategoriaPainel; }, function(row, key){ return { Key:key, Label:key, Count:0, QuantidadeSolicitada:0 }; }).sort(sortDesc('Count', 'QuantidadeSolicitada'));
    const wrongByCollection = buildOperationalCounts(wrongRows, function(row){ return row.Colecao; }, function(row, key){ return { Key:key, Label:key, Count:0, QuantidadeSolicitada:0 }; }).sort(sortDesc('Count', 'QuantidadeSolicitada'));
    const clientBaseRows = clients.filter(hasSalesBase);
    const productBaseRows = products.filter(hasSalesBase);
    const criticalClients = clientBaseRows.filter(function(row){ return ctx.rollingFull && row.BaseMinima && row.ClassificacaoCriticidade === 'CRITICO'; }).sort(sortByCriticality);
    const criticalProducts = productBaseRows.filter(function(row){ return ctx.rollingFull && row.BaseMinima && row.ClassificacaoCriticidade === 'CRITICO'; }).sort(sortByCriticality);
    const monitorProducts = productBaseRows.filter(function(row){
      if(!ctx.rollingFull){ return false; }
      return (!row.BaseMinima && row.MaxTaxaPct >= 80) || (row.BaseMinima && row.ClassificacaoCriticidade === 'ATENCAO');
    }).sort(sortByCriticality);
    return { clients:clients, products:products, sizes:sizes, collections:collections, categories:categories, stores:stores, sale:sale, motives:motives, themes:themes, status:status, stage:stage, platforms:platforms, wrongRows:wrongRows, wrongByProduct:wrongByProduct, wrongByStore:wrongByStore, wrongByCategory:wrongByCategory, wrongByCollection:wrongByCollection, clientBaseRows:clientBaseRows, productBaseRows:productBaseRows, criticalClients:criticalClients, criticalProducts:criticalProducts, monitorProducts:monitorProducts };
  }
  function renderFilterStatus(currentState, ctx){
    const hero = document.getElementById('hero-sub-1');
    const fixedCoverage = ctx.rollingFull ? ('Base fixa das taxas: ' + formatDate(ctx.rollingStart) + ' a ' + formatDate(ctx.rollingEnd)) : ('Base fixa incompleta. Historico disponivel desde ' + formatDate(dataMin) + ', abaixo do inicio requerido em ' + formatDate(ctx.rollingStart) + '.');
    hero.textContent = 'Periodo visivel: ' + formatDate(ctx.periodStart) + ' a ' + formatDate(ctx.periodEnd) + '. ' + fixedCoverage;
    controls.statusText.innerHTML = '<strong>Periodo visivel:</strong> ' + escapeHtml(formatDate(ctx.periodStart) + ' a ' + formatDate(ctx.periodEnd)) + ' | <strong>Comparativo:</strong> ' + escapeHtml(formatDate(ctx.priorStart) + ' a ' + formatDate(ctx.priorEnd)) + ' | <strong>Base das taxas:</strong> ' + escapeHtml(formatDate(ctx.rollingStart) + ' a ' + formatDate(ctx.rollingEnd)) + (ctx.operationalFiltersApplied ? ' | <strong>Ponte operacional ativa:</strong> ' + escapeHtml(String(ctx.rollingBridgeCount) + ' ocorrencias cruzando financeiro por Pedido+SKU/Pedido.') : '');
    if(!ctx.rollingFull){
      controls.alertBox.style.display = 'block';
      controls.alertBox.textContent = 'A base fixa solicitada exige historico desde ' + formatDate(ctx.rollingStart) + ', mas a base carregada comeca em ' + formatDate(dataMin) + '. As taxas e criticidades ficam sinalizadas como base insuficiente.';
    } else {
      controls.alertBox.style.display = 'none';
      controls.alertBox.textContent = '';
    }
  }
  function monthlyChartHtml(monthly){
    if(!monthly.rows.length){ return emptyHtml('Sem dados para montar o comparativo mensal.'); }
    const width = 620, height = 220, padding = { top:20, right:18, bottom:34, left:34 };
    const innerWidth = width - padding.left - padding.right;
    const innerHeight = height - padding.top - padding.bottom;
    const maxRate = Math.max(5, Math.ceil(monthly.maxRate / 5) * 5);
    const xStep = monthly.rows.length > 1 ? innerWidth / (monthly.rows.length - 1) : innerWidth;
    function xPoint(index){ return padding.left + (monthly.rows.length === 1 ? innerWidth / 2 : index * xStep); }
    function yPoint(rate){ return padding.top + innerHeight - ((safeNum(rate) / maxRate) * innerHeight); }
    function linePoints(field){ return monthly.rows.map(function(row, index){ return xPoint(index) + ',' + yPoint(row[field]); }).join(' '); }
    const grid = [0, maxRate / 2, maxRate].map(function(value){ const y = yPoint(value); return '<line x1="' + padding.left + '" x2="' + (width - padding.right) + '" y1="' + y + '" y2="' + y + '" stroke="#d9e2dc" stroke-dasharray="3 4"></line><text x="8" y="' + (y + 4) + '" class="axis-label">' + escapeHtml(formatPercent(value)) + '</text>'; }).join('');
    const currentDots = monthly.rows.map(function(row, index){ return '<circle cx="' + xPoint(index) + '" cy="' + yPoint(row.CurrentTaxaValorPct) + '" r="4" fill="#0f766e"></circle>'; }).join('');
    const priorDots = monthly.rows.map(function(row, index){ return '<circle cx="' + xPoint(index) + '" cy="' + yPoint(row.PriorTaxaValorPct) + '" r="4" fill="#b42318"></circle>'; }).join('');
    const labels = monthly.rows.map(function(row, index){ return '<text x="' + xPoint(index) + '" y="' + (height - 10) + '" text-anchor="middle" class="axis-label">' + escapeHtml(row.Mes) + '</text>'; }).join('');
    return '<div class="chart-wrap"><div class="legend"><span><i class="dot" style="background:#0f766e"></i>' + escapeHtml(String(currentYear)) + ' % devolucao</span><span><i class="dot" style="background:#b42318"></i>' + escapeHtml(String(priorYear)) + ' % devolucao</span></div><svg viewBox="0 0 ' + width + ' ' + height + '" width="100%" height="220" role="img" aria-label="Comparativo mensal de taxa de devolucao">' + grid + '<polyline fill="none" stroke="#b42318" stroke-width="3" points="' + linePoints('PriorTaxaValorPct') + '"></polyline><polyline fill="none" stroke="#0f766e" stroke-width="3" points="' + linePoints('CurrentTaxaValorPct') + '"></polyline>' + priorDots + currentDots + labels + '</svg></div>';
  }
  function renderExecutive(ctx, views){
    const deltaReturnValue = round(ctx.periodSummary.ValorDevolvido - ctx.priorSummary.ValorDevolvido, 2);
    const totalClients = views.clientBaseRows.length;
    const totalProducts = views.productBaseRows.length;
    const topCollection = [].concat(views.collections).sort(sortDesc('DeltaPeriodValorDevolvido', 'ValorDevolvido'))[0];
    const topStore = [].concat(views.stores).sort(sortDesc('DeltaPeriodValorDevolvido', 'ValorDevolvido'))[0];
    const topProduct = [].concat(views.products).sort(sortDesc('ValorDevolvido', 'QuantidadeDevolvida'))[0];
    const topCriticalClient = views.criticalClients[0];
    const topCriticalProduct = views.criticalProducts[0];
    const topCategoryDelta = [].concat(views.categories).sort(sortDesc('DeltaPeriodValorDevolvido', 'ValorDevolvido'))[0];
    const execKpis = document.getElementById('exec-kpis');
    execKpis.innerHTML = ''
      + kpiHtml('Valor vendido no periodo', formatMoney(ctx.periodSummary.ValorVendido), 'Comparativo oficial do filtro global.', 'green')
      + kpiHtml('Valor devolvido no periodo', formatMoney(ctx.periodSummary.ValorDevolvido), 'Delta vs periodo comparavel: ' + formatMoney(deltaReturnValue), 'red')
      + kpiHtml('Taxa periodo (' + fixedPeriodLabel + ') - valor', formatPercent(ctx.rollingSummary.TaxaValorPct), fixedPeriodSubtitle, 'amber')
      + kpiHtml('Taxa periodo (' + fixedPeriodLabel + ') - quantidade', formatPercent(ctx.rollingSummary.TaxaQuantidadePct), fixedPeriodSubtitle, 'amber')
      + kpiHtml('Base periodo vendida', formatMoney(ctx.rollingSummary.ValorVendido), fixedPeriodSubtitle, 'blue')
      + kpiHtml('Base periodo devolvida', formatMoney(ctx.rollingSummary.ValorDevolvido), fixedPeriodSubtitle, 'blue')
      + kpiHtml('Clientes totais', intFmt.format(totalClients), fixedPeriodSubtitle, totalClients ? 'green' : 'slate')
      + kpiHtml('Clientes criticos', intFmt.format(views.criticalClients.length), ctx.rollingFull ? criticalShareText(views.criticalClients.length, totalClients) : 'Criticos desativados por base insuficiente.', views.criticalClients.length ? 'red' : 'slate')
      + kpiHtml('Produtos totais', intFmt.format(totalProducts), fixedPeriodSubtitle, totalProducts ? 'green' : 'slate')
      + kpiHtml('Produtos criticos', intFmt.format(views.criticalProducts.length), ctx.rollingFull ? criticalShareText(views.criticalProducts.length, totalProducts) : 'Criticos desativados por base insuficiente.', views.criticalProducts.length ? 'red' : 'slate')
      + kpiHtml('Recebimento errado', intFmt.format(views.wrongRows.length), 'Casos classificados no periodo visivel.', views.wrongRows.length ? 'sky' : 'slate');
    const callout = document.getElementById('exec-callout');
    callout.innerHTML = '<strong>Leitura executiva</strong>'
      + '<div class="muted">A taxa geral do periodo esta em ' + escapeHtml(formatPercent(ctx.rollingSummary.TaxaValorPct)) + ' por valor e ' + escapeHtml(formatPercent(ctx.rollingSummary.TaxaQuantidadePct)) + ' por quantidade, com base fixa de ' + escapeHtml(formatMoney(ctx.rollingSummary.ValorVendido)) + ' vendida e ' + escapeHtml(formatMoney(ctx.rollingSummary.ValorDevolvido)) + ' devolvida. Dentro dessa base, ha ' + escapeHtml(intFmt.format(totalClients)) + ' clientes, ' + escapeHtml(intFmt.format(views.criticalClients.length)) + ' clientes criticos, ' + escapeHtml(intFmt.format(totalProducts)) + ' produtos e ' + escapeHtml(intFmt.format(views.criticalProducts.length)) + ' produtos criticos.</div>'
      + '<div class="note-grid"><div class="note-box"><strong>Maior pressao no periodo visivel</strong><span class="muted">Colecao: ' + escapeHtml(topCollection ? topCollection.Label : 'Sem dados') + ' | Loja: ' + escapeHtml(topStore ? topStore.Label : 'Sem dados') + '</span></div><div class="note-box"><strong>Produto com maior impacto absoluto</strong><span class="muted">' + escapeHtml(topProduct ? (topProduct.Sku + ' - ' + clipText(topProduct.Produto, 42)) : 'Sem dados') + '</span></div></div>';
    document.getElementById('monthly-chart').innerHTML = monthlyChartHtml(ctx.monthly);
    renderTable('monthly-table', [
      { label:'Mes', key:'Mes' }, { label:String(priorYear) + ' vendido', render:function(row){ return formatMoney(row.PriorValorVendido); } },
      { label:String(priorYear) + ' devolvido', render:function(row){ return formatMoney(row.PriorValorDevolvido); } }, { label:String(priorYear) + ' %', render:function(row){ return formatPercent(row.PriorTaxaValorPct); } },
      { label:String(currentYear) + ' vendido', render:function(row){ return formatMoney(row.CurrentValorVendido); } }, { label:String(currentYear) + ' devolvido', render:function(row){ return formatMoney(row.CurrentValorDevolvido); } },
      { label:String(currentYear) + ' %', render:function(row){ return formatPercent(row.CurrentTaxaValorPct); } }, { label:'Delta p.p.', render:function(row){ return formatPercent(row.DeltaTaxaValorPct); } }
    ], ctx.monthly.rows, 'Sem dados para o comparativo mensal.');
    renderTable('exec-highlights', [
      { label:'Destaque', key:'Item' }, { label:'Valor', key:'Valor' }, { label:'Leitura', key:'Detalhe' }
    ], [
      { Item:'Clientes totais', Valor:intFmt.format(totalClients), Detalhe:fixedPeriodSubtitle },
      { Item:'Clientes criticos', Valor:intFmt.format(views.criticalClients.length), Detalhe:topCriticalClient ? (criticalShareText(views.criticalClients.length, totalClients) + ' Lider: ' + topCriticalClient.Cliente + ' | ' + formatPercent(topCriticalClient.MaxTaxaPct)) : 'Nenhum cliente critico no filtro atual.' },
      { Item:'Produtos totais', Valor:intFmt.format(totalProducts), Detalhe:fixedPeriodSubtitle },
      { Item:'Produtos criticos', Valor:intFmt.format(views.criticalProducts.length), Detalhe:topCriticalProduct ? (criticalShareText(views.criticalProducts.length, totalProducts) + ' Lider: ' + topCriticalProduct.Sku + ' | ' + formatPercent(topCriticalProduct.MaxTaxaPct)) : 'Nenhum produto critico no filtro atual.' },
      { Item:'Taxa geral do periodo', Valor:formatPercent(ctx.rollingSummary.TaxaValorPct), Detalhe:'Por quantidade: ' + formatPercent(ctx.rollingSummary.TaxaQuantidadePct) },
      { Item:'Maior delta por categoria', Valor:topCategoryDelta ? formatMoney(topCategoryDelta.DeltaPeriodValorDevolvido) : formatMoney(0), Detalhe:topCategoryDelta ? topCategoryDelta.Label : 'Sem dados' }
    ], 'Sem destaques no filtro atual.');
    renderBars('exec-products-bars', [].concat(views.products).sort(sortDesc('ValorDevolvido', 'QuantidadeDevolvida')).slice(0, 10), {
      label:function(row){ return row.Sku; },
      meta:function(row){ return clipText(row.Produto, 46) + ' | % max ' + formatPercent(row.MaxTaxaPct) + ' | % val ' + formatPercent(row.TaxaValorPct); },
      value:function(row){ return row.ValorDevolvido; }, format:function(value){ return formatMoney(value); }, color:function(row){ return criticalityColor(row, '#0f766e'); }, emptyText:'Sem produtos com devolucao na base do periodo.'
    });
    renderBars('exec-sizes-bars', [].concat(views.sizes).sort(sortByCriticality).slice(0, 6), {
      label:function(row){ return row.Tamanho; }, meta:function(row){ return 'Vend ' + formatQty(row.QuantidadeVendida) + ' | Dev ' + formatQty(row.QuantidadeDevolvida); },
      value:function(row){ return row.MaxTaxaPct; }, format:function(value){ return formatPercent(value); }, color:function(row){ return criticalityColor(row, '#b45309'); }, emptyText:'Sem tamanhos com devolucao no filtro atual.'
    });
    renderBars('exec-clients-bars', [].concat(views.clientBaseRows).sort(sortByCriticality).slice(0, 6), {
      label:function(row){ return clipText(row.Cliente, 36); }, meta:function(row){ return row.ClienteId + ' | % max ' + formatPercent(row.MaxTaxaPct); },
      value:function(row){ return row.MaxTaxaPct; }, format:function(value){ return formatPercent(value); }, color:function(row){ return criticalityColor(row, '#2563eb'); }, emptyText:'Sem clientes com devolucao no filtro atual.'
    });
  }
  function renderOrigin(views){
    function rowsFor(viewRows){ return [].concat(viewRows).sort(sortDesc('DeltaPeriodValorDevolvido', 'ValorDevolvido')).slice(0, 12); }
    const columns = [
      { label:'Dimensao', render:function(row){ return row.Label; } }, { label:'Dev periodo', render:function(row){ return formatMoney(row.PeriodValorDevolvido); } },
      { label:'Dev periodo ant.', render:function(row){ return formatMoney(row.PriorValorDevolvido); } }, { label:'Delta valor', render:function(row){ return formatMoney(row.DeltaPeriodValorDevolvido); } },
      { label:'Delta taxa p.p.', render:function(row){ return formatPercent(row.DeltaPeriodTaxaValorPct); } }
    ];
    renderTable('origem-colecao', columns, rowsFor(views.collections), 'Sem delta relevante por colecao.');
    renderTable('origem-categoria', columns, rowsFor(views.categories), 'Sem delta relevante por categoria.');
    renderTable('origem-loja', columns, rowsFor(views.stores), 'Sem delta relevante por loja.');
    renderTable('origem-sale', columns, rowsFor(views.sale), 'Sem delta relevante por SALE.');
  }
  function renderDimensionPage(targetBars, targetTable, rows, labelField, emptyText){
    const ordered = [].concat(rows).sort(sortByCriticality);
    renderBars(targetBars, ordered.slice(0, 12), {
      label:function(row){ return row[labelField]; }, meta:function(row){ return '% max ' + formatPercent(row.MaxTaxaPct) + ' | Dev periodo ' + formatMoney(row.ValorDevolvido); },
      value:function(row){ return row.MaxTaxaPct; }, format:function(value){ return formatPercent(value); }, color:function(row){ return criticalityColor(row, '#0f766e'); }, emptyText:emptyText
    });
    renderTable(targetTable, [
      { label:'Dimensao', render:function(row){ return row[labelField]; } }, { label:'Vend periodo', render:function(row){ return formatMoney(row.ValorVendido); } },
      { label:'Dev periodo', render:function(row){ return formatMoney(row.ValorDevolvido); } }, { label:'Qtd vend periodo', render:function(row){ return formatQty(row.QuantidadeVendida); } },
      { label:'Qtd dev periodo', render:function(row){ return formatQty(row.QuantidadeDevolvida); } }, { label:'% val periodo', render:function(row){ return formatPercent(row.TaxaValorPct); } },
      { label:'% qtd periodo', render:function(row){ return formatPercent(row.TaxaQuantidadePct); } }, { label:'% max periodo', render:function(row){ return formatPercent(row.MaxTaxaPct); } },
      { label:'Delta dev periodo', render:function(row){ return formatMoney(row.DeltaPeriodValorDevolvido); } }, { label:'Criticidade', render:function(row){ return criticalityTagHtml(row); }, html:true }
    ], ordered, emptyText);
  }
  function renderSalePage(rows){
    const ordered = [].concat(rows).sort(sortDesc('ValorDevolvido', 'QuantidadeDevolvida'));
    renderTable('sale-table', [
      { label:'Grupo', render:function(row){ return row.Label; } }, { label:'Vend periodo', render:function(row){ return formatMoney(row.ValorVendido); } },
      { label:'Dev periodo', render:function(row){ return formatMoney(row.ValorDevolvido); } }, { label:'Qtd vend periodo', render:function(row){ return formatQty(row.QuantidadeVendida); } },
      { label:'Qtd dev periodo', render:function(row){ return formatQty(row.QuantidadeDevolvida); } }, { label:'% val periodo', render:function(row){ return formatPercent(row.TaxaValorPct); } },
      { label:'% qtd periodo', render:function(row){ return formatPercent(row.TaxaQuantidadePct); } }, { label:'% max periodo', render:function(row){ return formatPercent(row.MaxTaxaPct); } },
      { label:'Delta dev periodo', render:function(row){ return formatMoney(row.DeltaPeriodValorDevolvido); } }, { label:'Criticidade', render:function(row){ return criticalityTagHtml(row); }, html:true }
    ], ordered, 'Sem dados para comparar SALE e NAO_SALE.');
    renderBars('sale-bars', ordered, {
      label:function(row){ return row.Label; }, meta:function(row){ return '% val ' + formatPercent(row.TaxaValorPct) + ' | qtd dev ' + formatQty(row.QuantidadeDevolvida); },
      value:function(row){ return row.ValorDevolvido; }, format:function(value){ return formatMoney(value); }, color:function(row){ return row.Label === 'SALE' ? '#0f766e' : '#2563eb'; }, emptyText:'Sem dados para o comparativo SALE.'
    });
  }
  function renderClientsPage(views){
    const byValue = [].concat(views.clients).sort(sortDesc('ValorDevolvido', 'QuantidadeDevolvida'));
    const byRate = [].concat(views.clientBaseRows).sort(sortByCriticality);
    const byQty = [].concat(views.clients).sort(sortDesc('QuantidadeDevolvida', 'ValorDevolvido'));
    const clientColumns = [
      { label:'Cliente', render:function(row){ return clipText(row.Cliente, 44); } }, { label:'Codigo', render:function(row){ return row.ClienteId; } },
      { label:'Vend periodo', render:function(row){ return formatMoney(row.ValorVendido); } }, { label:'Dev periodo', render:function(row){ return formatMoney(row.ValorDevolvido); } },
      { label:'Qtd vend periodo', render:function(row){ return formatQty(row.QuantidadeVendida); } }, { label:'Qtd dev periodo', render:function(row){ return formatQty(row.QuantidadeDevolvida); } },
      { label:'% val periodo', render:function(row){ return formatPercent(row.TaxaValorPct); } }, { label:'% qtd periodo', render:function(row){ return formatPercent(row.TaxaQuantidadePct); } },
      { label:'% max periodo', render:function(row){ return formatPercent(row.MaxTaxaPct); } }, { label:'Criticidade', render:function(row){ return criticalityTagHtml(row); }, html:true }
    ];
    renderTable('clients-critical', clientColumns, views.criticalClients, 'Nenhum cliente critico com base minima no filtro atual.');
    renderTable('clients-value', clientColumns, byValue.slice(0, 25), 'Sem clientes com devolucao no filtro atual.');
    renderTable('clients-rate', clientColumns, byRate.slice(0, 25), 'Sem clientes com base suficiente para ranking de taxa.');
    renderTable('clients-qty', clientColumns, byQty.slice(0, 25), 'Sem clientes com devolucao por quantidade no filtro atual.');
  }
  function renderProductsPage(views){
    const impact = [].concat(views.products).sort(sortDesc('ValorDevolvido', 'QuantidadeDevolvida'));
    const critical = views.criticalProducts;
    const monitor = views.monitorProducts;
    const byQty = [].concat(views.products).sort(sortDesc('QuantidadeDevolvida', 'ValorDevolvido'));
    const columns = [
      { label:'SKU', render:function(row){ return row.Sku; } }, { label:'Produto', render:function(row){ return clipText(row.Produto, 42); } },
      { label:'Colecao', render:function(row){ return row.Colecao; } }, { label:'Categoria', render:function(row){ return row.CategoriaPainel; } },
      { label:'Tamanho', render:function(row){ return row.Tamanho; } }, { label:'Loja', render:function(row){ return row.Loja; } },
      { label:'Vend periodo', render:function(row){ return formatMoney(row.ValorVendido); } }, { label:'Dev periodo', render:function(row){ return formatMoney(row.ValorDevolvido); } },
      { label:'Qtd vend periodo', render:function(row){ return formatQty(row.QuantidadeVendida); } }, { label:'Qtd dev periodo', render:function(row){ return formatQty(row.QuantidadeDevolvida); } },
      { label:'% val periodo', render:function(row){ return formatPercent(row.TaxaValorPct); } }, { label:'% qtd periodo', render:function(row){ return formatPercent(row.TaxaQuantidadePct); } },
      { label:'% max periodo', render:function(row){ return formatPercent(row.MaxTaxaPct); } }, { label:'Delta dev periodo', render:function(row){ return formatMoney(row.DeltaPeriodValorDevolvido); } },
      { label:'Criticidade', render:function(row){ return criticalityTagHtml(row); }, html:true }
    ];
    renderTable('products-critical', columns, critical.slice(0, 30), 'Nenhum produto critico com base minima no filtro atual.');
    renderTable('products-impact', columns, impact.slice(0, 30), 'Sem produtos com devolucao no filtro atual.');
    renderTable('products-monitor', columns, monitor.slice(0, 30), 'Nenhum produto em monitoramento no filtro atual.');
    renderTable('products-qty', columns, byQty.slice(0, 30), 'Sem produtos com devolucao por quantidade no filtro atual.');
  }
  function renderSizesPage(rows){
    const ordered = [].concat(rows).sort(sortByCriticality);
    renderTable('sizes-table', [
      { label:'Tamanho', render:function(row){ return row.Tamanho; } }, { label:'Qtd vend periodo', render:function(row){ return formatQty(row.QuantidadeVendida); } },
      { label:'Qtd dev periodo', render:function(row){ return formatQty(row.QuantidadeDevolvida); } }, { label:'% qtd periodo', render:function(row){ return formatPercent(row.TaxaQuantidadePct); } },
      { label:'Valor vend periodo', render:function(row){ return formatMoney(row.ValorVendido); } }, { label:'Valor dev periodo', render:function(row){ return formatMoney(row.ValorDevolvido); } },
      { label:'% val periodo', render:function(row){ return formatPercent(row.TaxaValorPct); } }, { label:'% max periodo', render:function(row){ return formatPercent(row.MaxTaxaPct); } },
      { label:'Criticidade', render:function(row){ return criticalityTagHtml(row); }, html:true }
    ], ordered, 'Sem tamanhos com devolucao no filtro atual.');
    renderBars('sizes-bars', ordered.slice(0, 8), {
      label:function(row){ return row.Tamanho; }, meta:function(row){ return 'Qtd vend ' + formatQty(row.QuantidadeVendida) + ' | Qtd dev ' + formatQty(row.QuantidadeDevolvida); },
      value:function(row){ return row.MaxTaxaPct; }, format:function(value){ return formatPercent(value); }, color:function(row){ return criticalityColor(row, '#b45309'); }, emptyText:'Sem tamanhos para analisar no filtro atual.'
    });
    const top = ordered[0];
    document.getElementById('sizes-note').innerHTML = top
      ? ('<strong>Exemplo de leitura</strong><div class="muted">Tamanho ' + escapeHtml(top.Tamanho) + ' teve ' + escapeHtml(formatQty(top.QuantidadeDevolvida)) + ' devolucoes, equivalentes a ' + escapeHtml(formatPercent(top.TaxaQuantidadePct)) + ' da quantidade vendida e ' + escapeHtml(formatPercent(top.TaxaValorPct)) + ' do valor vendido na base fixa do periodo.</div>')
      : '<strong>Sem destaque</strong><div class="muted">Nao ha tamanhos com devolucao no filtro atual.</div>';
  }
  function renderMotivosPage(ctx, views){
    renderBars('motivos-bars', views.motives.slice(0, 12), {
      label:function(row){ return row.Label; }, meta:function(row){ return 'Qtd solicitada ' + formatQty(row.QuantidadeSolicitada); },
      value:function(row){ return row.Count; }, format:function(value){ return intFmt.format(value) + ' casos'; }, color:function(){ return '#2563eb'; }, emptyText:'Sem motivos operacionais no filtro atual.'
    });
    renderTable('motivos-table', [
      { label:'Motivo', render:function(row){ return row.Label; } }, { label:'Casos', render:function(row){ return intFmt.format(row.Count); } },
      { label:'Qtd solicitada', render:function(row){ return formatQty(row.QuantidadeSolicitada); } }, { label:'Participacao', render:function(row){ return shareText(row.Count, ctx.periodOperational.length); } }
    ], views.motives, 'Sem motivos operacionais no filtro atual.');
    renderBars('themes-bars', views.themes.slice(0, 12), {
      label:function(row){ return row.Label; }, meta:function(row){ return 'Participacao ' + shareText(row.Count, ctx.periodOperational.length); },
      value:function(row){ return row.Count; }, format:function(value){ return intFmt.format(value) + ' casos'; }, color:function(row){ return row.Label === 'RECEBIMENTO_ERRADO' ? '#b42318' : '#0f766e'; }, emptyText:'Sem comentarios classificados no filtro atual.'
    });
    renderTable('themes-table', [
      { label:'Comentario classificado', render:function(row){ return row.Label; } }, { label:'Casos', render:function(row){ return intFmt.format(row.Count); } },
      { label:'Qtd solicitada', render:function(row){ return formatQty(row.QuantidadeSolicitada); } }, { label:'Participacao', render:function(row){ return shareText(row.Count, ctx.periodOperational.length); } }
    ], views.themes, 'Sem comentarios classificados no filtro atual.');
    const wrongShare = ctx.periodOperational.length ? (views.wrongRows.length / ctx.periodOperational.length) * 100 : 0;
    const topWrongStore = views.wrongByStore[0];
    const topWrongProduct = views.wrongByProduct[0];
    document.getElementById('wrong-kpis').innerHTML = ''
      + kpiHtml('Casos RECEBIMENTO_ERRADO', intFmt.format(views.wrongRows.length), 'Participacao no operacional: ' + formatPercent(wrongShare), views.wrongRows.length ? 'red' : 'slate')
      + kpiHtml('Qtd solicitada', formatQty(views.wrongRows.reduce(function(acc, row){ return acc + safeNum(row.QuantidadeSolicitada); }, 0)), 'Somente base operacional.', 'sky')
      + kpiHtml('Loja mais citada', topWrongStore ? topWrongStore.Label : 'Sem dados', topWrongStore ? (intFmt.format(topWrongStore.Count) + ' casos') : 'Sem ocorrencias', 'amber')
      + kpiHtml('Produto mais citado', topWrongProduct ? topWrongProduct.Sku : 'Sem dados', topWrongProduct ? clipText(topWrongProduct.Produto, 34) : 'Sem ocorrencias', 'amber');
    renderTable('wrong-product', [{ label:'SKU', render:function(row){ return row.Sku; } }, { label:'Produto', render:function(row){ return clipText(row.Produto, 44); } }, { label:'Casos', render:function(row){ return intFmt.format(row.Count); } }, { label:'Qtd solicitada', render:function(row){ return formatQty(row.QuantidadeSolicitada); } }], views.wrongByProduct.slice(0, 20), 'Sem casos de recebimento errado por produto.');
    renderTable('wrong-store', [{ label:'Loja', render:function(row){ return row.Label; } }, { label:'Casos', render:function(row){ return intFmt.format(row.Count); } }, { label:'Qtd solicitada', render:function(row){ return formatQty(row.QuantidadeSolicitada); } }], views.wrongByStore.slice(0, 20), 'Sem casos de recebimento errado por loja.');
    renderTable('wrong-category', [{ label:'Categoria', render:function(row){ return row.Label; } }, { label:'Casos', render:function(row){ return intFmt.format(row.Count); } }, { label:'Qtd solicitada', render:function(row){ return formatQty(row.QuantidadeSolicitada); } }], views.wrongByCategory.slice(0, 20), 'Sem casos de recebimento errado por categoria.');
    renderTable('wrong-collection', [{ label:'Colecao', render:function(row){ return row.Label; } }, { label:'Casos', render:function(row){ return intFmt.format(row.Count); } }, { label:'Qtd solicitada', render:function(row){ return formatQty(row.QuantidadeSolicitada); } }], views.wrongByCollection.slice(0, 20), 'Sem casos de recebimento errado por colecao.');
    const commentRows = [].concat(views.wrongRows.length ? views.wrongRows : ctx.periodOperational).sort(function(left, right){ return compareLabel(right.Data, left.Data); }).slice(0, 40).map(function(row){ return { Data:formatDate(row.Data), Plataforma:row.Plataforma, Pedido:row.Pedido || '-', Sku:row.Sku, Produto:clipText(row.Produto, 42), Motivo:row.Motivo, CategoriaComentario:row.CategoriaComentario, Comentario:clipText(row.Comentario || '-', 120) }; });
    renderTable('comments-table', [{ label:'Data', key:'Data' }, { label:'Plataforma', key:'Plataforma' }, { label:'Pedido', key:'Pedido' }, { label:'SKU', key:'Sku' }, { label:'Produto', key:'Produto' }, { label:'Motivo', key:'Motivo' }, { label:'Classificacao', key:'CategoriaComentario' }, { label:'Comentario', key:'Comentario' }], commentRows, 'Sem comentarios para o filtro atual.');
  }
  function renderOperationPage(currentState, ctx, views){
    renderBars('status-bars', views.status.slice(0, 12), { label:function(row){ return row.Label; }, meta:function(row){ return 'Qtd solicitada ' + formatQty(row.QuantidadeSolicitada); }, value:function(row){ return row.Count; }, format:function(value){ return intFmt.format(value) + ' casos'; }, color:function(){ return '#0f766e'; }, emptyText:'Sem status operacionais para o filtro atual.' });
    renderTable('stage-table', [{ label:'Estagio', render:function(row){ return row.Label; } }, { label:'Casos', render:function(row){ return intFmt.format(row.Count); } }, { label:'Qtd solicitada', render:function(row){ return formatQty(row.QuantidadeSolicitada); } }], views.stage, 'Sem estagios operacionais para o filtro atual.');
    renderTable('platform-table', [{ label:'Plataforma', render:function(row){ return row.Label; } }, { label:'Casos', render:function(row){ return intFmt.format(row.Count); } }, { label:'Qtd solicitada', render:function(row){ return formatQty(row.QuantidadeSolicitada); } }], views.platforms, 'Sem plataformas operacionais para o filtro atual.');
    const financeRows = ctx.periodFinance.length;
    const categoryMapped = ctx.periodFinance.filter(function(row){ return row.CategoriaPainel !== 'SEM_GRUPO / SEM_SUBGRUPO'; }).length;
    const storeMapped = ctx.periodFinance.filter(function(row){ return row.Loja !== 'SEM_LOJA'; }).length;
    const saleTagged = ctx.periodFinance.filter(function(row){ return row.SaleBucket !== 'NAO_SALE' || row.FlagSale === false; }).length;
    renderTable('coverage-table', [{ label:'Indicador', key:'Indicador' }, { label:'Leitura', key:'Leitura' }], [
      { Indicador:'Financeiro filtrado', Leitura:intFmt.format(financeRows) + ' linhas' }, { Indicador:'Operacional filtrado', Leitura:intFmt.format(ctx.periodOperational.length) + ' linhas' },
      { Indicador:'Base fixa do periodo', Leitura:ctx.rollingFull ? ('Completa de ' + formatDate(ctx.rollingStart) + ' a ' + formatDate(ctx.rollingEnd)) : ('Incompleta. Historico disponivel desde ' + formatDate(dataMin) + ' para uma base requerida desde ' + formatDate(ctx.rollingStart)) },
      { Indicador:'Categoria mapeada no financeiro', Leitura:financeRows ? (shareText(categoryMapped, financeRows) + ' (' + intFmt.format(categoryMapped) + '/' + intFmt.format(financeRows) + ')') : 'Sem linhas financeiras no periodo' },
      { Indicador:'Loja mapeada no financeiro', Leitura:financeRows ? (shareText(storeMapped, financeRows) + ' (' + intFmt.format(storeMapped) + '/' + intFmt.format(financeRows) + ')') : 'Sem linhas financeiras no periodo' },
      { Indicador:'SALE identificado', Leitura:financeRows ? (shareText(saleTagged, financeRows) + ' (' + intFmt.format(saleTagged) + '/' + intFmt.format(financeRows) + ')') : 'Sem linhas financeiras no periodo' },
      { Indicador:'Ponte operacional ativa', Leitura:ctx.operationalFiltersApplied ? ('Sim. ' + intFmt.format(ctx.rollingBridgeCount) + ' ocorrencias operacionais filtrando o financeiro.') : 'Nao. Filtro financeiro direto pelos campos compartilhados.' }
    ], 'Sem indicadores de cobertura.');
    renderTable('exports-table', [{ label:'Artefato', key:'Artefato' }, { label:'Caminho', key:'Caminho' }], [
      { Artefato:'Painel HTML', Caminho:rawText(meta.htmlFile) || '-' }, { Artefato:'Fato financeiro', Caminho:rawText(meta.factFinanceFile) || '-' },
      { Artefato:'Fato operacional', Caminho:rawText(meta.factOperationalFile) || '-' }, { Artefato:'Fato legado', Caminho:rawText(meta.factLegacyFile) || '-' }
    ], 'Sem artefatos exportados.');
  }
  function renderDetailsPage(ctx){
    const financeRows = [].concat(ctx.periodFinance).sort(function(left, right){ return compareLabel(right.Data, left.Data); }).slice(0, 120);
    renderTable('detail-finance', [{ label:'Data', render:function(row){ return formatDate(row.Data); } }, { label:'Pedido', render:function(row){ return row.Pedido || '-'; } }, { label:'Cliente', render:function(row){ return clipText(row.ClienteLabel, 36); } }, { label:'SKU', render:function(row){ return row.Sku; } }, { label:'Produto', render:function(row){ return clipText(row.Produto, 36); } }, { label:'Colecao', render:function(row){ return row.Colecao; } }, { label:'Categoria', render:function(row){ return row.CategoriaPainel; } }, { label:'Loja', render:function(row){ return row.Loja; } }, { label:'SALE', render:function(row){ return row.SaleBucket; } }, { label:'Valor vendido', render:function(row){ return formatMoney(row.ValorVendido); } }, { label:'Valor devolvido', render:function(row){ return formatMoney(row.ValorDevolvido); } }, { label:'Qtd vendida', render:function(row){ return formatQty(row.QuantidadeVendida); } }, { label:'Qtd devolvida', render:function(row){ return formatQty(row.QuantidadeDevolvida); } }], financeRows, 'Sem detalhamento financeiro no filtro atual.');
    const opRows = [].concat(ctx.periodOperational).sort(function(left, right){ return compareLabel(right.Data, left.Data); }).slice(0, 120);
    renderTable('detail-operational', [{ label:'Data', render:function(row){ return formatDate(row.Data); } }, { label:'Plataforma', render:function(row){ return row.Plataforma; } }, { label:'Pedido', render:function(row){ return row.Pedido || '-'; } }, { label:'Cliente', render:function(row){ return clipText(row.Cliente, 30); } }, { label:'SKU', render:function(row){ return row.Sku; } }, { label:'Produto', render:function(row){ return clipText(row.Produto, 34); } }, { label:'Motivo', render:function(row){ return row.Motivo; } }, { label:'Comentario classificado', render:function(row){ return row.CategoriaComentario; } }, { label:'Status', render:function(row){ return row.StatusSolicitacao; } }, { label:'Comentario', render:function(row){ return clipText(row.Comentario || '-', 120); } }], opRows, 'Sem detalhamento operacional no filtro atual.');
  }
  function renderFooter(currentState, ctx){
    const footer = document.getElementById('footer-note');
    footer.innerHTML = '<strong>Regras aplicadas</strong><br>'
      + '1. Valores financeiros seguem apenas a base oficial. Nenhum valor de Genius ou Troque Facil entra em vendido/devolvido.<br>'
      + '2. Comentarios, motivos, status e plataforma sao operacionais e servem para contexto, ranking e filtros.<br>'
      + '3. Taxas por cliente, produto, tamanho, colecao, categoria, loja e SALE usam a base fixa de ' + escapeHtml(formatDate(ctx.rollingStart)) + ' a ' + escapeHtml(formatDate(ctx.rollingEnd)) + '.<br>'
      + '4. Criticidade = MAX(% devolucao por valor, % devolucao por quantidade). Classificacao: CRITICO >= 80%, ATENCAO entre 50% e 79%, SAUDAVEL abaixo de 50%.<br>'
      + '5. Base minima: cliente com qtd vendida >= 2 ou valor vendido >= R$ 200; produto com qtd vendida >= 5; tamanho com qtd vendida >= 10. Sem base minima, a linha fica como base baixa.<br>'
      + '6. Se o historico nao cobrir todo o periodo base solicitado, os alertas criticos ficam bloqueados e o painel sinaliza base insuficiente.';
  }
  function attachNav(){
    Array.from(document.querySelectorAll('.nav button')).forEach(function(button){
      button.addEventListener('click', function(){
        Array.from(document.querySelectorAll('.nav button')).forEach(function(item){ item.classList.remove('active'); });
        Array.from(document.querySelectorAll('.page')).forEach(function(page){ page.classList.remove('active'); });
        button.classList.add('active');
        const page = document.getElementById('page-' + button.getAttribute('data-page'));
        if(page){ page.classList.add('active'); }
      });
    });
  }
  function renderAll(){
    const validation = validateState(readStateFromControls());
    if(!validation.ok){ controls.alertBox.style.display = 'block'; controls.alertBox.textContent = validation.message; return; }
    state = validation.value;
    const ctx = buildContext(state);
    const views = buildViews(ctx);
    renderFilterStatus(state, ctx);
    renderExecutive(ctx, views);
    renderOrigin(views);
    renderDimensionPage('colecoes-bars', 'colecoes-table', views.collections, 'Label', 'Sem colecoes com devolucao no filtro atual.');
    renderDimensionPage('categorias-bars', 'categorias-table', views.categories, 'Label', 'Sem categorias com devolucao no filtro atual.');
    renderDimensionPage('lojas-bars', 'lojas-table', views.stores, 'Label', 'Sem lojas com devolucao no filtro atual.');
    renderSalePage(views.sale);
    renderClientsPage(views);
    renderProductsPage(views);
    renderSizesPage(views.sizes);
    renderMotivosPage(ctx, views);
    renderOperationPage(state, ctx, views);
    renderDetailsPage(ctx);
    renderFooter(state, ctx);
  }

  controls.apply.addEventListener('click', renderAll);
  controls.reset.addEventListener('click', function(){ resetControls(); renderAll(); });
  populateControls();
  attachNav();
  renderAll();
})();









