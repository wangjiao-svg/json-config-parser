(function () {
  const jsonInput = document.getElementById('jsonInput');
  const parseError = document.getElementById('parseError');
  const tableHeadRow = document.getElementById('tableHeadRow');
  const tableBody = document.querySelector('#tableOutput tbody');
  const formatBtn = document.getElementById('formatBtn');
  const clearBtn = document.getElementById('clearBtn');
  const mappingCsv = document.getElementById('mappingCsv');
  const standardCsv = document.getElementById('standardCsv');
  const roiStatus = document.getElementById('roiStatus');

  /** 参考表写死在代码里，不对用户展示 */
  const REFERENCE_SHEET_ID = '1Fw8CuAG_MC9HKReY-U--xVgzqKpYle_UDEQU3YR5Q0o';
  /** 道具映射在第二个 sheet 时必填：打开表格 → 点「道具映射」标签 → 地址栏 #gid= 后面的数字填这里（填一次即可） */
  const REFERENCE_MAPPING_GID = '945420337';

  let lastValidJson = null;

  /** 从 Google 表格链接解析 spreadsheet ID */
  function getSpreadsheetId(url) {
    const u = (url || '').trim();
    const m = u.match(/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    return m ? m[1] : null;
  }

  /** 从表格页面 HTML 里解析出所有 sheet 的 gid（用于自动找道具映射 sheet） */
  function extractSheetGidsFromHtml(html) {
    const gids = new Set();
    [/gid[=:]["']?(\d+)/gi, /["']gid["']\s*:\s*["']?(\d+)/gi, /sheet[^a-z]*id["']?\s*:\s*["']?(\d+)/gi].forEach((re) => {
      let m;
      while ((m = re.exec(html)) !== null) gids.add(m[1]);
    });
    return Array.from(gids);
  }

  /** 判断 CSV 首行是否像道具映射表头（含 道具id、对应类型、倍数） */
  function looksLikeMappingHeader(headerRow) {
    const norm = (s) => String(s).trim().toLowerCase().replace(/\s/g, '');
    const hasId = headerRow.some((h) => norm(h) === '道具id' || (norm(h).includes('道具') && norm(h).includes('id')));
    const hasType = headerRow.some((h) => norm(h) === '对应类型' || norm(h).includes('类型'));
    const hasMult = headerRow.some((h) => norm(h) === '倍数' || norm(h).includes('倍数'));
    return hasId && hasType && hasMult;
  }

  /** 从代码内写死的表格链接加载参考表：先加载 gid=0（礼包标准），再自动发现并加载道具映射 sheet */
  function loadFromGoogleSheet() {
    const id = REFERENCE_SHEET_ID;
    if (!id) {
      if (roiStatus) roiStatus.textContent = '未配置参考表';
      return;
    }
    const url0 = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=csv&gid=0';
    const proxy0 = 'https://api.allorigins.win/raw?url=' + encodeURIComponent(url0);
    if (roiStatus) roiStatus.textContent = '正在加载参考表（礼包标准 + 道具映射）…';

    function parseCsvToRows(csvText) {
      const raw = (csvText || '').replace(/^\uFEFF/, '').trim();
      const lines = raw.split(/\r?\n/).filter(Boolean);
      const sep = lines[0] && lines[0].includes('\t') ? '\t' : ',';
      return lines.map((line) => {
        const out = [];
        let cur = '';
        let inQ = false;
        for (let i = 0; i < line.length; i++) {
          const ch = line[i];
          if (ch === '"') inQ = !inQ;
          else if (!inQ && (ch === sep || ch === '\t')) { out.push(cur.replace(/^"|"$/g, '').trim()); cur = ''; }
          else cur += ch;
        }
        out.push(cur.replace(/^"|"$/g, '').trim());
        return out;
      });
    }

    function finishWithMapping(mapCsvText) {
      if (mappingCsv && mapCsvText) mappingCsv.value = mapCsvText;
      const itemMap = buildItemMap(mappingCsv ? mappingCsv.value : '');
      const standardMap = buildStandardMap(standardCsv ? standardCsv.value : '');
      updateRoiStatus(itemMap, standardMap);
      if (roiStatus) roiStatus.textContent = '参考表已加载（礼包标准 + 道具映射 ' + itemMap.size + ' 条）。';
      if (lastValidJson && lastValidJson.arraysList) renderAll(lastValidJson.arraysList, lastValidJson.pricesList);
    }

    fetch(proxy0)
      .then((r) => r.text())
      .then((csvText) => {
        const rows = parseCsvToRows(csvText);
        if (rows.length < 2) {
          if (roiStatus) roiStatus.textContent = '礼包标准内容过少，请确认表格已「发布到网页」';
          return;
        }
        const headers = rows[0];
        const sep = headers.join('').includes('\t') ? '\t' : ',';
        const mappingCol = headers.findIndex((h) => String(h).trim() === '道具id' || String(h).replace(/\s/g, '').toLowerCase().includes('道具id'));
        const standardText = mappingCol >= 1
          ? rows.map((row) => row.slice(0, mappingCol).join(sep)).join('\n')
          : rows.map((row) => row.join(sep)).join('\n');
        if (standardCsv) standardCsv.value = standardText;

        let mappingText = '';
        if (mappingCol >= 1) {
          mappingText = rows.map((row) => row.slice(mappingCol).join(sep)).join('\n');
          if (mappingCsv) mappingCsv.value = mappingText;
        }

        if (mappingCol >= 1 && mappingText.split(/\r?\n/).length >= 2) {
          finishWithMapping(mappingText);
          return;
        }

        const mappingGid = (REFERENCE_MAPPING_GID && String(REFERENCE_MAPPING_GID).trim()) || '';
        if (mappingGid) {
          if (roiStatus) roiStatus.textContent = '正在加载道具映射 sheet（gid=' + mappingGid + '）…';
          const urlMap = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=csv&gid=' + mappingGid;
          const proxyMap = 'https://api.allorigins.win/raw?url=' + encodeURIComponent(urlMap);
          return fetch(proxyMap)
            .then((r) => r.text())
            .then((mapCsv) => {
              const mapRows = parseCsvToRows(mapCsv);
              if (mapRows.length >= 2 && looksLikeMappingHeader(mapRows[0])) {
                const mapSep = (mapRows[0] && mapRows[0].join('').includes('\t')) ? '\t' : ',';
                finishWithMapping(mapRows.map((row) => row.join(mapSep)).join('\n'));
                return;
              }
              if (roiStatus) roiStatus.textContent = '礼包标准已加载；gid=' + mappingGid + ' 的表头不是道具映射格式，道具映射: 0 条。';
              if (lastValidJson && lastValidJson.arraysList) renderAll(lastValidJson.arraysList, lastValidJson.pricesList);
            })
            .catch((err) => {
              if (roiStatus) roiStatus.textContent = '礼包标准已加载；道具映射 sheet 加载失败: ' + (err.message || '网络错误') + '。请确认 REFERENCE_MAPPING_GID 正确且表格已发布。';
              roiStatus.className = 'roi-status warn';
              if (lastValidJson && lastValidJson.arraysList) renderAll(lastValidJson.arraysList, lastValidJson.pricesList);
            });
        }

        if (roiStatus) roiStatus.textContent = '礼包标准已加载，正在查找道具映射 sheet…';
        const editUrl = 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
        const proxyEdit = 'https://api.allorigins.win/raw?url=' + encodeURIComponent(editUrl);
        return fetch(proxyEdit)
          .then((r) => r.text())
          .then((html) => {
            let gids = extractSheetGidsFromHtml(html);
            let otherGids = gids.filter((g) => g !== '0').filter((g, i, a) => a.indexOf(g) === i);
            if (otherGids.length === 0) {
              otherGids = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
            }
            let next = 0;
            function tryNext() {
              if (next >= otherGids.length) {
                if (roiStatus) roiStatus.textContent = '礼包标准已加载；未找到含「道具id、对应类型、倍数」的 sheet，道具映射: 0 条。';
                if (lastValidJson && lastValidJson.arraysList) renderAll(lastValidJson.arraysList, lastValidJson.pricesList);
                return;
              }
              const gid = otherGids[next++];
              const urlMap = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=csv&gid=' + gid;
              const proxyMap = 'https://api.allorigins.win/raw?url=' + encodeURIComponent(urlMap);
              fetch(proxyMap)
                .then((r) => r.text())
                .then((mapCsv) => {
                  const mapRows = parseCsvToRows(mapCsv);
                  if (mapRows.length >= 2 && looksLikeMappingHeader(mapRows[0])) {
                    const mapSep = (mapRows[0] && mapRows[0].join('').includes('\t')) ? '\t' : ',';
                    const mapText = mapRows.map((row) => row.join(mapSep)).join('\n');
                    finishWithMapping(mapText);
                    return;
                  }
                  tryNext();
                })
                .catch(() => tryNext());
            }
            tryNext();
          })
          .catch(() => {
            if (roiStatus) roiStatus.textContent = '礼包标准已加载；无法读取表格结构，道具映射: 0 条。';
            if (lastValidJson && lastValidJson.arraysList) renderAll(lastValidJson.arraysList, lastValidJson.pricesList);
          });
      })
      .catch((err) => {
        if (roiStatus) roiStatus.textContent = '加载失败: ' + (err.message || '网络错误') + '。请确认表格已「发布到网页」。';
        roiStatus.className = 'roi-status warn';
      })
      .finally(() => {});
  }

  /** 解析单行 CSV（支持引号内逗号/制表符） */
  function parseCsvLine(line, sep) {
    const out = [];
    let cur = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        inQuotes = !inQuotes;
      } else if (inQuotes) {
        cur += ch;
      } else if (ch === sep) {
        out.push(cur.replace(/^"|"$/g, '').trim());
        cur = '';
      } else {
        cur += ch;
      }
    }
    out.push(cur.replace(/^"|"$/g, '').trim());
    return out;
  }

  /** 解析 CSV 行（支持制表符、逗号或分号，引号内不分割） */
  function parseCsvLines(text) {
    if (!text || !text.trim()) return [];
    const raw = text.replace(/^\uFEFF/, '').trim();
    const lines = raw.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
    let sep = ',';
    if (lines[0]) {
      if (lines[0].includes('\t')) sep = '\t';
      else if (lines[0].includes(';') && !lines[0].includes(',')) sep = ';';
    }
    return lines.map((line) => parseCsvLine(line, sep));
  }

  /** 全角数字转半角 */
  function halfWidthDigits(s) {
    return String(s).replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  }

  /** 规范道具 id（处理科学计数法、逗号、首尾引号、全角数字等，保证数字 id 一致） */
  function normalizeItemId(rawId) {
    if (rawId == null) return '';
    let s = String(rawId).trim().replace(/^["']+/, '').replace(/["']+$/, '').trim();
    s = halfWidthDigits(s).replace(/[\uFEFF\u200B-\u200D\u2060]/g, '').replace(/,/g, '.');
    const num = parseFloat(s);
    if (Number.isFinite(num) && num === Math.floor(num)) return String(Math.floor(num));
    return s || '';
  }

  /** 构建道具映射：id -> { type, multiplier }。仅解析 道具id、对应类型、倍数 三列，道具描述列不参与。 */
  function buildItemMap(csvText) {
    const rows = parseCsvLines(csvText);
    if (rows.length < 2) return new Map();
    const headerRow = rows[0].map((h) => String(h).trim().replace(/\uFEFF/g, ''));
    const headerNorm = headerRow.map((h) => h.toLowerCase().replace(/\s/g, '').replace(/\uFEFF/g, ''));
    let idCol = headerNorm.findIndex((h) => h === '道具id');
    if (idCol < 0) idCol = headerNorm.findIndex((h) => h === 'id' || (h.includes('道具') && h.includes('id') && !h.includes('类型')));
    if (idCol < 0) idCol = headerNorm.findIndex((h) => h.includes('id') && !h.includes('类型'));
    let typeCol = headerNorm.findIndex((h) => h === '对应类型');
    if (typeCol < 0) typeCol = headerNorm.findIndex((h) => h === '类型' || (h.includes('类型') && h.includes('对应')));
    if (typeCol < 0) typeCol = headerNorm.findIndex((h) => h.includes('类型'));
    let multCol = headerNorm.findIndex((h) => h === '倍数');
    if (multCol < 0) multCol = headerNorm.findIndex((h) => h.includes('倍数'));
    if (idCol < 0 || typeCol < 0 || multCol < 0) return new Map();
    const map = new Map();
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      let rawId = row[idCol];
      if ((rawId == null || String(rawId).trim() === '') && row.length > idCol + 1) rawId = row[idCol + 1];
      const id = normalizeItemId(rawId);
      if (!id) continue;
      let typeColUse = typeCol;
      let multColUse = multCol;
      if (rawId !== row[idCol] && row.length > typeCol + 1) {
        typeColUse = typeCol + 1;
        multColUse = multCol + 1;
      }
      let type = String(row[typeColUse] ?? '').trim();
      let mult = parseFloat(String(row[multColUse]).replace(',', '.')) || 0;
      // 道具描述中含逗号（如 50,000粮食、1,234,567钢铁）时 CSV 会多拆多列，类型/倍数整体右移；向前扫描直到找到「类型非数字片段 + 倍数>0」
      const looksLikeFragment = (t) => t && (/^\d+$/.test(t) || /^\d+[^（(]*$/.test(t));
      if ((mult === 0 || looksLikeFragment(type)) && row.length > multColUse + 1) {
        const maxShift = Math.min(10, row.length - multColUse - 1);
        for (let s = 1; s <= maxShift; s++) {
          const nextType = String(row[typeColUse + s] ?? '').trim();
          const nextMult = parseFloat(String(row[multColUse + s]).replace(',', '.')) || 0;
          if (nextType && nextMult > 0 && !looksLikeFragment(nextType)) {
            type = nextType;
            mult = nextMult;
            break;
          }
        }
      }
      const entry = { type, multiplier: mult };
      map.set(id, entry);
      const rawStr = rawId != null ? String(rawId).trim() : '';
      if (rawStr && rawStr !== id) map.set(rawStr, entry);
      if (/^\d+$/.test(id)) {
        const numId = parseInt(id, 10);
        map.set(numId, entry);
        map.set(Number(id), entry);
      }
    }
    return map;
  }

  /** 构建礼包标准：(价格, 类型) -> 标准数量。首行为表头；价格在首列或第二列（首列为 —/空/序号 或 1,2,3 则价格在第二列） */
  function buildStandardMap(csvText) {
    const rows = parseCsvLines(csvText);
    if (rows.length < 2) return new Map();
    const headers = rows[0].map((h) => String(h).trim());
    const firstRow = rows[1];
    let priceCol = 0;
    const v0 = parseFloat(String(firstRow[0] ?? '').replace(',', '.'));
    const v1 = parseFloat(String(firstRow[1] ?? '').replace(',', '.'));
    const h0 = (headers[0] ?? '').trim();
    const looksLikeRowNum = (v) => Number.isInteger(v) && v >= 1 && v <= 200;
    const looksLikePrice = (v) => Number.isFinite(v) && v >= 0.99 && v < 10000;
    if (looksLikeRowNum(v0) && looksLikePrice(v1)) {
      priceCol = 1;
    } else if (!looksLikePrice(v0) && looksLikePrice(v1) && (h0 === '—' || h0 === '' || h0 === 'P2礼包标准' || /^\d+$/.test(h0))) {
      priceCol = 1;
    }
    const typeCols = headers.slice(priceCol + 1).map((h) => String(h).trim()).filter(Boolean);
    const map = new Map();
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const priceRaw = String(row[priceCol] ?? '').trim().replace(',', '.');
      const priceNum = parseFloat(priceRaw);
      if (!Number.isFinite(priceNum) || priceNum < 0) continue;
      const priceKey = String(priceNum);
      const rowMap = new Map();
      for (let c = 0; c < typeCols.length; c++) {
        const type = typeCols[c];
        const val = parseFloat(String(row[priceCol + 1 + c] ?? '').replace(',', '.')) || 0;
        rowMap.set(type, val);
        rowMap.set(type.trim(), val);
        rowMap.set(type.replace(/\s/g, ''), val);
      }
      map.set(priceKey, rowMap);
    }
    return map;
  }

  function normalizePrice(p) {
    const s = String(p).trim().replace(',', '.');
    const n = parseFloat(s);
    if (Number.isFinite(n)) return String(n);
    return s || p;
  }

  /**
   * 礼包标准查表：价格 → 行（P2礼包标准列），类型 → 列（如高抽列），取该单元格值作为分母。
   * 示例：价格 4.99 + 类型「高抽」=> 标准值(分母)=5
   */
  function findStandardQty(standardMap, priceStr, typeName) {
    const typeKey = String(typeName).trim();
    const typeKeyNoSpace = typeKey.replace(/\s/g, '');
    const priceKey = normalizePrice(priceStr);
    function normParen(s) { return String(s || '').replace(/（/g, '(').replace(/）/g, ')'); }
    function findInRow(row) {
      if (!row) return null;
      let q = row.get(typeKey) ?? row.get(typeKeyNoSpace);
      if (q != null && Number(q) > 0) return Number(q);
      const typeP = normParen(typeKey);
      for (const [k, v] of row) {
        const kt = (k || '').trim().replace(/\s/g, '');
        const kp = normParen(kt);
        if ((kt === typeKey || kt === typeKeyNoSpace || kp === typeP || kt.includes(typeKey) || typeKey.includes(kt)) && Number(v) > 0) return Number(v);
      }
      return null;
    }
    const row = standardMap.get(priceKey);
    let q = findInRow(row);
    if (q != null) return q;
    const keys = Array.from(standardMap.keys()).filter((k) => !isNaN(parseFloat(k))).map(Number).sort((a, b) => a - b);
    const p = parseFloat(priceStr);
    if (!Number.isFinite(p) || keys.length === 0) return null;
    let best = keys[0];
    for (const k of keys) {
      if (Math.abs(k - p) < Math.abs(best - p)) best = k;
    }
    return findInRow(standardMap.get(String(best)));
  }

  /**
   * 单行 ROI：对每个道具 → 用道具id 在【道具映射】取 类型+倍数，用 价格+类型 在【礼包标准】取 分母，
   * 该项贡献 = 道具数量×倍数/分母；整行 ROI = 各项贡献之和。
   */
  function calcRowRoi(itemMap, standardMap, priceStr, items) {
    const result = calcPerItemRois(itemMap, standardMap, priceStr, items);
    return result ? result.total : null;
  }

  /** 从道具映射中查找条目，兼容多种 id 形式（字符串、数字、去前导零等） */
  function getItemMapEntry(itemMap, rawId) {
    if (rawId == null || rawId === '') return null;
    const idNorm = normalizeItemId(rawId);
    if (!idNorm) return null;
    const numId = /^\d+$/.test(idNorm) ? parseInt(idNorm, 10) : NaN;
    let entry = itemMap.get(idNorm) || itemMap.get(rawId);
    if (entry) return entry;
    if (Number.isFinite(numId)) entry = itemMap.get(numId) || itemMap.get(Number(rawId));
    if (entry) return entry;
    if (typeof rawId === 'string' && /^\d+$/.test(rawId.replace(/^0+/, ''))) entry = itemMap.get(parseInt(rawId.replace(/^0+/, ''), 10));
    if (entry) return entry;
    for (const [k, v] of itemMap) {
      if (normalizeItemId(k) === idNorm) return v;
    }
    return null;
  }

  /**
   * 返回每个道具的 ROI 及总 ROI：rois[i] 为第 i 个道具的贡献，total 为总和。
   */
  function calcPerItemRois(itemMap, standardMap, priceStr, items) {
    if (!itemMap.size || !standardMap.size || !items.length) return null;
    const priceKey = normalizePrice(String(priceStr || '').trim());
    const rois = [];
    let total = 0;
    for (const it of items) {
      const rawId = it.itemId;
      const qty = Number(it.itemCount) || 0;
      const info = getItemMapEntry(itemMap, rawId);
      if (!info) {
        rois.push(null);
        continue;
      }
      const denominator = findStandardQty(standardMap, priceKey, info.type);
      if (denominator == null || denominator <= 0) {
        rois.push(null);
        continue;
      }
      const contrib = (qty * info.multiplier) / denominator;
      rois.push(contrib);
      total += contrib;
    }
    return rois.some((r) => r != null) ? { rois, total } : null;
  }

  function showError(msg) {
    parseError.textContent = msg;
    parseError.hidden = false;
  }

  function hideError() {
    parseError.hidden = true;
  }

  /** 解析多行输入；每行可为 "价格\tJSON数组" 或纯 JSON 数组 */
  function parseJson() {
    const raw = jsonInput.value.trim();
    if (!raw) {
      lastValidJson = null;
      hideError();
      renderEmpty();
      return null;
    }
    const lines = raw.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
    const arraysList = [];
    const pricesList = [];
    let firstError = null;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      let jsonPart = line;
      let price = '';
      if (line.includes('\t')) {
        const idx = line.indexOf('\t');
        price = line.slice(0, idx).trim();
        jsonPart = line.slice(idx + 1).trim();
      }
      try {
        const data = JSON.parse(jsonPart);
        if (Array.isArray(data)) {
          arraysList.push(data);
          pricesList.push(price);
        } else if (data != null && typeof data === 'object') {
          for (const v of Object.values(data)) {
            if (Array.isArray(v)) {
              arraysList.push(v);
              pricesList.push(price);
            }
          }
        }
      } catch (e) {
        if (!firstError) firstError = `第 ${i + 1} 行 JSON 解析错误：` + e.message;
      }
    }
    if (arraysList.length === 0 && firstError) {
      showError(firstError);
      renderEmpty();
      return null;
    }
    if (firstError) showError(firstError);
    else hideError();
    lastValidJson = { arraysList, pricesList };
    renderAll(arraysList, pricesList);
    return arraysList;
  }

  function renderEmpty() {
    if (tableHeadRow) tableHeadRow.innerHTML = '<th class="group">组</th><th class="price">价格</th><th class="item-id">道具id1</th><th class="item-count">道具数量1</th><th class="roi">ROI1</th><th class="roi total-roi">总ROI</th>';
    if (tableBody) tableBody.innerHTML = '<tr><td colspan="6" class="empty-state">输入或粘贴 JSON 数组后自动解析</td></tr>';
    updateRoiStatus(new Map(), new Map());
  }

  /** 诊断某道具 ROI 为 — 的原因：未找到道具id 还是类型未匹配 */
  function diagnoseRoiItem(itemMap, standardMap, itemId, samplePrice) {
    const idStr = String(itemId);
    const entry = getItemMapEntry(itemMap, itemId);
    if (!entry) return idStr + ': 未在道具映射中找到（检查道具id列是否为纯文本、表中是否有该id）';
    const denom = findStandardQty(standardMap, samplePrice, entry.type);
    if (denom == null || denom <= 0) return idStr + ': 已找到映射(类型=' + entry.type + ')，但礼包标准中无该类型列（检查表头是否含「' + entry.type + '」）';
    return idStr + ': 已匹配(类型=' + entry.type + '，分母=' + denom + ')';
  }

  function updateRoiStatus(itemMap, standardMap, opts) {
    const el = document.getElementById('roiStatus');
    if (!el) return;
    const priceCount = standardMap.size;
    const typeCount = standardMap.size > 0 ? (standardMap.values().next().value?.size ?? 0) : 0;
    let msg = `道具映射: ${itemMap.size} 条 | 礼包标准: ${priceCount} 个价格档位，${typeCount} 种类型`;
    const hasRoi = itemMap.size > 0 && standardMap.size > 0;
    if (hasRoi) {
      let diagPrice = '4.99';
      if (opts && opts.pricesList && opts.groupsDataForRoi) {
        const idx = opts.groupsDataForRoi.findIndex((row) =>
          row.some((it) => it && (String(it.itemId) === '11111002' || Number(it.itemId) === 11111002))
        );
        if (idx >= 0 && opts.pricesList[idx] != null && String(opts.pricesList[idx]).trim() !== '')
          diagPrice = String(opts.pricesList[idx]).trim();
      }
      msg += ' | 诊断 11111002(价格=' + diagPrice + '): ' + diagnoseRoiItem(itemMap, standardMap, '11111002', diagPrice);
    } else if (itemMap.size === 0 && standardMap.size > 0) {
      msg += ' — 道具映射未加载：请打开该表格，点击底部「道具映射」标签，看地址栏末尾 #gid= 后面的数字，复制后填到 app.js 里 REFERENCE_MAPPING_GID = \'数字\'，保存并刷新页面。';
    } else if ((mappingCsv && mappingCsv.value && mappingCsv.value.trim()) || (standardCsv && standardCsv.value && standardCsv.value.trim())) {
      msg += ' — 请检查 CSV 表头（道具映射：道具id、对应类型、倍数；礼包标准：P2礼包标准列为价格，高抽等为类型列）';
    }
    el.textContent = msg;
    el.className = 'roi-status' + (hasRoi ? '' : ' warn');
  }

  function getTypeLabel(val) {
    if (val === null) return 'null';
    if (Array.isArray(val)) return 'array';
    return typeof val;
  }

  function getTypeClass(val) {
    if (val === null) return 'null';
    if (typeof val === 'string') return 'string';
    if (typeof val === 'number') return 'number';
    if (typeof val === 'boolean') return 'boolean';
    if (Array.isArray(val)) return 'bracket';
    return 'key';
  }

  function formatValue(val) {
    if (val === null) return 'null';
    if (typeof val === 'string') return '"' + val + '"';
    if (typeof val === 'boolean') return val ? 'true' : 'false';
    if (typeof val === 'number') return String(val);
    if (Array.isArray(val)) return '[ ... ]';
    return '{ ... }';
  }

  function buildTree(obj, indent = 0) {
    const lines = [];
    const pad = '  '.repeat(indent);

    if (obj === null) {
      lines.push(`${pad}<span class="null">null</span>`);
      return lines.join('\n');
    }

    if (typeof obj !== 'object') {
      const cls = getTypeClass(obj);
      lines.push(`${pad}<span class="${cls}">${escapeHtml(formatValue(obj))}</span>`);
      return lines.join('\n');
    }

    const isArray = Array.isArray(obj);
    const open = isArray ? '[' : '{';
    const close = isArray ? ']' : '}';
    const entries = isArray ? obj.map((v, i) => [String(i), v]) : Object.entries(obj);

    lines.push(`${pad}<span class="bracket">${open}</span>`);
    entries.forEach(([key, val], i) => {
      const isLast = i === entries.length - 1;
      const comma = isLast ? '' : ',';
      const keyPart = isArray ? '' : `<span class="key">"${escapeHtml(key)}"</span>: `;
      const valIsObject = val !== null && typeof val === 'object';
      if (valIsObject) {
        const sub = buildTree(val, indent + 2);
        const subLines = sub.split('\n');
        lines.push(`${pad}  ${keyPart}${subLines[0]}`);
        for (let j = 1; j < subLines.length - 1; j++) lines.push(subLines[j]);
        lines.push(`${pad}  ${subLines[subLines.length - 1]}${comma}`);
      } else {
        const cls = getTypeClass(val);
        lines.push(`${pad}  ${keyPart}<span class="${cls}">${escapeHtml(formatValue(val))}</span>${comma}`);
      }
    });
    lines.push(`${pad}<span class="bracket">${close}</span>`);
    return lines.join('\n');
  }

  function escapeHtml(s) {
    const div = document.createElement('div');
    div.textContent = s;
    return div.innerHTML;
  }

  /** 从数组元素中取 id/val：支持顶层 id/val，或 asset.id、asset.val */
  function pickIdVal(item) {
    if (item == null || typeof item !== 'object') return null;
    const source = item.asset != null && typeof item.asset === 'object' ? item.asset : item;
    if ('id' in source || 'val' in source) {
      return {
        itemId: source.id !== undefined ? source.id : '',
        itemCount: source.val !== undefined ? source.val : ''
      };
    }
    return null;
  }

  /** 判断 id 是否「以 1116 开头」：仅用开头匹配（startsWith），不用包含（includes），避免误伤中间含 1116 的 id */
  function idStartsWith1116(id) {
    const s = String(id).trim();
    return s.startsWith('1116');
  }

  /** 不参与表格展示与 ROI 的 id：仅下列指定 id；若需过滤「以 1116 开头的 id」可设 FILTER_IDS_STARTING_WITH_1116 为 true（仍用 startsWith 判断） */
  const FILTER_ID_SET = new Set([
    '11114303', '11114316', '11114317', '11114318', '11114319', '11114320'
  ]);
  const FILTER_IDS_STARTING_WITH_1116 = true;
  function shouldFilterId(id) {
    const s = String(id).trim();
    if (FILTER_IDS_STARTING_WITH_1116 && idStartsWith1116(s)) return true;
    return FILTER_ID_SET.has(s);
  }

  /** 对单个 [] 内的内容：过滤后按 id 合并，val 累加，每个 id 只输出一行（用于表格展示） */
  function getItemRowsForOneArray(arr) {
    const rows = [];
    if (!Array.isArray(arr)) return rows;
    for (const item of arr) {
      const r = pickIdVal(item);
      if (r && !shouldFilterId(r.itemId)) rows.push(r);
    }
    const byId = new Map();
    for (const r of rows) {
      const id = r.itemId;
      const val = Number(r.itemCount) || 0;
      byId.set(id, (byId.get(id) || 0) + val);
    }
    return Array.from(byId.entries(), ([itemId, itemCount]) => ({ itemId, itemCount }));
  }

  /** 对单个 [] 内全部道具按 id 合并（排除应过滤的 id，如 1116 开头），用于 ROI 计算 */
  function getItemRowsForRoi(arr) {
    const rows = [];
    if (!Array.isArray(arr)) return rows;
    for (const item of arr) {
      const r = pickIdVal(item);
      if (r && !shouldFilterId(r.itemId)) rows.push(r);
    }
    const byId = new Map();
    for (const r of rows) {
      const id = r.itemId;
      const val = Number(r.itemCount) || 0;
      byId.set(id, (byId.get(id) || 0) + val);
    }
    return Array.from(byId.entries(), ([itemId, itemCount]) => ({ itemId, itemCount }));
  }

  /** 多个 [] 时：每组单独合并，返回每组一行；每行为 [ { itemId, itemCount }, ... ] */
  function getGroupsAsRows(arraysList) {
    return arraysList.map((arr) => getItemRowsForOneArray(arr));
  }

  function renderAll(arraysList, pricesList) {
    if (!pricesList) pricesList = arraysList.map(() => '');
    const groupsData = getGroupsAsRows(arraysList);
    const groupsDataForRoi = arraysList.map((arr) => getItemRowsForRoi(arr));
    const maxItems = Math.max(1, ...groupsData.map((g) => g.length));
    const mappingText = (mappingCsv && mappingCsv.value) ? mappingCsv.value : '';
    const standardText = (standardCsv && standardCsv.value) ? standardCsv.value : '';
    const itemMap = buildItemMap(mappingText);
    const standardMap = buildStandardMap(standardText);
    const hasRoi = itemMap.size > 0 && standardMap.size > 0;

    updateRoiStatus(itemMap, standardMap, { pricesList, groupsDataForRoi });

    tableHeadRow.innerHTML =
      '<th class="group">组</th><th class="price">价格</th>' +
      Array.from({ length: maxItems }, (_, i) => `<th class="item-id">道具id${i + 1}</th><th class="item-count">道具数量${i + 1}</th><th class="roi">ROI${i + 1}</th>`).join('') +
      '<th class="roi total-roi">总ROI</th>';

    const colCount = 2 + maxItems * 3 + 1;
    if (groupsData.every((g) => g.length === 0)) {
      tableBody.innerHTML = '<tr><td colspan="' + colCount + '" class="empty-state">未找到包含 id/val 的数组项，请输入如 [{ "id": 1001, "val": 10 }] 的 JSON</td></tr>';
    } else {
      tableBody.innerHTML = groupsData
        .map((items, groupIndex) => {
          const groupNum = groupIndex + 1;
          const price = pricesList[groupIndex] != null ? String(pricesList[groupIndex]) : '';
          const itemsForRoi = groupsDataForRoi[groupIndex] || [];
          const roiResult = hasRoi ? calcPerItemRois(itemMap, standardMap, price, itemsForRoi) : null;
          const perItemRois = roiResult ? roiResult.rois : [];
          const totalRoi = roiResult ? roiResult.total : null;
          const cells = [
            '<td class="group">' + escapeHtml(String(groupNum)) + '</td>',
            '<td class="price">' + escapeHtml(price) + '</td>'
          ];
          for (let i = 0; i < maxItems; i++) {
            const item = items[i];
            if (item) {
              cells.push('<td class="item-id">' + escapeHtml(String(item.itemId)) + '</td>');
              cells.push('<td class="item-count">' + escapeHtml(String(item.itemCount)) + '</td>');
              const r = perItemRois[i];
              cells.push('<td class="roi">' + (r != null ? Number(r).toFixed(2) : '—') + '</td>');
            } else {
              cells.push('<td class="item-id"></td><td class="item-count"></td><td class="roi"></td>');
            }
          }
          cells.push('<td class="roi total-roi">' + (totalRoi != null ? Number(totalRoi).toFixed(2) : '—') + '</td>');
          return '<tr>' + cells.join('') + '</tr>';
        })
        .join('');
    }

  }

  function copyTableToClipboard() {
    const table = document.getElementById('tableOutput');
    if (!table) return;
    const rows = table.querySelectorAll('tr');
    const lines = [];
    rows.forEach((tr) => {
      const cells = tr.querySelectorAll('th, td');
      const text = Array.from(cells).map((c) => (c.textContent || '').trim().replace(/\n/g, ' ')).join('\t');
      lines.push(text);
    });
    const text = lines.join('\n');
    if (!text.trim()) return;
    navigator.clipboard.writeText(text).then(() => {
      const btn = document.getElementById('copyTableBtn');
      if (btn) {
        const orig = btn.textContent;
        btn.textContent = '已复制';
        setTimeout(() => { btn.textContent = orig; }, 1500);
      }
    }).catch(() => {
      const btn = document.getElementById('copyTableBtn');
      if (btn) btn.textContent = '复制失败';
    });
  }

  jsonInput.addEventListener('input', () => parseJson());
  jsonInput.addEventListener('paste', () => setTimeout(parseJson, 10));
  mappingCsv.addEventListener('input', () => {
    if (lastValidJson && lastValidJson.arraysList) renderAll(lastValidJson.arraysList, lastValidJson.pricesList);
  });
  standardCsv.addEventListener('input', () => {
    if (lastValidJson && lastValidJson.arraysList) renderAll(lastValidJson.arraysList, lastValidJson.pricesList);
  });

  formatBtn.addEventListener('click', () => {
    const data = parseJson();
    if (data !== null) jsonInput.value = JSON.stringify(data, null, 2);
  });

  clearBtn.addEventListener('click', () => {
    jsonInput.value = '';
    parseError.hidden = true;
    lastValidJson = null;
    renderEmpty();
  });

  // 参考表链接写死在代码内，页面打开即自动加载，用户无需填写或点击
  loadFromGoogleSheet();

  const copyTableBtn = document.getElementById('copyTableBtn');
  if (copyTableBtn) copyTableBtn.addEventListener('click', copyTableToClipboard);

  parseJson();
})();
