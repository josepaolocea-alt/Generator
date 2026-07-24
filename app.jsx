    const { useState, useEffect, useLayoutEffect, useMemo, useRef, useCallback } = React;

    /* ============================================================
       Constants
       ============================================================ */
    const HOURS_24 = ['12AM','1AM','2AM','3AM','4AM','5AM','6AM','7AM','8AM','9AM','10AM','11AM','12PM','1PM','2PM','3PM','4PM','5PM','6PM','7PM','8PM','9PM','10PM','11PM'];
    const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    const LAYOUT_LABELS = { hourly: 'Hourly Block', flat: 'Flat Table', alarm: 'Alarm Log' };

    const FLAT_COLS = ['Date','Time','Trunk ID','Alias','Active Calls','CPS','Peak','CLZ Total','CLZ CPS','15min ASR','15min ACD','15min PDD'];

    const FONT_FAMILIES = ['Calibri','Arial','Helvetica','Times New Roman','Georgia','Verdana','Tahoma','Trebuchet MS','Courier New','Consolas','Comic Sans MS'];
    const FONT_SIZES = [8,9,10,11,12,14,16,18,20,24,28,32];

    const DATE_FORMATS = [
      { value: 'yyyy-mm-dd',         sample: '2026-06-01' },
      { value: 'yyyy/mm/dd',         sample: '2026/06/01' },
      { value: 'mm/dd/yyyy',         sample: '06/01/2026' },
      { value: 'dd/mm/yyyy',         sample: '01/06/2026' },
      { value: 'd-mmm-yy',           sample: '1-Jun-26' },
      { value: 'd-mmm-yyyy',         sample: '1-Jun-2026' },
      { value: 'mmm d, yyyy',        sample: 'Jun 1, 2026' },
      { value: 'mmmm d, yyyy',       sample: 'June 1, 2026' },
      { value: 'd mmmm yyyy',        sample: '1 June 2026' },
      { value: 'dddd, mmmm d, yyyy', sample: 'Monday, June 1, 2026' },
    ];
    const DATE_FORMAT_VALUES = new Set(DATE_FORMATS.map(d => d.value));

    function previewDateFormat(d, fmt) {
      const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
      const dayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
      const y = d.getFullYear();
      const m = d.getMonth();
      const day = d.getDate();
      const wd = d.getDay();
      const tokens = {
        yyyy: String(y),
        yy:   String(y).slice(-2),
        mmmm: months[m],
        mmm:  months[m].slice(0, 3),
        mm:   String(m + 1).padStart(2, '0'),
        m:    String(m + 1),
        dddd: dayNames[wd],
        ddd:  dayNames[wd].slice(0, 3),
        dd:   String(day).padStart(2, '0'),
        d:    String(day),
      };
      return String(fmt || '').replace(/yyyy|yy|mmmm|mmm|mm|dddd|ddd|dd|m|d/g, t => tokens[t] ?? t);
    }

    const DEFAULT_NOTE_STYLE = {
      noteFontFamily: 'Calibri',
      noteFontSize: 11,
      noteFontColor: 'FF6B7280',
      noteBold: false,
      noteItalic: true,
      noteUnderline: false,
      noteAlign: 'left',
    };

    function applyNoteStyle(ws, rowNum, colSpan, s) {
      ws.mergeCells(rowNum, 1, rowNum, Math.max(1, colSpan));
      const c = ws.getCell(rowNum, 1);
      c.value = s.note;
      c.font = {
        name:      s.noteFontFamily || 'Calibri',
        size:      Number(s.noteFontSize) || 11,
        color:     { argb: s.noteFontColor || 'FF6B7280' },
        bold:      !!s.noteBold,
        italic:    s.noteItalic !== undefined ? !!s.noteItalic : true,
        underline: !!s.noteUnderline,
      };
      c.alignment = {
        wrapText: true,
        vertical: 'middle',
        horizontal: s.noteAlign || 'left',
      };
      const sz = Number(s.noteFontSize) || 11;
      ws.getRow(rowNum).height = Math.max(20, Math.round(sz * 1.8));
    }

    const DEFAULT_SHEETS = [
      { id: 's_alarm',    name: 'SBC ALARM',       layout: 'alarm',  active: true, columns: ['Contains','Severity','Source','Count','Last Occur','SBC','Date Reported','Remarks from L2'], note: '' },
      { id: 's_sbc',      name: 'SBC',             layout: 'hourly', active: true, metrics: ['CIMB IN SBC (ACTIVE CALLS)','GLOBESIP (CIMB) [30 CH]','PLDTSIP (CIMB) [30 CH]','SBC (ACTIVE CALLS)','PLDT SIP','PLDT FCS','SBC2','ETPI SIP (CYN)','GLOBE SIP (CYN,FEDEX,TOKU,AIQON)'], hourStart: 0, hourEnd: 24, note: '' },
      { id: 's_vos',      name: 'VOS',             layout: 'hourly', active: true, metrics: ['IN_ETPI-SIP','IN_PLDT-FCS-CYN','IN_PLDT-SIP','V_PLDT-SIP','V_ETPI-SIP'], hourStart: 0, hourEnd: 24, note: '' },
      { id: 's_active',   name: 'ACTIVE S',        layout: 'hourly', active: true, metrics: ['Active Sessions'], hourStart: 0, hourEnd: 24, note: '' },
      { id: 's_feb',      name: 'Feb ALL OUT',     layout: 'flat',   active: true, columns: FLAT_COLS, note: '' },
      { id: 's_vc',       name: 'VC SIPFCS',       layout: 'flat',   active: true, columns: FLAT_COLS, note: '' },
      { id: 's_vos2',     name: 'VOS2 GSMGW',      layout: 'flat',   active: true, columns: FLAT_COLS, note: '' },
      { id: 's_voscyn',   name: 'VOS CYN GSMGW',   layout: 'hourly', active: true, metrics: ['Activate Calls','CPS','PEAK'], hourStart: 0, hourEnd: 24, note: '' },
      { id: 's_kingsf',   name: 'KINGSFORD INOUT', layout: 'hourly', active: true, metrics: ['Current Inbound','MaxInbound','Current Outbound','Max Outbound'], hourStart: 0, hourEnd: 24, note: '' },
      { id: 's_uno',      name: 'UNOBANK IN',      layout: 'hourly', active: true, metrics: ['Current Inbound','MaxInbound'], hourStart: 0, hourEnd: 24, note: '' },
      { id: 's_myvelox',  name: 'MYVELOX INOUT',   layout: 'hourly', active: true, metrics: ['Current Inbound','MaxInbound','Current Outbound','Max Outbound'], hourStart: 0, hourEnd: 24, note: '' },
    ];

    const STORAGE_KEY = 'mrg_state_v1';
    const TEMPLATES_KEY = 'mrg_templates_v1';
    const EDITOR_TEMPLATES_KEY = 'mrg_editor_templates_v1';
    const RCA_TEMPLATES_KEY = 'mrg_rca_templates_v1';
    const RCA_SIGNATORIES_KEY = 'mrg_rca_signatories_v1';

    /* ============================================================
       Helpers
       ============================================================ */
    const newId = () => 's_' + Math.random().toString(36).slice(2, 9);
    const daysInMonth = (y, m) => new Date(y, m + 1, 0).getDate();
    /* Items in `metrics` / `columns` can be either a string (visible) or
       { label, hidden?: true, dropdown?: Array<string | {label,color}> }
       (extended form kept in config; `hidden` skips on generate, `dropdown`
       adds list validation on alarm-log columns). */
    const itemLabel    = (it) => (typeof it === 'string' ? it : (it?.label ?? ''));
    const itemHidden   = (it) => (typeof it === 'string' ? false : !!it?.hidden);
    const itemDropdown = (it) => (typeof it === 'string' || !it) ? [] : (Array.isArray(it.dropdown) ? it.dropdown : []);
    const itemColor    = (it) => (typeof it === 'string' || !it) ? '' : (it.color || '');
    const itemColorExtend = (it) => (typeof it === 'string' || !it) ? false : !!it.colorExtend;
    const dropdownOptionLabel = (opt) => (typeof opt === 'string' ? opt : (opt?.label ?? ''));
    const dropdownOptionColor = (opt) => (typeof opt === 'string' || !opt) ? '' : (opt.color || '');
    const hourlyHasRowSeparator = (sheet) => sheet?.layout === 'hourly' && sheet.rowSeparator !== false;
    const visibleItems  = (arr) => (arr || []).filter(it => !itemHidden(it));
    const visibleLabels = (arr) => visibleItems(arr).map(itemLabel);

    /* In-app dialog (confirm / alert) — replaces native browser dialogs to match
       the app UI. App registers a handler at mount; callers get a Promise. */
    let __dialogHandler = null;
    function openDialog(opts) {
      return new Promise((resolve) => {
        if (!__dialogHandler) {
          if (opts.kind === 'confirm') resolve(window.confirm(opts.message));
          else { window.alert(opts.message); resolve(); }
          return;
        }
        __dialogHandler({ ...opts, resolve });
      });
    }
    const confirmDialog = (opts) => openDialog({ ...opts, kind: 'confirm' });
    const alertDialog   = (opts) => openDialog({ ...opts, kind: 'alert' });
    const pad = (n) => String(n).padStart(2, '0');
    const isWeekend = (y, m, d) => { const wd = new Date(y, m, d).getDay(); return wd === 0 || wd === 6; };
    const colLetter = (n) => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
    const safeSheetName = (name) => (name || 'Sheet').replace(/[\\/\?\*\[\]:]/g, '').slice(0, 31) || 'Sheet';

    /* ============================================================
       Import existing .xlsx → sheet configs
       ============================================================ */
    const HOUR_LABEL_RE = /^(1[0-2]|[1-9])\s*(AM|PM)$/i;
    const SKIP_IMPORT_NAMES = /^(index|summary|changelog)$/i;

    function cellText(c) {
      if (c == null) return '';
      const v = c.value;
      if (v == null) return '';
      if (typeof v === 'object') {
        if (v.richText) return v.richText.map(p => p.text || '').join('').trim();
        if (v.text) return String(v.text).trim();
        if (v.result != null) return String(v.result).trim();
        return '';
      }
      return String(v).trim();
    }

    function detectHeaderRow(ws) {
      const maxScan = Math.min(ws.rowCount || 8, 8);
      for (let r = 1; r <= maxScan; r++) {
        const row = ws.getRow(r);
        const cells = [];
        const colCount = ws.columnCount || 30;
        for (let c = 1; c <= Math.max(colCount, 2); c++) {
          const t = cellText(row.getCell(c));
          if (t) cells.push({ col: c, text: t });
        }
        if (cells.length < 2) continue;
        const labels = cells.map(x => x.text);
        const lc = labels.map(x => x.toLowerCase());
        const hourCount = labels.filter(v => HOUR_LABEL_RE.test(v)).length;
        if (hourCount >= 3) return { row: r, labels, kind: 'hourly' };
        if (lc.includes('contains') && lc.includes('severity')) return { row: r, labels, kind: 'alarm' };
        if (lc.includes('date') && lc.includes('time')) return { row: r, labels, kind: 'flat' };
      }
      // Fallback: first non-empty row
      for (let r = 1; r <= maxScan; r++) {
        const row = ws.getRow(r);
        const labels = [];
        const colCount = ws.columnCount || 30;
        for (let c = 1; c <= Math.max(colCount, 2); c++) {
          const t = cellText(row.getCell(c));
          if (t) labels.push(t);
        }
        if (labels.length >= 2) return { row: r, labels, kind: 'flat' };
      }
      return null;
    }

    function inferHourlyRowSeparator(ws, headerRow, metricCount) {
      if (!metricCount) return true;
      const afterFirstBlock = headerRow + metricCount + 1;
      if (afterFirstBlock > (ws.rowCount || afterFirstBlock)) return true;
      const firstCell = cellText(ws.getRow(afterFirstBlock).getCell(1));
      const metricCell = cellText(ws.getRow(afterFirstBlock).getCell(2));
      return !(firstCell || metricCell);
    }

    async function parseXlsxToSheets(file) {
      const buf = await file.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const sheets = [];
      wb.eachSheet((ws) => {
        if (SKIP_IMPORT_NAMES.test(ws.name)) return;
        const hdr = detectHeaderRow(ws);
        if (!hdr) return;

        if (hdr.kind === 'hourly') {
          const hourIndices = hdr.labels
            .filter(v => HOUR_LABEL_RE.test(v))
            .map(v => HOURS_24.indexOf(v.toUpperCase().replace(/\s+/g, '')))
            .filter(i => i >= 0)
            .sort((a, b) => a - b);
          const hourStart = hourIndices[0] ?? 0;
          const hourEnd   = (hourIndices[hourIndices.length - 1] ?? 23) + 1;

          // Metrics live in column 2 below the header row; collect uniques in order.
          const seen = new Set();
          const metrics = [];
          const last = Math.min(ws.rowCount || (hdr.row + 400), hdr.row + 400);
          for (let r = hdr.row + 1; r <= last; r++) {
            const t = cellText(ws.getRow(r).getCell(2));
            if (!t) continue;
            if (HOUR_LABEL_RE.test(t)) continue;
            if (seen.has(t)) continue;
            seen.add(t);
            metrics.push(t);
            if (metrics.length >= 40) break;
          }
          sheets.push({
            id: newId(),
            name: ws.name,
            layout: 'hourly',
            active: true,
            metrics: metrics.length ? metrics : ['Metric 1'],
            rowSeparator: inferHourlyRowSeparator(ws, hdr.row, metrics.length || 1),
            hourStart,
            hourEnd,
            note: '',
          });
        } else {
          // flat or alarm — columns come from the detected header row
          sheets.push({
            id: newId(),
            name: ws.name,
            layout: hdr.kind === 'alarm' ? 'alarm' : 'flat',
            active: true,
            columns: hdr.labels,
            note: '',
          });
        }
      });
      return sheets;
    }

    const BMR_CLIENT_IMPORT_HEADER_KEYS = {
      clientname: 'name',
      client: 'name',
      carriername: 'name',
      carrier: 'name',
      allowbal: 'allowBalanceLabel',
      allowballabel: 'allowBalanceLabel',
      allowablebalance: 'allowBalanceLabel',
      allowablebalancelabel: 'allowBalanceLabel',
      accountmanager: 'accountManagerName',
      accountmanagername: 'accountManagerName',
      manager: 'accountManagerName',
      am: 'accountManagerName',
      clientcolor: 'color',
      clientcolour: 'color',
      color: 'color',
      colour: 'color',
      accountmanagercolor: 'accountManagerColor',
      accountmanagercolour: 'accountManagerColor',
      managercolor: 'accountManagerColor',
      managercolour: 'accountManagerColor',
      amcolor: 'accountManagerColor',
      amcolour: 'accountManagerColor',
      hidden: 'hidden',
      hide: 'hidden',
    };

    const BMR_RULE_IMPORT_HEADER_KEYS = {
      rulename: 'name',
      rule: 'name',
      name: 'name',
      clientname: 'clientName',
      client: 'clientName',
      carriername: 'clientName',
      carrier: 'clientName',
      condition: 'kind',
      operator: 'kind',
      kind: 'kind',
      comparison: 'kind',
      value: 'value',
      min: 'min',
      minimum: 'min',
      max: 'max',
      maximum: 'max',
      fill: 'color',
      fillcolor: 'color',
      background: 'color',
      backgroundcolor: 'color',
      color: 'color',
      font: 'fontColor',
      fontcolor: 'fontColor',
      textcolor: 'fontColor',
      bold: 'bold',
      italic: 'italic',
      underline: 'underline',
      enabled: 'enabled',
      active: 'enabled',
    };

    const BMR_RULE_KIND_LABELS = {
      gte: 'greater than or equal to',
      gt: 'greater than',
      lte: 'less than or equal to',
      lt: 'less than',
      eq: 'equal to',
      between: 'between',
    };

    function bmrImportKey(value) {
      return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    }

    function bmrImportColor(value) {
      const raw = String(value || '').trim().replace(/^#/, '');
      if (/^[0-9a-f]{6}$/i.test(raw)) return '#' + raw.toUpperCase();
      if (/^[0-9a-f]{8}$/i.test(raw)) return '#' + raw.slice(-6).toUpperCase();
      return '';
    }

    function bmrImportHidden(value) {
      const raw = bmrImportKey(value);
      if (['1', 'true', 'yes', 'y', 'hide', 'hidden'].includes(raw)) return true;
      if (['0', 'false', 'no', 'n', 'show', 'visible'].includes(raw)) return false;
      return undefined;
    }

    function bmrImportBoolean(value) {
      const raw = bmrImportKey(value);
      if (['1', 'true', 'yes', 'y', 'enabled', 'active', 'on', 'bold', 'italic', 'underline'].includes(raw)) return true;
      if (['0', 'false', 'no', 'n', 'disabled', 'inactive', 'off', 'none'].includes(raw)) return false;
      return undefined;
    }

    function bmrImportNumber(value) {
      const raw = String(value ?? '').replace(/\u00a0/g, ' ').trim().replace(/,/g, '');
      if (!raw) return undefined;
      const n = Number(raw);
      return Number.isFinite(n) ? n : undefined;
    }

    function bmrImportRuleKind(value) {
      const raw = bmrImportKey(value);
      const map = {
        gte: 'gte',
        ge: 'gte',
        greaterthanorequalto: 'gte',
        greaterthanorequals: 'gte',
        greaterorequal: 'gte',
        greaterorequalto: 'gte',
        atleastequal: 'gte',
        atleast: 'gte',
        gt: 'gt',
        greaterthan: 'gt',
        greater: 'gt',
        lte: 'lte',
        le: 'lte',
        lessthanorequalto: 'lte',
        lessthanorequals: 'lte',
        lessorequal: 'lte',
        lessorequalto: 'lte',
        atmost: 'lte',
        lt: 'lt',
        lessthan: 'lt',
        less: 'lt',
        eq: 'eq',
        equal: 'eq',
        equalto: 'eq',
        equals: 'eq',
        between: 'between',
      };
      const symbol = String(value || '').trim();
      if (symbol === '>=' || symbol === '≥') return 'gte';
      if (symbol === '>') return 'gt';
      if (symbol === '<=' || symbol === '≤') return 'lte';
      if (symbol === '<') return 'lt';
      if (symbol === '=') return 'eq';
      return map[raw] || '';
    }

    function bmrRuleKindLabel(kind) {
      return BMR_RULE_KIND_LABELS[kind] || kind || '';
    }

    function detectBmrClientHeader(ws) {
      const rowCount = Math.min(ws.rowCount || 10, 10);
      const colCount = Math.min(Math.max(ws.columnCount || 12, 12), 80);
      for (let r = 1; r <= rowCount; r++) {
        const cols = {};
        for (let c = 1; c <= colCount; c++) {
          const field = BMR_CLIENT_IMPORT_HEADER_KEYS[bmrImportKey(cellText(ws.getRow(r).getCell(c)))];
          if (field && !cols[field]) cols[field] = c;
        }
        if (cols.name) return { row: r, cols };
      }
      return null;
    }

    function detectBmrRuleHeader(ws, requireClient = false) {
      const rowCount = Math.min(ws.rowCount || 10, 10);
      const colCount = Math.min(Math.max(ws.columnCount || 12, 12), 80);
      for (let r = 1; r <= rowCount; r++) {
        const cols = {};
        for (let c = 1; c <= colCount; c++) {
          const field = BMR_RULE_IMPORT_HEADER_KEYS[bmrImportKey(cellText(ws.getRow(r).getCell(c)))];
          if (field && !cols[field]) cols[field] = c;
        }
        if (cols.kind && (!requireClient || cols.clientName)) return { row: r, cols };
      }
      return null;
    }

    function parseBmrRuleRows(ws, header, type) {
      const rows = [];
      const lastRow = Math.min(ws.rowCount || header.row, header.row + 5000);
      for (let r = header.row + 1; r <= lastRow; r++) {
        const read = (field) => header.cols[field] ? cellText(ws.getRow(r).getCell(header.cols[field])) : undefined;
        const name = read('name');
        const kind = header.cols.kind ? bmrImportRuleKind(read('kind')) : '';
        const clientName = read('clientName');
        const hasAny = ['name','clientName','kind','value','min','max','color','fontColor','bold','italic','underline','enabled']
          .some(field => header.cols[field] && String(read(field) ?? '').trim());
        if (!hasAny) continue;
        if (!kind) continue;
        if (type === 'target' && !String(clientName || '').trim()) continue;
        const imported = { type, kind };
        if (String(name || '').trim()) imported.name = String(name).trim();
        if (type === 'target') imported.clientName = String(clientName || '').trim();
        if (header.cols.value) {
          const value = bmrImportNumber(read('value'));
          if (value !== undefined) imported.value = value;
        }
        if (header.cols.min) {
          const min = bmrImportNumber(read('min'));
          if (min !== undefined) imported.min = min;
        }
        if (header.cols.max) {
          const max = bmrImportNumber(read('max'));
          if (max !== undefined) imported.max = max;
        }
        if (header.cols.color) {
          const color = bmrImportColor(read('color'));
          if (color) imported.color = color;
        }
        if (header.cols.fontColor) {
          const fontColor = bmrImportColor(read('fontColor'));
          if (fontColor) imported.fontColor = fontColor;
        }
        ['bold', 'italic', 'underline', 'enabled'].forEach((field) => {
          if (!header.cols[field]) return;
          const value = bmrImportBoolean(read(field));
          if (value !== undefined) imported[field] = value;
        });
        rows.push(imported);
      }
      return rows;
    }

    async function parseBmrImportXlsx(file) {
      const buf = await file.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const rows = [];
      const targetRules = [];
      const usageRules = [];
      wb.eachSheet((ws) => {
        const targetHeader = detectBmrRuleHeader(ws, true);
        if (targetHeader) {
          targetRules.push(...parseBmrRuleRows(ws, targetHeader, 'target'));
          return;
        }
        const sheetKey = bmrImportKey(ws.name);
        const usageHeader = detectBmrRuleHeader(ws, false);
        if (usageHeader && (sheetKey.includes('usage') || sheetKey.includes('30min') || !usageHeader.cols.clientName)) {
          usageRules.push(...parseBmrRuleRows(ws, usageHeader, 'usage'));
          return;
        }
        const header = detectBmrClientHeader(ws);
        if (!header) return;
        const lastRow = Math.min(ws.rowCount || header.row, header.row + 5000);
        for (let r = header.row + 1; r <= lastRow; r++) {
          const read = (field) => header.cols[field] ? cellText(ws.getRow(r).getCell(header.cols[field])) : undefined;
          const name = read('name');
          if (!name) continue;
          const imported = { name };
          if (header.cols.allowBalanceLabel) imported.allowBalanceLabel = read('allowBalanceLabel') || '';
          if (header.cols.accountManagerName) imported.accountManagerName = read('accountManagerName') || '';
          if (header.cols.color) {
            const color = bmrImportColor(read('color'));
            if (color) imported.color = color;
          }
          if (header.cols.accountManagerColor) {
            const accountManagerColor = bmrImportColor(read('accountManagerColor'));
            if (accountManagerColor) imported.accountManagerColor = accountManagerColor;
          }
          if (header.cols.hidden) {
            const hidden = bmrImportHidden(read('hidden'));
            if (hidden !== undefined) imported.hidden = hidden;
          }
          rows.push(imported);
        }
      });
      return { clients: rows, targetRules, usageRules };
    }

    async function parseBmrClientsXlsx(file) {
      const parsed = await parseBmrImportXlsx(file);
      return parsed.clients;
    }

    function mergeBmrImportedClients(bmr, importedRows) {
      const normalizeName = (value) => String(value || '').trim().toLowerCase();
      const compactRowsByName = new Map();
      importedRows.forEach((row) => {
        const key = normalizeName(row.name);
        if (!key) return;
        compactRowsByName.set(key, { ...(compactRowsByName.get(key) || {}), ...row });
      });
      const rows = Array.from(compactRowsByName.values());
      const accountManagers = [...(bmr.accountManagers || [])];
      const managerByName = new Map(accountManagers.map(manager => [normalizeName(manager.name), manager]));
      const ensureManager = (name, color) => {
        const key = normalizeName(name);
        if (!key) return null;
        let manager = managerByName.get(key);
        if (!manager) {
          manager = { id: bmrId('am'), name: String(name).trim(), color: color || '' };
          accountManagers.push(manager);
          managerByName.set(key, manager);
        } else if (color && manager.color !== color) {
          manager = { ...manager, color };
          const index = accountManagers.findIndex(item => item.id === manager.id);
          accountManagers[index] = manager;
          managerByName.set(key, manager);
        }
        return manager;
      };

      const clients = [...(bmr.clients || [])];
      const clientIndexByName = new Map(clients.map((client, index) => [normalizeName(client.name), index]));
      let added = 0;
      let updated = 0;

      rows.forEach((row) => {
        const key = normalizeName(row.name);
        const accountManager = row.accountManagerName ? ensureManager(row.accountManagerName, row.accountManagerColor) : null;
        const patch = {};
        if (Object.prototype.hasOwnProperty.call(row, 'allowBalanceLabel')) patch.allowBalanceLabel = row.allowBalanceLabel;
        if (row.color) patch.color = row.color;
        if (Object.prototype.hasOwnProperty.call(row, 'hidden')) patch.hidden = row.hidden;
        if (accountManager) patch.accountManagerId = accountManager.id;

        if (clientIndexByName.has(key)) {
          const index = clientIndexByName.get(key);
          clients[index] = { ...clients[index], ...patch, name: row.name };
          updated += 1;
        } else {
          clients.push({
            id: bmrId(),
            name: row.name,
            color: row.color || '',
            allowBalanceLabel: Object.prototype.hasOwnProperty.call(row, 'allowBalanceLabel') ? row.allowBalanceLabel : '',
            accountManagerId: accountManager?.id || '',
            hidden: row.hidden === true,
          });
          clientIndexByName.set(key, clients.length - 1);
          added += 1;
        }
      });

      return {
        bmr: { ...bmr, clients, accountManagers },
        stats: { rows: rows.length, added, updated },
      };
    }

    function bmrImportedRulePatch(row) {
      const patch = {};
      if (row.kind) patch.kind = row.kind;
      ['value', 'min', 'max'].forEach((field) => {
        if (Object.prototype.hasOwnProperty.call(row, field)) patch[field] = row[field];
      });
      ['color', 'fontColor'].forEach((field) => {
        if (row[field]) patch[field] = row[field];
      });
      ['bold', 'italic', 'underline', 'enabled'].forEach((field) => {
        if (Object.prototype.hasOwnProperty.call(row, field)) patch[field] = row[field];
      });
      return patch;
    }

    function mergeBmrImportedRules(bmr, importedTargetRules = [], importedUsageRules = []) {
      const normalizeName = (value) => String(value || '').trim().toLowerCase();
      const clients = bmr.clients || [];
      const clientByName = new Map(clients.map(client => [normalizeName(client.name), client]));
      const clientNameById = new Map(clients.map(client => [client.id, client.name || '']));
      const targetRules = [...(bmr.targetRules || [])];
      const usageRules = [...(bmr.usageRules || [])];
      const targetKey = (clientName) => normalizeName(clientName);
      const usageKey = (name) => normalizeName(name);
      const targetIndex = new Map();
      targetRules.forEach((rule, index) => {
        const key = targetKey(clientNameById.get(rule.clientId) || '');
        if (key && !targetIndex.has(key)) targetIndex.set(key, index);
      });
      const usageIndex = new Map(usageRules.map((rule, index) => [usageKey(rule.name), index]));
      const stats = { targetAdded: 0, targetUpdated: 0, targetSkipped: 0, usageAdded: 0, usageUpdated: 0 };

      importedTargetRules.forEach((row, index) => {
        const client = clientByName.get(normalizeName(row.clientName));
        if (!client) {
          stats.targetSkipped += 1;
          return;
        }
        const importedName = row.name || '';
        const name = importedName || `Client rule ${index + 1}`;
        const patch = bmrImportedRulePatch(row);
        const key = targetKey(client.name);
        if (targetIndex.has(key)) {
          const existingIndex = targetIndex.get(key);
          targetRules[existingIndex] = {
            ...targetRules[existingIndex],
            ...patch,
            ...(importedName ? { name: importedName } : {}),
            clientId: client.id,
          };
          stats.targetUpdated += 1;
        } else {
          targetRules.push({
            id: bmrId('tr'),
            name,
            clientId: client.id,
            kind: 'gte',
            value: 5,
            color: '#EF4444',
            fontColor: '#FFFFFF',
            bold: false,
            italic: false,
            underline: false,
            enabled: true,
            ...patch,
          });
          targetIndex.set(key, targetRules.length - 1);
          stats.targetAdded += 1;
        }
      });

      importedUsageRules.forEach((row, index) => {
        const name = row.name || `Usage rule ${index + 1}`;
        const patch = bmrImportedRulePatch(row);
        const key = usageKey(name);
        if (usageIndex.has(key)) {
          const existingIndex = usageIndex.get(key);
          usageRules[existingIndex] = { ...usageRules[existingIndex], ...patch, name };
          stats.usageUpdated += 1;
        } else {
          usageRules.push({
            id: bmrId('ur'),
            name,
            kind: 'between',
            min: -49,
            max: -0.001,
            color: '#F59E0B',
            fontColor: '#1F2937',
            bold: false,
            italic: false,
            underline: false,
            enabled: true,
            ...patch,
          });
          usageIndex.set(key, usageRules.length - 1);
          stats.usageAdded += 1;
        }
      });

      return {
        bmr: { ...bmr, targetRules, usageRules },
        stats,
      };
    }

    const BMR_CLIENT_IMPORT_COLUMNS = [
      { header: 'Client Name', key: 'name', width: 24 },
      { header: 'Allow Bal label', key: 'allowBalanceLabel', width: 32 },
      { header: 'Account Manager', key: 'accountManagerName', width: 24 },
      { header: 'Client color', key: 'color', width: 16 },
      { header: 'Account Manager color', key: 'accountManagerColor', width: 24 },
      { header: 'Hidden', key: 'hidden', width: 12 },
    ];

    const BMR_TARGET_RULE_IMPORT_COLUMNS = [
      { header: 'Rule Name', key: 'name', width: 30 },
      { header: 'Client Name', key: 'clientName', width: 28 },
      { header: 'Condition', key: 'kind', width: 26 },
      { header: 'Value', key: 'value', width: 14 },
      { header: 'Min', key: 'min', width: 14 },
      { header: 'Max', key: 'max', width: 14 },
      { header: 'Fill color', key: 'color', width: 16 },
      { header: 'Font color', key: 'fontColor', width: 16 },
      { header: 'Bold', key: 'bold', width: 10 },
      { header: 'Italic', key: 'italic', width: 10 },
      { header: 'Underline', key: 'underline', width: 12 },
      { header: 'Enabled', key: 'enabled', width: 12 },
    ];

    const BMR_USAGE_RULE_IMPORT_COLUMNS = BMR_TARGET_RULE_IMPORT_COLUMNS.filter(col => col.key !== 'clientName');

    function styleBmrImportSheet(ws) {
      ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
      ws.getRow(1).alignment = { vertical: 'middle' };
      ws.views = [{ state: 'frozen', ySplit: 1 }];
    }

    async function downloadBmrClientImportWorkbook(sheetName, rows, filename, options = {}) {
      const wb = new ExcelJS.Workbook();
      wb.creator = 'BMR Generator';
      wb.created = new Date();
      const ws = wb.addWorksheet(sheetName);
      ws.columns = BMR_CLIENT_IMPORT_COLUMNS;
      ws.addRows(rows);
      styleBmrImportSheet(ws);
      const targetRules = options.targetRules || [];
      const usageRules = options.usageRules || [];
      if (targetRules.length || options.includeRuleSheets) {
        const targetWs = wb.addWorksheet('Target Rules');
        targetWs.columns = BMR_TARGET_RULE_IMPORT_COLUMNS;
        targetWs.addRows(targetRules);
        styleBmrImportSheet(targetWs);
      }
      if (usageRules.length || options.includeRuleSheets) {
        const usageWs = wb.addWorksheet('30-min Usage Rules');
        usageWs.columns = BMR_USAGE_RULE_IMPORT_COLUMNS;
        usageWs.addRows(usageRules);
        styleBmrImportSheet(usageWs);
      }
      const buf = await wb.xlsx.writeBuffer();
      saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);
    }

    async function downloadBmrClientImportSample() {
      await downloadBmrClientImportWorkbook('Client Import Sample', [
        { name: 'SAMPLE CLIENT A', allowBalanceLabel: 'Block once debit', accountManagerName: 'Jason Cruz', color: '', accountManagerColor: '#60A5FA', hidden: 'No' },
        { name: 'SAMPLE CLIENT B', allowBalanceLabel: '$5 debit poc', accountManagerName: 'Jason Cruz', color: '#FACC15', accountManagerColor: '#60A5FA', hidden: 'No' },
      ], 'BMR_Client_Import_Sample.xlsx', {
        includeRuleSheets: true,
        targetRules: [
          { name: 'Reached allowable balance', clientName: 'SAMPLE CLIENT B', kind: 'greater than or equal to', value: 5, color: '#EF4444', fontColor: '#FFFFFF', bold: 'Yes', italic: 'No', underline: 'No', enabled: 'Yes' },
        ],
        usageRules: [
          { name: 'Light usage (-49 to -0.001)', kind: 'between', min: -49, max: -0.001, color: '#F59E0B', fontColor: '#1F2937', bold: 'Yes', italic: 'No', underline: 'No', enabled: 'Yes' },
        ],
      });
    }

    async function downloadBmrClientData(bmr) {
      const accountManagersById = bmrAccountManagersById(bmr);
      const clientsById = new Map((bmr.clients || []).map(client => [client.id, client]));
      const rows = (bmr.clients || []).map((client) => {
        const manager = accountManagersById.get(client.accountManagerId);
        return {
          name: client.name || '',
          allowBalanceLabel: client.allowBalanceLabel || '',
          accountManagerName: manager?.name || '',
          color: client.color || '',
          accountManagerColor: manager?.color || '',
          hidden: client.hidden ? 'Yes' : 'No',
        };
      });
      const targetRules = (bmr.targetRules || []).map(rule => ({
        name: rule.name || '',
        clientName: clientsById.get(rule.clientId)?.name || '',
        kind: bmrRuleKindLabel(rule.kind),
        value: rule.kind === 'between' ? '' : rule.value,
        min: rule.kind === 'between' ? rule.min : '',
        max: rule.kind === 'between' ? rule.max : '',
        color: rule.color || '',
        fontColor: rule.fontColor || '',
        bold: rule.bold ? 'Yes' : 'No',
        italic: rule.italic ? 'Yes' : 'No',
        underline: rule.underline ? 'Yes' : 'No',
        enabled: rule.enabled !== false ? 'Yes' : 'No',
      }));
      const usageRules = (bmr.usageRules || []).map(rule => ({
        name: rule.name || '',
        kind: bmrRuleKindLabel(rule.kind),
        value: rule.kind === 'between' ? '' : rule.value,
        min: rule.kind === 'between' ? rule.min : '',
        max: rule.kind === 'between' ? rule.max : '',
        color: rule.color || '',
        fontColor: rule.fontColor || '',
        bold: rule.bold ? 'Yes' : 'No',
        italic: rule.italic ? 'Yes' : 'No',
        underline: rule.underline ? 'Yes' : 'No',
        enabled: rule.enabled !== false ? 'Yes' : 'No',
      }));
      await downloadBmrClientImportWorkbook('Current Client Data', rows, `BMR_Client_Data_${bmrTodayString()}.xlsx`, {
        includeRuleSheets: true,
        targetRules,
        usageRules,
      });
    }

    /* ============================================================
       Excel Generation (ExcelJS — Google Sheets compatible)
       ============================================================ */
    // Builds the SIP/FCS workbook in memory. Returns { blob, filename } so both
    // the download button and the Google Sheets sync can reuse one build.
    // Snapshot each finished worksheet's used extent (last row / column that
    // actually holds a value) so the Google Sheets sync can shrink the converted
    // sheet's grid back to exactly this range. Drive's xlsx→Sheets conversion
    // pads every tab out to its 1000-row default, leaving empty rows below the
    // totals row; this lets us trim them off so the last row is the totals row.
    function collectSheetGridDims(wb) {
      return wb.worksheets.map(ws => ({
        title: ws.name,
        rows: ws.rowCount || 0,
        cols: ws.columnCount || 0,
      }));
    }

    async function buildSipFcsBlob(state) {
      const wb = new ExcelJS.Workbook();
      wb.creator = 'Monthly Report Generator';
      wb.created = new Date();
      const { year, month, sheets, rules, includeIndex, changelog } = state;
      const activeSheets = sheets.filter(s => s.active);

      if (includeIndex) {
        const idx = wb.addWorksheet('INDEX');
        idx.columns = [
          { header: 'Sheet',             key: 'sheet',  width: 26 },
          { header: 'Layout',            key: 'layout', width: 16 },
          { header: 'Metrics / Columns', key: 'items',  width: 80 },
        ];
        idx.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        idx.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
        activeSheets.forEach(s => {
          idx.addRow({
            sheet: s.name,
            layout: LAYOUT_LABELS[s.layout],
            items: visibleLabels(s.layout === 'hourly' ? s.metrics : s.columns).join(', '),
          });
        });
        idx.addRow({});
        idx.addRow({ sheet: 'Period',     layout: `${MONTHS[month]} ${year}` });
        idx.addRow({ sheet: 'Days',       layout: String(daysInMonth(year, month)) });
        idx.addRow({ sheet: 'Generated',  layout: new Date().toLocaleString() });
        idx.views = [{ state: 'frozen', ySplit: 1 }];
      }

      for (const s of activeSheets) {
        const ws = wb.addWorksheet(safeSheetName(s.name));
        ws.properties.defaultRowHeight = 18;
        if (s.layout === 'hourly')      buildHourlySheet(ws, s, year, month, rules);
        else if (s.layout === 'flat')   buildFlatSheet(ws, s, year, month, rules);
        else if (s.layout === 'alarm')  buildAlarmSheet(ws, s, rules);
      }

      if (changelog && changelog.trim()) {
        const cl = wb.addWorksheet('CHANGELOG');
        cl.getCell('A1').value = `${MONTHS[month]} ${year} — Notes`;
        cl.getCell('A1').font = { bold: true, size: 14 };
        cl.getCell('A3').value = changelog;
        cl.getCell('A3').alignment = { wrapText: true, vertical: 'top' };
        cl.getColumn(1).width = 100;
      }

      const buf = await wb.xlsx.writeBuffer();
      const filename = `${MONTHS[month]}_${year}_SIP_FCS_Hourly_Record.xlsx`;
      return { blob: new Blob([buf], { type: XLSX_MIME }), filename, sheets: collectSheetGridDims(wb) };
    }

    async function generateWorkbook(state) {
      const { blob, filename } = await buildSipFcsBlob(state);
      saveAs(blob, filename);
      return filename;
    }

    const SIP_LAYOUT_RULE_SCOPES = [
      { value: '__layout_alarm__',  label: 'Global rules for all Alarm Log',    layout: 'alarm' },
      { value: '__layout_hourly__', label: 'Global rules for all Hourly Block', layout: 'hourly' },
      { value: '__layout_hourly_separator__', label: 'Global rules for all Hourly Block with row separator', layout: 'hourly', rowSeparator: true },
      { value: '__layout_hourly_no_separator__', label: 'Global rules for all Hourly Block with no row separator', layout: 'hourly', rowSeparator: false },
      { value: '__layout_flat__',   label: 'Global rules for all Flat Table',   layout: 'flat' },
    ];

    function sipLayoutScopeMatches(layoutScope, sheet) {
      if (!layoutScope || !sheet || sheet.layout !== layoutScope.layout) return false;
      if (layoutScope.layout === 'hourly' && typeof layoutScope.rowSeparator === 'boolean') {
        return hourlyHasRowSeparator(sheet) === layoutScope.rowSeparator;
      }
      return true;
    }

    function sipRuleScopeMatches(rule, scope, sheet) {
      if (scope === '__view_all__') return true;
      if (scope === '__all__') return !rule.targetSheet || rule.targetSheet === '__all__';
      const layoutScope = SIP_LAYOUT_RULE_SCOPES.find(item => item.value === scope);
      if (layoutScope) return rule.targetSheet === scope;
      return rule.targetSheet === scope;
    }

    function rulesForSheet(allRules, sheet) {
      return (allRules || []).filter(r => {
        const scope = r.targetSheet || '__all__';
        if (scope === '__all__') return true;
        const layoutScope = SIP_LAYOUT_RULE_SCOPES.find(item => item.value === scope);
        if (layoutScope) return sipLayoutScopeMatches(layoutScope, sheet);
        return scope === sheet.id;
      });
    }

    function sipRuleStyle(rule, fallbackColor = 'FF166534') {
      return {
        fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: rule.color || fallbackColor } },
        font: { bold: !!rule.bold, color: { argb: rule.fontColor || 'FFFFFFFF' } },
      };
    }

    function applyTabularSheetRules(ws, sheet, allRules, dataStartRow, dataEndRow, totalCols) {
      if (dataEndRow < dataStartRow || totalCols < 1) return;
      const rules = rulesForSheet(allRules, sheet);
      if (!rules.length) return;
      const firstCol = colLetter(1);
      const lastCol = colLetter(totalCols);
      const dataRange = `${firstCol}${dataStartRow}:${lastCol}${dataEndRow}`;
      rules.forEach((r, idx) => {
        if (!r.enabled) return;
        const priority = 10 + idx;
        const style = sipRuleStyle(r);
        if (r.type === 'above') {
          const t = Number(r.threshold) || 0;
          ws.addConditionalFormatting({
            ref: dataRange,
            rules: [{ type: 'expression', formulae: [`AND(ISNUMBER(${firstCol}${dataStartRow}),${firstCol}${dataStartRow}>${t})`], style, priority }],
          });
        } else if (r.type === 'below') {
          const t = Number(r.threshold) || 0;
          ws.addConditionalFormatting({
            ref: dataRange,
            rules: [{ type: 'expression', formulae: [`AND(ISNUMBER(${firstCol}${dataStartRow}),${firstCol}${dataStartRow}<${t})`], style, priority }],
          });
        } else if (r.type === 'blank') {
          ws.addConditionalFormatting({
            ref: dataRange,
            rules: [{ type: 'expression', formulae: [`ISBLANK(${firstCol}${dataStartRow})`], style: sipRuleStyle(r, 'FF374151'), priority }],
          });
        } else if (r.type === 'zero') {
          ws.addConditionalFormatting({
            ref: dataRange,
            rules: [{ type: 'expression', formulae: [`${firstCol}${dataStartRow}=0`], style, priority }],
          });
        }
      });
    }

    function buildHourlySheet(ws, s, year, month, allRules) {
      const hourStart = s.hourStart ?? 0;
      const hourEnd   = s.hourEnd ?? 24;
      const hours = HOURS_24.slice(hourStart, hourEnd);
      const totalCols = 2 + hours.length;
      const rules = rulesForSheet(allRules, s);
      const hasRowSeparator = hourlyHasRowSeparator(s);

      let cur = 1;

      // Header note
      if (s.note && s.note.trim()) {
        applyNoteStyle(ws, cur, totalCols, s);
        cur += 2;
      }

      // Header row
      const headerRow = ws.getRow(cur);
      headerRow.values = [s.dateHeader || 'DATE', s.metricHeader ?? '', ...hours];
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top:    { style: 'thin', color: { argb: 'FF374151' } },
          bottom: { style: 'thin', color: { argb: 'FF374151' } },
        };
      });
      ws.getRow(cur).height = 22;
      const headerRowNum = cur;
      cur += 1;

      const days = daysInMonth(year, month);
      const visibleMetrics = visibleItems(s.metrics);
      const metricCount = visibleMetrics.length;
      const dataStartRow = cur;
      const weekendRule = rules.find(r => r.type === 'weekend' && r.enabled);
      const separatorRule = hasRowSeparator ? rules.find(r => r.type === 'separator' && r.enabled) : null;

      const borderOn = !!s.bordersOnData;
      const borderStyle = s.borderStyle || 'thin';
      const borderColor = s.borderColor || 'FF374151';
      const borderSide = borderOn ? { style: borderStyle, color: { argb: borderColor } } : null;

      for (let d = 1; d <= days; d++) {
        const blockStart = cur;
        const dateObj = new Date(Date.UTC(year, month, d));
        for (let mi = 0; mi < metricCount; mi++) {
          const row = ws.getRow(cur);
          if (mi === 0) {
            row.getCell(1).value = dateObj;
            row.getCell(1).numFmt = s.dateFormat || 'yyyy-mm-dd';
          }
          const metric = visibleMetrics[mi];
          row.getCell(2).value = itemLabel(metric);
          row.getCell(2).alignment = { horizontal: 'left', vertical: 'middle' };
          if (itemColor(metric)) {
            const metricFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argbHex(itemColor(metric)) } };
            row.getCell(2).fill = metricFill;
            if (itemColorExtend(metric)) {
              for (let cIdx = 3; cIdx <= totalCols; cIdx++) row.getCell(cIdx).fill = metricFill;
            }
          }
          for (let h = 0; h < hours.length; h++) {
            row.getCell(3 + h).value = null;
            row.getCell(3 + h).alignment = { horizontal: 'center' };
          }
          if (borderSide) {
            for (let cIdx = 1; cIdx <= totalCols; cIdx++) {
              ws.getCell(cur, cIdx).border = {
                top: borderSide, left: borderSide, bottom: borderSide, right: borderSide,
              };
            }
          }
          cur += 1;
        }
        if (metricCount > 1) ws.mergeCells(blockStart, 1, blockStart + metricCount - 1, 1);
        const dateCell = ws.getCell(blockStart, 1);
        dateCell.font = { bold: true };
        dateCell.alignment = { vertical: 'middle', horizontal: 'center' };

        if (weekendRule && isWeekend(year, month, d)) {
          const fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: weekendRule.color || 'FF3F2A2A' } };
          for (let rIdx = blockStart; rIdx < blockStart + metricCount; rIdx++) {
            for (let cIdx = 1; cIdx <= totalCols; cIdx++) ws.getCell(rIdx, cIdx).fill = fill;
          }
        }
        if (hasRowSeparator && d < days) {
          if (separatorRule) {
            const fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: separatorRule.color || 'FFE5E7EB' } };
            for (let cIdx = 1; cIdx <= totalCols; cIdx++) ws.getCell(cur, cIdx).fill = fill;
          }
          cur += 1; // blank row separator between day blocks
        }
      }

      ws.getColumn(1).width = 14;
      ws.getColumn(2).width = 34;
      for (let i = 3; i <= totalCols; i++) ws.getColumn(i).width = 8;

      ws.views = [{ state: 'frozen', xSplit: 2, ySplit: headerRowNum }];

      // Conditional formatting (Google-Sheets-compatible formula rules)
      const dataEndRow = cur - 1;
      if (dataEndRow >= dataStartRow) {
        const firstCol = colLetter(3);
        const lastCol  = colLetter(totalCols);
        const dataRange = `${firstCol}${dataStartRow}:${lastCol}${dataEndRow}`;

        rules.forEach((r, idx) => {
          if (!r.enabled) return;
          const style = {
            fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: r.color || 'FF166534' } },
            font: { bold: !!r.bold, color: { argb: r.fontColor || 'FFFFFFFF' } },
          };
          const priority = 10 + idx;
          if (r.type === 'maxInRow') {
            ws.addConditionalFormatting({
              ref: dataRange,
              rules: [{
                type: 'expression',
                formulae: [`AND(ISNUMBER(${firstCol}${dataStartRow}),${firstCol}${dataStartRow}=MAX($${firstCol}${dataStartRow}:$${lastCol}${dataStartRow}))`],
                style, priority,
              }],
            });
          } else if (r.type === 'maxInCol') {
            ws.addConditionalFormatting({
              ref: dataRange,
              rules: [{
                type: 'expression',
                formulae: [`AND(ISNUMBER(${firstCol}${dataStartRow}),${firstCol}${dataStartRow}=MAX(${firstCol}$${dataStartRow}:${firstCol}$${dataEndRow}))`],
                style, priority,
              }],
            });
          } else if (r.type === 'above') {
            const t = Number(r.threshold) || 0;
            ws.addConditionalFormatting({
              ref: dataRange,
              rules: [{
                type: 'expression',
                formulae: [`AND(ISNUMBER(${firstCol}${dataStartRow}),${firstCol}${dataStartRow}>${t})`],
                style, priority,
              }],
            });
          } else if (r.type === 'below') {
            const t = Number(r.threshold) || 0;
            ws.addConditionalFormatting({
              ref: dataRange,
              rules: [{
                type: 'expression',
                formulae: [`AND(ISNUMBER(${firstCol}${dataStartRow}),${firstCol}${dataStartRow}<${t})`],
                style, priority,
              }],
            });
          } else if (r.type === 'blank') {
            ws.addConditionalFormatting({
              ref: dataRange,
              rules: [{
                type: 'expression',
                formulae: [`ISBLANK(${firstCol}${dataStartRow})`],
                style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: r.color || 'FF374151' } } },
                priority,
              }],
            });
          } else if (r.type === 'zero') {
            ws.addConditionalFormatting({
              ref: dataRange,
              rules: [{
                type: 'expression',
                formulae: [`${firstCol}${dataStartRow}=0`],
                style, priority,
              }],
            });
          }
        });

        // Static: bold specific metric rows
        rules.forEach(r => {
          if (!r.enabled || r.type !== 'boldMetric') return;
          const targets = (r.metricMatch || '').split(',').map(x => x.trim().toUpperCase()).filter(Boolean);
          if (!targets.length) return;
          for (let row = dataStartRow; row <= dataEndRow; row++) {
            const mv = ws.getCell(row, 2).value;
            if (mv && targets.some(t => String(mv).toUpperCase().includes(t))) {
              for (let c = 1; c <= totalCols; c++) {
                const cell = ws.getCell(row, c);
                cell.font = { ...(cell.font || {}), bold: true };
              }
            }
          }
        });
      }
    }

    function buildFlatSheet(ws, s, year, month, allRules) {
      const cols = visibleLabels(s.columns);
      let cur = 1;
      if (s.note && s.note.trim()) {
        applyNoteStyle(ws, cur, cols.length, s);
        cur += 2;
      }
      const hdr = ws.getRow(cur);
      hdr.values = cols;
      hdr.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
        cell.alignment = { horizontal: 'center' };
      });
      hdr.height = 22;
      cols.forEach((_, i) => { ws.getColumn(i + 1).width = 14; });

      // Prefill first data row: Date column → first of month, Time column → 12:00 AM.
      const dateFmt = s.dateFormat || 'mmm d, yyyy';
      const firstOfMonth = new Date(Date.UTC(year, month, 1, 0, 0, 0));
      const firstRow = ws.getRow(cur + 1);
      cols.forEach((c, i) => {
        const lc = String(c).toLowerCase().trim();
        if (lc === 'date') {
          firstRow.getCell(i + 1).value = firstOfMonth;
          firstRow.getCell(i + 1).numFmt = dateFmt;
        } else if (lc === 'time') {
          firstRow.getCell(i + 1).value = firstOfMonth;
          firstRow.getCell(i + 1).numFmt = 'h:mm AM/PM';
        }
      });
      applyTabularSheetRules(ws, s, allRules, cur + 1, cur + 500, cols.length);

      ws.views = [{ state: 'frozen', ySplit: cur }];
    }

    function buildAlarmSheet(ws, s, allRules) {
      const visibleCols = visibleItems(s.columns || []);
      const labels = visibleCols.map(itemLabel);
      let cur = 1;
      if (s.note && s.note.trim()) {
        applyNoteStyle(ws, cur, labels.length, s);
        cur += 2;
      }
      const hdr = ws.getRow(cur);
      hdr.values = labels;
      hdr.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
        cell.alignment = { horizontal: 'center' };
      });
      hdr.height = 22;
      labels.forEach((_, i) => { ws.getColumn(i + 1).width = 18; });

      // Apply list-based data validation (dropdowns) for any column with dropdown options.
      // Commas in option text are replaced with spaces — inline list formulae are comma-delimited.
      const VALIDATION_ROWS = 500;
      const dataStartRow = cur + 1;
      const dataEndRow   = cur + VALIDATION_ROWS;
      visibleCols.forEach((col, i) => {
        const opts = itemDropdown(col);
        if (!opts.length) return;
        const safe = opts
          .map(o => ({
            label: String(dropdownOptionLabel(o)).replace(/,/g, ' ').trim(),
            color: dropdownOptionColor(o),
          }))
          .filter(o => o.label);
        if (!safe.length) return;
        const formula = `"${safe.map(o => o.label).join(',')}"`;
        const colRef = colLetter(i + 1);
        const validationRange = `${colRef}${dataStartRow}:${colRef}${dataEndRow}`;
        for (let r = dataStartRow; r <= dataEndRow; r++) {
          ws.getCell(r, i + 1).dataValidation = {
            type: 'list',
            allowBlank: true,
            formulae: [formula],
            showErrorMessage: false,
          };
        }
        safe.forEach((opt, optIdx) => {
          const color = argbHex(opt.color, '');
          if (!color) return;
          const text = opt.label.replace(/"/g, '""');
          ws.addConditionalFormatting({
            ref: validationRange,
            rules: [{
              type: 'expression',
              formulae: [`${colRef}${dataStartRow}="${text}"`],
              style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: color } } },
              priority: 200 + (i * 50) + optIdx,
            }],
          });
        });
      });
      applyTabularSheetRules(ws, s, allRules, dataStartRow, dataEndRow, labels.length);

      ws.views = [{ state: 'frozen', ySplit: cur }];
    }

    /* ============================================================
       BMR module — Balance Day & Night daily generator
       ============================================================ */
    const BMR_BLOCK_COLS = 12;
    const BMR_BLOCK_HEADERS = ['CarrierName ','Credit Limit ','Critical Balance ','Billing Cycle ','IsPrepay ','Inv Upto ','Prev Balance ','Pay&Adj ','Current Usage ','Amnt Receivable ','30mins usage','Total usage'];
    const BMR_BLOCK_HEADER_OPTIONS = BMR_BLOCK_HEADERS.map((label, offset) => ({ offset, label: label.trim() }));
    const BMR_AMOUNT_OFFSET = 9;
    const BMR_USAGE_OFFSET  = 10;
    const BMR_TOTAL_OFFSET  = 11;
    const BMR_DECIMAL_OFFSETS = [1, 2, 5, 6, 7, 8, BMR_AMOUNT_OFFSET, BMR_USAGE_OFFSET, BMR_TOTAL_OFFSET];
    // Always hide these per-block columns in the generated XLSX (CarrierName, Amnt Receivable, 30mins usage, Total usage remain visible).
    const BMR_FORCE_HIDDEN_OFFSETS = [1, 2, 3, 4, 5, 6, 7, 8];
    const BMR_SLOT_COUNT = 25;
    const BMR_EXTRA_ROWS = 10;
    const BMR_SMS_BLOCK_COLS = 9;
    const BMR_SMS_USAGE_OFFSET = 7;
    const BMR_SMS_TOTAL_OFFSET = 8;
    const BMR_SMS_MARKETS = {
      retail: {
        label: 'Retail',
        code: 'RES',
        clientsKey: 'retailClients',
        targetOffset: 5,
        balanceOffset: 5,
        overdraftOffset: 6,
        headers: ['Client', 'Name', '', '', '', 'BALANCE', 'Over Draft', '30mins usage', 'Total usage'],
      },
      wholesale: {
        label: 'Wholesale',
        code: 'WHS',
        clientsKey: 'wholesaleClients',
        targetOffset: 6,
        balanceOffset: 6,
        overdraftOffset: 5,
        headers: ['Client', 'Name', '', '', '', 'OD', 'BALANCE', '30mins usage', 'Total usage'],
      },
    };

    function bmrSlotsForShift(shift) {
      const slots = [];
      const startHour = shift === 'night' ? 19 : 7;
      for (let i = 0; i < BMR_SLOT_COUNT; i++) {
        const totalMin = startHour * 60 + i * 30;
        const h24 = Math.floor(totalMin / 60) % 24;
        const m = totalMin % 60;
        const ampm = h24 >= 12 ? 'PM' : 'AM';
        const h12 = h24 % 12 || 12;
        slots.push(`${h12}:${String(m).padStart(2, '0')}${ampm}`);
      }
      return slots;
    }

    const BMR_TARGET_COL_LETTERS = Array.from({length: BMR_SLOT_COUNT}, (_, k) => colLetter(12 + 12 * k));
    const BMR_USAGE_COL_LETTERS  = Array.from({length: BMR_SLOT_COUNT}, (_, k) => colLetter(13 + 12 * k));

    function argbHex(hex, fallback = 'FFFFFF00') {
      if (!hex) return fallback;
      const h = String(hex).replace('#', '').toUpperCase();
      if (h.length === 6) return 'FF' + h;
      if (h.length === 8) return h;
      return fallback;
    }

    function bmrAccountManagersById(bmr) {
      return new Map((bmr.accountManagers || []).map(manager => [manager.id, manager]));
    }

    function bmrClientColor(bmr, client, accountManagersById = bmrAccountManagersById(bmr)) {
      const manager = accountManagersById.get(client.accountManagerId);
      return manager?.color || client.color || '';
    }

    function bmrBuildConditionFormula(rule, cellRef) {
      const v = (x) => isFinite(Number(x)) ? Number(x) : 0;
      const cleanText = `TRIM(SUBSTITUTE(${cellRef}&"",CHAR(160),""))`;
      const hasValue = `LEN(${cleanText})>0`;
      const numericRef = `IFERROR(VALUE(SUBSTITUTE(${cleanText},",","")),${cellRef})`;
      switch (rule.kind) {
        case 'between': return `AND(${hasValue},ISNUMBER(${numericRef}),${numericRef}>=${v(rule.min)},${numericRef}<=${v(rule.max)})`;
        case 'gte':     return `AND(${hasValue},ISNUMBER(${numericRef}),${numericRef}>=${v(rule.value)})`;
        case 'lte':     return `AND(${hasValue},ISNUMBER(${numericRef}),${numericRef}<=${v(rule.value)})`;
        case 'gt':      return `AND(${hasValue},ISNUMBER(${numericRef}),${numericRef}>${v(rule.value)})`;
        case 'lt':      return `AND(${hasValue},ISNUMBER(${numericRef}),${numericRef}<${v(rule.value)})`;
        case 'eq':      return `AND(${hasValue},ISNUMBER(${numericRef}),${numericRef}=${v(rule.value)})`;
        default:        return 'FALSE';
      }
    }

    function bmrNumericExpression(cellRef) {
      const cleanText = `TRIM(SUBSTITUTE(${cellRef}&"",CHAR(160),""))`;
      return `IFERROR(VALUE(SUBSTITUTE(${cleanText},",","")),${cellRef})`;
    }

    function bmrSmsOverdraftConditionFormula(balanceRef, overdraftRef) {
      const balanceText = `TRIM(SUBSTITUTE(${balanceRef}&"",CHAR(160),""))`;
      const overdraftText = `TRIM(SUBSTITUTE(${overdraftRef}&"",CHAR(160),""))`;
      const balanceNum = bmrNumericExpression(balanceRef);
      const overdraftNum = bmrNumericExpression(overdraftRef);
      return `AND(LEN(${balanceText})>0,LEN(${overdraftText})>0,ISNUMBER(${balanceNum}),ISNUMBER(${overdraftNum}),${overdraftNum}>0,${balanceNum}<=-ABS(${overdraftNum}))`;
    }

    // True when the same-row "30mins usage" cell holds a non-zero number. Used to
    // bold the BALANCE cell so slots with usage stand out and BALANCE is easy to
    // pick out from the Over Draft / OD column (which sit in swapped positions on
    // the RES vs WHS sheets).
    function bmrSmsUsageNonZeroFormula(usageRef) {
      const usageText = `TRIM(SUBSTITUTE(${usageRef}&"",CHAR(160),""))`;
      const usageNum = bmrNumericExpression(usageRef);
      return `AND(LEN(${usageText})>0,ISNUMBER(${usageNum}),${usageNum}<>0)`;
    }

    function bmrBuildClientConditionFormula(rule, cellRef, clientName, rowNum) {
      const safeName = String(clientName || '').replace(/"/g, '""');
      if (!safeName) return 'FALSE';
      return `AND(${bmrBuildConditionFormula(rule, cellRef)},EXACT($A${rowNum},"${safeName}"))`;
    }

    /* ------------------------------------------------------------
       Red-cell counters — one big number per time slot showing how
       many Amnt Receivable / BALANCE cells that slot's rules highlight.
       Conditional formatting is dynamic, so we replicate the exact same
       conditions in a SUMPRODUCT that counts matches over the whole
       column. AND() collapses an array to a single scalar, so these
       versions rebuild the logic with multiplication over ranges. Kept
       byte-for-byte in sync with bmrBuildConditionFormula so the count
       always matches what the eye sees.
       ------------------------------------------------------------ */
    function bmrArrayNumericExpr(rangeRef) {
      const cleanText = `TRIM(SUBSTITUTE(${rangeRef}&"",CHAR(160),""))`;
      return `IFERROR(VALUE(SUBSTITUTE(${cleanText},",","")),${rangeRef})`;
    }
    function bmrArrayHasValue(rangeRef) {
      return `(LEN(TRIM(SUBSTITUTE(${rangeRef}&"",CHAR(160),"")))>0)`;
    }
    function bmrArrayCompare(rule, numExpr) {
      const v = (x) => isFinite(Number(x)) ? Number(x) : 0;
      switch (rule.kind) {
        case 'between': return `(${numExpr}>=${v(rule.min)})*(${numExpr}<=${v(rule.max)})`;
        case 'gte':     return `(${numExpr}>=${v(rule.value)})`;
        case 'lte':     return `(${numExpr}<=${v(rule.value)})`;
        case 'gt':      return `(${numExpr}>${v(rule.value)})`;
        case 'lt':      return `(${numExpr}<${v(rule.value)})`;
        case 'eq':      return `(${numExpr}=${v(rule.value)})`;
        default:        return '0';
      }
    }
    // 0/1-per-row expression: which cells in rangeRef a set of client-scoped
    // rules (each rule targets one client, matched by exact name in aRangeRef)
    // would highlight. Returns null when there are no usable rules.
    function bmrClientRulesArrayExpr(ruleEntries, rangeRef, aRangeRef) {
      const valid = (ruleEntries || []).filter(e => e.client?.name);
      if (!valid.length) return null;
      const num = bmrArrayNumericExpr(rangeRef);
      const inner = valid.map(({ rule, client }) => {
        const safe = String(client.name).replace(/"/g, '""');
        return `(${bmrArrayCompare(rule, num)})*EXACT(${aRangeRef},"${safe}")`;
      }).join('+');
      return `${bmrArrayHasValue(rangeRef)}*ISNUMBER(${num})*--((${inner})>0)`;
    }
    // 0/1-per-row expression mirroring bmrSmsOverdraftConditionFormula.
    function bmrSmsOverdraftArrayExpr(balRange, odRange) {
      const balNum = bmrArrayNumericExpr(balRange);
      const odNum = bmrArrayNumericExpr(odRange);
      return `${bmrArrayHasValue(balRange)}*${bmrArrayHasValue(odRange)}*ISNUMBER(${balNum})*ISNUMBER(${odNum})*(${odNum}>0)*(${balNum}<=-ABS(${odNum}))`;
    }
    // Combine the per-row 0/1 expressions into a single count. A cell is red
    // if any expression flags it, so we sum then test >0 to avoid double
    // counting a cell matched by more than one rule.
    function bmrRedCountFormula(rowExprs) {
      const parts = (rowExprs || []).filter(Boolean);
      if (!parts.length) return null;
      return `SUMPRODUCT(--((${parts.join('+')})>0))`;
    }

    const BMR_RED_COUNT_FILL = 'FFFEE2E2';
    const BMR_RED_COUNT_INK  = 'FF991B1B';
    function applyBmrRedCountStyle(cell) {
      // 2x the default 11pt font so the tally reads at a glance, on a red wash.
      cell.font = { bold: true, size: 22, color: { argb: BMR_RED_COUNT_INK } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BMR_RED_COUNT_FILL } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.numFmt = '0';
      const edge = { style: 'medium', color: { argb: BMR_RED_COUNT_INK } };
      cell.border = { top: edge, left: edge, bottom: edge, right: edge };
    }
    function applyBmrRedCountLabel(cell, text) {
      cell.value = text;
      cell.font = { bold: true, size: 12, color: { argb: BMR_RED_COUNT_INK } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BMR_RED_COUNT_FILL } };
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
    }

    function bmrStyleFromRule(rule) {
      const style = {};
      if (rule.color)     style.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: argbHex(rule.color) } };
      const fontParts = {};
      if (rule.fontColor) fontParts.color = { argb: argbHex(rule.fontColor) };
      if (rule.bold)      fontParts.bold = true;
      if (rule.italic)    fontParts.italic = true;
      if (rule.underline) fontParts.underline = true;
      if (Object.keys(fontParts).length) style.font = fontParts;
      return style;
    }

    function bmrEvaluateRule(rule, value) {
      const n = Number(value);
      if (!isFinite(n)) return false;
      const v = (x) => Number(x) || 0;
      switch (rule.kind) {
        case 'between': return n >= v(rule.min) && n <= v(rule.max);
        case 'gte':     return n >= v(rule.value);
        case 'lte':     return n <= v(rule.value);
        case 'gt':      return n >  v(rule.value);
        case 'lt':      return n <  v(rule.value);
        case 'eq':      return n === v(rule.value);
        default:        return false;
      }
    }

    function bmrTodayString() {
      const d = new Date();
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }

    function applyBmrBorders(ws, bmr, totalCols, dataEndRow, slotCount, blockCols = BMR_BLOCK_COLS) {
      if (bmr.bordersOnData === false) return;

      const borderColor = bmr.borderColor || 'FF374151';
      const cellBorder = { style: bmr.borderStyle || 'thin', color: { argb: borderColor } };
      const timeBorder = { style: bmr.timeSeparatorBorderStyle || 'thick', color: { argb: borderColor } };

      for (let r = 2; r <= dataEndRow; r++) {
        for (let c = 1; c <= totalCols; c++) {
          ws.getCell(r, c).border = {
            top: cellBorder,
            left: cellBorder,
            bottom: cellBorder,
            right: cellBorder,
          };
        }
      }

      for (let k = 0; k < slotCount; k++) {
        const blockStart = 3 + k * blockCols;
        const blockEnd = blockStart + blockCols - 1;
        ws.getCell(1, blockStart).border = {
          top: cellBorder,
          left: timeBorder,
          bottom: cellBorder,
          right: timeBorder,
        };

        for (let r = 2; r <= dataEndRow; r++) {
          const firstCell = ws.getCell(r, blockStart);
          const lastCell = ws.getCell(r, blockEnd);
          firstCell.border = { ...firstCell.border, left: timeBorder };
          lastCell.border = { ...lastCell.border, right: timeBorder };
        }
      }
    }

    function applyBmrNumberFormats(ws, dataStartRow, dataEndRow, slotCount) {
      for (let k = 0; k < slotCount; k++) {
        const blockStart = 3 + k * BMR_BLOCK_COLS;
        for (let r = dataStartRow; r <= dataEndRow; r++) {
          BMR_DECIMAL_OFFSETS.forEach((offset) => {
            ws.getCell(r, blockStart + offset).numFmt = '0.00';
          });
        }
      }
    }

    function bmrCarrierNameWidth(clients) {
      const longest = (clients || []).reduce((max, client) => Math.max(max, String(client.name || '').length), 0);
      return Math.max(BMR_BLOCK_HEADERS[0].trim().length + 2, longest + 2);
    }

    function buildBmrSheet(ws, bmr, shift) {
      const slots = bmrSlotsForShift(shift);
      const visibleClients = bmrNamedClients((bmr.clients || []).filter(c => !c.hidden));
      const accountManagersById = bmrAccountManagersById(bmr);
      const totalCols = 2 + slots.length * BMR_BLOCK_COLS;
      const dataStartRow = 3;
      const dataRowCount = visibleClients.length + BMR_EXTRA_ROWS;
      const dataEndRow   = dataStartRow + dataRowCount - 1;

      // Row 1: time labels merged across each 12-col block
      for (let k = 0; k < slots.length; k++) {
        const c = 3 + k * BMR_BLOCK_COLS;
        const cell = ws.getCell(1, c);
        cell.value = slots[k];
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        ws.mergeCells(1, c, 1, c + BMR_BLOCK_COLS - 1);
      }
      ws.getRow(1).height = 22;

      // Row 2: per-block headers
      ws.getCell(2, 1).value = 'CarrierName ';
      ws.getCell(2, 2).value = 'Allow Bal';
      for (let k = 0; k < slots.length; k++) {
        const baseCol = 3 + k * BMR_BLOCK_COLS;
        for (let i = 0; i < BMR_BLOCK_HEADERS.length; i++) {
          ws.getCell(2, baseCol + i).value = BMR_BLOCK_HEADERS[i];
        }
      }
      ws.getRow(2).eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });
      ws.getRow(2).height = 24;

      // Client rows plus blank allowance rows with formulas for 30-min usage / total usage.
      for (let i = 0; i < dataRowCount; i++) {
        const cl = visibleClients[i];
        const r = dataStartRow + i;
        if (cl) {
          const clientColor = bmrClientColor(bmr, cl, accountManagersById);
          ws.getCell(r, 1).value = cl.name;
          ws.getCell(r, 2).value = cl.allowBalanceLabel || '';
          if (clientColor) {
            const fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argbHex(clientColor) } };
            ws.getCell(r, 1).fill = fill;
          }
        }
        for (let k = 0; k < slots.length; k++) {
          const baseCol = 3 + k * BMR_BLOCK_COLS;
          const amntCol  = colLetter(baseCol + BMR_AMOUNT_OFFSET);
          const usageCol = colLetter(baseCol + BMR_USAGE_OFFSET);
          if (k === 0) {
            ws.getCell(r, baseCol + BMR_USAGE_OFFSET).value = 0;
            ws.getCell(r, baseCol + BMR_TOTAL_OFFSET).value = 0;
          } else {
            const prevAmnt = colLetter(baseCol + BMR_AMOUNT_OFFSET - BMR_BLOCK_COLS);
            ws.getCell(r, baseCol + BMR_USAGE_OFFSET).value = { formula: `${prevAmnt}${r}-${amntCol}${r}` };
            const prevTotal = colLetter(baseCol + BMR_TOTAL_OFFSET - BMR_BLOCK_COLS);
            ws.getCell(r, baseCol + BMR_TOTAL_OFFSET).value = { formula: `${prevTotal}${r}+${usageCol}${r}` };
          }
        }
      }

      // Column widths
      const carrierNameWidth = bmrCarrierNameWidth(visibleClients);
      ws.getColumn(1).width = carrierNameWidth;
      ws.getColumn(2).width = 22;
      for (let c = 3; c <= totalCols; c++) ws.getColumn(c).width = 12;
      for (let k = 0; k < slots.length; k++) ws.getColumn(3 + k * BMR_BLOCK_COLS).width = carrierNameWidth;

      applyBmrNumberFormats(ws, dataStartRow, dataEndRow, slots.length);
      applyBmrBorders(ws, bmr, totalCols, dataEndRow, slots.length);

      // Conditional formatting — Amnt Receivable target rules for their selected clients
      let priorityCounter = 100;
      const visibleClientsById = new Map(visibleClients.map(c => [c.id, c]));
      (bmr.targetRules || []).filter(r => r.enabled).forEach((rule) => {
        const client = visibleClientsById.get(rule.clientId);
        if (!client?.name) return;
        BMR_TARGET_COL_LETTERS.forEach((cl) => {
          const ref = `${cl}${dataStartRow}:${cl}${dataEndRow}`;
          ws.addConditionalFormatting({
            ref,
            rules: [{
              type: 'expression',
              formulae: [bmrBuildClientConditionFormula(rule, `${cl}${dataStartRow}`, client.name, dataStartRow)],
              style: bmrStyleFromRule(rule),
              priority: priorityCounter++,
            }],
          });
        });
      });

      // 30-min usage rules
      (bmr.usageRules || []).filter(r => r.enabled).forEach((rule) => {
        BMR_USAGE_COL_LETTERS.forEach((cl) => {
          const ref = `${cl}${dataStartRow}:${cl}${dataEndRow}`;
          ws.addConditionalFormatting({
            ref,
            rules: [{
              type: 'expression',
              formulae: [bmrBuildConditionFormula(rule, `${cl}${dataStartRow}`)],
              style: bmrStyleFromRule(rule),
              priority: priorityCounter++,
            }],
          });
        });
      });

      // Client text-match colouring on carrier cells so the color follows pasted names.
      visibleClients.forEach((cl) => {
        const clientColor = bmrClientColor(bmr, cl, accountManagersById);
        if (!clientColor) return;
        const safe = (cl.name || '').replace(/"/g, '""');
        const carrierCols = ['A'];
        if (bmr.colorBlockCarrierNames !== false) {
          for (let k = 0; k < slots.length; k++) carrierCols.push(colLetter(3 + k * BMR_BLOCK_COLS));
        }
        ws.addConditionalFormatting({
          ref: carrierCols.map((carrierCol) => `${carrierCol}${dataStartRow}:${carrierCol}${dataEndRow}`).join(' '),
          rules: [{
            type: 'expression',
            formulae: [`EXACT(TRIM(SUBSTITUTE(A${dataStartRow}&"",CHAR(160),"")),"${safe}")`],
            style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: argbHex(clientColor) } } },
            priority: priorityCounter++,
          }],
        });
      });

      // Totals row: per-slot count of Amnt Receivable cells the target rules
      // paint red, printed 2x size so it is readable at a glance.
      const redCountRow = dataEndRow + 2;
      const enabledTargetRules = (bmr.targetRules || [])
        .filter(r => r.enabled)
        .map(rule => ({ rule, client: visibleClientsById.get(rule.clientId) }))
        .filter(entry => entry.client?.name);
      if (enabledTargetRules.length) {
        applyBmrRedCountLabel(ws.getCell(redCountRow, 1), 'Colored Amnt Receivable count');
        const aRange = `$A${dataStartRow}:$A${dataEndRow}`;
        for (let k = 0; k < slots.length; k++) {
          const amntCol = colLetter(3 + k * BMR_BLOCK_COLS + BMR_AMOUNT_OFFSET);
          const rangeRef = `${amntCol}${dataStartRow}:${amntCol}${dataEndRow}`;
          const formula = bmrRedCountFormula([
            bmrClientRulesArrayExpr(enabledTargetRules, rangeRef, aRange),
          ]);
          const cell = ws.getCell(redCountRow, 3 + k * BMR_BLOCK_COLS + BMR_AMOUNT_OFFSET);
          cell.value = formula ? { formula } : 0;
          applyBmrRedCountStyle(cell);
        }
        ws.getRow(redCountRow).height = 34;
      }

      // Hidden columns
      (bmr.hiddenCols || []).forEach((c) => {
        try { ws.getColumn(c).hidden = true; } catch (_) {}
      });
      const forcedHidden = new Set(BMR_FORCE_HIDDEN_OFFSETS);
      (bmr.hiddenBlockCols || []).forEach((offset) => forcedHidden.add(Number(offset)));
      forcedHidden.forEach((offset) => {
        const colOffset = Number(offset);
        if (!Number.isInteger(colOffset) || colOffset < 0 || colOffset >= BMR_BLOCK_COLS) return;
        for (let k = 0; k < slots.length; k++) {
          ws.getColumn(3 + k * BMR_BLOCK_COLS + colOffset).hidden = true;
        }
      });

      // Freeze: top 2 rows + first 2 cols
      ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 2 }];
    }

    async function buildBmrBlob(state) {
      const wb = new ExcelJS.Workbook();
      wb.creator = 'BMR Generator';
      wb.created = new Date();
      const bmr = state.bmr;
      const shifts = [];
      if (bmr.includeDay !== false)   shifts.push('day');
      if (bmr.includeNight !== false) shifts.push('night');
      if (!shifts.length) shifts.push('day');
      for (const shift of shifts) {
        const name = shift === 'day' ? '7AM-7PM' : '7PM-7AM';
        const ws = wb.addWorksheet(name);
        buildBmrSheet(ws, bmr, shift);
      }
      const buf = await wb.xlsx.writeBuffer();
      const filename = `BMR_VOIP_${bmrTodayString()}.xlsx`;
      return { blob: new Blob([buf], { type: XLSX_MIME }), filename, sheets: collectSheetGridDims(wb) };
    }

    async function generateBmrWorkbook(state) {
      const { blob, filename } = await buildBmrBlob(state);
      saveAs(blob, filename);
      return filename;
    }

    function bmrSmsClientNameWidth(clients) {
      const longest = (clients || []).reduce((max, client) => Math.max(max, String(client.name || '').length), 0);
      return Math.max(16, Math.min(42, longest + 2));
    }

    // Drop every blank-name client (not just trailing ones) so filler entries —
    // e.g. from pasting into a 1000-row range, which leaves hundreds of nameless
    // rows scattered through the list — don't emit empty rows that balloon the tab
    // (and the synced Google Sheet) up to its 1000-row default. Any client with a
    // name is always kept; the 10 manual-entry rows below the list are added
    // separately, so writing space is unaffected.
    function bmrNamedClients(clients) {
      return (clients || []).filter(client => String(client?.name || '').trim() !== '');
    }

    function bmrSmsTargetColLetters(market) {
      return Array.from({ length: BMR_SLOT_COUNT }, (_, k) =>
        colLetter(3 + k * BMR_SMS_BLOCK_COLS + market.targetOffset)
      );
    }

    function bmrSmsUsageColLetters() {
      return Array.from({ length: BMR_SLOT_COUNT }, (_, k) =>
        colLetter(3 + k * BMR_SMS_BLOCK_COLS + BMR_SMS_USAGE_OFFSET)
      );
    }

    function applyBmrSmsNumberFormats(ws, dataStartRow, dataEndRow) {
      for (let k = 0; k < BMR_SLOT_COUNT; k++) {
        const baseCol = 3 + k * BMR_SMS_BLOCK_COLS;
        [5, 6, BMR_SMS_USAGE_OFFSET, BMR_SMS_TOTAL_OFFSET].forEach((offset) => {
          for (let r = dataStartRow; r <= dataEndRow; r++) {
            ws.getCell(r, baseCol + offset).numFmt = '0.00';
          }
        });
      }
    }

    function bmrSmsSheetName(market, shift) {
      return `${market.code} ${shift === 'day' ? 'DAY (7AM-7PM)' : 'NIGHT (7PM-7AM)'}`;
    }

    function buildBmrSmsSheet(ws, sms, marketKey, shift) {
      const market = BMR_SMS_MARKETS[marketKey];
      const slots = bmrSlotsForShift(shift);
      const clients = bmrNamedClients((sms[market.clientsKey] || []).filter(c => !c.hidden));
      const accountManagersById = bmrAccountManagersById(sms);
      const totalCols = 2 + slots.length * BMR_SMS_BLOCK_COLS;
      const dataStartRow = 3;
      const dataRowCount = clients.length + BMR_EXTRA_ROWS;
      const dataEndRow = dataStartRow + dataRowCount - 1;

      for (let k = 0; k < slots.length; k++) {
        const c = 3 + k * BMR_SMS_BLOCK_COLS;
        const cell = ws.getCell(1, c);
        cell.value = slots[k];
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        ws.mergeCells(1, c, 1, c + BMR_SMS_BLOCK_COLS - 1);
      }
      ws.getRow(1).height = 22;

      ws.getCell(2, 1).value = 'CLIENT';
      ws.getCell(2, 2).value = 'Allow Bal';
      for (let k = 0; k < slots.length; k++) {
        const baseCol = 3 + k * BMR_SMS_BLOCK_COLS;
        market.headers.forEach((header, offset) => {
          ws.getCell(2, baseCol + offset).value = header;
        });
      }
      ws.getRow(2).eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });
      ws.getRow(2).height = 24;

      for (let i = 0; i < dataRowCount; i++) {
        const client = clients[i];
        const r = dataStartRow + i;
        if (client) {
          const clientColor = bmrClientColor(sms, client, accountManagersById);
          ws.getCell(r, 1).value = client.name;
          ws.getCell(r, 2).value = client.allowBalanceLabel || '';
          if (clientColor) {
            ws.getCell(r, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argbHex(clientColor) } };
          }
        }

        for (let k = 0; k < slots.length; k++) {
          const baseCol = 3 + k * BMR_SMS_BLOCK_COLS;
          const balanceCol = colLetter(baseCol + market.balanceOffset);
          const usageCol = colLetter(baseCol + BMR_SMS_USAGE_OFFSET);
          if (k === 0) {
            ws.getCell(r, baseCol + BMR_SMS_USAGE_OFFSET).value = 0;
            ws.getCell(r, baseCol + BMR_SMS_TOTAL_OFFSET).value = 0;
          } else {
            const prevBalance = colLetter(baseCol + market.balanceOffset - BMR_SMS_BLOCK_COLS);
            const prevTotal = colLetter(baseCol + (k === 1 ? BMR_SMS_USAGE_OFFSET : BMR_SMS_TOTAL_OFFSET) - BMR_SMS_BLOCK_COLS);
            ws.getCell(r, baseCol + BMR_SMS_USAGE_OFFSET).value = { formula: `${prevBalance}${r}-${balanceCol}${r}` };
            ws.getCell(r, baseCol + BMR_SMS_TOTAL_OFFSET).value = { formula: `${prevTotal}${r}+${usageCol}${r}` };
          }
        }
      }

      const clientWidth = bmrSmsClientNameWidth(clients);
      ws.getColumn(1).width = clientWidth;
      ws.getColumn(2).width = 22;
      for (let k = 0; k < slots.length; k++) {
        const baseCol = 3 + k * BMR_SMS_BLOCK_COLS;
        ws.getColumn(baseCol).width = 16;
        ws.getColumn(baseCol + 1).width = clientWidth;
        ws.getColumn(baseCol + 2).width = 4;
        ws.getColumn(baseCol + 3).width = 4;
        ws.getColumn(baseCol + 4).width = 4;
        ws.getColumn(baseCol + 5).width = 14;
        ws.getColumn(baseCol + 6).width = 14;
        ws.getColumn(baseCol + 7).width = 14;
        ws.getColumn(baseCol + 8).width = 14;
        ws.getColumn(baseCol + 2).hidden = true;
        ws.getColumn(baseCol + 3).hidden = true;
        ws.getColumn(baseCol + 4).hidden = true;
      }

      applyBmrSmsNumberFormats(ws, dataStartRow, dataEndRow);
      applyBmrBorders(ws, sms, totalCols, dataEndRow, slots.length, BMR_SMS_BLOCK_COLS);

      let priorityCounter = 300;
      const visibleClientsById = new Map(clients.map(c => [c.id, c]));
      (sms.overdraftRules || []).filter(r => r.enabled).forEach((rule) => {
        for (let k = 0; k < slots.length; k++) {
          const baseCol = 3 + k * BMR_SMS_BLOCK_COLS;
          const balanceCol = colLetter(baseCol + market.balanceOffset);
          const overdraftCol = colLetter(baseCol + market.overdraftOffset);
          ws.addConditionalFormatting({
            ref: `${balanceCol}${dataStartRow}:${balanceCol}${dataEndRow}`,
            rules: [{
              type: 'expression',
              formulae: [bmrSmsOverdraftConditionFormula(`${balanceCol}${dataStartRow}`, `${overdraftCol}${dataStartRow}`)],
              style: bmrStyleFromRule(rule),
              priority: priorityCounter++,
            }],
          });
        }
      });
      const targetCols = bmrSmsTargetColLetters(market);
      (sms.targetRules || []).filter(r => r.enabled).forEach((rule) => {
        const client = visibleClientsById.get(rule.clientId);
        if (!client?.name) return;
        targetCols.forEach((cl) => {
          ws.addConditionalFormatting({
            ref: `${cl}${dataStartRow}:${cl}${dataEndRow}`,
            rules: [{
              type: 'expression',
              formulae: [bmrBuildClientConditionFormula(rule, `${cl}${dataStartRow}`, client.name, dataStartRow)],
              style: bmrStyleFromRule(rule),
              priority: priorityCounter++,
            }],
          });
        });
      });

      bmrSmsUsageColLetters().forEach((cl) => {
        (sms.usageRules || []).filter(r => r.enabled).forEach((rule) => {
          ws.addConditionalFormatting({
            ref: `${cl}${dataStartRow}:${cl}${dataEndRow}`,
            rules: [{
              type: 'expression',
              formulae: [bmrBuildConditionFormula(rule, `${cl}${dataStartRow}`)],
              style: bmrStyleFromRule(rule),
              priority: priorityCounter++,
            }],
          });
        });
      });

      // Bold each BALANCE cell whose same-row 30mins usage is non-zero, so slots
      // with usage stand out and BALANCE reads apart from the Over Draft / OD
      // column. Bold-only style layers over the overdraft/target fills above.
      if (sms.boldBalanceOnUsage !== false) {
        for (let k = 0; k < slots.length; k++) {
          const baseCol = 3 + k * BMR_SMS_BLOCK_COLS;
          const balanceCol = colLetter(baseCol + market.balanceOffset);
          const usageCol = colLetter(baseCol + BMR_SMS_USAGE_OFFSET);
          ws.addConditionalFormatting({
            ref: `${balanceCol}${dataStartRow}:${balanceCol}${dataEndRow}`,
            rules: [{
              type: 'expression',
              formulae: [bmrSmsUsageNonZeroFormula(`${usageCol}${dataStartRow}`)],
              style: { font: { bold: true } },
              priority: priorityCounter++,
            }],
          });
        }
      }

      clients.forEach((client) => {
        const clientColor = bmrClientColor(sms, client, accountManagersById);
        if (!clientColor) return;
        const safe = (client.name || '').replace(/"/g, '""');
        const clientCols = ['A'];
        if (sms.colorBlockCarrierNames !== false) {
          for (let k = 0; k < slots.length; k++) {
            const baseCol = 3 + k * BMR_SMS_BLOCK_COLS;
            clientCols.push(colLetter(baseCol), colLetter(baseCol + 1));
          }
        }
        ws.addConditionalFormatting({
          ref: clientCols.map((clientCol) => `${clientCol}${dataStartRow}:${clientCol}${dataEndRow}`).join(' '),
          rules: [{
            type: 'expression',
            formulae: [`EXACT(TRIM(SUBSTITUTE(A${dataStartRow}&"",CHAR(160),"")),"${safe}")`],
            style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: argbHex(clientColor) } } },
            priority: priorityCounter++,
          }],
        });
      });

      // Totals row: per-slot count of BALANCE cells the overdraft + target
      // rules paint red, printed 2x size so it is readable at a glance.
      const redCountRow = dataEndRow + 2;
      const overdraftEnabled = (sms.overdraftRules || []).some(r => r.enabled);
      const enabledTargetRules = (sms.targetRules || [])
        .filter(r => r.enabled)
        .map(rule => ({ rule, client: visibleClientsById.get(rule.clientId) }))
        .filter(entry => entry.client?.name);
      if (overdraftEnabled || enabledTargetRules.length) {
        applyBmrRedCountLabel(ws.getCell(redCountRow, 1), `Red ${market.label} balance count`);
        const aRange = `$A${dataStartRow}:$A${dataEndRow}`;
        for (let k = 0; k < slots.length; k++) {
          const baseCol = 3 + k * BMR_SMS_BLOCK_COLS;
          const balanceCol = colLetter(baseCol + market.balanceOffset);
          const overdraftCol = colLetter(baseCol + market.overdraftOffset);
          const balRange = `${balanceCol}${dataStartRow}:${balanceCol}${dataEndRow}`;
          const odRange = `${overdraftCol}${dataStartRow}:${overdraftCol}${dataEndRow}`;
          const formula = bmrRedCountFormula([
            overdraftEnabled ? bmrSmsOverdraftArrayExpr(balRange, odRange) : null,
            bmrClientRulesArrayExpr(enabledTargetRules, balRange, aRange),
          ]);
          const cell = ws.getCell(redCountRow, baseCol + market.balanceOffset);
          cell.value = formula ? { formula } : 0;
          applyBmrRedCountStyle(cell);
        }
        ws.getRow(redCountRow).height = 34;
      }

      ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 2 }];
    }

    async function buildBmrSmsBlob(state) {
      const wb = new ExcelJS.Workbook();
      wb.creator = 'BMR SMS Generator';
      wb.created = new Date();
      const sms = state.bmrSms || DEFAULT_BMR_SMS_STATE;
      const shifts = [];
      if (sms.includeDay !== false) shifts.push('day');
      if (sms.includeNight !== false) shifts.push('night');
      if (!shifts.length) shifts.push('day');

      shifts.forEach((shift) => {
        ['retail', 'wholesale'].forEach((marketKey) => {
          const market = BMR_SMS_MARKETS[marketKey];
          const ws = wb.addWorksheet(bmrSmsSheetName(market, shift));
          buildBmrSmsSheet(ws, sms, marketKey, shift);
        });
      });

      const buf = await wb.xlsx.writeBuffer();
      const filename = `BMR_SMS_${bmrTodayString()}.xlsx`;
      return { blob: new Blob([buf], { type: XLSX_MIME }), filename, sheets: collectSheetGridDims(wb) };
    }

    async function generateBmrSmsWorkbook(state) {
      const { blob, filename } = await buildBmrSmsBlob(state);
      saveAs(blob, filename);
      return filename;
    }

    /* ============================================================
       Google Sheets sync (Drive upload → Sheets conversion)
       ------------------------------------------------------------
       Browser-only OAuth via Google Identity Services. Instead of
       re-implementing every cell/format via the Sheets API, we upload the
       exact same .xlsx the download button produces and let Drive convert it
       into a Google Sheet — so all exceljs formatting, formulas and tabs carry
       over unchanged. Access tokens live in memory only; nothing sensitive is
       persisted. Only the created file IDs (non-secret) are saved in state.
       ============================================================ */
    const googleSheetsSync = (() => {
      let tokenClient = null;
      let accessToken = null;
      let tokenExpiry = 0;      // epoch ms; token considered valid until this
      let grantedEmail = '';    // best-effort account label for the sync panel
      let grantedScopes = '';   // space-delimited scopes actually granted

      // Did the granted token include the Drive permission we need? Without it,
      // Drive calls fail with 403 "insufficient authentication scopes".
      const hasDriveScope = () => /https:\/\/www\.googleapis\.com\/auth\/drive(\.file)?/.test(grantedScopes);
      // Effective Client ID. Defaults to the hard-coded constant but can be set
      // at runtime from the in-app field (stored in state.googleSheets.clientId),
      // so users configure sync without editing source.
      let clientId = GOOGLE_OAUTH_CLIENT_ID || '';

      // We also request openid+email (non-sensitive) so the panel can show
      // which Google account is connected. drive.file is the only sensitive
      // scope and keeps the app limited to sheets it creates.
      const REQUEST_SCOPE = 'openid email ' + GOOGLE_DRIVE_SCOPE;

      const isConfigured = () => !!clientId;
      const gisReady = () => !!(window.google && window.google.accounts && window.google.accounts.oauth2);

      // Point the sync at a Client ID. Rebuilds the token client on change so
      // the next request uses the new ID.
      function setClientId(id) {
        const v = String(id || '').trim();
        if (v === clientId) return false;
        clientId = v;
        tokenClient = null;
        return true;
      }

      function ensureTokenClient() {
        if (!isConfigured()) throw new Error('Google OAuth Client ID is not set. Paste it into the Google Sheets sync field in the sidebar.');
        if (!gisReady()) throw new Error('Google sign-in library has not loaded yet. Check your connection and try again.');
        if (!tokenClient) {
          tokenClient = window.google.accounts.oauth2.initTokenClient({
            client_id: clientId,
            scope: REQUEST_SCOPE,
            callback: () => {}, // replaced per-request in requestToken()
          });
        }
        return tokenClient;
      }

      // Resolve an access token. interactive=true shows the account/consent
      // popup; interactive=false attempts a silent grant (prompt: '') which
      // succeeds once the user has consented before and has a live Google
      // session — no popup.
      function requestToken(interactive) {
        return new Promise((resolve, reject) => {
          let client;
          try { client = ensureTokenClient(); } catch (e) { reject(e); return; }
          client.callback = (resp) => {
            if (resp && resp.error) { reject(new Error(resp.error_description || resp.error)); return; }
            accessToken = resp.access_token;
            grantedScopes = resp.scope || '';
            const ttlMs = (Number(resp.expires_in) || 3600) * 1000;
            tokenExpiry = Date.now() + ttlMs - 60000; // refresh a minute early
            // A token can come back without the Drive permission if the user
            // left its checkbox unticked (or a prior grant lacked it). Treat
            // that as a distinct failure so getToken() can force re-consent.
            if (!hasDriveScope()) {
              accessToken = null; tokenExpiry = 0;
              const err = new Error('DRIVE_SCOPE_NOT_GRANTED');
              err.code = 'DRIVE_SCOPE_NOT_GRANTED';
              reject(err);
              return;
            }
            resolve(accessToken);
          };
          try {
            client.requestAccessToken({ prompt: interactive ? 'consent' : '' });
          } catch (e) { reject(e); }
        });
      }

      const SCOPE_HELP = 'Drive access wasn\'t granted. Click Connect again, and in the Google window keep the checkbox for "See, edit, create and delete only the specific Google Drive files you use with this app" ticked before continuing.';

      async function getToken({ interactive = false } = {}) {
        if (accessToken && Date.now() < tokenExpiry && hasDriveScope()) return accessToken;
        try {
          return await requestToken(interactive);
        } catch (e) {
          // Missing Drive scope (or a failed silent refresh) → escalate to an
          // interactive consent prompt once so the user can grant Drive access.
          if (!interactive) {
            try { return await requestToken(true); }
            catch (e2) { throw (e2.code === 'DRIVE_SCOPE_NOT_GRANTED' ? new Error(SCOPE_HELP) : e2); }
          }
          throw (e.code === 'DRIVE_SCOPE_NOT_GRANTED' ? new Error(SCOPE_HELP) : e);
        }
      }

      // Best-effort: fetch the connected account's email for display only.
      async function fetchEmail(token) {
        try {
          const r = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
            headers: { Authorization: 'Bearer ' + token },
          });
          if (r.ok) { const info = await r.json(); grantedEmail = info.email || ''; }
        } catch {}
        return grantedEmail;
      }

      async function connect() {
        const token = await getToken({ interactive: true });
        await fetchEmail(token);
        return { email: grantedEmail };
      }

      function disconnect() {
        if (accessToken && gisReady()) {
          try { window.google.accounts.oauth2.revoke(accessToken, () => {}); } catch {}
        }
        accessToken = null; tokenExpiry = 0; grantedEmail = ''; grantedScopes = '';
      }

      const isConnected = () => !!accessToken && Date.now() < tokenExpiry && hasDriveScope();
      const email = () => grantedEmail;
      const sheetUrl = (fileId) => 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(fileId) + '/edit';

      async function driveError(r) {
        let detail = '';
        try { const body = await r.json(); detail = body?.error?.message || ''; } catch {}
        const err = new Error('Google Drive error ' + r.status + (detail ? ': ' + detail : ''));
        err.status = r.status;
        return err;
      }

      // Create a new Google Sheet from xlsx bytes. Drive converts the upload
      // because the metadata mimeType is the Google Sheets type. Returns id.
      async function createSheet(name, blob) {
        const token = await getToken();
        const metadata = { name, mimeType: 'application/vnd.google-apps.spreadsheet' };
        const form = new FormData();
        form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
        form.append('file', blob, name + '.xlsx');
        const r = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
          method: 'POST',
          headers: { Authorization: 'Bearer ' + token },
          body: form,
        });
        if (!r.ok) throw await driveError(r);
        return (await r.json()).id;
      }

      // Replace an existing sheet's content with new xlsx bytes. The file keeps
      // its Google Sheets type (so the same id/url stays valid) and Drive
      // converts the uploaded xlsx into it.
      async function updateSheet(fileId, blob) {
        const token = await getToken();
        const r = await fetch('https://www.googleapis.com/upload/drive/v3/files/' + encodeURIComponent(fileId) + '?uploadType=media&fields=id', {
          method: 'PATCH',
          headers: { Authorization: 'Bearer ' + token, 'Content-Type': XLSX_MIME },
          body: blob,
        });
        if (!r.ok) throw await driveError(r);
        return fileId;
      }

      // After Drive converts an uploaded xlsx into a Google Sheet it pads every
      // tab out to the default 1000-row grid, leaving a long tail of empty rows
      // below the real data (so the totals row is no longer the last row). Shrink
      // each tab's grid back to exactly the rows/columns the workbook uses. The
      // Sheets API is reachable under the same drive.file scope for files this
      // app created, so no extra consent is needed. `dims` is the per-sheet
      // extent list from collectSheetGridDims (title/rows/cols).
      async function trimSheetGrids(fileId, dims) {
        const wanted = (dims || []).filter(d => d && d.title && d.rows > 0 && d.cols > 0);
        if (!fileId || !wanted.length) return;
        const token = await getToken();
        const base = 'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(fileId);
        const metaResp = await fetch(base + '?fields=sheets.properties(sheetId,title,gridProperties(rowCount,columnCount))', {
          headers: { Authorization: 'Bearer ' + token },
        });
        if (!metaResp.ok) throw await driveError(metaResp);
        const meta = await metaResp.json();
        const byTitle = new Map((meta.sheets || []).map(s => [s.properties && s.properties.title, s.properties]));
        const requests = [];
        wanted.forEach((want) => {
          const props = byTitle.get(want.title);
          if (!props) return;
          const grid = props.gridProperties || {};
          const curRows = grid.rowCount || 0;
          const curCols = grid.columnCount || 0;
          if (!curRows || !curCols) return;
          // Only ever shrink toward the used range — never grow, and never clip
          // below the data. Drive pads upward, so current >= used in practice.
          const newRows = curRows > want.rows ? want.rows : curRows;
          const newCols = curCols > want.cols ? want.cols : curCols;
          if (newRows === curRows && newCols === curCols) return;
          requests.push({
            updateSheetProperties: {
              properties: { sheetId: props.sheetId, gridProperties: { rowCount: newRows, columnCount: newCols } },
              fields: 'gridProperties.rowCount,gridProperties.columnCount',
            },
          });
        });
        if (!requests.length) return;
        const r = await fetch(base + ':batchUpdate', {
          method: 'POST',
          headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: JSON.stringify({ requests }),
        });
        if (!r.ok) throw await driveError(r);
      }

      return { isConfigured, setClientId, gisReady, getToken, connect, disconnect, isConnected, email, createSheet, updateSheet, trimSheetGrids, sheetUrl };
    })();

    // Modules that support Google Sheets sync → the in-memory workbook builder
    // and the (static) name used when the app first creates that module's sheet.
    // Only these four sync; other modules just download as before.
    const GOOGLE_SYNC_MODULES = {
      sip_fcs:  { build: buildSipFcsBlob,   sheetName: 'SIP FCS — Hourly Record' },
      bmr:      { build: buildBmrBlob,      sheetName: 'BMR VOIP — Balance Day & Night' },
      bmr_sms:  { build: buildBmrSmsBlob,   sheetName: 'BMR SMS — Balance Day & Night' },
      // buildRecorderBlob is a hoisted function declaration defined with the
      // Recorder module code further down this file.
      recorder: { build: buildRecorderBlob, sheetName: 'Recorder — VOS Hourly Record' },
    };

    /* ============================================================
       Editor module - pasted table editor
       ============================================================ */
    const editorRuleId = () => 'er_' + Math.random().toString(36).slice(2, 9);
    const DEFAULT_EDITOR_STATE = {
      name: 'Pasted Table',
      cells: [['']],
      headerRow: 1,
      labelColumn: 1,
      headerFill: '',
      headerFontColor: '',
      headerBold: true,
      headerItalic: false,
      headerUnderline: false,
      labelFill: '',
      labelFontColor: '',
      labelBold: false,
      labelItalic: false,
      labelUnderline: false,
      rowStyles: [],
      rules: [],
    };

    function editorNormalizeGrid(rows) {
      const source = Array.isArray(rows) && rows.length ? rows : [['']];
      const width = Math.max(1, ...source.map(row => Array.isArray(row) ? row.length : 1));
      return source.map(row => Array.from({ length: width }, (_, col) => String(Array.isArray(row) ? (row[col] ?? '') : (col ? '' : (row ?? '')))));
    }

    function editorParseGridText(text) {
      const raw = String(text ?? '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      if (!raw) return [['']];
      return editorNormalizeGrid(raw.split('\n').map(line => line.split('\t')));
    }

    function editorHeaderRow(editor, rowCount) {
      const value = Math.floor(Number(editor?.headerRow));
      return value > 0 ? Math.min(value, Math.max(1, rowCount)) : 0;
    }

    function editorLabelColumn(editor, colCount) {
      const value = Math.floor(Number(editor?.labelColumn));
      return value > 0 ? Math.min(value, Math.max(1, colCount)) : 0;
    }

    function editorCustomRowStyles(editor, rowCount) {
      return (Array.isArray(editor?.rowStyles) ? editor.rowStyles : [])
        .map((style) => ({ ...style, row: Math.floor(Number(style?.row)) }))
        .filter((style) => style.row > 0 && style.row <= rowCount && (style.fill || style.fontColor || style.bold || style.italic || style.underline));
    }

    function editorCustomRowStyleFor(editor, rowCount, rowNum) {
      const styles = editorCustomRowStyles(editor, rowCount).filter((style) => style.row === rowNum);
      return styles.length ? styles[styles.length - 1] : null;
    }

    function editorParseRowNumbers(text, rowCount) {
      const rows = new Set();
      String(text || '').split(/[\s,;]+/).filter(Boolean).forEach((part) => {
        const range = part.match(/^(\d+)-(\d+)$/);
        if (range) {
          const start = Number(range[1]);
          const end = Number(range[2]);
          const min = Math.min(start, end);
          const max = Math.max(start, end);
          for (let row = min; row <= max; row++) {
            if (row >= 1 && row <= rowCount) rows.add(row);
          }
          return;
        }
        const row = Number(part);
        if (Number.isInteger(row) && row >= 1 && row <= rowCount) rows.add(row);
      });
      return Array.from(rows).sort((a, b) => a - b);
    }

    function editorSnapshot(editor) {
      const merged = { ...DEFAULT_EDITOR_STATE, ...(editor || {}) };
      return JSON.parse(JSON.stringify({
        ...merged,
        cells: editorNormalizeGrid(merged.cells),
        rowStyles: Array.isArray(merged.rowStyles) ? merged.rowStyles : [],
        rules: Array.isArray(merged.rules) ? merged.rules : [],
      }));
    }

    function editorStatesEqual(a, b) {
      return JSON.stringify(a) === JSON.stringify(b);
    }

    function editorTemplateData(editor) {
      const source = editorSnapshot(editor);
      const keys = [
        'headerRow', 'labelColumn',
        'headerFill', 'headerFontColor', 'headerBold', 'headerItalic', 'headerUnderline',
        'labelFill', 'labelFontColor', 'labelBold', 'labelItalic', 'labelUnderline',
        'rowStyles', 'rules',
      ];
      return keys.reduce((acc, key) => ({ ...acc, [key]: source[key] }), {});
    }

    function editorApplyTemplate(editor, templateData) {
      const current = editorSnapshot(editor);
      const next = { ...current, ...(templateData || {}) };
      return editorSnapshot({ ...next, name: current.name, cells: current.cells });
    }

    function mergeLocalSheetActiveFlags(remote) {
      if (!remote || !Array.isArray(remote.sheets)) return remote;
      try {
        const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
        if (!saved || !Array.isArray(saved.sheets)) return remote;
        const activeById = new Map(saved.sheets
          .filter(sheet => sheet && typeof sheet.active === 'boolean')
          .map(sheet => [sheet.id, sheet.active]));
        if (!activeById.size) return remote;
        return {
          ...remote,
          sheets: remote.sheets.map(sheet => activeById.has(sheet.id) ? { ...sheet, active: activeById.get(sheet.id) } : sheet),
        };
      } catch {
        return remote;
      }
    }

    function editorNumber(value) {
      const clean = String(value ?? '').replace(/\u00a0/g, ' ').trim().replace(/,/g, '');
      if (!clean || !/^[-+]?(?:\d+|\d*\.\d+)$/.test(clean)) return null;
      const n = Number(clean);
      return Number.isFinite(n) ? n : null;
    }

    function editorEvaluateRule(rule, value, rowValues, dataStartCol) {
      if (!rule || rule.enabled === false) return false;
      const n = editorNumber(value);
      if (n === null) return false;
      const threshold = (x) => Number.isFinite(Number(x)) ? Number(x) : 0;
      if (rule.kind === 'maxInRow') {
        const numericRow = rowValues.slice(Math.max(0, dataStartCol - 1)).map(editorNumber).filter(v => v !== null);
        return numericRow.length > 0 && n === Math.max(...numericRow);
      }
      switch (rule.kind) {
        case 'between': return n >= threshold(rule.min) && n <= threshold(rule.max);
        case 'gte':     return n >= threshold(rule.value);
        case 'lte':     return n <= threshold(rule.value);
        case 'gt':      return n >  threshold(rule.value);
        case 'lt':      return n <  threshold(rule.value);
        case 'eq':      return n === threshold(rule.value);
        default:        return false;
      }
    }

    function editorNumericExport(value) {
      const raw = String(value ?? '');
      const text = raw.trim();
      const numericPattern = /^[-+]?(?:(?:0|[1-9]\d*)|(?:[1-9]\d{0,2}(?:,\d{3})+))(?:\.\d+)?$/;
      if (!numericPattern.test(text)) return null;
      const n = Number(text.replace(/,/g, ''));
      if (!Number.isFinite(n)) return null;
      const hasComma = text.includes(',');
      const decimals = text.match(/\.(\d+)$/)?.[1]?.length || 0;
      return {
        value: n,
        numFmt: `${hasComma ? '#,##' : ''}0${decimals ? `.${'0'.repeat(decimals)}` : ''}`,
      };
    }

    function editorCellExportValue(value, isDataCell = false) {
      const text = String(value ?? '');
      if (text.startsWith('=') && text.length > 1) return { value: { formula: text.slice(1) } };
      if (isDataCell) {
        const numeric = editorNumericExport(text);
        if (numeric) return numeric;
      }
      return { value: text };
    }

    function editorBuildConditionFormula(rule, cellRef) {
      const v = (x) => Number.isFinite(Number(x)) ? Number(x) : 0;
      switch (rule.kind) {
        case 'between': return `AND(ISNUMBER(${cellRef}),${cellRef}>=${v(rule.min)},${cellRef}<=${v(rule.max)})`;
        case 'gte':     return `AND(ISNUMBER(${cellRef}),${cellRef}>=${v(rule.value)})`;
        case 'lte':     return `AND(ISNUMBER(${cellRef}),${cellRef}<=${v(rule.value)})`;
        case 'gt':      return `AND(ISNUMBER(${cellRef}),${cellRef}>${v(rule.value)})`;
        case 'lt':      return `AND(ISNUMBER(${cellRef}),${cellRef}<${v(rule.value)})`;
        case 'eq':      return `AND(ISNUMBER(${cellRef}),${cellRef}=${v(rule.value)})`;
        default:        return 'FALSE';
      }
    }

    function editorRuleFormula(rule, cellRef, firstDataCol, lastDataCol, rowNum) {
      if (rule.kind !== 'maxInRow') return editorBuildConditionFormula(rule, cellRef);
      const rowRange = `$${firstDataCol}${rowNum}:$${lastDataCol}${rowNum}`;
      return `AND(ISNUMBER(${cellRef}),${cellRef}=MAX(${rowRange}))`;
    }

    function editorFontStylePatch(style = {}, fontColor = '') {
      const patch = {};
      if (fontColor) patch.color = { argb: argbHex(fontColor) };
      if (typeof style.bold === 'boolean') patch.bold = style.bold;
      if (typeof style.italic === 'boolean') patch.italic = style.italic;
      if (typeof style.underline === 'boolean') patch.underline = style.underline;
      return patch;
    }

    function editorApplyStaticStyle(cell, fillColor, fontColor, fontStyle = {}) {
      if (fillColor) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argbHex(fillColor) } };
      const patch = editorFontStylePatch(fontStyle, fontColor);
      if (Object.keys(patch).length) cell.font = { ...(cell.font || {}), ...patch };
    }

    function editorApplyCustomRowStyle(cell, rowStyle) {
      if (!rowStyle) return;
      if (rowStyle.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argbHex(rowStyle.fill) } };
      const patch = editorFontStylePatch(rowStyle, rowStyle.fontColor);
      if (Object.keys(patch).length) cell.font = { ...(cell.font || {}), ...patch };
    }

    function buildEditorSheet(ws, editor) {
      const grid = editorNormalizeGrid(editor?.cells);
      const rowCount = grid.length;
      const colCount = grid[0].length;
      const headerRow = editorHeaderRow(editor, rowCount);
      const labelCol = editorLabelColumn(editor, colCount);
      const dataStartRow = headerRow ? headerRow + 1 : 1;
      const dataStartCol = labelCol ? labelCol + 1 : 1;
      const thinBorder = { style: 'thin', color: { argb: 'FFD4D4D8' } };

      grid.forEach((row, rowIndex) => {
        row.forEach((value, colIndex) => {
          const cell = ws.getCell(rowIndex + 1, colIndex + 1);
          const isDataCell = rowIndex + 1 >= dataStartRow && colIndex + 1 >= dataStartCol;
          const exported = editorCellExportValue(value, isDataCell);
          cell.value = exported.value;
          if (exported.numFmt) cell.numFmt = exported.numFmt;
          cell.alignment = { vertical: 'middle', wrapText: true };
          cell.border = { top: thinBorder, left: thinBorder, bottom: thinBorder, right: thinBorder };
        });
      });

      if (headerRow) {
        const headerFontStyle = { bold: editor.headerBold !== false, italic: !!editor.headerItalic, underline: !!editor.headerUnderline };
        for (let col = 1; col <= colCount; col++) editorApplyStaticStyle(ws.getCell(headerRow, col), editor.headerFill, editor.headerFontColor, headerFontStyle);
      }
      if (labelCol) {
        const labelFontStyle = { bold: !!editor.labelBold, italic: !!editor.labelItalic, underline: !!editor.labelUnderline };
        for (let row = 1; row <= rowCount; row++) {
          if (row !== headerRow) editorApplyStaticStyle(ws.getCell(row, labelCol), editor.labelFill, editor.labelFontColor, labelFontStyle);
        }
      }
      editorCustomRowStyles(editor, rowCount).forEach((rowStyle) => {
        for (let col = 1; col <= colCount; col++) editorApplyCustomRowStyle(ws.getCell(rowStyle.row, col), rowStyle);
      });

      for (let col = 1; col <= colCount; col++) {
        const longest = Math.max(...grid.map(row => String(row[col - 1] ?? '').length), 4);
        ws.getColumn(col).width = Math.min(34, Math.max(col === labelCol ? 14 : 10, longest + 2));
      }

      if (dataStartRow <= rowCount && dataStartCol <= colCount) {
        const firstCol = colLetter(dataStartCol);
        const lastCol = colLetter(colCount);
        const ref = `${firstCol}${dataStartRow}:${lastCol}${rowCount}`;
        (editor.rules || []).filter(rule => rule.enabled !== false).forEach((rule, index) => {
          ws.addConditionalFormatting({
            ref,
            rules: [{
              type: 'expression',
              formulae: [editorRuleFormula(rule, `${firstCol}${dataStartRow}`, firstCol, lastCol, dataStartRow)],
              style: bmrStyleFromRule(rule),
              priority: index + 1,
            }],
          });
        });
      }

      ws.views = [{ state: 'frozen', xSplit: labelCol || 0, ySplit: headerRow || 0 }];
    }

    async function generateEditorWorkbook(state) {
      const editor = { ...DEFAULT_EDITOR_STATE, ...(state.editor || {}) };
      const wb = new ExcelJS.Workbook();
      wb.creator = 'Table Editor';
      wb.created = new Date();
      const ws = wb.addWorksheet(safeSheetName(editor.name || 'Pasted Table'));
      buildEditorSheet(ws, editor);
      const buf = await wb.xlsx.writeBuffer();
      const filename = `Editor_${bmrTodayString()}.xlsx`;
      saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);
      return filename;
    }

    /* ============================================================
       Whitelist SMS module - paste content blocks, pick test numbers,
       generate xlsx blast file
       ============================================================ */
    const WL_SMS_NETWORKS = {
      smart: { id: 'smart', label: 'PH Smart Mob', shortLabel: 'Smart', color: '#22C55E' },
      globe: { id: 'globe', label: 'PH Globe Mob', shortLabel: 'Globe', color: '#3B82F6' },
      dito:  { id: 'dito',  label: 'PH DITO',      shortLabel: 'DITO',  color: '#F59E0B' },
    };
    const WL_SMS_NETWORK_ORDER = ['smart', 'globe', 'dito'];

    const wlSmsId = (prefix = 'wl') => `${prefix}_${Math.random().toString(36).slice(2, 9)}`;
    // Identity key for a test number — used to dedupe across the shared database.
    // Matching by digits + network avoids duplicates from formatting differences (spaces, +, etc.).
    const wlSmsNumberKey = (n) => `${String(n?.number || '').replace(/\D/g, '')}|${n?.network || ''}`;
    const SHARED_WL_SMS_NUMBERS_DOC = ['shared', 'whitelistSmsTestNumbers'];

    const DEFAULT_WL_SMS_STATE = {
      rawPaste: '',
      contents: [],
      senderId: '',
      networks: ['smart', 'globe', 'dito'],
      testNumbers: [
        { id: 'tn_smart_1', label: 'Smart Test (Pao)', number: '639081989019', network: 'smart' },
        { id: 'tn_globe_1', label: 'Globe Test (Pao)', number: '639766707548', network: 'globe' },
      ],
      selectedNumberIds: ['tn_smart_1', 'tn_globe_1'],
      sheetName: 'Sheet1',
      separatorMode: 'auto',
      linkBuilderRaw: '',
      notes: '',
    };

    const WL_SMS_LINK_BUILDER_SAMPLE = `Your account ({{phone-10}}) received 200PHP reward, send out within 24 hours. Tap link:

https://bit.ly/4dVLLXR
https://bit.ly/4fqbmed
https://bit.ly/434ofmN
https://bit.ly/4dSm7mV
https://bit.ly/4fl8bUV
https://bit.ly/4voKT5n
https://bit.ly/4dWCWxg
https://bit.ly/43bhG1L
https://bit.ly/434WOJx
https://bit.ly/4vou9Lv
https://bit.ly/4ahjc6b
https://bit.ly/4fgsOSc
https://bit.ly/4vrcu64`;

    function wlSmsCleanUrl(value) {
      return String(value || '').trim().replace(/[)\],.;]+$/g, '');
    }

    function wlSmsExtractUrls(text) {
      return (String(text || '').match(/https?:\/\/[^\s<>"']+/gi) || [])
        .map(wlSmsCleanUrl)
        .filter(Boolean);
    }

    function wlSmsLinkTemplate(text) {
      return String(text || '')
        .replace(/https?:\/\/[^\s<>"']+/gi, ' ')
        .split(/\n/)
        .map(line => line.trim())
        .filter(Boolean)
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();
    }

    function wlSmsBuildLinkMessages(text) {
      const urls = wlSmsExtractUrls(text);
      const template = wlSmsLinkTemplate(text);
      const messages = urls.map(url => {
        if (!template) return url;
        if (/\{\{\s*(link|url)\s*\}\}/i.test(template)) {
          return template.replace(/\{\{\s*(link|url)\s*\}\}/gi, url).trim();
        }
        return `${template}${/\s$/.test(template) ? '' : ' '}${url}`.trim();
      });
      return {
        urls,
        template,
        messages,
        output: messages.join('\n\n'),
      };
    }

    async function copyTextToClipboard(text) {
      if (!text) return false;
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(text);
          return true;
        }
      } catch (e) {
        console.warn('clipboard api failed', e);
      }
      const area = document.createElement('textarea');
      area.value = text;
      area.setAttribute('readonly', '');
      area.style.position = 'fixed';
      area.style.left = '-9999px';
      document.body.appendChild(area);
      area.select();
      try {
        return document.execCommand('copy');
      } catch (e) {
        console.warn('clipboard fallback failed', e);
        return false;
      } finally {
        document.body.removeChild(area);
      }
    }

    function stripWlSmsLeadingNumber(block) {
      return block.replace(/^\s*\d+[.\)]\s+/, '');
    }

    function parseWlSmsBlasts(text, mode = 'auto') {
      if (text === null || text === undefined) return [];
      const normalized = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      const lines = normalized.split('\n');
      if (mode === 'lines') {
        return lines.map(l => l.trim()).filter(l => l.length).map(stripWlSmsLeadingNumber);
      }
      const isDivider = (line) => {
        const trimmed = line.trim();
        if (trimmed.length < 3) return false;
        return /^[—–\-=_*\s]+$/.test(trimmed) && /[—–\-=_*]/.test(trimmed);
      };
      const isBlank = (line) => line.trim() === '';
      const hasDivider = lines.some(isDivider);
      const isSeparator = hasDivider ? isDivider : isBlank;
      const blocks = [];
      let current = [];
      for (const line of lines) {
        if (isSeparator(line)) {
          const block = current.join('\n').trim();
          if (block) blocks.push(stripWlSmsLeadingNumber(block));
          current = [];
        } else {
          current.push(line);
        }
      }
      const tail = current.join('\n').trim();
      if (tail) blocks.push(stripWlSmsLeadingNumber(tail));
      return blocks;
    }

    // GSM 7-bit default alphabet + extension table. Any character outside this
    // set forces the whole SMS into UCS-2 (UTF-16) encoding, which drops the
    // per-part limit from 160 to 70 characters.
    const WL_SMS_GSM7_SET = new Set(
      '@£$¥èéùìòÇ\nØø\rÅåΔ_ΦΓΛΩΠΨΣΘΞÆæßÉ !"#¤%&\'()*+,-./0123456789:;<=>?¡ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÑܧ¿abcdefghijklmnopqrstuvwxyzäöñüà'
      + '^{}\\[~]|€'
    );

    // Friendly names + GSM-safe replacements for the special characters that
    // most often sneak in from Word / Google Docs copy-paste. fix: '' means
    // "safe to remove"; entries without fix are flagged but never auto-fixed.
    const WL_SMS_SPECIAL_CHAR_INFO = {
      '—': { name: 'Em dash', fix: '-' },
      '–': { name: 'En dash', fix: '-' },
      '―': { name: 'Horizontal bar', fix: '-' },
      '‐': { name: 'Hyphen (Unicode)', fix: '-' },
      '‑': { name: 'Non-breaking hyphen', fix: '-' },
      '‒': { name: 'Figure dash', fix: '-' },
      '−': { name: 'Minus sign', fix: '-' },
      '‘': { name: 'Left single quote', fix: "'" },
      '’': { name: 'Right single quote (curly apostrophe)', fix: "'" },
      '‚': { name: 'Low single quote', fix: "'" },
      '‛': { name: 'Reversed single quote', fix: "'" },
      '′': { name: 'Prime', fix: "'" },
      '`': { name: 'Backtick', fix: "'" },
      '´': { name: 'Acute accent', fix: "'" },
      '“': { name: 'Left double quote', fix: '"' },
      '”': { name: 'Right double quote', fix: '"' },
      '„': { name: 'Low double quote', fix: '"' },
      '″': { name: 'Double prime', fix: '"' },
      '«': { name: 'Left angle quote', fix: '"' },
      '»': { name: 'Right angle quote', fix: '"' },
      '‹': { name: 'Left single angle quote', fix: "'" },
      '›': { name: 'Right single angle quote', fix: "'" },
      '…': { name: 'Ellipsis', fix: '...' },
      '•': { name: 'Bullet', fix: '-' },
      '·': { name: 'Middle dot', fix: '-' },
      ' ': { name: 'Non-breaking space', fix: ' ' },
      ' ': { name: 'Narrow non-breaking space', fix: ' ' },
      ' ': { name: 'En space', fix: ' ' },
      ' ': { name: 'Em space', fix: ' ' },
      ' ': { name: 'Thin space', fix: ' ' },
      '\u2028': { name: 'Line separator', fix: '\n' },
      '\u2029': { name: 'Paragraph separator', fix: '\n' },
      '­': { name: 'Soft hyphen', fix: '' },
      '​': { name: 'Zero-width space', fix: '' },
      '‌': { name: 'Zero-width non-joiner', fix: '' },
      '‍': { name: 'Zero-width joiner', fix: '' },
      '﻿': { name: 'Zero-width no-break space (BOM)', fix: '' },
      '™': { name: 'Trademark sign', fix: 'TM' },
      '®': { name: 'Registered sign', fix: '(R)' },
      '©': { name: 'Copyright sign', fix: '(C)' },
    };

    function wlSmsCodePointLabel(ch) {
      return 'U+' + ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0');
    }

    // Invisible offenders (nbsp, zero-width chars) render as their codepoint
    // so the highlight has something visible to show.
    function wlSmsCharVisible(ch) {
      return /[\s\u00AD\u200B-\u200D\uFEFF\u2028\u2029]/.test(ch) ? wlSmsCodePointLabel(ch) : ch;
    }

    function wlSmsAnalyzeUtf(text) {
      const segments = []; // [{ text, special }] — consecutive runs, in order
      const offenders = new Map(); // char -> count
      for (const ch of String(text ?? '')) {
        const special = !WL_SMS_GSM7_SET.has(ch);
        if (special) offenders.set(ch, (offenders.get(ch) || 0) + 1);
        const last = segments[segments.length - 1];
        if (last && last.special === special) last.text += ch;
        else segments.push({ text: ch, special });
      }
      return { segments, offenders, isUtf: offenders.size > 0 };
    }

    function wlSmsApplySafeFixes(text) {
      let out = '';
      for (const ch of String(text ?? '')) {
        const info = WL_SMS_GSM7_SET.has(ch) ? null : WL_SMS_SPECIAL_CHAR_INFO[ch];
        out += (info && info.fix !== undefined) ? info.fix : ch;
      }
      return out;
    }

    function wlSmsNormalizeState(raw = {}) {
      const merged = { ...DEFAULT_WL_SMS_STATE, ...(raw || {}) };
      // testNumbers in user state is legacy — the live source is the shared
      // Firestore database. Keep this field so old caches still parse, but it
      // will normally be []. Migration in App() drains it into the shared doc.
      const testNumbers = Array.isArray(merged.testNumbers) ? merged.testNumbers.map(n => ({
        id: n?.id || wlSmsId('tn'),
        label: String(n?.label || ''),
        number: String(n?.number || ''),
        network: WL_SMS_NETWORKS[n?.network] ? n.network : 'smart',
      })) : [];
      const networks = Array.isArray(merged.networks)
        ? merged.networks.filter(n => WL_SMS_NETWORKS[n])
        : DEFAULT_WL_SMS_STATE.networks;
      // Don't filter selectedNumberIds against the legacy personal list — the
      // valid IDs now live in sharedTestNumbers and aren't visible here.
      // wlSmsEffectiveNumbers() does the final filtering at use-time.
      const selectedNumberIds = Array.isArray(merged.selectedNumberIds)
        ? merged.selectedNumberIds.filter(id => typeof id === 'string' && id)
        : [];
      const contents = Array.isArray(merged.contents)
        ? merged.contents.map(s => String(s ?? '')).filter(s => s.trim())
        : [];
      return {
        ...merged,
        rawPaste: String(merged.rawPaste || ''),
        senderId: String(merged.senderId || ''),
        sheetName: String(merged.sheetName || 'Sheet1'),
        linkBuilderRaw: String(merged.linkBuilderRaw || ''),
        notes: String(merged.notes || ''),
        separatorMode: merged.separatorMode === 'lines' ? 'lines' : 'auto',
        contents,
        testNumbers,
        networks: networks.length ? networks : ['smart', 'globe', 'dito'],
        selectedNumberIds,
      };
    }

    function wlSmsEffectiveNumbers(wl, sharedTestNumbers) {
      const networks = new Set(wl.networks || []);
      const selected = new Set(wl.selectedNumberIds || []);
      // sharedTestNumbers is the live cloud database. Fall back to whatever is
      // in the (legacy) wl.testNumbers when no shared list has been provided,
      // so this still works in transient states (before sign-in / Firebase load).
      const source = Array.isArray(sharedTestNumbers) && sharedTestNumbers.length
        ? sharedTestNumbers
        : (wl.testNumbers || []);
      return source.filter(n => selected.has(n.id) && networks.has(n.network));
    }

    function buildWlSmsSheet(ws, wl, sharedTestNumbers) {
      const contents = (wl.contents || []).map(s => String(s ?? '')).filter(s => s.trim());
      const numbers = wlSmsEffectiveNumbers(wl, sharedTestNumbers);

      ws.columns = [
        { key: 'number',  width: 16 },
        { key: 'content', width: 110 },
      ];

      const headerRow = ws.getRow(1);
      headerRow.getCell(1).value = 'Number';
      headerRow.getCell(2).value = 'Content';
      headerRow.font = { bold: true };
      headerRow.alignment = { vertical: 'middle' };

      let rowIdx = 2;
      for (const num of numbers) {
        const cleanNum = String(num.number || '').replace(/\D/g, '');
        const numericNumber = cleanNum && /^\d+$/.test(cleanNum) ? Number(cleanNum) : (num.number || '');
        for (const content of contents) {
          const row = ws.getRow(rowIdx);
          row.getCell(1).value = numericNumber;
          row.getCell(1).numFmt = '0';
          row.getCell(2).value = content;
          row.getCell(2).alignment = { wrapText: true, vertical: 'top' };
          row.height = 57.6;
          rowIdx++;
        }
      }

      ws.views = [{ state: 'frozen', ySplit: 1 }];
    }

    async function generateWlSmsWorkbook(state, sharedTestNumbers) {
      const wl = wlSmsNormalizeState(state.whitelistSms);
      const contents = (wl.contents || []).filter(s => s && s.trim());
      const numbers = wlSmsEffectiveNumbers(wl, sharedTestNumbers);
      if (!contents.length) throw new Error('No content blocks to whitelist. Paste content first.');
      if (!numbers.length) throw new Error('No test numbers selected for the chosen network(s).');

      const wb = new ExcelJS.Workbook();
      wb.creator = 'Whitelist SMS Generator';
      wb.created = new Date();
      const sheetName = safeSheetName(wl.sheetName || 'Sheet1');
      const ws = wb.addWorksheet(sheetName);
      buildWlSmsSheet(ws, wl, sharedTestNumbers);

      const buf = await wb.xlsx.writeBuffer();
      const senderPart = (wl.senderId || 'sender').replace(/[^A-Za-z0-9_-]+/g, '_').slice(0, 40) || 'sender';
      const filename = `Whitelist_SMS_${senderPart}_${bmrTodayString()}.xlsx`;
      saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);
      return filename;
    }

    /* ============================================================
       UI Primitives
       ============================================================ */
    function Pill({ children, tone = 'default', onClick, className = '' }) {
      const tones = {
        default: 'border-neutral-800 bg-neutral-900/80 text-neutral-300',
        accent:  'border-blue-500/40 bg-blue-500/10 text-blue-300',
        muted:   'border-neutral-900 bg-transparent text-neutral-600',
        success: 'border-emerald-500/40 bg-emerald-500/10 text-emerald-300',
      };
      return (
        <span onClick={onClick}
          className={`inline-flex items-center gap-1.5 rounded-full border px-2.5 py-0.5 text-[11px] font-medium tracking-wide ${tones[tone]} ${onClick ? 'cursor-pointer hover:bg-neutral-800' : ''} ${className}`}>
          {children}
        </span>
      );
    }

    function Btn({ children, onClick, variant = 'primary', size = 'md', disabled, className = '', type = 'button', title }) {
      const variants = {
        primary: 'bg-white text-black hover:bg-neutral-200',
        ghost:   'border border-neutral-800 bg-transparent text-neutral-200 hover:bg-neutral-900 hover:border-neutral-700',
        danger:  'border border-red-500/30 bg-red-500/5 text-red-300 hover:bg-red-500/15',
        accent:  'bg-blue-500 text-white hover:bg-blue-400',
      };
      const sizes = { sm: 'px-2.5 py-1 text-xs', md: 'px-3.5 py-2 text-sm', lg: 'px-5 py-2.5 text-sm' };
      return (
        <button type={type} title={title} onClick={onClick} disabled={disabled}
          className={`inline-flex items-center justify-center gap-2 rounded-md font-medium transition-colors disabled:opacity-40 disabled:cursor-not-allowed ${variants[variant]} ${sizes[size]} ${className}`}>
          {children}
        </button>
      );
    }

    function SectionLabel({ children, hint }) {
      return (
        <div className="flex items-baseline justify-between mb-3 gap-4">
          <h3 className="text-[10px] font-semibold uppercase tracking-[0.18em] text-neutral-500">{children}</h3>
          {hint && <span className="text-[10px] text-neutral-600">{hint}</span>}
        </div>
      );
    }

    // UI-only preferences (which editor sections are collapsed). Stored under a
    // SEPARATE localStorage key so it never touches app data (mrg_state_v1),
    // templates, or Firestore sync — collapsing a section can't affect saves.
    const UI_PREFS_KEY = 'mrg_ui_v1';
    function readUiPrefs() {
      try { return JSON.parse(localStorage.getItem(UI_PREFS_KEY)) || {}; } catch { return {}; }
    }
    function setSectionCollapsed(sid, collapsed) {
      try {
        const p = readUiPrefs();
        p.collapsed = p.collapsed || {};
        if (collapsed) p.collapsed[sid] = 1; else delete p.collapsed[sid];
        localStorage.setItem(UI_PREFS_KEY, JSON.stringify(p));
      } catch {}
    }

    // A SectionLabel whose content can collapse. `sid` remembers the open/closed
    // choice across sessions and sheet switches. `defaultCollapsed` only applies
    // until the user clicks once.
    function CollapsibleSection({ sid, title, hint, defaultCollapsed = false, children }) {
      const [collapsed, setCollapsed] = useState(() => {
        const c = readUiPrefs().collapsed || {};
        return sid in c ? !!c[sid] : defaultCollapsed;
      });
      const toggle = () => setCollapsed(v => { const n = !v; setSectionCollapsed(sid, n); return n; });
      return (
        <div>
          <button type="button" onClick={toggle} aria-expanded={!collapsed}
            className="no-press group flex w-full items-baseline justify-between gap-4 mb-3 text-left">
            <span className="flex items-center gap-1.5 min-w-0">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"
                className={`w-3 h-3 shrink-0 text-neutral-600 group-hover:text-neutral-400 transition-transform ${collapsed ? '-rotate-90' : ''}`}>
                <polyline points="6 9 12 15 18 9" />
              </svg>
              <span className="text-[10px] font-semibold uppercase tracking-[0.18em] text-neutral-500 group-hover:text-neutral-300 transition-colors truncate">{title}</span>
            </span>
            {hint && <span className="text-[10px] text-neutral-600 shrink-0">{hint}</span>}
          </button>
          {!collapsed && children}
        </div>
      );
    }

    function Card({ children, className = '' }) {
      return <div className={`rounded-lg border border-neutral-900 bg-[#232327] ${className}`}>{children}</div>;
    }

    function IconSheet() {
      return <svg className="w-3.5 h-3.5 text-emerald-500 shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/></svg>;
    }

    // Compact Google Sheets sync panel shown above the Generate button in the
    // SIP/FCS, BMR VOIP and BMR SMS sidebars. Connect once, then Generate (or
    // "Sync now") pushes the generated workbook into that module's Google Sheet.
    function GoogleSheetSync({ gsheets, moduleId, sheetId }) {
      const [draftId, setDraftId] = useState('');
      const [editing, setEditing] = useState(false);
      if (!gsheets) return null;
      const { conn, connect, disconnect, sync, sheetUrl, configured, clientId, onSetClientId } = gsheets;
      const busy = conn.busyModule === moduleId;
      const connecting = conn.busyModule === '__connect__';
      const last = conn.lastSync?.[moduleId];
      const url = sheetId ? sheetUrl(sheetId) : null;

      const saveId = () => { if (draftId.trim()) { onSetClientId(draftId); setEditing(false); } };

      // Client ID entry form — shown when no ID is set, or when changing it.
      if (!configured || editing) {
        return (
          <div className="rounded-md border border-neutral-900 bg-neutral-950 px-3 py-2.5 space-y-2">
            <div className="flex items-center gap-2 text-[11px]">
              <IconSheet />
              <span className="font-medium text-neutral-300">Google Sheets sync</span>
            </div>
            <p className="text-[10px] text-neutral-600 leading-relaxed">
              Paste your Google OAuth <span className="text-neutral-400">Client ID</span> (Web application) to enable one-click sync to Google Sheets.
            </p>
            <Input value={draftId} onChange={e => setDraftId(e.target.value)}
              onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); saveId(); } }}
              placeholder="xxxx.apps.googleusercontent.com"
              className="text-[11px] py-1.5 font-mono" />
            <div className="flex items-center gap-1.5">
              <Btn variant="ghost" size="sm" onClick={saveId} disabled={!draftId.trim()} className="flex-1">Save Client ID</Btn>
              {configured && <Btn variant="ghost" size="sm" onClick={() => { setEditing(false); setDraftId(''); }}>Cancel</Btn>}
            </div>
          </div>
        );
      }

      return (
        <div className="rounded-md border border-neutral-900 bg-neutral-950 px-3 py-2.5 space-y-2">
          <div className="flex items-center gap-2 text-[11px]">
            <IconSheet />
            <span className="font-medium text-neutral-300">Google Sheets</span>
            {conn.connected
              ? <span className="ml-auto inline-flex items-center gap-1 text-emerald-400"><span className="w-1.5 h-1.5 rounded-full bg-emerald-400"></span>Connected</span>
              : <span className="ml-auto text-neutral-500">Not connected</span>}
          </div>

          {conn.connected && conn.email && (
            <div className="text-[10px] text-neutral-500 truncate" title={conn.email}>{conn.email}</div>
          )}

          {!conn.connected ? (
            <Btn variant="ghost" size="sm" onClick={connect} disabled={connecting} className="w-full">
              {connecting ? <><span className="loader"></span> Connecting…</> : <>Connect Google Drive</>}
            </Btn>
          ) : (
            <div className="space-y-1.5">
              <Btn variant="ghost" size="sm" onClick={() => sync(moduleId, { interactive: true, confirmOverwrite: true })} disabled={busy} className="w-full">
                {busy ? <><span className="loader"></span> Syncing…</> : <>Sync now</>}
              </Btn>
              <div className="flex items-center justify-between text-[10px]">
                {url
                  ? <a href={url} target="_blank" rel="noopener noreferrer" className="text-blue-400 hover:text-blue-300 transition-colors">Open sheet ↗</a>
                  : <span className="text-neutral-600">No sheet yet</span>}
                <button onClick={disconnect} className="text-neutral-500 hover:text-neutral-300 transition-colors">Disconnect</button>
              </div>
              {last?.at && (
                <div className="text-[10px] text-neutral-600">Synced {new Date(last.at).toLocaleTimeString()}</div>
              )}
            </div>
          )}

          <button onClick={() => { setDraftId(clientId || ''); setEditing(true); }}
            className="text-[10px] text-neutral-600 hover:text-neutral-400 transition-colors">Change Client ID</button>
        </div>
      );
    }

    function Input(props) {
      const { className = '', ...rest } = props;
      return <input {...rest} className={`w-full rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2 text-sm text-neutral-100 placeholder-neutral-600 outline-none focus:border-blue-500/60 focus:ring-1 focus:ring-blue-500/30 transition-colors ${className}`} />;
    }

    function ClearableInput({ value, onClear, clearTitle = 'Clear input', className = '', inputClassName = '', ...rest }) {
      return (
        <div className={`relative ${className}`}>
          <Input {...rest} value={value} className={`pr-9 ${inputClassName}`} />
          {!!value && (
            <button type="button" onMouseDown={e => e.preventDefault()} onClick={onClear}
              className="no-press input-clear-button absolute right-2 top-1/2 inline-flex h-6 w-6 -translate-y-1/2 items-center justify-center rounded-full text-neutral-500 hover:text-neutral-100 focus:outline-none focus-visible:ring-1 focus-visible:ring-blue-500/40"
              title={clearTitle} aria-label={clearTitle}>
              <IconX />
            </button>
          )}
        </div>
      );
    }

    function Select({ children, className = '', ...rest }) {
      return <select {...rest} className={`w-full rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2 text-sm text-neutral-100 outline-none focus:border-blue-500/60 focus:ring-1 focus:ring-blue-500/30 transition-colors ${className}`}>{children}</select>;
    }

    function Textarea(props) {
      const { className = '', ...rest } = props;
      return <textarea {...rest} className={`w-full rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2 text-sm text-neutral-100 placeholder-neutral-600 outline-none focus:border-blue-500/60 focus:ring-1 focus:ring-blue-500/30 transition-colors resize-y ${className}`} />;
    }

    function IconUp() { return <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="18 15 12 9 6 15"/></svg>; }
    function IconDown() { return <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="6 9 12 15 18 9"/></svg>; }
    function IconReset() { return <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><polyline points="3 3 3 8 8 8"/></svg>; }
    function IconUndo() { return <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.3" strokeLinecap="round" strokeLinejoin="round"><path d="M9 14 4 9l5-5"/><path d="M4 9h10a6 6 0 0 1 0 12h-1"/></svg>; }
    function IconGrip() { return <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="currentColor"><circle cx="9" cy="5" r="1.4"/><circle cx="15" cy="5" r="1.4"/><circle cx="9" cy="12" r="1.4"/><circle cx="15" cy="12" r="1.4"/><circle cx="9" cy="19" r="1.4"/><circle cx="15" cy="19" r="1.4"/></svg>; }
    function IconX() { return <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>; }
    function IconPlus() { return <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>; }
    function IconDownload() { return <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>; }
    function IconUpload() { return <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>; }
    function IconSun() { return <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41"/></svg>; }
    function IconMoon() { return <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>; }
    function IconEye() { return <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7S1 12 1 12z"/><circle cx="12" cy="12" r="3"/></svg>; }
    function IconEyeOff() { return <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><path d="M17.94 17.94A10.94 10.94 0 0 1 12 19c-7 0-11-7-11-7a19.6 19.6 0 0 1 4.22-5.14"/><path d="M9.9 4.24A10.94 10.94 0 0 1 12 4c7 0 11 7 11 7a19.6 19.6 0 0 1-3.16 4.19"/><path d="M14.12 14.12a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></svg>; }

    const PRESET_COLOR_ROWS = [
      ['#000000','#434343','#666666','#999999','#B7B7B7','#CCCCCC','#D9D9D9','#EFEFEF','#F3F3F3','#FFFFFF'],
      ['#980000','#FF0000','#FF9900','#FFFF00','#00FF00','#00FFFF','#4A86E8','#0000FF','#9900FF','#FF00FF'],
      ['#E6B8AF','#F4CCCC','#FCE5CD','#FFF2CC','#D9EAD3','#D0E0E3','#C9DAF8','#CFE2F3','#D9D2E9','#EAD1DC'],
      ['#DD7E6B','#EA9999','#F9CB9C','#FFE599','#B6D7A8','#A2C4C9','#A4C2F4','#9FC5E8','#B4A7D6','#D5A6BD'],
      ['#CC4125','#E06666','#F6B26B','#FFD966','#93C47D','#76A5AF','#6D9EEB','#6FA8DC','#8E7CC3','#C27BA0'],
      ['#A61C00','#CC0000','#E69138','#F1C232','#6AA84F','#45818E','#3C78D8','#3D85C6','#674EA7','#A64D79'],
      ['#85200C','#990000','#B45F06','#BF9000','#38761D','#134F5C','#1155CC','#0B5394','#351C75','#741B47'],
      ['#5B0F00','#660000','#783F04','#7F6000','#274E13','#0C343D','#1C4587','#073763','#20124D','#4C1130'],
    ];

    function presetColorHex(value, fallback = '#000000') {
      const raw = String(value || fallback).replace('#', '').slice(-6).toUpperCase();
      return /^[0-9A-F]{6}$/.test(raw) ? '#' + raw : fallback;
    }

    function PresetColorPicker({ value, onChange, title = 'Choose color', showHex = true, className = '', buttonClassName = '' }) {
      const [open, setOpen] = useState(false);
      const [popupPos, setPopupPos] = useState(null);
      const rootRef = useRef(null);
      const buttonRef = useRef(null);
      const popupRef = useRef(null);
      const selected = presetColorHex(value);

      const computePopupPos = useCallback(() => {
        const btn = buttonRef.current;
        if (!btn) return null;
        const rect = btn.getBoundingClientRect();
        const POPUP_W = 254;
        const POPUP_H = 220;
        const margin = 8;
        const gap = 4;
        const spaceBelow = window.innerHeight - rect.bottom;
        const spaceAbove = rect.top;
        const above = spaceBelow < POPUP_H + margin && spaceAbove > spaceBelow;
        const top = above ? Math.max(margin, rect.top - POPUP_H - gap) : rect.bottom + gap;
        let left = rect.left;
        if (left + POPUP_W + margin > window.innerWidth) left = window.innerWidth - POPUP_W - margin;
        if (left < margin) left = margin;
        return { top, left };
      }, []);

      const togglePicker = () => {
        setOpen(prev => {
          const next = !prev;
          if (next) setPopupPos(computePopupPos());
          return next;
        });
      };

      useEffect(() => {
        if (!open) return;
        const isInside = (target) =>
          rootRef.current?.contains(target) || popupRef.current?.contains(target);
        const closeOutside = (e) => { if (!isInside(e.target)) setOpen(false); };
        const closeOnEscape = (e) => { if (e.key === 'Escape') setOpen(false); };
        window.addEventListener('pointerdown', closeOutside);
        window.addEventListener('keydown', closeOnEscape);
        return () => {
          window.removeEventListener('pointerdown', closeOutside);
          window.removeEventListener('keydown', closeOnEscape);
        };
      }, [open]);

      return (
        <span ref={rootRef} className={`relative inline-flex min-w-0 ${className}`}>
          <button ref={buttonRef} type="button" onClick={togglePicker} title={title} aria-label={title} aria-expanded={open}
            className={`inline-flex min-w-0 items-center gap-2 rounded-md border border-neutral-800 bg-neutral-950 px-2 py-1.5 text-neutral-300 hover:border-neutral-700 hover:bg-neutral-900 transition-colors ${showHex ? 'w-full' : ''} ${buttonClassName}`}>
            <span className="h-5 w-5 shrink-0 rounded border border-neutral-700" style={{ background: selected }}></span>
            {showHex && <span className="min-w-0 flex-1 truncate text-left text-[11px] font-mono text-neutral-400">{selected}</span>}
            <IconDown />
          </button>
          {open && popupPos && ReactDOM.createPortal(
            <span
              ref={popupRef}
              style={{ position: 'fixed', top: popupPos.top, left: popupPos.left, zIndex: 50 }}
              className="w-[254px] rounded-md border border-neutral-800 bg-[#232327] p-2 shadow-2xl">
              <span className="block space-y-1">
                {PRESET_COLOR_ROWS.map((row, rowIndex) => (
                  <span key={rowIndex} className="grid grid-cols-10 gap-1">
                    {row.map(color => (
                      <button key={color} type="button" onClick={() => { onChange(color); setOpen(false); }}
                        title={color} aria-label={`Use color ${color}`}
                        className={`h-5 w-5 rounded border transition-transform hover:scale-110 ${selected === color ? 'border-blue-300 ring-2 ring-blue-400/60' : 'border-neutral-700'}`}
                        style={{ background: color }}></button>
                    ))}
                  </span>
                ))}
              </span>
            </span>,
            document.body
          )}
        </span>
      );
    }

    /* ============================================================
       Sidebar
       ============================================================ */
    function SyncBadge({ sync, onRetry }) {
      const map = {
        connecting:  { dot: 'bg-amber-400 animate-pulse',   text: 'Connecting…', color: 'text-amber-300' },
        loading:     { dot: 'bg-blue-400 animate-pulse',    text: 'Loading…',    color: 'text-blue-300' },
        saving:      { dot: 'bg-blue-400 animate-pulse',    text: 'Saving…',     color: 'text-blue-300' },
        synced:      { dot: 'bg-emerald-500',               text: 'Synced',      color: 'text-emerald-300' },
        offline:     { dot: 'bg-neutral-500',               text: sync?.uid ? 'Local only' : 'Offline', color: 'text-neutral-400' },
        'signed-out':{ dot: 'bg-neutral-500',               text: 'Signed out',  color: 'text-neutral-400' },
      };
      const m = map[sync?.status] || map.offline;
      const title = sync?.message || (sync?.email ? `${sync.email}\nUID: ${sync.uid}` : 'Not signed in');
      return (
        <div className="flex items-center gap-2 mb-2" title={title}>
          <div className={`w-1.5 h-1.5 rounded-full ${m.dot}`}></div>
          <span className={`text-[10px] uppercase tracking-[0.2em] ${m.color}`}>{m.text}</span>
          {sync?.status === 'offline' && sync?.uid && onRetry && (
            <button type="button" onClick={onRetry}
              className="no-press rounded border border-neutral-800 px-1.5 py-0.5 text-[9px] uppercase tracking-wide text-neutral-500 hover:border-neutral-700 hover:text-neutral-200"
              title={sync?.message || 'Retry cloud sync'}>
              Retry
            </button>
          )}
        </div>
      );
    }

    function AccountChip({ sync }) {
      if (!sync?.email) return null;
      const onSignOut = async () => {
        if (!window.__fb || !window.__fbm) return;
        const ok = await confirmDialog({
          title: 'Sign out?',
          message: 'Your data stays safely in the cloud and will be there next time you sign in.',
          confirmText: 'Sign out',
          tone: 'danger',
        });
        if (!ok) return;
        window.__fbm.signOut(window.__fb.auth).catch(() => {});
      };
      return (
        <div className="mt-3 flex items-center justify-between gap-2 rounded-md border border-neutral-900 bg-neutral-950 px-2.5 py-1.5">
          <span className="text-[11px] text-neutral-400 truncate" title={sync.email}>{sync.email}</span>
          <button onClick={onSignOut} title="Sign out"
            className="text-[10px] uppercase tracking-wide text-neutral-500 hover:text-red-300 transition-colors shrink-0">
            Sign out
          </button>
        </div>
      );
    }

    function humanizeAuthError(e) {
      const code = e?.code || '';
      if (code === 'auth/invalid-email')           return 'Invalid email address.';
      if (code === 'auth/missing-password')        return 'Please enter a password.';
      if (code === 'auth/invalid-credential' ||
          code === 'auth/wrong-password')          return 'Email or password is incorrect.';
      if (code === 'auth/user-not-found')          return 'No account found for that email.';
      if (code === 'auth/email-already-in-use')    return 'An account with that email already exists.';
      if (code === 'auth/weak-password')           return 'Password must be at least 6 characters.';
      if (code === 'auth/too-many-requests')       return 'Too many attempts. Try again later.';
      if (code === 'auth/network-request-failed')  return 'Network error. Check your connection.';
      if (code === 'auth/operation-not-allowed')   return 'Email/password sign-in is disabled in Firebase Console.';
      return e?.message || 'Authentication failed.';
    }

    function syncErrorMessage(error, fallback = 'Cloud sync failed') {
      const code = error?.code ? ` (${error.code})` : '';
      const detail = error?.message ? `: ${error.message}` : '';
      return `${fallback}${code}${detail}`;
    }

    function Dialog({ dialog, onResolve }) {
      useEffect(() => {
        if (!dialog) return;
        const onKey = (e) => {
          if (e.key === 'Escape') { e.preventDefault(); onResolve(dialog.kind === 'confirm' ? false : undefined); }
          else if (e.key === 'Enter') { e.preventDefault(); onResolve(dialog.kind === 'confirm' ? true : undefined); }
        };
        window.addEventListener('keydown', onKey);
        return () => window.removeEventListener('keydown', onKey);
      }, [dialog, onResolve]);

      if (!dialog) return null;
      const { kind, title, message, confirmText = 'OK', cancelText = 'Cancel', tone = 'default' } = dialog;
      const cancelValue = kind === 'confirm' ? false : undefined;
      const okValue     = kind === 'confirm' ? true  : undefined;
      const okVariant   = tone === 'danger' ? 'danger' : 'primary';

      return (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4 anim-overlay"
          onClick={() => onResolve(cancelValue)}>
          <div
            className="w-full max-w-sm rounded-xl border border-neutral-800 bg-[#232327] p-6 shadow-2xl anim-panel"
            onClick={e => e.stopPropagation()}>
            {title && <h3 className="text-sm font-bold tracking-tight mb-1.5">{title}</h3>}
            {message && <p className="text-[12.5px] text-neutral-400 leading-relaxed">{message}</p>}
            <div className="flex justify-end gap-2 mt-5">
              {kind === 'confirm' && (
                <Btn variant="ghost" size="md" onClick={() => onResolve(cancelValue)}>{cancelText}</Btn>
              )}
              <Btn variant={okVariant} size="md" onClick={() => onResolve(okValue)}>{confirmText}</Btn>
            </div>
          </div>
        </div>
      );
    }

    function AuthModal() {
      const [mode, setMode] = useState('signin');
      const [email, setEmail] = useState('');
      const [password, setPassword] = useState('');
      const [busy, setBusy] = useState(false);
      const [error, setError] = useState(null);
      const [info, setInfo] = useState(null);

      const submit = async (e) => {
        e.preventDefault();
        if (!window.__fb || !window.__fbm) return;
        setBusy(true); setError(null); setInfo(null);
        try {
          const { auth } = window.__fb;
          const { signInWithEmailAndPassword, createUserWithEmailAndPassword } = window.__fbm;
          if (mode === 'signup') {
            await createUserWithEmailAndPassword(auth, email.trim(), password);
          } else {
            await signInWithEmailAndPassword(auth, email.trim(), password);
          }
        } catch (err) {
          setError(humanizeAuthError(err));
        } finally {
          setBusy(false);
        }
      };

      const resetPassword = async () => {
        if (!window.__fb || !window.__fbm) return;
        if (!email.trim()) { setError('Enter your email above first, then click "Forgot password".'); return; }
        setBusy(true); setError(null); setInfo(null);
        try {
          await window.__fbm.sendPasswordResetEmail(window.__fb.auth, email.trim());
          setInfo(`Password reset email sent to ${email.trim()}.`);
        } catch (err) {
          setError(humanizeAuthError(err));
        } finally {
          setBusy(false);
        }
      };

      return (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4 anim-overlay">
          <div className="w-full max-w-sm rounded-xl border border-neutral-800 bg-[#232327] p-6 shadow-2xl anim-panel">
            <div className="flex items-center gap-3 mb-1">
              <svg viewBox="0 0 512 512" className="w-9 h-9 rounded-md shrink-0" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                <rect width="512" height="512" rx="100" fill="#2196F3"/>
                <text x="256" y="370" textAnchor="middle" fontFamily="Inter, system-ui, sans-serif" fontSize="340" fontWeight="900" fill="white">G</text>
              </svg>
              <div>
                <h2 className="text-base font-bold tracking-tight">Monthly Report Generator</h2>
                <p className="text-[11px] text-neutral-500 mt-0.5">Sign in to sync across sessions</p>
              </div>
            </div>

            <div className="flex items-center gap-0.5 border border-neutral-900 rounded-md p-0.5 bg-[#1a1a1d] mt-5">
              {[['signin','Sign in'], ['signup','Create account']].map(([m, label]) => (
                <button key={m} type="button"
                  onClick={() => { setMode(m); setError(null); setInfo(null); }}
                  className={`flex-1 px-3 py-1.5 text-xs font-medium rounded transition-colors ${mode === m ? 'bg-neutral-800 text-neutral-100' : 'text-neutral-500 hover:text-neutral-200'}`}>
                  {label}
                </button>
              ))}
            </div>

            <form onSubmit={submit} className="mt-4 space-y-3">
              <div>
                <label className="block text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-1.5">Email</label>
                <Input type="email" autoComplete="email" required disabled={busy}
                  value={email} onChange={e => setEmail(e.target.value)}
                  placeholder="you@example.com" />
              </div>
              <div>
                <label className="block text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-1.5">Password</label>
                <Input type="password"
                  autoComplete={mode === 'signup' ? 'new-password' : 'current-password'}
                  required minLength={6} disabled={busy}
                  value={password} onChange={e => setPassword(e.target.value)}
                  placeholder={mode === 'signup' ? 'At least 6 characters' : 'Your password'} />
              </div>

              {error && (
                <div className="rounded-md border border-red-500/30 bg-red-500/5 px-3 py-2 text-[12px] text-red-300">
                  {error}
                </div>
              )}
              {info && (
                <div className="rounded-md border border-emerald-500/30 bg-emerald-500/5 px-3 py-2 text-[12px] text-emerald-300">
                  {info}
                </div>
              )}

              <Btn type="submit" variant="primary" size="lg" disabled={busy} className="w-full mt-1">
                {busy ? <><span className="loader"></span> Please wait…</> : (mode === 'signin' ? 'Sign in' : 'Create account')}
              </Btn>

              {mode === 'signin' && (
                <button type="button" onClick={resetPassword} disabled={busy}
                  className="block w-full text-center text-[11px] text-neutral-500 hover:text-neutral-300 transition-colors mt-1 disabled:opacity-40">
                  Forgot password?
                </button>
              )}
            </form>
          </div>
        </div>
      );
    }

    /* ============================================================
       Uploaded Sidebar — staging area for sheets parsed from .xlsx
       ============================================================ */
    function UploadedSidebar({ state, setState }) {
      const [collapsed, setCollapsed] = useState(true);
      const imported = state.imported || [];

      const removeOne = (id) =>
        setState(s => ({ ...s, imported: (s.imported || []).filter(x => x.id !== id) }));

      const clearAll = async () => {
        if (!imported.length) return;
        const ok = await confirmDialog({
          title: 'Clear uploaded sheets?',
          message: `${imported.length} imported sheet${imported.length === 1 ? '' : 's'} will be removed from this panel. Your main configuration is not affected.`,
          confirmText: 'Clear',
          tone: 'danger',
        });
        if (!ok) return;
        setState(s => ({ ...s, imported: [] }));
      };

      const addToMain = (item) => {
        setState(s => {
          const existingNames = new Set(s.sheets.map(x => x.name.toLowerCase()));
          let name = item.name;
          if (existingNames.has(name.toLowerCase())) {
            let n = 2;
            while (existingNames.has(`${item.name} (${n})`.toLowerCase())) n += 1;
            name = `${item.name} (${n})`;
          }
          const copy = { ...item, id: newId(), name, active: true };
          return { ...s, sheets: [...s.sheets, copy], selectedSheetId: copy.id };
        });
      };

      if (collapsed) {
        return (
          <aside className="w-10 shrink-0 border-r border-neutral-900 bg-[#17171a] h-screen sticky top-0 flex flex-col items-center py-4 gap-3">
            <button onClick={() => setCollapsed(false)}
              title="Expand uploaded sheets"
              className="text-neutral-500 hover:text-neutral-200 transition-colors">
              <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="9 18 15 12 9 6"/></svg>
            </button>
            <div className="[writing-mode:vertical-rl] rotate-180 text-[10px] uppercase tracking-[0.2em] text-neutral-500 mt-2">
              Uploaded {imported.length ? `· ${imported.length}` : ''}
            </div>
          </aside>
        );
      }

      return (
        <aside className="w-[280px] shrink-0 border-r border-neutral-900 bg-[#17171a] h-screen sticky top-0 flex flex-col">
          <div className="p-5 border-b border-neutral-900">
            <div className="flex items-center justify-between gap-2">
              <SectionLabel hint={imported.length ? `${imported.length} sheet${imported.length === 1 ? '' : 's'}` : ''}>Uploaded</SectionLabel>
              <button onClick={() => setCollapsed(true)}
                title="Collapse"
                className="text-neutral-500 hover:text-neutral-200 transition-colors -mt-2">
                <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="15 18 9 12 15 6"/></svg>
              </button>
            </div>
            <p className="text-[11px] text-neutral-500 -mt-1">Sheets parsed from an uploaded .xlsx. Add the ones you want into the main sidebar.</p>
          </div>

          <div className="p-5 flex-1 overflow-y-auto min-h-0">
            {imported.length === 0 ? (
              <div className="rounded-lg border border-dashed border-neutral-800 px-4 py-10 text-center">
                <p className="text-[12px] text-neutral-400">No uploaded sheets yet</p>
                <p className="text-[11px] text-neutral-600 mt-1.5 leading-relaxed">Open the <span className="text-neutral-300">Config</span> tab and use <span className="text-neutral-300">Import .xlsx</span> in the Templates section.</p>
              </div>
            ) : (
              <div className="space-y-1.5">
                {imported.map((it) => {
                  const count = it.layout === 'hourly'
                    ? `${(it.metrics || []).length} metrics`
                    : `${(it.columns || []).length} cols`;
                  return (
                    <div key={it.id}
                      className="group rounded-md border border-neutral-900 bg-neutral-950 px-2.5 py-2 hover:border-neutral-800 transition-colors">
                      <div className="flex items-center gap-2 min-w-0">
                        <div className="flex-1 min-w-0">
                          <div className="text-sm font-medium truncate text-neutral-100" title={it.name}>{it.name}</div>
                          <div className="text-[10px] text-neutral-500 mt-0.5 uppercase tracking-wide">{LAYOUT_LABELS[it.layout]} · {count}</div>
                        </div>
                        <button onClick={() => removeOne(it.id)}
                          title="Remove from uploaded"
                          className="p-1.5 text-neutral-600 hover:text-red-400 opacity-0 group-hover:opacity-100 transition-opacity shrink-0">
                          <IconX />
                        </button>
                      </div>
                      <Btn variant="ghost" size="sm" onClick={() => addToMain(it)} className="w-full mt-2">
                        <IconPlus /> Add to main
                      </Btn>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {imported.length > 0 && (
            <div className="p-5 border-t border-neutral-900">
              <Btn variant="ghost" size="sm" onClick={clearAll} className="w-full">Clear all</Btn>
            </div>
          )}
        </aside>
      );
    }

    function Sidebar({ state, setState, onGenerate, busy, sync, onRetrySync, gsheets }) {
      const { year, month, sheets, selectedSheetId } = state;
      const theme = state.theme || 'dark';
      const toggleTheme = () => setState(s => ({ ...s, theme: (s.theme || 'dark') === 'dark' ? 'light' : 'dark' }));

      const updateSheet = (id, patch) => setState(s => ({ ...s, sheets: s.sheets.map(x => x.id === id ? { ...x, ...patch } : x) }));
      const addSheet = () => {
        const id = newId();
        setState(s => ({ ...s, sheets: [...s.sheets, { id, name: 'New Sheet', layout: 'hourly', active: true, metrics: ['Metric 1'], rowSeparator: true, hourStart: 0, hourEnd: 24, note: '' }], selectedSheetId: id }));
      };
      const deleteSheet = async (sheet) => {
        const ok = await confirmDialog({
          title: 'Delete sheet?',
          message: `"${sheet.name}" and all of its configuration will be removed. This can't be undone.`,
          confirmText: 'Delete',
          tone: 'danger',
        });
        if (!ok) return;
        setState(s => {
          const remaining = s.sheets.filter(x => x.id !== sheet.id);
          const selectedStillExists = remaining.some(x => x.id === s.selectedSheetId);
          return { ...s, sheets: remaining, selectedSheetId: selectedStillExists ? s.selectedSheetId : (remaining[0]?.id || null) };
        });
      };
      const reorderSheets = (from, to) => setState(s => {
        if (from === to || from < 0 || to < 0 || from >= s.sheets.length || to >= s.sheets.length) return s;
        const arr = [...s.sheets];
        const [moved] = arr.splice(from, 1);
        arr.splice(to, 0, moved);
        return { ...s, sheets: arr };
      });
      const [dragIndex, setDragIndex] = useState(null);
      const [overIndex, setOverIndex] = useState(null);
      const [dragRow, setDragRow] = useState(null);

      return (
        <aside className="w-[320px] shrink-0 border-r border-neutral-900 bg-[#17171a] h-screen sticky top-0 flex flex-col">
          {/* Brand */}
          <div className="p-5 border-b border-neutral-900">
            <div className="flex items-center justify-between gap-2">
              <SyncBadge sync={sync} onRetry={onRetrySync} />
              <button onClick={toggleTheme}
                title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}
                className="inline-flex items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2 py-1 text-[10px] uppercase tracking-wide text-neutral-400 hover:text-neutral-100 hover:border-neutral-700 transition-colors">
                {theme === 'dark' ? <IconSun /> : <IconMoon />}
                {theme === 'dark' ? 'Light' : 'Dark'}
              </button>
            </div>
            <h1 className="text-[17px] font-bold tracking-tight leading-tight mt-2">Monthly Report Generator</h1>
            <p className="text-xs text-neutral-500 mt-1">24/7 ops · SIP/FCS hourly record</p>
            <AccountChip sync={sync} />
          </div>

          {/* Period */}
          <div className="p-5 border-b border-neutral-900">
            <SectionLabel>Period</SectionLabel>
            <div className="grid grid-cols-3 gap-2">
              <Select value={month} onChange={e => setState(s => ({ ...s, month: Number(e.target.value) }))} className="col-span-2">
                {MONTHS.map((m, i) => <option key={i} value={i}>{m}</option>)}
              </Select>
              <Input type="number" value={year} onChange={e => setState(s => ({ ...s, year: Number(e.target.value) }))} />
            </div>
            <p className="text-[11px] text-neutral-500 mt-2 font-mono">{daysInMonth(year, month)} days · 24h cycle</p>
          </div>

          {/* Sheet list */}
          <div className="p-5 flex-1 overflow-y-auto min-h-0">
            <SectionLabel hint={`${sheets.filter(s => s.active).length}/${sheets.length} active`}>Sheets</SectionLabel>
            <div className="space-y-0.5">
              {sheets.map((s, i) => {
                const isDragging = dragIndex === i;
                const showDropAbove = overIndex === i && dragIndex !== null && dragIndex !== i && dragIndex > i;
                const showDropBelow = overIndex === i && dragIndex !== null && dragIndex !== i && dragIndex < i;
                return (
                <div key={s.id}
                  draggable={dragRow === i}
                  onDragStart={(e) => {
                    setDragIndex(i);
                    e.dataTransfer.effectAllowed = 'move';
                    try { e.dataTransfer.setData('text/plain', String(i)); } catch (_) {}
                  }}
                  onDragOver={(e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; if (overIndex !== i) setOverIndex(i); }}
                  onDragLeave={() => { if (overIndex === i) setOverIndex(null); }}
                  onDrop={(e) => {
                    e.preventDefault();
                    if (dragIndex !== null && dragIndex !== i) reorderSheets(dragIndex, i);
                    setDragIndex(null); setOverIndex(null); setDragRow(null);
                  }}
                  onDragEnd={() => { setDragIndex(null); setOverIndex(null); setDragRow(null); }}
                  onClick={() => setState(st => ({ ...st, selectedSheetId: s.id }))}
                  className={`group flex items-center gap-2.5 rounded-md border px-2.5 py-2 cursor-pointer transition-colors ${selectedSheetId === s.id ? 'border-neutral-700 bg-neutral-900' : 'border-transparent hover:bg-neutral-900/60'} ${isDragging ? 'opacity-30' : ''} ${showDropAbove ? 'border-t-2 border-t-blue-500/70' : ''} ${showDropBelow ? 'border-b-2 border-b-blue-500/70' : ''}`}>
                  <span
                    onMouseDown={(e) => { e.stopPropagation(); setDragRow(i); }}
                    onMouseUp={() => setDragRow(null)}
                    onTouchStart={(e) => { e.stopPropagation(); setDragRow(i); }}
                    onTouchEnd={() => setDragRow(null)}
                    onClick={e => e.stopPropagation()}
                    className="cursor-grab active:cursor-grabbing p-0.5 -ml-0.5 text-neutral-700 hover:text-neutral-300 opacity-0 group-hover:opacity-100 transition-all select-none shrink-0"
                    title="Drag to reorder">
                    <IconGrip />
                  </span>
                  <input type="checkbox" checked={s.active}
                    onChange={(e) => updateSheet(s.id, { active: e.target.checked })}
                    onClick={e => e.stopPropagation()}
                    className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500 focus:ring-blue-500/30 cursor-pointer" />
                  <div className="flex-1 min-w-0">
                    <div className={`text-sm font-medium truncate ${s.active ? 'text-neutral-100' : 'text-neutral-500 line-through'}`}>{s.name}</div>
                    <div className="text-[10px] text-neutral-500 mt-0.5 uppercase tracking-wide">{LAYOUT_LABELS[s.layout]}</div>
                  </div>
                  <button onClick={(e) => { e.stopPropagation(); deleteSheet(s); }}
                    className="shrink-0 rounded p-1 text-neutral-500 opacity-0 transition-all hover:text-red-400 group-hover:opacity-100 focus:opacity-100"
                    title="Delete sheet">
                    <IconX />
                  </button>
                </div>
                );
              })}
            </div>
            <Btn variant="ghost" size="sm" onClick={addSheet} className="w-full mt-3"><IconPlus /> Add sheet</Btn>
          </div>

          {/* Generate */}
          <div className="p-5 border-t border-neutral-900 space-y-3">
            <label className="flex items-center gap-2 text-xs text-neutral-400 cursor-pointer">
              <input type="checkbox" checked={state.includeIndex}
                onChange={e => setState(s => ({ ...s, includeIndex: e.target.checked }))}
                className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500 focus:ring-blue-500/30 cursor-pointer" />
              Include INDEX tab
            </label>
            <GoogleSheetSync gsheets={gsheets} moduleId="sip_fcs" sheetId={state.googleSheets?.sheetIds?.sip_fcs} />
            <Btn variant="primary" size="lg" onClick={onGenerate} disabled={busy} className="w-full">
              {busy ? <><span className="loader"></span> Generating…</> : <><IconDownload /> Generate {MONTHS[month]} {year}</>}
            </Btn>
            <p className="text-[10px] text-neutral-600 text-center font-mono">.xlsx · Google Sheets compatible</p>
          </div>
        </aside>
      );
    }

    /* ============================================================
       Inline list editor
       ============================================================ */
    function ListEditor({ items, onChange, placeholder = 'New item', label = 'item', withDropdown = false, withColor = false }) {
      const [draft, setDraft] = useState('');
      const [dragIndex, setDragIndex] = useState(null);
      const [overIndex, setOverIndex] = useState(null);
      const [dragRow, setDragRow] = useState(null); // index that is currently allowed to start a drag
      const [expandedDropdown, setExpandedDropdown] = useState(null); // index with dropdown editor open
      const add = () => { if (!draft.trim()) return; onChange([...items, draft.trim()]); setDraft(''); };
      const setLabel = (i, v) => onChange(items.map((x, j) => {
        if (j !== i) return x;
        if (typeof x === 'string') return v;
        return { ...x, label: v };
      }));
      const toggleHidden = (i) => onChange(items.map((x, j) => {
        if (j !== i) return x;
        const dd = itemDropdown(x);
        const color = itemColor(x);
        const colorExtend = itemColorExtend(x);
        if (itemHidden(x)) {
          // unhide
          if (dd.length || color) return { label: itemLabel(x), ...(dd.length ? { dropdown: dd } : {}), ...(color ? { color } : {}), ...(color && colorExtend ? { colorExtend: true } : {}) };
          return itemLabel(x);
        }
        // hide
        return { label: itemLabel(x), hidden: true, ...(dd.length ? { dropdown: dd } : {}), ...(color ? { color } : {}), ...(color && colorExtend ? { colorExtend: true } : {}) };
      }));
      const setDropdown = (i, opts) => onChange(items.map((x, j) => {
        if (j !== i) return x;
        const lab = itemLabel(x);
        const hidden = itemHidden(x);
        const color = itemColor(x);
        const colorExtend = itemColorExtend(x);
        const clean = (opts || [])
          .map((opt) => {
            const label = String(dropdownOptionLabel(opt)).trim();
            const color = dropdownOptionColor(opt);
            if (!label) return null;
            return color ? { label, color } : label;
          })
          .filter(Boolean);
        if (!clean.length) {
          if (hidden || color) return { label: lab, ...(hidden ? { hidden: true } : {}), ...(color ? { color } : {}), ...(color && colorExtend ? { colorExtend: true } : {}) };
          return lab;
        }
        return { label: lab, ...(hidden ? { hidden: true } : {}), ...(color ? { color } : {}), ...(color && colorExtend ? { colorExtend: true } : {}), dropdown: clean };
      }));
      const setColor = (i, color) => onChange(items.map((x, j) => {
        if (j !== i) return x;
        return { ...(typeof x === 'string' ? { label: x } : x), color };
      }));
      const toggleColorExtend = (i) => onChange(items.map((x, j) => {
        if (j !== i) return x;
        const base = { ...(typeof x === 'string' ? { label: x } : x) };
        if (!itemColor(base)) return base;
        return { ...base, colorExtend: !itemColorExtend(base) };
      }));
      const resetColor = (i) => onChange(items.map((x, j) => {
        if (j !== i || typeof x === 'string') return x;
        const { color, colorExtend, ...rest } = x;
        if (!rest.hidden && !itemDropdown(rest).length) return itemLabel(rest);
        return rest;
      }));
      const remove = (i) => onChange(items.filter((_, j) => j !== i));
      const reorder = (from, to) => {
        if (from === to || to < 0 || to >= items.length) return;
        const arr = [...items];
        const [moved] = arr.splice(from, 1);
        arr.splice(to, 0, moved);
        onChange(arr);
      };
      return (
        <div className="space-y-1.5">
          {items.map((it, i) => {
            const hidden = itemHidden(it);
            const dropdown = itemDropdown(it);
            const isExpanded = expandedDropdown === i;
            const isDragging = dragIndex === i;
            const showDropAbove = overIndex === i && dragIndex !== null && dragIndex !== i && dragIndex > i;
            const showDropBelow = overIndex === i && dragIndex !== null && dragIndex !== i && dragIndex < i;
            return (
              <div key={i}>
                <div
                  draggable={dragRow === i}
                  onDragStart={(e) => {
                    setDragIndex(i);
                    e.dataTransfer.effectAllowed = 'move';
                    try { e.dataTransfer.setData('text/plain', String(i)); } catch (_) {}
                  }}
                  onDragOver={(e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; if (overIndex !== i) setOverIndex(i); }}
                  onDragLeave={() => { if (overIndex === i) setOverIndex(null); }}
                  onDrop={(e) => {
                    e.preventDefault();
                    if (dragIndex !== null && dragIndex !== i) reorder(dragIndex, i);
                    setDragIndex(null); setOverIndex(null); setDragRow(null);
                  }}
                  onDragEnd={() => { setDragIndex(null); setOverIndex(null); setDragRow(null); }}
                  className={`group flex items-center gap-2 rounded-md transition-all ${hidden ? 'opacity-50' : ''} ${isDragging ? 'opacity-30' : ''} ${showDropAbove ? 'border-t-2 border-blue-500/70' : ''} ${showDropBelow ? 'border-b-2 border-blue-500/70' : ''}`}>
                  <span
                    onMouseDown={() => setDragRow(i)}
                    onMouseUp={() => setDragRow(null)}
                    onTouchStart={() => setDragRow(i)}
                    onTouchEnd={() => setDragRow(null)}
                    className="cursor-grab active:cursor-grabbing p-1 text-neutral-700 hover:text-neutral-300 transition-colors select-none shrink-0"
                    title="Drag to reorder">
                    <IconGrip />
                  </span>
                  <span className="text-[10px] text-neutral-600 font-mono w-6 text-right shrink-0">{String(i + 1).padStart(2, '0')}</span>
                  <Input value={itemLabel(it)} onChange={e => setLabel(i, e.target.value)}
                    className={hidden ? 'line-through' : ''} />
                  {withColor && (
                    <span className="flex shrink-0 items-center gap-1">
                      <PresetColorPicker value={itemColor(it) || '#FFFFFF'} onChange={color => setColor(i, color)}
                        title="Choose metric cell color" showHex={false} buttonClassName="h-8 px-1.5 py-1" />
                      <button onClick={() => toggleColorExtend(i)}
                        disabled={!itemColor(it)}
                        title={itemColorExtend(it) ? 'Color extends across hourly cells' : 'Extend color across hourly cells'}
                        className={`h-8 rounded-md border px-2 text-[10px] font-semibold uppercase tracking-wide transition-colors disabled:cursor-not-allowed disabled:opacity-35 ${itemColorExtend(it) ? 'border-blue-500/40 bg-blue-500/10 text-blue-300' : 'border-neutral-800 bg-neutral-950 text-neutral-500 hover:border-neutral-700 hover:text-neutral-200'}`}>
                        Row
                      </button>
                      {itemColor(it) && (
                        <button onClick={() => resetColor(i)}
                          title="Reset metric cell color"
                          className="p-1 text-neutral-500 hover:text-neutral-200 transition-colors"><IconReset /></button>
                      )}
                    </span>
                  )}
                  {withDropdown && !hidden && (
                    <button onClick={() => setExpandedDropdown(isExpanded ? null : i)}
                      title={dropdown.length ? `${dropdown.length} dropdown option${dropdown.length === 1 ? '' : 's'} — click to edit` : 'Add dropdown options'}
                      className={`inline-flex items-center gap-1 shrink-0 px-2 h-7 rounded-md border text-[10px] font-medium uppercase tracking-wide transition-colors ${dropdown.length ? 'border-blue-500/40 bg-blue-500/10 text-blue-300 hover:bg-blue-500/15' : 'border-neutral-800 bg-neutral-950 text-neutral-500 hover:text-neutral-200 hover:border-neutral-700'}`}>
                      <IconDown />
                      {dropdown.length ? `List · ${dropdown.length}` : 'List'}
                    </button>
                  )}
                  <div className="flex shrink-0 opacity-0 group-hover:opacity-100 transition-opacity text-neutral-500">
                    <button onClick={() => toggleHidden(i)}
                      className={`p-1.5 ${hidden ? 'text-amber-400 hover:text-amber-300 opacity-100' : 'hover:text-neutral-200'}`}
                      title={hidden ? 'Show (include in output)' : 'Hide (exclude from output)'}>
                      {hidden ? <IconEyeOff /> : <IconEye />}
                    </button>
                    <button onClick={() => remove(i)} className="p-1.5 hover:text-red-400" title="Remove"><IconX /></button>
                  </div>
                </div>
                {withDropdown && isExpanded && !hidden && (
                  <DropdownOptionsEditor
                    options={dropdown}
                    onChange={(opts) => setDropdown(i, opts)}
                    columnName={itemLabel(it)} />
                )}
              </div>
            );
          })}
          <div className="flex gap-2 pt-2">
            <Input value={draft} onChange={e => setDraft(e.target.value)} onKeyDown={e => e.key === 'Enter' && (e.preventDefault(), add())} placeholder={placeholder} />
            <Btn variant="ghost" size="md" onClick={add}>Add {label}</Btn>
          </div>
        </div>
      );
    }

    function DropdownOptionsEditor({ options, onChange, columnName }) {
      const [draft, setDraft] = useState('');
      const add = () => {
        const v = draft.trim();
        if (!v) return;
        if (options.some(opt => dropdownOptionLabel(opt).trim().toLowerCase() === v.toLowerCase())) { setDraft(''); return; }
        onChange([...options, v]);
        setDraft('');
      };
      const remove = (i) => onChange(options.filter((_, j) => j !== i));
      const updateAt = (i, v) => onChange(options.map((o, j) => {
        if (j !== i) return o;
        const color = dropdownOptionColor(o);
        return color ? { label: v, color } : v;
      }));
      const updateColor = (i, color) => onChange(options.map((o, j) => {
        if (j !== i) return o;
        const label = dropdownOptionLabel(o);
        return color ? { label, color } : label;
      }));
      return (
        <div className="ml-8 mr-2 mt-1.5 mb-2 rounded-md border border-blue-500/20 bg-blue-500/5 p-3">
          <div className="text-[10px] uppercase tracking-wide text-neutral-500 mb-2">
            Dropdown options{columnName ? <> for <span className="text-neutral-300">"{columnName}"</span></> : ''}
          </div>
          {options.length === 0 ? (
            <p className="text-[11px] text-neutral-500 mb-2">No options yet — cells stay free-text. Add options below to turn this column into a dropdown.</p>
          ) : (
            <div className="space-y-1 mb-2">
              {options.map((o, i) => (
                <div key={i} className="grid grid-cols-[28px_minmax(160px,1fr)_112px_28px] items-center gap-2">
                  <span className="text-[10px] text-neutral-600 font-mono w-6 text-right shrink-0">{String(i + 1).padStart(2, '0')}</span>
                  <Input value={dropdownOptionLabel(o)} onChange={e => updateAt(i, e.target.value)} placeholder="Option text" />
                  <div className="flex items-center gap-1">
                    <PresetColorPicker value={dropdownOptionColor(o) || '#FFFFFF'}
                      onChange={color => updateColor(i, color)}
                      title="Choose option color" showHex={false} buttonClassName="h-8 px-1.5 py-1" />
                    {dropdownOptionColor(o) && (
                      <button onClick={() => updateColor(i, '')}
                        className="text-[10px] text-neutral-500 hover:text-red-300 px-1"
                        title="Clear option color">clear</button>
                    )}
                  </div>
                  <button onClick={() => remove(i)} className="p-1.5 text-neutral-500 hover:text-red-400 shrink-0" title="Remove option"><IconX /></button>
                </div>
              ))}
            </div>
          )}
          <div className="flex gap-2">
            <Input value={draft} onChange={e => setDraft(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && (e.preventDefault(), add())}
              placeholder="e.g. SBC 1" />
            <Btn variant="ghost" size="md" onClick={add}>Add option</Btn>
          </div>
          <p className="text-[10px] text-neutral-600 mt-2 leading-relaxed">
            Applies to ~500 cells below the header in the generated .xlsx (Excel & Google Sheets compatible). Commas inside option text are converted to spaces.
          </p>
        </div>
      );
    }

    /* ============================================================
       Sheet Editor
       ============================================================ */
    function SheetEditor({ sheet, onChange, onDelete, onDuplicate }) {
      if (!sheet) {
        return (
          <div className="rounded-lg border border-dashed border-neutral-800 px-6 py-16 text-center">
            <p className="text-sm text-neutral-400">No sheet selected</p>
            <p className="text-xs text-neutral-600 mt-1">Pick a sheet from the sidebar or add a new one.</p>
          </div>
        );
      }
      const update = (patch) => onChange({ ...sheet, ...patch });
      const setLayout = (layout) => {
        if (layout === sheet.layout) return;
        const next = { ...sheet, layout };
        if (layout === 'hourly') {
          next.metrics = sheet.metrics?.length ? sheet.metrics : (sheet.columns?.length ? sheet.columns : ['Metric 1']);
          next.hourStart = sheet.hourStart ?? 0;
          next.hourEnd = sheet.hourEnd ?? 24;
          next.rowSeparator = sheet.rowSeparator ?? true;
        } else {
          next.columns = sheet.columns?.length ? sheet.columns : (sheet.metrics?.length ? sheet.metrics : ['Column 1']);
        }
        onChange(next);
      };

      return (
        <div className="space-y-10">
          {/* Header */}
          <div className="flex items-start justify-between gap-6 border-b border-neutral-900 pb-7">
            <div className="flex-1 min-w-0">
              <SectionLabel>Sheet</SectionLabel>
              <input value={sheet.name} onChange={e => update({ name: e.target.value })}
                className="text-3xl font-bold tracking-tight bg-transparent border-none outline-none w-full focus:bg-neutral-900/50 rounded px-1.5 -mx-1.5 py-0.5" />
              <div className="flex items-center gap-2 mt-3 flex-wrap">
                <Pill tone="accent">{LAYOUT_LABELS[sheet.layout]}</Pill>
                <Pill tone={sheet.active ? 'success' : 'muted'}>{sheet.active ? '● Active' : '○ Inactive'}</Pill>
                {sheet.layout === 'hourly' && <Pill>{visibleItems(sheet.metrics).length}/{(sheet.metrics || []).length} metrics</Pill>}
                {sheet.layout === 'hourly' && <Pill>{hourlyHasRowSeparator(sheet) ? 'With separator' : 'No separator'}</Pill>}
                {sheet.layout !== 'hourly' && <Pill>{visibleItems(sheet.columns).length}/{(sheet.columns || []).length} columns</Pill>}
              </div>
            </div>
            <div className="flex flex-col gap-2 shrink-0 w-44">
              <Select value={sheet.layout} onChange={e => setLayout(e.target.value)}>
                <option value="hourly">Hourly Block</option>
                <option value="flat">Flat Table</option>
                <option value="alarm">Alarm Log</option>
              </Select>
              <div className="flex gap-2">
                <Btn variant="ghost" size="md" onClick={onDuplicate} className="flex-1">Duplicate</Btn>
                <Btn variant="danger" size="md" onClick={onDelete} className="flex-1">Delete</Btn>
              </div>
            </div>
          </div>

          {/* Note */}
          <CollapsibleSection sid="hdr-note" title="Header note" hint="Shown at top of the generated sheet">
            <Textarea rows={2} value={sheet.note || ''} onChange={e => update({ note: e.target.value })} placeholder="Optional instructional note (e.g. 'Highlight peak utilization hours in red')" />
            <div className="mt-3 grid grid-cols-2 md:grid-cols-6 gap-3 items-end">
              <div className="md:col-span-2">
                <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Font family</label>
                <Select value={sheet.noteFontFamily || DEFAULT_NOTE_STYLE.noteFontFamily}
                  onChange={e => update({ noteFontFamily: e.target.value })}>
                  {FONT_FAMILIES.map(f => <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>)}
                </Select>
              </div>
              <div>
                <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Size</label>
                <Select value={sheet.noteFontSize ?? DEFAULT_NOTE_STYLE.noteFontSize}
                  onChange={e => update({ noteFontSize: Number(e.target.value) })}>
                  {FONT_SIZES.map(n => <option key={n} value={n}>{n}pt</option>)}
                </Select>
              </div>
              <div>
                <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Color</label>
                <div className="flex items-center gap-1">
                  <PresetColorPicker value={'#' + (sheet.noteFontColor || DEFAULT_NOTE_STYLE.noteFontColor).slice(2)}
                    onChange={color => update({ noteFontColor: 'FF' + color.slice(1) })}
                    title="Choose header note font color" className="flex-1" />
                  {(sheet.noteFontColor || DEFAULT_NOTE_STYLE.noteFontColor) !== DEFAULT_NOTE_STYLE.noteFontColor && (
                    <button onClick={() => update({ noteFontColor: DEFAULT_NOTE_STYLE.noteFontColor })}
                      title={`Reset to default (#${DEFAULT_NOTE_STYLE.noteFontColor.slice(2)})`}
                      className="text-neutral-500 hover:text-neutral-200 transition-colors"><IconReset /></button>
                  )}
                </div>
              </div>
              <div>
                <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Align</label>
                <Select value={sheet.noteAlign || DEFAULT_NOTE_STYLE.noteAlign}
                  onChange={e => update({ noteAlign: e.target.value })}>
                  <option value="left">Left</option>
                  <option value="center">Center</option>
                  <option value="right">Right</option>
                </Select>
              </div>
              <div>
                <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Style</label>
                <div className="flex gap-1">
                  <button onClick={() => update({ noteBold: !sheet.noteBold })}
                    title="Bold"
                    className={`flex-1 h-9 rounded-md border text-sm font-bold transition-colors ${sheet.noteBold ? 'border-blue-500/40 bg-blue-500/10 text-blue-300' : 'border-neutral-800 bg-neutral-950 text-neutral-500 hover:text-neutral-200 hover:border-neutral-700'}`}>B</button>
                  <button onClick={() => update({ noteItalic: !(sheet.noteItalic !== undefined ? sheet.noteItalic : true) })}
                    title="Italic"
                    className={`flex-1 h-9 rounded-md border text-sm italic transition-colors ${(sheet.noteItalic !== undefined ? sheet.noteItalic : true) ? 'border-blue-500/40 bg-blue-500/10 text-blue-300' : 'border-neutral-800 bg-neutral-950 text-neutral-500 hover:text-neutral-200 hover:border-neutral-700'}`}>I</button>
                  <button onClick={() => update({ noteUnderline: !sheet.noteUnderline })}
                    title="Underline"
                    className={`flex-1 h-9 rounded-md border text-sm underline transition-colors ${sheet.noteUnderline ? 'border-blue-500/40 bg-blue-500/10 text-blue-300' : 'border-neutral-800 bg-neutral-950 text-neutral-500 hover:text-neutral-200 hover:border-neutral-700'}`}>U</button>
                </div>
              </div>
            </div>
            {sheet.note && sheet.note.trim() && (
              <div className="mt-3 rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2">
                <div className="text-[10px] uppercase tracking-wide text-neutral-600 mb-1">Live preview</div>
                <div style={{
                  fontFamily: sheet.noteFontFamily || DEFAULT_NOTE_STYLE.noteFontFamily,
                  fontSize: `${sheet.noteFontSize ?? DEFAULT_NOTE_STYLE.noteFontSize}pt`,
                  color: '#' + (sheet.noteFontColor || DEFAULT_NOTE_STYLE.noteFontColor).slice(2),
                  fontWeight: sheet.noteBold ? 700 : 400,
                  fontStyle: (sheet.noteItalic !== undefined ? sheet.noteItalic : true) ? 'italic' : 'normal',
                  textDecoration: sheet.noteUnderline ? 'underline' : 'none',
                  textAlign: sheet.noteAlign || 'left',
                }}>{sheet.note}</div>
              </div>
            )}
          </CollapsibleSection>

          {/* Metrics / Columns */}
          {sheet.layout === 'hourly' ? (
            <>
              <CollapsibleSection sid="cols" title="Column headers" hint="Shown in the header row above each column">
                <div className="grid grid-cols-2 gap-3 max-w-lg">
                  <div>
                    <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Date column label</label>
                    <Input value={sheet.dateHeader ?? 'DATE'} onChange={e => update({ dateHeader: e.target.value })} placeholder="DATE" />
                  </div>
                  <div>
                    <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Metric column label</label>
                    <Input value={sheet.metricHeader ?? ''} onChange={e => update({ metricHeader: e.target.value })} placeholder="METRIC" />
                  </div>
                </div>
              </CollapsibleSection>
              <CollapsibleSection sid="dayblock" title="Day block spacing" hint="Controls whether a blank row is inserted between each daily block" defaultCollapsed>
                <div className="inline-flex rounded-md border border-neutral-900 bg-neutral-950 p-0.5">
                  {[
                    [true, 'With row separator'],
                    [false, 'No row separator'],
                  ].map(([value, label]) => (
                    <button key={String(value)} type="button"
                      onClick={() => update({ rowSeparator: value })}
                      className={`rounded px-3 py-1.5 text-xs font-medium transition-colors ${hourlyHasRowSeparator(sheet) === value ? 'bg-neutral-800 text-neutral-100' : 'text-neutral-500 hover:text-neutral-200'}`}>
                      {label}
                    </button>
                  ))}
                </div>
              </CollapsibleSection>
              <CollapsibleSection sid="dateformat" title="Date format" hint="Applied to the date column in the .xlsx (Excel & Google Sheets compatible)" defaultCollapsed>
                {(() => {
                  const fmt = sheet.dateFormat || 'yyyy-mm-dd';
                  const isCustom = !DATE_FORMAT_VALUES.has(fmt);
                  const sample = previewDateFormat(new Date(2026, 5, 1), fmt); // June 1, 2026
                  return (
                    <div className="space-y-2 max-w-lg">
                      <div className="grid grid-cols-2 gap-3">
                        <div>
                          <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Preset</label>
                          <Select value={isCustom ? '__custom__' : fmt}
                            onChange={e => update({ dateFormat: e.target.value === '__custom__' ? fmt : e.target.value })}>
                            {DATE_FORMATS.map(d => <option key={d.value} value={d.value}>{d.sample}</option>)}
                            <option value="__custom__">Custom…</option>
                          </Select>
                        </div>
                        <div>
                          <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">{isCustom ? 'Custom format code' : 'Format code'}</label>
                          <Input value={fmt} onChange={e => update({ dateFormat: e.target.value })} placeholder="yyyy-mm-dd" />
                        </div>
                      </div>
                      <div className="rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2 flex items-center justify-between gap-3">
                        <span className="text-[10px] uppercase tracking-wide text-neutral-600">Preview (Jun 1, 2026)</span>
                        <span className="text-sm font-mono text-neutral-200">{sample}</span>
                      </div>
                      <p className="text-[11px] text-neutral-500 font-mono">
                        Tokens: yyyy yy · mmmm mmm mm m · dddd ddd dd d
                      </p>
                    </div>
                  );
                })()}
              </CollapsibleSection>
              <CollapsibleSection sid="metrics" title="Metric rows" hint={`${visibleItems(sheet.metrics).length} of ${(sheet.metrics || []).length} rows per day block`}>
                <ListEditor items={sheet.metrics || []} onChange={(v) => update({ metrics: v })} placeholder="e.g. Active Calls" label="metric" withColor />
              </CollapsibleSection>
              <CollapsibleSection sid="hourrange" title="Hour range" hint="24h = full 24/7 cycle">
                <div className="grid grid-cols-2 gap-3 max-w-lg">
                  <div>
                    <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Start hour</label>
                    <Select value={sheet.hourStart ?? 0} onChange={e => update({ hourStart: Number(e.target.value) })}>
                      {HOURS_24.map((h, i) => <option key={i} value={i}>{h}</option>)}
                    </Select>
                  </div>
                  <div>
                    <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">End hour (inclusive)</label>
                    <Select value={(sheet.hourEnd ?? 24) - 1} onChange={e => update({ hourEnd: Number(e.target.value) + 1 })}>
                      {HOURS_24.map((h, i) => <option key={i} value={i}>{h}</option>)}
                    </Select>
                  </div>
                </div>
                <p className="text-[11px] text-neutral-500 mt-2 font-mono">
                  {HOURS_24[sheet.hourStart ?? 0]} → {HOURS_24[(sheet.hourEnd ?? 24) - 1]} · {(sheet.hourEnd ?? 24) - (sheet.hourStart ?? 0)} columns
                </p>
              </CollapsibleSection>
              <CollapsibleSection sid="cellborders" title="Cell borders" hint="Borders on every fillable data cell in each day block" defaultCollapsed>
                <label className="inline-flex min-h-10 items-center gap-3 rounded-md border border-neutral-900 bg-neutral-950 px-3 py-2 text-sm text-neutral-300 hover:border-neutral-800 hover:bg-neutral-900/50 select-none mb-3" style={{ cursor: 'pointer' }}>
                  <input type="checkbox" checked={!!sheet.bordersOnData}
                    onChange={e => update({ bordersOnData: e.target.checked })}
                    className="h-4 w-4 shrink-0 rounded border-neutral-700 bg-neutral-900 text-blue-500 focus:ring-blue-500/30 cursor-pointer" />
                  Add borders to data range
                </label>
                {sheet.bordersOnData && (
                  <div className="grid grid-cols-2 gap-3 max-w-lg">
                    <div>
                      <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Style</label>
                      <Select value={sheet.borderStyle || 'thin'} onChange={e => update({ borderStyle: e.target.value })}>
                        <option value="thin">Thin</option>
                        <option value="medium">Medium</option>
                        <option value="thick">Thick</option>
                        <option value="dotted">Dotted</option>
                        <option value="dashed">Dashed</option>
                      </Select>
                    </div>
                    <div>
                      <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Color</label>
                      <div className="flex items-center gap-1">
                        <PresetColorPicker value={'#' + (sheet.borderColor || 'FF374151').slice(2)}
                          onChange={color => update({ borderColor: 'FF' + color.slice(1) })}
                          title="Choose border color" className="flex-1" />
                        {(sheet.borderColor || 'FF374151') !== 'FF374151' && (
                          <button onClick={() => update({ borderColor: 'FF374151' })}
                            title="Reset to default (#374151)"
                            className="text-neutral-500 hover:text-neutral-200 transition-colors"><IconReset /></button>
                        )}
                      </div>
                    </div>
                  </div>
                )}
              </CollapsibleSection>
            </>
          ) : (
            <div>
              <SectionLabel hint={sheet.layout === 'alarm' ? 'Click the List badge to add per-column dropdown options' : null}>Columns</SectionLabel>
              <ListEditor
                items={sheet.columns || []}
                onChange={(v) => update({ columns: v })}
                placeholder={sheet.layout === 'alarm' ? 'e.g. SBC' : 'e.g. Trunk ID'}
                label="column"
                withDropdown={sheet.layout === 'alarm'} />
            </div>
          )}
        </div>
      );
    }

    /* ============================================================
       Formatting Rules
       ============================================================ */
    const RULE_TYPES = [
      { value: 'maxInRow',   label: 'Highlight max value in row',   desc: 'Per-row maximum across hour columns (Hourly Block only)' },
      { value: 'maxInCol',   label: 'Highlight max value in column', desc: 'Per-column maximum down each hour column (Hourly Block only)' },
      { value: 'above',      label: 'Highlight above threshold',    needsThreshold: true },
      { value: 'below',      label: 'Highlight below threshold',    needsThreshold: true },
      { value: 'blank',      label: 'Highlight blank cells',        desc: 'Flag missing data' },
      { value: 'zero',       label: 'Highlight zero values' },
      { value: 'weekend',    label: 'Highlight weekend rows',       desc: 'Saturday & Sunday entire row' },
      { value: 'separator',  label: 'Color separator rows',         desc: 'The blank row between each day block (Hourly Block only)' },
      { value: 'boldMetric', label: 'Bold specific metric rows',    needsMatch: true, desc: 'Static formatting applied at generation' },
    ];

    function FormattingRules({ rules, sheets, onChange }) {
      const [ruleScope, setRuleScope] = useState('__view_all__');
      const add = (target = ruleScope) => {
        const targetSheet = target === '__view_all__' ? '__all__' : target;
        const targetLayout = sheets.find(sheet => sheet.id === targetSheet)?.layout || SIP_LAYOUT_RULE_SCOPES.find(scope => scope.value === targetSheet)?.layout;
        const defaultType = targetLayout && targetLayout !== 'hourly' ? 'blank' : 'maxInRow';
        onChange([...rules, {
          id: newId(), type: defaultType, enabled: true,
          color: 'FF166534', fontColor: 'FFFFFFFF', bold: false,
          threshold: '', metricMatch: '', targetSheet,
        }]);
      };
      const update = (id, patch) => onChange(rules.map(r => r.id === id ? { ...r, ...patch } : r));
      const remove = (id) => onChange(rules.filter(r => r.id !== id));
      const move = (id, dir) => {
        const idx = rules.findIndex(r => r.id === id);
        const next = idx + dir;
        if (idx < 0 || next < 0 || next >= rules.length) return;
        const arr = [...rules];
        [arr[idx], arr[next]] = [arr[next], arr[idx]];
        onChange(arr);
      };
      const scopedRules = rules
        .map((rule, index) => ({ rule, index }))
        .filter(({ rule }) => sipRuleScopeMatches(rule, ruleScope));
      const scopeLabel = ruleScope === '__view_all__'
        ? 'all rules'
        : ruleScope === '__all__'
          ? 'all sheets'
          : (SIP_LAYOUT_RULE_SCOPES.find(scope => scope.value === ruleScope)?.label || sheets.find(sheet => sheet.id === ruleScope)?.name || 'selected sheet');

      return (
        <div className="space-y-3">
          <div className="flex flex-wrap items-end justify-between gap-3 rounded-md border border-neutral-900 bg-neutral-950/60 p-3">
            <div className="w-full max-w-sm">
              <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Rule set</label>
              <Select value={ruleScope} onChange={e => setRuleScope(e.target.value)}>
                <option value="__view_all__">All rules</option>
                <option value="__all__">Global rules - all sheets</option>
                {SIP_LAYOUT_RULE_SCOPES.map(scope => <option key={scope.value} value={scope.value}>{scope.label}</option>)}
                {sheets.map(sheet => <option key={sheet.id} value={sheet.id}>{sheet.name}</option>)}
              </Select>
            </div>
            <div className="flex flex-wrap items-center gap-2 pb-0.5">
              <Pill>{scopedRules.filter(({ rule }) => rule.enabled).length} enabled</Pill>
              <Pill tone="muted">{scopedRules.length} shown</Pill>
              <Btn variant="ghost" size="md" onClick={() => add()}><IconPlus /> Add rule</Btn>
            </div>
            <p className="w-full text-[11px] text-neutral-500">
              Rules can be global, layout-wide, row-separator-specific, or assigned to one sheet. Unchecked rules are disabled and will not be applied during generation. Max-in-row, max-in-column, weekend, separator, and metric-name rules are hourly-only; threshold, blank, and zero rules also work on flat/alarm sheets.
            </p>
          </div>
          {rules.length === 0 && (
            <div className="rounded-lg border border-dashed border-neutral-800 px-5 py-12 text-center">
              <p className="text-sm text-neutral-300">No formatting rules yet</p>
              <p className="text-xs text-neutral-600 mt-1">Add rules to highlight peaks, flag missing data, mark weekends, etc.</p>
              <Btn variant="ghost" size="md" onClick={() => add()} className="mt-4"><IconPlus /> Add your first rule</Btn>
            </div>
          )}
          {rules.length > 0 && scopedRules.length === 0 && (
            <div className="rounded-lg border border-dashed border-neutral-800 px-5 py-10 text-center">
              <p className="text-sm text-neutral-300">No rules for {scopeLabel}</p>
              <p className="text-xs text-neutral-600 mt-1">Add a rule here to keep this sheet's formatting separate.</p>
              <Btn variant="ghost" size="md" onClick={() => add()} className="mt-4"><IconPlus /> Add rule</Btn>
            </div>
          )}
          {scopedRules.map(({ rule: r, index: idx }, displayIdx) => {
            const typeDef = RULE_TYPES.find(t => t.value === r.type);
            return (
              <Card key={r.id} className="p-4">
                <div className="flex items-center gap-3 mb-4">
                  <input type="checkbox" checked={r.enabled} onChange={e => update(r.id, { enabled: e.target.checked })}
                    className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500 focus:ring-blue-500/30 cursor-pointer" />
                  <span className="text-[10px] font-mono text-neutral-600 w-6">{String(displayIdx + 1).padStart(2, '0')}</span>
                  <Select value={r.type} onChange={e => update(r.id, { type: e.target.value })} className="flex-1">
                    {RULE_TYPES.map(t => <option key={t.value} value={t.value}>{t.label}</option>)}
                  </Select>
                  <div className="flex text-neutral-500">
                    <button onClick={() => move(r.id, -1)} className="p-1.5 hover:text-neutral-200" title="Move up"><IconUp /></button>
                    <button onClick={() => move(r.id, 1)} className="p-1.5 hover:text-neutral-200" title="Move down"><IconDown /></button>
                    <button onClick={() => remove(r.id)} className="p-1.5 hover:text-red-400" title="Remove"><IconX /></button>
                  </div>
                </div>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                  <div>
                    <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Applies to</label>
                    <Select value={r.targetSheet || '__all__'} onChange={e => update(r.id, { targetSheet: e.target.value })}>
                      <option value="__all__">All sheets</option>
                      {SIP_LAYOUT_RULE_SCOPES.map(scope => <option key={scope.value} value={scope.value}>{scope.label}</option>)}
                      {sheets.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                    </Select>
                  </div>
                  <div>
                    <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Background</label>
                    <PresetColorPicker value={'#' + (r.color || 'FF166534').slice(2)}
                      onChange={color => update(r.id, { color: 'FF' + color.slice(1) })}
                      title="Choose rule background color" className="w-full" />
                  </div>
                  <div>
                    <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Font color</label>
                    <PresetColorPicker value={'#' + (r.fontColor || 'FFFFFFFF').slice(2)}
                      onChange={color => update(r.id, { fontColor: 'FF' + color.slice(1) })}
                      title="Choose rule font color" className="w-full" />
                  </div>
                  {typeDef?.needsThreshold && (
                    <div>
                      <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Threshold</label>
                      <Input type="number" value={r.threshold} onChange={e => update(r.id, { threshold: e.target.value })} placeholder="0" />
                    </div>
                  )}
                  {typeDef?.needsMatch && (
                    <div className="col-span-2">
                      <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Metric name contains (comma-separated)</label>
                      <Input value={r.metricMatch} onChange={e => update(r.id, { metricMatch: e.target.value })} placeholder="PEAK, MAX" />
                    </div>
                  )}
                  <div>
                    <label className="text-[10px] uppercase tracking-wide text-neutral-500 block mb-1.5">Bold</label>
                    <button onClick={() => update(r.id, { bold: !r.bold })}
                      className={`w-full h-9 rounded-md border text-sm font-bold transition-colors ${r.bold ? 'border-blue-500/40 bg-blue-500/10 text-blue-300' : 'border-neutral-800 bg-neutral-950 text-neutral-500 hover:text-neutral-200 hover:border-neutral-700'}`}>
                      B
                    </button>
                  </div>
                </div>
                {typeDef?.desc && <p className="text-[11px] text-neutral-500 mt-3">{typeDef.desc}</p>}
              </Card>
            );
          })}
          {rules.length > 0 && <Btn variant="ghost" size="md" onClick={() => add()}><IconPlus /> Add rule</Btn>}
        </div>
      );
    }

    /* ============================================================
       Preview
       ============================================================ */
    function Preview({ sheet, year, month }) {
      if (!sheet) return null;
      if (sheet.layout === 'hourly') {
        const hours = HOURS_24.slice(sheet.hourStart ?? 0, sheet.hourEnd ?? 24);
        const sample = Math.min(2, daysInMonth(year, month));
        const showSeparator = hourlyHasRowSeparator(sheet);
        return (
          <div className="overflow-x-auto rounded-lg border border-neutral-900">
            <table className="min-w-full text-[11px]">
              <thead>
                <tr className="border-b border-neutral-800 bg-neutral-950">
                  <th className="text-left px-3 py-2.5 font-semibold text-neutral-300 sticky left-0 bg-neutral-950">DATE</th>
                  <th className="text-left px-3 py-2.5 font-semibold text-neutral-300">{sheet.metricHeader ?? ''}</th>
                  {hours.map(h => <th key={h} className="text-center px-2 py-2.5 font-semibold text-neutral-400 font-mono">{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {[...Array(sample)].map((_, dIdx) => {
                  const wk = isWeekend(year, month, dIdx + 1);
                  return (
                    <React.Fragment key={dIdx}>
                      {visibleItems(sheet.metrics).map((m, mIdx) => (
                        <tr key={mIdx} className={`border-b border-neutral-900/60 ${wk ? 'bg-red-950/20' : ''}`}>
                          <td className="px-3 py-1.5 text-neutral-400 font-mono sticky left-0 bg-[#232327]">{mIdx === 0 ? `${year}-${pad(month + 1)}-${pad(dIdx + 1)}` : ''}</td>
                          <td className="px-3 py-1.5 text-neutral-200" style={itemColor(m) ? { background: itemColor(m) } : {}}>{itemLabel(m)}</td>
                          {hours.map((_, hi) => <td key={hi} className="px-2 py-1.5 text-neutral-700 text-center" style={itemColor(m) && itemColorExtend(m) ? { background: itemColor(m) } : {}}>·</td>)}
                        </tr>
                      ))}
                      {showSeparator && dIdx < sample - 1 && <tr className="h-2 bg-[#1c1c1f]"><td colSpan={2 + hours.length}></td></tr>}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        );
      }
      return (
        <div className="overflow-x-auto rounded-lg border border-neutral-900">
          <table className="min-w-full text-[11px]">
            <thead>
              <tr className="border-b border-neutral-800 bg-neutral-950">
                {visibleLabels(sheet.columns).map(c => <th key={c} className="text-left px-3 py-2.5 font-semibold text-neutral-300">{c}</th>)}
              </tr>
            </thead>
            <tbody>
              {[...Array(3)].map((_, i) => (
                <tr key={i} className="border-b border-neutral-900/60">
                  {visibleLabels(sheet.columns).map(c => <td key={c} className="px-3 py-1.5 text-neutral-700">·</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
    }

    /* ============================================================
       Template bar
       ============================================================ */
    function TemplateBar({ state, setState, sync }) {
      const [templates, setTemplates] = useState(() => {
        try { return JSON.parse(localStorage.getItem(TEMPLATES_KEY)) || []; } catch { return []; }
      });
      const [name, setName] = useState('');

      // Hydrate templates from Firestore once auth is ready
      useEffect(() => {
        if (!window.__fb || !window.__fbm || !sync?.uid) return;
        let cancelled = false;
        (async () => {
          try {
            const { doc, getDoc } = window.__fbm;
            const ref = doc(window.__fb.db, 'users', sync.uid, 'state', 'templates');
            const snap = await getDoc(ref);
            if (cancelled) return;
            const remote = snap.exists() ? (firestoreDesanitize(snap.data()).items || []) : null;
            if (Array.isArray(remote) && remote.length) {
              const merged = [...remote];
              for (const local of templates) {
                if (!merged.some(t => t.name === local.name)) merged.push(local);
              }
              setTemplates(merged);
              localStorage.setItem(TEMPLATES_KEY, JSON.stringify(merged));
            }
          } catch (e) { console.warn('templates hydrate failed', e); }
        })();
        return () => { cancelled = true; };
      }, [sync?.uid]);

      const persist = (next) => {
        setTemplates(next);
        localStorage.setItem(TEMPLATES_KEY, JSON.stringify(next));
        if (window.__fb && window.__fbm && sync?.uid) {
          const { doc, setDoc, serverTimestamp } = window.__fbm;
          const ref = doc(window.__fb.db, 'users', sync.uid, 'state', 'templates');
          setDoc(ref, { items: firestoreSanitize(next), updatedAt: serverTimestamp() })
            .catch(e => console.warn('templates save failed', e));
        }
      };

      const save = () => {
        if (!name.trim()) return;
        const t = { id: newId(), name: name.trim(), savedAt: new Date().toISOString(), data: { sheets: state.sheets, rules: state.rules, includeIndex: state.includeIndex } };
        persist([...templates.filter(x => x.name !== t.name), t]);
        setName('');
      };
      const load = (id) => {
        const t = templates.find(x => x.id === id);
        if (!t) return;
        setState(s => ({ ...s, ...t.data, selectedSheetId: t.data.sheets[0]?.id || null }));
      };
      const del = (id) => persist(templates.filter(x => x.id !== id));
      const exportAll = () => {
        const blob = new Blob([JSON.stringify({ templates, current: { sheets: state.sheets, rules: state.rules } }, null, 2)], { type: 'application/json' });
        saveAs(blob, `mrg-templates-${new Date().toISOString().slice(0,10)}.json`);
      };
      const importFile = (e) => {
        const f = e.target.files[0];
        if (!f) return;
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const parsed = JSON.parse(reader.result);
            if (Array.isArray(parsed.templates)) persist([...templates, ...parsed.templates]);
            else if (parsed.sheets) setState(s => ({ ...s, sheets: parsed.sheets, rules: parsed.rules || s.rules }));
          } catch { alertDialog({ title: 'Invalid file', message: "That file isn't a valid template export.", confirmText: 'OK', tone: 'danger' }); }
        };
        reader.readAsText(f);
        e.target.value = '';
      };

      const importXlsx = async (e) => {
        const f = e.target.files[0];
        e.target.value = '';
        if (!f) return;
        let detected;
        try {
          detected = await parseXlsxToSheets(f);
        } catch (err) {
          console.warn('xlsx import failed', err);
          alertDialog({ title: "Couldn't read file", message: "That file couldn't be parsed as an .xlsx workbook.", confirmText: 'OK', tone: 'danger' });
          return;
        }
        if (!detected.length) {
          alertDialog({ title: 'Nothing to import', message: "No sheets with a recognizable header row were found in that workbook.", confirmText: 'OK' });
          return;
        }
        // Loaded into the Uploaded sidebar — does NOT change the main config.
        // From there, the user adds individual tabs into the main sidebar.
        setState(s => ({ ...s, imported: [...(s.imported || []), ...detected] }));
      };

      return (
        <Card className="p-4">
          <div className="flex items-center gap-2 flex-wrap">
            <Input value={name} onChange={e => setName(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && (e.preventDefault(), save())}
              placeholder="Template name (e.g. Q3 2026 Standard)" className="max-w-xs" />
            <Btn variant="ghost" size="md" onClick={save} disabled={!name.trim()}>Save current</Btn>
            <div className="w-px h-6 bg-neutral-800 mx-1"></div>
            <Select onChange={e => { if (e.target.value) load(e.target.value); e.target.value = ''; }} className="max-w-xs">
              <option value="">Load template…</option>
              {templates.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
            </Select>
            <Btn variant="ghost" size="md" onClick={exportAll}>Export JSON</Btn>
            <label className="inline-flex items-center justify-center gap-2 rounded-md font-medium transition-colors px-3.5 py-2 text-sm border border-neutral-800 bg-transparent text-neutral-200 hover:bg-neutral-900 hover:border-neutral-700 cursor-pointer">
              Import JSON
              <input type="file" accept=".json,application/json" onChange={importFile} className="hidden" />
            </label>
            <label title="Upload an existing .xlsx — auto-detects sheets, layouts, metrics and columns"
              className="inline-flex items-center justify-center gap-2 rounded-md font-medium transition-colors px-3.5 py-2 text-sm border border-neutral-800 bg-transparent text-neutral-200 hover:bg-neutral-900 hover:border-neutral-700 cursor-pointer">
              Import .xlsx
              <input type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={importXlsx} className="hidden" />
            </label>
          </div>
          {templates.length > 0 && (
            <div className="flex flex-wrap gap-1.5 mt-4 pt-4 border-t border-neutral-900">
              <span className="text-[10px] uppercase tracking-wide text-neutral-600 self-center mr-1">Saved</span>
              {templates.map(t => (
                <span key={t.id} className="inline-flex items-center gap-1.5 rounded-full border border-neutral-800 bg-neutral-900 pl-2.5 pr-1.5 py-0.5 text-[11px] text-neutral-300">
                  {t.name}
                  <button onClick={() => del(t.id)} className="text-neutral-500 hover:text-red-400 p-0.5"><IconX /></button>
                </span>
              ))}
            </div>
          )}
        </Card>
      );
    }

    /* ============================================================
       BMR — UI Components
       ============================================================ */
    const bmrId = (p='bc') => p + '_' + Math.random().toString(36).slice(2, 9);

    const DEFAULT_BMR_CLIENTS = [
      { id: bmrId(), name: 'APNTEL', color: '#FACC15', allowBalance: 0, allowBalanceLabel: 'Block once debit', hidden: false },
      { id: bmrId(), name: 'ASIACON', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'ASIAPAY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'BBI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CCPI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'COMMSOL-DID', color: '#60A5FA', allowBalance: 5, allowBalanceLabel: '$5 debit poc', hidden: false },
      { id: bmrId(), name: 'COMQUEST', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'COMQUEST HYA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'COMQUEST MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN 5STONE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ACOM', color: '#34D399', allowBalance: 700, allowBalanceLabel: '$700 debit / month', hidden: false },
      { id: bmrId(), name: 'CYN ACTIVEONE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ADMEREX', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ADSPARK', color: '', allowBalance: 2, allowBalanceLabel: 'advise sila pag $2 credit nalang', hidden: false },
      { id: bmrId(), name: 'CYN ADVANTAGE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AEI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AGO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AI RUDDER', color: '', allowBalance: 20000, allowBalanceLabel: '$20000 debit', hidden: false },
      { id: bmrId(), name: 'CYN AIQON', color: '', allowBalance: 5000, allowBalanceLabel: '$5000 debit', hidden: false },
      { id: bmrId(), name: 'CYN ALEQX', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ALLIANZPNB', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ALLINONE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ALLYOURVOIP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ALPHATECH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AMADA', color: '', allowBalance: 10, allowBalanceLabel: '$10 credit block', hidden: false },
      { id: bmrId(), name: 'CYN AMADEUS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AMEYO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AMG', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AMIHAN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AMZ', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ANCHOR', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ANCHOR2', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ANSR', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN AOCS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN APLUS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN APLUS JAKA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN APPXCONNECT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN APS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AQUOZ', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ARANDA-PP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ARBOUR', color: '', allowBalance: 200, allowBalanceLabel: '$200 debit', hidden: false },
      { id: bmrId(), name: 'CYN ARLENE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ASCII POC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ASHOK', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ASIALINK', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ATENTO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ATI CLINICA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ATICASIA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AUCTUS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AURORA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AVANTICE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AVANZA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AVON SSNL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN AZTEL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN B1BPO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN B2P', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BAC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BALIGOL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BBX', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BEEINFO', color: '', allowBalance: 0, allowBalanceLabel: 'mrc', hidden: false },
      { id: bmrId(), name: 'CYN BENTLEY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BETVISION', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BFB', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BICS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BIOMET', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BLUE SEO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BLUESHIFT', color: '', allowBalance: 0, allowBalanceLabel: 'poc', hidden: false },
      { id: bmrId(), name: 'CYN BMiTC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BOUNCETEL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN BYOC HYPHEN', color: '', allowBalance: 8000, allowBalanceLabel: '$8000 usage / month', hidden: false },
      { id: bmrId(), name: 'CYN BZI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CABMCS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CALLBEST', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CASH EXPRESS', color: '', allowBalance: 200, allowBalanceLabel: '$200 debit', hidden: false },
      { id: bmrId(), name: 'CYN CEPAT KREDIT', color: '', allowBalance: 920, allowBalanceLabel: '$920 debit', hidden: false },
      { id: bmrId(), name: 'CYN CEREBRO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CG8Rockwell', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CGP-IBP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CGPI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CHERRY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CiD', color: '', allowBalance: 0, allowBalanceLabel: 'currently blocked', hidden: false },
      { id: bmrId(), name: 'CYN CIMB', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CIMB AZURE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CIMB-MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CJM', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CLARKOUTSOURCING', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CLINK', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CLOUDCORE', color: '', allowBalance: 50, allowBalanceLabel: '$50 debit (temp credit)\ncurrently blocked', hidden: false },
      { id: bmrId(), name: 'CYN CLOUDDIAL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CLPAY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CMS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN COLLECTIUS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN COMSERV', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN COMSERV2', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN COMSERVECO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN COOLWAVE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CORVENTURES', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CTGLOBAL', color: '', allowBalance: 20000, allowBalanceLabel: '$20000 debit', hidden: false },
      { id: bmrId(), name: 'CYN CXMARKETING', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CYBER365', color: '', allowBalance: 3, allowBalanceLabel: '$3 debit', hidden: false },
      { id: bmrId(), name: 'CYN CYBERSOFT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CYBERVOIP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN CYBERVOIP2', color: '', allowBalance: 0, allowBalanceLabel: 'Block once debit', hidden: false },
      { id: bmrId(), name: 'CYN CYLAROS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DAGUPAN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DBMCI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DECATHLON', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DELTAPATH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DELTAPATH MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DEMO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DEOCAMPO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DESIERTO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DEUTSCHE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DIAB', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DIALOGX', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DIDLOGIC', color: '', allowBalance: 300, allowBalanceLabel: '$300 debit', hidden: false },
      { id: bmrId(), name: 'CYN DIDLOGIC BYOC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DIDLOGIC GSM', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DIGITALMINDS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DINEZ', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DM WENCESLAO', color: '', allowBalance: 10, allowBalanceLabel: '$10 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN DOHLE', color: '', allowBalance: 0, allowBalanceLabel: 'DO NOT BLOCK UNTIL ADVISE', hidden: false },
      { id: bmrId(), name: 'CYN DOMUS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DRAGON', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DRAGON PALACE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN DSIT', color: '', allowBalance: 0, allowBalanceLabel: 'mrc', hidden: false },
      { id: bmrId(), name: 'CYN ED', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN EKBET', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN EMERITUS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ENGINE RS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ETIQA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN EXODO', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN EXOTEL', color: '', allowBalance: 100, allowBalanceLabel: 'RCBC - Please add $100 soft credit. For POC only', hidden: false },
      { id: bmrId(), name: 'CYN EXOTEL POC', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN EXVOILINE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN FASTTAT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN FBPO WirelessDNA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN FBPO-AvonBLAST', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN FBPO-HOLCIM CEMENT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN FINANCE-TECH', color: '', allowBalance: 10, allowBalanceLabel: '$10 soft credit', hidden: false },
      { id: bmrId(), name: 'CYN FIRST DIGITAL', color: '', allowBalance: 2000, allowBalanceLabel: '$2000 credit advise Ana/Christine', hidden: false },
      { id: bmrId(), name: 'CYN FOCUSINC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN FusionBPO-Avon', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN FVL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GEAR INC', color: '', allowBalance: 0, allowBalanceLabel: 'wala allowable, but do not block as per bea', hidden: false },
      { id: bmrId(), name: 'CYN GENPACT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GETKLEAN VOIP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GLOBAL PO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GLOBALB2B', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GLOBALOS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GLOBALQUEST', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GLOBOASIATICO', color: '', allowBalance: 1500, allowBalanceLabel: '$1500 debit\n$5000 debit (temporary) as of Nov 24, 2025', hidden: false },
      { id: bmrId(), name: 'CYN GLORYCOMM', color: '', allowBalance: 2500, allowBalanceLabel: '$2,500 debit', hidden: false },
      { id: bmrId(), name: 'CYN GoAutoDial', color: '', allowBalance: 20, allowBalanceLabel: '$20 credit advise Client\n$50 debit & no payment, block', hidden: false },
      { id: bmrId(), name: 'CYN GOLDMAN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GORMSBY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GPAY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GRAB FIN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GRABASIAKREDIT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GRACEONE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GRALVAT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GRANDWEST', color: '', allowBalance: 100, allowBalanceLabel: '$100 debit advise Sir J', hidden: false },
      { id: bmrId(), name: 'CYN GTH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN GTS', color: '', allowBalance: 2000, allowBalanceLabel: '$2000 debit (postpaid)', hidden: false },
      { id: bmrId(), name: 'CYN HCS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN HCX', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN HDFINITE', color: '', allowBalance: 0, allowBalanceLabel: 'Block once debit', hidden: false },
      { id: bmrId(), name: 'CYN HEADSTRONG', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN HELLOWORLD', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN HEROES', color: '', allowBalance: 0, allowBalanceLabel: 'currently blocked', hidden: false },
      { id: bmrId(), name: 'CYN HGC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN HIRATEL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN HOMEHEALTH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN HORIZON', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ICAVHOST', color: '', allowBalance: 0, allowBalanceLabel: 'Block once debit', hidden: false },
      { id: bmrId(), name: 'CYN IMC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN IMONEY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INCOHO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INETCOM', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN', color: '', allowBalance: 1200, allowBalanceLabel: '$1200 debit advise bithiah\n$1500 debit all infinivan', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN FUJITRANS', color: '', allowBalance: 5, allowBalanceLabel: '$5 soft credit', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN HIS', color: '', allowBalance: 1200, allowBalanceLabel: '$1200 debit advise bithiah\n$1500 debit all infinivan', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN HIS CEBU', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN HSP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN JRI', color: '', allowBalance: 1200, allowBalanceLabel: '$1200 debit advise bithiah\n$1500 debit all infinivan', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN KOSHIN2', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN KURABE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN SABASEI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN STERLING', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN-DRAGON ROYALE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN-MC LOGISTICS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN-QUICKWAY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INFINIVAN-SANKO GOSEI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INFOCOM', color: '', allowBalance: 50, allowBalanceLabel: '$50 debit', hidden: false },
      { id: bmrId(), name: 'CYN INFOCOM MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN INLIFE', color: '', allowBalance: 0, allowBalanceLabel: 'poc', hidden: false },
      { id: bmrId(), name: 'CYN INTELL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN IRIGATEL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ISOC', color: '', allowBalance: 10, allowBalanceLabel: '$10 debit (poc)', hidden: false },
      { id: bmrId(), name: 'CYN ITG', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ITNIO TECH', color: '', allowBalance: 2000, allowBalanceLabel: '$2000 debit (postpaid)', hidden: false },
      { id: bmrId(), name: 'CYN ITVAS', color: '', allowBalance: 0, allowBalanceLabel: 'advance billing', hidden: false },
      { id: bmrId(), name: 'CYN ITWORKS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN IXL', color: '', allowBalance: 0, allowBalanceLabel: 'DO NOT BLOCK UNTIL FUTHER NOTICE', hidden: false },
      { id: bmrId(), name: 'CYN JASCOR', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN JDM', color: '', allowBalance: 2, allowBalanceLabel: '$2 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN JIREH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN JMB', color: '', allowBalance: 0, allowBalanceLabel: 'poc', hidden: false },
      { id: bmrId(), name: 'CYN JOY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN JRF', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN JTMS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN KDDI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN KIMSTORE', color: '', allowBalance: 0, allowBalanceLabel: 'mrc', hidden: false },
      { id: bmrId(), name: 'CYN KINGSFORD', color: '', allowBalance: 2000, allowBalanceLabel: '$2000 debit', hidden: false },
      { id: bmrId(), name: 'CYN KITAL', color: '', allowBalance: 5000, allowBalanceLabel: '$5000 debit (postpaid)', hidden: false },
      { id: bmrId(), name: 'CYN KMC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN KONEK-IT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN KORE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN KTP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN LARX', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN LBC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN LEDWORKS', color: '', allowBalance: 0, allowBalanceLabel: 'mrc', hidden: false },
      { id: bmrId(), name: 'CYN LEGION CREDIT', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit', hidden: false },
      { id: bmrId(), name: 'CYN LGOSOFT', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit', hidden: false },
      { id: bmrId(), name: 'CYN LINSEED', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN LINSEED MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN LIVEWIRE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN LORRAINE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN LRDATA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN M2INTERNATIONAL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MABROX', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MALACASH', color: '', allowBalance: 10, allowBalanceLabel: '$10 debit (poc)', hidden: false },
      { id: bmrId(), name: 'CYN MARVDATA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MASHIN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MASHUNKASIA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MAXICARE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MAYA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MCCS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MCCS.2', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MCGS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MCM', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN MEDGROCER', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MEDIA', color: '', allowBalance: 200, allowBalanceLabel: '$200 debit', hidden: false },
      { id: bmrId(), name: 'CYN MEDIA MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MERCURY-FII', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MERJJ', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN METACOM', color: '', allowBalance: 0, allowBalanceLabel: 'currently blocked', hidden: false },
      { id: bmrId(), name: 'CYN MHR', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MICROENSURE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MONEYMAX', color: '', allowBalance: 0, allowBalanceLabel: 'Postpaid', hidden: false },
      { id: bmrId(), name: 'CYN MONEYMAX2', color: '', allowBalance: 10, allowBalanceLabel: '$10 debit', hidden: false },
      { id: bmrId(), name: 'CYN MONTNETS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MOOLA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MOOLA2', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MORPHEUS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MPHTC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MPOTECH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MSEED', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MSI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MTB', color: '', allowBalance: 15, allowBalanceLabel: '$15 debit', hidden: false },
      { id: bmrId(), name: 'CYN MY TAXIPH', color: '', allowBalance: 0, allowBalanceLabel: 'Postpaid', hidden: false },
      { id: bmrId(), name: 'CYN MYLESTONE BPO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MYTAXIPH2', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN MYVELOX', color: '', allowBalance: 15, allowBalanceLabel: '$15 debit', hidden: false },
      { id: bmrId(), name: 'CYN MYVELOX MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NALURI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NAUTILUS', color: '', allowBalance: 2, allowBalanceLabel: '$2 debit (for testing)', hidden: false },
      { id: bmrId(), name: 'CYN NEARSOL', color: '', allowBalance: 0, allowBalanceLabel: 'DO NOT BLOCK UNTIL FUTHER ADVISE', hidden: false },
      { id: bmrId(), name: 'CYN NEIL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NETCHASE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NETIX', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NETVOICE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NMB', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NOTUS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NSYS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NTU', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NUTRAL HERBAL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN NZASSIST', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN OFCC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN OLIVIA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN OLP', color: '', allowBalance: 6000, allowBalanceLabel: '$6000 debit', hidden: false },
      { id: bmrId(), name: 'CYN OLP MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN OLP2 MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN OMSMPC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ONE FELICITY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ONE VISAYA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ONESEARCHPRO', color: '', allowBalance: 0, allowBalanceLabel: '1K debit', hidden: false },
      { id: bmrId(), name: 'CYN ONET', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ONETOUCH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN OODC', color: '', allowBalance: 10, allowBalanceLabel: '$10 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN OOSCORP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN OPHIR', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN OPTIMUM', color: '', allowBalance: 0, allowBalanceLabel: 'updated invoice', hidden: false },
      { id: bmrId(), name: 'CYN OPULENCE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN OTG', color: '', allowBalance: 0, allowBalanceLabel: 'Block once debit', hidden: false },
      { id: bmrId(), name: 'CYN OTG2', color: '', allowBalance: 0, allowBalanceLabel: 'Block once debit', hidden: false },
      { id: bmrId(), name: 'CYN OTP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PAL', color: '', allowBalance: 62, allowBalanceLabel: '$62 debit', hidden: false },
      { id: bmrId(), name: 'CYN PANDR', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PAY8', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PAYMEINDIA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PEGER', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PEMC', color: '', allowBalance: 10, allowBalanceLabel: '$10 debit\ncurrently blocked', hidden: false },
      { id: bmrId(), name: 'CYN PEMC AH', color: '', allowBalance: 6500, allowBalanceLabel: '$6500 debit advise Sir CJ\n$7000 auto block', hidden: false },
      { id: bmrId(), name: 'CYN PEMC MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PEREIRA', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit (poc)', hidden: false },
      { id: bmrId(), name: 'CYN PHILTOWER', color: '', allowBalance: 0, allowBalanceLabel: 'mrc', hidden: false },
      { id: bmrId(), name: 'CYN PISCORP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PLAYMATE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PNBHOLDINGS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PONG', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN POWERVISION', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PRABEL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PRFCI MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PRIME-NJPA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PRIMMT PHARMA', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit', hidden: false },
      { id: bmrId(), name: 'CYN PRINCESS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PRLI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PROJECTOS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN PUREIP', color: '', allowBalance: 10, allowBalanceLabel: '$10 debit', hidden: false },
      { id: bmrId(), name: 'CYN PURPLE ROSE', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit', hidden: false },
      { id: bmrId(), name: 'CYN QE SOL', color: '', allowBalance: 0, allowBalanceLabel: 'mrc', hidden: false },
      { id: bmrId(), name: 'CYN QWERTY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN RAYUSA', color: '', allowBalance: 0, allowBalanceLabel: 'currently blocked', hidden: false },
      { id: bmrId(), name: 'CYN RBS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN RCGS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN RCGS2', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN RD SOLUTIONS', color: '', allowBalance: 0, allowBalanceLabel: 'currently blocked', hidden: false },
      { id: bmrId(), name: 'CYN RELIANCE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN RESULTS MASTERS', color: '', allowBalance: 0, allowBalanceLabel: 'currently blocked', hidden: false },
      { id: bmrId(), name: 'CYN RGB', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN RGB IR', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN RGS', color: '', allowBalance: 20, allowBalanceLabel: '$20 debit (poc)', hidden: false },
      { id: bmrId(), name: 'CYN RGS-DOT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN RIHNO', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN RIPTEC', color: '', allowBalance: 3000, allowBalanceLabel: '$3000 debit (postpaid)', hidden: false },
      { id: bmrId(), name: 'CYN RLINK', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ROBOCASH', color: '', allowBalance: 500, allowBalanceLabel: '$500 debit', hidden: false },
      { id: bmrId(), name: 'CYN RSU', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN RVN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SAVANT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SAXONMSI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SEAD', color: '', allowBalance: 0, allowBalanceLabel: 'currently blocked', hidden: false },
      { id: bmrId(), name: 'CYN SERVICE4U', color: '', allowBalance: 2000, allowBalanceLabel: '$2000 debit', hidden: false },
      { id: bmrId(), name: 'CYN SGDLABS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SHAN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SHAN OUTBOUND', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SIEGREICH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SIEGREICH AZ', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SIEGREICH KL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SIEGREICH TAIWAN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SIEGREICH VIETNAM', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SIGMAPAY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SKYLINE', color: '', allowBalance: 1800, allowBalanceLabel: '$1,800 advise client\n$2000 block', hidden: false },
      { id: bmrId(), name: 'CYN SMARTWORK', color: '', allowBalance: 20, allowBalanceLabel: '$20 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN SOLIDDOUBLE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SPEEDVOIP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SPM A-Z', color: '', allowBalance: 1000, allowBalanceLabel: '$1000 debit, advise client on gc\nDO NOT BLOCK UNTIL ADVISE (NO LIMIT)', hidden: false },
      { id: bmrId(), name: 'CYN SPMADRID', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ST24', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN STOPMANUALDIAL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN STRATPRO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SUPPLYSTATION', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SUPPLYSTATION2', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SUPPLYSTATION3', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SYNCCENTRAL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SYNERG', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN SYSTRONIC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN T2G', color: '', allowBalance: 0, allowBalanceLabel: 'mrc', hidden: false },
      { id: bmrId(), name: 'CYN TAAS', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit (poc)\nblocked outbound', hidden: false },
      { id: bmrId(), name: 'CYN TAKENAKA-CAD', color: '', allowBalance: 0, allowBalanceLabel: 'postpaid', hidden: false },
      { id: bmrId(), name: 'CYN TAKENAKA-CIVIL', color: '', allowBalance: 0, allowBalanceLabel: 'postpaid', hidden: false },
      { id: bmrId(), name: 'CYN TALKPUSH', color: '', allowBalance: 20000, allowBalanceLabel: '$20000 debit', hidden: false },
      { id: bmrId(), name: 'CYN TCHGLB-SYCIP LAW', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TCN INC', color: '', allowBalance: 0, allowBalanceLabel: 'Block once debit', hidden: false },
      { id: bmrId(), name: 'CYN TECHNO', color: '', allowBalance: 0, allowBalanceLabel: 'currently blocked', hidden: false },
      { id: bmrId(), name: 'CYN TECHNO DREAM', color: '', allowBalance: 20, allowBalanceLabel: '$20 credit advise client\n$20 debit block', hidden: false },
      { id: bmrId(), name: 'CYN TECHPLAIN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TECHSUPP', color: '', allowBalance: 0, allowBalanceLabel: 'mrc', hidden: false },
      { id: bmrId(), name: 'CYN TEKTITE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TELAN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TELIAX', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TELIVOZ', color: '', allowBalance: 20, allowBalanceLabel: '$20 debit advise bea\n$50, $100 debit advise Client\n$150 debit block if no payment', hidden: false },
      { id: bmrId(), name: 'CYN TELNET', color: '', allowBalance: 0, allowBalanceLabel: 'Block once debit', hidden: false },
      { id: bmrId(), name: 'CYN TELTEL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TENDOPAY', color: '', allowBalance: 0, allowBalanceLabel: 'WITH SECURITY DEPOSIT', hidden: false },
      { id: bmrId(), name: 'CYN TESCOLAB', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TEXXEN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TEXXEN MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN THOUGHTFOCUS', color: '', allowBalance: 0, allowBalanceLabel: 'mrc', hidden: false },
      { id: bmrId(), name: 'CYN TI-NET', color: '', allowBalance: 1000, allowBalanceLabel: '$1000 debit', hidden: false },
      { id: bmrId(), name: 'CYN TIFIA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TITANS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TKACZ', color: '', allowBalance: 20, allowBalanceLabel: '$20 debit', hidden: false },
      { id: bmrId(), name: 'CYN TOKU', color: '', allowBalance: 0, allowBalanceLabel: 'BEFORE DEBIT SHOULD BE BLOCK ( MAKE SURE TO EMAIL THEM IF THEY PRE-PAYMENT IS REACH BY 30%)', hidden: false },
      { id: bmrId(), name: 'CYN TONIK', color: '', allowBalance: 8000, allowBalanceLabel: '$8000 debit', hidden: false },
      { id: bmrId(), name: 'CYN TOYOTA TSUSHO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TRADEPEDIA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TRAIL LTD', color: '', allowBalance: 100, allowBalanceLabel: '$100 debit', hidden: false },
      { id: bmrId(), name: 'CYN TRAVERSE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TRENDS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TRG', color: '', allowBalance: 0, allowBalanceLabel: 'mrc', hidden: false },
      { id: bmrId(), name: 'CYN TRI7SOLUTIONS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TRI7SOLUTIONS 2', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TRIPLE DIAMOND', color: '', allowBalance: 0, allowBalanceLabel: 'advise client if debit', hidden: false },
      { id: bmrId(), name: 'CYN TRUEGRIT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TWILIO', color: '', allowBalance: 250000, allowBalanceLabel: '$250000 debit', hidden: false },
      { id: bmrId(), name: 'CYN TWILIO TRAINING', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN TWINLAKES', color: '', allowBalance: 100, allowBalanceLabel: '$100 debit', hidden: false },
      { id: bmrId(), name: 'CYN UDERNAL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN UNACASH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN UNAMARKET', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN UNIONBANK', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN UNOBANK', color: '', allowBalance: 8000, allowBalanceLabel: '$8000 debit advise am', hidden: false },
      { id: bmrId(), name: 'CYN UNOBANK PBX', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN UNOCOLLECTIONS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN UPTEL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN UTPI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VALLEYTECH', color: '', allowBalance: 300, allowBalanceLabel: '$300 debit', hidden: false },
      { id: bmrId(), name: 'CYN VAULT TEL', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit', hidden: false },
      { id: bmrId(), name: 'CYN VEGAGLOBAL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VENTEL', color: '', allowBalance: 10000, allowBalanceLabel: '$10000 debit\n$20000 extended debit', hidden: false },
      { id: bmrId(), name: 'CYN VENTIRA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VERGE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VESPERTELECOM', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VIRTUEDGE TECH', color: '', allowBalance: 0, allowBalanceLabel: 'poc, inbound', hidden: false },
      { id: bmrId(), name: 'CYN VIRTUO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VISION', color: '', allowBalance: 0, allowBalanceLabel: 'Block once debit', hidden: false },
      { id: bmrId(), name: 'CYN VITEL365', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VMCT', color: '', allowBalance: 10000, allowBalanceLabel: '$10000 debit', hidden: false },
      { id: bmrId(), name: 'CYN VOCTIV', color: '', allowBalance: 11, allowBalanceLabel: '$11 debit (temp)', hidden: false },
      { id: bmrId(), name: 'CYN VOICEHOLD', color: '', allowBalance: 0, allowBalanceLabel: 'advise am if debit', hidden: false },
      { id: bmrId(), name: 'CYN VOICEHOLD2', color: '', allowBalance: 0, allowBalanceLabel: 'Block once debit', hidden: false },
      { id: bmrId(), name: 'CYN VOICEHOLD3', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VOICEHOLD4', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VOIP POC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VOIP4U', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VOIPHY', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VOIPHYNX', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VOIPTECH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VOS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN VOXIMPLANT', color: '', allowBalance: 5000, allowBalanceLabel: '$5000 credit advise client\n$40000 debit advise am', hidden: false },
      { id: bmrId(), name: 'CYN VOXIMPLANT GSM', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN WAVECELL88', color: '', allowBalance: 20000, allowBalanceLabel: '$20000 debit', hidden: false },
      { id: bmrId(), name: 'CYN WESTERN', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit (poc)', hidden: false },
      { id: bmrId(), name: 'CYN WGS', color: '', allowBalance: 800, allowBalanceLabel: '$800 debit advise am', hidden: false },
      { id: bmrId(), name: 'CYN WGS MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN WGS1', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN WIZ', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN YELLOW AI', color: '', allowBalance: 0, allowBalanceLabel: 'postpaid', hidden: false },
      { id: bmrId(), name: 'CYN ZATHURA', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ZOOM', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ZOOM AQUIS', color: '', allowBalance: 50, allowBalanceLabel: '$50 debit (outbound)', hidden: false },
      { id: bmrId(), name: 'CYN ZOOM AVVANZ', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN ZOOM EECHECK', color: '', allowBalance: 10, allowBalanceLabel: '$10 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN ZOOM GERI', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN ZOOM HEALTHY OPTION', color: '', allowBalance: 10, allowBalanceLabel: '$10 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN ZOOM WRISE', color: '', allowBalance: 10, allowBalanceLabel: '$10 debit (soft credit)', hidden: false },
      { id: bmrId(), name: 'CYN ZORRA', color: '', allowBalance: 5, allowBalanceLabel: '$5 debit', hidden: false },
      { id: bmrId(), name: 'CYN1 JOSIAH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN1 LALAFOOD', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN1 NMS', color: '', allowBalance: 0, allowBalanceLabel: 'DO NOT BLOCK UNTIL ADVISE', hidden: false },
      { id: bmrId(), name: 'CYN1 NMS CC', color: '', allowBalance: 0, allowBalanceLabel: 'DO NOT BLOCK UNTIL ADVISE', hidden: false },
      { id: bmrId(), name: 'CYN1 NMS-GE', color: '', allowBalance: 0, allowBalanceLabel: 'DO NOT BLOCK UNTIL ADVISE', hidden: false },
      { id: bmrId(), name: 'CYN1 NORTHSTAR', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN1 NOVA', color: '', allowBalance: 10, allowBalanceLabel: '$10 debit', hidden: false },
      { id: bmrId(), name: 'CYN1 NOVA MRC', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN1 ROYALCABLE', color: '', allowBalance: 0, allowBalanceLabel: 'DO NOT BLOCK UNTIL ADVISE', hidden: false },
      { id: bmrId(), name: 'CYN1 SKYKOMISH', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'CYN1 TIM', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'ECHO', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'ELISHATEL', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'IPVOICE', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'KHIDESIGN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'OPENACCESS', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'RFTECH PP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'SUNPAGE-PP', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'TARADESIGN', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'TRI7', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'TRI7.2', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'WANCOMM', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
      { id: bmrId(), name: 'ZAPPORT', color: '', allowBalance: 0, allowBalanceLabel: '', hidden: false },
    ];

    const DEFAULT_BMR_STATE = {
      clients: DEFAULT_BMR_CLIENTS,
      accountManagers: [],
      // Rules for Amnt Receivable target columns (L, X, AJ, ...)
      targetRules: [
        { id: bmrId('tr'), name: 'Reached allowable balance', clientId: DEFAULT_BMR_CLIENTS.find(c => c.name === 'COMMSOL-DID')?.id || DEFAULT_BMR_CLIENTS[0]?.id || '', kind: 'gte', value: 5, color: '#EF4444', fontColor: '#FFFFFF', bold: true, italic: false, underline: false, enabled: true },
      ],
      // Rules for 30-min usage columns (M, Y, AK, ...)
      usageRules: [
        { id: bmrId('ur'), name: 'Heavy debit (-0.001 to -49)', kind: 'between', min: -49, max: -0.001, color: '#F59E0B', fontColor: '#1F2937', bold: true, italic: false, underline: false, enabled: true },
      ],
      hiddenCols: [],
      hiddenBlockCols: [],
      colorBlockCarrierNames: true,
      bordersOnData: true,
      borderStyle: 'thin',
      timeSeparatorBorderStyle: 'thick',
      borderColor: 'FF374151',
      includeDay:   true,
      includeNight: true,
      notes: '',
    };

    const BMR_SMS_RETAIL_CLIENT_ROWS = [["CYN 168WATC",""],["CYN 168WATC2",""],["CYN 168WATC3",""],["CYN 2XLUCKY",""],["CYN 747",""],["CYN 88CASH","1000000.0"],["CYN ACTIVEONE",""],["CYN AEI",""],["CYN AHS",""],["CYN AHS2",""],["CYN AIQON",""],["CYN ALLENSMS",""],["CYN AMEYO QATAR",""],["CYN AMND","$2 debit"],["CYN ANYBET","$5 debit"],["CYN ASIALINK","$2 debit"],["CYN AZN-SKY",""],["CYN BILLESE","$10 debit"],["CYN BRGYBEL-AIR","$1000 debit"],["CYN CASH EXPRESS","$10 debit"],["CYN CEPAT KREDIT","$1200 debit"],["CYN CRISPY",""],["CYN CS TEST",""],["CYN D88",""],["CYN DEMO",""],["CYN DRAGON ROYALE",""],["CYN DWC","$2 debit"],["CYN EU9",""],["CYN EXOTEL-ASURION",""],["CYN FACHAI",""],["CYN FEETSMART",""],["CYN FGH",""],["CYN FINASIATECHSMS","$40 debit advise Client\n$60 debit block"],["CYN FOVTY","$5000 debit"],["CYN FRANZ TEST",""],["CYN FUSION BPO",""],["CYN GAMBIT",""],["CYN GDM",""],["CYN GetKleanSMS","$300 debit"],["CYN GG",""],["CYN GREATMANGO",""],["CYN HOLOJILI",""],["CYN HOLOJILI 2",""],["CYN IB8",""],["CYN INT-SB365",""],["CYN JILI711",""],["CYN JLS",""],["CYN JLSOTP",""],["CYN KARECO",""],["CYN KasadoBet",""],["CYN KasadoOTP",""],["CYN KITAL",""],["CYN LPIT2",""],["CYN LUCKYME PLUS","$35.71 debit"],["CYN LUXE",""],["CYN MalaCash","$100 debit"],["CYN MOTION2",""],["CYN MWG",""],["CYN NEWFENIX",""],["CYN OchoWin",""],["CYN OLP","$5500 debit"],["CYN OODC",""],["CYN PANDR SMS",""],["CYN PARE",""],["CYN PEMC",""],["CYN PinasWins",""],["CYN PinoyGo",""],["CYN PJC SMS",""],["CYN PLAYVERSE",""],["CYN POWERPLAY",""],["CYN SAMLUK",""],["CYN SG8","$500 debit"],["CYN SG82","$500 debit"],["CYN SGDLABS",""],["CYN SPM SMS","$6900 debit"],["CYN SWN",""],["CYN SWN10",""],["CYN SWN11",""],["CYN SWN12",""],["CYN SWN13",""],["CYN SWN14",""],["CYN SWN15",""],["CYN SWN16",""],["CYN SWN17",""],["CYN SWN18",""],["CYN SWN19",""],["CYN SWN2",""],["CYN SWN20",""],["CYN SWN21",""],["CYN SWN22",""],["CYN SWN23",""],["CYN SWN24",""],["CYN SWN25",""],["CYN SWN26",""],["CYN SWN27",""],["CYN SWN28",""],["CYN SWN29",""],["CYN SWN3",""],["CYN SWN30",""],["CYN SWN31",""],["CYN SWN32",""],["CYN SWN33",""],["CYN SWN34",""],["CYN SWN35",""],["CYN SWN36",""],["CYN SWN37",""],["CYN SWN38",""],["CYN SWN39",""],["CYN SWN4",""],["CYN SWN40",""],["CYN SWN41",""],["CYN SWN42",""],["CYN SWN43",""],["CYN SWN44",""],["CYN SWN45",""],["CYN SWN46",""],["CYN SWN47",""],["CYN SWN48","$1300 debit"],["CYN SWN49","$1000 debit"],["CYN SWN5",""],["CYN SWN50",""],["CYN SWN51",""],["CYN SWN52",""],["CYN SWN53","$5 debit"],["CYN SWN54",""],["CYN SWN6",""],["CYN SWN7",""],["CYN SWN8",""],["CYN SWN9",""],["CYN TALKPUSH","$10 credit advise Client\n$3000 debit"],["CYN TeleservSMS","$500 credit"],["CYN TEMASEK","$100 debit"],["CYN TopSky",""],["CYN TopSkyMKT",""],["CYN TRG",""],["CYN ULTRA",""],["CYN WINFORDBET",""],["CYN WINFORDBET2",""],["SMS CYN QILIANG",""]];
    const BMR_SMS_WHOLESALE_CLIENT_ROWS = [["C_SMS-CYN AIQON","$5000 debit"],["C_SMS-CYN DEMO",""],["C_SMS-CYN GAEI-TF",""],["C_SMS-CYN ITVAS",""],["C_SMS-CYN PRLI",""],["C_SMS_CYN RIZAL PROV","$5 debit"],["C_SMS-CYN SUNORO",""],["C_SMS-CYN SWN",""],["C_SMS-CYN SWN2",""],["C_SMS-CYN SWN3",""],["C_SMS-CYN SWN4",""],["C_SMS-CYN SWN5",""],["C_SMS-CYN SWN6",""],["C_SMS-CYN SWN7",""],["C_SMS-CYN SWN8",""],["C_SMS-CYN TEMASEK","$200 debit"],["C_SMS-CYN TEST NOC",""],["C_SMS-CYN TEST OTP",""],["C_SMS-CYN TEST2",""],["C_SMS-CYN TEXCELL",""],["C_CYN 168WATC3",""],["C_SMS-CYN",""],["C_SMS-CYN GAEI",""],["C_SMS-CYN HEYSCLOUDSMS",""],["C_SMS-CYN KITAL","$2000 debit"],["C_SMS-CYN OODC",""],["C_SMS-CYN ROYALCABLE",""],["C_SMS-CYN VTS",""],["C_SMS-CYN 1Cyxdynamiq",""],["C_SMS-CYN 2Cyxdynamiq","$82 debit"],["C_SMS-CYN ASMSC PHOENIX",""],["C_SMS-CYN CHONRY","$24 debit"],["C_SMS-CYN CLOOPENSMS",""],["C_SMS-CYN DIDLOGIC",""],["C_SMS-CYN ELFO","$3000 debit"],["C_SMS-CYN EXOTEL","$40000 debit"],["C_SMS-CYN GLORY","$5000 debit"],["C_SMS-CYN GREENPACKET","$2000 debit"],["C_SMS-CYN GTS",""],["C_SMS-CYN IBAZAAR",""],["C_SMS-CYN ILOMILO",""],["C_SMS-CYN INNOVOCOM",""],["C_SMS-CYN ITNIO","$20000 debit"],["C_SMS-CYN ITNIO INTL","$2000 debit"],["C_SMS-CYN ITVAS",""],["C_SMS-CYN KING PH","$7000 debit"],["C_SMS-CYN LANCK",""],["C_SMS-CYN LPIT",""],["C_SMS-CYN M360",""],["C_SMS-CYN MCT",""],["C_SMS-CYN MONTY MOBILE",""],["C_SMS-CYN MOTION",""],["C_SMS-CYN NEXTGEN","$3000 debit"],["C_SMS-CYN NXTGN S&C",""],["C_SMS-CYN OAKTEL",""],["C_SMS-CYN PLUTO",""],["C_SMS-CYN SHENTECH",""],["C_SMS-CYN SKYLINE","$40000 debit"],["C_SMS-CYN SKYLINE DOMESTIC","$2000 debit"],["C_SMS-CYN SKYLINE MRKTNG",""],["C_SMS-CYN SMART TECH","$2000 debit"],["C_SMS-CYN SOLUTIONS4U","$100 debit"],["C_SMS-CYN SUNORO","$2000 debit"],["C_SMS-CYN TELESIGN",""],["C_SMS-CYN TESTELIUM (SMS Test Tool)",""],["C_SMS-CYN TEXCELL","$1000 debit"],["C_SMS-CYN TIG",""],["C_SMS-CYN TIG",""],["C_SMS-CYN TIGO","$3000 debit"],["C_SMS-CYN TM",""],["C_SMS-CYN TOKU",""],["C_SMS-CYN TOKU2",""],["C_SMS-CYN TOPYING",""],["C_SMS-CYN TOPYING",""],["C_SMS-CYN TWILIO",""],["C_SMS-CYN ULTRAMARINE","$3000 debit"],["C_SMS-CYN VOIPCONNECT",""],["C_SMS-CYN VONAGE","$5000 debit"],["C_SMS-CYN VOXCO","$1000 debit"],["C_SMS-CYN VTS",""],["C_SMS-CYN VTSI",""],["C_SMS-CYN VTSMP","$4000 debit"],["C_SMS-CYN VTSMP 2","$1200 debit"],["C_SMS-CYN VTSMP 3",""],["C_SMS-CYN VTSMP 4",""],["C_SMS-CYN VTSMP 5",""],["C_SMS-CYN ZIFFY",""],["CYN VTSMP_2",""]];

    function bmrSmsClientsFromRows(rows, prefix) {
      return rows.map(([name, allowBalanceLabel], index) => ({
        id: `${prefix}_${String(index + 1).padStart(3, '0')}`,
        name,
        color: '',
        allowBalanceLabel: allowBalanceLabel || '',
        accountManagerId: '',
        hidden: false,
      }));
    }

    const DEFAULT_BMR_SMS_STATE = {
      retailClients: bmrSmsClientsFromRows(BMR_SMS_RETAIL_CLIENT_ROWS, 'sms_res'),
      wholesaleClients: bmrSmsClientsFromRows(BMR_SMS_WHOLESALE_CLIENT_ROWS, 'sms_whs'),
      accountManagers: [],
      targetRules: [],
      overdraftRules: [
        { id: bmrId('smsod'), name: 'Balance reached overdraft', color: '#EF4444', fontColor: '#FFFFFF', bold: true, italic: false, underline: false, enabled: true },
      ],
      usageRules: [
        { id: bmrId('smsur'), name: 'Heavy debit (-0.001 to -49)', kind: 'between', min: -49, max: -0.001, color: '#F59E0B', fontColor: '#1F2937', bold: true, italic: false, underline: false, enabled: true },
      ],
      colorBlockCarrierNames: true,
      boldBalanceOnUsage: true,
      bordersOnData: true,
      borderStyle: 'thin',
      timeSeparatorBorderStyle: 'thick',
      borderColor: 'FF374151',
      includeDay: true,
      includeNight: true,
      notes: '',
    };

    function BmrColorSwatch({ value, onChange, allowClear = true }) {
      return (
        <span className="inline-flex items-center gap-1.5">
          <PresetColorPicker value={value || '#000000'} onChange={onChange}
            title="Choose BMR color" showHex={false} buttonClassName="h-7 px-1.5 py-1" />
          {allowClear && value && (
            <button onClick={() => onChange('')} className="text-[10px] text-neutral-500 hover:text-red-300 px-1" title="Clear color">clear</button>
          )}
        </span>
      );
    }

    function BmrClientImporter({ bmr, onChange }) {
      const [busy, setBusy] = useState(false);
      const [downloadBusy, setDownloadBusy] = useState(false);
      const [sampleBusy, setSampleBusy] = useState(false);
      const [notice, setNotice] = useState('');
      const importClients = async (e) => {
        const file = e.target.files?.[0];
        e.target.value = '';
        if (!file) return;
        setBusy(true);
        setNotice('');
        try {
          const parsed = await parseBmrImportXlsx(file);
          if (!parsed.clients.length && !parsed.targetRules.length && !parsed.usageRules.length) {
            await alertDialog({
              title: 'No import rows found',
              message: 'Upload an .xlsx workbook with a Client Name / CarrierName sheet, Target Rules sheet, or 30-min Usage Rules sheet.',
              confirmText: 'OK',
            });
            return;
          }
          const mergedClients = mergeBmrImportedClients(bmr, parsed.clients);
          const mergedRules = mergeBmrImportedRules(mergedClients.bmr, parsed.targetRules, parsed.usageRules);
          onChange(mergedRules.bmr);
          const parts = [
            `${mergedClients.stats.added} clients added`,
            `${mergedClients.stats.updated} clients updated`,
            `${mergedRules.stats.targetAdded} target rules added`,
            `${mergedRules.stats.targetUpdated} target rules updated`,
            `${mergedRules.stats.usageAdded} usage rules added`,
            `${mergedRules.stats.usageUpdated} usage rules updated`,
          ];
          if (mergedRules.stats.targetSkipped) parts.push(`${mergedRules.stats.targetSkipped} target rules skipped (missing client)`);
          setNotice(parts.join(' · '));
        } catch (err) {
          console.warn('BMR client import failed', err);
          await alertDialog({
            title: "Couldn't import BMR data",
            message: "That file couldn't be read as a BMR client/rules .xlsx workbook.",
            confirmText: 'OK',
            tone: 'danger',
          });
        } finally {
          setBusy(false);
        }
      };
      const downloadSample = async () => {
        setSampleBusy(true);
        try {
          await downloadBmrClientImportSample();
        } catch (err) {
          console.warn('BMR client sample download failed', err);
          await alertDialog({
            title: "Couldn't create sample",
            message: "The client-import sample workbook couldn't be generated.",
            confirmText: 'OK',
            tone: 'danger',
          });
        } finally {
          setSampleBusy(false);
        }
      };
      const downloadData = async () => {
        setDownloadBusy(true);
        try {
          await downloadBmrClientData(bmr);
        } catch (err) {
          console.warn('BMR client data download failed', err);
          await alertDialog({
            title: "Couldn't download data",
            message: "The current client-data workbook couldn't be generated.",
            confirmText: 'OK',
            tone: 'danger',
          });
        } finally {
          setDownloadBusy(false);
        }
      };

      return (
        <div className="rounded-md border border-neutral-900 bg-neutral-950/60 p-3">
          <SectionLabel hint=".xlsx">Client information</SectionLabel>
          <div className="flex items-center gap-2 flex-wrap">
            <label
              title="Client sheet: Client Name or CarrierName. Rule sheets: Target Rules and 30-min Usage Rules."
              className={`inline-flex items-center justify-center gap-2 rounded-md border border-neutral-800 bg-neutral-950 px-3.5 py-2 text-sm font-medium text-neutral-300 transition-colors ${busy ? 'opacity-40 cursor-wait' : 'cursor-pointer hover:border-neutral-700 hover:bg-neutral-900 hover:text-neutral-100'}`}>
              {busy ? <span className="loader"></span> : <IconUpload />}
              {busy ? 'Importing' : 'Import clients/rules'}
              <input type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                onChange={importClients} disabled={busy} className="hidden" />
            </label>
            <Btn variant="ghost" size="md" onClick={downloadData} disabled={downloadBusy}>
              {downloadBusy ? <span className="loader"></span> : <IconDownload />}
              {downloadBusy ? 'Downloading' : 'Download data'}
            </Btn>
            <Btn variant="ghost" size="md" onClick={downloadSample} disabled={sampleBusy}>
              {sampleBusy ? <span className="loader"></span> : <IconDownload />}
              {sampleBusy ? 'Creating' : 'Sample format'}
            </Btn>
            {notice && <span className="text-[11px] text-emerald-300">{notice}</span>}
          </div>
        </div>
      );
    }

    function BmrBlockColumnHider({ bmr, onChange }) {
      const hidden = new Set((bmr.hiddenBlockCols || []).map(Number));
      const toggle = (offset) => {
        const next = hidden.has(offset)
          ? (bmr.hiddenBlockCols || []).filter(x => Number(x) !== offset)
          : [...(bmr.hiddenBlockCols || []), offset];
        onChange({ ...bmr, hiddenBlockCols: next });
      };

      return (
        <div className="rounded-md border border-neutral-900 bg-neutral-950/60 p-3">
          <SectionLabel hint={`${hidden.size}/${BMR_BLOCK_COLS} hidden in each time block`}>Downloaded columns</SectionLabel>
          <div className="grid gap-1.5 sm:grid-cols-2 lg:grid-cols-3">
            {BMR_BLOCK_HEADER_OPTIONS.map(col => (
              <label key={col.offset}
                className={`flex min-w-0 items-center gap-2 rounded-md border px-2.5 py-2 text-xs cursor-pointer transition-colors ${hidden.has(col.offset) ? 'border-amber-500/35 bg-amber-500/10 text-amber-100' : 'border-neutral-900 bg-neutral-950 text-neutral-300 hover:border-neutral-800'}`}>
                <input type="checkbox" checked={hidden.has(col.offset)} onChange={() => toggle(col.offset)}
                  className="h-3.5 w-3.5 rounded border-neutral-700 bg-neutral-900 text-amber-500" />
                <span className="truncate">{col.label}</span>
              </label>
            ))}
          </div>
        </div>
      );
    }

    function BmrBorderSettings({ bmr, onChange }) {
      const bordersOnData = bmr.bordersOnData !== false;
      const borderColor = bmr.borderColor || 'FF374151';
      const update = (patch) => onChange({ ...bmr, ...patch });
      const styleOptions = (
        <>
          <option value="thin">Thin</option>
          <option value="medium">Medium</option>
          <option value="thick">Thick</option>
          <option value="dotted">Dotted</option>
          <option value="dashed">Dashed</option>
        </>
      );

      return (
        <div className="rounded-md border border-neutral-900 bg-neutral-950/60 p-3">
          <SectionLabel hint="Generated workbook">Borders</SectionLabel>
          <label className="mb-3 flex cursor-pointer items-center gap-2 text-sm text-neutral-300">
            <input type="checkbox" checked={bordersOnData}
              onChange={e => update({ bordersOnData: e.target.checked })}
              className="h-3.5 w-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500 focus:ring-blue-500/30 cursor-pointer" />
            Add cell borders
          </label>
          {bordersOnData && (
            <div key="bmr-border-panel" className="anim-fade-in grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
              <div>
                <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">Cells</label>
                <Select value={bmr.borderStyle || 'thin'} onChange={e => update({ borderStyle: e.target.value })}>
                  {styleOptions}
                </Select>
              </div>
              <div>
                <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">Time blocks</label>
                <Select value={bmr.timeSeparatorBorderStyle || 'thick'} onChange={e => update({ timeSeparatorBorderStyle: e.target.value })}>
                  {styleOptions}
                </Select>
              </div>
              <div>
                <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">Color</label>
                <div className="flex items-center gap-1">
                  <PresetColorPicker value={'#' + borderColor.slice(2)}
                    onChange={color => update({ borderColor: 'FF' + color.slice(1) })}
                    title="Choose BMR border color" className="flex-1" />
                  {borderColor !== 'FF374151' && (
                    <button onClick={() => update({ borderColor: 'FF374151' })}
                      title="Reset to default (#374151)"
                      className="text-neutral-500 hover:text-neutral-200 transition-colors"><IconReset /></button>
                  )}
                </div>
              </div>
            </div>
          )}
        </div>
      );
    }

    function BmrCarrierNameSettings({ bmr, onChange }) {
      const colorBlockCarrierNames = bmr.colorBlockCarrierNames !== false;
      return (
        <div className="rounded-md border border-neutral-900 bg-neutral-950/60 p-3">
          <SectionLabel hint="Generated workbook">Carrier names</SectionLabel>
          <label className="flex cursor-pointer items-center gap-2 text-sm text-neutral-300">
            <input type="checkbox" checked={colorBlockCarrierNames}
              onChange={e => onChange({ ...bmr, colorBlockCarrierNames: e.target.checked })}
              className="h-3.5 w-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500 focus:ring-blue-500/30 cursor-pointer" />
            Match carrier colors in time blocks
          </label>
        </div>
      );
    }

    function BmrAccountManagersEditor({ bmr, onChange }) {
      const [draftName, setDraftName] = useState('');
      const accountManagers = bmr.accountManagers || [];
      const clientCountByManager = (bmr.clients || []).reduce((counts, client) => {
        if (client.accountManagerId) counts[client.accountManagerId] = (counts[client.accountManagerId] || 0) + 1;
        return counts;
      }, {});
      const updateAccountManager = useCallback((id, patch) => onChange((prev) => ({
        ...prev,
        accountManagers: (prev.accountManagers || []).map(manager =>
          manager.id === id ? { ...manager, ...patch } : manager
        ),
      })), [onChange]);
      const removeAccountManager = useCallback((id) => onChange((prev) => ({
        ...prev,
        accountManagers: (prev.accountManagers || []).filter(manager => manager.id !== id),
        clients: (prev.clients || []).map(client =>
          client.accountManagerId === id ? { ...client, accountManagerId: '' } : client
        ),
      })), [onChange]);
      const addAccountManager = () => {
        const name = draftName.trim();
        if (!name) return;
        const hasName = accountManagers.some(manager => (manager.name || '').trim().toLowerCase() === name.toLowerCase());
        if (hasName) return;
        onChange((prev) => ({
          ...prev,
          accountManagers: [...(prev.accountManagers || []), { id: bmrId('am'), name, color: '' }],
        }));
        setDraftName('');
      };

      return (
        <div className="rounded-md border border-neutral-900 bg-neutral-950/60 p-3 space-y-2">
          <SectionLabel hint={`${accountManagers.length} saved`}>Account managers</SectionLabel>
          {accountManagers.length > 0 && (
            <div className="space-y-1">
              {accountManagers.map(manager => (
                <div key={manager.id} className="grid grid-cols-[minmax(180px,1fr)_72px_84px_28px] gap-2 items-center rounded-md px-1 py-1 hover:bg-neutral-900/40">
                  <Input value={manager.name || ''} onChange={e => updateAccountManager(manager.id, { name: e.target.value })} placeholder="Account manager" />
                  <span className="text-[10px] uppercase tracking-wide text-neutral-500 text-right">{clientCountByManager[manager.id] || 0} clients</span>
                  <span className="flex items-center justify-center">
                    <BmrColorSwatch value={manager.color} onChange={color => updateAccountManager(manager.id, { color })} />
                  </span>
                  <button onClick={() => removeAccountManager(manager.id)} className="p-1 text-neutral-500 hover:text-red-400" title="Remove account manager"><IconX /></button>
                </div>
              ))}
            </div>
          )}
          <div className="flex items-end gap-2 flex-wrap">
            <div className="flex-1 min-w-[200px]">
              <label className="block text-[10px] uppercase tracking-wide text-neutral-500 mb-1">New account manager</label>
              <Input value={draftName} onChange={e => setDraftName(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && (e.preventDefault(), addAccountManager())}
                placeholder="e.g. Jason Cruz" />
            </div>
            <Btn variant="ghost" size="md" onClick={addAccountManager} disabled={!draftName.trim()}>
              <IconPlus /> Add
            </Btn>
          </div>
        </div>
      );
    }

    const BmrClientRow = React.memo(function BmrClientRow({
      client, accountManagers, index, selected, dragIndex, overIndex, dragRow,
      setDragIndex, setOverIndex, setDragRow,
      reorder, updateClient, removeClient, toggleSelected,
    }) {
      const isDragging = dragIndex === index;
      const above = overIndex === index && dragIndex !== null && dragIndex !== index && dragIndex > index;
      const below = overIndex === index && dragIndex !== null && dragIndex !== index && dragIndex < index;
      const accountManager = accountManagers.find(manager => manager.id === client.accountManagerId);
      const clientColor = accountManager?.color || client.color || '';

      return (
        <div
          draggable={dragRow === index}
          onDragStart={(e) => { setDragIndex(index); e.dataTransfer.effectAllowed = 'move'; try { e.dataTransfer.setData('text/plain', String(index)); } catch (_) {} }}
          onDragOver={(e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; if (overIndex !== index) setOverIndex(index); }}
          onDragLeave={() => { if (overIndex === index) setOverIndex(null); }}
          onDrop={(e) => { e.preventDefault(); if (dragIndex !== null && dragIndex !== index) reorder(dragIndex, index); setDragIndex(null); setOverIndex(null); setDragRow(null); }}
          onDragEnd={() => { setDragIndex(null); setOverIndex(null); setDragRow(null); }}
          className={`grid grid-cols-[20px_20px_28px_28px_minmax(180px,1fr)_minmax(160px,1fr)_minmax(140px,180px)_28px_84px_28px] gap-2 items-center rounded-md px-1 py-1 transition-colors ${client.hidden ? 'opacity-50' : ''} ${selected ? 'bg-blue-500/10' : ''} ${isDragging ? 'opacity-30' : ''} ${above ? 'border-t-2 border-blue-500/70' : ''} ${below ? 'border-b-2 border-blue-500/70' : ''} hover:bg-neutral-900/40`}>
          <input type="checkbox" checked={selected} onChange={() => toggleSelected(client.id)}
            title="Select client"
            className="h-3.5 w-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500 focus:ring-blue-500/30 cursor-pointer" />
          <span
            onMouseDown={() => setDragRow(index)}
            onMouseUp={() => setDragRow(null)}
            onTouchStart={() => setDragRow(index)}
            onTouchEnd={() => setDragRow(null)}
            className="cursor-grab active:cursor-grabbing p-0.5 text-neutral-700 hover:text-neutral-300 select-none"
            title="Drag to reorder">
            <IconGrip />
          </span>
          <span className="text-[10px] text-neutral-600 font-mono text-right">{String(index + 1).padStart(2, '0')}</span>
          <span className="w-6 h-6 rounded border border-neutral-800" style={clientColor ? { background: clientColor } : {}} title={accountManager?.color ? `${accountManager.name || 'Account manager'} color` : (client.color || 'no color')}></span>
          <Input value={client.name} onChange={e => updateClient(client.id, { name: e.target.value })} placeholder="Client name" className={client.hidden ? 'line-through' : ''} />
          <Input value={client.allowBalanceLabel || ''} onChange={e => updateClient(client.id, { allowBalanceLabel: e.target.value })} placeholder="e.g. Block once debit" />
          <Select value={client.accountManagerId || ''} onChange={e => updateClient(client.id, { accountManagerId: e.target.value })}>
            <option value="">No manager</option>
            {accountManagers.map(manager => <option key={manager.id} value={manager.id}>{manager.name || 'Unnamed manager'}</option>)}
          </Select>
          <button onClick={() => updateClient(client.id, { hidden: !client.hidden })}
            className={`p-1 rounded ${client.hidden ? 'text-amber-400' : 'text-neutral-500 hover:text-neutral-200'}`}
            title={client.hidden ? 'Hidden — click to show' : 'Hide row in generated file'}>
            {client.hidden ? <IconEyeOff /> : <IconEye />}
          </button>
          <span className="flex items-center justify-center">
            <BmrColorSwatch value={client.color} onChange={color => updateClient(client.id, { color })} />
          </span>
          <button onClick={() => removeClient(client.id)} className="p-1 text-neutral-500 hover:text-red-400" title="Remove client"><IconX /></button>
        </div>
      );
    });

    function BmrClientsEditor({ bmr, onChange, showBlockColumnHider = true, showCarrierNameSettings = true }) {
      const [draftName, setDraftName] = useState('');
      const [insertAt, setInsertAt] = useState('');
      const [managerFilter, setManagerFilter] = useState('__all__');
      const [clientSearch, setClientSearch] = useState('');
      const [selectedClientIds, setSelectedClientIds] = useState([]);
      const [bulkManagerId, setBulkManagerId] = useState('__unchanged__');
      const [bulkColor, setBulkColor] = useState('#FACC15');
      const [dragIndex, setDragIndex] = useState(null);
      const [overIndex, setOverIndex] = useState(null);
      const [dragRow, setDragRow]   = useState(null);

      const clients = bmr.clients || [];
      const accountManagers = bmr.accountManagers || [];
      const accountManagersById = useMemo(() => new Map(accountManagers.map(manager => [manager.id, manager])), [accountManagers]);
      useEffect(() => {
        if (managerFilter === '__all__' || managerFilter === '__none__') return;
        if (!accountManagers.some(manager => manager.id === managerFilter)) setManagerFilter('__all__');
      }, [accountManagers, managerFilter]);
      useEffect(() => {
        if (bulkManagerId === '__unchanged__' || bulkManagerId === '__none__') return;
        if (!accountManagers.some(manager => manager.id === bulkManagerId)) setBulkManagerId('__unchanged__');
      }, [accountManagers, bulkManagerId]);
      useEffect(() => {
        const clientIds = new Set(clients.map(client => client.id));
        setSelectedClientIds((ids) => {
          const kept = ids.filter(id => clientIds.has(id));
          return kept.length === ids.length ? ids : kept;
        });
      }, [clients]);
      const clientSearchTerm = clientSearch.trim().toLowerCase();
      const filteredClients = clients
        .map((client, index) => ({ client, index }))
        .filter(({ client }) => {
          const managerMatches = managerFilter === '__all__'
            || (managerFilter === '__none__' ? !client.accountManagerId : client.accountManagerId === managerFilter);
          if (!managerMatches || !clientSearchTerm) return managerMatches;
          const managerName = accountManagersById.get(client.accountManagerId)?.name || '';
          return [client.name, client.allowBalanceLabel, managerName]
            .some(value => String(value || '').toLowerCase().includes(clientSearchTerm));
        });
      const selectedClientIdSet = new Set(selectedClientIds);
      const visibleClientIds = filteredClients.map(({ client }) => client.id);
      const allVisibleSelected = visibleClientIds.length > 0 && visibleClientIds.every(id => selectedClientIdSet.has(id));
      const setClients = useCallback((next) => onChange((prev) => ({
        ...prev,
        clients: typeof next === 'function' ? next(prev.clients || []) : next,
      })), [onChange]);
      const updateClient = useCallback((id, patch) => setClients((list) =>
        list.map((c) => c.id === id ? { ...c, ...patch } : c)
      ), [setClients]);
      const removeClient = useCallback((id) => setClients((list) =>
        list.filter((c) => c.id !== id)
      ), [setClients]);
      const addClient = (pos) => {
        const name = draftName.trim();
        if (!name) return;
        const newC = { id: bmrId(), name, color: '', allowBalanceLabel: '', accountManagerId: '', hidden: false };
        if (pos == null || pos < 0 || pos > clients.length) setClients([...clients, newC]);
        else setClients([...clients.slice(0, pos), newC, ...clients.slice(pos)]);
        setDraftName('');
        setInsertAt('');
      };
      const reorder = useCallback((from, to) => setClients((list) => {
        if (from === to || to < 0 || to >= list.length) return list;
        const arr = [...list];
        const [moved] = arr.splice(from, 1);
        arr.splice(to, 0, moved);
        return arr;
      }), [setClients]);
      const toggleSelected = useCallback((id) => setSelectedClientIds((ids) =>
        ids.includes(id) ? ids.filter(item => item !== id) : [...ids, id]
      ), []);
      const toggleVisibleSelection = () => setSelectedClientIds((ids) => {
        const next = new Set(ids);
        if (allVisibleSelected) visibleClientIds.forEach(id => next.delete(id));
        else visibleClientIds.forEach(id => next.add(id));
        return Array.from(next);
      });
      const updateSelectedClients = (patch) => {
        if (!selectedClientIds.length) return;
        setClients((list) => list.map(client => selectedClientIdSet.has(client.id) ? { ...client, ...patch } : client));
      };
      const applyBulkManager = () => {
        if (bulkManagerId === '__unchanged__') return;
        updateSelectedClients({ accountManagerId: bulkManagerId === '__none__' ? '' : bulkManagerId });
      };

      return (
        <div className="space-y-3">
          <BmrClientImporter bmr={bmr} onChange={onChange} />
          {showBlockColumnHider && <BmrBlockColumnHider bmr={bmr} onChange={onChange} />}
          {showCarrierNameSettings && <BmrCarrierNameSettings bmr={bmr} onChange={onChange} />}
          <BmrBorderSettings bmr={bmr} onChange={onChange} />
          <BmrAccountManagersEditor bmr={bmr} onChange={onChange} />
          <div className="flex flex-wrap items-end justify-between gap-3 rounded-md border border-neutral-900 bg-neutral-950/60 p-3">
            <div className="flex min-w-0 flex-1 flex-wrap items-end gap-3">
              <div className="w-full max-w-xs">
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Search clients</label>
                <ClearableInput value={clientSearch} onChange={e => setClientSearch(e.target.value)}
                  onClear={() => setClientSearch('')} clearTitle="Clear client search"
                  placeholder="Client, label, or manager" />
              </div>
              <div className="w-full max-w-xs">
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Account manager filter</label>
                <Select value={managerFilter} onChange={e => setManagerFilter(e.target.value)}>
                  <option value="__all__">All account managers</option>
                  <option value="__none__">No manager</option>
                  {accountManagers.map(manager => <option key={manager.id} value={manager.id}>{manager.name || 'Unnamed manager'}</option>)}
                </Select>
              </div>
            </div>
            <span className="pb-2 text-[11px] text-neutral-500">{filteredClients.length}/{clients.length} clients shown</span>
          </div>
          {selectedClientIds.length > 0 && (
            <div className="rounded-md border border-blue-500/40 bg-blue-500/10 p-3">
              <SectionLabel hint={`${selectedClientIds.length} selected`}>Bulk edit</SectionLabel>
              <div className="flex flex-wrap items-end gap-2">
                <div className="w-full max-w-xs">
                  <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Account manager</label>
                  <Select value={bulkManagerId} onChange={e => setBulkManagerId(e.target.value)}>
                    <option value="__unchanged__">Choose manager</option>
                    <option value="__none__">No manager</option>
                    {accountManagers.map(manager => <option key={manager.id} value={manager.id}>{manager.name || 'Unnamed manager'}</option>)}
                  </Select>
                </div>
                <Btn variant="ghost" size="md" onClick={applyBulkManager} disabled={bulkManagerId === '__unchanged__'}>Apply manager</Btn>
                <span className="inline-flex h-9 items-center gap-1 rounded-md border border-neutral-800 bg-neutral-950 px-2">
                  <span className="text-[10px] uppercase tracking-wide text-neutral-500">Client color</span>
                  <BmrColorSwatch value={bulkColor} onChange={setBulkColor} allowClear={false} />
                </span>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedClients({ color: bulkColor })}>Apply color</Btn>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedClients({ color: '' })}>Clear colors</Btn>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedClients({ hidden: true })}>Hide</Btn>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedClients({ hidden: false })}>Show</Btn>
                <Btn variant="ghost" size="md" onClick={() => setSelectedClientIds([])}>Clear selection</Btn>
              </div>
            </div>
          )}
          <div className="overflow-x-auto pb-1">
            <div className="min-w-[820px]">
              <div className="grid grid-cols-[20px_20px_28px_28px_minmax(180px,1fr)_minmax(160px,1fr)_minmax(140px,180px)_28px_84px_28px] gap-2 px-1 text-[10px] uppercase tracking-wide text-neutral-500">
                <div className="flex items-center">
                  <input type="checkbox" checked={allVisibleSelected} onChange={toggleVisibleSelection}
                    disabled={!visibleClientIds.length} title="Select shown clients"
                    className="h-3.5 w-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500 focus:ring-blue-500/30 cursor-pointer disabled:cursor-not-allowed disabled:opacity-40" />
                </div>
                <div></div><div></div><div></div>
                <div>Client name</div>
                <div>Allow Bal label</div>
                <div>Account manager</div>
                <div className="text-center">Hide</div>
                <div className="text-center">Client color</div>
                <div></div>
              </div>
              <div className="space-y-1">
                {filteredClients.map(({ client: c, index: i }) => (
                  <BmrClientRow key={c.id}
                    client={c}
                    accountManagers={accountManagers}
                    index={i}
                    selected={selectedClientIdSet.has(c.id)}
                    dragIndex={dragIndex}
                    overIndex={overIndex}
                    dragRow={dragRow}
                    setDragIndex={setDragIndex}
                    setOverIndex={setOverIndex}
                    setDragRow={setDragRow}
                    reorder={reorder}
                    updateClient={updateClient}
                    removeClient={removeClient}
                    toggleSelected={toggleSelected} />
                ))}
              </div>
            </div>
          </div>
          <div className="pt-3 border-t border-neutral-900 space-y-2">
            <div className="flex items-end gap-2 flex-wrap">
              <div className="flex-1 min-w-[200px]">
                <label className="block text-[10px] uppercase tracking-wide text-neutral-500 mb-1">New client name</label>
                <Input value={draftName} onChange={e => setDraftName(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && (e.preventDefault(), addClient(insertAt === '' ? null : Number(insertAt) - 1))}
                  placeholder="e.g. ABC" />
              </div>
              <div className="w-32">
                <label className="block text-[10px] uppercase tracking-wide text-neutral-500 mb-1">Insert at row #</label>
                <Input type="number" min="1" max={clients.length + 1} value={insertAt}
                  onChange={e => setInsertAt(e.target.value)}
                  placeholder="(end)" />
              </div>
              <Btn variant="primary" size="md" onClick={() => addClient(insertAt === '' ? null : Number(insertAt) - 1)} disabled={!draftName.trim()}>
                <IconPlus /> Insert
              </Btn>
            </div>
            <p className="text-[10px] text-neutral-600 leading-relaxed">
              Leave row # empty to append at the bottom. Use a row # to insert between existing clients (e.g. <span className="text-neutral-400">4</span> inserts between current rows 3 and 4).
            </p>
          </div>
        </div>
      );
    }

    function BmrRuleEditor({ rule, onChange, onDelete, kind = 'target', clients = [], selected, onSelect }) {
      const ops = kind === 'target'
        ? [{v:'gte',l:'greater than or equal to'},{v:'gt',l:'greater than'},{v:'lte',l:'less than or equal to'},{v:'lt',l:'less than'},{v:'eq',l:'equal to'},{v:'between',l:'between'}]
        : [{v:'between',l:'between'},{v:'gte',l:'greater than or equal to'},{v:'gt',l:'greater than'},{v:'lte',l:'less than or equal to'},{v:'lt',l:'less than'},{v:'eq',l:'equal to'}];
      const targetClientMissing = kind === 'target' && rule.clientId && !clients.some(c => c.id === rule.clientId);
      const set = (patch) => onChange({ ...rule, ...patch });
      return (
        <Card className={`p-3 ${selected ? 'border-blue-500/40 bg-blue-500/10' : ''}`}>
          <div className="flex items-center gap-2 mb-2">
            <input type="checkbox" checked={!!selected} onChange={onSelect}
              title="Select rule"
              className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500 focus:ring-blue-500/30 cursor-pointer" />
            <input type="checkbox" checked={rule.enabled !== false} onChange={e => set({ enabled: e.target.checked })}
              title="Enable rule"
              className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500" />
            <Input value={rule.name || ''} onChange={e => set({ name: e.target.value })} placeholder="Rule name" />
            <button onClick={onDelete} className="p-1.5 text-neutral-500 hover:text-red-400" title="Delete rule"><IconX /></button>
          </div>
          {kind === 'target' && (
            <div className="mb-2">
              <label className="block text-[10px] uppercase tracking-wide text-neutral-500 mb-1">Client</label>
              <Select value={rule.clientId || ''} onChange={e => set({ clientId: e.target.value })}>
                <option value="">Choose client</option>
                {targetClientMissing && <option value={rule.clientId}>Missing client</option>}
                {clients.map(c => <option key={c.id} value={c.id}>{c.name || 'Unnamed client'}{c.hidden ? ' (hidden)' : ''}</option>)}
              </Select>
            </div>
          )}
          <div className="grid grid-cols-[220px_1fr_1fr] gap-2 items-end mb-2">
            <Select value={rule.kind} onChange={e => set({ kind: e.target.value })}>
              {ops.map(o => <option key={o.v} value={o.v}>{o.l}</option>)}
            </Select>
            {rule.kind === 'between' ? (
              <>
                <div>
                  <label className="block text-[10px] uppercase tracking-wide text-neutral-500 mb-1">Min</label>
                  <Input type="number" step="0.001" value={rule.min ?? ''} onChange={e => set({ min: Number(e.target.value) })} />
                </div>
                <div>
                  <label className="block text-[10px] uppercase tracking-wide text-neutral-500 mb-1">Max</label>
                  <Input type="number" step="0.001" value={rule.max ?? ''} onChange={e => set({ max: Number(e.target.value) })} />
                </div>
              </>
            ) : (
              <>
                <div className="col-span-2">
                  <label className="block text-[10px] uppercase tracking-wide text-neutral-500 mb-1">Value</label>
                  <Input type="number" step="0.001" value={rule.value ?? ''} onChange={e => set({ value: Number(e.target.value) })} />
                </div>
              </>
            )}
          </div>
          <div className="flex flex-wrap items-center gap-3 text-[11px] text-neutral-400">
            <span className="inline-flex items-center gap-1">
              <span className="text-[10px] uppercase tracking-wide text-neutral-500">Fill</span>
              <BmrColorSwatch value={rule.color} onChange={v => set({ color: v })} />
            </span>
            <span className="inline-flex items-center gap-1">
              <span className="text-[10px] uppercase tracking-wide text-neutral-500">Font</span>
              <BmrColorSwatch value={rule.fontColor} onChange={v => set({ fontColor: v })} />
            </span>
            <label className={`inline-flex items-center gap-1.5 cursor-pointer ${rule.bold ? 'text-neutral-100 font-bold' : ''}`}>
              <input type="checkbox" checked={!!rule.bold} onChange={e => set({ bold: e.target.checked })} className="w-3.5 h-3.5" />
              <span className="font-bold">B</span>
            </label>
            <label className={`inline-flex items-center gap-1.5 cursor-pointer ${rule.italic ? 'text-neutral-100 italic' : ''}`}>
              <input type="checkbox" checked={!!rule.italic} onChange={e => set({ italic: e.target.checked })} className="w-3.5 h-3.5" />
              <span className="italic">I</span>
            </label>
            <label className={`inline-flex items-center gap-1.5 cursor-pointer ${rule.underline ? 'text-neutral-100 underline' : ''}`}>
              <input type="checkbox" checked={!!rule.underline} onChange={e => set({ underline: e.target.checked })} className="w-3.5 h-3.5" />
              <span className="underline">U</span>
            </label>
          </div>
        </Card>
      );
    }

    function BmrRulesPanel({
      bmr,
      onChange,
      targetTitle = 'Target column rules - Amnt Receivable',
      targetHint = 'Each rule targets one client across all 25 Amnt Receivable columns (L, X, AJ, ..., KN)',
      usageHint = 'Applies to all 25 "30mins usage" columns (M, Y, AK, ..., KO)',
    }) {
      const [ruleSearch, setRuleSearch] = useState('');
      const [selectedRuleKeys, setSelectedRuleKeys] = useState([]);
      const [bulkRuleKind, setBulkRuleKind] = useState('gte');
      const [bulkRuleValue, setBulkRuleValue] = useState('5');
      const [bulkRuleMin, setBulkRuleMin] = useState('-49');
      const [bulkRuleMax, setBulkRuleMax] = useState('-0.001');
      const [bulkFillColor, setBulkFillColor] = useState('#EF4444');
      const [bulkFontColor, setBulkFontColor] = useState('#FFFFFF');
      const [generalStyleScope, setGeneralStyleScope] = useState('target');
      const [generalFillColor, setGeneralFillColor] = useState('#EF4444');
      const [generalFontColor, setGeneralFontColor] = useState('#FFFFFF');
      const [generalBold, setGeneralBold] = useState(true);
      const [generalItalic, setGeneralItalic] = useState(false);
      const [generalUnderline, setGeneralUnderline] = useState(false);
      const targetClients = bmr.clients || [];
      const targetClientsById = useMemo(() => new Map(targetClients.map(client => [client.id, client])), [targetClients]);
      const defaultTargetClient = targetClients.find(c => !c.hidden) || targetClients[0];
      const addTarget = () => onChange({ ...bmr, targetRules: [...(bmr.targetRules || []), { id: bmrId('tr'), name: 'New client rule', clientId: defaultTargetClient?.id || '', kind: 'gte', value: 5, color: '#EF4444', fontColor: '#FFFFFF', bold: true, italic: false, underline: false, enabled: true }] });
      const addUsage  = () => onChange({ ...bmr, usageRules:  [...(bmr.usageRules  || []), { id: bmrId('ur'), name: 'New rule', kind: 'between', min: -49, max: -0.001, color: '#F59E0B', fontColor: '#1F2937', bold: true, italic: false, underline: false, enabled: true }] });
      const updT = (id, next) => onChange({ ...bmr, targetRules: bmr.targetRules.map(r => r.id === id ? next : r) });
      const updU = (id, next) => onChange({ ...bmr, usageRules:  bmr.usageRules.map(r  => r.id === id ? next : r) });
      const delT = (id) => onChange({ ...bmr, targetRules: bmr.targetRules.filter(r => r.id !== id) });
      const delU = (id) => onChange({ ...bmr, usageRules:  bmr.usageRules.filter(r  => r.id !== id) });
      const ruleSearchTerm = ruleSearch.trim().toLowerCase();
      const ruleMatches = (rule, clientName = '') => !ruleSearchTerm || [
        rule.name,
        clientName,
        rule.kind,
        rule.value,
        rule.min,
        rule.max,
      ].some(value => String(value ?? '').toLowerCase().includes(ruleSearchTerm));
      const targetRules = bmr.targetRules || [];
      const usageRules = bmr.usageRules || [];
      useEffect(() => {
        const ruleKeys = new Set([
          ...targetRules.map(rule => `target:${rule.id}`),
          ...usageRules.map(rule => `usage:${rule.id}`),
        ]);
        setSelectedRuleKeys((keys) => {
          const kept = keys.filter(key => ruleKeys.has(key));
          return kept.length === keys.length ? keys : kept;
        });
      }, [targetRules, usageRules]);
      const filteredTargetRules = targetRules.filter(rule => ruleMatches(rule, targetClientsById.get(rule.clientId)?.name || ''));
      const filteredUsageRules = usageRules.filter(rule => ruleMatches(rule));
      const filteredTargetRuleIds = new Set(filteredTargetRules.map(rule => rule.id));
      const filteredUsageRuleIds = new Set(filteredUsageRules.map(rule => rule.id));
      const selectedRuleKeySet = new Set(selectedRuleKeys);
      const shownRuleKeys = [
        ...filteredTargetRules.map(rule => `target:${rule.id}`),
        ...filteredUsageRules.map(rule => `usage:${rule.id}`),
      ];
      const allShownRulesSelected = shownRuleKeys.length > 0 && shownRuleKeys.every(key => selectedRuleKeySet.has(key));
      const toggleSelectedRule = (key) => setSelectedRuleKeys((keys) =>
        keys.includes(key) ? keys.filter(item => item !== key) : [...keys, key]
      );
      const toggleShownRules = () => setSelectedRuleKeys((keys) => {
        const next = new Set(keys);
        if (allShownRulesSelected) shownRuleKeys.forEach(key => next.delete(key));
        else shownRuleKeys.forEach(key => next.add(key));
        return Array.from(next);
      });
      const updateSelectedRules = (patch) => {
        if (!selectedRuleKeys.length) return;
        onChange({
          ...bmr,
          targetRules: targetRules.map(rule => selectedRuleKeySet.has(`target:${rule.id}`) ? { ...rule, ...patch } : rule),
          usageRules: usageRules.map(rule => selectedRuleKeySet.has(`usage:${rule.id}`) ? { ...rule, ...patch } : rule),
        });
      };
      const generalStyleCount =
        (generalStyleScope === 'target' || generalStyleScope === 'all' ? filteredTargetRules.length : 0)
        + (generalStyleScope === 'usage' || generalStyleScope === 'all' ? filteredUsageRules.length : 0);
      const updateGeneralStyleRules = (patch) => {
        if (!generalStyleCount) return;
        onChange({
          ...bmr,
          targetRules: targetRules.map(rule =>
            (generalStyleScope === 'target' || generalStyleScope === 'all') && filteredTargetRuleIds.has(rule.id)
              ? { ...rule, ...patch }
              : rule
          ),
          usageRules: usageRules.map(rule =>
            (generalStyleScope === 'usage' || generalStyleScope === 'all') && filteredUsageRuleIds.has(rule.id)
              ? { ...rule, ...patch }
              : rule
          ),
        });
      };
      const applyGeneralStyle = () => updateGeneralStyleRules({
        color: generalFillColor,
        fontColor: generalFontColor,
        bold: generalBold,
        italic: generalItalic,
        underline: generalUnderline,
      });
      const applyBulkCondition = () => {
        const patch = { kind: bulkRuleKind };
        if (bulkRuleKind === 'between') {
          patch.min = Number(bulkRuleMin);
          patch.max = Number(bulkRuleMax);
        } else {
          patch.value = Number(bulkRuleValue);
        }
        updateSelectedRules(patch);
      };

      return (
        <div className="space-y-10">
          <div className="flex flex-wrap items-end justify-between gap-3 rounded-md border border-neutral-900 bg-neutral-950/60 p-3">
            <div className="w-full max-w-sm">
              <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Search rules</label>
              <ClearableInput value={ruleSearch} onChange={e => setRuleSearch(e.target.value)}
                onClear={() => setRuleSearch('')} clearTitle="Clear rules search"
                placeholder="Rule name, client, or value" />
            </div>
            <div className="flex flex-wrap items-center gap-3 pb-2">
              <label className="inline-flex cursor-pointer items-center gap-2 text-[11px] text-neutral-400">
                <input type="checkbox" checked={allShownRulesSelected} onChange={toggleShownRules}
                  disabled={!shownRuleKeys.length}
                  className="h-3.5 w-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500 focus:ring-blue-500/30 disabled:cursor-not-allowed disabled:opacity-40" />
                Select shown rules
              </label>
              <span className="text-[11px] text-neutral-500">{filteredTargetRules.length}/{targetRules.length} client rules / {filteredUsageRules.length}/{usageRules.length} usage rules</span>
            </div>
          </div>
          <div className="rounded-md border border-neutral-900 bg-neutral-950/60 p-3 space-y-3">
            <SectionLabel hint={`${generalStyleCount} shown rule${generalStyleCount === 1 ? '' : 's'}`}>General style edit</SectionLabel>
            <div className="flex flex-wrap items-end gap-3">
              <div className="w-full max-w-[240px]">
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Apply to</label>
                <Select value={generalStyleScope} onChange={e => setGeneralStyleScope(e.target.value)}>
                  <option value="target">Target/client rules shown</option>
                  <option value="usage">30-min usage rules shown</option>
                  <option value="all">All shown rules</option>
                </Select>
              </div>
              <span className="inline-flex h-9 items-center gap-1 rounded-md border border-neutral-800 bg-neutral-950 px-2">
                <span className="text-[10px] uppercase tracking-wide text-neutral-500">Fill</span>
                <BmrColorSwatch value={generalFillColor} onChange={setGeneralFillColor} allowClear={false} />
              </span>
              <span className="inline-flex h-9 items-center gap-1 rounded-md border border-neutral-800 bg-neutral-950 px-2">
                <span className="text-[10px] uppercase tracking-wide text-neutral-500">Font</span>
                <BmrColorSwatch value={generalFontColor} onChange={setGeneralFontColor} allowClear={false} />
              </span>
              <label className={`inline-flex h-9 cursor-pointer items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2 text-[12px] ${generalBold ? 'text-neutral-100' : 'text-neutral-500'}`}>
                <input type="checkbox" checked={generalBold} onChange={e => setGeneralBold(e.target.checked)} className="h-3.5 w-3.5" />
                <span className="font-bold">B</span>
              </label>
              <label className={`inline-flex h-9 cursor-pointer items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2 text-[12px] ${generalItalic ? 'text-neutral-100' : 'text-neutral-500'}`}>
                <input type="checkbox" checked={generalItalic} onChange={e => setGeneralItalic(e.target.checked)} className="h-3.5 w-3.5" />
                <span className="italic">I</span>
              </label>
              <label className={`inline-flex h-9 cursor-pointer items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2 text-[12px] ${generalUnderline ? 'text-neutral-100' : 'text-neutral-500'}`}>
                <input type="checkbox" checked={generalUnderline} onChange={e => setGeneralUnderline(e.target.checked)} className="h-3.5 w-3.5" />
                <span className="underline">U</span>
              </label>
              <Btn variant="accent" size="md" onClick={applyGeneralStyle} disabled={!generalStyleCount}>Apply style</Btn>
              <Btn variant="ghost" size="md" onClick={() => updateGeneralStyleRules({ color: '' })} disabled={!generalStyleCount}>Clear fill</Btn>
              <Btn variant="ghost" size="md" onClick={() => updateGeneralStyleRules({ fontColor: '' })} disabled={!generalStyleCount}>Clear font</Btn>
            </div>
            <p className="text-[10px] text-neutral-600 leading-relaxed">
              Uses the current search filter. Clear search first if you want to update every rule in the selected group.
            </p>
          </div>
          {selectedRuleKeys.length > 0 && (
            <div className="rounded-md border border-blue-500/40 bg-blue-500/10 p-3 space-y-3">
              <SectionLabel hint={`${selectedRuleKeys.length} selected`}>Bulk edit rules</SectionLabel>
              <div className="flex flex-wrap items-end gap-2">
                <div className="w-full max-w-[220px]">
                  <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Condition</label>
                  <Select value={bulkRuleKind} onChange={e => setBulkRuleKind(e.target.value)}>
                    <option value="gte">greater than or equal to</option>
                    <option value="gt">greater than</option>
                    <option value="lte">less than or equal to</option>
                    <option value="lt">less than</option>
                    <option value="eq">equal to</option>
                    <option value="between">between</option>
                  </Select>
                </div>
                {bulkRuleKind === 'between' ? (
                  <>
                    <div className="w-28">
                      <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Min</label>
                      <Input type="number" step="0.001" value={bulkRuleMin} onChange={e => setBulkRuleMin(e.target.value)} />
                    </div>
                    <div className="w-28">
                      <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Max</label>
                      <Input type="number" step="0.001" value={bulkRuleMax} onChange={e => setBulkRuleMax(e.target.value)} />
                    </div>
                  </>
                ) : (
                  <div className="w-32">
                    <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Value</label>
                    <Input type="number" step="0.001" value={bulkRuleValue} onChange={e => setBulkRuleValue(e.target.value)} />
                  </div>
                )}
                <Btn variant="ghost" size="md" onClick={applyBulkCondition}>Apply condition</Btn>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <span className="inline-flex h-9 items-center gap-1 rounded-md border border-neutral-800 bg-neutral-950 px-2">
                  <span className="text-[10px] uppercase tracking-wide text-neutral-500">Fill</span>
                  <BmrColorSwatch value={bulkFillColor} onChange={setBulkFillColor} allowClear={false} />
                </span>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedRules({ color: bulkFillColor })}>Apply fill</Btn>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedRules({ color: '' })}>Clear fill</Btn>
                <span className="inline-flex h-9 items-center gap-1 rounded-md border border-neutral-800 bg-neutral-950 px-2">
                  <span className="text-[10px] uppercase tracking-wide text-neutral-500">Font</span>
                  <BmrColorSwatch value={bulkFontColor} onChange={setBulkFontColor} allowClear={false} />
                </span>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedRules({ fontColor: bulkFontColor })}>Apply font</Btn>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedRules({ fontColor: '' })}>Clear font</Btn>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <Btn variant="ghost" size="md" onClick={() => updateSelectedRules({ enabled: true })}>Enable</Btn>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedRules({ enabled: false })}>Disable</Btn>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedRules({ bold: true })}><span className="font-bold">B</span></Btn>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedRules({ italic: true })}><span className="italic">I</span></Btn>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedRules({ underline: true })}><span className="underline">U</span></Btn>
                <Btn variant="ghost" size="md" onClick={() => updateSelectedRules({ bold: false, italic: false, underline: false })}>Clear text style</Btn>
                <Btn variant="ghost" size="md" onClick={() => setSelectedRuleKeys([])}>Clear selection</Btn>
              </div>
            </div>
          )}
          <div>
            <SectionLabel hint={targetHint}>{targetTitle}</SectionLabel>
            <div className="space-y-2">
              {filteredTargetRules.map(r => (
                <BmrRuleEditor key={r.id} rule={r} kind="target" clients={targetClients}
                  selected={selectedRuleKeySet.has(`target:${r.id}`)}
                  onSelect={() => toggleSelectedRule(`target:${r.id}`)}
                  onChange={(next) => updT(r.id, next)} onDelete={() => delT(r.id)} />
              ))}
              <Btn variant="ghost" size="sm" onClick={addTarget}><IconPlus /> Add target rule</Btn>
            </div>
          </div>
          <div aria-hidden="true" className="border-t-2 border-neutral-700"></div>
          <div>
            <SectionLabel hint={usageHint}>30-min usage column rules</SectionLabel>
            <div className="space-y-2">
              {filteredUsageRules.map(r => (
                <BmrRuleEditor key={r.id} rule={r} kind="usage"
                  selected={selectedRuleKeySet.has(`usage:${r.id}`)}
                  onSelect={() => toggleSelectedRule(`usage:${r.id}`)}
                  onChange={(next) => updU(r.id, next)} onDelete={() => delU(r.id)} />
              ))}
              <Btn variant="ghost" size="sm" onClick={addUsage}><IconPlus /> Add usage rule</Btn>
            </div>
          </div>
        </div>
      );
    }

    function BmrPreview({ bmr }) {
      const visibleClients = (bmr.clients || []).filter(c => !c.hidden).slice(0, 8);
      const accountManagersById = bmrAccountManagersById(bmr);
      const slots = bmrSlotsForShift('day').slice(0, 3); // first 3 timeslots
      return (
        <Card className="p-4 overflow-x-auto">
          <p className="text-[10px] text-neutral-600 mb-3 font-mono">Day sheet preview · first {visibleClients.length} client(s) × {slots.length} timeslot(s)</p>
          <table className="min-w-full text-[11px] border-collapse">
            <thead>
              <tr className="bg-neutral-900 text-neutral-300">
                <th className="px-2 py-1 border border-neutral-800 text-left">CarrierName</th>
                <th className="px-2 py-1 border border-neutral-800 text-left">Allow Bal</th>
                {slots.map((s, i) => (
                  <th key={i} colSpan={3} className="px-2 py-1 border border-neutral-800 text-center">{s}</th>
                ))}
              </tr>
              <tr className="bg-neutral-950 text-neutral-500 text-[10px]">
                <th className="px-2 py-1 border border-neutral-800"></th>
                <th className="px-2 py-1 border border-neutral-800"></th>
                {slots.map((_, i) => (
                  <React.Fragment key={i}>
                    <th className="px-2 py-1 border border-neutral-800">Amnt Rcv</th>
                    <th className="px-2 py-1 border border-neutral-800">30min</th>
                    <th className="px-2 py-1 border border-neutral-800">Total</th>
                  </React.Fragment>
                ))}
              </tr>
            </thead>
            <tbody>
              {visibleClients.map((c) => (
                <tr key={c.id}>
                  <td className="px-2 py-1 border border-neutral-800" style={bmrClientColor(bmr, c, accountManagersById) ? { background: bmrClientColor(bmr, c, accountManagersById), color: '#000' } : {}}>{c.name}</td>
                  <td className="px-2 py-1 border border-neutral-800 text-neutral-500">{c.allowBalanceLabel || '—'}</td>
                  {slots.map((_, i) => (
                    <React.Fragment key={i}>
                      <td className="px-2 py-1 border border-neutral-800 text-neutral-700">·</td>
                      <td className="px-2 py-1 border border-neutral-800 text-neutral-700">·</td>
                      <td className="px-2 py-1 border border-neutral-800 text-neutral-700">·</td>
                    </React.Fragment>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
          <p className="text-[10px] text-neutral-600 mt-3 leading-relaxed">
            Generated file: 2 sheets (<span className="text-neutral-400">7AM-7PM</span> and <span className="text-neutral-400">7PM-7AM</span>), 25 timeslots × 12 cols each, <span className="text-neutral-400">{(bmr.clients || []).filter(c => !c.hidden).length}</span> visible client rows. 30-min usage and Total usage are auto-formulas referencing the previous slot.
          </p>
        </Card>
      );
    }

    function BmrSidebar({ state, setState, onGenerate, busy, sync, onRetrySync, gsheets }) {
      const theme = state.theme || 'dark';
      const toggleTheme = () => setState(s => ({ ...s, theme: (s.theme || 'dark') === 'dark' ? 'light' : 'dark' }));
      const bmr = state.bmr || DEFAULT_BMR_STATE;
      const accountManagersById = bmrAccountManagersById(bmr);
      const setBmr = (next) => setState(s => ({ ...s, bmr: next }));

      return (
        <aside className="w-[320px] shrink-0 border-r border-neutral-900 bg-[#17171a] h-screen sticky top-0 flex flex-col">
          <div className="p-5 border-b border-neutral-900">
            <div className="flex items-center justify-between gap-2">
              <SyncBadge sync={sync} onRetry={onRetrySync} />
              <button onClick={toggleTheme}
                title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}
                className="inline-flex items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2 py-1 text-[10px] uppercase tracking-wide text-neutral-400 hover:text-neutral-100 hover:border-neutral-700 transition-colors">
                {theme === 'dark' ? <IconSun /> : <IconMoon />}
                {theme === 'dark' ? 'Light' : 'Dark'}
              </button>
            </div>
            <h1 className="text-[17px] font-bold tracking-tight leading-tight mt-2">BMR VOIP Generator</h1>
            <p className="text-xs text-neutral-500 mt-1">Daily · Balance Day &amp; Night</p>
            <AccountChip sync={sync} />
          </div>

          <div className="p-5 border-b border-neutral-900">
            <SectionLabel>Today</SectionLabel>
            <div className="text-2xl font-bold tracking-tight">{bmrTodayString()}</div>
            <p className="text-[11px] text-neutral-500 mt-1 font-mono">Daily generation</p>
          </div>

          <div className="p-5 border-b border-neutral-900 space-y-2">
            <SectionLabel>Shifts</SectionLabel>
            <label className="flex items-center gap-2 text-xs text-neutral-300 cursor-pointer">
              <input type="checkbox" checked={bmr.includeDay !== false}
                onChange={e => setBmr({ ...bmr, includeDay: e.target.checked })}
                className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500" />
              <span>7AM–7PM (Day)</span>
            </label>
            <label className="flex items-center gap-2 text-xs text-neutral-300 cursor-pointer">
              <input type="checkbox" checked={bmr.includeNight !== false}
                onChange={e => setBmr({ ...bmr, includeNight: e.target.checked })}
                className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500" />
              <span>7PM–7AM (Night)</span>
            </label>
          </div>

          <div className="p-5 border-b border-neutral-900">
            <SectionLabel>Hidden columns</SectionLabel>
            <div className="flex flex-wrap gap-1 mb-2">
              {(bmr.hiddenCols || []).length === 0 && <span className="text-[11px] text-neutral-600">None</span>}
              {(bmr.hiddenCols || []).map(c => (
                <span key={c} className="inline-flex items-center gap-1 rounded-full border border-neutral-800 bg-neutral-950 pl-2 pr-1 py-0.5 text-[10px] text-neutral-300 font-mono">
                  {c}
                  <button onClick={() => setBmr({ ...bmr, hiddenCols: bmr.hiddenCols.filter(x => x !== c) })}
                    className="text-neutral-500 hover:text-red-400 p-0.5"><IconX /></button>
                </span>
              ))}
            </div>
            <BmrHiddenColInput bmr={bmr} setBmr={setBmr} />
          </div>

          <div className="p-5 flex-1 overflow-y-auto min-h-0">
            <SectionLabel hint={`${(bmr.clients || []).filter(c => !c.hidden).length}/${(bmr.clients || []).length} visible`}>Clients quick view</SectionLabel>
            <div className="space-y-0.5 max-h-[40vh] overflow-y-auto">
              {(bmr.clients || []).map((c, i) => (
                <div key={c.id} className={`flex items-center gap-2 rounded px-1.5 py-1 ${c.hidden ? 'opacity-40' : ''}`}>
                  <span className="w-3 h-3 rounded border border-neutral-800 shrink-0" style={bmrClientColor(bmr, c, accountManagersById) ? { background: bmrClientColor(bmr, c, accountManagersById) } : {}}></span>
                  <span className="text-[11px] text-neutral-300 truncate flex-1">{c.name}</span>
                </div>
              ))}
            </div>
          </div>

          <div className="p-5 border-t border-neutral-900 space-y-3">
            <GoogleSheetSync gsheets={gsheets} moduleId="bmr" sheetId={state.googleSheets?.sheetIds?.bmr} />
            <Btn variant="primary" size="lg" onClick={onGenerate} disabled={busy} className="w-full">
              {busy ? <><span className="loader"></span> Generating…</> : <><IconDownload /> Generate BMR VOIP</>}
            </Btn>
            <p className="text-[10px] text-neutral-600 text-center font-mono">.xlsx · Google Sheets compatible</p>
          </div>
        </aside>
      );
    }

    function BmrHiddenColInput({ bmr, setBmr }) {
      const [v, setV] = useState('');
      const add = () => {
        const raw = v.trim().toUpperCase();
        if (!raw) return;
        const list = raw.split(/[\s,]+/).filter(Boolean);
        const next = Array.from(new Set([...(bmr.hiddenCols || []), ...list]));
        setBmr({ ...bmr, hiddenCols: next });
        setV('');
      };
      return (
        <div className="flex gap-1">
          <Input value={v} onChange={e => setV(e.target.value)} onKeyDown={e => e.key === 'Enter' && (e.preventDefault(), add())} placeholder="e.g. L, X, AJ" />
          <Btn variant="ghost" size="sm" onClick={add}>Hide</Btn>
        </div>
      );
    }

    function bmrSmsCombinedRuleClients(sms) {
      return [
        ...((sms.retailClients || []).map(client => ({ ...client, name: `RES - ${client.name}` }))),
        ...((sms.wholesaleClients || []).map(client => ({ ...client, name: `WHS - ${client.name}` }))),
      ];
    }

    function BmrSmsClientsEditor({ sms, onChange }) {
      const [marketKey, setMarketKey] = useState('retail');
      const market = BMR_SMS_MARKETS[marketKey];
      const otherMarket = marketKey === 'retail' ? BMR_SMS_MARKETS.wholesale : BMR_SMS_MARKETS.retail;
      const marketClients = sms[market.clientsKey] || [];
      const visibleRetail = (sms.retailClients || []).filter(c => !c.hidden).length;
      const visibleWholesale = (sms.wholesaleClients || []).filter(c => !c.hidden).length;

      const marketBmr = useMemo(() => ({
        ...sms,
        clients: marketClients,
      }), [sms, marketClients]);

      const setMarketBmr = useCallback((next) => {
        onChange((prevSms) => {
          const base = { ...DEFAULT_BMR_SMS_STATE, ...(prevSms || {}) };
          const currentView = { ...base, clients: base[market.clientsKey] || [] };
          const resolved = typeof next === 'function' ? next(currentView) : next;
          const accountManagers = resolved.accountManagers || [];
          const managerIds = new Set(accountManagers.map(manager => manager.id));
          const cleanManagerLinks = (clients) => (clients || []).map(client =>
            client.accountManagerId && !managerIds.has(client.accountManagerId)
              ? { ...client, accountManagerId: '' }
              : client
          );
          return {
            ...base,
            accountManagers,
            colorBlockCarrierNames: resolved.colorBlockCarrierNames,
            bordersOnData: resolved.bordersOnData,
            borderStyle: resolved.borderStyle,
            timeSeparatorBorderStyle: resolved.timeSeparatorBorderStyle,
            borderColor: resolved.borderColor,
            [market.clientsKey]: cleanManagerLinks(resolved.clients || []),
            [otherMarket.clientsKey]: cleanManagerLinks(base[otherMarket.clientsKey] || []),
          };
        });
      }, [market.clientsKey, onChange, otherMarket.clientsKey]);

      return (
        <div className="space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-3 rounded-md border border-neutral-900 bg-neutral-950/60 p-3">
            <div>
              <SectionLabel hint="Separate management for RES and WHS">SMS clients</SectionLabel>
              <p className="text-xs text-neutral-500">Retail and Wholesale keep separate client orders, colors, managers, and hidden rows.</p>
            </div>
            <div className="inline-flex rounded-md border border-neutral-800 bg-neutral-950 p-0.5">
              {[
                { id: 'retail', label: 'Retail / RES', count: visibleRetail },
                { id: 'wholesale', label: 'Wholesale / WHS', count: visibleWholesale },
              ].map(item => (
                <button key={item.id} onClick={() => setMarketKey(item.id)}
                  className={`rounded px-3 py-1.5 text-xs font-medium transition-colors ${marketKey === item.id ? 'bg-blue-500/20 text-blue-200 border border-blue-500/40' : 'text-neutral-500 hover:text-neutral-200'}`}>
                  {item.label} · {item.count}
                </button>
              ))}
            </div>
          </div>
          <BmrClientsEditor
            key={marketKey}
            bmr={marketBmr}
            onChange={setMarketBmr}
            showBlockColumnHider={false}
            showCarrierNameSettings={true}
          />
        </div>
      );
    }

    function BmrSmsOverdraftRuleEditor({ rule, onChange, onDelete }) {
      const set = (patch) => onChange({ ...rule, ...patch });
      return (
        <Card className="p-3">
          <div className="flex items-center gap-2 mb-2">
            <input type="checkbox" checked={rule.enabled !== false} onChange={e => set({ enabled: e.target.checked })}
              title="Enable rule"
              className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500" />
            <Input value={rule.name || ''} onChange={e => set({ name: e.target.value })} placeholder="Rule name" />
            <button onClick={onDelete} className="p-1.5 text-neutral-500 hover:text-red-400" title="Delete rule"><IconX /></button>
          </div>
          <p className="mb-3 text-[11px] leading-relaxed text-neutral-500">
            Colors the BALANCE cell when the balance is debit and the debit amount reaches the Over Draft / OD value in the same time block.
          </p>
          <div className="flex flex-wrap items-center gap-3 text-[11px] text-neutral-400">
            <span className="inline-flex items-center gap-1">
              <span className="text-[10px] uppercase tracking-wide text-neutral-500">Fill</span>
              <BmrColorSwatch value={rule.color} onChange={v => set({ color: v })} />
            </span>
            <span className="inline-flex items-center gap-1">
              <span className="text-[10px] uppercase tracking-wide text-neutral-500">Font</span>
              <BmrColorSwatch value={rule.fontColor} onChange={v => set({ fontColor: v })} />
            </span>
            <label className={`inline-flex items-center gap-1.5 cursor-pointer ${rule.bold ? 'text-neutral-100 font-bold' : ''}`}>
              <input type="checkbox" checked={!!rule.bold} onChange={e => set({ bold: e.target.checked })} className="w-3.5 h-3.5" />
              <span className="font-bold">B</span>
            </label>
            <label className={`inline-flex items-center gap-1.5 cursor-pointer ${rule.italic ? 'text-neutral-100 italic' : ''}`}>
              <input type="checkbox" checked={!!rule.italic} onChange={e => set({ italic: e.target.checked })} className="w-3.5 h-3.5" />
              <span className="italic">I</span>
            </label>
            <label className={`inline-flex items-center gap-1.5 cursor-pointer ${rule.underline ? 'text-neutral-100 underline' : ''}`}>
              <input type="checkbox" checked={!!rule.underline} onChange={e => set({ underline: e.target.checked })} className="w-3.5 h-3.5" />
              <span className="underline">U</span>
            </label>
          </div>
        </Card>
      );
    }

    function BmrSmsRulesPanel({ sms, onChange }) {
      const ruleClients = useMemo(() => bmrSmsCombinedRuleClients(sms), [sms]);
      const rulesBmr = useMemo(() => ({
        ...sms,
        clients: ruleClients,
      }), [sms, ruleClients]);
      const overdraftRules = sms.overdraftRules || DEFAULT_BMR_SMS_STATE.overdraftRules;
      const addOverdraftRule = () => onChange(prev => ({
        ...(prev || DEFAULT_BMR_SMS_STATE),
        overdraftRules: [
          ...((prev || DEFAULT_BMR_SMS_STATE).overdraftRules || []),
          { id: bmrId('smsod'), name: 'Balance reached overdraft', color: '#EF4444', fontColor: '#FFFFFF', bold: true, italic: false, underline: false, enabled: true },
        ],
      }));
      const updateOverdraftRule = (id, next) => onChange(prev => ({
        ...(prev || DEFAULT_BMR_SMS_STATE),
        overdraftRules: ((prev || DEFAULT_BMR_SMS_STATE).overdraftRules || []).map(rule => rule.id === id ? next : rule),
      }));
      const deleteOverdraftRule = (id) => onChange(prev => ({
        ...(prev || DEFAULT_BMR_SMS_STATE),
        overdraftRules: ((prev || DEFAULT_BMR_SMS_STATE).overdraftRules || []).filter(rule => rule.id !== id),
      }));
      const setRulesBmr = useCallback((next) => {
        onChange((prevSms) => {
          const base = { ...DEFAULT_BMR_SMS_STATE, ...(prevSms || {}) };
          const currentView = { ...base, clients: bmrSmsCombinedRuleClients(base) };
          const resolved = typeof next === 'function' ? next(currentView) : next;
          return {
            ...base,
            targetRules: resolved.targetRules || [],
            usageRules: resolved.usageRules || [],
          };
        });
      }, [onChange]);

      return (
        <div className="space-y-10">
          <div>
            <SectionLabel hint="BMR SMS only - compares BALANCE with Over Draft / OD in the same time block">Overdraft reached rules</SectionLabel>
            <div className="space-y-2">
              {overdraftRules.map(rule => (
                <BmrSmsOverdraftRuleEditor key={rule.id} rule={rule}
                  onChange={(next) => updateOverdraftRule(rule.id, next)}
                  onDelete={() => deleteOverdraftRule(rule.id)} />
              ))}
              <Btn variant="ghost" size="sm" onClick={addOverdraftRule}><IconPlus /> Add overdraft rule</Btn>
            </div>
          </div>
          <div>
            <SectionLabel hint="BMR SMS only - applies to both RES and WHS sheets">Balance emphasis</SectionLabel>
            <div className="rounded-md border border-neutral-900 bg-neutral-950/60 p-3">
              <label className="flex cursor-pointer items-start gap-2 text-sm text-neutral-300">
                <input type="checkbox" checked={sms.boldBalanceOnUsage !== false}
                  onChange={e => onChange(prev => ({ ...(prev || DEFAULT_BMR_SMS_STATE), boldBalanceOnUsage: e.target.checked }))}
                  className="mt-0.5 h-3.5 w-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500 focus:ring-blue-500/30 cursor-pointer" />
                <span>Bold the BALANCE cell when its 30mins usage is non-zero
                  <span className="block text-[11px] text-neutral-500">Only balance cells whose same-row 30mins usage has a value turn bold — makes it easy to spot BALANCE apart from the Over Draft / OD column.</span>
                </span>
              </label>
            </div>
          </div>
          <div aria-hidden="true" className="border-t-2 border-neutral-700"></div>
          <BmrRulesPanel
            bmr={rulesBmr}
            onChange={setRulesBmr}
            targetTitle="Balance column rules"
            targetHint="Each rule targets one SMS client across the 25 Balance columns. Retail uses BALANCE; Wholesale uses BALANCE after OD."
            usageHint="Applies to all 25 SMS 30mins usage columns in RES and WHS sheets."
          />
        </div>
      );
    }

    function BmrSmsPreview({ sms }) {
      const slots = bmrSlotsForShift('day').slice(0, 3);
      const accountManagersById = bmrAccountManagersById(sms);
      const renderMarket = (marketKey) => {
        const market = BMR_SMS_MARKETS[marketKey];
        const clients = (sms[market.clientsKey] || []).filter(c => !c.hidden).slice(0, 6);
        const detailHeaders = marketKey === 'retail'
          ? ['Type', 'Name', 'Balance', 'Over Draft']
          : ['Type', 'Name', 'OD Limit', 'Balance'];
        return (
          <Card key={marketKey} className="p-4 overflow-x-auto">
            <p className="text-[10px] text-neutral-600 mb-3 font-mono">{market.label} preview · first {clients.length} client(s) × {slots.length} timeslot(s)</p>
            <table className="min-w-full text-[11px] border-collapse">
              <thead>
                <tr className="bg-neutral-900 text-neutral-300">
                  <th className="px-2 py-1 border border-neutral-800 text-left">CLIENT</th>
                  <th className="px-2 py-1 border border-neutral-800 text-left">Allow Bal</th>
                  {slots.map((slot, index) => (
                    <th key={index} colSpan={4} className="px-2 py-1 border border-neutral-800 text-center">{slot}</th>
                  ))}
                </tr>
                <tr className="bg-neutral-950 text-neutral-500 text-[10px]">
                  <th className="px-2 py-1 border border-neutral-800"></th>
                  <th className="px-2 py-1 border border-neutral-800"></th>
                  {slots.map((_, slotIndex) => detailHeaders.map(header => (
                    <th key={`${slotIndex}-${header}`} className="px-2 py-1 border border-neutral-800">{header}</th>
                  )))}
                </tr>
              </thead>
              <tbody>
                {clients.map(client => {
                  const clientColor = bmrClientColor(sms, client, accountManagersById);
                  return (
                    <tr key={client.id}>
                      <td className="px-2 py-1 border border-neutral-800" style={clientColor ? { background: clientColor, color: '#000' } : {}}>{client.name}</td>
                      <td className="px-2 py-1 border border-neutral-800 text-neutral-500">{client.allowBalanceLabel || '-'}</td>
                      {slots.map((_, slotIndex) => detailHeaders.map(header => (
                        <td key={`${slotIndex}-${header}`} className="px-2 py-1 border border-neutral-800 text-neutral-700">·</td>
                      )))}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </Card>
        );
      };

      return (
        <div className="space-y-3">
          {renderMarket('retail')}
          {renderMarket('wholesale')}
          <p className="text-[10px] text-neutral-600 leading-relaxed">
            Generated file: 4 sheets when both shifts are enabled: RES DAY, WHS DAY, RES NIGHT, and WHS NIGHT. Each sheet keeps the sample workbook's hidden in-block column pattern and adds 10 blank rows below the visible client list.
          </p>
        </div>
      );
    }

    function BmrSmsSidebar({ state, setState, onGenerate, busy, sync, onRetrySync, gsheets }) {
      const theme = state.theme || 'dark';
      const toggleTheme = () => setState(s => ({ ...s, theme: (s.theme || 'dark') === 'dark' ? 'light' : 'dark' }));
      const sms = state.bmrSms || DEFAULT_BMR_SMS_STATE;
      const accountManagersById = bmrAccountManagersById(sms);
      const setSms = (next) => setState(s => {
        const prevSms = s.bmrSms || DEFAULT_BMR_SMS_STATE;
        return { ...s, bmrSms: typeof next === 'function' ? next(prevSms) : next };
      });
      const visibleRetail = (sms.retailClients || []).filter(c => !c.hidden).length;
      const visibleWholesale = (sms.wholesaleClients || []).filter(c => !c.hidden).length;
      const quickClients = [
        ...(sms.retailClients || []).slice(0, 8).map(client => ({ ...client, market: 'RES' })),
        ...(sms.wholesaleClients || []).slice(0, 8).map(client => ({ ...client, market: 'WHS' })),
      ];

      return (
        <aside className="w-[320px] shrink-0 border-r border-neutral-900 bg-[#17171a] h-screen sticky top-0 flex flex-col">
          <div className="p-5 border-b border-neutral-900">
            <div className="flex items-center justify-between gap-2">
              <SyncBadge sync={sync} onRetry={onRetrySync} />
              <button onClick={toggleTheme}
                title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}
                className="inline-flex items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2 py-1 text-[10px] uppercase tracking-wide text-neutral-400 hover:text-neutral-100 hover:border-neutral-700 transition-colors">
                {theme === 'dark' ? <IconSun /> : <IconMoon />}
                {theme === 'dark' ? 'Light' : 'Dark'}
              </button>
            </div>
            <h1 className="text-[17px] font-bold tracking-tight leading-tight mt-2">BMR SMS Generator</h1>
            <p className="text-xs text-neutral-500 mt-1">Retail and Wholesale · Day &amp; Night</p>
            <AccountChip sync={sync} />
          </div>

          <div className="p-5 border-b border-neutral-900">
            <SectionLabel>Today</SectionLabel>
            <div className="text-2xl font-bold tracking-tight">{bmrTodayString()}</div>
            <p className="text-[11px] text-neutral-500 mt-1 font-mono">4-sheet SMS generation</p>
          </div>

          <div className="p-5 border-b border-neutral-900 space-y-2">
            <SectionLabel>Shifts</SectionLabel>
            <label className="flex items-center gap-2 text-xs text-neutral-300 cursor-pointer">
              <input type="checkbox" checked={sms.includeDay !== false}
                onChange={e => setSms({ ...sms, includeDay: e.target.checked })}
                className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500" />
              <span>7AM-7PM (Day)</span>
            </label>
            <label className="flex items-center gap-2 text-xs text-neutral-300 cursor-pointer">
              <input type="checkbox" checked={sms.includeNight !== false}
                onChange={e => setSms({ ...sms, includeNight: e.target.checked })}
                className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500" />
              <span>7PM-7AM (Night)</span>
            </label>
          </div>

          <div className="p-5 flex-1 overflow-y-auto min-h-0">
            <SectionLabel hint={`${visibleRetail} RES / ${visibleWholesale} WHS visible`}>Clients quick view</SectionLabel>
            <div className="space-y-0.5 max-h-[44vh] overflow-y-auto">
              {quickClients.map((client, index) => {
                const color = bmrClientColor(sms, client, accountManagersById);
                return (
                  <div key={`${client.market}-${client.id}-${index}`} className={`flex items-center gap-2 rounded px-1.5 py-1 ${client.hidden ? 'opacity-40' : ''}`}>
                    <span className="w-8 text-[10px] text-neutral-600 font-mono">{client.market}</span>
                    <span className="w-3 h-3 rounded border border-neutral-800 shrink-0" style={color ? { background: color } : {}}></span>
                    <span className="text-[11px] text-neutral-300 truncate flex-1">{client.name}</span>
                  </div>
                );
              })}
            </div>
          </div>

          <div className="p-5 border-t border-neutral-900 space-y-3">
            <GoogleSheetSync gsheets={gsheets} moduleId="bmr_sms" sheetId={state.googleSheets?.sheetIds?.bmr_sms} />
            <Btn variant="primary" size="lg" onClick={onGenerate} disabled={busy} className="w-full">
              {busy ? <><span className="loader"></span> Generating...</> : <><IconDownload /> Generate BMR SMS</>}
            </Btn>
            <p className="text-[10px] text-neutral-600 text-center font-mono">.xlsx · Google Sheets compatible</p>
          </div>
        </aside>
      );
    }

    /* ============================================================
       Editor - UI Components
       ============================================================ */
    function EditorColorSwatch({ value, onChange, title = 'Choose color', allowClear = true, fallback = '#FFFFFF' }) {
      return (
        <span className="inline-flex items-center gap-1.5">
          <PresetColorPicker value={value || fallback} onChange={onChange}
            title={title} showHex={false} buttonClassName="h-8 px-1.5 py-1" />
          {allowClear && value && (
            <button onClick={() => onChange('')} className="text-[10px] text-neutral-500 hover:text-red-300 px-1" title={`Clear ${title.toLowerCase()}`}>clear</button>
          )}
        </span>
      );
    }

    function EditorSidebar({ state, setState, onGenerate, busy, sync, onRetrySync }) {
      const theme = state.theme || 'dark';
      const editor = { ...DEFAULT_EDITOR_STATE, ...(state.editor || {}) };
      const grid = editorNormalizeGrid(editor.cells);
      const toggleTheme = () => setState(s => ({ ...s, theme: (s.theme || 'dark') === 'dark' ? 'light' : 'dark' }));
      return (
        <aside className="w-[320px] shrink-0 border-r border-neutral-900 bg-[#17171a] h-screen sticky top-0 flex flex-col">
          <div className="p-5 border-b border-neutral-900">
            <div className="flex items-center justify-between gap-2">
              <SyncBadge sync={sync} onRetry={onRetrySync} />
              <button onClick={toggleTheme}
                title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}
                className="inline-flex items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2 py-1 text-[10px] uppercase tracking-wide text-neutral-400 hover:text-neutral-100 hover:border-neutral-700 transition-colors">
                {theme === 'dark' ? <IconSun /> : <IconMoon />}
                {theme === 'dark' ? 'Light' : 'Dark'}
              </button>
            </div>
            <h1 className="text-[17px] font-bold tracking-tight leading-tight mt-2">Table Editor</h1>
            <p className="text-xs text-neutral-500 mt-1">Pasted table workbook</p>
            <AccountChip sync={sync} />
          </div>
          <div className="p-5 flex-1">
            <SectionLabel>Table</SectionLabel>
            <div className="rounded-md border border-neutral-900 bg-neutral-950 p-3 space-y-2">
              <div className="truncate text-sm font-medium text-neutral-100" title={editor.name || 'Pasted Table'}>{editor.name || 'Pasted Table'}</div>
              <div className="flex flex-wrap gap-1.5">
                <Pill>{grid.length} rows</Pill>
                <Pill>{grid[0].length} columns</Pill>
                <Pill tone={(editor.rules || []).some(rule => rule.enabled !== false) ? 'accent' : 'default'}>{(editor.rules || []).length} rules</Pill>
              </div>
            </div>
          </div>
          <div className="p-5 border-t border-neutral-900 space-y-3">
            <Btn variant="primary" size="lg" onClick={onGenerate} disabled={busy} className="w-full">
              {busy ? <><span className="loader"></span> Generating</> : <><IconDownload /> Export Editor</>}
            </Btn>
            <p className="text-[10px] text-neutral-600 text-center font-mono">.xlsx / Google Sheets compatible</p>
          </div>
        </aside>
      );
    }

    function EditorPastePanel({ editor, onChange }) {
      const [pasteText, setPasteText] = useState('');
      useEffect(() => {
        const grid = editorNormalizeGrid(editor.cells);
        if (grid.length === 1 && grid[0].length === 1 && !grid[0][0]) setPasteText('');
      }, [editor.cells]);
      const loadText = (text) => {
        setPasteText(text);
        onChange({ ...editor, cells: editorParseGridText(text) });
      };
      const paste = (e) => {
        const text = e.clipboardData?.getData('text/plain');
        if (text === undefined || text === null || text === '') return;
        e.preventDefault();
        loadText(text);
      };
      return (
        <Card className="p-4">
          <SectionLabel hint="Clipboard table">Paste</SectionLabel>
          <Textarea rows={4} value={pasteText} onPaste={paste} onChange={e => setPasteText(e.target.value)}
            placeholder={`Paste cells copied from Google Sheets here`} />
          <div className="mt-2 flex flex-wrap items-center gap-2">
            <Btn variant="ghost" size="md" onClick={() => loadText(pasteText)} disabled={!pasteText}>Load pasted table</Btn>
            <span className="text-[11px] text-neutral-500">{editorNormalizeGrid(editor.cells).length} rows / {editorNormalizeGrid(editor.cells)[0].length} columns</span>
          </div>
        </Card>
      );
    }

    function EditorColorField({ label, value, onChange, title, fallback }) {
      return (
        <div>
          <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">{label}</label>
          <EditorColorSwatch value={value} onChange={onChange} title={title} fallback={fallback} />
        </div>
      );
    }

    function EditorFontStyleControls({ label, bold, italic, underline, onChange }) {
      const update = (patch) => onChange({ bold, italic, underline, ...patch });
      return (
        <div>
          <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">{label}</label>
          <div className="flex flex-wrap items-center gap-2 text-xs text-neutral-400">
            <label className={`inline-flex h-8 items-center gap-1.5 rounded-md border border-neutral-800 px-2 cursor-pointer ${bold ? 'bg-neutral-900 text-neutral-100 font-bold' : ''}`}>
              <input type="checkbox" checked={!!bold} onChange={e => update({ bold: e.target.checked })} className="h-3.5 w-3.5" />
              <span className="font-bold">B</span>
            </label>
            <label className={`inline-flex h-8 items-center gap-1.5 rounded-md border border-neutral-800 px-2 cursor-pointer ${italic ? 'bg-neutral-900 text-neutral-100 italic' : ''}`}>
              <input type="checkbox" checked={!!italic} onChange={e => update({ italic: e.target.checked })} className="h-3.5 w-3.5" />
              <span className="italic">I</span>
            </label>
            <label className={`inline-flex h-8 items-center gap-1.5 rounded-md border border-neutral-800 px-2 cursor-pointer ${underline ? 'bg-neutral-900 text-neutral-100 underline' : ''}`}>
              <input type="checkbox" checked={!!underline} onChange={e => update({ underline: e.target.checked })} className="h-3.5 w-3.5" />
              <span className="underline">U</span>
            </label>
          </div>
        </div>
      );
    }

    function EditorHeaderSettings({ editor, onChange }) {
      const update = (patch) => onChange({ ...editor, ...patch });
      const grid = editorNormalizeGrid(editor.cells);
      const rowCount = grid.length;
      const rowStyles = Array.isArray(editor.rowStyles) ? editor.rowStyles : [];
      const [bulkRows, setBulkRows] = useState('');
      const [bulkRowFill, setBulkRowFill] = useState('#000000');
      const [bulkRowFontColor, setBulkRowFontColor] = useState('#FFFFFF');
      const [bulkRowFontStyle, setBulkRowFontStyle] = useState({ bold: true, italic: false, underline: false });
      const updateRowStyle = (id, patch) => update({ rowStyles: rowStyles.map(style => style.id === id ? { ...style, ...patch } : style) });
      const removeRowStyle = (id) => update({ rowStyles: rowStyles.filter(style => style.id !== id) });
      const addRowStyle = () => {
        const lastRow = Math.max(0, ...rowStyles.map(style => Math.floor(Number(style.row)) || 0));
        const defaultRow = Math.min(rowCount, Math.max(1, lastRow ? lastRow + 1 : (Math.floor(Number(editor.headerRow)) || 1) + 6));
        update({
          rowStyles: [
            ...rowStyles,
            { id: editorRuleId(), row: defaultRow, fill: '#000000', fontColor: '#FFFFFF', bold: true, italic: false, underline: false },
          ],
        });
      };
      const addBulkRowStyles = () => {
        const rows = editorParseRowNumbers(bulkRows, rowCount);
        if (!rows.length) return;
        const rowSet = new Set(rows);
        const nextRows = [
          ...rowStyles.filter(style => !rowSet.has(Math.floor(Number(style.row)) || 0)),
          ...rows.map(row => ({
            id: editorRuleId(),
            row,
            fill: bulkRowFill,
            fontColor: bulkRowFontColor,
            bold: !!bulkRowFontStyle.bold,
            italic: !!bulkRowFontStyle.italic,
            underline: !!bulkRowFontStyle.underline,
          })),
        ].sort((a, b) => (Number(a.row) || 0) - (Number(b.row) || 0));
        update({ rowStyles: nextRows });
        setBulkRows('');
      };
      return (
        <Card className="p-4">
          <SectionLabel hint="Header setup">Table style</SectionLabel>
          <div className="grid gap-3 lg:grid-cols-4">
            <div className="lg:col-span-2">
              <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">Sheet name</label>
              <Input value={editor.name || ''} onChange={e => update({ name: e.target.value })} placeholder="Pasted Table" />
            </div>
            <div>
              <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">Header row #</label>
              <Input type="number" min="0" value={editor.headerRow ?? 1} onChange={e => update({ headerRow: Math.max(0, Number(e.target.value) || 0) })} />
            </div>
            <div>
              <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">Date / label col #</label>
              <Input type="number" min="0" value={editor.labelColumn ?? 1} onChange={e => update({ labelColumn: Math.max(0, Number(e.target.value) || 0) })} />
            </div>
          </div>
          <div className="mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
            <EditorColorField label="Header fill" value={editor.headerFill} onChange={headerFill => update({ headerFill })} title="Choose header fill" />
            <EditorColorField label="Header font" value={editor.headerFontColor} onChange={headerFontColor => update({ headerFontColor })} title="Choose header font color" fallback="#000000" />
            <EditorColorField label="Date / label fill" value={editor.labelFill} onChange={labelFill => update({ labelFill })} title="Choose label fill" />
            <EditorColorField label="Date / label font" value={editor.labelFontColor} onChange={labelFontColor => update({ labelFontColor })} title="Choose label font color" fallback="#000000" />
          </div>
          <div className="mt-4 grid gap-3 sm:grid-cols-2">
            <EditorFontStyleControls label="Header font style"
              bold={editor.headerBold !== false} italic={!!editor.headerItalic} underline={!!editor.headerUnderline}
              onChange={style => update({ headerBold: style.bold, headerItalic: style.italic, headerUnderline: style.underline })} />
            <EditorFontStyleControls label="Date / label font style"
              bold={!!editor.labelBold} italic={!!editor.labelItalic} underline={!!editor.labelUnderline}
              onChange={style => update({ labelBold: style.bold, labelItalic: style.italic, labelUnderline: style.underline })} />
          </div>
          <div className="mt-5 border-t border-neutral-900 pt-4">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
              <SectionLabel hint="Color one row or many rows by row number">Custom row colors</SectionLabel>
              <Btn variant="ghost" size="sm" onClick={addRowStyle}><IconPlus /> Row color</Btn>
            </div>
            <div className="mb-3 rounded-md border border-neutral-900 bg-neutral-950/60 p-3">
              <div className="grid gap-3 md:grid-cols-[minmax(180px,1fr)_minmax(150px,auto)_minmax(150px,auto)_190px_auto] md:items-end">
                <div>
                  <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">Rows</label>
                  <Input value={bulkRows} onChange={e => setBulkRows(e.target.value)} onKeyDown={e => e.key === 'Enter' && (e.preventDefault(), addBulkRowStyles())}
                    placeholder="7, 14, 21 or 7-12" />
                </div>
                <EditorColorField label="Fill" value={bulkRowFill} onChange={setBulkRowFill} title="Choose bulk row fill" />
                <EditorColorField label="Font" value={bulkRowFontColor} onChange={setBulkRowFontColor} title="Choose bulk row font color" fallback="#000000" />
                <EditorFontStyleControls label="Style"
                  bold={bulkRowFontStyle.bold} italic={bulkRowFontStyle.italic} underline={bulkRowFontStyle.underline}
                  onChange={setBulkRowFontStyle} />
                <Btn variant="ghost" size="sm" onClick={addBulkRowStyles} disabled={!bulkRows.trim()}><IconPlus /> Add rows</Btn>
              </div>
            </div>
            {rowStyles.length === 0 ? (
              <div className="rounded-md border border-dashed border-neutral-800 px-3 py-4 text-xs text-neutral-500">
                Add row colors for repeated date/header rows such as row 7.
              </div>
            ) : (
              <div className="space-y-2">
                {rowStyles.map((style) => (
                  <div key={style.id} className="grid gap-2 rounded-md border border-neutral-900 bg-neutral-950/60 p-2 md:grid-cols-[96px_minmax(150px,1fr)_minmax(150px,1fr)_190px_auto] md:items-end">
                    <div>
                      <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">Row #</label>
                      <Input type="number" min="1" max={rowCount} value={style.row ?? 1}
                        onChange={e => updateRowStyle(style.id, { row: Math.min(rowCount, Math.max(1, Number(e.target.value) || 1)) })} />
                    </div>
                    <EditorColorField label="Fill" value={style.fill} onChange={fill => updateRowStyle(style.id, { fill })} title="Choose row fill" />
                    <EditorColorField label="Font" value={style.fontColor} onChange={fontColor => updateRowStyle(style.id, { fontColor })} title="Choose row font color" fallback="#000000" />
                    <EditorFontStyleControls label="Style"
                      bold={style.bold !== false} italic={!!style.italic} underline={!!style.underline}
                      onChange={next => updateRowStyle(style.id, next)} />
                    <button onClick={() => removeRowStyle(style.id)} title="Remove row color" className="h-8 px-2 text-neutral-500 hover:text-red-400 justify-self-start md:justify-self-end"><IconX /></button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </Card>
      );
    }

    function EditorTemplatesPanel({ editor, onChange, sync }) {
      const [templates, setTemplates] = useState(() => {
        try { return JSON.parse(localStorage.getItem(EDITOR_TEMPLATES_KEY)) || []; } catch { return []; }
      });
      const [name, setName] = useState('');

      useEffect(() => {
        if (!window.__fb || !window.__fbm || !sync?.uid) return;
        let cancelled = false;
        (async () => {
          try {
            const { doc, getDoc } = window.__fbm;
            const ref = doc(window.__fb.db, 'users', sync.uid, 'state', 'editorTemplates');
            const snap = await getDoc(ref);
            if (cancelled) return;
            const remote = snap.exists() ? (firestoreDesanitize(snap.data()).items || []) : null;
            if (Array.isArray(remote) && remote.length) {
              const merged = [...remote];
              for (const local of templates) {
                if (!merged.some(t => t.name === local.name)) merged.push(local);
              }
              setTemplates(merged);
              localStorage.setItem(EDITOR_TEMPLATES_KEY, JSON.stringify(merged));
            }
          } catch (e) { console.warn('editor templates hydrate failed', e); }
        })();
        return () => { cancelled = true; };
      }, [sync?.uid]);

      const persist = (next) => {
        setTemplates(next);
        localStorage.setItem(EDITOR_TEMPLATES_KEY, JSON.stringify(next));
        if (window.__fb && window.__fbm && sync?.uid) {
          const { doc, setDoc, serverTimestamp } = window.__fbm;
          const ref = doc(window.__fb.db, 'users', sync.uid, 'state', 'editorTemplates');
          setDoc(ref, { items: firestoreSanitize(next), updatedAt: serverTimestamp() })
            .catch(e => console.warn('editor templates save failed', e));
        }
      };
      const save = () => {
        const cleanName = name.trim();
        if (!cleanName) return;
        const t = { id: newId(), name: cleanName, savedAt: new Date().toISOString(), data: editorTemplateData(editor) };
        persist([...templates.filter(x => x.name !== t.name), t]);
        setName('');
      };
      const load = (id) => {
        const t = templates.find(x => x.id === id);
        if (!t) return;
        onChange(current => editorApplyTemplate(current, t.data));
      };
      const del = (id) => persist(templates.filter(x => x.id !== id));

      return (
        <Card className="p-4">
          <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
            <SectionLabel hint="Save style and rules without table data">Formatting templates</SectionLabel>
            <Pill>{templates.length} saved</Pill>
          </div>
          <div className="grid gap-2 md:grid-cols-[minmax(220px,1fr)_auto]">
            <Input value={name} onChange={e => setName(e.target.value)} onKeyDown={e => e.key === 'Enter' && (e.preventDefault(), save())}
              placeholder="Template name" />
            <Btn variant="ghost" size="md" onClick={save} disabled={!name.trim()}>Save current</Btn>
          </div>
          {templates.length > 0 && (
            <div className="mt-3 space-y-2">
              {templates.map(t => (
                <div key={t.id} className="flex flex-wrap items-center gap-2 rounded-md border border-neutral-900 bg-neutral-950/60 px-3 py-2">
                  <div className="min-w-0 flex-1">
                    <div className="truncate text-sm font-medium text-neutral-200">{t.name}</div>
                    <div className="text-[10px] text-neutral-600 font-mono">{new Date(t.savedAt).toLocaleDateString()}</div>
                  </div>
                  <Btn variant="ghost" size="sm" onClick={() => load(t.id)}>Use</Btn>
                  <button onClick={() => del(t.id)} title="Delete template" className="p-1.5 text-neutral-500 hover:text-red-400"><IconX /></button>
                </div>
              ))}
            </div>
          )}
        </Card>
      );
    }

    function editorCellPreviewStyle(editor, grid, rowIndex, colIndex) {
      const headerRow = editorHeaderRow(editor, grid.length);
      const labelCol = editorLabelColumn(editor, grid[0].length);
      const rowNum = rowIndex + 1;
      const colNum = colIndex + 1;
      const style = {};
      if (headerRow && rowNum === headerRow) {
        if (editor.headerFill) style.background = editor.headerFill;
        if (editor.headerFontColor) style.color = editor.headerFontColor;
        style.fontWeight = editor.headerBold !== false ? 700 : 400;
        if (editor.headerItalic) style.fontStyle = 'italic';
        if (editor.headerUnderline) style.textDecoration = 'underline';
      }
      if (labelCol && colNum === labelCol && rowNum !== headerRow) {
        if (editor.labelFill) style.background = editor.labelFill;
        if (editor.labelFontColor) style.color = editor.labelFontColor;
        if (editor.labelBold) style.fontWeight = 700;
        if (editor.labelItalic) style.fontStyle = 'italic';
        if (editor.labelUnderline) style.textDecoration = 'underline';
      }
      const customRowStyle = editorCustomRowStyleFor(editor, grid.length, rowNum);
      if (customRowStyle) {
        if (customRowStyle.fill) style.background = customRowStyle.fill;
        if (customRowStyle.fontColor) style.color = customRowStyle.fontColor;
        if (typeof customRowStyle.bold === 'boolean') style.fontWeight = customRowStyle.bold ? 700 : 400;
        if (customRowStyle.italic) style.fontStyle = 'italic';
        if (customRowStyle.underline) style.textDecoration = 'underline';
      }
      const isData = rowNum > (headerRow || 0) && colNum > (labelCol || 0);
      if (!isData) return style;
      (editor.rules || []).forEach((rule) => {
        if (!editorEvaluateRule(rule, grid[rowIndex][colIndex], grid[rowIndex], labelCol ? labelCol + 1 : 1)) return;
        if (rule.color) style.background = rule.color;
        if (rule.fontColor) style.color = rule.fontColor;
        if (rule.bold) style.fontWeight = 700;
        if (rule.italic) style.fontStyle = 'italic';
        if (rule.underline) style.textDecoration = 'underline';
      });
      return style;
    }

    function EditorTable({ editor, onChange }) {
      const grid = editorNormalizeGrid(editor.cells);
      const updateCell = (rowIndex, colIndex, value) => {
        const next = grid.map(row => row.slice());
        next[rowIndex][colIndex] = value;
        onChange({ ...editor, cells: next });
      };
      const addRow = () => onChange({ ...editor, cells: [...grid, Array(grid[0].length).fill('')] });
      const addColumn = () => onChange({ ...editor, cells: grid.map(row => [...row, '']) });
      return (
        <Card className="p-4">
          <div className="flex flex-wrap items-center justify-between gap-3 mb-3">
            <SectionLabel hint={`${grid.length} rows x ${grid[0].length} columns`}>Table</SectionLabel>
            <div className="flex gap-2">
              <Btn variant="ghost" size="sm" onClick={addRow}><IconPlus /> Row</Btn>
              <Btn variant="ghost" size="sm" onClick={addColumn}><IconPlus /> Column</Btn>
            </div>
          </div>
          <div className="overflow-auto rounded-md border border-neutral-900">
            <table className="border-collapse text-xs">
              <thead>
                <tr className="bg-neutral-950 text-neutral-500">
                  <th className="sticky left-0 z-10 min-w-10 border border-neutral-900 bg-neutral-950 px-2 py-1"></th>
                  {grid[0].map((_, colIndex) => (
                    <th key={colIndex} className="min-w-[136px] border border-neutral-900 px-2 py-1 font-mono">{colLetter(colIndex + 1)}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {grid.map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    <th className="sticky left-0 z-10 border border-neutral-900 bg-neutral-950 px-2 py-1 text-neutral-500 font-mono">{rowIndex + 1}</th>
                    {row.map((value, colIndex) => {
                      const style = editorCellPreviewStyle(editor, grid, rowIndex, colIndex);
                      const inputStyle = {
                        color: style.color,
                        fontWeight: style.fontWeight,
                        fontStyle: style.fontStyle,
                        textDecoration: style.textDecoration,
                      };
                      return (
                        <td key={colIndex} className="border border-neutral-800 p-0" style={style}>
                          <input value={value} onChange={e => updateCell(rowIndex, colIndex, e.target.value)}
                            aria-label={`Cell ${colLetter(colIndex + 1)}${rowIndex + 1}`}
                            style={inputStyle}
                            className="h-9 w-full min-w-[136px] bg-transparent px-2 text-current outline-none focus:ring-2 focus:ring-blue-500/40" />
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      );
    }

    function EditorRuleEditor({ rule, onChange, onDelete }) {
      const update = (patch) => onChange({ ...rule, ...patch });
      return (
        <Card className="p-3">
          <div className="mb-2 flex items-center gap-2">
            <input type="checkbox" checked={rule.enabled !== false} onChange={e => update({ enabled: e.target.checked })}
              className="h-3.5 w-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500" />
            <Input value={rule.name || ''} onChange={e => update({ name: e.target.value })} placeholder="Rule name" />
            <button onClick={onDelete} title="Delete rule" className="p-1.5 text-neutral-500 hover:text-red-400"><IconX /></button>
          </div>
          <div className="mb-2 grid gap-2 md:grid-cols-[240px_1fr_1fr] items-end">
            <Select value={rule.kind || 'gte'} onChange={e => update({ kind: e.target.value })}>
              <option value="gte">greater than or equal to</option>
              <option value="gt">greater than</option>
              <option value="lte">less than or equal to</option>
              <option value="lt">less than</option>
              <option value="eq">equal to</option>
              <option value="between">between</option>
              <option value="maxInRow">highest value in row</option>
            </Select>
            {rule.kind === 'maxInRow' ? (
              <div className="md:col-span-2 h-9 rounded-md border border-neutral-900 bg-neutral-950/60"></div>
            ) : rule.kind === 'between' ? (
              <>
                <div>
                  <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Min</label>
                  <Input type="number" step="0.001" value={rule.min ?? ''} onChange={e => update({ min: Number(e.target.value) })} />
                </div>
                <div>
                  <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Max</label>
                  <Input type="number" step="0.001" value={rule.max ?? ''} onChange={e => update({ max: Number(e.target.value) })} />
                </div>
              </>
            ) : (
              <div className="md:col-span-2">
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Value</label>
                <Input type="number" step="0.001" value={rule.value ?? ''} onChange={e => update({ value: Number(e.target.value) })} />
              </div>
            )}
          </div>
          <div className="flex flex-wrap items-center gap-3 text-[11px] text-neutral-400">
            <span className="inline-flex items-center gap-1"><span className="text-[10px] uppercase tracking-wide text-neutral-500">Fill</span><EditorColorSwatch value={rule.color} onChange={color => update({ color })} title="Choose rule fill" /></span>
            <span className="inline-flex items-center gap-1"><span className="text-[10px] uppercase tracking-wide text-neutral-500">Font</span><EditorColorSwatch value={rule.fontColor} onChange={fontColor => update({ fontColor })} title="Choose rule font color" fallback="#000000" /></span>
            <label className={`inline-flex items-center gap-1.5 cursor-pointer ${rule.bold ? 'text-neutral-100 font-bold' : ''}`}><input type="checkbox" checked={!!rule.bold} onChange={e => update({ bold: e.target.checked })} className="h-3.5 w-3.5" /><span className="font-bold">B</span></label>
            <label className={`inline-flex items-center gap-1.5 cursor-pointer ${rule.italic ? 'text-neutral-100 italic' : ''}`}><input type="checkbox" checked={!!rule.italic} onChange={e => update({ italic: e.target.checked })} className="h-3.5 w-3.5" /><span className="italic">I</span></label>
            <label className={`inline-flex items-center gap-1.5 cursor-pointer ${rule.underline ? 'text-neutral-100 underline' : ''}`}><input type="checkbox" checked={!!rule.underline} onChange={e => update({ underline: e.target.checked })} className="h-3.5 w-3.5" /><span className="underline">U</span></label>
          </div>
        </Card>
      );
    }

    function EditorRulesPanel({ editor, onChange }) {
      const rules = editor.rules || [];
      const addRule = (kind = 'gte') => onChange({
        ...editor,
        rules: [...rules, {
          id: editorRuleId(),
          name: kind === 'maxInRow' ? 'Highest in row' : 'Reached value',
          kind,
          value: 5,
          min: 0,
          max: 5,
          color: kind === 'maxInRow' ? '#FACC15' : '#EF4444',
          fontColor: '#FFFFFF',
          bold: true,
          italic: false,
          underline: false,
          enabled: true,
        }],
      });
      const updateRule = (id, next) => onChange({ ...editor, rules: rules.map(rule => rule.id === id ? next : rule) });
      const deleteRule = (id) => onChange({ ...editor, rules: rules.filter(rule => rule.id !== id) });
      return (
        <div className="space-y-3">
          <SectionLabel hint="Data cells">Table rules</SectionLabel>
          {rules.map(rule => <EditorRuleEditor key={rule.id} rule={rule} onChange={next => updateRule(rule.id, next)} onDelete={() => deleteRule(rule.id)} />)}
          <div className="flex flex-wrap gap-2">
            <Btn variant="ghost" size="md" onClick={() => addRule('gte')}><IconPlus /> Value rule</Btn>
            <Btn variant="ghost" size="md" onClick={() => addRule('maxInRow')}><IconPlus /> Highest in row</Btn>
          </div>
        </div>
      );
    }

    /* ============================================================
       Whitelist SMS - UI Components
       ============================================================ */
    function WlSmsSidebar({ state, setState, onGenerate, busy, sync, onRetrySync, sharedTestNumbers }) {
      const theme = state.theme || 'dark';
      const wl = wlSmsNormalizeState(state.whitelistSms);
      const toggleTheme = () => setState(s => ({ ...s, theme: (s.theme || 'dark') === 'dark' ? 'light' : 'dark' }));
      const setWl = (next) => setState(s => {
        const prev = wlSmsNormalizeState(s.whitelistSms);
        return { ...s, whitelistSms: typeof next === 'function' ? next(prev) : next };
      });
      const toggleNetwork = (id) => setWl(prev => {
        const set = new Set(prev.networks || []);
        if (set.has(id)) set.delete(id); else set.add(id);
        return { ...prev, networks: WL_SMS_NETWORK_ORDER.filter(k => set.has(k)) };
      });
      const numbers = wlSmsEffectiveNumbers(wl, sharedTestNumbers);
      return (
        <aside className="w-[320px] shrink-0 border-r border-neutral-900 bg-[#17171a] h-screen sticky top-0 flex flex-col">
          <div className="p-5 border-b border-neutral-900">
            <div className="flex items-center justify-between gap-2">
              <SyncBadge sync={sync} onRetry={onRetrySync} />
              <button onClick={toggleTheme}
                title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}
                className="inline-flex items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2 py-1 text-[10px] uppercase tracking-wide text-neutral-400 hover:text-neutral-100 hover:border-neutral-700 transition-colors">
                {theme === 'dark' ? <IconSun /> : <IconMoon />}
                {theme === 'dark' ? 'Light' : 'Dark'}
              </button>
            </div>
            <h1 className="text-[17px] font-bold tracking-tight leading-tight mt-2">Whitelist SMS</h1>
            <p className="text-xs text-neutral-500 mt-1">Content blast file generator</p>
            <AccountChip sync={sync} />
          </div>

          <div className="p-5 border-b border-neutral-900">
            <SectionLabel>Sender ID</SectionLabel>
            <Input value={wl.senderId} placeholder="e.g. JILIBET"
              onChange={e => setWl({ ...wl, senderId: e.target.value })} />
          </div>

          <div className="p-5 border-b border-neutral-900 space-y-2">
            <SectionLabel hint="Multi-select">Networks</SectionLabel>
            {WL_SMS_NETWORK_ORDER.map(key => {
              const net = WL_SMS_NETWORKS[key];
              const on = (wl.networks || []).includes(key);
              return (
                <label key={key} className="flex items-center gap-2 text-xs text-neutral-300 cursor-pointer">
                  <input type="checkbox" checked={on}
                    onChange={() => toggleNetwork(key)}
                    className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500" />
                  <span className="inline-block w-2 h-2 rounded-full" style={{ background: net.color }}></span>
                  <span>{net.label}</span>
                </label>
              );
            })}
          </div>

          <div className="p-5 flex-1 overflow-y-auto min-h-0">
            <SectionLabel hint={`${numbers.length} ready`}>Will generate</SectionLabel>
            <div className="rounded-md border border-neutral-900 bg-neutral-950 p-3 space-y-2">
              <div className="flex flex-wrap gap-1.5">
                <Pill>{(wl.contents || []).length} contents</Pill>
                <Pill>{numbers.length} numbers</Pill>
                <Pill tone={(wl.contents || []).length && numbers.length ? 'accent' : 'muted'}>
                  {(wl.contents || []).length * numbers.length} rows
                </Pill>
              </div>
              {numbers.slice(0, 8).map(n => {
                const net = WL_SMS_NETWORKS[n.network];
                return (
                  <div key={n.id} className="flex items-center gap-2 text-[11px] text-neutral-300">
                    <span className="inline-block w-2 h-2 rounded-full shrink-0" style={{ background: net?.color || '#666' }}></span>
                    <span className="font-mono">{n.number}</span>
                    <span className="text-neutral-600 truncate">· {n.label}</span>
                  </div>
                );
              })}
            </div>
          </div>

          <div className="p-5 border-t border-neutral-900 space-y-3">
            <Btn variant="primary" size="lg" onClick={onGenerate} disabled={busy} className="w-full">
              {busy ? <><span className="loader"></span> Generating...</> : <><IconDownload /> Generate Whitelist</>}
            </Btn>
            <p className="text-[10px] text-neutral-600 text-center font-mono">.xlsx · Number + Content per row</p>
          </div>
        </aside>
      );
    }

    function WlSmsLinkBuilder({ wl, onChange }) {
      const raw = wl.linkBuilderRaw || '';
      const built = wlSmsBuildLinkMessages(raw);
      const [copyStatus, setCopyStatus] = useState('');
      const setRaw = (value) => onChange({ ...wl, linkBuilderRaw: value });
      const copyOutput = async () => {
        if (!built.output) return;
        const ok = await copyTextToClipboard(built.output);
        setCopyStatus(ok ? 'Copied' : 'Copy failed');
        setTimeout(() => setCopyStatus(''), 1600);
      };
      const useAsContents = () => {
        if (!built.messages.length) return;
        onChange({
          ...wl,
          linkBuilderRaw: raw,
          rawPaste: built.output,
          contents: built.messages,
          separatorMode: 'auto',
        });
      };
      return (
        <Card className="p-4">
          <SectionLabel hint={`${built.urls.length} link${built.urls.length === 1 ? '' : 's'} detected`}>Link message builder</SectionLabel>
          <div className="grid gap-4 xl:grid-cols-2">
            <div>
              <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Paste template + links</label>
              <Textarea rows={10} value={raw}
                onChange={e => setRaw(e.target.value)}
                placeholder={`Paste one message template followed by the links.\n\nTip: use {{link}} or {{url}} inside the template if the link should appear in the middle.`} />
              <div className="mt-2 flex flex-wrap items-center gap-2">
                <Btn variant="ghost" size="md" onClick={() => setRaw(WL_SMS_LINK_BUILDER_SAMPLE)}>Load sample</Btn>
                <Btn variant="ghost" size="md" onClick={() => setRaw('')} disabled={!raw}>Clear</Btn>
                {raw && !built.urls.length && <span className="text-[11px] text-amber-400">No links detected yet.</span>}
              </div>
            </div>
            <div>
              <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Generated messages</label>
              <Textarea rows={10} readOnly value={built.output}
                className="font-mono text-xs"
                placeholder="Each detected link will become one complete SMS message here." />
              <div className="mt-2 flex flex-wrap items-center gap-2">
                <Btn variant="accent" size="md" onClick={useAsContents} disabled={!built.messages.length}>Use as contents</Btn>
                <Btn variant="ghost" size="md" onClick={copyOutput} disabled={!built.output}>Copy output</Btn>
                <span className="text-[11px] text-neutral-500">{copyStatus || `${built.messages.length} message${built.messages.length === 1 ? '' : 's'}`}</span>
              </div>
              {built.template && (
                <div className="mt-3 rounded-md border border-neutral-900 bg-neutral-950 px-3 py-2">
                  <div className="text-[10px] uppercase tracking-wide text-neutral-500 mb-1">Template detected</div>
                  <div className="text-[12px] text-neutral-300 break-words">{built.template}</div>
                </div>
              )}
            </div>
          </div>
        </Card>
      );
    }

    function WlSmsContentPanel({ wl, onChange }) {
      const [pasteText, setPasteText] = useState(wl.rawPaste || '');
      const [copiedIndex, setCopiedIndex] = useState(null);
      const mode = wl.separatorMode === 'lines' ? 'lines' : 'auto';
      useEffect(() => {
        setPasteText(wl.rawPaste || '');
      }, [wl.rawPaste]);
      const loadText = (text, useMode = mode) => {
        const blocks = parseWlSmsBlasts(text, useMode);
        onChange({ ...wl, rawPaste: text, contents: blocks, separatorMode: useMode });
      };
      const onPaste = (e) => {
        const text = e.clipboardData?.getData('text/plain');
        if (text === undefined || text === null || text === '') return;
        e.preventDefault();
        setPasteText(text);
        loadText(text);
      };
      const setMode = (nextMode) => {
        if (nextMode === mode) return;
        if (pasteText) loadText(pasteText, nextMode);
        else onChange({ ...wl, separatorMode: nextMode });
      };
      const updateBlock = (index, value) => {
        const next = [...(wl.contents || [])];
        next[index] = value;
        onChange({ ...wl, contents: next });
      };
      const deleteBlock = (index) => {
        const next = (wl.contents || []).filter((_, i) => i !== index);
        onChange({ ...wl, contents: next });
      };
      const copyBlock = async (block, index) => {
        if (!String(block || '').trim()) return;
        const ok = await copyTextToClipboard(block);
        if (!ok) return;
        setCopiedIndex(index);
        setTimeout(() => setCopiedIndex(null), 1400);
      };
      const addBlock = () => onChange({ ...wl, contents: [...(wl.contents || []), ''] });
      const clearAll = () => { setPasteText(''); onChange({ ...wl, rawPaste: '', contents: [] }); };
      // Live UTF check on whatever is in the paste box, mirroring the UTF Check
      // tab so special characters are flagged the moment content is pasted.
      const pasteAnalysis = wlSmsAnalyzeUtf(pasteText);
      const pasteFixable = wlSmsApplySafeFixes(pasteText) !== pasteText;
      const applyPasteFixes = () => {
        const fixed = wlSmsApplySafeFixes(pasteText);
        setPasteText(fixed);
        loadText(fixed);
      };
      const modeHint = mode === 'lines'
        ? 'Every non-empty line becomes one content'
        : 'Splits on ———— / === lines, or blank lines if no dashes';
      return (
        <div className="space-y-4">
          <WlSmsLinkBuilder wl={wl} onChange={onChange} />

          <Card className="p-4">
            <SectionLabel hint={modeHint}>Paste content</SectionLabel>
            <div className="mb-2 flex flex-wrap items-center gap-1 rounded-md border border-neutral-800 bg-[#1a1a1d] p-0.5 w-fit">
              {[
                { id: 'auto',  label: 'With separator' },
                { id: 'lines', label: 'Without separator' },
              ].map(opt => (
                <button key={opt.id} onClick={() => setMode(opt.id)}
                  className={`px-3 py-1 text-[11px] font-semibold uppercase tracking-wider rounded transition-colors ${mode === opt.id ? 'bg-blue-500/20 text-blue-200 border border-blue-500/40' : 'text-neutral-500 hover:text-neutral-200'}`}>
                  {opt.label}
                </button>
              ))}
            </div>
            <Textarea rows={6} value={pasteText} onPaste={onPaste}
              onChange={e => setPasteText(e.target.value)}
              placeholder={mode === 'lines'
                ? `Paste blast content here. Each non-empty line will be treated as one content.`
                : `Paste blast content here. Use a divider line like ———————— between each content, or leave a blank line between them.`} />
            <div className="mt-2 flex flex-wrap items-center gap-2">
              <Btn variant="ghost" size="md" onClick={() => loadText(pasteText)} disabled={!pasteText}>Parse pasted text</Btn>
              <Btn variant="ghost" size="md" onClick={clearAll} disabled={!pasteText && !(wl.contents || []).length}>Clear all</Btn>
              <span className="text-[11px] text-neutral-500">{(wl.contents || []).length} content block{(wl.contents || []).length === 1 ? '' : 's'}</span>
              {pasteText.trim() && <WlSmsUtfStatusPill analysis={pasteAnalysis} />}
            </div>
            {pasteText.trim() && pasteAnalysis.isUtf && (
              <div className="mt-3 space-y-2 rounded-md border border-red-500/30 bg-red-500/5 p-3">
                <div className="flex flex-wrap items-center gap-2">
                  <span className="text-[11px] font-semibold text-red-300">Special characters detected — this will be sent as UTF (70 chars per SMS part).</span>
                  {pasteFixable && (
                    <Btn variant="ghost" size="sm" onClick={applyPasteFixes}>Apply safe replacements</Btn>
                  )}
                </div>
                <WlSmsUtfHighlight segments={pasteAnalysis.segments} />
                <WlSmsUtfOffenderChips offenders={pasteAnalysis.offenders} />
              </div>
            )}
          </Card>

          <Card className="p-4">
            <SectionLabel hint="One block = one row in column B">Parsed contents</SectionLabel>
            <div className="space-y-2">
              {(wl.contents || []).length === 0 && (
                <div className="text-[12px] text-neutral-500 italic">Nothing parsed yet. Paste content above.</div>
              )}
              {(wl.contents || []).map((block, i) => (
                <div key={i} className="rounded-md border border-neutral-900 bg-neutral-950 p-2">
                  <div className="flex items-center justify-between mb-1">
                    <span className="text-[10px] uppercase tracking-wide text-neutral-500">Content {i + 1}</span>
                    <div className="flex items-center gap-3">
                      <button onClick={() => copyBlock(block, i)}
                        className="text-[10px] text-neutral-500 hover:text-blue-300 disabled:opacity-40"
                        disabled={!String(block || '').trim()}>
                        {copiedIndex === i ? 'Copied' : 'Copy'}
                      </button>
                      <button onClick={() => deleteBlock(i)} className="text-[10px] text-neutral-500 hover:text-red-300">Remove</button>
                    </div>
                  </div>
                  <Textarea rows={3} value={block} onChange={e => updateBlock(i, e.target.value)} />
                </div>
              ))}
              <div>
                <Btn variant="ghost" size="sm" onClick={addBlock}><IconPlus /> Add empty content</Btn>
              </div>
            </div>
          </Card>
        </div>
      );
    }

    function WlSmsUtfHighlight({ segments }) {
      return (
        <div className="whitespace-pre-wrap break-words rounded-md border border-neutral-900 bg-neutral-950 p-3 text-[12px] leading-relaxed text-neutral-200">
          {segments.length === 0 && <span className="text-neutral-500 italic">Empty</span>}
          {segments.map((s, i) => s.special ? (
            <span key={i} className="rounded border border-red-500/30 bg-red-500/10 px-0.5 font-semibold text-red-300">
              {[...s.text].map(wlSmsCharVisible).join('')}
            </span>
          ) : (
            <span key={i}>{s.text}</span>
          ))}
        </div>
      );
    }

    function WlSmsUtfOffenderChips({ offenders }) {
      const entries = Array.from(offenders.entries());
      if (!entries.length) return null;
      return (
        <div className="mt-2 flex flex-wrap gap-1.5">
          {entries.map(([ch, count]) => {
            const info = WL_SMS_SPECIAL_CHAR_INFO[ch];
            const name = info?.name || (ch.codePointAt(0) > 0xFFFF ? 'Emoji / symbol' : 'Non-GSM character');
            return (
              <span key={ch} className="inline-flex items-center gap-1.5 rounded-md border border-amber-500/35 bg-amber-500/10 px-2 py-1 text-[11px] text-amber-300">
                <span className="font-mono text-[13px] font-bold">{wlSmsCharVisible(ch)}</span>
                <span className="font-mono text-[10px] opacity-70">{wlSmsCodePointLabel(ch)}</span>
                <span>{name}</span>
                <span className="opacity-70">×{count}</span>
                {info?.fix !== undefined && (
                  <span className="opacity-70">→ {info.fix === '' ? 'remove' : JSON.stringify(info.fix)}</span>
                )}
              </span>
            );
          })}
        </div>
      );
    }

    function WlSmsUtfStatusPill({ analysis }) {
      const specialCount = Array.from(analysis.offenders.values()).reduce((a, b) => a + b, 0);
      return analysis.isUtf ? (
        <span className="inline-flex items-center rounded-md border border-red-500/30 bg-red-500/10 px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wider text-red-300">
          UTF · {specialCount} special char{specialCount === 1 ? '' : 's'}
        </span>
      ) : (
        <span className="inline-flex items-center rounded-md border border-emerald-500/30 bg-emerald-500/10 px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wider text-emerald-300">
          GSM-7 · safe
        </span>
      );
    }

    function WlSmsUtfCheckPanel({ wl, onChange }) {
      const [adhoc, setAdhoc] = useState('');
      const [copied, setCopied] = useState(null); // 'all' | 'adhoc' | null
      const contents = wl.contents || [];
      const analyses = contents.map(c => wlSmsAnalyzeUtf(c));
      const utfBlocks = analyses.filter(a => a.isUtf).length;
      const adhocAnalysis = wlSmsAnalyzeUtf(adhoc);
      const isFixable = (text) => wlSmsApplySafeFixes(text) !== text;
      const fixBlock = (i) => {
        const next = [...contents];
        next[i] = wlSmsApplySafeFixes(next[i]);
        onChange({ ...wl, contents: next });
      };
      const fixAll = () => onChange({ ...wl, contents: contents.map(wlSmsApplySafeFixes) });
      const anyFixable = contents.some(isFixable);
      const copyText = async (text, which) => {
        const ok = await copyTextToClipboard(text);
        if (!ok) return;
        setCopied(which);
        setTimeout(() => setCopied(null), 1400);
      };
      const copyAll = () => copyText(contents.filter(s => s && s.trim()).join('\n\n'), 'all');
      // Mirrors the Content tab's Clear all (drops parsed blocks + raw paste),
      // and also empties the quick-check box.
      const clearAll = () => { setAdhoc(''); onChange({ ...wl, rawPaste: '', contents: [] }); };
      return (
        <div className="space-y-4">
          <Card className="p-4">
            <SectionLabel hint="Paste anything here to check it without touching the content blocks">Quick check</SectionLabel>
            <Textarea rows={4} value={adhoc} onChange={e => setAdhoc(e.target.value)}
              placeholder="Paste text here to see which characters force UTF..." />
            {adhoc.trim() && (
              <div className="mt-2 space-y-2">
                <div className="flex flex-wrap items-center gap-2">
                  <WlSmsUtfStatusPill analysis={adhocAnalysis} />
                  {adhocAnalysis.isUtf && isFixable(adhoc) && (
                    <Btn variant="ghost" size="sm" onClick={() => setAdhoc(wlSmsApplySafeFixes(adhoc))}>Apply safe replacements</Btn>
                  )}
                  <Btn variant="ghost" size="sm" onClick={() => copyText(adhoc, 'adhoc')}>{copied === 'adhoc' ? 'Copied' : 'Copy text'}</Btn>
                </div>
                {adhocAnalysis.isUtf && <WlSmsUtfHighlight segments={adhocAnalysis.segments} />}
                <WlSmsUtfOffenderChips offenders={adhocAnalysis.offenders} />
              </div>
            )}
          </Card>

          <Card className="p-4">
            <SectionLabel hint="Checks every parsed content block from the Content tab">Content blocks</SectionLabel>
            <div className="mb-3 flex flex-wrap items-center gap-2">
              <span className="text-[11px] text-neutral-500">
                {contents.length === 0
                  ? 'No content blocks yet.'
                  : utfBlocks === 0
                    ? `All ${contents.length} content block${contents.length === 1 ? '' : 's'} are GSM-7 safe.`
                    : `${utfBlocks} of ${contents.length} content block${contents.length === 1 ? '' : 's'} will be sent as UTF.`}
              </span>
              {anyFixable && (
                <Btn variant="ghost" size="sm" onClick={fixAll}>Fix all fixable</Btn>
              )}
              {contents.some(s => s && s.trim()) && (
                <Btn variant="ghost" size="sm" onClick={copyAll}>{copied === 'all' ? 'Copied' : 'Copy all'}</Btn>
              )}
              {(contents.length > 0 || adhoc) && (
                <Btn variant="ghost" size="sm" onClick={clearAll}>Clear all</Btn>
              )}
            </div>
            <div className="space-y-2">
              {contents.length === 0 && (
                <div className="text-[12px] text-neutral-500 italic">Nothing to check. Paste content in the Content tab first.</div>
              )}
              {contents.map((block, i) => {
                const a = analyses[i];
                return (
                  <div key={i} className="rounded-md border border-neutral-900 bg-neutral-950 p-2">
                    <div className="mb-1 flex flex-wrap items-center justify-between gap-2">
                      <div className="flex items-center gap-2">
                        <span className="text-[10px] uppercase tracking-wide text-neutral-500">Content {i + 1}</span>
                        <WlSmsUtfStatusPill analysis={a} />
                      </div>
                      {a.isUtf && isFixable(block) && (
                        <button onClick={() => fixBlock(i)} className="text-[10px] text-neutral-500 hover:text-blue-300">Fix this block</button>
                      )}
                    </div>
                    <WlSmsUtfHighlight segments={a.segments} />
                    <WlSmsUtfOffenderChips offenders={a.offenders} />
                  </div>
                );
              })}
            </div>
          </Card>
        </div>
      );
    }

    function WlSmsNumbersPanel({ wl, onChange, sharedTestNumbers, onChangeSharedTestNumbers, sharedReady }) {
      const [newLabel, setNewLabel] = useState('');
      const [newNumber, setNewNumber] = useState('');
      const [newNetwork, setNewNetwork] = useState('smart');
      const numbers = Array.isArray(sharedTestNumbers) ? sharedTestNumbers : [];
      const selected = new Set(wl.selectedNumberIds || []);
      const toggle = (id) => {
        const next = new Set(selected);
        if (next.has(id)) next.delete(id); else next.add(id);
        onChange({ ...wl, selectedNumberIds: Array.from(next) });
      };
      const selectAllVisible = () => {
        const visibleIds = numbers.filter(n => (wl.networks || []).includes(n.network)).map(n => n.id);
        onChange({ ...wl, selectedNumberIds: visibleIds });
      };
      const clearSelection = () => onChange({ ...wl, selectedNumberIds: [] });
      const addNumber = () => {
        const trimmed = newNumber.trim();
        if (!trimmed) return;
        const entry = { id: wlSmsId('tn'), label: newLabel.trim() || trimmed, number: trimmed, network: newNetwork };
        // Dedupe by digits + network so two devices adding the same number
        // don't both succeed.
        const key = wlSmsNumberKey(entry);
        if (numbers.some(n => wlSmsNumberKey(n) === key)) {
          // Already exists — just auto-select the existing entry.
          const existing = numbers.find(n => wlSmsNumberKey(n) === key);
          if (existing && !selected.has(existing.id)) {
            onChange({ ...wl, selectedNumberIds: [...(wl.selectedNumberIds || []), existing.id] });
          }
        } else {
          onChangeSharedTestNumbers([...numbers, entry]);
          onChange({ ...wl, selectedNumberIds: [...(wl.selectedNumberIds || []), entry.id] });
        }
        setNewLabel(''); setNewNumber('');
      };
      const updateNumber = (id, patch) => onChangeSharedTestNumbers(numbers.map(n => n.id === id ? { ...n, ...patch } : n));
      const deleteNumber = (id) => {
        onChangeSharedTestNumbers(numbers.filter(n => n.id !== id));
        // Drop the id from this user's selection too. Other users' selections
        // will get cleaned the next time wlSmsEffectiveNumbers runs.
        onChange({ ...wl, selectedNumberIds: (wl.selectedNumberIds || []).filter(x => x !== id) });
      };
      return (
        <div className="space-y-4">
          <Card className="p-4">
            <SectionLabel hint={sharedReady ? 'Shared with all users · check to include in the generated file' : 'Loading shared database…'}>Test number database</SectionLabel>
            <div className="flex flex-wrap items-center gap-2 mb-3">
              <Btn variant="ghost" size="sm" onClick={selectAllVisible}>Select all visible</Btn>
              <Btn variant="ghost" size="sm" onClick={clearSelection}>Clear selection</Btn>
              <span className="text-[11px] text-neutral-500">{selected.size} selected / {numbers.length} total</span>
            </div>
            <div className="space-y-1.5">
              {numbers.length === 0 && (
                <div className="text-[12px] text-neutral-500 italic">No numbers yet. Add one below.</div>
              )}
              {numbers.map(n => {
                const net = WL_SMS_NETWORKS[n.network] || WL_SMS_NETWORKS.smart;
                const visible = (wl.networks || []).includes(n.network);
                return (
                  <div key={n.id} className={`flex flex-wrap items-center gap-2 rounded-md border border-neutral-900 bg-neutral-950 px-2 py-1.5 ${visible ? '' : 'opacity-40'}`}>
                    <input type="checkbox" checked={selected.has(n.id)} onChange={() => toggle(n.id)}
                      className="w-3.5 h-3.5 rounded border-neutral-700 bg-neutral-900 text-blue-500" />
                    <span className="inline-block w-2 h-2 rounded-full shrink-0" style={{ background: net.color }}></span>
                    <Input className="!w-40" value={n.label} onChange={e => updateNumber(n.id, { label: e.target.value })} placeholder="Label" />
                    <Input className="!w-44" value={n.number} onChange={e => updateNumber(n.id, { number: e.target.value })} placeholder="639XXXXXXXXX" />
                    <Select className="!w-32" value={n.network} onChange={e => updateNumber(n.id, { network: e.target.value })}>
                      {WL_SMS_NETWORK_ORDER.map(k => <option key={k} value={k}>{WL_SMS_NETWORKS[k].shortLabel}</option>)}
                    </Select>
                    <button onClick={() => deleteNumber(n.id)} className="ml-auto text-[10px] text-neutral-500 hover:text-red-300">Delete</button>
                  </div>
                );
              })}
            </div>
          </Card>

          <Card className="p-4">
            <SectionLabel>Add test number</SectionLabel>
            <div className="flex flex-wrap items-end gap-2">
              <div className="flex-1 min-w-[160px]">
                <label className="block text-[10px] uppercase tracking-wide text-neutral-500 mb-1">Label</label>
                <Input value={newLabel} onChange={e => setNewLabel(e.target.value)} placeholder="e.g. Globe Test" />
              </div>
              <div className="flex-1 min-w-[180px]">
                <label className="block text-[10px] uppercase tracking-wide text-neutral-500 mb-1">Number</label>
                <Input value={newNumber} onChange={e => setNewNumber(e.target.value)} placeholder="639XXXXXXXXX" />
              </div>
              <div className="min-w-[120px]">
                <label className="block text-[10px] uppercase tracking-wide text-neutral-500 mb-1">Network</label>
                <Select value={newNetwork} onChange={e => setNewNetwork(e.target.value)}>
                  {WL_SMS_NETWORK_ORDER.map(k => <option key={k} value={k}>{WL_SMS_NETWORKS[k].label}</option>)}
                </Select>
              </div>
              <Btn variant="accent" size="md" onClick={addNumber} disabled={!newNumber.trim()}><IconPlus /> Add</Btn>
            </div>
          </Card>
        </div>
      );
    }

    function WlSmsPreview({ wl, sharedTestNumbers }) {
      const contents = (wl.contents || []).filter(s => s && s.trim());
      const numbers = wlSmsEffectiveNumbers(wl, sharedTestNumbers);
      const rows = [];
      for (const n of numbers) {
        for (const c of contents) rows.push({ number: n.number, content: c, network: n.network });
      }
      return (
        <Card className="p-0 overflow-hidden">
          <div className="px-4 py-3 border-b border-neutral-900 flex items-center justify-between">
            <span className="text-[10px] uppercase tracking-[0.18em] text-neutral-500">Preview · same as generated file</span>
            <span className="text-[11px] text-neutral-500">{rows.length} row{rows.length === 1 ? '' : 's'}</span>
          </div>
          <div className="overflow-x-auto max-h-[60vh]">
            <table className="min-w-full text-[12px]">
              <thead className="bg-neutral-950 sticky top-0">
                <tr>
                  <th className="text-left font-semibold text-neutral-300 px-3 py-2 border-b border-neutral-900 w-40">Number</th>
                  <th className="text-left font-semibold text-neutral-300 px-3 py-2 border-b border-neutral-900">Content</th>
                </tr>
              </thead>
              <tbody>
                {rows.length === 0 && (
                  <tr><td colSpan={2} className="px-3 py-6 text-center text-neutral-500 italic">Nothing to preview. Add content and select numbers.</td></tr>
                )}
                {rows.map((r, i) => {
                  const net = WL_SMS_NETWORKS[r.network];
                  return (
                    <tr key={i} className="border-b border-neutral-900 align-top">
                      <td className="px-3 py-2 font-mono text-neutral-200">
                        <div className="flex items-center gap-2">
                          <span className="inline-block w-2 h-2 rounded-full shrink-0" style={{ background: net?.color || '#666' }}></span>
                          <span>{r.number}</span>
                        </div>
                      </td>
                      <td className="px-3 py-2 text-neutral-200 whitespace-pre-wrap">{r.content}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Card>
      );
    }

    /* ============================================================
       RCA module — Root Cause Analysis / Incident Report editor
       ============================================================ */
    const rcaId = (prefix = 'rca') => `${prefix}_${Math.random().toString(36).slice(2, 9)}`;

    // CYN logo embedded so RCAs work offline. Replaceable per-document
    // via the "Header logo" upload in the editor.
    const RCA_DEFAULT_LOGO_DATA_URL =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAH4AAACqCAYAAABrjPuxAAAgAElEQVR4nO2deXjV1ZnHP3e/yU1uEpJAVsImkVWg7EVsHUGUsojySJBFRqjt9OlMa1EghL2yuU5nOo86lg61j2hhGqtFrY5FkH3fEgibYQnZ95vc/f7mD3qO594EDJEsoO/z5Inyn99v/M73vPt73qPT/YP4jr41FAgEArrDhw8fHjhw4MC2Hsx31Ho0c+bMmfq2HsR31Db0HfDfUvoO+G8pGdt6AG1BgUAAv9+PXn9t3et0OvnzbaFvJfCapuH1etHpdBiNRvR6/bcKdPgWi3qTyYTBYMDj8bT1UNqEvpXA63Q6vF4veXl5vP/++7hcrrYeUqvTHS/qNU1D0zR0Oh2BQAC4puPPnj1LZmYmR48exePxMHXqVCwWixT5d7rOv+OBB/D7/dKgAzh+/DjLly9n586d+Hw+FixYgNPp5IknniAsLOyOBx2+JcDDNQ7W6/UcOnSIrKwsdu/ejU6nw+/3U15ezvLly3G73Tz55JNYrVbMZnNbD7lF6VsDvKZpHDt2TILu9/vRNA2DwYCmaVRVVfHCCy+gaRpz5sy544G/44w7odPFj3jt+PHjLFiwgF27duHz+aQoV3354uJiXnrpJTZs2EB9fX2j17pT6I7keK/Xi16vx+/3o9PpyMnJYdmyZezZs0cC7vP5AHC73cC1xREIBCgsLGT9+vVYLBZmzZqFyWTCaDTecXr/juN4FdhAIMDx48dZsmQJf//73yWAjVEgEJDgFxcXs2rVKjZv3ozX68Xr9UrD8E6hOw54v98vufnkyZMsWLCAzz77DE3T8Pl8Urdfj0QUr6SkhEWLFvH73/8et9uNwWBorUdoFbrjgAcwGAwcPXqUZ555hl27dhEIBPD5fGiahtHYuHYTVr9w/QwGAxUVFaxZs4bf//73VFdXt/JTtCzd9jre4/FIy1xY6UePHmXRokUcPHgQTdMkmOoCCKVAICC5XYh8gJKSEl544QV8Ph9z584lIiIiyDC8XfX+bQ+8TqfD4/FI4I4fP87ixYvZt2+ftN5V/XwjMS/AVv/W6/UUFRXx6quvYjabmTNnjozwCY/gdqTbd+T/IJ1Oh8FgQKfTcfr0abKysti7d68EXQXzZkksEmHtr1mzhk2bNuFyub7RddsD3fYcL4y1U6dOsWjRInbs2BGUazcajc2yyFXQBXdXVFSwdOlSdDod06ZNw2QyfSfqW4OE7tXr9RIYEYZdtGgRu3fvlq8Ly/6bBF6E7SBcPZ1OR3l5OcuWLcPhcDB37lxsNpu8z+2k82874D0ejzTA9Ho9R44cYdGiRezatUty9q0Sw6qkEAYiQFFRES+88AJ+v5+nnnqKyMhIuTBuF+BvKx0vKmbgGhDHjh1j4cKF7N+/Xy6Elry3GgIuLi7m5ZdfZsOGDbhcrtsurHtbAa+6WWfPnmXRokVB1ntLcpuqWvR6PYFAgNLSUl544QU2bNggS7luF2rXwIcmSQTXCUNu165dwFdgtLR7JYAVYzEYDJSXl7Nu3TrefvttnE5ngzG3VynQrnW8CLYIt8xgMJCTk0NWVhbbt2+XFr2q01tyokUwSJDX6wWu6fwlS5bg9Xp54oknCA8PB75yNdsjtWuONxgMBAIBGaDZu3cvCxYs4PPPP5cumwC/LbhLLEq4FuFbuXIlb775Jk6nM8jzaI/UrjleNdgOHTpEZmYmhw8flqlW8butJlhIAOHGlZaW8vLLL6NpGnPnziUyMrJNxtUUatccLyb25MmTLFq0iP379weJ2rbmKmFQquHd0tJSXnnlFTZs2IDD4WizsX0dtRuOF2lTETQRlJeXx+LFi4OKKIR4bw9WtJA4qqtXUlLCunXrsFqtzJo1SxZwCjulPYy73QAPBGXP9Ho9p0+fZvHixWzbtk2GTpuacGkNCgXc5/MFif1Vq1ZhNpt5/PHHCQ8PD/JO2pralag3GAyS448fP878+fP5/PPP5XttDXRTSA3vFhcXk5WVxR/+8AdZw9deqF0BL7hdFFF88cUXQRN5O6RBQzOCFRUVrFy5kjfeeKNd7dhps5lUkx9iogwGA4cPH2bRokUcOnSowS7W2yEVqiZq1KzeK6+8wmuvvUZdXV3Qc7eVFGgzHS/8c2GZi2rYJUuWcODAAfx+f1Bi5HYhYYMIQEWQp7CwkFdeeQWj0chTTz2FzWaTtkxbBHnajOPFA4utTSdPnmTJkiWyRq49GEC3gkKDPCK8K3bpthXHtxnwQszp9XpOnTpFZmamtN6F/96ejKHmUmiYt7y8nJUrV/LWW2/hdrvbzG5pM+CFEXTkyBHmz5/P9u3bgzj9dipquBGpBZxwzY4pKytjxYoVbNiwgdra2jYZV6vpeKHPARnIEEUUauXM1xk8YhJvZPQJu0FcR/2OoNB7hC4y9btikYZm59TrNnY9YcCp7wtJVlJSwvPPP4/b7W5QvdsaQZ5WA150nxCgi9j7vn37mgy6ILGAhJ0gtkypHoLZbJYBIaPRSCAQwGg0ypKs0Do8cc3rBVlUI1StABLPphqiInmkJpHEtcXrmqZRVlbGK6+8QiAQ4Kc//Sk2m63VjNlWA15dmZqTk8PixYvZv39/EEc0BXQx8SaTibCwMLp06UJSUhIdO3bEarVKQJ1OJ+Xl5Zw/f56ioiJcLhd+v18alKHXVC1rv9+P2WzGZrPRpUsXunbtSlRUFGazGZ1Oh9vtpqamhvLycnJycqitrcXv92MymeT9TSaTXGShJOoGAYqLi3n11VexWCw89dRTWK3WZs3vzVKrAS+s99OnT5OZmcnOnTsxmUxyd8v1JkklnU6HxWIhNjaWESNGcP/999OvXz9sNhsWiwWTyQRc40Cn04nX66WsrIwTJ07w97//nb1791JVVdXoAhOLwmQykZ6ezvjx4xk2bBgpKSlYrVYsFousvBGSpba2lsuXL3Py5Enef/99zp8/T319vbze9Z5BzIfQ/+Xl5axZswabzca0adOIiIho7jQ3mVoE+OtVw544cYLFixezffv2oAjX9Xa3hFJ0dDRjxozh8ccfp2fPnoSFhQENdbzf7ycsLAyLxUJERASdO3dmzJgxHDlyhM2bN7Nt2zYcDkeQSNfpdHTp0oWMjAwefPBBYmNjZQBGVSFi3Hq9ng4dOhAdHU2/fv0YO3Ys27ZtY9OmTZw7dy6oFCu0Xi80qyjAX7p0KYFAIKiYo6Wqdw1PP/3004mJiYm39KpcM+b8fj9er5dAIMDhw4fl/nTB/eK9G+l2vV6P2WymZ8+ePPvss8ycOZPU1NQGKdHQiVH1rRDlnTp14p/+6Z9ITEzk/Pnz1NTUAGCxWJg8eTJLly5l5MiR2O32Boac+FsFQtX5NpuNPn36MHLkSNxuNxcuXJDBm9B6QHURqAagw+HgwIEDWK1WevXqhdlsbpHq3ezs7OwWA15MjN/v58iRIyxcuJBDhw5J46YphpxOp8NqtdK7d29Wr17NoEGDCAsLk5MvFo1qmAENfGPV4NPpdHTr1o309HTy8vJwu93Mnj2befPmkZiYGMTl6vhUo0/lZDEG8bx2u50hQ4YAcOLEiaD3bwSeGHNdXR1Hjx7FYrFwzz33tEgfvhYDXqQo4drDP/vssxw8eLBBjPrrgNfr9XTv3p0lS5bQt29fmb1T8/JfN6ECSL1ej9frxWg0YjKZiI2NpVevXkRGRvLjH/8Yu90uDTzVlVPFsslkCpIw6gJT/zcajfTp0we3201ubm4Dy76xMaruotvt5tixY5jNZgYOHCgX7K2i7Ozs7Fui46/nw+bm5pKZmSl3rTbmT9+I4uLiePrpp+nfvz9GozEozGkwGBroX3VyVcDFQrNarbJhgtlsZtCgQfTu3RubzYbX65WuX6j/L348Hk+jolctsoBrxmV0dDRz586loKCA//u//7vhzh4xRnEtvV5PWVkZL7/8MmazmXnz5kl7Rr3nN6FbArxIqIgHEGHYJUuWsGPHjiBddiNS/fOwsDAmTJjAAw88gF6vx+VySXdNXEvYCjU1NbhcriB/3mg0EhYWRlRUlBSVwqbQ6/VyY0Z4eHhQZYwAVjxPXV0d1dXV8hnFezabjfDwcCIjI+X3xW+h4mJiYpgzZw45OTkUFRVdN+l0vcROSUkJa9aswWq1kpGRIcFXN5Y0l25JyFasdBGgOXjwIAsXLuSzzz4Dmr46BScDdOnShalTp8qHNRqNEjgBQGVlJQUFBVRWVuJwOHC5XDidTtxuN/X19fL92traIENQLBrVzlANOLgGRklJCQUFBbhcLurr6/H5fHg8HunHFxUVcfHiRbxeLx6PR4Kh2gMDBgxgypQpMojU1HlQEzvLli1j48aNuN3uW1ZcekuAF/rJYDBw8OBBnnvuOXbs2NHofvOmXMtgMDBu3DhSUlKCyqeFGDQYDJSWllJWVobX68XtduPz+SgvL6egoICqqir5mtPppKKiguLi4qDom0qqLtfpdNTV1VFcXExZWRl6vR6Px4PT6eTy5csUFBTgcDhwu91SPQhuFkEicR3xzA8//DBJSUk3xQBCCogyrueff57/+q//wuFw3JI07i3z40OtdyFeTSaT9GmbulITEhIYNWoU4eHhciJFgEXTNCorK6murkbTNC5dusTOnTvJz8+nvLycsrIyEhISpPE2ZMgQOnXqRCAQoKamRor2UPdK1eOVlZU4nU5cLhcnTpxg//79FBQUUFRUhE6no0OHDiQkJDB8+HD69u1LWFgYRUVFJCUlBS1QsbUrLi6OYcOGceHChSY9v2q7iIVeUVHBq6++ik6nY+7cucTExDQLJ0HNAl5tPyIoNzeXrKwsDh06FJRavZntymLSunTpQmpqqhTpgUBAhkOFmK2trWXbtm1s3bqVoqKiIP1eUVGBXq9n//797Nixg7Fjx3LfffdRXl6OxWLBarU2sLKF1V9XV0dtbS2FhYVs3ryZQ4cOUVdXF6Sfr169Sk5ODgcPHqRfv35kZGSQnJyMw+EgJiZGejVi8dvtdvr37897773XpJLrUDdS3Le0tJSXXnoJvV7PvHnzCA8Plwx1s4mdZgGv0wW3HxHW+549eyTowkC5WQoEAvTq1auBFStEX01NDVVVVWRnZ/PRRx9RVVXV4DPib7/fz6lTp7h69SrV1dVMnjwZl8sl256pzQ3F5x0OB5cvX2bDhg2cOHEC+CqyGLp46+vrOXjwIKWlpfzkJz/BYDBgt9uDjFSdTofP56NHjx7Y7fZmAa82eigtLWX9+vWYzWZmz54tK3lulpql44U+1+l0EnTRJvSbbHIQKzY5OTlI3Amu1zQNp9PJnj17+OCDD5rca97hcPDBBx9w6NChoIlXuT0QCOD1eiktLeXPf/4zx44dw+fzBXXBVEksbq/Xy6VLl9i4cSM1NTUyVi8+IxZjbGxsg8V8M6QaokLnv/POO3Kj5s1Ss4BXmwf+6le/Ytu2bZIrvknZlOAqu91+bXAhwRO/309+fj5/+ctfcLvduFyuJt+rurqa9957j6tXr0puFwtKbOSoq6tj9+7d7Ny5M8iHv9EzCe8gNzeXTz/9FIfDgdFoxGg0SuYwGAyYzeZmG2WNhacrKipYsmRJ0P78m6FmAW8wGNi3bx/z589n7969QRwpJqI5pOa6RTNC9TVN09i3b580sppqMAqgz5w5Q25ublDkUA2H1tfX8/e//z1IvArjqrH7iO+73W4CgQDbt2+ntLRUWvvwVVbym7hhoeMUi6miooK1a9fy+uuv43Q6JUM2JRzeLODFnnCxgdHv98ukzDepjFV9anF0iAq62+2mqKhIVvPcrIsoJIawTQQni4IJkWMXC1g8z/Wyh0Lci4rhkpISLl68GKTbVS7/JvV1oYErcf+Kigpef/11tm7dKmMYItZwI2rWSCIiIpgwYQIdO3a85cWCapxfXb16vV4WX9zsWTJqA0S73R6UuRP3EdEwNSN2s1RXV0dRURGAFPWqUXirqmuEtBMqZNiwYTKsLWoSvi6y1yzUjEYjGRkZ/PKXvyQuLq45l7ghOZ1OKR5FdgqQGaubrVIRhlt4eDj9+vWT7pwqNcxmM0lJSfTt27fZGx2E8al20gy10G8FiTFbLBbGjx9PVlYWnTt3lvmLprh2zQLe6/Vis9mYM2cOWVlZxMbGNucy16X6+nppDYvO0QKo73//+wwaNEjq36Zypk6n44EHHmD48OFAsH8sgi2RkZFMmjSJ6OjoZo9duIpqRu5WF1EImjRpEsuXL6dHjx4SdLWy50bUbFGv0+mIiIhg1qxZrFu3jqSkJHnD5op/8V2HwyHjBEKkCcMlKSmJOXPm0K1bN/kdNRPX2L1Ftcxjjz0mtyyrnxORRb1ez5AhQ/jBD37QbLBEhFH1HJqrOtTxm0wmLBYLRqNRFo6sXLmSzp07y2IVNW3dIsCrboXVamXatGmsXr2aTp06SaOsuRQIBLh06RKVlZVBbpRqxQ8ePJjnnnuOnj17yu6SqgUeSiKKJpIpamUMIPWi1+ulurq6yaVgoaTaJ2oY+JuCL2oBRRGoAD0pKQkI7hegLrgbUbMjd4JECvTRRx/F4/GwfPlyiouLm3NZed1z585RWlpKTEwMdrs9yLWCaxMxcuRIoqKiePvttzl8+DDl5eW43e7rAiZEe0FBAdHR0URFRQUtKofDQU1NDTU1NdIwa04r1NC5Ea+pwNzsohKxBpvNxkMPPcTy5cuJi4v7Rg0db1khhk6nIyMjA5fLxbp16ygoKLjp64hJKygo4OzZs8TFxWE2m6WlHZpGHTBgAN26dSMvL4+TJ09y5coVTp8+zf79+4OuK7i9qKiI+Ph4Kisr8Xq9REdHYzabKSsro6amBr/fT21tLRUVFc2azNDiDDE3qlXfrCibXo/VauXhhx9m+fLldOnSJajsrDkLtFnAq4MXLovgkjlz5hAIBFi7di3FxcVBoq6p162pqeHAgQP07t2b8PBwYmNjg/S9WuwQHh7OsGHDGDJkCF6vl7/+9a8NgBel0L/73e+YNm0a3/ve96R/bjQaqampQafTUVRUxJYtWzh8+HBzpgUgSIqIZ2qKzm3sGvBVvH/cuHEsXbqUrl27ymuKk7Kao1qbBXxjNxKFBmFhYTz11FOEhYWxYsUKioqKMJvNTWoKICx3TdPYs2cP3//+97HZbOh0OmJiYqThouYEhKsmwqSNHRsmdG9+fj5vvPEGDz/8MOPGjZM6MxAIcOrUKbZs2cKJEyduqDJuRKp+VXfR3GyxpDDUPB4PZrOZ8ePHs2LFCpKTk/H7/fJZvwndsny88LkF+NOnT8dms7FgwQIZ1Pg6EpPk8/morKzk/fffp1OnTvK9uLg4CbzgetVyFtbtja5fVlbGu+++y5dffsnUqVNJSEjgo48+YuvWrZSVlQUZlDcLfmjETy3IuBm3E76KNj7yyCOsWLGCjh07yhbs7aYQQ+gu1YgxGAxMnjwZr9dLVlYWV69ebdJEqvVnBw8eJC0tjSlTpkiujYmJkbV3oXH867lzYjLVhol79+6lsLCQrl27snfvXurr64NKs5pLqmWthlZvpuxK5NcnTJjA6tWrSUhICHr9VgSCbgnw1/OdDQYDjz32GA6Hg+eff/5rOV8tKhAbLrZu3Up8fDz3338/gUAAp9OJzWYjKipKbsESdCPjSX1dgJOfn09+fn5Q5ksdx81SaM5e5fKmAi+k1sSJE1m+fLkEXVVtt+IotBbbOyc2DYaFhfHkk08SCAR44YUXuHr1atBmCJXEpKlRL5fLxbvvvotOp+P73/8+YWFhMu9ttVqJjIyUQY3rcZa4lloc0lJdMVXx3hTf2mAwBBViWiwWJkyYQGZmJikpKUG1D4JuRSTwlgAfOhDxwMIACQsLY86cORiNRlavXi0LH6/HBarL5vf7qa6uZuPGjVy5coUHH3yQpKQkmY+vra3FZDLJ61VWVjZ53Lca9FDbIPQZr3c/wc16vZ4HH3yQBQsWkJaWFlSyfauTYS3G8SJaJnzN8PBwZs+ejdVqZcmSJU0K8qjWscPhYOvWrRw7dowxY8YwePBgWUQpODlU1LYlibGrXH+9Sh7hVk6ePJm1a9fK5xI/LXHAcYsBryZRdDqdBOfxxx/HbDYzf/78rwVfNdrE/19++SX//d//zV//+lf69+/P0KFDsdvtdOjQQaqBtiKV00O5/3qLUvjjkyZN4qWXXpLVs6p6EC7craQWBR6CLWrxMI8++igul4ulS5cGBXlCRf/1jDVN07h8+TJXr17l008/JSIigk6dOhEREUFFRUVLPVKTKLRWUJD4W910odPppE5fsWIFHTp0kJ9TQW+JdmgtAvzXGR8Wi4Xp06fj9XpZs2aNLKVqrLjiRla6KIZ0u92Ul5ffkrF/EwoV56qLq6aAjUYjbrebsLAwxo0bR1ZWFklJSVKfq4cf32rdLqjVOmKogRZN0zCbzWRkZACwZs0arl692lpDaTFSOV3l1tDGyz6fD6vVyoQJE1i0aBFpaWnyO7eqSufrqFVboaikaRomk4np06cD8Otf//qOAB+Cdb3gelXvi+CMyKcDsoy7tYzTVutzp+bMhW9qMpkIDw9n1qxZMix5O5Na4BFqt6ixiUceeYS1a9eSlpYmc/8t4bLdiFqN4xszdISfarVamTFjBhaLhYULF1JcXBxUdXM7kWq9C5e2vr6eK1euYDKZeOyxx1ixYgWxsbEySaQC3loHGbQJ8PBVQkb4+iaTiUcffRS3283KlSspLCxsNz75zZBaEm42m6mtraWsrAydTsf48eOZPXs2CQkJQW3XWgPoUGqz7tUin67qQYPBQEZGBn6/n9WrV3PlypXbDnyxEUQ0VSgpKcHr9ZKSksLw4cNlVbJa5dsW3avbDPjGYtBwzSDKyMjA6/WyevVqCgsL22iEN08CTNFnr6SkBI/HQ1JSEl27diUiIkJGMuH6vfBag9oU+FAym834fD7MZjNz5szBZDKxYsUKSkpKvlGZUWuREPFOp5OysjJcLhcdO3akR48eREREBFn44vNtIeahjQ8jaqwwUfwtDD6r1UpmZibl5eXtGnS4Nu76+nqKi4vx+XzExcXJ8rHGYvZteQpluzrkRXC1cPnMZjOPPfYYa9euJTY29rbQ94FAALfbTWxsrOyopW7SbC/UroC/XhJjypQpLFmyRJZhtWfS6/XEx8eTnp4u25KqW7LbC/jt6tw5taxIzefrdDpmzJiBx+Nh/fr1FBYWytBoW06kKrZ1umubS5KSkujWrZvcsSPeg/Z1+EK74Xi1UiU0vCny+XPmzGHx4sUkJCR8bWFla5DY1gRgtVr50Y9+JPvRqd24v64Kpy2o3QDfGKncHwgECA8P54knnmDVqlUkJyc3qdV5S5IQ3SaTicmTJ7N8+XI6d+4st5G19fhuRO1K1IeSmFhh/Qpx+vjjj2O1Wnn22WfbNLEjpNOECRNYv349HTt2DOoM0l64uzFq1xyv7vUWET6fz4fJZGLKlCmsWLGC+Pj4Vh2Tqq/FONauXUtcXFxQMWdrplibQ+2a40MTO6JiRU3per1e1q5dy5UrV+TiaAl/X9gU4tpWq5WJEyeyePFiOnXqJN3P0Oxce+X62wZ4+CokKv42m83MmDGDQCAQZO235HiEFBo/fjzz588nJSVFFluENihqz9SugQ8lwc3C7RM6f9asWZhMJpnVa6l7i2rYSZMmsWrVKlJSUoIkTHsHW6XbCvhQvxm+2mM2ffp0DAYDmZmZlJSUtMi9RR+A0Eiiunu3NYspvgndVsBDw/i2KEw0Go088cQTmM1mFixYcMPq3aaQWs6s011rv/bII4+wevVq4uLigpo0AG2WXm0u3VbANyZK1dp9Uczh8/nIysqitLQUaN4myMaqYZcsWRLE6a1RDdtSdFsB3xipxRyapmGxWCT4q1ev5vLly826rtDp4eHhTJo0Se5lg/bvqjWFbq9l2gipWS/BfSaTiWnTpvHss8/S3HOWhNcwceJEli1bRo8ePeQevdBdMrcj3fbAi1h+aIVrWFgYs2bNYuHChXTs2PFrLW41Vy6uKfrIpaWlSctd3YN/O1nxoXTbi/pQ3WowGGTnSpvNxj//8z9jt9tZuHAhhYWFmM1m2X9WkAgMATIyOHnyZJ5//nni4+PRNK1B+5FbvZentem25/hQCt2iLAy+9evXB51xc73vqeHgDh063NQhQrcT3XHAi44R6k5bwcHLli1rVOeLzwo/fdWqVXTp0gWLxSJ3uNxp4N9xwKs6WPwWEbfHH3+c5557rgH4Is7+yCOPsGzZMpKSkuQOF9E5806j21tRXYdC9b6o3hVtWXQ6HatXr5Z+vmgTumjRIlJSUmQ8Xj1y9E4D/44D/npBHtGvNiwsTO7SXb9+PWVlZUyaNEm2/lZtg9sl4dIcuuOADyURiBFBHoDIyEhmz55NXFwcn3/+OStWrMBut0tdL2r772S644GH4LCuMNSEwffQQw/J06FE9Yw4WPhO5HRB3xrgVVK3b4nfak2/aIV2J9MdD3xjAKpRPkFqZu12yrI1l+44d+47ahp9B/y3lL4D/ltK3wH/LaXvgP+W0nfAf0up3bhzt6qapak7aK/XlKE597vZ67SHGEG7AR6+OvhPrXYRhRFer5eXz58nNzeXuro6XC4XYWFhxMfH06lTJ3r37k1cXBwGgwGn00ldXR2BQEAusKioqKBjSuvr66murpa+vU6nw+fzERUVhdFolBxnsViwWCwAUjeJevyamhpqamoabNs2GAxYrVbsdnu7iA20K+DhqyNH1WoXURPv8/koLi5m586d7N69m/PnzwO1NhscHEy3bt0YOXIkAwYMICEhAY/HQ2VlpYzAiSgcQM+ePYNa+vh8PoqLi/F4PJjNZpxOJ16vF7vdjsfjkU4dgUAAm82GzWZDp9NJ4APBYJqamio8Hg9OpzMolKvT6YiKiiI6Otw94+8Y4kKxRyhrZTAYqK6u5osvvuD5559n3759JCcnk5GRwfDhw4mKisLpdHL69Gn+93//l1//+tcsXLgwoyMjg6tXr5Kfny+5T0gQs9mM2WzGZrPRsWPHIO73eDxcunQJp9OJxWKhrq4Ot9tNbGysVCcGgwGr1Up0dDQ2mw3wOK0aVlZW4nQ6cTgcQbtndDodNpvtO+CbS6FZJrUaVqfTUVNTw1//+leeffZZampq+M1vfsP06dPp1KkTRqMRr9eL2+1m6tSpzJ8/n7feeotvfvkjLk+exNixY4PEsXp+sQpcaIm4eF0FU63Z9/v9siGwOg5Aviae6evO5G0NaperLnRyVU776KOPyMzMxOPxsHTpUmbOnInJZArqYW80GklISGD8+PFs2bKFOXPmkJ6eHmRMSDXLnh2GAAAGsklEQVQVKgEa+1u9Z2NjVT8XOq72QO1uNI3p+tDDfo4cOcLChQs5fPgws2bN4ic/+QnR0dHA8WjMTQAQEhKCxWLh8U+xYsXz5895/PHHGTJkSAOdHKquQu/d1ENBVH1+I3uhrandcXxjOl3T/D9DBYqLi8nKyuLgwYMMHTqUf/3XfyU6Olr67mLSQ+vpdDodaWlpzJs3jzNnzvD8889z9erVoCKKxsamcrLqp4emXFXfH75qsxKaolXt7HZA7c64C/15p/RC6c1G/v3v//Mff/93k5+fT0pKCv/yL//8Tw9R7RC1Rh6+arQv6vKMRiODBg1i7ty5/PnPf+a3vw3z6quv0qFDB4xGI16vNyiCpzZBVDt5C9DUTpnq9wRDqAtAfFb0xRG/n376adwsLi4uaJtVaxlz7Q54iAaSqGKB4HJqaW3rdNc6Vfz1r39l27Zt6PV6Ro8eTb9+/eRkqlJC/G02m4PerlirVy/L/n4//ehHP+L48eNs376dvn37MmrUKBn5C233oRpzwjMQUkkAFcrFwoNQK21CgVeNS/H8DoeDzZs3y4WrxhpEdaztqjBDjbsLPV5SUkJlZWVQYwNN+r9XKpZRrGE36IcffsiYMWNkJ27hY9bV1ZGfn8/JkyfJyclBp9PRu3fvoOukpqYyZcoUnE4nW7du5ezZsxJsq9VKWVkZubm5HD16FK/Xi91up6CggKtXr+JyuTAajRiNRpkAOnv2LF988QXl5eW4XC727dvHkSNHKCkpkddHEzfp9Xrcbjenj/y9vPLKK7zwwgvU1NRgNps5d+4cFy9eDDpdsi2Mw3bn3oNbB7G2tpYrV67IZkbXjhEEUVJSUlJyMjLfffvthQsfFvfff58dOnRgwoQJcvKMRiPHjh3jhTfeIDc3l/T0dGw2G/feey/Hjh3D6/Vit9vlVqz09HQuXrxIYWEhJ06coLi4mKioKAYNGsTOnTu5cuUKDoeDpKQk7r//ftLT0+UpkQ6Hg82bN3PixAm6detGYWEhV69eJTk5mZqaGgYNGoTNZuPYsWMMHToUm83G1q1bWb9+PR06dKBfv36cOnWKbdu24Xa70TSNH/zgB/zwhz/EZrPJYzxay7hrl8DDV2HOyspKiouLqaurC9Lt6sQJfQ7XOkpqmsbHH3/Mhx9+SO/evdHr9bz//vt89tlnPProo3IhVVdXc/HiRRwOB1arVRZ8DBs2DKfTKQM1Op2OkydPsmnTJrxeL507d8ZkMnHixAlOnz4tGyhdvHiRyspKtm/fjsfjkc16P/30U7Zu3UpVVRX9+/cnNzeXjz/+mLi4OFwuFx9//DGapnHmzBmKi4uxWCxBhSauu+9Wt55tTYBb83vXr0eAuqQ50CGqp50e58OHDfPbZZ5SXl3PXXXf95g60n+mFAo5z9SWtq3hu1HtEdb/UPjnLly+ntLSU2NhYzpw5Q1lZGZcuXSI/P58DBw5IUSyu3atXLwoLCzl69CgGg4FAIEBKSgo2m43c3FwOHToky8d8Ph9Xr16loqICt9tNZGQkPp+P3NxcSkpKuPvuu+nQoQNXr16lqKgIvV5PUlISkZGRdO/enUOHDpGfn88PfvADvF4vBQUFlJWVybCu3+8nIyODBwoUSJyHo3ZJxorYWFwLAi6jU7Vbg0K3Bbo7vL30/wfeS6e7nDuk5W0lQfwzlfaQ7gAAAABJRU5ErkJggg==';

    const RCA_SAMPLE_SIGNATURE_DATA_URL =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAHoAAABCCAIAAAAXElVzAAACPklEQVR4nO2b25KEIAwFYWv//5fdB2opSu4hOUFIP045I7bHENHxz/M4A8WP9gDuwnRDMd1QTDcU0w3FdEMx3VBMNxTTDcV0QzHdUEw3FNMNRVm39153AGAs3VJ47/Mw/aoM5VRSv8UHCaabTq0SNp7YaOr+VuHOR/s8z+vD7qMxNd3e+3y4WiMZ2SyqjNuHQ5ja19XFJIgbUZbOeytPd3V0E3LBPgDXE9ed9whcke5i2e1uLxGIc3Q3SvC4ODnRAQXdsZLMlpT27LroSFp0YPd0v+RKFH2M6ABadzfajU6W1zVLpzGLfrpzv0WtXB26iuUIVHc41PQ2wf0fdpp6uV07JcsRQd1FccVK0lVMvgXdxHJERHdx8iE3JMUfz38/36CxjRbMumuzfENxTG78brHa5J8IdYSi8Ogev2Zr3mtnKF+kb6zDxblhW+N03ePXbDfaje/mjUqt7qfFCtlKTzEdhB2WTCPxZGzr94X4dddoQtpzXVo0GoP8iuiArG5aGd2te2NE/64yMPsU6qMo6y4GeavpgRfZu0pCzd25jVtHSne7+XOVcnFwrgO4YnLwBDiO1JpJ/prAoOWzT4ZsuscL8fFlJMCve+U27+xoO0AxMVKYXzimhfqe08P/fjdN3CXGOXVfomwFNt3k1uKqk8Sj+yplK3Cmm7bWWnux5Ejsr1BQ2HRbtEewdENh0L1Stdf3/i0s3VD0/6R9VcYt3VA0dd8Wbceim6zsNtcO8BaVkWK1G4rphmK6oZhuKH8gG1nmI3dIxgAAAABJRU5ErkJggg==';

    const RCA_BULLET_STYLES = [
      { value: '•',  label: '• Bullet' },
      { value: '◦',  label: '◦ Open circle' },
      { value: '▪',  label: '▪ Square' },
      { value: '–',  label: '– Dash' },
      { value: '*',  label: '* Asterisk' },
      { value: '1.', label: '1. Numbered' },
      { value: 'a.', label: 'a. Lower alpha' },
    ];

    // Letterhead defaults — logo and the "INCIDENT REPORT" title stay baked
    // in. The company text fields (name/address/contact/website) are the
    // initial values used when the user has not customized them yet; the
    // live values live in state.rcaCompany so they can be edited and
    // persisted without being touched by any reset-to-default action that
    // wipes the regular RCA editor state.
    const RCA_LETTERHEAD = Object.freeze({
      logo: RCA_DEFAULT_LOGO_DATA_URL,
      companyName: 'CLOUD YOUR NETWORK SOLUTIONS, INC.',
      companyAddress: '55 Sta. Rosa Street, Barrio Kapitolyo, Pasig City 1603',
      companyContact: 'PH Tel: +632 396-6399 / US Tel: +1213 550-3937 / Email: info@cynsolutions.ph',
      companyWebsite: 'www.cynsolutions.ph',
      docTitle: 'INCIDENT REPORT',
    });

    const DEFAULT_RCA_COMPANY = Object.freeze({
      companyName: RCA_LETTERHEAD.companyName,
      companyAddress: RCA_LETTERHEAD.companyAddress,
      companyContact: RCA_LETTERHEAD.companyContact,
      companyWebsite: RCA_LETTERHEAD.companyWebsite,
    });

    function rcaNormalizeCompany(raw) {
      const src = (raw && typeof raw === 'object') ? raw : {};
      const pick = (key) => {
        const v = src[key];
        return (typeof v === 'string') ? v : DEFAULT_RCA_COMPANY[key];
      };
      return {
        companyName: pick('companyName'),
        companyAddress: pick('companyAddress'),
        companyContact: pick('companyContact'),
        companyWebsite: pick('companyWebsite'),
      };
    }

    function rcaCompanyLine(value, fallback) {
      const v = typeof value === 'string' ? value.trim() : '';
      return v ? value : fallback;
    }

    // Section blueprints in render order. Each section gets its own toggle,
    // formatting state, and content textarea in the editor.
    const RCA_SECTION_DEFS = [
      { key: 'summary',     title: 'Summary of Issue',  defaultBullets: false },
      { key: 'details',     title: 'Details of Incident', defaultBullets: false },
      { key: 'resolved',    title: 'Resolved',          defaultBullets: false },
      { key: 'findings',    title: 'Findings',          defaultBullets: true  },
      { key: 'actionTaken', title: 'Action Taken',      defaultBullets: true  },
      { key: 'actionPlan',  title: 'Action Plan',       defaultBullets: true  },
    ];

    function makeRcaDefaultSections() {
      const out = {};
      RCA_SECTION_DEFS.forEach((def) => {
        out[def.key] = {
          title: def.title,
          show: true,
          content: '',
          bullets: def.defaultBullets,
          bold: false,
          italic: false,
          underline: false,
          justify: false,
          indent: !def.defaultBullets,
        };
      });
      return out;
    }

    function rcaTodayIso() {
      const d = new Date();
      const y = d.getFullYear();
      const m = pad(d.getMonth() + 1);
      const day = pad(d.getDate());
      return `${y}-${m}-${day}`;
    }

    function rcaFormatLongDate(iso) {
      if (!iso) return '';
      const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (!m) return iso;
      const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
      return d.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    }

    function rcaFormatDateTime(iso) {
      if (!iso) return '';
      const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
      if (!m) return iso;
      const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]));
      return d.toLocaleString('en-US', { year: 'numeric', month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
    }

    const DEFAULT_RCA_SIGNATORIES = [
      { id: 'sig_aries', name: 'Aries Yuson',     position: 'Senior Technical Support', signature: RCA_SAMPLE_SIGNATURE_DATA_URL },
      { id: 'sig_efren', name: 'Efren Bonavente', position: '', signature: '' },
    ];

    const DEFAULT_RCA_STATE = {
      fileName: '',
      date: rcaTodayIso(),
      customerCid: '',
      incidentStart: '',
      incidentEnd: '',
      reportedBy: '',
      sections: makeRcaDefaultSections(),
      font: 'Calibri',
      fontSize: 11,
      bulletStyle: '•',
      signatoryIds: ['sig_aries', 'sig_efren'],
      notes: '',
      // NOTE: letterhead (logo, company name, address, contact, website, doc
      // title) is intentionally NOT in state — it lives in the RCA_LETTERHEAD
      // constant above so every generated RCA carries the same branding
      // without per-report edits.
    };

    function rcaNormalizeSections(raw) {
      const base = makeRcaDefaultSections();
      if (!raw || typeof raw !== 'object') return base;
      RCA_SECTION_DEFS.forEach((def) => {
        const src = raw[def.key];
        if (!src || typeof src !== 'object') return;
        base[def.key] = {
          ...base[def.key],
          ...src,
          title: typeof src.title === 'string' && src.title.trim() ? src.title : base[def.key].title,
          show: src.show !== false,
          content: typeof src.content === 'string' ? src.content : '',
          bullets: !!src.bullets,
          bold: !!src.bold,
          italic: !!src.italic,
          underline: !!src.underline,
        };
      });
      return base;
    }

    function rcaNormalizeState(raw) {
      const merged = { ...DEFAULT_RCA_STATE, ...(raw || {}) };
      // Strip any stale letterhead fields off old persisted state so they
      // can never accidentally override the locked RCA_LETTERHEAD constant.
      delete merged.headerLogo;
      delete merged.companyName;
      delete merged.companyAddress;
      delete merged.companyContact;
      delete merged.companyWebsite;
      delete merged.docTitle;
      return {
        ...merged,
        sections: rcaNormalizeSections(merged.sections),
        signatoryIds: Array.isArray(merged.signatoryIds) ? merged.signatoryIds.filter(Boolean) : [],
        fontSize: Number(merged.fontSize) || 11,
        font: merged.font || 'Calibri',
        bulletStyle: merged.bulletStyle || '•',
        date: merged.date || '',
        customerCid: merged.customerCid || '',
        incidentStart: merged.incidentStart || '',
        incidentEnd: merged.incidentEnd || '',
        reportedBy: merged.reportedBy || '',
      };
    }

    // Template payload — captures everything except the live, per-report fields
    // (date, incident times). Lets a "CYN MAYA" template carry section
    // content, signatories, and styling without overwriting today's incident
    // metadata when loaded. Letterhead is intentionally NOT part of templates
    // — it's locked in RCA_LETTERHEAD and shared across every RCA.
    function rcaTemplateData(rca) {
      const r = rcaNormalizeState(rca);
      return {
        sections: r.sections,
        font: r.font,
        fontSize: r.fontSize,
        bulletStyle: r.bulletStyle,
        signatoryIds: r.signatoryIds,
        reportedBy: r.reportedBy,
        customerCid: r.customerCid,
      };
    }

    function rcaApplyTemplate(current, data) {
      const cur = rcaNormalizeState(current);
      const next = rcaNormalizeState({ ...cur, ...(data || {}) });
      // Preserve the user-facing per-report fields the template should NOT clobber.
      return { ...next, date: cur.date, incidentStart: cur.incidentStart, incidentEnd: cur.incidentEnd, fileName: cur.fileName };
    }

    function rcaFileNameStem(rca) {
      const r = rcaNormalizeState(rca);
      if (r.fileName && r.fileName.trim()) return r.fileName.trim();
      const cid = (r.customerCid || '').trim().replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_') || 'CLIENT';
      const date = (r.date || '').trim() || rcaTodayIso();
      return `RCA_${cid}_${date}`;
    }

    /* ---------------- RCA DOCX / PDF generation ---------------- */

    function rcaSplitLines(text) {
      return String(text || '')
        .replace(/\r\n/g, '\n')
        .split('\n')
        .map(s => s.replace(/\s+$/, ''))
        .filter(s => s.length > 0);
    }

    function rcaDataUrlToUint8(dataUrl) {
      if (!dataUrl || typeof dataUrl !== 'string') return null;
      // Tolerate data URLs with extra params (e.g. ;charset=utf-8;base64,...).
      const m = dataUrl.match(/^data:([^;,]+)(?:;[^,]*)?;base64,([\s\S]+)$/);
      if (!m) return null;
      // Strip whitespace and any stray non-base64 characters before atob so
      // a single corrupted image never crashes the export.
      let cleaned = m[2].replace(/[^A-Za-z0-9+/=]/g, '');
      // Reject `=` anywhere except the final 1-2 chars — that combo passes
      // length % 4 === 0 but makes atob throw, which is precisely the
      // "Failed to execute atob on Window" error users have hit.
      const eqIdx = cleaned.indexOf('=');
      if (eqIdx !== -1 && eqIdx < cleaned.length - 2) return null;
      // Pad to a multiple of 4 instead of dropping the image when the
      // stored data is missing its trailing `=` padding.
      const padNeeded = (4 - (cleaned.length % 4)) % 4;
      if (padNeeded === 3) return null; // not recoverable
      cleaned += '='.repeat(padNeeded);
      if (!cleaned) return null;
      let binary;
      try { binary = atob(cleaned); } catch (_) { return null; }
      if (!binary || !binary.length) return null;
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      return { bytes, mime: m[1] };
    }

    function rcaImageType(mime) {
      if (!mime) return 'png';
      const t = mime.toLowerCase();
      if (t.includes('jpeg') || t.includes('jpg')) return 'jpg';
      if (t.includes('gif')) return 'gif';
      if (t.includes('bmp')) return 'bmp';
      // docx 8.x SVG handling needs a fallback raster + the SVG string —
      // we don't ship that, and shipping bare SVG bytes makes the docx
      // library try to atob the data. Coerce to png so the image is
      // packed as raw bytes safely.
      if (t.includes('svg')) return 'png';
      return 'png';
    }

    // Build an ImageRun safely: returns null if the image data cannot be
    // expressed as a Uint8Array. This is the single chokepoint that
    // prevents docx's internal atob from ever seeing a string and
    // throwing "Failed to execute 'atob' on 'Window'".
    function rcaSafeImageRun(D, opts) {
      const bytes = opts && opts.data;
      if (!(bytes instanceof Uint8Array) || bytes.length === 0) return null;
      try {
        return new D.ImageRun({
          type: opts.type || 'png',
          data: bytes,
          transformation: opts.transformation,
          altText: opts.altText,
        });
      } catch (_) {
        return null;
      }
    }

    async function rcaImageNaturalSize(dataUrl) {
      if (!dataUrl) return null;
      return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => resolve({ width: img.naturalWidth || 1, height: img.naturalHeight || 1 });
        img.onerror = () => resolve(null);
        img.src = dataUrl;
      });
    }

    async function generateRcaDocx(state) {
      if (!window.docx) throw new Error('docx library not loaded');
      const D = window.docx;
      const rca = rcaNormalizeState(state.rca);
      const signatories = rcaResolveSignatories(state, rca.signatoryIds);
      const company = rcaNormalizeCompany(state.rcaCompany);
      const companyName    = rcaCompanyLine(company.companyName,    DEFAULT_RCA_COMPANY.companyName);
      const companyAddress = rcaCompanyLine(company.companyAddress, DEFAULT_RCA_COMPANY.companyAddress);
      const companyContact = rcaCompanyLine(company.companyContact, DEFAULT_RCA_COMPANY.companyContact);
      const companyWebsite = rcaCompanyLine(company.companyWebsite, DEFAULT_RCA_COMPANY.companyWebsite);

      const baseFont = rca.font || 'Calibri';
      const baseHalf = Math.max(8, Math.round((Number(rca.fontSize) || 11) * 2)); // half-points for docx

      const runText = (text, opts = {}) => new D.TextRun({
        text: String(text == null ? '' : text),
        font: baseFont,
        size: opts.size || baseHalf,
        bold: !!opts.bold,
        italics: !!opts.italic,
        underline: opts.underline ? { type: 'single' } : undefined,
        color: opts.color || undefined,
      });

      const para = (children, opts = {}) => new D.Paragraph({
        children: Array.isArray(children) ? children : [children],
        spacing: opts.spacing || { after: 80 },
        alignment: opts.alignment,
      });

      const children = [];

      // ---- Locked letterhead ----
      // Logo (left) and company info (right) sit on the same row via a
      // single 2-column table — matches the screenshot the user shared.
      const logoSource = rca.logo || RCA_LETTERHEAD.logo;
      const logo = rcaDataUrlToUint8(logoSource);
      const infoLines = [
        new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [runText(companyName, { bold: true, size: Math.max(18, baseHalf) })] }),
        new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [runText(companyAddress, { size: Math.max(16, baseHalf - 2) })] }),
        new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [runText(companyContact,  { size: Math.max(16, baseHalf - 2) })] }),
        new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [runText('Website: ' + companyWebsite, { size: Math.max(16, baseHalf - 2) })] }),
      ];
      const logoCellChildren = [];
      let logoRun = null;
      if (logo) {
        const size = await rcaImageNaturalSize(logoSource);
        const targetH = 96;
        const ratio = size ? (size.width / size.height) : 1;
        const targetW = Math.max(40, Math.round(targetH * ratio));
        logoRun = rcaSafeImageRun(D, {
          type: rcaImageType(logo.mime),
          data: logo.bytes,
          transformation: { width: targetW, height: targetH },
          altText: { title: 'Logo', description: companyName, name: 'logo' },
        });
      }
      if (logoRun) {
        logoCellChildren.push(new D.Paragraph({
          alignment: D.AlignmentType.LEFT,
          children: [logoRun],
        }));
      } else {
        logoCellChildren.push(new D.Paragraph({ children: [runText(' ')] }));
      }
      const noBorder = (() => {
        const b = { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
        return { top: b, bottom: b, left: b, right: b };
      })();
      children.push(new D.Table({
        width: { size: 10080, type: D.WidthType.DXA },
        columnWidths: [2400, 7680],
        rows: [new D.TableRow({
          children: [
            new D.TableCell({
              width: { size: 2400, type: D.WidthType.DXA },
              margins: { top: 0, bottom: 0, left: 0, right: 80 },
              borders: noBorder,
              verticalAlign: D.VerticalAlign.CENTER,
              children: logoCellChildren,
            }),
            new D.TableCell({
              width: { size: 7680, type: D.WidthType.DXA },
              margins: { top: 0, bottom: 0, left: 80, right: 0 },
              borders: noBorder,
              verticalAlign: D.VerticalAlign.CENTER,
              children: infoLines,
            }),
          ],
        })],
      }));

      // Title (locked)
      children.push(new D.Paragraph({
        alignment: D.AlignmentType.CENTER,
        spacing: { before: 280, after: 200 },
        children: [runText(RCA_LETTERHEAD.docTitle, { bold: true, size: Math.max(28, baseHalf + 8) })],
      }));

      // Metadata table — 2 columns
      const labelRun = (t) => runText(t, { bold: true });
      const metaRow = (l1, v1, l2, v2) => new D.TableRow({
        children: [
          new D.TableCell({
            width: { size: 5040, type: D.WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            borders: rcaCellBorders(),
            children: [new D.Paragraph({ children: [labelRun(l1 + ' '), runText(v1)] })],
          }),
          new D.TableCell({
            width: { size: 5040, type: D.WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            borders: rcaCellBorders(),
            children: [new D.Paragraph({ children: [labelRun(l2 + ' '), runText(v2)] })],
          }),
        ],
      });
      const dateStr = rcaFormatLongDate(rca.date);
      const startStr = rcaFormatDateTime(rca.incidentStart);
      const endStr = rcaFormatDateTime(rca.incidentEnd);
      const metaRows = [];
      metaRows.push(metaRow('DATE:', dateStr, 'CUSTOMER ID:', rca.customerCid || ''));
      metaRows.push(metaRow('DATE & TIME OF INCIDENT:', startStr, 'DATE & TIME OF RESOLVE:', endStr));
      metaRows.push(new D.TableRow({
        children: [
          new D.TableCell({
            width: { size: 10080, type: D.WidthType.DXA },
            columnSpan: 2,
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            borders: rcaCellBorders(),
            children: [new D.Paragraph({ children: [labelRun('REPORTED BY: '), runText(rca.reportedBy || '')] })],
          }),
        ],
      }));
      children.push(new D.Table({
        width: { size: 10080, type: D.WidthType.DXA },
        columnWidths: [5040, 5040],
        rows: metaRows,
      }));
      children.push(para([runText('')]));

      // Each enabled section: heading + content lines (paragraph or bullet).
      // Heading spacing is bumped (before: 480) so each section is separated
      // by ~2 blank lines from the previous one, per the user's spec.
      RCA_SECTION_DEFS.forEach((def) => {
        const sec = rca.sections[def.key];
        if (!sec || sec.show === false) return;
        const lines = rcaSplitLines(sec.content);
        children.push(new D.Paragraph({
          spacing: { before: 480, after: 160 },
          children: [runText(sec.title, { bold: true, size: baseHalf + 2 })],
        }));
        if (!lines.length) {
          children.push(para([runText('—', { italic: true, color: '999999' })]));
          return;
        }
        if (sec.bullets) {
          const bullet = rca.bulletStyle || '•';
          const numbered = /^\d+\./.test(bullet);
          const alpha = /^[a-z]\./i.test(bullet);
          lines.forEach((ln, idx) => {
            let marker;
            if (numbered) marker = `${idx + 1}. `;
            else if (alpha) marker = `${String.fromCharCode(97 + (idx % 26))}. `;
            else marker = `${bullet}  `;
            children.push(new D.Paragraph({
              indent: { left: 360, hanging: 240 },
              spacing: { after: 40 },
              children: [
                runText(marker, { bold: sec.bold, italic: sec.italic, underline: sec.underline }),
                runText(ln, { bold: sec.bold, italic: sec.italic, underline: sec.underline }),
              ],
            }));
          });
        } else {
          const paraAlign = sec.justify ? D.AlignmentType.JUSTIFIED : undefined;
          // 720 twips = 0.5 inch first-line indent (Word's default tab stop).
          const paraIndent = sec.indent ? { firstLine: 720 } : undefined;
          lines.forEach((ln) => {
            children.push(new D.Paragraph({
              spacing: { after: 80 },
              alignment: paraAlign,
              indent: paraIndent,
              children: [runText(ln, { bold: sec.bold, italic: sec.italic, underline: sec.underline })],
            }));
          });
        }
      });

      // Signatories
      if (signatories.length) {
        children.push(new D.Paragraph({
          spacing: { before: 360, after: 120 },
          children: [runText('Signed by:', { bold: true })],
        }));
        const sigCount = signatories.length;
        const sigCellWidth = sigCount === 1 ? 3120 : Math.floor(10080 / sigCount);
        const sigCells = await Promise.all(signatories.map(async (s) => {
          const cellChildren = [];
          const img = rcaDataUrlToUint8(s.signature);
          let sigRun = null;
          if (img) {
            const sz = await rcaImageNaturalSize(s.signature);
            const targetH = 50;
            const ratio = sz ? (sz.width / sz.height) : 2;
            const targetW = Math.max(80, Math.min(220, Math.round(targetH * ratio)));
            sigRun = rcaSafeImageRun(D, {
              type: rcaImageType(img.mime),
              data: img.bytes,
              transformation: { width: targetW, height: targetH },
              altText: { title: s.name || 'Signature', description: s.name || 'Signature', name: 'signature' },
            });
          }
          if (sigRun) {
            cellChildren.push(new D.Paragraph({
              alignment: D.AlignmentType.CENTER,
              spacing: { after: 40 },
              children: [sigRun],
            }));
          } else {
            cellChildren.push(new D.Paragraph({ spacing: { after: 320 }, children: [runText(' ')] }));
          }
          cellChildren.push(new D.Paragraph({
            alignment: D.AlignmentType.CENTER,
            border: { top: { style: D.BorderStyle.SINGLE, size: 4, color: '666666', space: 1 } },
            spacing: { before: 20, after: 40 },
            children: [runText(s.name || '', { bold: true })],
          }));
          if (s.position) cellChildren.push(new D.Paragraph({
            alignment: D.AlignmentType.CENTER,
            children: [runText(s.position, { italic: true, size: Math.max(14, baseHalf - 2) })],
          }));
          return new D.TableCell({
            width: { size: sigCellWidth, type: D.WidthType.DXA },
            margins: { top: 120, bottom: 120, left: 160, right: 160 },
            borders: rcaCellBorders('FFFFFF'),
            children: cellChildren,
          });
        }));
        const sigRowCells = [...sigCells];
        const sigColumnWidths = signatories.map(() => sigCellWidth);
        if (sigCount === 1) {
          const fillerWidth = 10080 - sigCellWidth;
          sigRowCells.push(new D.TableCell({
            width: { size: fillerWidth, type: D.WidthType.DXA },
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            borders: rcaCellBorders('FFFFFF'),
            children: [new D.Paragraph({ children: [runText(' ')] })],
          }));
          sigColumnWidths.push(fillerWidth);
        }
        children.push(new D.Table({
          width: { size: 10080, type: D.WidthType.DXA },
          columnWidths: sigColumnWidths,
          rows: [new D.TableRow({ children: sigRowCells })],
        }));
      }

      const doc = new D.Document({
        creator: 'RCA Editor',
        title: RCA_LETTERHEAD.docTitle,
        styles: { default: { document: { run: { font: baseFont, size: baseHalf } } } },
        sections: [{
          properties: {
            page: {
              size: { width: 12240, height: 15840 },
              margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
            },
          },
          children,
        }],
      });

      let blob;
      try {
        blob = await D.Packer.toBlob(doc);
      } catch (e) {
        const msg = String(e && e.message || e);
        if (/atob/i.test(msg)) {
          throw new Error("Couldn't read an embedded image (logo or signature). Re-upload it and try again.");
        }
        throw e;
      }
      const filename = rcaFileNameStem(rca) + '.docx';
      saveAs(blob, filename);
      return filename;
    }

    function rcaCellBorders(color = 'CCCCCC') {
      const D = window.docx;
      if (!D) return undefined;
      const b = { style: D.BorderStyle.SINGLE, size: 4, color };
      return { top: b, bottom: b, left: b, right: b };
    }

    function rcaResolveSignatories(state, ids) {
      const list = Array.isArray(state.rcaSignatories) ? state.rcaSignatories : DEFAULT_RCA_SIGNATORIES;
      const byId = new Map(list.map(s => [s.id, s]));
      return (Array.isArray(ids) ? ids : []).map(id => byId.get(id)).filter(Boolean);
    }

    async function generateRcaPdf(state, previewNode) {
      if (!window.html2pdf) throw new Error('html2pdf library not loaded');
      const rca = rcaNormalizeState(state.rca);
      const filename = rcaFileNameStem(rca) + '.pdf';
      const sourceNode = previewNode || document.querySelector('[data-rca-preview-root]');
      if (!sourceNode) throw new Error('Preview is not visible — open the RCA editor and try again.');

      // Clone the preview into a hidden in-flow container so the PDF
      // renderer doesn't capture any dark-mode wrappers or scrollbars.
      // html2canvas measures the source element via getBoundingClientRect
      // on its captured layout — `position: fixed; left: -10000px` returns
      // height: 0 from html2canvas's perspective and produces a blank PDF,
      // so we keep the clone in normal document flow but collapse the
      // wrapper to height: 0 with opacity: 0 to avoid any visible flash.
      const wrapper = document.createElement('div');
      wrapper.style.position = 'relative';
      wrapper.style.height = '0';
      wrapper.style.overflow = 'hidden';
      wrapper.style.opacity = '0';
      wrapper.style.pointerEvents = 'none';
      wrapper.setAttribute('aria-hidden', 'true');
      const clone = sourceNode.cloneNode(true);
      clone.style.width = '8.5in';
      clone.style.maxWidth = '8.5in';
      clone.style.padding = '0.5in';
      clone.style.background = '#ffffff';
      clone.style.color = '#111';
      wrapper.appendChild(clone);
      document.body.appendChild(wrapper);

      try {
        await window.html2pdf().set({
          filename,
          margin: 0,
          image: { type: 'jpeg', quality: 0.95 },
          html2canvas: { scale: 2, useCORS: true, backgroundColor: '#ffffff' },
          jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' },
          pagebreak: { mode: ['css', 'legacy', 'avoid-all'] },
        }).from(clone).save();
      } finally {
        document.body.removeChild(wrapper);
      }
      return filename;
    }

    /* ---------------- RCA components ---------------- */

    function RcaSidebar({ state, setState, onGenerateDocx, busy, sync, onRetrySync }) {
      const theme = state.theme || 'dark';
      const rca = rcaNormalizeState(state.rca);
      const toggleTheme = () => setState(s => ({ ...s, theme: (s.theme || 'dark') === 'dark' ? 'light' : 'dark' }));
      const update = (patch) => setState(s => ({ ...s, rca: { ...rcaNormalizeState(s.rca), ...patch } }));
      const enabledSections = RCA_SECTION_DEFS.filter(def => rca.sections[def.key]?.show !== false).length;
      const fileStem = rcaFileNameStem(rca);
      return (
        <aside className="w-[320px] shrink-0 border-r border-neutral-900 bg-[#17171a] h-screen sticky top-0 flex flex-col">
          <div className="p-5 border-b border-neutral-900">
            <div className="flex items-center justify-between gap-2">
              <SyncBadge sync={sync} onRetry={onRetrySync} />
              <button onClick={toggleTheme}
                title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}
                className="inline-flex items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2 py-1 text-[10px] uppercase tracking-wide text-neutral-400 hover:text-neutral-100 hover:border-neutral-700 transition-colors">
                {theme === 'dark' ? <IconSun /> : <IconMoon />}
                {theme === 'dark' ? 'Light' : 'Dark'}
              </button>
            </div>
            <h1 className="text-[17px] font-bold tracking-tight leading-tight mt-2">RCA Editor</h1>
            <p className="text-xs text-neutral-500 mt-1">Incident reports · DOCX</p>
            <AccountChip sync={sync} />
          </div>
          <div className="p-5 flex-1 overflow-y-auto space-y-4">
            <div>
              <SectionLabel>Output file name</SectionLabel>
              <Input value={rca.fileName} onChange={e => update({ fileName: e.target.value })} placeholder={fileStem} />
              <p className="mt-1 text-[10px] text-neutral-600 font-mono break-all">{fileStem}.docx</p>
            </div>
            <div>
              <SectionLabel>Status</SectionLabel>
              <div className="rounded-md border border-neutral-900 bg-neutral-950 p-3 space-y-2">
                <div className="flex flex-wrap gap-1.5">
                  <Pill>{enabledSections} of {RCA_SECTION_DEFS.length} sections</Pill>
                  <Pill tone={rca.signatoryIds.length ? 'accent' : 'muted'}>{rca.signatoryIds.length} signers</Pill>
                </div>
                <div className="text-[11px] text-neutral-400 truncate">{rca.customerCid || '—'}</div>
                <div className="text-[10px] text-neutral-600">{rca.date || '—'}</div>
              </div>
            </div>
          </div>
          <div className="p-5 border-t border-neutral-900 space-y-2">
            <Btn variant="primary" size="lg" onClick={onGenerateDocx} disabled={busy} className="w-full">
              {busy ? <><span className="loader"></span> Generating</> : <><IconDownload /> Export .docx</>}
            </Btn>
            <p className="text-[10px] text-neutral-600 text-center font-mono">Word output</p>
          </div>
        </aside>
      );
    }

    function RcaSectionEditor({ sectionKey, section, bulletStyle, onChange }) {
      const update = (patch) => onChange({ ...section, ...patch });
      const placeholder = section.bullets
        ? 'One bullet per line…'
        : 'Paragraph text (use blank lines for paragraph breaks)…';

      // Strip leading bullet markers from pasted text so users can paste
      // bulleted content from Word/email/Slack without doubling up with the
      // bullet style the export applies on its own.
      const handlePaste = (e) => {
        const cd = e.clipboardData || window.clipboardData;
        if (!cd) return;
        const raw = cd.getData('text/plain');
        if (!raw) return;
        const stripRe = /^[ \t ]*(?:[•·●○◦▪▫■□▶►▸➢❖\-–—*]|\d+[.)]|[a-z][.)])[ \t ]+/;
        const cleaned = raw
          .replace(/\r\n/g, '\n')
          .split('\n')
          .map(line => line.replace(stripRe, ''))
          .join('\n');
        if (cleaned === raw) return;
        e.preventDefault();
        const ta = e.target;
        ta.focus();
        let inserted = false;
        if (document.execCommand) {
          try { inserted = document.execCommand('insertText', false, cleaned); } catch (_) {}
        }
        if (!inserted) {
          const start = ta.selectionStart ?? (section.content || '').length;
          const end = ta.selectionEnd ?? start;
          const current = section.content || '';
          update({ content: current.slice(0, start) + cleaned + current.slice(end) });
        }
      };

      return (
        <Card className={`p-4 ${section.show === false ? 'opacity-50' : ''}`}>
          <div className="flex flex-wrap items-center justify-between gap-3 mb-2">
            <label className="inline-flex items-center gap-2 cursor-pointer text-sm font-semibold text-neutral-100">
              <input type="checkbox" checked={section.show !== false} onChange={e => update({ show: e.target.checked })}
                className="h-4 w-4 rounded border-neutral-700 bg-neutral-950" />
              <span>{section.title}</span>
              {section.show === false && <span className="text-[10px] text-neutral-500 font-normal">(hidden in output)</span>}
            </label>
            <div className="flex items-center gap-2 text-[11px] text-neutral-400">
              <label className="inline-flex items-center gap-1.5 cursor-pointer">
                <input type="checkbox" checked={!!section.bullets} onChange={e => update({ bullets: e.target.checked })}
                  className="h-3.5 w-3.5" />
                <span>Bullets</span>
                {section.bullets && <span className="text-neutral-500 font-mono">{bulletStyle}</span>}
              </label>
              <label className={`inline-flex h-7 items-center gap-1 rounded-md border border-neutral-800 px-2 cursor-pointer ${section.bold ? 'bg-neutral-900 text-neutral-100' : ''}`}>
                <input type="checkbox" checked={!!section.bold} onChange={e => update({ bold: e.target.checked })} className="sr-only" />
                <span className="font-bold">B</span>
              </label>
              <label className={`inline-flex h-7 items-center gap-1 rounded-md border border-neutral-800 px-2 cursor-pointer ${section.italic ? 'bg-neutral-900 text-neutral-100' : ''}`}>
                <input type="checkbox" checked={!!section.italic} onChange={e => update({ italic: e.target.checked })} className="sr-only" />
                <span className="italic">I</span>
              </label>
              <label className={`inline-flex h-7 items-center gap-1 rounded-md border border-neutral-800 px-2 cursor-pointer ${section.underline ? 'bg-neutral-900 text-neutral-100' : ''}`}>
                <input type="checkbox" checked={!!section.underline} onChange={e => update({ underline: e.target.checked })} className="sr-only" />
                <span className="underline">U</span>
              </label>
              <label className={`inline-flex h-7 items-center gap-1 rounded-md border border-neutral-800 px-2 cursor-pointer ${section.justify ? 'bg-neutral-900 text-neutral-100' : ''} ${section.bullets ? 'opacity-40 pointer-events-none' : ''}`}
                title="Justify paragraph">
                <input type="checkbox" checked={!!section.justify} onChange={e => update({ justify: e.target.checked })} className="sr-only" disabled={!!section.bullets} />
                <span className="font-semibold tracking-tight">≡</span>
              </label>
              <label className={`inline-flex h-7 items-center gap-1 rounded-md border border-neutral-800 px-2 cursor-pointer ${section.indent ? 'bg-neutral-900 text-neutral-100' : ''} ${section.bullets ? 'opacity-40 pointer-events-none' : ''}`}
                title="Indent first line">
                <input type="checkbox" checked={!!section.indent} onChange={e => update({ indent: e.target.checked })} className="sr-only" disabled={!!section.bullets} />
                <span className="font-mono text-[11px]">→¶</span>
              </label>
            </div>
          </div>
          <Textarea rows={section.bullets ? 5 : 4} value={section.content || ''}
            onChange={e => update({ content: e.target.value })} onPaste={handlePaste}
            placeholder={placeholder} disabled={section.show === false} />
        </Card>
      );
    }

    function RcaSignatoriesPanel({ state, setState, sync }) {
      const signatories = Array.isArray(state.rcaSignatories) ? state.rcaSignatories : DEFAULT_RCA_SIGNATORIES;
      // The whole app state is auto-saved to localStorage and synced to
      // Firestore (see the localStorage/Firestore sync effects in App).
      // Signatories live in state.rcaSignatories so they ride that pipeline
      // and stay in sync across devices for the signed-in user.
      const setSignatories = (next) => setState(s => ({ ...s, rcaSignatories: next }));
      const rca = rcaNormalizeState(state.rca);
      const selectedIds = new Set(rca.signatoryIds || []);
      const setSelected = (ids) => setState(s => ({ ...s, rca: { ...rcaNormalizeState(s.rca), signatoryIds: ids } }));

      const addSignatory = () => setSignatories([
        ...signatories,
        { id: rcaId('sig'), name: '', position: '', signature: '' },
      ]);
      const updateSignatory = (id, patch) =>
        setSignatories(signatories.map(s => s.id === id ? { ...s, ...patch } : s));
      const removeSignatory = async (id) => {
        const ok = await confirmDialog({
          title: 'Remove signatory?',
          message: 'This will remove them from your saved signatory list and from any RCA where they were selected.',
          confirmText: 'Remove',
          tone: 'danger',
        });
        if (!ok) return;
        setSignatories(signatories.filter(s => s.id !== id));
        if (selectedIds.has(id)) setSelected(Array.from(selectedIds).filter(x => x !== id));
      };
      const toggleSelected = (id) => {
        const next = new Set(selectedIds);
        if (next.has(id)) next.delete(id); else next.add(id);
        // Preserve order across signatories list
        const ordered = signatories.map(s => s.id).filter(x => next.has(x));
        setSelected(ordered);
      };
      const onSigUpload = (id) => (e) => {
        const file = e.target.files?.[0];
        e.target.value = '';
        if (!file) return;
        const reader = new FileReader();
        reader.onload = () => updateSignatory(id, { signature: String(reader.result || '') });
        reader.readAsDataURL(file);
      };
      const moveSig = (idx, delta) => {
        const next = signatories.slice();
        const j = idx + delta;
        if (j < 0 || j >= next.length) return;
        const [item] = next.splice(idx, 1);
        next.splice(j, 0, item);
        setSignatories(next);
      };

      return (
        <Card className="p-4">
          <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
            <SectionLabel hint="Add to your library and tick who signs this RCA">Signatories</SectionLabel>
            <Btn variant="ghost" size="sm" onClick={addSignatory}><IconPlus /> Add signatory</Btn>
          </div>
          {signatories.length === 0 ? (
            <div className="rounded-md border border-dashed border-neutral-800 px-3 py-4 text-xs text-neutral-500">
              No signatories saved yet. Click "Add signatory" to create one.
            </div>
          ) : (
            <div className="space-y-3">
              {signatories.map((s, idx) => (
                <div key={s.id} className="rounded-md border border-neutral-900 bg-neutral-950/60 p-3">
                  <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
                    <label className="inline-flex items-center gap-2 cursor-pointer text-sm">
                      <input type="checkbox" checked={selectedIds.has(s.id)} onChange={() => toggleSelected(s.id)}
                        className="h-4 w-4 rounded border-neutral-700 bg-neutral-950" />
                      <span className="font-medium text-neutral-100">{s.name || '(unnamed)'}</span>
                      {selectedIds.has(s.id) && <Pill tone="accent">in this RCA</Pill>}
                    </label>
                    <div className="flex items-center gap-1">
                      <button onClick={() => moveSig(idx, -1)} disabled={idx === 0}
                        title="Move up" className="p-1.5 text-neutral-500 hover:text-neutral-200 disabled:opacity-30"><IconUp /></button>
                      <button onClick={() => moveSig(idx, 1)} disabled={idx === signatories.length - 1}
                        title="Move down" className="p-1.5 text-neutral-500 hover:text-neutral-200 disabled:opacity-30"><IconDown /></button>
                      <button onClick={() => removeSignatory(s.id)} title="Remove signatory"
                        className="p-1.5 text-neutral-500 hover:text-red-400"><IconX /></button>
                    </div>
                  </div>
                  <div className="grid gap-2 md:grid-cols-[1fr_1fr_minmax(180px,220px)]">
                    <div>
                      <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Name</label>
                      <Input value={s.name || ''} onChange={e => updateSignatory(s.id, { name: e.target.value })} placeholder="e.g. Aries Yuson" />
                    </div>
                    <div>
                      <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Position</label>
                      <Input value={s.position || ''} onChange={e => updateSignatory(s.id, { position: e.target.value })} placeholder="e.g. Senior Technical Support" />
                    </div>
                    <div>
                      <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Signature image</label>
                      <div className="flex items-center gap-2">
                        {s.signature
                          ? <img src={s.signature} alt={s.name || 'sig'} className="h-9 max-w-[120px] object-contain rounded bg-white p-1 border border-neutral-800" />
                          : <span className="text-[11px] text-neutral-600">no image</span>}
                        <label className="inline-flex items-center gap-1 cursor-pointer rounded-md border border-neutral-800 bg-transparent px-2.5 py-1 text-xs text-neutral-200 hover:bg-neutral-900 hover:border-neutral-700">
                          <IconUpload /> Upload
                          <input type="file" accept="image/*" onChange={onSigUpload(s.id)} className="sr-only" />
                        </label>
                        {s.signature && (
                          <button onClick={() => updateSignatory(s.id, { signature: '' })}
                            className="text-[11px] text-neutral-500 hover:text-red-300" title="Clear signature">clear</button>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </Card>
      );
    }

    function RcaTemplatesPanel({ state, setState, sync }) {
      const [templates, setTemplates] = useState(() => {
        try { return JSON.parse(localStorage.getItem(RCA_TEMPLATES_KEY)) || []; } catch { return []; }
      });
      const [name, setName] = useState('');

      useEffect(() => {
        if (!window.__fb || !window.__fbm || !sync?.uid) return;
        let cancelled = false;
        (async () => {
          try {
            const { doc, getDoc } = window.__fbm;
            const ref = doc(window.__fb.db, 'users', sync.uid, 'state', 'rcaTemplates');
            const snap = await getDoc(ref);
            if (cancelled) return;
            const remote = snap.exists() ? (firestoreDesanitize(snap.data()).items || []) : null;
            if (Array.isArray(remote) && remote.length) {
              const merged = [...remote];
              for (const local of templates) {
                if (!merged.some(t => t.name === local.name)) merged.push(local);
              }
              setTemplates(merged);
              localStorage.setItem(RCA_TEMPLATES_KEY, JSON.stringify(merged));
            }
          } catch (e) { console.warn('rca templates hydrate failed', e); }
        })();
        return () => { cancelled = true; };
      }, [sync?.uid]);

      const persist = (next) => {
        setTemplates(next);
        try { localStorage.setItem(RCA_TEMPLATES_KEY, JSON.stringify(next)); } catch {}
        if (window.__fb && window.__fbm && sync?.uid) {
          const { doc, setDoc, serverTimestamp } = window.__fbm;
          const ref = doc(window.__fb.db, 'users', sync.uid, 'state', 'rcaTemplates');
          setDoc(ref, { items: firestoreSanitize(next), updatedAt: serverTimestamp() })
            .catch(e => console.warn('rca templates save failed', e));
        }
      };
      const save = () => {
        const cleanName = name.trim();
        if (!cleanName) return;
        const t = { id: rcaId('tpl'), name: cleanName, savedAt: new Date().toISOString(), data: rcaTemplateData(state.rca) };
        persist([...templates.filter(x => x.name !== t.name), t]);
        setName('');
      };
      const load = (id) => {
        const t = templates.find(x => x.id === id);
        if (!t) return;
        setState(s => ({ ...s, rca: rcaApplyTemplate(s.rca, t.data) }));
      };
      const del = (id) => persist(templates.filter(x => x.id !== id));

      return (
        <Card className="p-4">
          <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
            <SectionLabel hint="Save per-client headings, signatories, and styling (not date/incident times)">Client templates</SectionLabel>
            <Pill>{templates.length} saved</Pill>
          </div>
          <div className="grid gap-2 md:grid-cols-[minmax(220px,1fr)_auto]">
            <Input value={name} onChange={e => setName(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && (e.preventDefault(), save())}
              placeholder="e.g. CYN MAYA" />
            <Btn variant="ghost" size="md" onClick={save} disabled={!name.trim()}>Save current as template</Btn>
          </div>
          {templates.length > 0 && (
            <div className="mt-3 space-y-2">
              {templates.map(t => (
                <div key={t.id} className="flex flex-wrap items-center gap-2 rounded-md border border-neutral-900 bg-neutral-950/60 px-3 py-2">
                  <div className="min-w-0 flex-1">
                    <div className="truncate text-sm font-medium text-neutral-200">{t.name}</div>
                    <div className="text-[10px] text-neutral-600 font-mono">{new Date(t.savedAt).toLocaleDateString()}</div>
                  </div>
                  <Btn variant="ghost" size="sm" onClick={() => load(t.id)}>Use</Btn>
                  <button onClick={() => del(t.id)} title="Delete template" className="p-1.5 text-neutral-500 hover:text-red-400"><IconX /></button>
                </div>
              ))}
            </div>
          )}
        </Card>
      );
    }

    function RcaPreview({ state }) {
      const rca = rcaNormalizeState(state.rca);
      const signatories = rcaResolveSignatories(state, rca.signatoryIds);
      const company = rcaNormalizeCompany(state.rcaCompany);
      const companyName    = rcaCompanyLine(company.companyName,    DEFAULT_RCA_COMPANY.companyName);
      const companyAddress = rcaCompanyLine(company.companyAddress, DEFAULT_RCA_COMPANY.companyAddress);
      const companyContact = rcaCompanyLine(company.companyContact, DEFAULT_RCA_COMPANY.companyContact);
      const companyWebsite = rcaCompanyLine(company.companyWebsite, DEFAULT_RCA_COMPANY.companyWebsite);
      const baseFont = rca.font || 'Calibri';
      const baseSize = Number(rca.fontSize) || 11;
      return (
        <div data-rca-preview-root className="rca-preview" style={{
          background: '#ffffff',
          color: '#111',
          fontFamily: `'${baseFont}', Calibri, Arial, sans-serif`,
          fontSize: baseSize,
          lineHeight: 1.45,
          padding: '0.6in 0.7in',
          minHeight: '11in',
          boxShadow: '0 1px 4px rgba(0,0,0,0.25)',
          borderRadius: 4,
          width: '100%',
          boxSizing: 'border-box',
        }}>
          {/* Letterhead — company text locked, logo overridable via rca.logo */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 16, marginBottom: 12 }}>
            {(rca.logo || RCA_LETTERHEAD.logo) && (
              <img src={rca.logo || RCA_LETTERHEAD.logo} alt="Logo" style={{ height: 78, width: 'auto', objectFit: 'contain', flexShrink: 0 }} />
            )}
            <div style={{ flex: 1, fontSize: baseSize - 1, lineHeight: 1.35, textAlign: 'right' }}>
              <div style={{ fontWeight: 700, fontSize: baseSize + 1 }}>{companyName}</div>
              <div>{companyAddress}</div>
              <div>{companyContact}</div>
              <div>Website: {companyWebsite}</div>
            </div>
          </div>

          {/* Title (locked) */}
          <h1 style={{ textAlign: 'center', fontSize: baseSize + 9, fontWeight: 800, margin: '18px 0 14px', letterSpacing: 1 }}>
            {RCA_LETTERHEAD.docTitle}
          </h1>

          {/* Metadata grid */}
          <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: 14, fontSize: baseSize - 1 }}>
            <tbody>
              <tr>
                <td style={previewMetaCell}><strong>DATE:</strong> {rcaFormatLongDate(rca.date) || <em style={{ color: '#999' }}>—</em>}</td>
                <td style={previewMetaCell}><strong>CUSTOMER ID:</strong> {rca.customerCid || <em style={{ color: '#999' }}>—</em>}</td>
              </tr>
              <tr>
                <td style={previewMetaCell}><strong>DATE & TIME OF INCIDENT:</strong> {rcaFormatDateTime(rca.incidentStart) || <em style={{ color: '#999' }}>—</em>}</td>
                <td style={previewMetaCell}><strong>DATE & TIME OF RESOLVE:</strong> {rcaFormatDateTime(rca.incidentEnd) || <em style={{ color: '#999' }}>—</em>}</td>
              </tr>
              <tr>
                <td colSpan={2} style={previewMetaCell}><strong>REPORTED BY:</strong> {rca.reportedBy || <em style={{ color: '#999' }}>—</em>}</td>
              </tr>
            </tbody>
          </table>

          {/* Sections */}
          {RCA_SECTION_DEFS.map((def) => {
            const sec = rca.sections[def.key];
            if (!sec || sec.show === false) return null;
            const lines = rcaSplitLines(sec.content);
            const segStyle = {
              fontWeight: sec.bold ? 700 : 'inherit',
              fontStyle: sec.italic ? 'italic' : 'inherit',
              textDecoration: sec.underline ? 'underline' : 'inherit',
            };
            return (
              <div key={def.key} style={{ marginBottom: 32, breakInside: 'avoid' }}>
                <h2 style={{ fontSize: baseSize + 1, fontWeight: 700, margin: '32px 0 8px' }}>{sec.title}</h2>
                {lines.length === 0
                  ? <div style={{ color: '#aaa', fontStyle: 'italic' }}>—</div>
                  : sec.bullets
                    ? <RcaBulletList lines={lines} style={segStyle} bulletStyle={rca.bulletStyle} />
                    : lines.map((ln, i) => (
                        <p key={i} style={{
                          margin: '4px 0',
                          textAlign: sec.justify ? 'justify' : 'left',
                          textIndent: sec.indent ? '0.5in' : 0,
                          ...segStyle,
                        }}>{ln}</p>
                      ))
                }
              </div>
            );
          })}

          {/* Signatures */}
          {signatories.length > 0 && (
            <div style={{ marginTop: 36, breakInside: 'avoid' }}>
              <div style={{ fontWeight: 700, marginBottom: 12 }}>Signed by:</div>
              <div style={{
                display: 'grid',
                gridTemplateColumns: signatories.length === 1 ? '260px' : `repeat(${signatories.length}, 1fr)`,
                gap: 18,
                justifyContent: 'start',
              }}>
                {signatories.map(s => (
                  <div key={s.id} style={{ textAlign: 'center' }}>
                    <div style={{ minHeight: 56, display: 'flex', alignItems: 'flex-end', justifyContent: 'center', marginBottom: 4 }}>
                      {s.signature
                        ? <img src={s.signature} alt={s.name || 'signature'} style={{ maxHeight: 56, maxWidth: '90%', objectFit: 'contain' }} />
                        : <span style={{ fontSize: baseSize - 2, color: '#bbb' }}>(no signature image)</span>}
                    </div>
                    <div style={{ borderTop: '1px solid #666', padding: '4px 8px 0' }}>
                      <div style={{ fontWeight: 700 }}>{s.name || '—'}</div>
                      {s.position && <div style={{ fontStyle: 'italic', fontSize: baseSize - 2, color: '#444' }}>{s.position}</div>}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      );
    }

    const previewMetaCell = { border: '1px solid #d0d0d0', padding: '6px 10px', verticalAlign: 'top' };

    function RcaBulletList({ lines, style, bulletStyle }) {
      const numbered = /^\d+\./.test(bulletStyle || '');
      const alpha = /^[a-z]\./i.test(bulletStyle || '');
      if (numbered) {
        return <ol style={{ paddingLeft: 24, margin: '4px 0', ...style }}>{lines.map((ln, i) => <li key={i}>{ln}</li>)}</ol>;
      }
      if (alpha) {
        return <ol style={{ paddingLeft: 24, margin: '4px 0', listStyleType: 'lower-alpha', ...style }}>{lines.map((ln, i) => <li key={i}>{ln}</li>)}</ol>;
      }
      return (
        <ul style={{ paddingLeft: 24, margin: '4px 0', listStyleType: 'none', ...style }}>
          {lines.map((ln, i) => (
            <li key={i} style={{ position: 'relative', paddingLeft: 14 }}>
              <span style={{ position: 'absolute', left: 0 }}>{bulletStyle || '•'}</span>{ln}
            </li>
          ))}
        </ul>
      );
    }

    function RcaForm({ state, setState, sync }) {
      const rca = rcaNormalizeState(state.rca);
      const update = (patch) => setState(s => ({ ...s, rca: { ...rcaNormalizeState(s.rca), ...patch } }));
      const updateSection = (key, sec) => setState(s => {
        const cur = rcaNormalizeState(s.rca);
        return { ...s, rca: { ...cur, sections: { ...cur.sections, [key]: sec } } };
      });

      // Company info lives in state.rcaCompany — separate from state.rca so
      // it stays untouched by the per-report "Reset to default" actions
      // (e.g. logo reset). Edits persist via the same localStorage +
      // Firestore sync pipeline as signatories.
      const company = rcaNormalizeCompany(state.rcaCompany);
      const updateCompany = (patch) => setState(s => ({
        ...s,
        rcaCompany: { ...rcaNormalizeCompany(s.rcaCompany), ...patch },
      }));
      const resetCompany = () => setState(s => ({ ...s, rcaCompany: { ...DEFAULT_RCA_COMPANY } }));
      const companyIsDefault =
        company.companyName    === DEFAULT_RCA_COMPANY.companyName &&
        company.companyAddress === DEFAULT_RCA_COMPANY.companyAddress &&
        company.companyContact === DEFAULT_RCA_COMPANY.companyContact &&
        company.companyWebsite === DEFAULT_RCA_COMPANY.companyWebsite;

      const onLogoUpload = (e) => {
        const file = e.target.files?.[0];
        e.target.value = '';
        if (!file) return;
        const reader = new FileReader();
        reader.onload = () => update({ logo: String(reader.result || '') });
        reader.readAsDataURL(file);
      };
      const activeLogo = rca.logo || RCA_LETTERHEAD.logo;

      // Wipe the per-report fields (incident metadata + all section bodies +
      // file name) while keeping logo, company info, signatories, saved
      // templates, and document styling intact. Section titles and the
      // show/bullet/bold flags stay too — only the typed content is cleared.
      const clearReport = async () => {
        const ok = await confirmDialog({
          title: 'Clear report?',
          message: 'This clears the incident details and the body of every section (Summary, Details, Resolved, Findings, Action Taken, Action Plan). Your company info, logo, signatories, saved templates, and document styling are kept.',
          confirmText: 'Clear report',
          tone: 'danger',
        });
        if (!ok) return;
        setState(s => {
          const cur = rcaNormalizeState(s.rca);
          const clearedSections = {};
          Object.keys(cur.sections).forEach((key) => {
            clearedSections[key] = { ...cur.sections[key], content: '' };
          });
          return {
            ...s,
            rca: {
              ...cur,
              fileName: '',
              date: rcaTodayIso(),
              customerCid: '',
              incidentStart: '',
              incidentEnd: '',
              reportedBy: '',
              sections: clearedSections,
            },
          };
        });
      };

      return (
        <div className="space-y-4">
          {/* Company info and the "INCIDENT REPORT" title stay locked in
             RCA_LETTERHEAD. Logo is overridable here so a clean PNG/JPG can
             replace the embedded fallback when needed. */}

          <Card className="flex flex-wrap items-center justify-between gap-2 px-3 py-2">
            <div className="text-[11px] text-neutral-500">
              Start a new incident — clears report fields only, keeps your saved setup.
            </div>
            <button onClick={clearReport}
              title="Clear incident details and all section bodies"
              className="inline-flex items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2.5 py-1 text-xs text-neutral-300 hover:text-red-300 hover:border-red-500/40">
              Clear report
            </button>
          </Card>

          <Card className="p-4">
            <SectionLabel hint="PNG or JPG. Falls back to the embedded CYN logo when empty.">Header logo</SectionLabel>
            <div className="flex flex-wrap items-center gap-3">
              {activeLogo
                ? <img src={activeLogo} alt="logo" className="h-14 max-w-[140px] object-contain rounded bg-white p-1 border border-neutral-800" />
                : <span className="text-[11px] text-neutral-600">no logo</span>}
              <label className="inline-flex items-center gap-1 cursor-pointer rounded-md border border-neutral-800 bg-transparent px-2.5 py-1 text-xs text-neutral-200 hover:bg-neutral-900 hover:border-neutral-700">
                <input type="file" accept="image/*" className="sr-only" onChange={onLogoUpload} />
                Upload
              </label>
              {rca.logo && (
                <button onClick={() => update({ logo: '' })}
                  className="inline-flex items-center gap-1 rounded-md border border-neutral-800 bg-transparent px-2.5 py-1 text-xs text-neutral-400 hover:text-red-400 hover:border-red-500/40">
                  Reset to default
                </button>
              )}
            </div>
          </Card>

          <Card className="p-4">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
              <SectionLabel hint="Shown on the right side of the letterhead. Persists across resets — update here when our number or address changes.">Company info</SectionLabel>
              <button onClick={resetCompany} disabled={companyIsDefault}
                title="Reset company info to the built-in defaults"
                className="inline-flex items-center gap-1 rounded-md border border-neutral-800 bg-transparent px-2.5 py-1 text-xs text-neutral-400 hover:text-neutral-100 hover:border-neutral-700 disabled:cursor-not-allowed disabled:opacity-40">
                Reset company info
              </button>
            </div>
            <div className="grid gap-3 md:grid-cols-2">
              <div className="md:col-span-2">
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Company name</label>
                <Input value={company.companyName}
                  onChange={e => updateCompany({ companyName: e.target.value })}
                  placeholder={DEFAULT_RCA_COMPANY.companyName} />
              </div>
              <div className="md:col-span-2">
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Address</label>
                <Input value={company.companyAddress}
                  onChange={e => updateCompany({ companyAddress: e.target.value })}
                  placeholder={DEFAULT_RCA_COMPANY.companyAddress} />
              </div>
              <div className="md:col-span-2">
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Contact (phone / email)</label>
                <Input value={company.companyContact}
                  onChange={e => updateCompany({ companyContact: e.target.value })}
                  placeholder={DEFAULT_RCA_COMPANY.companyContact} />
              </div>
              <div className="md:col-span-2">
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Website</label>
                <Input value={company.companyWebsite}
                  onChange={e => updateCompany({ companyWebsite: e.target.value })}
                  placeholder={DEFAULT_RCA_COMPANY.companyWebsite} />
              </div>
            </div>
          </Card>

          <Card className="p-4">
            <SectionLabel hint="Incident metadata that shows in the report header">Incident details</SectionLabel>
            <div className="grid gap-3 md:grid-cols-2">
              <div>
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Date</label>
                <Input type="date" value={rca.date || ''} onChange={e => update({ date: e.target.value })} />
              </div>
              <div>
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Customer CID</label>
                <Input value={rca.customerCid || ''} onChange={e => update({ customerCid: e.target.value })} placeholder="e.g. CYN MAYA" />
              </div>
              <div>
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Date & time of incident</label>
                <Input type="datetime-local" value={rca.incidentStart || ''} onChange={e => update({ incidentStart: e.target.value })} />
              </div>
              <div>
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Date & time of resolve</label>
                <Input type="datetime-local" value={rca.incidentEnd || ''} onChange={e => update({ incidentEnd: e.target.value })} />
              </div>
              <div className="md:col-span-2">
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Reported by</label>
                <Input value={rca.reportedBy || ''} onChange={e => update({ reportedBy: e.target.value })} placeholder="e.g. Sir Paul of MAYA" />
              </div>
            </div>
          </Card>

          <Card className="p-4">
            <SectionLabel hint="Applied to all sections in the output">Document styling</SectionLabel>
            <div className="grid gap-3 md:grid-cols-3">
              <div>
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Font family</label>
                <Select value={rca.font || 'Calibri'} onChange={e => update({ font: e.target.value })}>
                  {FONT_FAMILIES.map(f => <option key={f} value={f}>{f}</option>)}
                </Select>
              </div>
              <div>
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Font size</label>
                <Select value={String(rca.fontSize || 11)} onChange={e => update({ fontSize: Number(e.target.value) || 11 })}>
                  {FONT_SIZES.map(s => <option key={s} value={s}>{s}</option>)}
                </Select>
              </div>
              <div>
                <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Bullet style</label>
                <Select value={rca.bulletStyle || '•'} onChange={e => update({ bulletStyle: e.target.value })}>
                  {RCA_BULLET_STYLES.map(b => <option key={b.value} value={b.value}>{b.label}</option>)}
                </Select>
              </div>
            </div>
          </Card>

          {RCA_SECTION_DEFS.map((def) => (
            <RcaSectionEditor key={def.key} sectionKey={def.key}
              section={rca.sections[def.key]}
              bulletStyle={rca.bulletStyle}
              onChange={(sec) => updateSection(def.key, sec)} />
          ))}

          <RcaSignatoriesPanel state={state} setState={setState} sync={sync} />

          <RcaTemplatesPanel state={state} setState={setState} sync={sync} />
        </div>
      );
    }

    /* ============================================================
       Image Editor module
       ------------------------------------------------------------
       A paint-style quick-annotation surface: paste a screenshot (or drop /
       pick any image file), draw arrows, circle things, blur or black out the
       private parts, then copy the result back to the clipboard or download it.

       WHERE THE PIXELS LIVE — read before adding anything here.
       The image and its annotations are held in this component's own state and
       are deliberately NOT part of `state` (the mrg_state_v1 object). That whole
       object is JSON-stringified into localStorage on every edit and mirrored to
       one Firestore document. A pasted screenshot is several MB once encoded —
       it would blow the ~5MB localStorage quota and exceed Firestore's 1MB
       document limit, which would break saving for every other module. Only the
       small tool preferences in DEFAULT_IMAGE_STATE get persisted.
       ============================================================ */

    const IMAGE_FONT_STACK = 'ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, sans-serif';
    const IMAGE_BLANK_SIZE = { w: 1280, h: 720 };
    const IMAGE_COLORS = ['#EF4444', '#F97316', '#FACC15', '#22C55E', '#3B82F6', '#A855F7', '#FFFFFF', '#000000'];

    const IMAGE_TOOLS = [
      { id: 'select',    label: 'Select',    hotkey: 'V', hint: 'Move or delete an annotation' },
      { id: 'pen',       label: 'Pen',       hotkey: 'P', hint: 'Freehand draw' },
      { id: 'arrow',     label: 'Arrow',     hotkey: 'A', hint: 'Point at something' },
      { id: 'line',      label: 'Line',      hotkey: 'L', hint: 'Straight line' },
      { id: 'rect',      label: 'Box',       hotkey: 'R', hint: 'Rectangle outline' },
      { id: 'ellipse',   label: 'Circle',    hotkey: 'C', hint: 'Circle / ellipse outline' },
      { id: 'highlight', label: 'Highlight', hotkey: 'H', hint: 'Translucent marker' },
      { id: 'text',      label: 'Text',      hotkey: 'T', hint: 'Click, then type' },
      { id: 'blur',      label: 'Blur',      hotkey: 'B', hint: 'Blur out a region' },
      { id: 'pixelate',  label: 'Pixelate',  hotkey: 'X', hint: 'Mosaic a region' },
      { id: 'redact',    label: 'Redact',    hotkey: 'D', hint: 'Solid block-out' },
      { id: 'crop',      label: 'Crop',      hotkey: 'K', hint: 'Drag a region, then apply' },
    ];
    const IMAGE_TOOL_IDS = new Set(IMAGE_TOOLS.map(t => t.id));

    // Persisted preferences ONLY — every field here is a scalar.
    const DEFAULT_IMAGE_STATE = {
      tool: 'arrow',
      color: '#EF4444',
      strokeWidth: 5,
      fontSize: 28,
      blurStrength: 14,
      pixelSize: 12,
      exportFormat: 'png',
    };

    function imageNormalizeState(raw = {}) {
      const src = raw && typeof raw === 'object' ? raw : {};
      const num = (value, fallback, min, max) => {
        const n = Number(value);
        return Number.isFinite(n) ? Math.min(max, Math.max(min, n)) : fallback;
      };
      // Rebuilt field by field rather than spread, so a stale cache can never
      // smuggle a multi-MB data URL into the synced app state.
      return {
        tool: IMAGE_TOOL_IDS.has(src.tool) ? src.tool : DEFAULT_IMAGE_STATE.tool,
        color: /^#[0-9a-f]{6}$/i.test(String(src.color || '')) ? String(src.color) : DEFAULT_IMAGE_STATE.color,
        strokeWidth: num(src.strokeWidth, DEFAULT_IMAGE_STATE.strokeWidth, 1, 40),
        fontSize: num(src.fontSize, DEFAULT_IMAGE_STATE.fontSize, 10, 160),
        blurStrength: num(src.blurStrength, DEFAULT_IMAGE_STATE.blurStrength, 2, 50),
        pixelSize: num(src.pixelSize, DEFAULT_IMAGE_STATE.pixelSize, 3, 60),
        exportFormat: src.exportFormat === 'jpeg' ? 'jpeg' : 'png',
      };
    }

    const imageId = () => `im_${Math.random().toString(36).slice(2, 9)}`;
    const imageClamp = (v, min, max) => Math.min(max, Math.max(min, v));
    const imageRectFrom = (a, b) => ({
      x: Math.min(a.x, b.x), y: Math.min(a.y, b.y),
      w: Math.abs(b.x - a.x), h: Math.abs(b.y - a.y),
    });

    // ctx.filter (used for the real gaussian blur) is missing on some older
    // browsers. Probe once; fall back to a mosaic, which redacts just as well.
    let __imageFilterSupport = null;
    function imageSupportsFilter() {
      if (__imageFilterSupport === null) {
        try {
          const ctx = document.createElement('canvas').getContext('2d');
          ctx.filter = 'blur(2px)';
          __imageFilterSupport = ctx.filter === 'blur(2px)';
        } catch { __imageFilterSupport = false; }
      }
      return __imageFilterSupport;
    }

    // A canvas backing store is measured in real screen pixels, so anything drawn
    // for display has to know the ratio. It is not fixed for the life of the page:
    // dragging the window to a monitor with different scaling changes it, and a
    // resolution media query is the only thing that reports that.
    function useDevicePixelRatio() {
      const [dpr, setDpr] = useState(() => (typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1));
      useEffect(() => {
        if (typeof window === 'undefined' || !window.matchMedia) return;
        let query = null;
        let stopped = false;
        // The query only ever fires to say "this is no longer the ratio", so each
        // firing has to re-arm a fresh query against the new value.
        const arm = () => {
          if (stopped) return;
          const current = window.devicePixelRatio || 1;
          setDpr(current);
          query = window.matchMedia(`(resolution: ${current}dppx)`);
          query.addEventListener('change', arm, { once: true });
        };
        arm();
        return () => { stopped = true; if (query) query.removeEventListener('change', arm); };
      }, []);
      return dpr;
    }

    let __imageMeasureCtx = null;
    function imageMeasureText(text, size) {
      if (!__imageMeasureCtx) __imageMeasureCtx = document.createElement('canvas').getContext('2d');
      __imageMeasureCtx.font = `600 ${size}px ${IMAGE_FONT_STACK}`;
      const lines = String(text || '').split('\n');
      let w = 0;
      for (const line of lines) w = Math.max(w, __imageMeasureCtx.measureText(line || ' ').width);
      return { w, h: lines.length * size * 1.25 };
    }

    function imageDrawPixelated(ctx, source, rect, size) {
      const w = Math.round(rect.w), h = Math.round(rect.h);
      if (w < 1 || h < 1) return;
      const block = Math.max(2, Math.round(size) || 8);
      const tw = Math.max(1, Math.round(w / block));
      const th = Math.max(1, Math.round(h / block));
      const tmp = document.createElement('canvas');
      tmp.width = tw; tmp.height = th;
      // Shrink the region down to one pixel per block, then blow it back up with
      // smoothing off — that's the mosaic.
      tmp.getContext('2d').drawImage(source, Math.round(rect.x), Math.round(rect.y), w, h, 0, 0, tw, th);
      ctx.save();
      ctx.imageSmoothingEnabled = false;
      ctx.drawImage(tmp, 0, 0, tw, th, Math.round(rect.x), Math.round(rect.y), w, h);
      ctx.restore();
    }

    function imageDrawBlurred(ctx, source, rect, strength) {
      if (rect.w < 1 || rect.h < 1) return;
      const radius = Math.max(1, Math.round(strength) || 10);
      if (!imageSupportsFilter()) {
        imageDrawPixelated(ctx, source, rect, Math.max(6, radius));
        return;
      }
      // Sample a padded region so the blur kernel has real neighbouring pixels
      // to average. Without the padding the edges of the region fade toward
      // transparent instead of blending into the surrounding image.
      const pad = radius * 2;
      const sx = Math.max(0, Math.floor(rect.x - pad));
      const sy = Math.max(0, Math.floor(rect.y - pad));
      const sw = Math.min(source.width, Math.ceil(rect.x + rect.w + pad)) - sx;
      const sh = Math.min(source.height, Math.ceil(rect.y + rect.h + pad)) - sy;
      if (sw < 1 || sh < 1) return;
      const tmp = document.createElement('canvas');
      tmp.width = sw; tmp.height = sh;
      tmp.getContext('2d').drawImage(source, sx, sy, sw, sh, 0, 0, sw, sh);
      ctx.save();
      ctx.beginPath();
      ctx.rect(rect.x, rect.y, rect.w, rect.h);
      ctx.clip();
      ctx.filter = `blur(${radius}px)`;
      ctx.drawImage(tmp, sx, sy);
      ctx.restore();
    }

    function imageDrawShape(ctx, shape, canvas) {
      const width = Math.max(1, Number(shape.width) || 1);
      const color = shape.color || '#EF4444';
      ctx.save();
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.strokeStyle = color;
      ctx.fillStyle = color;
      ctx.lineWidth = width;
      switch (shape.type) {
        case 'pen': {
          const pts = shape.points || [];
          if (!pts.length) break;
          if (pts.length === 1) {
            ctx.beginPath();
            ctx.arc(pts[0][0], pts[0][1], width / 2, 0, Math.PI * 2);
            ctx.fill();
            break;
          }
          ctx.beginPath();
          ctx.moveTo(pts[0][0], pts[0][1]);
          for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i][0], pts[i][1]);
          ctx.stroke();
          break;
        }
        case 'line':
          ctx.beginPath();
          ctx.moveTo(shape.x1, shape.y1);
          ctx.lineTo(shape.x2, shape.y2);
          ctx.stroke();
          break;
        case 'arrow': {
          const angle = Math.atan2(shape.y2 - shape.y1, shape.x2 - shape.x1);
          const head = Math.max(12, width * 3.4);
          const len = Math.hypot(shape.x2 - shape.x1, shape.y2 - shape.y1);
          // Stop the shaft short of the head so the tip stays sharp.
          const back = Math.min(head * 0.8, len);
          ctx.beginPath();
          ctx.moveTo(shape.x1, shape.y1);
          ctx.lineTo(shape.x2 - back * Math.cos(angle), shape.y2 - back * Math.sin(angle));
          ctx.stroke();
          ctx.beginPath();
          ctx.moveTo(shape.x2, shape.y2);
          ctx.lineTo(shape.x2 - head * Math.cos(angle - Math.PI / 7), shape.y2 - head * Math.sin(angle - Math.PI / 7));
          ctx.lineTo(shape.x2 - head * Math.cos(angle + Math.PI / 7), shape.y2 - head * Math.sin(angle + Math.PI / 7));
          ctx.closePath();
          ctx.fill();
          break;
        }
        case 'rect':
          ctx.strokeRect(shape.x, shape.y, shape.w, shape.h);
          break;
        case 'ellipse':
          ctx.beginPath();
          ctx.ellipse(shape.x + shape.w / 2, shape.y + shape.h / 2, Math.max(1, shape.w / 2), Math.max(1, shape.h / 2), 0, 0, Math.PI * 2);
          ctx.stroke();
          break;
        case 'highlight':
          ctx.globalAlpha = 0.35;
          ctx.fillRect(shape.x, shape.y, shape.w, shape.h);
          break;
        case 'redact':
          ctx.fillRect(shape.x, shape.y, shape.w, shape.h);
          break;
        case 'text': {
          const size = Math.max(10, Number(shape.size) || 24);
          ctx.font = `600 ${size}px ${IMAGE_FONT_STACK}`;
          ctx.textBaseline = 'top';
          // Dark halo so light-coloured text stays readable on a light screenshot.
          ctx.lineWidth = Math.max(2, size / 7);
          ctx.strokeStyle = 'rgba(0,0,0,0.6)';
          String(shape.text || '').split('\n').forEach((line, i) => {
            const y = shape.y + i * size * 1.25;
            ctx.strokeText(line, shape.x, y);
            ctx.fillText(line, shape.x, y);
          });
          break;
        }
        case 'blur':
          imageDrawBlurred(ctx, canvas, shape, shape.strength);
          break;
        case 'pixelate':
          imageDrawPixelated(ctx, canvas, shape, shape.size);
          break;
        default:
          break;
      }
      ctx.restore();
    }

    // Blur/pixelate sample the canvas itself, so shapes must be painted in order:
    // whatever sits under a blur region at draw time is what gets blurred.
    function imageRenderScene(canvas, scene, extraShape) {
      const ctx = canvas.getContext('2d');
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      if (scene.img) ctx.drawImage(scene.img, 0, 0, canvas.width, canvas.height);
      for (const shape of scene.shapes) imageDrawShape(ctx, shape, canvas);
      if (extraShape) imageDrawShape(ctx, extraShape, canvas);
    }

    function imageShapeBounds(shape) {
      switch (shape.type) {
        case 'pen': {
          const pts = shape.points || [];
          if (!pts.length) return null;
          let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
          for (const [px, py] of pts) {
            x1 = Math.min(x1, px); y1 = Math.min(y1, py);
            x2 = Math.max(x2, px); y2 = Math.max(y2, py);
          }
          return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
        }
        case 'line':
        case 'arrow':
          return imageRectFrom({ x: shape.x1, y: shape.y1 }, { x: shape.x2, y: shape.y2 });
        case 'text': {
          const size = Math.max(10, Number(shape.size) || 24);
          const m = imageMeasureText(shape.text, size);
          return { x: shape.x, y: shape.y, w: m.w, h: m.h };
        }
        default:
          return { x: shape.x, y: shape.y, w: shape.w, h: shape.h };
      }
    }

    function imageHitTest(shapes, pt) {
      for (let i = shapes.length - 1; i >= 0; i--) {
        const b = imageShapeBounds(shapes[i]);
        if (!b) continue;
        const pad = Math.max(6, (Number(shapes[i].width) || 0) / 2 + 4);
        if (pt.x >= b.x - pad && pt.x <= b.x + b.w + pad
          && pt.y >= b.y - pad && pt.y <= b.y + b.h + pad) return shapes[i].id;
      }
      return null;
    }

    function imageMoveShape(shape, dx, dy) {
      if (shape.type === 'pen') {
        return { ...shape, points: (shape.points || []).map(([px, py]) => [px + dx, py + dy]) };
      }
      if (shape.type === 'line' || shape.type === 'arrow') {
        return { ...shape, x1: shape.x1 + dx, y1: shape.y1 + dy, x2: shape.x2 + dx, y2: shape.y2 + dy };
      }
      return { ...shape, x: shape.x + dx, y: shape.y + dy };
    }

    // Drop the degenerate shapes a stray click produces (a zero-size box, a
    // zero-length arrow) instead of littering the scene with invisible entries.
    function imageShapeIsUsable(shape) {
      if (!shape) return false;
      if (shape.type === 'pen') return (shape.points || []).length > 1;
      if (shape.type === 'line' || shape.type === 'arrow') {
        return Math.hypot(shape.x2 - shape.x1, shape.y2 - shape.y1) > 4;
      }
      if (shape.type === 'text') return String(shape.text || '').trim().length > 0;
      return Math.abs(shape.w) > 4 && Math.abs(shape.h) > 4;
    }

    const IMAGE_STROKE_TOOLS = new Set(['pen', 'line', 'arrow', 'rect', 'ellipse']);

    // Which slider the toolbar shows, and which field on the shape it writes.
    function imageSizeControlFor(type) {
      if (IMAGE_STROKE_TOOLS.has(type)) return { key: 'strokeWidth', shapeKey: 'width', label: 'Stroke', min: 1, max: 40 };
      if (type === 'text') return { key: 'fontSize', shapeKey: 'size', label: 'Text size', min: 10, max: 160 };
      if (type === 'blur') return { key: 'blurStrength', shapeKey: 'strength', label: 'Blur', min: 2, max: 50 };
      if (type === 'pixelate') return { key: 'pixelSize', shapeKey: 'size', label: 'Pixel size', min: 3, max: 60 };
      return null;
    }

    function IconTool({ id }) {
      const p = { className: 'w-4 h-4', viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.9, strokeLinecap: 'round', strokeLinejoin: 'round' };
      switch (id) {
        case 'select':    return <svg {...p}><path d="M4 3l7.5 17 2.3-6.9 6.9-2.3z" /></svg>;
        case 'pen':       return <svg {...p}><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4z" /></svg>;
        case 'arrow':     return <svg {...p}><line x1="5" y1="19" x2="19" y2="5" /><polyline points="10 5 19 5 19 14" /></svg>;
        case 'line':      return <svg {...p}><line x1="5" y1="19" x2="19" y2="5" /></svg>;
        case 'rect':      return <svg {...p}><rect x="3.5" y="5.5" width="17" height="13" rx="1.5" /></svg>;
        case 'ellipse':   return <svg {...p}><circle cx="12" cy="12" r="8.5" /></svg>;
        case 'highlight': return <svg {...p}><path d="M4 15l5-5 5 5-5 5H4z" /><path d="M11 8l5-5 5 5-5 5" /></svg>;
        case 'text':      return <svg {...p}><polyline points="5 6 5 4 19 4 19 6" /><line x1="12" y1="4" x2="12" y2="20" /><line x1="9" y1="20" x2="15" y2="20" /></svg>;
        case 'blur':      return <svg {...p}><path d="M12 3s6 6.2 6 10a6 6 0 0 1-12 0c0-3.8 6-10 6-10z" /><path d="M9.5 14a2.5 2.5 0 0 0 2.5 2.5" /></svg>;
        case 'pixelate':  return <svg {...p}><rect x="4" y="4" width="6" height="6" /><rect x="14" y="4" width="6" height="6" /><rect x="4" y="14" width="6" height="6" /><rect x="14" y="14" width="6" height="6" /></svg>;
        case 'redact':    return <svg {...p} fill="currentColor" stroke="none"><rect x="3.5" y="7" width="17" height="10" rx="1.5" /></svg>;
        case 'crop':      return <svg {...p}><path d="M6.5 2v15.5H22" /><path d="M2 6.5h15.5V22" /></svg>;
        default:          return null;
      }
    }
    function IconCopy() { return <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="9" y="9" width="12" height="12" rx="2" /><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" /></svg>; }
    function IconRedo() { return <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.3" strokeLinecap="round" strokeLinejoin="round"><path d="m15 14 5-5-5-5" /><path d="M20 9H10a6 6 0 0 0 0 12h1" /></svg>; }
    function IconImage() { return <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.9" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" /><circle cx="8.5" cy="8.5" r="1.6" /><path d="m21 15-5-5L5 21" /></svg>; }

    function ImageSidebar({ state, setState, sync, onRetrySync, scene, onPickFile, onBlank, onCopy, onDownload, busy }) {
      const theme = state.theme || 'dark';
      const prefs = imageNormalizeState(state.imageEditor);
      const fileRef = useRef(null);
      const toggleTheme = () => setState(s => ({ ...s, theme: (s.theme || 'dark') === 'dark' ? 'light' : 'dark' }));
      const hasImage = !!scene.img;
      return (
        <aside className="w-[320px] shrink-0 border-r border-neutral-900 bg-[#17171a] h-screen sticky top-0 flex flex-col">
          <div className="p-5 border-b border-neutral-900">
            <div className="flex items-center justify-between gap-2">
              <SyncBadge sync={sync} onRetry={onRetrySync} />
              <button onClick={toggleTheme}
                title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}
                className="inline-flex items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2 py-1 text-[10px] uppercase tracking-wide text-neutral-400 hover:text-neutral-100 hover:border-neutral-700 transition-colors">
                {theme === 'dark' ? <IconSun /> : <IconMoon />}
                {theme === 'dark' ? 'Light' : 'Dark'}
              </button>
            </div>
            <h1 className="text-[17px] font-bold tracking-tight leading-tight mt-2">Image Editor</h1>
            <p className="text-xs text-neutral-500 mt-1">Paste a screenshot, mark it up, copy it back</p>
            <AccountChip sync={sync} />
          </div>

          <div className="p-5 flex-1 overflow-y-auto min-h-0 space-y-5">
            <div>
              <SectionLabel hint="Any image type">Source</SectionLabel>
              <div className="rounded-md border border-neutral-900 bg-neutral-950 p-3 space-y-2">
                <div className="flex items-center gap-2 text-xs text-neutral-400">
                  <kbd className="rounded border border-neutral-700 bg-neutral-900 px-1.5 py-0.5 font-mono text-[10px] text-neutral-300">Ctrl</kbd>
                  <span className="text-neutral-600">+</span>
                  <kbd className="rounded border border-neutral-700 bg-neutral-900 px-1.5 py-0.5 font-mono text-[10px] text-neutral-300">V</kbd>
                  <span>anywhere to paste</span>
                </div>
                <input ref={fileRef} type="file" accept="image/*" className="hidden"
                  onChange={e => { const f = e.target.files?.[0]; if (f) onPickFile(f); e.target.value = ''; }} />
                <Btn variant="ghost" size="md" className="w-full" onClick={() => fileRef.current?.click()}>
                  <IconUpload /> Upload image
                </Btn>
                <Btn variant="ghost" size="md" className="w-full" onClick={onBlank}>
                  <IconImage /> Blank canvas
                </Btn>
              </div>
            </div>

            {hasImage && (
              <div>
                <SectionLabel>Image</SectionLabel>
                <div className="rounded-md border border-neutral-900 bg-neutral-950 p-3 space-y-2">
                  <div className="truncate text-sm font-medium text-neutral-100" title={scene.name}>{scene.name || 'Untitled'}</div>
                  <div className="flex flex-wrap gap-1.5">
                    <Pill>{scene.width} × {scene.height}</Pill>
                    <Pill tone={scene.shapes.length ? 'accent' : 'default'}>{scene.shapes.length} edit{scene.shapes.length === 1 ? '' : 's'}</Pill>
                  </div>
                </div>
              </div>
            )}

            <div>
              <SectionLabel hint="Not saved to the cloud">Scratch space</SectionLabel>
              <p className="text-[11px] leading-relaxed text-neutral-600">
                Images stay in this browser tab only — they are never uploaded or synced.
                Copy or download before you leave the module.
              </p>
            </div>
          </div>

          <div className="p-5 border-t border-neutral-900 space-y-3">
            <div>
              <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">Format</label>
              <Select value={prefs.exportFormat}
                onChange={e => setState(s => ({ ...s, imageEditor: imageNormalizeState({ ...imageNormalizeState(s.imageEditor), exportFormat: e.target.value }) }))}>
                <option value="png">PNG · keeps transparency</option>
                <option value="jpeg">JPG · smaller file</option>
              </Select>
            </div>
            <Btn variant="ghost" size="lg" onClick={onCopy} disabled={!hasImage || busy} className="w-full">
              <IconCopy /> Copy image
            </Btn>
            <Btn variant="primary" size="lg" onClick={onDownload} disabled={!hasImage || busy} className="w-full">
              {busy ? <><span className="loader"></span> Saving</> : <><IconDownload /> Download</>}
            </Btn>
            <p className="text-[10px] text-neutral-600 text-center font-mono">
              {hasImage ? `image_${bmrTodayString()}.${prefs.exportFormat === 'jpeg' ? 'jpg' : 'png'}` : 'no image loaded'}
            </p>
          </div>
        </aside>
      );
    }

    function ImageEditorModule({ state, setState, sync, onRetrySync, moduleSwitch, onToast }) {
      const prefs = imageNormalizeState(state.imageEditor);
      const setPrefs = useCallback((patch) => setState(s => ({
        ...s,
        imageEditor: imageNormalizeState({ ...imageNormalizeState(s.imageEditor), ...patch }),
      })), [setState]);

      // Session-only — see the module note above. `img` is any CanvasImageSource
      // (an <img> for a pasted/loaded file, a <canvas> after a crop or a blank
      // start), so drawImage handles both without a branch.
      const [scene, setSceneState] = useState({ img: null, name: '', width: 0, height: 0, shapes: [] });
      const sceneRef = useRef(scene);
      const setScene = useCallback((next) => {
        const value = typeof next === 'function' ? next(sceneRef.current) : next;
        sceneRef.current = value;
        setSceneState(value);
      }, []);

      // Undo/redo stacks live in refs, not state. If they were state, undo() would
      // read whichever `history` its render closure captured — so two undos that
      // React batches into one render would both pop the same entry and the second
      // would silently do nothing. Refs are always current, so the stacks stay
      // correct no matter how the updates get batched. The lengths are mirrored
      // into state purely to drive the buttons' disabled state.
      const historyRef = useRef([]);
      const futureRef = useRef([]);
      const [historyLen, setHistoryLen] = useState(0);
      const [futureLen, setFutureLen] = useState(0);
      const syncStackLengths = useCallback(() => {
        setHistoryLen(historyRef.current.length);
        setFutureLen(futureRef.current.length);
      }, []);

      const [draft, setDraftState] = useState(null);
      const draftRef = useRef(null);
      const setDraft = useCallback((value) => { draftRef.current = value; setDraftState(value); }, []);
      const [textDraft, setTextDraftState] = useState(null);
      const textDraftRef = useRef(null);
      const setTextDraft = useCallback((value) => { textDraftRef.current = value; setTextDraftState(value); }, []);
      const [selectedId, setSelectedId] = useState(null);
      const [cropRect, setCropRect] = useState(null);
      const [zoom, setZoom] = useState('fit');
      const [dragOver, setDragOver] = useState(false);
      const [busy, setBusy] = useState(false);
      const [stageWidth, setStageWidth] = useState(0);

      const dpr = useDevicePixelRatio();

      const canvasRef = useRef(null);
      // The scene is always rendered at the image's own resolution, never the
      // display's. Blur and pixelate sample this canvas back to build their
      // effect, and export encodes it, so one canvas pixel has to stay one image
      // pixel. The visible canvas below is a separate, device-pixel-sized copy.
      const sceneCanvasRef = useRef(null);
      if (!sceneCanvasRef.current && typeof document !== 'undefined') {
        sceneCanvasRef.current = document.createElement('canvas');
      }
      const stageRef = useRef(null);
      const dragRef = useRef(null);
      const objectUrlRef = useRef(null);
      const textInputRef = useRef(null);
      const editSessionRef = useRef(null);

      const hasImage = !!scene.img;
      const selected = scene.shapes.find(s => s.id === selectedId) || null;

      const pushHistory = useCallback((snapshot) => {
        historyRef.current = [...historyRef.current.slice(-49), snapshot];
        futureRef.current = [];
        syncStackLengths();
      }, [syncStackLengths]);

      const commit = useCallback((next) => {
        const prev = sceneRef.current;
        const value = typeof next === 'function' ? next(prev) : next;
        if (value === prev) return;
        pushHistory(prev);
        setScene(value);
      }, [pushHistory, setScene]);

      /* ---------- loading ---------- */

      const resetEditingState = useCallback(() => {
        historyRef.current = [];
        futureRef.current = [];
        syncStackLengths();
        setDraft(null); setTextDraft(null);
        setSelectedId(null); setCropRect(null);
        setZoom('fit');
        dragRef.current = null;
        editSessionRef.current = null;
      }, [setDraft, setTextDraft, syncStackLengths]);

      const loadFile = useCallback((file) => {
        if (!file) return;
        if (!String(file.type || '').startsWith('image/')) {
          onToast('err', 'That file is not an image');
          return;
        }
        const url = URL.createObjectURL(file);
        const img = new Image();
        img.onload = () => {
          if (objectUrlRef.current) URL.revokeObjectURL(objectUrlRef.current);
          objectUrlRef.current = url;
          resetEditingState();
          setScene({
            img,
            name: file.name || 'pasted-image',
            width: img.naturalWidth,
            height: img.naturalHeight,
            shapes: [],
          });
          onToast('ok', `Loaded ${img.naturalWidth} × ${img.naturalHeight} image`);
        };
        img.onerror = () => {
          URL.revokeObjectURL(url);
          onToast('err', 'Could not read that image');
        };
        img.src = url;
      }, [onToast, resetEditingState, setScene]);

      const loadBlank = useCallback(() => {
        const canvas = document.createElement('canvas');
        canvas.width = IMAGE_BLANK_SIZE.w;
        canvas.height = IMAGE_BLANK_SIZE.h;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        resetEditingState();
        setScene({ img: canvas, name: 'blank-canvas', width: canvas.width, height: canvas.height, shapes: [] });
      }, [resetEditingState, setScene]);

      useEffect(() => () => {
        if (objectUrlRef.current) URL.revokeObjectURL(objectUrlRef.current);
      }, []);

      // Paste anywhere in the module. Only mounted while the Image Editor is the
      // active module, so it can't steal pastes from the other modules' inputs.
      useEffect(() => {
        const onPaste = (e) => {
          const items = e.clipboardData?.items || [];
          for (const item of items) {
            if (item.kind === 'file' && String(item.type || '').startsWith('image/')) {
              const file = item.getAsFile();
              if (file) { e.preventDefault(); loadFile(file); return; }
            }
          }
        };
        window.addEventListener('paste', onPaste);
        return () => window.removeEventListener('paste', onPaste);
      }, [loadFile]);

      /* ---------- render ---------- */

      useEffect(() => {
        const stage = stageRef.current;
        if (!stage || typeof ResizeObserver === 'undefined') return;
        const ro = new ResizeObserver(entries => {
          for (const entry of entries) setStageWidth(entry.contentRect.width);
        });
        ro.observe(stage);
        setStageWidth(stage.clientWidth);
        return () => ro.disconnect();
      }, [hasImage]);

      // Zoom counts device pixels per image pixel, so 1:1 puts one image pixel on
      // one screen pixel and a pasted screenshot looks exactly as sharp as it did
      // on screen. `scale` converts that to the CSS pixels layout needs: on a
      // 150%-scaled display, 1:1 is a CSS scale of 0.667. Fit never goes above 1:1
      // because magnifying past it only invents pixels the image does not have.
      const fitZoom = (scene.width && stageWidth)
        ? Math.min(1, ((stageWidth - 8) * dpr) / scene.width)
        : 1;
      const zoomLevel = zoom === 'fit' ? fitZoom : zoom;
      const scale = zoomLevel / dpr;
      // Deliberately not rounded to whole CSS pixels. Rounding here and then
      // multiplying back up by the ratio lands a fraction off the device grid —
      // at 1:1 on a 150% display that alone resamples the image by half a pixel.
      const displayW = Math.max(1, scene.width * scale);
      const displayH = Math.max(1, scene.height * scale);

      // Paint the scene at the image's own resolution. Zoom is deliberately not a
      // dependency: this is the expensive pass (a blur region re-filters on every
      // run) and the result does not depend on how large it is being shown.
      useEffect(() => {
        const offscreen = sceneCanvasRef.current;
        if (!offscreen || !scene.img) return;
        if (offscreen.width !== scene.width) offscreen.width = scene.width;
        if (offscreen.height !== scene.height) offscreen.height = scene.height;
        imageRenderScene(offscreen, scene, draft);
      }, [scene, draft]);

      // Blit that onto a canvas whose backing store is the physical pixels the CSS
      // box covers. Sizing the visible canvas in image pixels instead left the
      // browser stretching it onto the device grid, which softened every image the
      // stage was not already shrinking. Runs after the pass above, same commit.
      useEffect(() => {
        const offscreen = sceneCanvasRef.current;
        const canvas = canvasRef.current;
        if (!offscreen || !canvas || !scene.img) return;
        // Derived from the image, not from displayW, so at 1:1 this is exactly the
        // image's own pixel count and the draw below is a straight copy.
        const bw = Math.max(1, Math.round(scene.width * zoomLevel));
        const bh = Math.max(1, Math.round(scene.height * zoomLevel));
        if (canvas.width !== bw) canvas.width = bw;
        if (canvas.height !== bh) canvas.height = bh;
        const ctx = canvas.getContext('2d');
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        ctx.clearRect(0, 0, bw, bh);
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        ctx.drawImage(offscreen, 0, 0, offscreen.width, offscreen.height, 0, 0, bw, bh);
      }, [scene, draft, zoomLevel]);

      /* ---------- pointer ---------- */

      const toImagePoint = useCallback((e) => {
        const canvas = canvasRef.current;
        const rect = canvas.getBoundingClientRect();
        const sx = rect.width ? scene.width / rect.width : 1;
        const sy = rect.height ? scene.height / rect.height : 1;
        return {
          x: imageClamp((e.clientX - rect.left) * sx, 0, scene.width),
          y: imageClamp((e.clientY - rect.top) * sy, 0, scene.height),
        };
      }, [scene.width, scene.height]);

      const makeShape = useCallback((tool, a, b) => {
        const base = { id: imageId(), type: tool, color: prefs.color, width: prefs.strokeWidth };
        switch (tool) {
          case 'pen':      return { ...base, points: [[a.x, a.y]] };
          case 'line':
          case 'arrow':    return { ...base, x1: a.x, y1: a.y, x2: b.x, y2: b.y };
          case 'blur':     return { ...base, ...imageRectFrom(a, b), strength: prefs.blurStrength };
          case 'pixelate': return { ...base, ...imageRectFrom(a, b), size: prefs.pixelSize };
          default:         return { ...base, ...imageRectFrom(a, b) };
        }
      }, [prefs.color, prefs.strokeWidth, prefs.blurStrength, prefs.pixelSize]);

      const commitText = useCallback(() => {
        const t = textDraftRef.current;
        if (!t) return;
        setTextDraft(null);
        if (!String(t.value || '').trim()) return;
        commit(prev => ({
          ...prev,
          shapes: [...prev.shapes, {
            id: imageId(), type: 'text', x: t.x, y: t.y,
            text: t.value, color: prefs.color, size: prefs.fontSize,
          }],
        }));
      }, [commit, prefs.color, prefs.fontSize, setTextDraft]);

      // Keeps the drag alive when the pointer leaves the canvas mid-stroke.
      // It can throw if the pointer is already gone — never let that abort the
      // rest of the handler, or the drag silently never starts.
      const capturePointer = (e) => {
        try { e.currentTarget.setPointerCapture(e.pointerId); } catch {}
      };

      const onPointerDown = (e) => {
        if (!hasImage || e.button !== 0) return;
        if (textDraftRef.current) { commitText(); return; }
        const pt = toImagePoint(e);
        const tool = prefs.tool;

        if (tool === 'select') {
          const id = imageHitTest(scene.shapes, pt);
          setSelectedId(id);
          if (id) {
            capturePointer(e);
            dragRef.current = { mode: 'move', id, last: pt, before: sceneRef.current, moved: false };
          }
          return;
        }

        setSelectedId(null);
        capturePointer(e);

        if (tool === 'text') {
          setTextDraft({ x: pt.x, y: pt.y, value: '' });
          return;
        }
        if (tool === 'crop') {
          dragRef.current = { mode: 'crop', start: pt };
          setCropRect({ x: pt.x, y: pt.y, w: 0, h: 0 });
          return;
        }
        dragRef.current = { mode: 'draw', tool, start: pt };
        setDraft(makeShape(tool, pt, pt));
      };

      const onPointerMove = (e) => {
        const drag = dragRef.current;
        if (!drag || !hasImage) return;
        const pt = toImagePoint(e);

        if (drag.mode === 'move') {
          const dx = pt.x - drag.last.x;
          const dy = pt.y - drag.last.y;
          if (!dx && !dy) return;
          drag.last = pt;
          drag.moved = true;
          // Live-move without touching history — one undo entry is pushed for
          // the whole drag when the pointer comes up.
          setScene(prev => ({
            ...prev,
            shapes: prev.shapes.map(s => s.id === drag.id ? imageMoveShape(s, dx, dy) : s),
          }));
          return;
        }
        if (drag.mode === 'crop') {
          setCropRect(imageRectFrom(drag.start, pt));
          return;
        }
        if (drag.mode === 'draw') {
          const current = draftRef.current;
          if (!current) return;
          if (drag.tool === 'pen') {
            setDraft({ ...current, points: [...current.points, [pt.x, pt.y]] });
          } else if (drag.tool === 'line' || drag.tool === 'arrow') {
            setDraft({ ...current, x2: pt.x, y2: pt.y });
          } else {
            setDraft({ ...current, ...imageRectFrom(drag.start, pt) });
          }
        }
      };

      const onPointerUp = () => {
        const drag = dragRef.current;
        dragRef.current = null;
        if (!drag) return;

        if (drag.mode === 'move') {
          if (drag.moved && drag.before) pushHistory(drag.before);
          return;
        }
        if (drag.mode === 'crop') return; // the rect stays pending until Apply
        if (drag.mode === 'draw') {
          const shape = draftRef.current;
          setDraft(null);
          if (!imageShapeIsUsable(shape)) return;
          commit(prev => ({ ...prev, shapes: [...prev.shapes, shape] }));
        }
      };

      /* ---------- actions ---------- */

      const undo = useCallback(() => {
        const stack = historyRef.current;
        if (!stack.length) return;
        historyRef.current = stack.slice(0, -1);
        futureRef.current = [sceneRef.current, ...futureRef.current].slice(0, 50);
        syncStackLengths();
        setScene(stack[stack.length - 1]);
        setSelectedId(null); setDraft(null); setTextDraft(null); setCropRect(null);
      }, [setScene, setDraft, setTextDraft, syncStackLengths]);

      const redo = useCallback(() => {
        const stack = futureRef.current;
        if (!stack.length) return;
        futureRef.current = stack.slice(1);
        historyRef.current = [...historyRef.current.slice(-49), sceneRef.current];
        syncStackLengths();
        setScene(stack[0]);
        setSelectedId(null); setDraft(null); setTextDraft(null); setCropRect(null);
      }, [setScene, setDraft, setTextDraft, syncStackLengths]);

      const deleteSelected = useCallback(() => {
        if (!selectedId) return;
        commit(prev => ({ ...prev, shapes: prev.shapes.filter(s => s.id !== selectedId) }));
        setSelectedId(null);
      }, [commit, selectedId]);

      const clearShapes = useCallback(() => {
        if (!sceneRef.current.shapes.length) return;
        commit(prev => ({ ...prev, shapes: [] }));
        setSelectedId(null);
      }, [commit]);

      const closeImage = useCallback(async () => {
        const ok = await confirmDialog({
          title: 'Close this image?',
          message: 'The image and every annotation on it will be discarded. Copy or download first if you still need it.',
          confirmText: 'Close image',
          tone: 'danger',
        });
        if (!ok) return;
        if (objectUrlRef.current) { URL.revokeObjectURL(objectUrlRef.current); objectUrlRef.current = null; }
        resetEditingState();
        setScene({ img: null, name: '', width: 0, height: 0, shapes: [] });
      }, [resetEditingState, setScene]);

      const applyCrop = useCallback(() => {
        const r = cropRect;
        if (!r) return;
        const x = Math.round(imageClamp(r.x, 0, scene.width));
        const y = Math.round(imageClamp(r.y, 0, scene.height));
        const w = Math.round(Math.min(r.w, scene.width - x));
        const h = Math.round(Math.min(r.h, scene.height - y));
        if (w < 8 || h < 8) { onToast('err', 'That crop is too small'); return; }
        // Crop the base image only and shift the annotations to match, so every
        // arrow and blur stays a separate, still-editable object.
        const out = document.createElement('canvas');
        out.width = w; out.height = h;
        out.getContext('2d').drawImage(scene.img, x, y, w, h, 0, 0, w, h);
        commit(prev => ({
          ...prev,
          img: out, width: w, height: h,
          shapes: prev.shapes.map(s => imageMoveShape(s, -x, -y)),
        }));
        setCropRect(null);
        setPrefs({ tool: 'select' });
      }, [commit, cropRect, onToast, scene.img, scene.width, scene.height, setPrefs]);

      // Encodes the image-resolution scene, not the visible canvas — the visible
      // one is sized for the screen and would export at whatever the current zoom
      // happened to be.
      const toBlob = useCallback((format) => new Promise((resolve, reject) => {
        const canvas = sceneCanvasRef.current;
        if (!canvas || !sceneRef.current.img) { reject(new Error('There is no image to export')); return; }
        const done = (blob) => blob ? resolve(blob) : reject(new Error('The browser could not encode the image'));
        if (format === 'jpeg') {
          // JPEG has no alpha — flatten onto white so transparent PNGs don't
          // come out with black backgrounds.
          const flat = document.createElement('canvas');
          flat.width = canvas.width; flat.height = canvas.height;
          const ctx = flat.getContext('2d');
          ctx.fillStyle = '#FFFFFF';
          ctx.fillRect(0, 0, flat.width, flat.height);
          ctx.drawImage(canvas, 0, 0);
          flat.toBlob(done, 'image/jpeg', 0.92);
          return;
        }
        canvas.toBlob(done, 'image/png');
      }), []);

      const copyImage = useCallback(async () => {
        if (!sceneRef.current.img) return;
        setBusy(true);
        try {
          if (!navigator.clipboard || typeof window.ClipboardItem === 'undefined') {
            throw new Error('this browser cannot put images on the clipboard');
          }
          const blob = await toBlob('png');
          await navigator.clipboard.write([new window.ClipboardItem({ 'image/png': blob })]);
          onToast('ok', 'Copied — paste it anywhere');
        } catch (e) {
          onToast('err', 'Copy failed: ' + (e.message || e) + '. Use Download instead.');
        } finally {
          setBusy(false);
        }
      }, [onToast, toBlob]);

      const downloadImage = useCallback(async () => {
        if (!sceneRef.current.img) return;
        setBusy(true);
        try {
          const format = prefs.exportFormat;
          const blob = await toBlob(format);
          const ext = format === 'jpeg' ? 'jpg' : 'png';
          saveAs(blob, `image_${bmrTodayString()}_${Date.now().toString().slice(-5)}.${ext}`);
          onToast('ok', `Downloaded .${ext}`);
        } catch (e) {
          onToast('err', 'Download failed: ' + (e.message || e));
        } finally {
          setBusy(false);
        }
      }, [onToast, prefs.exportFormat, toBlob]);

      /* ---------- toolbar bindings ---------- */

      // Changing colour or size with a shape selected retargets that shape too,
      // so a wrong-coloured arrow is a two-click fix instead of undo-and-redraw.
      //
      // A slider drag or a colour-picker drag fires dozens of change events. They
      // apply live but share ONE undo entry: the scene is snapshotted when the
      // interaction starts and pushed to history when it ends, rather than each
      // tick pushing its own entry and flushing the rest of the stack out.
      const beginPropEdit = useCallback(() => {
        if (selectedId && !editSessionRef.current) editSessionRef.current = sceneRef.current;
      }, [selectedId]);

      const endPropEdit = useCallback(() => {
        const before = editSessionRef.current;
        editSessionRef.current = null;
        if (before && before !== sceneRef.current) pushHistory(before);
      }, [pushHistory]);

      const patchSelected = useCallback((patch) => {
        if (!selectedId) return;
        beginPropEdit();
        setScene(prev => ({
          ...prev,
          shapes: prev.shapes.map(s => s.id === selectedId ? { ...s, ...patch } : s),
        }));
      }, [beginPropEdit, selectedId, setScene]);

      const setColor = useCallback((color, { live = false } = {}) => {
        setPrefs({ color });
        patchSelected({ color });
        if (!live) endPropEdit(); // swatch clicks are a single discrete edit
      }, [endPropEdit, patchSelected, setPrefs]);

      const sizeControl = imageSizeControlFor(selected ? selected.type : prefs.tool);
      const sizeValue = sizeControl
        ? (selected ? (Number(selected[sizeControl.shapeKey]) || prefs[sizeControl.key]) : prefs[sizeControl.key])
        : 0;

      const setSize = useCallback((value) => {
        if (!sizeControl) return;
        const n = imageClamp(Number(value) || sizeControl.min, sizeControl.min, sizeControl.max);
        setPrefs({ [sizeControl.key]: n });
        patchSelected({ [sizeControl.shapeKey]: n });
      }, [patchSelected, setPrefs, sizeControl]);

      /* ---------- keyboard ---------- */

      useEffect(() => {
        const onKeyDown = (e) => {
          const el = e.target;
          const typing = el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable);
          const mod = e.ctrlKey || e.metaKey;

          if (mod && !typing) {
            const key = e.key.toLowerCase();
            if (key === 'z') { e.preventDefault(); if (e.shiftKey) redo(); else undo(); return; }
            if (key === 'y') { e.preventDefault(); redo(); return; }
            if (key === 'c') { e.preventDefault(); copyImage(); return; }
            if (key === 's') { e.preventDefault(); downloadImage(); return; }
            return;
          }
          if (typing || mod || e.altKey) return;

          if (e.key === 'Delete' || e.key === 'Backspace') {
            if (selectedId) { e.preventDefault(); deleteSelected(); }
            return;
          }
          if (e.key === 'Escape') {
            setSelectedId(null); setCropRect(null); setTextDraft(null);
            return;
          }
          const tool = IMAGE_TOOLS.find(t => t.hotkey.toLowerCase() === e.key.toLowerCase());
          if (tool) { e.preventDefault(); setPrefs({ tool: tool.id }); }
        };
        window.addEventListener('keydown', onKeyDown);
        return () => window.removeEventListener('keydown', onKeyDown);
      }, [copyImage, deleteSelected, downloadImage, redo, selectedId, setPrefs, setTextDraft, undo]);

      useEffect(() => {
        if (textDraft && textInputRef.current) textInputRef.current.focus();
      }, [textDraft?.x, textDraft?.y]);

      /* ---------- view ---------- */

      const selectionBox = selected ? imageShapeBounds(selected) : null;
      const cursor = prefs.tool === 'select' ? 'default' : (prefs.tool === 'text' ? 'text' : 'crosshair');

      return (
        <>
          <ImageSidebar
            state={state} setState={setState} sync={sync} onRetrySync={onRetrySync}
            scene={scene} onPickFile={loadFile} onBlank={loadBlank}
            onCopy={copyImage} onDownload={downloadImage} busy={busy} />

          <main className="flex-1 min-w-0 flex flex-col">
            <header className="app-topbar border-b border-neutral-900 px-4 sm:px-6 lg:px-8 py-4 sm:py-5 flex flex-wrap items-start justify-between gap-4 sticky top-0 bg-[#1c1c1f]/95 backdrop-blur z-10">
              <div className="app-title-block min-w-0">
                <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-1">Module</div>
                <h2 className="truncate text-[22px] font-bold tracking-tight">
                  Image Editor{hasImage ? ` · ${scene.width} × ${scene.height}` : ''}
                </h2>
              </div>
              <div className="app-pill-row hidden xl:flex items-center gap-2">
                <Pill tone={scene.shapes.length ? 'accent' : 'muted'}>{scene.shapes.length} edits</Pill>
                <Pill>{Math.round(zoomLevel * 100)}%</Pill>
              </div>
              <div className="app-header-controls flex flex-wrap items-center gap-2 sm:gap-3">
                {moduleSwitch}
                <Btn variant="ghost" size="sm" onClick={undo} disabled={!historyLen} title="Undo (Ctrl+Z)"><IconUndo /> Undo</Btn>
                <Btn variant="ghost" size="sm" onClick={redo} disabled={!futureLen} title="Redo (Ctrl+Y)"><IconRedo /> Redo</Btn>
                {hasImage && <Btn variant="ghost" size="sm" onClick={closeImage}>Close</Btn>}
              </div>
            </header>

            <div className="p-4 sm:p-6 lg:p-8 w-full max-w-[1700px] space-y-4">
              {hasImage && (
                <Card className="p-3">
                  <div className="flex flex-wrap items-center gap-x-4 gap-y-3">
                    {/* Tools */}
                    <div className="flex flex-wrap items-center gap-1 rounded-md border border-neutral-900 bg-[#1a1a1d] p-1">
                      {IMAGE_TOOLS.map(tool => (
                        <button key={tool.id} type="button"
                          onClick={() => setPrefs({ tool: tool.id })}
                          title={`${tool.hint} (${tool.hotkey})`}
                          className={`inline-flex items-center gap-1.5 rounded px-2.5 py-1.5 text-xs font-medium transition-colors ${prefs.tool === tool.id ? 'app-tab-active bg-neutral-800 text-neutral-100' : 'app-tab-muted text-neutral-500 hover:text-neutral-200'}`}>
                          <IconTool id={tool.id} />
                          <span className="hidden lg:inline">{tool.label}</span>
                        </button>
                      ))}
                    </div>

                    {/* Colour */}
                    <div className="flex items-center gap-1.5">
                      {IMAGE_COLORS.map(c => (
                        <button key={c} type="button" onClick={() => setColor(c)} title={c}
                          className={`h-6 w-6 rounded-full border transition-transform hover:scale-110 ${(selected ? selected.color : prefs.color).toLowerCase() === c.toLowerCase() ? 'border-blue-400 ring-2 ring-blue-500/40' : 'border-neutral-700'}`}
                          style={{ background: c }} />
                      ))}
                      <input type="color" value={selected ? (selected.color || prefs.color) : prefs.color}
                        onChange={e => setColor(e.target.value, { live: true })}
                        onPointerUp={endPropEdit} onBlur={endPropEdit}
                        title="Custom colour"
                        className="h-6 w-8 cursor-pointer rounded border border-neutral-700 bg-transparent p-0" />
                    </div>

                    {/* Contextual size */}
                    {sizeControl && (
                      <div className="flex items-center gap-2 min-w-[190px]">
                        <span className="text-[10px] uppercase tracking-wide text-neutral-500 shrink-0">{sizeControl.label}</span>
                        <input type="range" min={sizeControl.min} max={sizeControl.max} value={sizeValue}
                          onChange={e => setSize(e.target.value)}
                          onPointerUp={endPropEdit} onKeyUp={endPropEdit} onBlur={endPropEdit}
                          className="flex-1 accent-blue-500" />
                        <span className="w-7 text-right font-mono text-[11px] text-neutral-400">{Math.round(sizeValue)}</span>
                      </div>
                    )}

                    <div className="ml-auto flex items-center gap-2">
                      {selected && (
                        <Btn variant="danger" size="sm" onClick={deleteSelected} title="Delete the selected annotation (Del)">
                          <IconX /> Delete
                        </Btn>
                      )}
                      <Btn variant="ghost" size="sm" onClick={clearShapes} disabled={!scene.shapes.length}>Clear all</Btn>
                      <div className="flex items-center gap-0.5 rounded-md border border-neutral-900 bg-[#1a1a1d] p-0.5">
                        <button type="button" onClick={() => setZoom(z => imageClamp((z === 'fit' ? fitZoom : z) - 0.25, 0.25, 4))}
                          className="px-2 py-1 text-xs text-neutral-400 hover:text-neutral-100" title="Zoom out">−</button>
                        <button type="button" onClick={() => setZoom('fit')}
                          className={`px-2 py-1 text-xs font-medium ${zoom === 'fit' ? 'text-neutral-100' : 'text-neutral-500 hover:text-neutral-200'}`}>Fit</button>
                        <button type="button" onClick={() => setZoom(1)}
                          className={`px-2 py-1 text-xs font-medium ${zoom === 1 ? 'text-neutral-100' : 'text-neutral-500 hover:text-neutral-200'}`}>1:1</button>
                        <button type="button" onClick={() => setZoom(z => imageClamp((z === 'fit' ? fitZoom : z) + 0.25, 0.25, 4))}
                          className="px-2 py-1 text-xs text-neutral-400 hover:text-neutral-100" title="Zoom in">+</button>
                      </div>
                    </div>
                  </div>

                  {cropRect && cropRect.w > 4 && cropRect.h > 4 && (
                    <div className="mt-3 flex flex-wrap items-center gap-2 rounded-md border border-blue-500/40 bg-blue-500/10 px-3 py-2">
                      <span className="text-xs text-blue-200">
                        Crop to {Math.round(cropRect.w)} × {Math.round(cropRect.h)}
                      </span>
                      <div className="ml-auto flex items-center gap-2">
                        <Btn variant="ghost" size="sm" onClick={() => setCropRect(null)}>Cancel</Btn>
                        <Btn variant="accent" size="sm" onClick={applyCrop}>Apply crop</Btn>
                      </div>
                    </div>
                  )}
                </Card>
              )}

              {/* Stage */}
              <div ref={stageRef}
                onDragOver={e => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={e => {
                  e.preventDefault();
                  setDragOver(false);
                  const file = e.dataTransfer?.files?.[0];
                  if (file) loadFile(file);
                }}
                className={`rounded-lg border ${dragOver ? 'border-blue-500 bg-blue-500/5' : 'border-neutral-900 bg-[#1a1a1d]'} transition-colors`}>
                {hasImage ? (
                  <div className="flex justify-center p-4 overflow-auto">
                    <div className="relative shrink-0" style={{ width: displayW, height: displayH }}>
                      <canvas ref={canvasRef}
                        onPointerDown={onPointerDown}
                        onPointerMove={onPointerMove}
                        onPointerUp={onPointerUp}
                        onPointerCancel={onPointerUp}
                        style={{ width: displayW, height: displayH, cursor, touchAction: 'none' }}
                        className="block rounded shadow-2xl shadow-black/40 bg-white" />

                      {/* Overlays live in the DOM, not on the canvas, so they can
                          never end up baked into the exported image. */}
                      {selectionBox && (
                        <div className="pointer-events-none absolute border border-dashed border-blue-400 bg-blue-400/10"
                          style={{
                            left: (selectionBox.x * scale) - 4,
                            top: (selectionBox.y * scale) - 4,
                            width: (selectionBox.w * scale) + 8,
                            height: (selectionBox.h * scale) + 8,
                          }} />
                      )}
                      {cropRect && cropRect.w > 0 && cropRect.h > 0 && (
                        <div className="pointer-events-none absolute border-2 border-blue-400"
                          style={{
                            left: cropRect.x * scale,
                            top: cropRect.y * scale,
                            width: cropRect.w * scale,
                            height: cropRect.h * scale,
                            boxShadow: '0 0 0 9999px rgba(0,0,0,0.55)',
                          }} />
                      )}
                      {textDraft && (
                        <textarea ref={textInputRef}
                          value={textDraft.value}
                          onChange={e => setTextDraft({ ...textDraft, value: e.target.value })}
                          onBlur={commitText}
                          onKeyDown={e => {
                            e.stopPropagation();
                            if (e.key === 'Escape') { e.preventDefault(); setTextDraft(null); }
                            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitText(); }
                          }}
                          rows={1}
                          spellCheck={false}
                          placeholder="Type, then Enter"
                          style={{
                            position: 'absolute',
                            left: textDraft.x * scale,
                            top: textDraft.y * scale,
                            minWidth: 140,
                            fontSize: Math.max(11, prefs.fontSize * scale),
                            lineHeight: 1.25,
                            color: prefs.color,
                            fontWeight: 600,
                            fontFamily: IMAGE_FONT_STACK,
                            padding: 0,
                            margin: 0,
                            background: 'rgba(0,0,0,0.35)',
                            border: '1px dashed rgba(96,165,250,0.9)',
                            outline: 'none',
                            resize: 'none',
                            overflow: 'hidden',
                          }} />
                      )}
                    </div>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center gap-4 px-6 py-24 text-center">
                    <div className="flex h-14 w-14 items-center justify-center rounded-full border border-neutral-800 bg-neutral-950 text-neutral-600">
                      <IconImage />
                    </div>
                    <div>
                      <div className="text-sm font-medium text-neutral-200">Paste a screenshot to start</div>
                      <p className="mt-1 text-xs text-neutral-500">
                        Press <kbd className="rounded border border-neutral-700 bg-neutral-900 px-1.5 py-0.5 font-mono text-[10px] text-neutral-300">Ctrl</kbd>
                        {' + '}
                        <kbd className="rounded border border-neutral-700 bg-neutral-900 px-1.5 py-0.5 font-mono text-[10px] text-neutral-300">V</kbd>
                        , drop an image file here, or upload one from the sidebar.
                      </p>
                    </div>
                  </div>
                )}
              </div>

              {hasImage && (
                <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-[11px] text-neutral-600">
                  <span>Drag to draw</span>
                  <span>·</span>
                  <span><span className="text-neutral-400">V</span> select, move, <span className="text-neutral-400">Del</span> to remove</span>
                  <span>·</span>
                  <span><span className="text-neutral-400">Ctrl+Z</span> undo</span>
                  <span>·</span>
                  <span><span className="text-neutral-400">Ctrl+C</span> copy the result</span>
                  {!imageSupportsFilter() && (
                    <>
                      <span>·</span>
                      <span className="text-amber-500">This browser has no canvas blur — Blur falls back to a mosaic.</span>
                    </>
                  )}
                </div>
              )}
            </div>

            <footer className="mt-auto border-t border-neutral-900 px-8 py-4 flex items-center justify-between text-[11px] text-neutral-600">
              <span className="font-mono">in-browser only · nothing uploaded</span>
              <span>{hasImage ? `${scene.width} × ${scene.height} · ${scene.shapes.length} edits` : 'no image loaded'}</span>
            </footer>
          </main>
        </>
      );
    }

    /* ============================================================
       App
       ============================================================ */
    const APP_STATE_SCHEMA_VERSION = 2;

    const DEFAULT_STATE = {
      schemaVersion: APP_STATE_SCHEMA_VERSION,
      year: new Date().getFullYear(),
      month: new Date().getMonth(),
      sheets: DEFAULT_SHEETS,
      selectedSheetId: 's_sbc',
      rules: [],
      includeIndex: true,
      changelog: '',
      tab: 'config',
      theme: 'dark',
      imported: [],
      module: 'sip_fcs',
      stateOwnerUid: '',
      stateOwnerEmail: '',
      bmr: DEFAULT_BMR_STATE,
      bmrTab: 'clients',
      bmrSms: DEFAULT_BMR_SMS_STATE,
      bmrSmsTab: 'clients',
      editor: DEFAULT_EDITOR_STATE,
      editorTab: 'table',
      whitelistSms: DEFAULT_WL_SMS_STATE,
      whitelistSmsTab: 'content',
      rca: DEFAULT_RCA_STATE,
      rcaTab: 'editor',
      rcaSignatories: DEFAULT_RCA_SIGNATORIES,
      rcaCompany: { ...DEFAULT_RCA_COMPANY },
      // Image Editor: tool preferences only. The image itself is never stored
      // here — see the Image Editor module note.
      imageEditor: { ...DEFAULT_IMAGE_STATE },
      // Google Sheets sync: the OAuth Client ID (set in-app, non-secret) plus
      // durable file IDs of the app-owned sheets (one per module). No tokens
      // here — all non-secret, so it rides the normal localStorage + Firestore
      // sync. sheetIds are empty until the first sync creates them.
      googleSheets: { clientId: '', sheetIds: { sip_fcs: '', bmr: '', bmr_sms: '', recorder: '' } },
    };

    function timestampMs(value) {
      if (!value) return 0;
      if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
      if (typeof value === 'string') {
        const parsed = Date.parse(value);
        return Number.isFinite(parsed) ? parsed : 0;
      }
      if (typeof value.toMillis === 'function') return value.toMillis();
      if (typeof value.seconds === 'number') {
        return (value.seconds * 1000) + Math.floor((value.nanoseconds || 0) / 1000000);
      }
      return 0;
    }

    function stateEditedAt(data) {
      return Math.max(timestampMs(data?.localUpdatedAt), timestampMs(data?.updatedAt));
    }

    function stateOwnerUid(data) {
      return String(data?.stateOwnerUid || '').trim();
    }

    function withStateOwner(data, user) {
      return normalizeAppState({
        ...(data || {}),
        stateOwnerUid: user?.uid || '',
        stateOwnerEmail: user?.email || '',
      });
    }

    // Firestore does not allow arrays directly inside arrays. Wrap any nested
    // array as `{ __arr: [...] }` on write and unwrap on read. This is needed
    // for editor.cells (a 2D string grid) and future-proofs any other 2D field.
    const FIRESTORE_NESTED_ARRAY_KEY = '__arr';
    function firestoreSanitize(value) {
      if (Array.isArray(value)) {
        return value.map(v => Array.isArray(v)
          ? { [FIRESTORE_NESTED_ARRAY_KEY]: firestoreSanitize(v) }
          : firestoreSanitize(v));
      }
      if (value && typeof value === 'object') {
        const out = {};
        for (const [k, v] of Object.entries(value)) {
          if (v === undefined) continue;
          out[k] = firestoreSanitize(v);
        }
        return out;
      }
      return value;
    }
    function firestoreDesanitize(value) {
      if (Array.isArray(value)) {
        return value.map(v => {
          if (v && typeof v === 'object' && !Array.isArray(v)
              && Object.prototype.hasOwnProperty.call(v, FIRESTORE_NESTED_ARRAY_KEY)
              && Object.keys(v).length === 1) {
            return firestoreDesanitize(v[FIRESTORE_NESTED_ARRAY_KEY]);
          }
          return firestoreDesanitize(v);
        });
      }
      if (value && typeof value === 'object') {
        const out = {};
        for (const [k, v] of Object.entries(value)) out[k] = firestoreDesanitize(v);
        return out;
      }
      return value;
    }

    function migrateSheetItem(item) {
      const label = itemLabel(item);
      const nextLabel = label.trim().toLowerCase() === 'remarks from tse' ? 'Remarks from L2' : label;
      if (nextLabel === label) return item;
      return typeof item === 'string' ? nextLabel : { ...item, label: nextLabel };
    }

    function migrateSheets(sheets) {
      const list = Array.isArray(sheets) && sheets.length ? sheets : DEFAULT_SHEETS;
      return list.map(sheet => ({
        ...sheet,
        ...(Array.isArray(sheet.columns) ? { columns: sheet.columns.map(migrateSheetItem) } : {}),
        ...(Array.isArray(sheet.metrics) ? { metrics: sheet.metrics.map(migrateSheetItem) } : {}),
      }));
    }

    function normalizeBmrSmsState(raw = {}) {
      const sms = { ...DEFAULT_BMR_SMS_STATE, ...(raw || {}) };
      return {
        ...sms,
        retailClients: Array.isArray(sms.retailClients) ? sms.retailClients : DEFAULT_BMR_SMS_STATE.retailClients,
        wholesaleClients: Array.isArray(sms.wholesaleClients) ? sms.wholesaleClients : DEFAULT_BMR_SMS_STATE.wholesaleClients,
        accountManagers: Array.isArray(sms.accountManagers) ? sms.accountManagers : [],
        targetRules: Array.isArray(sms.targetRules) ? sms.targetRules : [],
        overdraftRules: Array.isArray(sms.overdraftRules) ? sms.overdraftRules : DEFAULT_BMR_SMS_STATE.overdraftRules,
        usageRules: Array.isArray(sms.usageRules) ? sms.usageRules : DEFAULT_BMR_SMS_STATE.usageRules,
      };
    }

    function normalizeAppState(raw = {}, options = {}) {
      const migrateLegacyLabels = !!options.migrateLegacyLabels
        && Number(raw?.schemaVersion || 0) < APP_STATE_SCHEMA_VERSION;
      const merged = { ...DEFAULT_STATE, ...(raw || {}) };
      const bmr = { ...DEFAULT_BMR_STATE, ...(merged.bmr || {}) };
      const bmrSms = normalizeBmrSmsState(merged.bmrSms);
      const editorRaw = merged.editor || {};
      const editor = { ...DEFAULT_EDITOR_STATE, ...editorRaw, cells: editorNormalizeGrid(editorRaw.cells || DEFAULT_EDITOR_STATE.cells) };
      const whitelistSms = wlSmsNormalizeState(merged.whitelistSms);
      const rca = rcaNormalizeState(merged.rca);
      const rcaSignatories = Array.isArray(merged.rcaSignatories) && merged.rcaSignatories.length
        ? merged.rcaSignatories : DEFAULT_RCA_SIGNATORIES;
      const rcaCompany = rcaNormalizeCompany(merged.rcaCompany);
      const imageEditor = imageNormalizeState(merged.imageEditor);
      const recorder = recorderNormalizeState(merged.recorder);
      const googleSheets = {
        clientId: (merged.googleSheets && typeof merged.googleSheets.clientId === 'string')
          ? merged.googleSheets.clientId
          : (DEFAULT_STATE.googleSheets.clientId || ''),
        sheetIds: {
          ...DEFAULT_STATE.googleSheets.sheetIds,
          ...((merged.googleSheets && merged.googleSheets.sheetIds) || {}),
        },
      };
      const sheets = migrateLegacyLabels
        ? migrateSheets(merged.sheets)
        : (Array.isArray(merged.sheets) && merged.sheets.length ? merged.sheets : DEFAULT_SHEETS);
      return {
        ...merged,
        schemaVersion: APP_STATE_SCHEMA_VERSION,
        sheets,
        stateOwnerUid: String(merged.stateOwnerUid || ''),
        stateOwnerEmail: String(merged.stateOwnerEmail || ''),
        bmr,
        bmrSms,
        editor,
        whitelistSms,
        rca,
        rcaSignatories,
        rcaCompany,
        imageEditor,
        recorder,
        googleSheets,
      };
    }

    /* ────────────────────────────────────────────────────────────────────
       Recorder module (team shared — every signed-in user)
       Paste VOS / dashboard screenshots → OCR (Tesseract.js, lazy-loaded)
       → match configured gateway/metric names → commit numbers into a
       per-client month grid → export the hourly record workbook.
       Screenshots are processed in-memory only; `state.recorder` holds
       nothing but plain numbers + config (same localStorage/Firestore
       budget rules as the Image Editor — see PROJECT_RESUME §11b).
       Values + config are ALSO mirrored to shared Firestore docs so the
       whole team works on one grid — see the shared team sync helpers.
       ──────────────────────────────────────────────────────────────────── */

    const RECORDER_MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    let recorderUidCounter = 0;
    function recorderUid() {
      recorderUidCounter += 1;
      return 'rc' + Date.now().toString(36) + (recorderUidCounter % 1296).toString(36) + Math.random().toString(36).slice(2, 6);
    }
    function recorderRowsSeed(prefix, defs) {
      return defs.map(([label, color, aliases], i) => ({
        id: `${prefix}_${i + 1}`, label, color, aliases: aliases || [],
      }));
    }
    // Seeded from VOS_MIRTA_Record_Aug_2026.xlsx — row order, group colors and
    // separator colors reproduce that workbook exactly.
    const RECORDER_DEFAULT_CLIENTS = [
      { id: 'rc_vos', name: 'VOS', type: 'gateway', bordered: false, headerBorder: false, sepColor: '#000000',
        rows: recorderRowsSeed('rv', [
          ['IN_ETPI-SIP', '#FFFF00'],
          ['IN_GLOBE-SIP', '#FFFF00'],
          ['IN_PLDT-FCS-CYN', '#FFFF00'],
          ['IN_PLDT-SIP', '#FFFF00'],
          ['IN_PLDT-SIP-FEDEX', '#FFFF00'],
          ['IN_PLDT-SIP-TOKU', '#FFFF00'],
          ['V_ETPI-SIP', '#FFE599'],
          ['V_ETPI-SIP RANDOM', '#FFE599'],
          ['V_ETPI-SIP RANDOM_MOB', '#FFE599'],
          ['V_PLDT-SIP', '#FFE599'],
          ['V_PLDT-FCS-CYN', '#FFD966'],
          ['V_PLDT-FCS-CYN RANDOM', '#FFD966'],
          ['V_PLDT-FCS-CYN RANDOM_MNL', '#FFD966'],
          ['V_PLDT-FCS-CYN RANDOM_MOB', '#FFD966'],
        ]) },
      { id: 'rc_twilio', name: 'VOS TWILIO', type: 'gateway', bordered: false, headerBorder: true, sepColor: '#000000',
        rows: recorderRowsSeed('rt', [
          ['C_CYN_TWILIO  STD', '#FFFF00'],
          ['C_CYN_TWILIO  STD2', '#FFFF00'],
          ['C_CYN_TWILIO  STD3', '#FFFF00'],
          ['C_CYN_TWILIO  STD5', '#FFFF00'],
        ]) },
      { id: 'rc_kingsford', name: 'KINGSFORD INOUT', type: 'cards', bordered: true, headerBorder: true, sepColor: '#000000',
        rows: recorderRowsSeed('rk', [
          ['Current Inbound', '#FFFF00'],
          ['MaxInbound', '#FFFF00', ['Max Inbound']],
          ['Current Outbound', '#FFFF00'],
          ['Max Outbound', '#FFFF00'],
        ]) },
      { id: 'rc_unobank', name: 'UNOBANK IN', type: 'cards', bordered: true, headerBorder: true, sepColor: '#FFE599',
        rows: recorderRowsSeed('ru', [
          ['Current Inbound', '#FFFF00'],
          ['MaxInbound', '#FFFF00', ['Max Inbound']],
        ]) },
      { id: 'rc_myvelox', name: 'MYVELOX INOUT', type: 'cards', bordered: true, headerBorder: true, sepColor: '#000000',
        rows: recorderRowsSeed('rm', [
          ['Current Inbound', '#FFFF00'],
          ['MaxInbound', '#FFFF00', ['Max Inbound']],
          ['Current Outbound', '#FFFF00'],
          ['Max Outbound', '#FFFF00'],
        ]) },
    ];

    function recorderNormalizeClient(c) {
      if (!c || typeof c !== 'object') return null;
      const rows = Array.isArray(c.rows) ? c.rows.map(r => {
        if (!r || typeof r !== 'object' || !r.label) return null;
        return {
          id: String(r.id || recorderUid()),
          label: String(r.label),
          color: presetColorHex(r.color, '#FFFF00'),
          aliases: Array.isArray(r.aliases) ? r.aliases.map(a => String(a)).filter(Boolean) : [],
        };
      }).filter(Boolean) : [];
      return {
        id: String(c.id || recorderUid()),
        name: String(c.name || 'Client'),
        type: c.type === 'cards' ? 'cards' : 'gateway',
        bordered: !!c.bordered,
        headerBorder: !!c.headerBorder,
        sepColor: presetColorHex(c.sepColor, '#000000'),
        rows,
      };
    }
    // Rebuilt field-by-field so a stale cache can never smuggle blobs or
    // malformed entries into the synced doc (same idea as imageNormalizeState).
    function recorderNormalizeState(raw) {
      const src = raw && typeof raw === 'object' ? raw : {};
      const clients = Array.isArray(src.clients) && src.clients.length
        ? src.clients.map(recorderNormalizeClient).filter(Boolean)
        : RECORDER_DEFAULT_CLIENTS.map(c => ({ ...c, rows: c.rows.map(r => ({ ...r, aliases: [...r.aliases] })) }));
      const values = {};
      if (src.values && typeof src.values === 'object') {
        for (const [cid, byDate] of Object.entries(src.values)) {
          if (!byDate || typeof byDate !== 'object') continue;
          const cOut = {};
          for (const [d, byRow] of Object.entries(byDate)) {
            if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !byRow || typeof byRow !== 'object') continue;
            const dOut = {};
            for (const [rid, byHour] of Object.entries(byRow)) {
              if (!byHour || typeof byHour !== 'object') continue;
              const hOut = {};
              for (const [h, v] of Object.entries(byHour)) {
                const hi = Number(h), vi = Number(v);
                if (Number.isInteger(hi) && hi >= 0 && hi < 24 && Number.isFinite(vi)) hOut[hi] = vi;
              }
              if (Object.keys(hOut).length) dOut[rid] = hOut;
            }
            if (Object.keys(dOut).length) cOut[d] = dOut;
          }
          if (Object.keys(cOut).length) values[cid] = cOut;
        }
      }
      return {
        tab: ['capture', 'data', 'config', 'rules'].includes(src.tab) ? src.tab : 'capture',
        clients,
        values,
        filePattern: (typeof src.filePattern === 'string' && src.filePattern.trim()) ? src.filePattern.trim() : 'VOS_MIRTA_Record_{Month}_{Year}',
        // Strict by default: sibling gateways can differ by just 2 characters
        // (RANDOM_MNL / RANDOM_MOB, V_GLOBE / IN_GLOBE), which 'normal' would
        // cross-match when the unwatched sibling appears in the screenshot.
        match: ['strict', 'normal', 'loose'].includes(src.match) ? src.match : 'strict',
        overwrite: src.overwrite !== false,
        upscale: src.upscale !== false,
        clampMax: (Number.isFinite(Number(src.clampMax)) && Number(src.clampMax) > 0) ? Math.floor(Number(src.clampMax)) : 9999,
        // "Highlight max value in row" — same rule as SIP/FCS (defaults copied
        // from its rule editor: green fill, white font). Applies to the export
        // as conditional formatting and to the Data grid as a live preview.
        maxInRow: {
          enabled: !!(src.maxInRow && src.maxInRow.enabled),
          color: presetColorHex(src.maxInRow && src.maxInRow.color, '#166534'),
          fontColor: presetColorHex(src.maxInRow && src.maxInRow.fontColor, '#FFFFFF'),
          bold: !!(src.maxInRow && src.maxInRow.bold),
        },
        // Export month/year live in the slice (not component state) so the
        // Google Sheets sync builder can reproduce the exact workbook the
        // Download button makes. exportYear stays a raw string so normalize
        // never fights the year input mid-typing.
        exportMonth: (src.exportMonth != null && Number.isInteger(Number(src.exportMonth)) && Number(src.exportMonth) >= 0 && Number(src.exportMonth) <= 11)
          ? Number(src.exportMonth) : new Date().getMonth(),
        exportYear: (src.exportYear != null && String(src.exportYear).trim() !== '') ? String(src.exportYear) : String(new Date().getFullYear()),
      };
    }

    /* ── Recorder: shared team sync (Firestore) ──
       Recorder data is shared across ALL signed-in users so the whole team
       sees (and fills) the same hourly grid:
         shared/recorderConfig            → clients/watch rows + settings/rules
         shared/recorderValues_{YYYY-MM}  → one doc per month of readings
       Same design decision as the Whitelist SMS shared number DB: plain
       getDoc/setDoc only, NO onSnapshot — Firestore's watch channel hits
       "INTERNAL ASSERTION FAILED" on some machines (see loadSharedTestNumbers).
       Freshness comes from a refresh on module open / month navigation, a 60s
       poll + window-focus refresh while the module is open, a manual Refresh
       button, and a flush+refresh right before every export.
       Writes are per-cell nested merges (setDoc {merge:true}; deletions become
       deleteField() at flush time), so two users recording different slots
       never clobber each other. Per-month docs keep each doc far below the
       1MB Firestore cap no matter how many months accumulate. */
    const RECORDER_SHARED_CONFIG_DOC = ['shared', 'recorderConfig'];
    const RECORDER_SHARED_CONFIG_FIELDS = ['clients', 'filePattern', 'match', 'overwrite', 'upscale', 'clampMax', 'maxInRow'];
    // One-time per-device flag: which months already had their local cells
    // additively uploaded into the shared docs (rollout migration; the values
    // in localStorage belong to the device, so the flag is device-wide).
    const RECORDER_SHARED_MIGRATED_KEY = 'mrg_recorder_shared_migrated_v1';
    const recorderMonthKey = iso => String(iso || '').slice(0, 7);
    const recorderSharedMonthDocId = mk => `recorderValues_${mk}`;

    // Slice rec.values down to one month: { [clientId]: { [dateIso]: rows } }.
    function recorderMonthSlice(values, mk) {
      const out = {};
      for (const [cid, byDate] of Object.entries(values || {})) {
        for (const [iso, rows] of Object.entries(byDate || {})) {
          if (recorderMonthKey(iso) !== mk) continue;
          (out[cid] = out[cid] || {})[iso] = rows;
        }
      }
      return out;
    }

    // Replace one month of `values` with a shared month slice, leaving every
    // other month untouched.
    function recorderApplyMonthSlice(values, mk, sharedMonth) {
      const next = {};
      for (const [cid, byDate] of Object.entries(values || {})) {
        for (const [iso, rows] of Object.entries(byDate || {})) {
          if (recorderMonthKey(iso) === mk) continue;
          (next[cid] = next[cid] || {})[iso] = rows;
        }
      }
      for (const [cid, byDate] of Object.entries(sharedMonth || {})) {
        for (const [iso, rows] of Object.entries(byDate || {})) {
          if (recorderMonthKey(iso) !== mk) continue;
          if (!rows || typeof rows !== 'object' || !Object.keys(rows).length) continue;
          (next[cid] = next[cid] || {})[iso] = rows;
        }
      }
      return next;
    }

    // Diff two sparse value trees → per-month nested merge objects. `null`
    // marks a deletion (converted to deleteField() only at flush time so the
    // queue stays plain JSON and can be overlaid onto fetched snapshots).
    function recorderDiffValues(prev, next) {
      const months = {};
      const put = (mk, path, v) => {
        let node = (months[mk] = months[mk] || {});
        for (let i = 0; i < path.length - 1; i++) {
          const k = path[i];
          if (node[k] == null || typeof node[k] !== 'object') node[k] = {};
          node = node[k];
        }
        node[path[path.length - 1]] = v;
      };
      const cids = new Set([...Object.keys(prev || {}), ...Object.keys(next || {})]);
      for (const cid of cids) {
        const pc = (prev || {})[cid] || {}, nc = (next || {})[cid] || {};
        const dates = new Set([...Object.keys(pc), ...Object.keys(nc)]);
        for (const iso of dates) {
          const mk = recorderMonthKey(iso);
          const pd = pc[iso], nd = nc[iso];
          if (pd && !nd) { put(mk, [cid, iso], null); continue; }
          const prows = pd || {}, nrows = nd || {};
          const rids = new Set([...Object.keys(prows), ...Object.keys(nrows)]);
          for (const rid of rids) {
            const ph = prows[rid], nh = nrows[rid];
            if (ph && !nh) { put(mk, [cid, iso, rid], null); continue; }
            const phh = ph || {}, nhh = nh || {};
            const hours = new Set([...Object.keys(phh), ...Object.keys(nhh)]);
            for (const h of hours) {
              const pv = phh[h], nv = nhh[h];
              if (pv === nv) continue;
              put(mk, [cid, iso, rid, h], nv == null ? null : nv);
            }
          }
        }
      }
      return months;
    }

    // Deep-merge one queued month diff into another (b wins; null deletion
    // markers replace whole subtrees just like they will in Firestore).
    function recorderMergeQueuedMonth(a, b) {
      if (b == null || typeof b !== 'object') return b;
      if (a == null || typeof a !== 'object') a = {};
      const out = { ...a };
      for (const [k, v] of Object.entries(b)) {
        out[k] = (v != null && typeof v === 'object') ? recorderMergeQueuedMonth(out[k], v) : v;
      }
      return out;
    }

    // Overlay pending (unflushed) queued writes onto a freshly fetched shared
    // month so a refresh can never revert cells the user just typed.
    function recorderOverlayQueuedMonth(sharedMonth, queued) {
      if (!queued) return sharedMonth;
      const walk = (base, q) => {
        const out = (base != null && typeof base === 'object') ? { ...base } : {};
        for (const [k, v] of Object.entries(q)) {
          if (v === null) delete out[k];
          else if (typeof v === 'object') out[k] = walk(out[k], v);
          else out[k] = v;
        }
        return out;
      };
      return walk(sharedMonth || {}, queued);
    }

    // Cells present locally but missing in the shared month → nested add
    // object for a one-time additive migration write (shared wins conflicts,
    // so a stale device can never overwrite the team's newer entries).
    function recorderMissingCells(localMonth, sharedMonth) {
      const out = {};
      let any = false;
      for (const [cid, byDate] of Object.entries(localMonth || {})) {
        for (const [iso, rows] of Object.entries(byDate || {})) {
          for (const [rid, byHour] of Object.entries(rows || {})) {
            for (const [h, v] of Object.entries(byHour || {})) {
              const has = sharedMonth?.[cid]?.[iso]?.[rid]?.[h];
              if (has != null || v == null) continue;
              ((((out[cid] = out[cid] || {})[iso] = out[cid][iso] || {})[rid] = out[cid][iso][rid] || {}))[h] = v;
              any = true;
            }
          }
        }
      }
      return any ? out : null;
    }

    // Queue → Firestore write shape: null markers become deleteField().
    function recorderQueueToFirestore(queued, deleteField) {
      const walk = (q) => {
        const out = {};
        for (const [k, v] of Object.entries(q)) {
          out[k] = v === null ? deleteField() : (typeof v === 'object' ? walk(v) : v);
        }
        return out;
      };
      return walk(queued || {});
    }

    /* ── Recorder: name matching ── */
    const recorderNorm = s => String(s || '').toUpperCase().replace(/\u00A0/g, ' ').replace(/\|/g, 'I').replace(/\s+/g, ' ').trim();
    const recorderSig = s => recorderNorm(s).replace(/[^A-Z0-9]/g, '');
    // Signature with classic OCR look-alikes folded to one class (5\u2194S, 0\u2194O/Q,
    // 1\u2194I/L, 2\u2194Z, 8\u2194B, 6\u2194G). Both sides fold, so "C_CYN_TWILIO STDS" (misread
    // 5) and the watched "C_CYN_TWILIO  STD5" land on the same string while
    // STD2 / STD3 stay distinct.
    const recorderFoldSig = s => recorderSig(s)
      .replace(/[OQ]/g, '0').replace(/[IL]/g, '1').replace(/S/g, '5')
      .replace(/Z/g, '2').replace(/B/g, '8').replace(/G/g, '6');
    function recorderLev(a, b, cap = 3) {
      if (a === b) return 0;
      const la = a.length, lb = b.length;
      if (Math.abs(la - lb) > cap) return cap + 1;
      if (!la || !lb) return Math.max(la, lb);
      let prev = Array.from({ length: lb + 1 }, (_, i) => i);
      for (let i = 1; i <= la; i++) {
        const cur = [i];
        let rowMin = i;
        for (let j = 1; j <= lb; j++) {
          cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
          if (cur[j] < rowMin) rowMin = cur[j];
        }
        if (rowMin > cap) return cap + 1;
        prev = cur;
      }
      return prev[lb] > cap ? cap + 1 : prev[lb];
    }
    function recorderWordLike(text, target) {
      const a = String(text || '').toLowerCase().replace(/[^a-z]/g, '');
      if (!a) return false;
      if (a === target) return true;
      const cap = target.length >= 6 ? 2 : 1;
      return recorderLev(a, target, cap) <= cap;
    }
    // Exact (normalized) match first, then signature match (separators removed),
    // then unique fuzzy match. Never prefix/contains — several watched names are
    // prefixes of each other (IN_PLDT-SIP vs IN_PLDT-SIP-TOKU, STD vs STD2 …).
    function recorderMatchGatewayRow(name, clients, mode) {
      const n = recorderNorm(name), g = recorderSig(name);
      if (!g) return null;
      const all = [];
      for (const c of clients) {
        if (c.type !== 'gateway') continue;
        for (const r of c.rows) all.push({ c, r });
      }
      let hit = all.find(({ r }) => recorderNorm(r.label) === n || r.aliases.some(a => recorderNorm(a) === n));
      if (!hit) hit = all.find(({ r }) => recorderSig(r.label) === g || r.aliases.some(a => recorderSig(a) === g));
      if (hit) return { clientId: hit.c.id, rowId: hit.r.id, label: hit.r.label, exact: true };
      // Look-alike fold (5↔S etc.) — deterministic, but only when exactly one
      // watched row folds onto the OCR text.
      const f = recorderFoldSig(name);
      const foldHits = all.filter(({ r }) => recorderFoldSig(r.label) === f || r.aliases.some(a => recorderFoldSig(a) === f));
      if (foldHits.length === 1) return { clientId: foldHits[0].c.id, rowId: foldHits[0].r.id, label: foldHits[0].r.label, exact: true };
      const maxD = mode === 'strict' ? 1 : mode === 'loose' ? 3 : 2;
      let best = null, bestD = maxD + 1, secondD = maxD + 1;
      for (const { c, r } of all) {
        let d = recorderLev(recorderFoldSig(r.label), f, maxD);
        for (const a of r.aliases) d = Math.min(d, recorderLev(recorderFoldSig(a), f, maxD));
        if (d < bestD) { secondD = bestD; bestD = d; best = { c, r }; }
        else if (d < secondD) secondD = d;
      }
      if (best && bestD <= maxD && bestD < secondD) {
        return { clientId: best.c.id, rowId: best.r.id, label: best.r.label, exact: false, dist: bestD };
      }
      return null;
    }
    function recorderMatchClientRow(name, client) {
      const n = recorderNorm(name), g = recorderSig(name), f = recorderFoldSig(name);
      const row = client.rows.find(r =>
        recorderNorm(r.label) === n || recorderSig(r.label) === g
        || r.aliases.some(a => recorderNorm(a) === n || recorderSig(a) === g));
      if (row) return row;
      const foldHits = client.rows.filter(r =>
        recorderFoldSig(r.label) === f || r.aliases.some(a => recorderFoldSig(a) === f));
      return foldHits.length === 1 ? foldHits[0] : null;
    }

    /* ── Recorder: OCR (Tesseract.js, injected on first use) ── */
    let recorderTessScriptPromise = null;
    function recorderLoadTesseract() {
      if (window.Tesseract) return Promise.resolve(window.Tesseract);
      if (!recorderTessScriptPromise) {
        recorderTessScriptPromise = new Promise((resolve, reject) => {
          const s = document.createElement('script');
          s.src = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
          s.onload = () => window.Tesseract ? resolve(window.Tesseract) : reject(new Error('OCR engine loaded but is unavailable'));
          s.onerror = () => { recorderTessScriptPromise = null; reject(new Error('Could not download the OCR engine — check the internet connection')); };
          document.head.appendChild(s);
        });
      }
      return recorderTessScriptPromise;
    }
    let recorderWorkerPromise = null;
    function recorderEnsureWorker() {
      if (!recorderWorkerPromise) {
        recorderWorkerPromise = (async () => {
          const T = await recorderLoadTesseract();
          const worker = await T.createWorker('eng', 1, {
            logger: m => { try { window.__recorderOcrProgress && window.__recorderOcrProgress(m); } catch {} },
          });
          await worker.setParameters({ tessedit_pageseg_mode: '6', preserve_interword_spaces: '1' });
          return worker;
        })().catch(err => { recorderWorkerPromise = null; throw err; });
      }
      return recorderWorkerPromise;
    }
    async function recorderRecognizeWords(worker, canvas, psm) {
      await worker.setParameters({ tessedit_pageseg_mode: String(psm) });
      const { data } = await worker.recognize(canvas);
      return recorderWordsFromOcr(data);
    }
    async function recorderFileToCanvas(file, upscale) {
      let source = null, w = 0, h = 0;
      try {
        source = await createImageBitmap(file);
        w = source.width; h = source.height;
      } catch {
        source = await new Promise((resolve, reject) => {
          const url = URL.createObjectURL(file);
          const img = new Image();
          img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
          img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Could not read that image')); };
          img.src = url;
        });
        w = source.naturalWidth; h = source.naturalHeight;
      }
      if (!w || !h) throw new Error('Could not read that image');
      // Small UI crops OCR poorly at native size; upscale so thin strokes get
      // enough pixels to survive (aim a large x-height for tight crops like the
      // TWILIO 4-row table). Target ~2400px wide, cap 5x.
      const scale = (upscale && w < 1600) ? Math.max(2, Math.min(5, Math.round(2400 / w))) : 1;
      const canvas = document.createElement('canvas');
      canvas.width = w * scale; canvas.height = h * scale;
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(source, 0, 0, canvas.width, canvas.height);
      // Grayscale + gamma darkening done PER PIXEL, not via ctx.filter. Canvas
      // `filter` is implemented differently by Blink (Chrome/Edge) and Gecko
      // (Firefox): on Blink it faded the thin leading "1" stroke enough that
      // Tesseract dropped the digit, so the SAME screenshot read 117 in Firefox
      // but 17 in Chrome/Edge. A per-pixel transform renders identically on
      // every engine, and gamma > 1 thickens thin anti-aliased strokes while
      // leaving the white background white.
      try {
        const im = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const d = im.data;
        const lut = new Uint8ClampedArray(256);
        for (let v = 0; v < 256; v++) lut[v] = Math.round(255 * Math.pow(v / 255, 1.7));
        for (let i = 0; i < d.length; i += 4) {
          const g = lut[(d[i] * 299 + d[i + 1] * 587 + d[i + 2] * 114) / 1000 | 0];
          d[i] = d[i + 1] = d[i + 2] = g;
        }
        ctx.putImageData(im, 0, 0);
      } catch {}
      return canvas;
    }
    function recorderThumb(canvas) {
      // Kept large enough that the click-to-enlarge view on the capture card
      // stays readable; the data URL lives in component state only (never
      // persisted), so the bigger size costs memory for this session only.
      const w = Math.min(1280, canvas.width);
      const h = Math.max(1, Math.round(canvas.height * (w / canvas.width)));
      const t = document.createElement('canvas');
      t.width = w; t.height = h;
      t.getContext('2d').drawImage(canvas, 0, 0, w, h);
      return t.toDataURL('image/png');
    }
    function recorderWordsFromOcr(data) {
      let words = Array.isArray(data && data.words) ? data.words : [];
      if (!words.length && Array.isArray(data && data.blocks)) {
        for (const b of data.blocks) {
          for (const p of (b.paragraphs || [])) {
            for (const l of (p.lines || [])) words.push(...(l.words || []));
          }
        }
      }
      return words
        .filter(w => w && w.text && w.bbox)
        .map(w => ({
          text: String(w.text).trim(),
          conf: Number(w.confidence != null ? w.confidence : 0),
          x0: w.bbox.x0, y0: w.bbox.y0, x1: w.bbox.x1, y1: w.bbox.y1,
        }))
        .filter(w => w.text);
    }
    function recorderClusterRows(words) {
      if (!words.length) return [];
      const sorted = [...words].sort((a, b) => ((a.y0 + a.y1) / 2) - ((b.y0 + b.y1) / 2));
      const heights = sorted.map(w => w.y1 - w.y0).sort((a, b) => a - b);
      const medH = heights[Math.floor(heights.length / 2)] || 12;
      const rows = [];
      for (const w of sorted) {
        const yc = (w.y0 + w.y1) / 2;
        const row = rows.find(r => Math.abs(r.yc - yc) <= medH * 0.72);
        if (row) { row.words.push(w); row.yc = row.yc + (yc - row.yc) / row.words.length; }
        else rows.push({ yc, words: [w] });
      }
      rows.forEach(r => r.words.sort((a, b) => a.x0 - b.x0));
      return { rows, medH };
    }
    // A value token must be digit-shaped BEFORE repair (0/O, 1/I/l, 5/S), so
    // name fragments like "STD5" can never be mistaken for a reading.
    function recorderNumFromToken(text, maxDigits) {
      // Leading/trailing pipes are table borders that merged into the word,
      // never digits (fixes "|19" → 119 and stray "|" → phantom 1).
      const t = String(text || '').replace(/^\|+/, '').replace(/\|+$/, '');
      if (!t) return null;
      if (/[#,.:%()]/.test(t)) return null;
      if (!/^[0-9OoQqIl|Ss]{1,6}$/.test(t)) return null;
      const s = t.replace(/[OoQq]/g, '0').replace(/[Il|]/g, '1').replace(/[Ss]/g, '5');
      if (!/^\d+$/.test(s) || s.length > maxDigits) return null;
      return parseInt(s, 10);
    }
    // Same shape check as recorderNumFromToken, but returns the repaired DIGIT
    // STRING (not a parsed int) so a number OCR split across words can be
    // stitched back together left-to-right. null when the token isn't a clean
    // digit run.
    function recorderDigitStr(text) {
      const t = String(text || '').replace(/^\|+/, '').replace(/\|+$/, '');
      if (!t || /[#,.:%()]/.test(t) || !/^[0-9OoQqIl|Ss]{1,6}$/.test(t)) return null;
      const s = t.replace(/[OoQq]/g, '0').replace(/[Il|]/g, '1').replace(/[Ss]/g, '5');
      return /^\d+$/.test(s) ? s : null;
    }
    // A value word that starts LEFT of the value column's text edge has likely
    // swallowed the cell border as a leading stroke. Only strip that stroke when
    // it came through as a NON-digit ('|', 'I', 'l') — those are unambiguous
    // border artifacts. A literal leading '1' is NOT stripped: it is
    // indistinguishable from a real hundreds digit, and deciding on the
    // sub-pixel x-threshold silently turned 117 into 17 — and did so only in
    // some browsers, whose canvas rendering shifts the OCR bbox by a pixel or
    // two. Keeping the full digits is both correct and browser-consistent.
    function recorderValueFromWord(w, valueX, maxDigits) {
      if (valueX != null && w.x0 < valueX - 2 && /^[Il|][0-9OoQqIl|Ss]{1,5}$/.test(w.text)) {
        const stripped = recorderNumFromToken(w.text.slice(1), maxDigits);
        if (stripped != null) return stripped;
      }
      return recorderNumFromToken(w.text, maxDigits);
    }
    // Gateway names never contain #, %, commas or long digit runs — prefix
    // cells always do. Used to cut the name off even when the column
    // boundary could not be located from the header.
    function recorderLooksPrefixy(text) {
      const t = String(text || '');
      return /[#%,;]/.test(t) || /^\d{5,}$/.test(t) || /\.{2,}/.test(t);
    }
    function recorderParseTable(words, imgW) {
      const { rows, medH } = recorderClusterRows(words);
      let headerRow = null, prefixX = null, valueX = null;
      for (const r of rows) {
        for (let i = 0; i < r.words.length; i++) {
          const w = r.words[i], next = r.words[i + 1];
          // OCR sometimes merges header cells into one token ("Gatewayname").
          if ((recorderWordLike(w.text, 'gateway') && next && recorderWordLike(next.text, 'name')) || recorderWordLike(w.text, 'gatewayname')) { headerRow = r; break; }
        }
        if (headerRow) break;
      }
      if (headerRow) {
        for (let i = 0; i < headerRow.words.length; i++) {
          const w = headerRow.words[i];
          if (prefixX == null && recorderWordLike(w.text, 'prefix')) {
            const before = headerRow.words[i - 1];
            prefixX = (before && recorderWordLike(before.text, 'gateway')) ? before.x0 : w.x0;
          }
          if (prefixX == null && recorderWordLike(w.text, 'gatewayprefix')) prefixX = w.x0;
          if (recorderWordLike(w.text, 'number') || recorderWordLike(w.text, 'numberofcalling')) valueX = w.x0;
          if (valueX == null && recorderWordLike(w.text, 'calling')) valueX = w.x0 - Math.round(medH * 4);
        }
      }
      const entries = [];
      const nameStop = (prefixX != null ? prefixX : (valueX != null ? valueX : imgW * 0.6)) - Math.max(6, medH * 0.4);
      for (const r of rows) {
        if (headerRow && r.yc <= headerRow.yc + medH * 0.8) continue;
        const nameParts = [];
        for (const w of r.words) {
          if (w.x0 >= nameStop) break;
          if (recorderLooksPrefixy(w.text)) break;
          if (nameParts.length && recorderNumFromToken(w.text, 6) != null) break;
          nameParts.push(w);
        }
        const nameText = nameParts.map(w => w.text).join(' ').trim();
        if (!nameText || recorderSig(nameText).length < 3) continue;
        let value = null, vconf = 100;
        if (valueX != null) {
          // Value-column tokens, left→right. OCR sometimes splits one number
          // across words ("117" → "1" + "17"), and on some browsers the
          // preprocessing canvas renders just differently enough to trigger the
          // split — so the SAME screenshot read 117 on one machine and 17 on
          // another (the old code kept only the rightmost fragment). Anchor on
          // the rightmost digit token and rebuild the number from the run of
          // tightly-adjacent digit fragments (gap < ~half a digit width);
          // a name token or a real gap ends the run.
          const windowed = r.words.filter(w => ((w.x0 + w.x1) / 2) >= valueX - 10);
          let end = -1;
          for (let i = windowed.length - 1; i >= 0; i--) { if (recorderDigitStr(windowed[i].text) != null) { end = i; break; } }
          if (end >= 0) {
            const run = [windowed[end]];
            for (let i = end - 1; i >= 0; i--) {
              const cur = windowed[i];
              if (recorderDigitStr(cur.text) == null) break;
              if ((run[0].x0 - cur.x1) > medH * 0.35) break;
              run.unshift(cur);
            }
            if (run.length > 1) {
              const digits = run.map(w => recorderDigitStr(w.text)).join('');
              if (/^\d+$/.test(digits) && digits.length <= 6) { value = parseInt(digits, 10); vconf = Math.min(...run.map(w => w.conf)); }
            }
            if (value == null) {
              // Single token (or oversized run): keep the original per-word read
              // so an attached leading border (I/l/|) is still stripped.
              const n = recorderValueFromWord(windowed[end], valueX, 6);
              if (n != null) { value = n; vconf = windowed[end].conf; }
            }
          }
        }
        if (value == null) {
          // No value column located — fall back to the rightmost short number
          // in the row (≤4 digits so prefix fragments can't qualify).
          const lastName = nameParts.length ? nameParts[nameParts.length - 1] : null;
          for (let i = r.words.length - 1; i >= 0; i--) {
            const w = r.words[i];
            if (lastName && w.x0 <= lastName.x0) break;
            const n = recorderNumFromToken(w.text, 4);
            if (n != null) { value = n; vconf = w.conf; break; }
          }
        }
        const nameConf = nameParts.length ? Math.min(...nameParts.map(w => w.conf)) : 0;
        entries.push({ name: nameText, value, conf: Math.round(Math.min(nameConf, vconf)), y: r.yc });
      }
      return { entries, meta: { valueX, nameStop, medH } };
    }
    // Same digit-dropping workaround as the cards: block mode loses lone
    // small digits, so rows that came back valueless get their value column
    // re-read from a sparse-mode (PSM 11) pass.
    function recorderFillTableValuesFromSparse(entries, meta, sparseWords) {
      const { valueX, nameStop, medH } = meta;
      for (const e of entries) {
        if (e.value != null) continue;
        let best = null;
        for (const w of sparseWords) {
          const yc = (w.y0 + w.y1) / 2;
          if (Math.abs(yc - e.y) > medH * 0.75) continue;
          const xc = (w.x0 + w.x1) / 2;
          if (valueX != null ? xc < valueX - 10 : w.x0 <= nameStop) continue;
          const n = recorderValueFromWord(w, valueX, valueX != null ? 6 : 4);
          if (n == null) continue;
          if (!best || w.x0 > best.w.x0) best = { n, w };
        }
        if (best) { e.value = best.n; e.conf = Math.round(Math.min(e.conf, best.w.conf)); }
      }
    }
    const RECORDER_CARD_LABELS = [
      ['current', 'inbound', 'Current Inbound'],
      ['max', 'inbound', 'Max Inbound'],
      ['current', 'outbound', 'Current Outbound'],
      ['max', 'outbound', 'Max Outbound'],
    ];
    function recorderParseCards(words) {
      const { rows } = recorderClusterRows(words);
      const found = [];
      for (const r of (rows || [])) {
        for (let i = 0; i < r.words.length; i++) {
          const a = r.words[i], b = r.words[i + 1];
          for (const [w1, w2, label] of RECORDER_CARD_LABELS) {
            if (found.some(f => f.label === label)) continue;
            if (b && recorderWordLike(a.text, w1) && recorderWordLike(b.text, w2)) {
              found.push({ label, x0: a.x0, x1: b.x1, y0: Math.min(a.y0, b.y0), conf: Math.min(a.conf, b.conf) });
            } else if (recorderWordLike(a.text, w1 + w2)) {
              found.push({ label, x0: a.x0, x1: a.x1, y0: a.y0, conf: a.conf });
            }
          }
        }
      }
      const entries = [];
      for (const f of found) {
        const width = Math.max(40, f.x1 - f.x0);
        const cands = words
          .filter(w => w.y1 <= f.y0 + 2)
          .filter(w => {
            const xc = (w.x0 + w.x1) / 2;
            return xc >= f.x0 - width * 0.8 && xc <= f.x1 + width * 0.8;
          })
          .sort((a, b) => (f.y0 - a.y1) - (f.y0 - b.y1));
        let value = null, conf = f.conf;
        for (const c of cands) {
          const n = recorderNumFromToken(c.text, 6);
          if (n != null) { value = n; conf = Math.min(conf, c.conf); break; }
        }
        entries.push({ name: f.label, value, conf: Math.round(conf), y: f.y0 });
      }
      entries.sort((a, b) => a.y - b.y || a.name.localeCompare(b.name));
      return entries;
    }
    function recorderClassifyWords(words) {
      let gateway = false, cardWords = 0;
      for (const w of words) {
        if (recorderWordLike(w.text, 'gateway') || recorderWordLike(w.text, 'gatewayname')) gateway = true;
        if (recorderWordLike(w.text, 'inbound') || recorderWordLike(w.text, 'outbound')
          || recorderWordLike(w.text, 'currentinbound') || recorderWordLike(w.text, 'maxinbound')
          || recorderWordLike(w.text, 'currentoutbound') || recorderWordLike(w.text, 'maxoutbound')) cardWords++;
      }
      if (gateway) return 'table';
      if (cardWords >= 1) return 'cards';
      return 'table';
    }

    /* ── Recorder: workbook build / import ── */
    const RECORDER_BORDER_COLOR = 'FF374151';
    function recorderThinBorder(sides) {
      const edge = { style: 'thin', color: { argb: RECORDER_BORDER_COLOR } };
      const b = {};
      for (const s of sides) b[s] = edge;
      return b;
    }
    function recorderSheetName(name) {
      const clean = String(name || 'Sheet').replace(/[\\/*?:\[\]]/g, ' ').trim().slice(0, 31);
      return clean || 'Sheet';
    }
    function recorderFileName(pattern, year, monthIdx) {
      let out = String(pattern || 'Record_{Month}_{Year}')
        .replace(/\{Month\}/g, RECORDER_MONTHS[monthIdx] || '')
        .replace(/\{Year\}/g, String(year));
      if (!/\.xlsx$/i.test(out)) out += '.xlsx';
      return out;
    }
    function recorderBuildWorkbook(rec, year, monthIdx) {
      const wb = new ExcelJS.Workbook();
      wb.creator = 'Generator App — Recorder';
      wb.created = new Date();
      const days = new Date(year, monthIdx + 1, 0).getDate();
      const headerFont = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      const baseFont = { name: 'Calibri', size: 11 };
      const boldFont = { name: 'Calibri', size: 11, bold: true };
      const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
      for (const client of rec.clients) {
        if (!client.rows.length) continue;
        const ws = wb.addWorksheet(recorderSheetName(client.name), {
          views: [{ state: 'frozen', xSplit: 2, ySplit: 1, topLeftCell: 'C2', activePane: 'bottomRight' }],
        });
        ws.getColumn(1).width = 14;
        ws.getColumn(2).width = 34;
        for (let c = 3; c <= 26; c++) ws.getColumn(c).width = 8;
        const headerBorder = client.headerBorder ? recorderThinBorder(['top', 'bottom']) : null;
        for (let c = 1; c <= 26; c++) {
          const cell = ws.getCell(1, c);
          cell.value = c === 1 ? 'DATE' : c === 2 ? '' : HOURS_24[c - 3];
          cell.font = headerFont;
          cell.fill = headerFill;
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          if (headerBorder) cell.border = headerBorder;
        }
        const byDate = (rec.values[client.id]) || {};
        let r = 2;
        for (let day = 1; day <= days; day++) {
          const iso = `${year}-${String(monthIdx + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
          const byRow = byDate[iso] || {};
          const blockLen = client.rows.length;
          client.rows.forEach((row, i) => {
            const rr = r + i;
            const aCell = ws.getCell(rr, 1);
            if (i === 0) {
              aCell.value = new Date(Date.UTC(year, monthIdx, day));
              aCell.numFmt = 'mmm d, yyyy';
              aCell.font = boldFont;
              aCell.alignment = { horizontal: 'center', vertical: 'middle' };
            }
            if (client.bordered) {
              const sides = ['left', 'right'];
              if (i === 0) sides.push('top');
              if (i === blockLen - 1) sides.push('bottom');
              aCell.border = recorderThinBorder(sides);
            }
            const bCell = ws.getCell(rr, 2);
            bCell.value = row.label;
            bCell.font = baseFont;
            bCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argbHex(row.color) } };
            bCell.alignment = { horizontal: 'left', vertical: 'middle' };
            if (client.bordered) bCell.border = recorderThinBorder(['left', 'right', 'top', 'bottom']);
            const byHour = byRow[row.id] || {};
            for (let h = 0; h < 24; h++) {
              const cell = ws.getCell(rr, 3 + h);
              const v = byHour[h];
              if (v != null) cell.value = v;
              cell.font = baseFont;
              cell.alignment = { horizontal: 'center' };
              if (client.bordered) cell.border = recorderThinBorder(['left', 'right', 'top', 'bottom']);
            }
          });
          const sepRow = r + blockLen;
          const sepFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argbHex(client.sepColor) } };
          for (let c = 1; c <= 26; c++) ws.getCell(sepRow, c).fill = sepFill;
          r = sepRow + 1;
        }
        // "Highlight max value in row" — same formula the SIP/FCS builder
        // writes (Google-Sheets-compatible). One relative-row rule covers the
        // whole grid; ISNUMBER skips date/separator/blank rows.
        if (rec.maxInRow && rec.maxInRow.enabled && r > 2) {
          ws.addConditionalFormatting({
            ref: `C2:Z${r - 2}`,
            rules: [{
              type: 'expression',
              formulae: ['AND(ISNUMBER(C2),C2=MAX($C2:$Z2))'],
              style: {
                fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: argbHex(rec.maxInRow.color) } },
                font: { bold: !!rec.maxInRow.bold, color: { argb: argbHex(rec.maxInRow.fontColor) } },
              },
              priority: 10,
            }],
          });
        }
      }
      return wb;
    }
    // Google Sheets sync builder — reproduces exactly what the Download
    // button makes for the month/year picked in the Recorder sidebar.
    async function buildRecorderBlob(state) {
      const rec = recorderNormalizeState(state.recorder);
      const year = recorderParseIntOr(rec.exportYear, new Date().getFullYear());
      const wb = recorderBuildWorkbook(rec, year, rec.exportMonth);
      const buf = await wb.xlsx.writeBuffer();
      const filename = recorderFileName(rec.filePattern, year, rec.exportMonth);
      return { blob: new Blob([buf], { type: XLSX_MIME }), filename, sheets: collectSheetGridDims(wb) };
    }
    function recorderCellText(v) {
      if (v == null) return '';
      if (typeof v === 'object') {
        if (typeof v.richText !== 'undefined') return (v.richText || []).map(t => t.text).join('');
        if (typeof v.text !== 'undefined') return String(v.text);
        if (typeof v.result !== 'undefined') return String(v.result);
        return '';
      }
      return String(v);
    }
    function recorderCoerceDateIso(v) {
      if (v instanceof Date && !isNaN(v)) {
        return `${v.getUTCFullYear()}-${String(v.getUTCMonth() + 1).padStart(2, '0')}-${String(v.getUTCDate()).padStart(2, '0')}`;
      }
      const n = Number(v);
      if (Number.isFinite(n) && n > 20000 && n < 80000) {
        const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
        return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
      }
      return null;
    }
    function recorderCoerceNumber(v) {
      if (v == null || v === '') return null;
      if (typeof v === 'object') {
        if (typeof v.result !== 'undefined') return recorderCoerceNumber(v.result);
        return null;
      }
      const n = Number(String(v).replace(/,/g, '').trim());
      return Number.isFinite(n) ? n : null;
    }
    // Merge an existing record workbook (same layout) into rec.values —
    // used to pick up a month that was partly recorded before this module.
    async function recorderImportWorkbook(rec, arrayBuffer) {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(arrayBuffer);
      const values = JSON.parse(JSON.stringify(rec.values || {}));
      let count = 0, sheets = 0;
      wb.eachSheet(ws => {
        const client = rec.clients.find(c => recorderSig(c.name) === recorderSig(ws.name));
        if (!client) return;
        sheets++;
        let curDate = null;
        ws.eachRow(row => {
          const iso = recorderCoerceDateIso(row.getCell(1).value);
          if (iso) curDate = iso;
          if (!curDate) return;
          const label = recorderCellText(row.getCell(2).value).trim();
          if (!label) return;
          const target = recorderMatchClientRow(label, client);
          if (!target) return;
          for (let h = 0; h < 24; h++) {
            const num = recorderCoerceNumber(row.getCell(3 + h).value);
            if (num == null) continue;
            const cv = values[client.id] = values[client.id] || {};
            const dv = cv[curDate] = cv[curDate] || {};
            const rv = dv[target.id] = dv[target.id] || {};
            rv[h] = num;
            count++;
          }
        });
      });
      return { values, count, sheets };
    }

    /* ── Recorder: shared UI helpers ── */
    const recorderTodayIso = () => {
      const d = new Date();
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    };
    function recorderPrettyDate(iso) {
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
      if (!m) return String(iso || '');
      return `${RECORDER_MONTHS[Number(m[2]) - 1]} ${Number(m[3])}`;
    }
    function recorderApplyWrites(values, writes, overwrite) {
      let written = 0, skipped = 0;
      const next = { ...values };
      for (const wr of writes) {
        const cv = next[wr.clientId] = { ...(next[wr.clientId] || {}) };
        const dv = cv[wr.dateIso] = { ...(cv[wr.dateIso] || {}) };
        const rv = dv[wr.rowId] = { ...(dv[wr.rowId] || {}) };
        if (wr.value == null) {
          if (rv[wr.hour] != null) { delete rv[wr.hour]; written++; }
        } else if (!overwrite && rv[wr.hour] != null) {
          skipped++;
        } else {
          rv[wr.hour] = wr.value; written++;
        }
        if (!Object.keys(rv).length) delete dv[wr.rowId];
        if (!Object.keys(dv).length) delete cv[wr.dateIso];
        if (!Object.keys(cv).length) delete next[wr.clientId];
      }
      return { next, written, skipped };
    }
    function recorderSlotFilled(rec, client, target) {
      const byRow = ((rec.values[client.id] || {})[target.date]) || {};
      return client.rows.some(r => (byRow[r.id] || {})[target.hour] != null);
    }
    function recorderSuggestCardsClient(rec, target, excludeIds) {
      const cards = rec.clients.filter(c => c.type === 'cards');
      const free = cards.find(c => !excludeIds.includes(c.id) && !recorderSlotFilled(rec, c, target));
      return free ? free.id : (cards[0] ? cards[0].id : null);
    }
    function recorderRowsFromEntries(entries, kind, rec, clientId) {
      const rows = [], ignored = [];
      for (const e of entries) {
        let match = null;
        if (kind === 'cards') {
          const client = rec.clients.find(c => c.id === clientId);
          const row = client ? recorderMatchClientRow(e.name, client) : null;
          if (row) match = { clientId, rowId: row.id, label: row.label, exact: true };
        } else {
          match = recorderMatchGatewayRow(e.name, rec.clients, rec.match);
        }
        if (match) {
          rows.push({
            key: recorderUid(), clientId: match.clientId, rowId: match.rowId, label: match.label,
            ocrName: e.name, value: e.value == null ? '' : String(e.value), conf: e.conf, exact: !!match.exact,
          });
        } else {
          ignored.push(`${e.name}${e.value != null ? ' = ' + e.value : ''}`);
        }
      }
      return { rows, ignored };
    }
    function recorderParseIntOr(raw, fallback) {
      const n = parseInt(String(raw), 10);
      return Number.isFinite(n) ? n : fallback;
    }

    /* ── Recorder: components ── */
    function RecorderDropZone({ onFiles, engine }) {
      const fileRef = useRef(null);
      const [over, setOver] = useState(false);
      return (
        <div
          onDragOver={e => { e.preventDefault(); setOver(true); }}
          onDragLeave={() => setOver(false)}
          onDrop={e => {
            e.preventDefault(); setOver(false);
            const files = [...((e.dataTransfer && e.dataTransfer.files) || [])].filter(f => String(f.type || '').startsWith('image/'));
            if (files.length) onFiles(files);
          }}
          className={`rounded-lg border-2 border-dashed px-6 py-9 text-center transition-colors ${over ? 'border-blue-500/70 bg-blue-500/5' : 'border-neutral-800 bg-neutral-950/60'}`}>
          <div className="flex items-center justify-center gap-2 text-sm text-neutral-300 font-medium">
            <kbd className="rounded border border-neutral-700 bg-neutral-900 px-1.5 py-0.5 font-mono text-[10px] text-neutral-300">Ctrl</kbd>
            <span className="text-neutral-600">+</span>
            <kbd className="rounded border border-neutral-700 bg-neutral-900 px-1.5 py-0.5 font-mono text-[10px] text-neutral-300">V</kbd>
            <span>paste a screenshot — or drop image files here</span>
          </div>
          <div className="mt-1.5 text-xs text-neutral-600">Gateway tables and Current/Max Inbound-Outbound cards are detected automatically. Queue as many screenshots as you need for the hour.</div>
          <div className="mt-4 flex items-center justify-center">
            <input ref={fileRef} type="file" accept="image/*" multiple className="hidden"
              onChange={e => { const files = [...(e.target.files || [])]; if (files.length) onFiles(files); e.target.value = ''; }} />
            <Btn variant="ghost" size="md" onClick={() => fileRef.current && fileRef.current.click()}><IconUpload /> Choose images</Btn>
          </div>
          {(engine.status === 'loading' || engine.status === 'working') && (
            <div className="mt-3 text-[11px] text-blue-300">{engine.label || 'Working…'}</div>
          )}
        </div>
      );
    }

    function RecorderCaptureCard({ cap, rec, target, onAssign, onValue, onCommit, onDiscard }) {
      const cardsClients = rec.clients.filter(c => c.type === 'cards');
      const [zoomed, setZoomed] = useState(false);
      useEffect(() => {
        if (!zoomed) return;
        const onKey = e => { if (e.key === 'Escape') setZoomed(false); };
        window.addEventListener('keydown', onKey);
        return () => window.removeEventListener('keydown', onKey);
      }, [zoomed]);
      const existing = (clientId, rowId) => ((((rec.values[clientId] || {})[target.date]) || {})[rowId] || {})[target.hour];
      return (
        <div className="rounded-lg border border-neutral-900 bg-neutral-950 p-4">
          <div className="flex flex-wrap items-start gap-4">
            <div className="w-[300px] max-w-full shrink-0">
              {cap.thumb
                ? <img src={cap.thumb} alt="" title="Click to enlarge" onClick={() => setZoomed(true)}
                    className="w-full cursor-zoom-in rounded border border-neutral-800" />
                : <div className="h-20 rounded border border-neutral-800 bg-neutral-900 animate-pulse" />}
              {zoomed && cap.thumb && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4 anim-overlay cursor-zoom-out"
                  onClick={() => setZoomed(false)}>
                  <img src={cap.thumb} alt="" className="max-h-[92vh] max-w-[96vw] rounded border border-neutral-700 object-contain shadow-2xl" />
                </div>
              )}
              <div className="mt-2 flex flex-wrap gap-1.5">
                {cap.status === 'ocr' && <Pill tone="accent">Reading…</Pill>}
                {cap.status === 'ready' && <Pill tone="accent">{cap.kind === 'cards' ? 'Cards' : 'Gateway table'}</Pill>}
                {cap.status === 'ready' && <Pill tone={cap.rows.length ? 'success' : 'muted'}>{cap.rows.length} matched</Pill>}
                {cap.status === 'empty' && <Pill tone="muted">Nothing readable</Pill>}
                {cap.status === 'error' && <Pill tone="muted">Failed</Pill>}
              </div>
            </div>
            <div className="flex-1 min-w-[260px]">
              {cap.status === 'error' && <p className="text-xs text-red-300">{cap.error}</p>}
              {cap.status === 'ocr' && <p className="text-xs text-neutral-500">Running OCR…</p>}
              {cap.status === 'empty' && <p className="text-xs text-neutral-500">No values were recognized. Try a tighter screenshot of the table, or type the numbers in the “Still missing” panel below.</p>}
              {cap.status === 'ready' && (
                <>
                  {cap.kind === 'cards' && (
                    <div className="mb-3 flex flex-wrap items-center gap-2">
                      <span className="text-[10px] uppercase tracking-wide text-neutral-500">Client</span>
                      <Select value={cap.clientId || ''} onChange={e => onAssign(cap.id, e.target.value)} className="!w-auto min-w-[180px]">
                        {cardsClients.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                      </Select>
                      <span className="text-[10px] text-amber-300">check this — card screenshots look alike</span>
                    </div>
                  )}
                  {cap.rows.length > 0 && (
                    <div className="grid gap-1.5 sm:grid-cols-2">
                      {cap.rows.map(row => {
                        const prev = existing(row.clientId, row.rowId);
                        const clientName = (rec.clients.find(c => c.id === row.clientId) || {}).name || '';
                        const flagged = row.conf < 70 || !row.exact;
                        return (
                          <div key={row.key} className={`flex items-center gap-2 rounded-md border px-2.5 py-1.5 ${flagged ? 'border-amber-500/40 bg-amber-500/5' : 'border-neutral-900 bg-neutral-900/50'}`}>
                            <div className="flex-1 min-w-0">
                              <div className="truncate text-xs text-neutral-200" title={`${clientName} · OCR read: ${row.ocrName}`}>{row.label}</div>
                              <div className="truncate text-[10px] text-neutral-500">
                                {clientName}
                                {!row.exact ? ' · fuzzy match' : ''}
                                {row.conf < 70 ? ` · low confidence` : ''}
                                {prev != null ? ` · overwrites ${prev}` : ''}
                              </div>
                            </div>
                            <input value={row.value} onChange={e => onValue(cap.id, row.key, e.target.value.replace(/[^0-9]/g, ''))}
                              inputMode="numeric" placeholder="—"
                              className="w-14 shrink-0 rounded border border-neutral-800 bg-neutral-950 px-1.5 py-1 text-center text-sm text-neutral-100 outline-none focus:border-blue-500/60" />
                          </div>
                        );
                      })}
                    </div>
                  )}
                  {cap.rows.length === 0 && (
                    <p className="text-xs text-amber-300">Text was read but none of it matches the watchlist. Check the Config tab (names/aliases) if this screenshot should be recorded.</p>
                  )}
                  {cap.ignored.length > 0 && (
                    <div className="mt-2 text-[10px] leading-relaxed text-neutral-600">
                      Ignored (not watched): {cap.ignored.join(' · ')}
                    </div>
                  )}
                </>
              )}
            </div>
            <div className="flex shrink-0 flex-col gap-2">
              {cap.status === 'ready' && cap.rows.length > 0 && <Btn variant="primary" size="md" onClick={onCommit}>Save</Btn>}
              <Btn variant="ghost" size="md" onClick={onDiscard}><IconX /> Discard</Btn>
            </div>
          </div>
        </div>
      );
    }

    function RecorderMissingPanel({ missing, manual, setManual, hourLabel }) {
      if (!missing.length) {
        return (
          <div className="rounded-lg border border-emerald-500/30 bg-emerald-500/5 px-4 py-3 text-xs text-emerald-300">
            Every watched row has a value for this slot (saved or pending).
          </div>
        );
      }
      return (
        <div className="rounded-lg border border-neutral-900 bg-neutral-950 p-4">
          <SectionLabel hint="type anything the screenshots did not cover">Still missing at {hourLabel}</SectionLabel>
          <div className="space-y-3">
            {missing.map(group => (
              <div key={group.client.id}>
                <div className="mb-1.5 text-[10px] uppercase tracking-wide text-neutral-500">{group.client.name}</div>
                <div className="grid gap-1.5 sm:grid-cols-2 lg:grid-cols-3">
                  {group.rows.map(row => {
                    const k = group.client.id + '|' + row.id;
                    return (
                      <div key={row.id} className="flex items-center gap-2 rounded-md border border-neutral-900 bg-neutral-900/40 px-2.5 py-1.5">
                        <div className="flex-1 truncate text-xs text-neutral-300" title={row.label}>{row.label}</div>
                        <input value={manual[k] || ''} onChange={e => { const v = e.target.value.replace(/[^0-9]/g, ''); setManual(m => ({ ...m, [k]: v })); }}
                          inputMode="numeric" placeholder="—"
                          className="w-14 shrink-0 rounded border border-neutral-800 bg-neutral-950 px-1.5 py-1 text-center text-sm text-neutral-100 outline-none focus:border-blue-500/60" />
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        </div>
      );
    }

    function RecorderDataTab({ rec, setRec, onToast, confirmDialog, onImport, onVisibleMonth }) {
      const [clientSel, setClientSel] = useState(() => (rec.clients[0] ? rec.clients[0].id : null));
      const [date, setDate] = useState(recorderTodayIso());
      const importRef = useRef(null);
      const monthPrefix = date.slice(0, 7);
      // Tell the module which month is on screen so the shared team sync can
      // pull it from Firestore the first time it's viewed this session.
      useEffect(() => { if (onVisibleMonth) onVisibleMonth(monthPrefix); }, [monthPrefix, onVisibleMonth]);
      const client = rec.clients.find(c => c.id === clientSel) || rec.clients[0];
      if (!client) return <p className="text-sm text-neutral-500">No clients configured — add one in the Config tab.</p>;
      const [yy, mm] = monthPrefix.split('-').map(Number);
      const daysInMonth = new Date(yy, mm, 0).getDate();
      const byDate = rec.values[client.id] || {};
      const byRow = byDate[date] || {};
      const today = recorderTodayIso();
      const curHour = new Date().getHours();
      const setCell = (rowId, hour, raw) => {
        const t = String(raw).trim();
        const v = t === '' ? null : Math.min(recorderParseIntOr(t, 0), rec.clampMax);
        setRec(r => ({ ...r, values: recorderApplyWrites(r.values, [{ clientId: client.id, dateIso: date, rowId, hour, value: v }], true).next }));
      };
      const shiftDay = (delta) => {
        const d = new Date(yy, mm - 1, Number(date.slice(8, 10)) + delta);
        setDate(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`);
      };
      const clearScope = async (scope) => {
        const ok = await confirmDialog({
          title: scope === 'day' ? `Clear ${client.name} — ${recorderPrettyDate(date)}?` : `Clear ${client.name} — entire ${RECORDER_MONTHS[mm - 1]} ${yy}?`,
          message: 'The recorded values in this range will be removed. This cannot be undone.',
          confirmText: 'Clear', tone: 'danger',
        });
        if (!ok) return;
        setRec(r => {
          const cv = { ...(r.values[client.id] || {}) };
          if (scope === 'day') delete cv[date];
          else for (const k of Object.keys(cv)) if (k.startsWith(monthPrefix + '-')) delete cv[k];
          const values = { ...r.values };
          if (Object.keys(cv).length) values[client.id] = cv; else delete values[client.id];
          return { ...r, values };
        });
        onToast('ok', scope === 'day' ? 'Day cleared' : 'Month cleared');
      };
      return (
        <div className="max-w-full space-y-4 anim-fade-in">
          <div className="flex flex-wrap items-center gap-2">
            {rec.clients.map(c => (
              <button key={c.id} onClick={() => setClientSel(c.id)}
                className={`rounded-md border px-3 py-1.5 text-xs font-medium transition-colors ${c.id === client.id ? 'bg-white text-black border-white' : 'border-neutral-800 text-neutral-400 hover:text-neutral-100 hover:border-neutral-600'}`}>
                {c.name}
              </button>
            ))}
            <div className="ml-auto flex items-center gap-2">
              <input ref={importRef} type="file" accept=".xlsx" className="hidden"
                onChange={e => { const f = e.target.files && e.target.files[0]; if (f) onImport(f); e.target.value = ''; }} />
              <Btn variant="ghost" size="md" onClick={() => importRef.current && importRef.current.click()} title="Merge an existing record workbook (same tab layout) into the saved values">
                <IconUpload /> Import month
              </Btn>
              <Btn variant="ghost" size="md" onClick={() => clearScope('day')}>Clear day</Btn>
              <Btn variant="danger" size="md" onClick={() => clearScope('month')}>Clear month</Btn>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <Btn variant="ghost" size="md" onClick={() => shiftDay(-1)}>‹</Btn>
            <Input type="date" value={date} onChange={e => e.target.value && setDate(e.target.value)} className="!w-auto" />
            <Btn variant="ghost" size="md" onClick={() => shiftDay(1)}>›</Btn>
            <span className="text-[11px] text-neutral-600">cells are editable — blank means not recorded</span>
          </div>
          <div className="flex flex-wrap gap-1">
            {Array.from({ length: daysInMonth }, (_, i) => {
              const dIso = `${monthPrefix}-${String(i + 1).padStart(2, '0')}`;
              const dRows = byDate[dIso] || {};
              let filled = 0;
              for (const r of client.rows) filled += Object.keys(dRows[r.id] || {}).length;
              const total = client.rows.length * 24;
              // text-neutral-950 (not text-black): html.light inverts .text-black
              // to near-white for primary buttons, which would blank these chips.
              const cls = filled === 0 ? 'bg-neutral-900 text-neutral-500' : filled >= total ? 'bg-emerald-500/70 text-neutral-950' : 'bg-blue-500/40 text-blue-200';
              return (
                <button key={dIso} onClick={() => setDate(dIso)} title={`${dIso} — ${filled}/${total} cells`}
                  className={`h-6 w-7 rounded text-[10px] font-medium transition-transform hover:scale-105 ${cls} ${dIso === date ? 'ring-1 ring-white/70' : ''}`}>
                  {i + 1}
                </button>
              );
            })}
          </div>
          <div className="overflow-x-auto rounded-lg border border-neutral-900">
            <table className="min-w-full border-collapse text-xs">
              <thead>
                <tr>
                  <th className="sticky left-0 z-10 border-b border-neutral-900 bg-neutral-950 px-3 py-2 text-left font-semibold text-neutral-400">{recorderPrettyDate(date)}</th>
                  {HOURS_24.map((h, i) => (
                    <th key={i} className={`border-b border-neutral-900 bg-neutral-950 px-1 py-2 text-center font-medium ${i === curHour && date === today ? 'text-blue-300' : 'text-neutral-500'}`}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {client.rows.map(row => {
                  const rowVals = byRow[row.id] || {};
                  // Live preview of the "highest value in a row" rule — same
                  // semantics as the exported conditional formatting (ties all
                  // highlight, single lone value counts as the max).
                  const rowMax = rec.maxInRow.enabled
                    ? Object.values(rowVals).reduce((m, x) => (x != null && (m == null || x > m)) ? x : m, null)
                    : null;
                  return (
                    <tr key={row.id} className="odd:bg-neutral-900/30">
                      <td className="sticky left-0 z-10 whitespace-nowrap border-b border-neutral-900/60 bg-[#1c1c1f] px-3 py-1">
                        <span className="mr-2 inline-block h-2 w-2 rounded-sm align-middle" style={{ background: row.color }} />
                        <span className="align-middle text-neutral-200">{row.label}</span>
                      </td>
                      {HOURS_24.map((_, h) => {
                        const v = rowVals[h];
                        const isMax = rowMax != null && v === rowMax;
                        return (
                          <td key={h} className="border-b border-neutral-900/60 p-0.5 text-center">
                            <input value={v == null ? '' : String(v)} onChange={e => setCell(row.id, h, e.target.value.replace(/[^0-9]/g, ''))}
                              inputMode="numeric"
                              style={isMax ? { background: rec.maxInRow.color, color: rec.maxInRow.fontColor, fontWeight: rec.maxInRow.bold ? 700 : 400 } : undefined}
                              className={`w-10 rounded bg-transparent px-0.5 py-1 text-center outline-none focus:bg-neutral-800 ${v != null ? 'text-neutral-100' : 'text-neutral-600'}`} />
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      );
    }

    function RecorderConfigTab({ rec, setRec, confirmDialog, onToast }) {
      const [selId, setSelId] = useState(() => (rec.clients[0] ? rec.clients[0].id : null));
      const client = rec.clients.find(c => c.id === selId) || rec.clients[0];
      const patchClient = (id, patch) => setRec(r => ({ ...r, clients: r.clients.map(c => c.id === id ? { ...c, ...patch } : c) }));
      const patchRow = (cid, rid, patch) => setRec(r => ({
        ...r,
        clients: r.clients.map(c => c.id !== cid ? c : { ...c, rows: c.rows.map(rw => rw.id === rid ? { ...rw, ...patch } : rw) }),
      }));
      const moveClient = (id, dir) => setRec(r => {
        const i = r.clients.findIndex(c => c.id === id), j = i + dir;
        if (i < 0 || j < 0 || j >= r.clients.length) return r;
        const clients = [...r.clients];
        [clients[i], clients[j]] = [clients[j], clients[i]];
        return { ...r, clients };
      });
      const moveRow = (cid, rid, dir) => setRec(r => ({
        ...r,
        clients: r.clients.map(c => {
          if (c.id !== cid) return c;
          const i = c.rows.findIndex(rw => rw.id === rid), j = i + dir;
          if (i < 0 || j < 0 || j >= c.rows.length) return c;
          const rows = [...c.rows];
          [rows[i], rows[j]] = [rows[j], rows[i]];
          return { ...c, rows };
        }),
      }));
      const addClient = () => {
        const id = recorderUid();
        setRec(r => ({ ...r, clients: [...r.clients, { id, name: 'NEW CLIENT', type: 'gateway', bordered: true, headerBorder: true, sepColor: '#000000', rows: [] }] }));
        setSelId(id);
      };
      const removeClient = async (c) => {
        const ok = await confirmDialog({
          title: `Delete client “${c.name}”?`,
          message: 'Its watch rows and its recorded values will be removed from the Recorder.',
          confirmText: 'Delete', tone: 'danger',
        });
        if (!ok) return;
        setRec(r => {
          const values = { ...r.values };
          delete values[c.id];
          return { ...r, clients: r.clients.filter(x => x.id !== c.id), values };
        });
      };
      const addRow = () => setRec(r => ({
        ...r,
        clients: r.clients.map(c => c.id !== client.id ? c : { ...c, rows: [...c.rows, { id: recorderUid(), label: 'NEW ROW', color: '#FFFF00', aliases: [] }] }),
      }));
      const removeRow = (rid) => setRec(r => ({
        ...r,
        clients: r.clients.map(c => c.id !== client.id ? c : { ...c, rows: c.rows.filter(rw => rw.id !== rid) }),
      }));
      const restoreDefaults = async () => {
        const ok = await confirmDialog({
          title: 'Restore the default clients & rows?',
          message: 'The client list goes back to the built-in 5-tab setup (VOS, VOS TWILIO, KINGSFORD INOUT, UNOBANK IN, MYVELOX INOUT). Recorded values are kept.',
          confirmText: 'Restore', tone: 'danger',
        });
        if (!ok) return;
        setRec(r => ({ ...r, clients: RECORDER_DEFAULT_CLIENTS.map(c => ({ ...c, rows: c.rows.map(rw => ({ ...rw, aliases: [...rw.aliases] })) })) }));
        onToast('ok', 'Default clients restored');
      };
      if (!client) {
        return (
          <div className="anim-fade-in">
            <Btn variant="primary" size="md" onClick={addClient}><IconPlus /> Add client</Btn>
          </div>
        );
      }
      return (
        <div className="grid max-w-6xl gap-6 lg:grid-cols-[280px,1fr] anim-fade-in">
          <div>
            <SectionLabel hint="workbook tab order">Clients</SectionLabel>
            <div className="space-y-1.5">
              {rec.clients.map((c, i) => (
                <div key={c.id} className={`flex items-center gap-1.5 rounded-md border px-2 py-1.5 ${c.id === client.id ? 'border-blue-500/50 bg-blue-500/5' : 'border-neutral-900 bg-neutral-950'}`}>
                  <button onClick={() => setSelId(c.id)} className="min-w-0 flex-1 truncate text-left text-xs text-neutral-200" title={c.name}>{c.name}</button>
                  <Pill tone="muted">{c.type === 'cards' ? 'cards' : 'table'}</Pill>
                  <button onClick={() => moveClient(c.id, -1)} disabled={i === 0} className="text-neutral-600 hover:text-neutral-200 disabled:opacity-30"><IconUp /></button>
                  <button onClick={() => moveClient(c.id, 1)} disabled={i === rec.clients.length - 1} className="text-neutral-600 hover:text-neutral-200 disabled:opacity-30"><IconDown /></button>
                  <button onClick={() => removeClient(c)} className="text-neutral-600 hover:text-red-300" title="Delete client"><IconX /></button>
                </div>
              ))}
            </div>
            <div className="mt-3 space-y-2">
              <Btn variant="ghost" size="md" className="w-full" onClick={addClient}><IconPlus /> Add client</Btn>
              <Btn variant="ghost" size="md" className="w-full" onClick={restoreDefaults}><IconReset /> Restore defaults</Btn>
            </div>
          </div>
          <div className="min-w-0 space-y-5">
            <div>
              <SectionLabel>Client settings</SectionLabel>
              <div className="grid gap-3 rounded-lg border border-neutral-900 bg-neutral-950 p-4 sm:grid-cols-2">
                <div>
                  <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Sheet / tab name</label>
                  <Input value={client.name} onChange={e => patchClient(client.id, { name: e.target.value })} />
                </div>
                <div>
                  <label className="mb-1 block text-[10px] uppercase tracking-wide text-neutral-500">Screenshot type</label>
                  <Select value={client.type} onChange={e => patchClient(client.id, { type: e.target.value })}>
                    <option value="gateway">Gateway table (name + Number of calling)</option>
                    <option value="cards">Metric cards (Current/Max In-Outbound)</option>
                  </Select>
                </div>
                <div className="flex items-center gap-4">
                  <label className="flex items-center gap-2 text-xs text-neutral-300">
                    <input type="checkbox" className="h-4 w-4 accent-blue-500" checked={client.bordered} onChange={e => patchClient(client.id, { bordered: e.target.checked })} />
                    Cell borders in the export
                  </label>
                  <label className="flex items-center gap-2 text-xs text-neutral-300">
                    <input type="checkbox" className="h-4 w-4 accent-blue-500" checked={client.headerBorder} onChange={e => patchClient(client.id, { headerBorder: e.target.checked })} />
                    Header border
                  </label>
                </div>
                <div className="flex items-center gap-3">
                  <span className="text-[10px] uppercase tracking-wide text-neutral-500">Separator row</span>
                  <PresetColorPicker value={client.sepColor} onChange={hex => patchClient(client.id, { sepColor: hex })} title="Separator row color" />
                </div>
              </div>
            </div>
            <div>
              <SectionLabel hint="order = row order inside each day block">Watched rows</SectionLabel>
              <div className="space-y-1.5">
                {client.rows.map((row, i) => (
                  <div key={row.id} className="flex flex-wrap items-center gap-2 rounded-md border border-neutral-900 bg-neutral-950 px-2.5 py-2">
                    <PresetColorPicker value={row.color} onChange={hex => patchRow(client.id, row.id, { color: hex })} title="Name cell color" showHex={false} />
                    <Input value={row.label} onChange={e => patchRow(client.id, row.id, { label: e.target.value })} className="!w-64 font-mono !text-xs" />
                    <Input key={row.id + '_al'} defaultValue={row.aliases.join(', ')}
                      onBlur={e => patchRow(client.id, row.id, { aliases: e.target.value.split(',').map(s => s.trim()).filter(Boolean) })}
                      placeholder="aliases (comma separated) — extra names OCR may read"
                      className="!w-72 flex-1 !text-xs" />
                    <div className="ml-auto flex items-center gap-1.5">
                      <button onClick={() => moveRow(client.id, row.id, -1)} disabled={i === 0} className="text-neutral-600 hover:text-neutral-200 disabled:opacity-30"><IconUp /></button>
                      <button onClick={() => moveRow(client.id, row.id, 1)} disabled={i === client.rows.length - 1} className="text-neutral-600 hover:text-neutral-200 disabled:opacity-30"><IconDown /></button>
                      <button onClick={() => removeRow(row.id)} className="text-neutral-600 hover:text-red-300" title="Remove row"><IconX /></button>
                    </div>
                  </div>
                ))}
              </div>
              <div className="mt-2">
                <Btn variant="ghost" size="md" onClick={addRow}><IconPlus /> Add row</Btn>
              </div>
            </div>
          </div>
        </div>
      );
    }

    function RecorderRulesTab({ rec, setRec, engine, onPreload }) {
      const patch = (p) => setRec(r => ({ ...r, ...p }));
      const Row = ({ label, hint, children }) => (
        <div className="flex flex-wrap items-center justify-between gap-3 rounded-md border border-neutral-900 bg-neutral-950 px-4 py-3">
          <div className="min-w-[240px] flex-1">
            <div className="text-sm text-neutral-200">{label}</div>
            {hint && <div className="mt-0.5 text-[11px] text-neutral-600">{hint}</div>}
          </div>
          <div className="flex items-center gap-2">{children}</div>
        </div>
      );
      return (
        <div className="max-w-3xl space-y-2.5 anim-fade-in">
          <Row label="Name matching" hint="How OCR text is matched to watched rows. Exact matches always win; this only sets how much OCR misreading is tolerated.">
            <Select value={rec.match} onChange={e => patch({ match: e.target.value })} className="!w-auto">
              <option value="strict">Strict — 1 misread character (recommended)</option>
              <option value="normal">Normal — 2 (may cross-match similar names)</option>
              <option value="loose">Loose — 3 (risky with sibling names)</option>
            </Select>
          </Row>
          <Row label="Overwrite on save" hint="When a cell already has a value for the slot, saving a capture replaces it. Off = existing values are kept.">
            <input type="checkbox" className="h-4 w-4 accent-blue-500" checked={rec.overwrite} onChange={e => patch({ overwrite: e.target.checked })} />
          </Row>
          <Row label="Upscale small screenshots" hint="2–3× enlargement before OCR. Helps with small UI text; turn off only if reading gets worse.">
            <input type="checkbox" className="h-4 w-4 accent-blue-500" checked={rec.upscale} onChange={e => patch({ upscale: e.target.checked })} />
          </Row>
          <Row label="Maximum value" hint="Readings above this are clamped — protects against a prefix number sneaking in as a value.">
            <Input type="number" value={rec.clampMax} onChange={e => patch({ clampMax: Math.max(1, recorderParseIntOr(e.target.value, 9999)) })} className="!w-28" />
          </Row>
          <Row label="Highlight highest value in a row" hint="Same rule as SIP/FCS — colors each row's maximum across the 24 hour columns, live in the Data grid and as conditional formatting in the exported workbook.">
            <span className="text-[10px] uppercase tracking-wide text-neutral-500">Fill</span>
            <PresetColorPicker value={rec.maxInRow.color} onChange={hex => patch({ maxInRow: { ...rec.maxInRow, color: hex } })} title="Highlight fill color" showHex={false} />
            <span className="text-[10px] uppercase tracking-wide text-neutral-500">Font</span>
            <PresetColorPicker value={rec.maxInRow.fontColor} onChange={hex => patch({ maxInRow: { ...rec.maxInRow, fontColor: hex } })} title="Highlight font color" showHex={false} />
            <label className={`inline-flex cursor-pointer items-center gap-1.5 text-xs ${rec.maxInRow.bold ? 'text-neutral-100' : 'text-neutral-500'}`} title="Bold the highlighted value">
              <input type="checkbox" className="h-3.5 w-3.5 accent-blue-500" checked={rec.maxInRow.bold} onChange={e => patch({ maxInRow: { ...rec.maxInRow, bold: e.target.checked } })} />
              <span className="font-bold">B</span>
            </label>
            <input type="checkbox" className="h-4 w-4 accent-blue-500" checked={rec.maxInRow.enabled} onChange={e => patch({ maxInRow: { ...rec.maxInRow, enabled: e.target.checked } })} />
          </Row>
          <Row label="Export file name" hint="{Month} and {Year} are replaced at download time.">
            <Input value={rec.filePattern} onChange={e => patch({ filePattern: e.target.value })} className="!w-80 font-mono !text-xs" />
          </Row>
          <Row label="OCR engine" hint="Tesseract.js — downloaded from CDN on first use (a few MB, then cached by the browser). Screenshots never leave this machine.">
            <Pill tone={engine.status === 'ready' ? 'success' : engine.status === 'idle' ? 'default' : 'accent'}>
              {engine.status === 'ready' ? 'Ready' : engine.status === 'idle' ? 'Not loaded' : (engine.label || 'Loading…')}
            </Pill>
            {engine.status !== 'ready' && <Btn variant="ghost" size="md" onClick={onPreload}>Load now</Btn>}
          </Row>
        </div>
      );
    }

    function RecorderSidebar({ state, setState, rec, sync, onRetrySync, target, setTarget, engine, slotStats, exportSel, setExportSel, onExport, busyExport, shared, onRefreshShared, gsheets }) {
      const theme = state.theme || 'dark';
      const toggleTheme = () => setState(s => ({ ...s, theme: (s.theme || 'dark') === 'dark' ? 'light' : 'dark' }));
      const exportYear = recorderParseIntOr(exportSel.year, new Date().getFullYear());
      return (
        <aside className="w-[320px] shrink-0 border-r border-neutral-900 bg-[#17171a] h-screen sticky top-0 flex flex-col">
          <div className="p-5 border-b border-neutral-900">
            <div className="flex items-center justify-between gap-2">
              <SyncBadge sync={sync} onRetry={onRetrySync} />
              <button onClick={toggleTheme}
                title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}
                className="inline-flex items-center gap-1.5 rounded-md border border-neutral-800 bg-neutral-950 px-2 py-1 text-[10px] uppercase tracking-wide text-neutral-400 hover:text-neutral-100 hover:border-neutral-700 transition-colors">
                {theme === 'dark' ? <IconSun /> : <IconMoon />}
                {theme === 'dark' ? 'Light' : 'Dark'}
              </button>
            </div>
            <h1 className="text-[17px] font-bold tracking-tight leading-tight mt-2">Recorder</h1>
            <p className="text-xs text-neutral-500 mt-1">Screenshots → hourly record workbook</p>
            <AccountChip sync={sync} />
          </div>
          <div className="p-5 flex-1 overflow-y-auto min-h-0 space-y-5">
            <div>
              <SectionLabel hint="every capture saves here">Recording slot</SectionLabel>
              <div className="rounded-md border border-neutral-900 bg-neutral-950 p-3 space-y-2">
                <Input type="date" value={target.date} onChange={e => { const v = e.target.value; if (v) setTarget(t => ({ ...t, date: v })); }} />
                <div className="flex gap-2">
                  <Select value={target.hour} onChange={e => setTarget(t => ({ ...t, hour: Number(e.target.value) }))}>
                    {HOURS_24.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </Select>
                  <Btn variant="ghost" size="md" onClick={() => setTarget({ date: recorderTodayIso(), hour: new Date().getHours() })}>Now</Btn>
                </div>
              </div>
            </div>
            <div>
              <SectionLabel hint="saved values at this slot">Coverage</SectionLabel>
              <div className="space-y-1.5">
                {slotStats.map(s => (
                  <div key={s.id} className="flex items-center justify-between gap-2 text-xs">
                    <span className="min-w-0 truncate text-neutral-400">{s.name}</span>
                    <Pill tone={s.filled >= s.total && s.total > 0 ? 'success' : s.filled ? 'accent' : 'default'}>{s.filled}/{s.total}</Pill>
                  </div>
                ))}
              </div>
            </div>
            <div>
              <SectionLabel hint="everyone signed in sees the same data">Team sync</SectionLabel>
              <div className="flex items-center justify-between gap-2">
                <Pill tone={shared.status === 'live' ? 'success' : shared.status === 'loading' ? 'accent' : 'muted'}>
                  {shared.status === 'live' ? 'Shared — live' : shared.status === 'loading' ? 'Refreshing…' : shared.status === 'error' ? 'Sync error' : 'Local only'}
                </Pill>
                <Btn variant="ghost" size="sm" onClick={onRefreshShared} disabled={shared.status === 'loading' || !sync.uid}><IconReset /> Refresh</Btn>
              </div>
              {shared.status === 'live' && shared.at > 0 && (
                <p className="mt-1.5 text-[10px] leading-relaxed text-neutral-600">
                  Team data pulled {new Date(shared.at).toLocaleTimeString()} — refreshes every minute and before each export.
                </p>
              )}
              {shared.message && <p className="mt-1.5 text-[10px] leading-relaxed text-amber-300">{shared.message}</p>}
            </div>
            <div>
              <SectionLabel hint={engine.status === 'ready' ? 'loaded' : 'loads on first capture'}>OCR engine</SectionLabel>
              <Pill tone={engine.status === 'ready' ? 'success' : engine.status === 'idle' ? 'default' : 'accent'}>
                {engine.status === 'ready' ? 'Ready' : engine.status === 'idle' ? 'Not loaded' : (engine.label || 'Loading…')}
              </Pill>
              <p className="mt-2 text-[10px] leading-relaxed text-neutral-600">
                Screenshots are read in this browser only — images are never stored in the app state or synced to the cloud.
              </p>
            </div>
          </div>
          <div className="p-5 border-t border-neutral-900 space-y-3">
            <GoogleSheetSync gsheets={gsheets} moduleId="recorder" sheetId={state.googleSheets?.sheetIds?.recorder} />
            <div>
              <label className="mb-1.5 block text-[10px] uppercase tracking-wide text-neutral-500">Export month</label>
              <div className="flex gap-2">
                <Select value={exportSel.month} onChange={e => setExportSel(x => ({ ...x, month: Number(e.target.value) }))}>
                  {RECORDER_MONTHS.map((m, i) => <option key={m} value={i}>{m}</option>)}
                </Select>
                <Input type="number" value={exportSel.year} onChange={e => setExportSel(x => ({ ...x, year: e.target.value }))} className="!w-24" />
              </div>
            </div>
            <Btn variant="primary" size="lg" onClick={onExport} disabled={busyExport} className="w-full">
              {busyExport ? <><span className="loader"></span> Building</> : <><IconDownload /> Download workbook</>}
            </Btn>
            <p className="truncate text-center font-mono text-[10px] text-neutral-600" title={recorderFileName(rec.filePattern, exportYear, exportSel.month)}>
              {recorderFileName(rec.filePattern, exportYear, exportSel.month)}
            </p>
          </div>
        </aside>
      );
    }

    function RecorderModule({ state, setState, sync, onRetrySync, moduleSwitch, onToast, confirmDialog, gsheets }) {
      const rec = useMemo(() => recorderNormalizeState(state.recorder), [state.recorder]);
      const setRec = useCallback((next) => setState(s => {
        const prev = recorderNormalizeState(s.recorder);
        return { ...s, recorder: typeof next === 'function' ? next(prev) : next };
      }), [setState]);
      const [captures, setCaptures] = useState([]);
      const [manual, setManual] = useState({});
      const [target, setTarget] = useState(() => ({ date: recorderTodayIso(), hour: new Date().getHours() }));
      const [engine, setEngine] = useState(() => ({ status: recorderWorkerPromise ? 'ready' : 'idle', label: '' }));
      const [busyExport, setBusyExport] = useState(false);
      const queueRef = useRef(Promise.resolve());
      const recRef = useRef(rec); recRef.current = rec;
      const targetRef = useRef(target); targetRef.current = target;
      // Export month/year are persisted on the slice (rec.exportMonth /
      // rec.exportYear) so the Google Sheets builder sees the same selection.
      const exportSel = { month: rec.exportMonth, year: rec.exportYear };
      const setExportSel = useCallback((updater) => setRec(r => {
        const cur = { month: r.exportMonth, year: r.exportYear };
        const nx = typeof updater === 'function' ? updater(cur) : updater;
        return { ...r, exportMonth: nx.month, exportYear: String(nx.year ?? '') };
      }), [setRec]);

      /* ── Shared team sync (design notes at the module-scope helpers) ── */
      const [shared, setShared] = useState({ status: 'idle', at: 0, message: '' });
      const sharedQueueRef = useRef({ config: null, months: {} });
      const sharedFlushTimerRef = useRef(null);
      const sharedFlushBusyRef = useRef(false);
      const sharedLoadedMonthsRef = useRef(new Set());
      const sharedRefreshRef = useRef({ inFlight: null, queued: new Set() });
      const sharedInitRef = useRef(false);
      const applyingRemoteRef = useRef(false);
      const prevRecShadowRef = useRef(rec);
      const flushSharedRef = useRef(() => {});
      const refreshSharedRef = useRef(() => null);
      const canShared = useCallback(() => !!(window.__fb && window.__fbm && sync.uid), [sync.uid]);
      const sharedErrMessage = (e) => {
        const code = e?.code ? ` (${e.code})` : '';
        return e?.code === 'permission-denied'
          ? `Firestore rules denied shared access${code} — allow signed-in reads/writes on shared/recorderConfig and shared/recorderValues_*.`
          : `Team sync failed${code}. Your entries stay saved locally — Refresh retries.`;
      };

      const scheduleSharedFlush = useCallback(() => {
        if (sharedFlushTimerRef.current) clearTimeout(sharedFlushTimerRef.current);
        sharedFlushTimerRef.current = setTimeout(() => { flushSharedRef.current(); }, 800);
      }, []);

      const flushShared = useCallback(async () => {
        const q = sharedQueueRef.current;
        if (!q.config && !Object.keys(q.months).length) return;
        if (!canShared()) return; // stays queued; retried on the next edit after sign-in
        if (sharedFlushBusyRef.current) { scheduleSharedFlush(); return; }
        sharedFlushBusyRef.current = true;
        sharedQueueRef.current = { config: null, months: {} };
        try {
          const { db } = window.__fb;
          const { doc, setDoc, serverTimestamp, deleteField } = window.__fbm;
          if (q.config) {
            await setDoc(doc(db, RECORDER_SHARED_CONFIG_DOC[0], RECORDER_SHARED_CONFIG_DOC[1]),
              { ...firestoreSanitize(q.config), updatedAt: serverTimestamp(), updatedBy: sync.uid, updatedByEmail: sync.email || '' },
              { merge: true });
          }
          for (const [mk, m] of Object.entries(q.months)) {
            await setDoc(doc(db, 'shared', recorderSharedMonthDocId(mk)),
              { values: recorderQueueToFirestore(m, deleteField), updatedAt: serverTimestamp(), updatedBy: sync.uid },
              { merge: true });
            sharedLoadedMonthsRef.current.add(mk);
          }
          setShared({ status: 'live', at: Date.now(), message: '' });
        } catch (e) {
          // Re-queue the failed ops under anything queued meanwhile (newer wins).
          const later = sharedQueueRef.current;
          const months = { ...q.months };
          for (const [mk, m] of Object.entries(later.months)) months[mk] = recorderMergeQueuedMonth(months[mk], m);
          sharedQueueRef.current = {
            config: (q.config || later.config) ? { ...(q.config || {}), ...(later.config || {}) } : null,
            months,
          };
          console.warn('recorder shared sync save failed', { code: e?.code, message: e?.message, uid: sync.uid, error: e });
          setShared({ status: 'error', at: Date.now(), message: sharedErrMessage(e) });
        } finally {
          sharedFlushBusyRef.current = false;
        }
      }, [canShared, scheduleSharedFlush, sync.uid, sync.email]);
      flushSharedRef.current = flushShared;

      const refreshShared = useCallback(async (monthKeys = []) => {
        if (!canShared()) {
          setShared({ status: window.__fb ? 'offline' : 'idle', at: Date.now(), message: window.__fb ? 'Sign in to share Recorder data with the team.' : '' });
          return null;
        }
        const want = [...new Set(monthKeys.filter(Boolean))];
        const rr = sharedRefreshRef.current;
        if (rr.inFlight) {
          want.forEach(mk => rr.queued.add(mk));
          return rr.inFlight;
        }
        const run = (async () => {
          setShared(s => ({ ...s, status: 'loading' }));
          try {
            const { db } = window.__fb;
            const { doc, getDoc, setDoc, serverTimestamp } = window.__fbm;
            const local = recRef.current;
            // Config doc: seed it the first time anyone opens the module,
            // otherwise adopt the shared copy (clients, rules, settings).
            const cref = doc(db, RECORDER_SHARED_CONFIG_DOC[0], RECORDER_SHARED_CONFIG_DOC[1]);
            const csnap = await getDoc(cref);
            let cfgPatch = null;
            if (!csnap.exists()) {
              const seed = {};
              for (const f of RECORDER_SHARED_CONFIG_FIELDS) seed[f] = firestoreSanitize(local[f]);
              await setDoc(cref, { ...seed, updatedAt: serverTimestamp(), updatedBy: sync.uid, updatedByEmail: sync.email || '' });
            } else {
              const data = firestoreDesanitize(csnap.data() || {});
              cfgPatch = {};
              for (const f of RECORDER_SHARED_CONFIG_FIELDS) if (data[f] !== undefined) cfgPatch[f] = data[f];
            }
            let migrated = {};
            try { migrated = JSON.parse(localStorage.getItem(RECORDER_SHARED_MIGRATED_KEY)) || {}; } catch {}
            const monthPatches = {};
            for (const mk of want) {
              const mref = doc(db, 'shared', recorderSharedMonthDocId(mk));
              const msnap = await getDoc(mref);
              const localMonth = recorderMonthSlice(local.values, mk);
              if (!msnap.exists()) {
                // First time anyone opens this month — seed it wholesale from
                // this device, so recordings made before the rollout (e.g. the
                // admin's existing months) are preserved in the shared copy.
                if (Object.keys(localMonth).length) {
                  await setDoc(mref, { values: localMonth, updatedAt: serverTimestamp(), updatedBy: sync.uid });
                }
                monthPatches[mk] = localMonth;
              } else {
                const sharedMonth = (firestoreDesanitize(msnap.data() || {})).values || {};
                if (!migrated[mk]) {
                  // One-time per-device top-up: upload cells this device has
                  // that the shared doc lacks. Shared wins on conflicts, so a
                  // stale laptop can never overwrite the team's newer entries.
                  const adds = recorderMissingCells(localMonth, sharedMonth);
                  if (adds) await setDoc(mref, { values: adds, updatedAt: serverTimestamp(), updatedBy: sync.uid }, { merge: true });
                  monthPatches[mk] = adds ? recorderMergeQueuedMonth(sharedMonth, adds) : sharedMonth;
                } else {
                  monthPatches[mk] = sharedMonth;
                }
              }
              migrated[mk] = 1;
              sharedLoadedMonthsRef.current.add(mk);
            }
            try { localStorage.setItem(RECORDER_SHARED_MIGRATED_KEY, JSON.stringify(migrated)); } catch {}
            // Apply to app state. Pending unflushed writes are overlaid so a
            // refresh can never revert cells the user typed a moment ago, and
            // an unchanged snapshot returns the SAME state object so the
            // quiet 60s poll doesn't bump localUpdatedAt / re-save the
            // per-user doc every minute.
            applyingRemoteRef.current = true;
            setState(s => {
              const prev = recorderNormalizeState(s.recorder);
              let values = prev.values;
              for (const [mk, m] of Object.entries(monthPatches)) {
                values = recorderApplyMonthSlice(values, mk, recorderOverlayQueuedMonth(m, sharedQueueRef.current.months[mk]));
              }
              const cfg = cfgPatch ? { ...cfgPatch, ...(sharedQueueRef.current.config || {}) } : {};
              const changed = JSON.stringify(values) !== JSON.stringify(prev.values)
                || Object.keys(cfg).some(f => JSON.stringify(cfg[f]) !== JSON.stringify(prev[f]));
              if (!changed) { applyingRemoteRef.current = false; return s; }
              return { ...s, recorder: { ...prev, ...cfg, values } };
            });
            setShared({ status: 'live', at: Date.now(), message: '' });
            return { monthPatches };
          } catch (e) {
            console.warn('recorder shared sync refresh failed', { code: e?.code, message: e?.message, uid: sync.uid, error: e });
            setShared({ status: 'error', at: Date.now(), message: sharedErrMessage(e) });
            return null;
          } finally {
            rr.inFlight = null;
            if (rr.queued.size) {
              const next = [...rr.queued];
              rr.queued.clear();
              refreshSharedRef.current(next);
            }
          }
        })();
        rr.inFlight = run;
        return run;
      }, [canShared, setState, sync.uid, sync.email]);
      refreshSharedRef.current = refreshShared;

      // Months the UI currently cares about: the recording slot, the export
      // selection, and the month open in the Data tab.
      const dataMonthRef = useRef(null);
      const monthsInView = useCallback(() => {
        const year = recorderParseIntOr(recRef.current.exportYear, new Date().getFullYear());
        return [...new Set([
          recorderMonthKey(targetRef.current.date),
          `${year}-${String(recRef.current.exportMonth + 1).padStart(2, '0')}`,
          dataMonthRef.current,
        ].filter(Boolean))];
      }, []);
      const onVisibleMonth = useCallback((mk) => {
        dataMonthRef.current = mk;
        if (canShared() && !sharedLoadedMonthsRef.current.has(mk)) refreshSharedRef.current([mk]);
      }, [canShared]);

      // Initial load, sign-in/out transitions, and newly viewed months.
      useEffect(() => {
        if (!canShared()) {
          sharedInitRef.current = false;
          setShared({ status: window.__fb ? 'offline' : 'idle', at: 0, message: (window.__fb && !sync.uid) ? 'Sign in to share Recorder data with the team.' : '' });
          return;
        }
        const missing = monthsInView().filter(mk => !sharedLoadedMonthsRef.current.has(mk));
        if (!sharedInitRef.current || missing.length) {
          sharedInitRef.current = true;
          refreshSharedRef.current(missing.length ? missing : monthsInView());
        }
      }, [sync.uid, target.date, rec.exportMonth, rec.exportYear, canShared, monthsInView]);

      // 60s poll + window-focus refresh while the module is open (plain gets —
      // see the no-onSnapshot note at the shared helpers).
      useEffect(() => {
        if (!canShared()) return;
        const tick = () => { if (!document.hidden) refreshSharedRef.current(monthsInView()); };
        const t = setInterval(tick, 60000);
        window.addEventListener('focus', tick);
        return () => { clearInterval(t); window.removeEventListener('focus', tick); };
      }, [sync.uid, canShared, monthsInView]);

      // Queue shared writes from every local recorder edit: diff the committed
      // slice against the previous one. Remote applies only refresh the shadow.
      useEffect(() => {
        const prev = prevRecShadowRef.current;
        prevRecShadowRef.current = rec;
        if (applyingRemoteRef.current) { applyingRemoteRef.current = false; return; }
        if (prev === rec || !canShared()) return;
        const q = sharedQueueRef.current;
        let any = false;
        for (const f of RECORDER_SHARED_CONFIG_FIELDS) {
          if (JSON.stringify(prev[f]) !== JSON.stringify(rec[f])) { (q.config = q.config || {})[f] = rec[f]; any = true; }
        }
        const months = recorderDiffValues(prev.values, rec.values);
        for (const [mk, m] of Object.entries(months)) { q.months[mk] = recorderMergeQueuedMonth(q.months[mk], m); any = true; }
        if (any) scheduleSharedFlush();
      }, [rec, canShared, scheduleSharedFlush]);

      // Best-effort flush when leaving the module / closing the tab soon after
      // an edit (the debounce is 800ms, so this rarely has anything to do).
      useEffect(() => () => {
        if (sharedFlushTimerRef.current) clearTimeout(sharedFlushTimerRef.current);
        flushSharedRef.current();
      }, []);

      useEffect(() => {
        window.__recorderOcrProgress = (m) => {
          if (!m || !m.status) return;
          if (m.status === 'recognizing text') {
            setEngine({ status: 'working', label: `Reading ${Math.round((m.progress || 0) * 100)}%` });
          } else if (/load|init/i.test(m.status)) {
            setEngine({ status: 'loading', label: 'Loading OCR engine…' });
          }
        };
        return () => { window.__recorderOcrProgress = null; };
      }, []);

      const processFile = useCallback((file) => {
        const id = recorderUid();
        setCaptures(list => [...list, { id, status: 'ocr', kind: null, thumb: null, entries: [], rows: [], ignored: [], clientId: null, error: null }]);
        setRec(r => r.tab === 'capture' ? r : { ...r, tab: 'capture' });
        queueRef.current = queueRef.current.then(async () => {
          try {
            const canvas = await recorderFileToCanvas(file, recRef.current.upscale);
            const thumb = recorderThumb(canvas);
            setCaptures(list => list.map(c => c.id === id ? { ...c, thumb } : c));
            const worker = await recorderEnsureWorker();
            const words = await recorderRecognizeWords(worker, canvas, 6);
            try { window.__recWords = words.map(w => w.text + ' [x' + Math.round(w.x0) + '-' + Math.round(w.x1) + ' y' + Math.round((w.y0 + w.y1) / 2) + ' c' + Math.round(w.conf) + ']'); window.__recCanvas = canvas.width + 'x' + canvas.height + ' upscale=' + recRef.current.upscale; } catch {}
            const kind = recorderClassifyWords(words);
            let entries, tableMeta = null;
            if (kind === 'cards') {
              entries = recorderParseCards(words);
            } else {
              const parsed = recorderParseTable(words, canvas.width);
              entries = parsed.entries;
              tableMeta = parsed.meta;
            }
            // Block mode (PSM 6) drops lone small digits (card values, "0"
            // readings in tight tables). Anything that came back valueless
            // gets a sparse-mode (PSM 11) second read.
            if (kind === 'cards' && (!entries.length || entries.some(e => e.value == null))) {
              const sparse = recorderParseCards(await recorderRecognizeWords(worker, canvas, 11));
              if (!entries.length) {
                entries = sparse;
              } else {
                entries = entries.map(e => {
                  if (e.value != null) return e;
                  const alt = sparse.find(x => x.name === e.name);
                  return (alt && alt.value != null) ? { ...e, value: alt.value, conf: Math.min(e.conf, alt.conf) } : e;
                });
                for (const x of sparse) if (!entries.some(e => e.name === x.name)) entries.push(x);
              }
            } else if (kind !== 'cards' && entries.length && entries.some(e => e.value == null)) {
              recorderFillTableValuesFromSparse(entries, tableMeta, await recorderRecognizeWords(worker, canvas, 11));
            }
            setEngine({ status: 'ready', label: '' });
            setCaptures(list => {
              const assigned = list.filter(c => c.id !== id && c.kind === 'cards' && c.clientId).map(c => c.clientId);
              const r = recRef.current;
              const clientId = kind === 'cards' ? recorderSuggestCardsClient(r, targetRef.current, assigned) : null;
              const { rows, ignored } = recorderRowsFromEntries(entries, kind, r, clientId);
              return list.map(c => c.id === id ? { ...c, status: entries.length ? 'ready' : 'empty', kind, entries, rows, ignored, clientId } : c);
            });
          } catch (err) {
            setEngine(e => e.status === 'working' || e.status === 'loading' ? { status: 'idle', label: '' } : e);
            setCaptures(list => list.map(c => c.id === id ? { ...c, status: 'error', error: String((err && err.message) || err) } : c));
          }
        });
      }, [setRec]);

      // Paste anywhere while this module is active (same pattern as Image Editor).
      useEffect(() => {
        const onPaste = (e) => {
          const items = (e.clipboardData && e.clipboardData.items) || [];
          let took = false;
          for (const item of items) {
            if (item.kind === 'file' && String(item.type || '').startsWith('image/')) {
              const f = item.getAsFile();
              if (f) { took = true; processFile(f); }
            }
          }
          if (took) e.preventDefault();
        };
        window.addEventListener('paste', onPaste);
        return () => window.removeEventListener('paste', onPaste);
      }, [processFile]);

      const assignCardsClient = (capId, clientId) => setCaptures(list => list.map(c => {
        if (c.id !== capId) return c;
        const { rows, ignored } = recorderRowsFromEntries(c.entries, 'cards', recRef.current, clientId);
        return { ...c, clientId, rows, ignored };
      }));
      const changeValue = (capId, rowKey, value) => setCaptures(list => list.map(c =>
        c.id !== capId ? c : { ...c, rows: c.rows.map(rw => rw.key === rowKey ? { ...rw, value } : rw) }));
      const discardCapture = (capId) => setCaptures(list => list.filter(c => c.id !== capId));

      const buildWrites = (caps, manualMap) => {
        const writes = [];
        const r = recRef.current;
        for (const cap of caps) {
          if (cap.status !== 'ready') continue;
          for (const row of cap.rows) {
            const t = String(row.value).trim();
            if (t === '' || !row.clientId) continue;
            writes.push({ clientId: row.clientId, dateIso: targetRef.current.date, rowId: row.rowId, hour: targetRef.current.hour, value: Math.min(recorderParseIntOr(t, 0), r.clampMax) });
          }
        }
        for (const [k, val] of Object.entries(manualMap)) {
          const t = String(val).trim();
          if (t === '') continue;
          const sep = k.indexOf('|');
          writes.push({ clientId: k.slice(0, sep), dateIso: targetRef.current.date, rowId: k.slice(sep + 1), hour: targetRef.current.hour, value: Math.min(recorderParseIntOr(t, 0), r.clampMax) });
        }
        return writes;
      };
      const commitWrites = (writes) => {
        if (!writes.length) { onToast('err', 'Nothing to save yet'); return false; }
        const preview = recorderApplyWrites(recRef.current.values, writes, recRef.current.overwrite);
        setRec(r => ({ ...r, values: recorderApplyWrites(r.values, writes, r.overwrite).next }));
        onToast('ok', `Saved ${preview.written} value${preview.written === 1 ? '' : 's'} → ${recorderPrettyDate(targetRef.current.date)} · ${HOURS_24[targetRef.current.hour]}${preview.skipped ? ` (${preview.skipped} kept — overwrite is off)` : ''}`);
        return true;
      };
      const commitOne = (cap) => { if (commitWrites(buildWrites([cap], {}))) discardCapture(cap.id); };
      const commitAll = () => { if (commitWrites(buildWrites(captures, manual))) { setCaptures([]); setManual({}); } };

      const missing = useMemo(() => {
        const covered = new Set();
        for (const cap of captures) {
          if (cap.status !== 'ready') continue;
          for (const row of cap.rows) if (String(row.value).trim() !== '') covered.add(row.clientId + '|' + row.rowId);
        }
        const out = [];
        for (const c of rec.clients) {
          const byRow = ((rec.values[c.id] || {})[target.date]) || {};
          const rows = c.rows.filter(r => (byRow[r.id] || {})[target.hour] == null && !covered.has(c.id + '|' + r.id));
          if (rows.length) out.push({ client: c, rows });
        }
        return out;
      }, [captures, rec, target]);

      const slotStats = useMemo(() => rec.clients.map(c => {
        const byRow = ((rec.values[c.id] || {})[target.date]) || {};
        return { id: c.id, name: c.name, filled: c.rows.filter(r => (byRow[r.id] || {})[target.hour] != null).length, total: c.rows.length };
      }), [rec, target]);

      const preloadEngine = async () => {
        try {
          setEngine({ status: 'loading', label: 'Loading OCR engine…' });
          await recorderEnsureWorker();
          setEngine({ status: 'ready', label: '' });
          onToast('ok', 'OCR engine ready');
        } catch (err) {
          setEngine({ status: 'idle', label: '' });
          onToast('err', String((err && err.message) || err));
        }
      };
      const doExport = async () => {
        setBusyExport(true);
        try {
          const year = recorderParseIntOr(recRef.current.exportYear, new Date().getFullYear());
          const monthIdx = recRef.current.exportMonth;
          const mk = `${year}-${String(monthIdx + 1).padStart(2, '0')}`;
          // Push pending edits and pull the team's latest for the export month
          // so the workbook includes everyone's inputs. Best effort — the
          // export still works offline from local data.
          let recForExport = recRef.current;
          try {
            await flushSharedRef.current();
            const res = await refreshSharedRef.current([mk]);
            if (res && res.monthPatches && res.monthPatches[mk]) {
              recForExport = {
                ...recForExport,
                values: recorderApplyMonthSlice(recForExport.values, mk,
                  recorderOverlayQueuedMonth(res.monthPatches[mk], sharedQueueRef.current.months[mk])),
              };
            }
          } catch {}
          const wb = recorderBuildWorkbook(recForExport, year, monthIdx);
          const buf = await wb.xlsx.writeBuffer();
          const fname = recorderFileName(recForExport.filePattern, year, monthIdx);
          saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fname);
          onToast('ok', `Downloaded ${fname}`);
          // Mirror to the module's Google Sheet when already connected —
          // silent, exactly like the other modules' Generate buttons.
          if (gsheets && gsheets.conn && gsheets.conn.connected) {
            try { await gsheets.sync('recorder', { interactive: false }); } catch { /* toast surfaced by the sync handler */ }
          }
        } catch (err) {
          onToast('err', `Export failed: ${String((err && err.message) || err)}`);
        } finally {
          setBusyExport(false);
        }
      };
      const importFile = async (file) => {
        try {
          const buf = await file.arrayBuffer();
          const { values, count, sheets } = await recorderImportWorkbook(recRef.current, buf);
          if (!count) { onToast('err', 'No matching sheets / rows found in that workbook'); return; }
          setRec(r => ({ ...r, values }));
          onToast('ok', `Imported ${count} values from ${sheets} sheet${sheets === 1 ? '' : 's'}`);
        } catch (err) {
          onToast('err', `Import failed: ${String((err && err.message) || err)}`);
        }
      };

      const hasPending = captures.some(c => c.status === 'ready' && c.rows.some(r => String(r.value).trim() !== ''))
        || Object.values(manual).some(v => String(v).trim() !== '');
      const TABS = [['capture', 'Capture'], ['data', 'Data'], ['config', 'Config'], ['rules', 'Rules']];

      return (
        <>
          <RecorderSidebar state={state} setState={setState} rec={rec} sync={sync} onRetrySync={onRetrySync}
            target={target} setTarget={setTarget} engine={engine} slotStats={slotStats}
            exportSel={exportSel} setExportSel={setExportSel} onExport={doExport} busyExport={busyExport}
            shared={shared} onRefreshShared={() => refreshSharedRef.current(monthsInView())} gsheets={gsheets} />
          <main className="flex-1 min-w-0 flex flex-col">
            <header className="app-topbar border-b border-neutral-900 px-4 sm:px-6 lg:px-8 py-4 sm:py-5 flex flex-wrap items-start justify-between gap-4 sticky top-0 bg-[#1c1c1f]/95 backdrop-blur z-20">
              <div className="app-title-block min-w-0">
                <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-1">Module · Team shared</div>
                <h2 className="truncate text-[22px] font-bold tracking-tight">Recorder</h2>
              </div>
              <div className="app-header-controls flex flex-wrap items-center gap-2 sm:gap-3">
                <div className="flex items-center gap-1.5">
                  {TABS.map(([id, label]) => (
                    <button key={id} onClick={() => setRec(r => ({ ...r, tab: id }))}
                      className={`rounded-md border px-3 py-1.5 text-xs font-medium transition-colors ${rec.tab === id ? 'bg-white text-black border-white' : 'border-neutral-800 text-neutral-400 hover:text-neutral-100 hover:border-neutral-600'}`}>
                      {label}
                    </button>
                  ))}
                </div>
                {moduleSwitch}
              </div>
            </header>
            <div className="flex-1 min-w-0 p-4 sm:p-6 lg:p-8">
              {rec.tab === 'capture' && (
                <div className="max-w-6xl space-y-5 anim-fade-in">
                  <RecorderDropZone onFiles={files => files.forEach(processFile)} engine={engine} />
                  {captures.map(cap => (
                    <RecorderCaptureCard key={cap.id} cap={cap} rec={rec} target={target}
                      onAssign={assignCardsClient} onValue={changeValue}
                      onCommit={() => commitOne(cap)} onDiscard={() => discardCapture(cap.id)} />
                  ))}
                  <RecorderMissingPanel missing={missing} manual={manual} setManual={setManual} hourLabel={`${recorderPrettyDate(target.date)} · ${HOURS_24[target.hour]}`} />
                  {hasPending && (
                    <div className="sticky bottom-4 z-10 flex flex-wrap items-center justify-between gap-3 rounded-lg border border-blue-500/40 bg-[#16203a]/95 px-4 py-3 backdrop-blur">
                      <div className="text-xs text-neutral-300">
                        Saving to <span className="font-semibold text-neutral-100">{recorderPrettyDate(target.date)}</span> · <span className="font-semibold text-neutral-100">{HOURS_24[target.hour]}</span>
                        <span className="text-neutral-500"> — wrong slot? change it in the sidebar first</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <Btn variant="ghost" size="md" onClick={() => { setCaptures([]); setManual({}); }}>Discard all</Btn>
                        <Btn variant="accent" size="md" onClick={commitAll}>Commit all</Btn>
                      </div>
                    </div>
                  )}
                </div>
              )}
              {rec.tab === 'data' && <RecorderDataTab rec={rec} setRec={setRec} onToast={onToast} confirmDialog={confirmDialog} onImport={importFile} onVisibleMonth={onVisibleMonth} />}
              {rec.tab === 'config' && <RecorderConfigTab rec={rec} setRec={setRec} confirmDialog={confirmDialog} onToast={onToast} />}
              {rec.tab === 'rules' && <RecorderRulesTab rec={rec} setRec={setRec} engine={engine} onPreload={preloadEngine} />}
            </div>
            <footer className="mt-auto border-t border-neutral-900 px-8 py-4 flex items-center justify-between text-[11px] text-neutral-600">
              <span className="font-mono">state.recorder · numbers only, images never stored</span>
              <span>{rec.clients.length} clients · {rec.clients.reduce((n, c) => n + c.rows.length, 0)} watched rows</span>
            </footer>
          </main>
        </>
      );
    }


    // Role-based access. Admin sees everything; semi-admin gets the
    // light-touch modules plus the team-shared Recorder. Anyone signed in
    // defaults to semi_admin, so every Firebase user can use the Recorder.
    const ROLE_MODULES = {
      admin:      ['sip_fcs', 'bmr', 'bmr_sms', 'whitelist_sms', 'editor', 'rca', 'image', 'recorder'],
      semi_admin: ['whitelist_sms', 'editor', 'rca'],
    };
    const ROLE_LABELS = { admin: 'Admin', semi_admin: 'Semi-admin' };
    function isSuperAdminEmail(email) {
      return !!email && SUPER_ADMINS.includes(email.trim().toLowerCase());
    }
    function effectiveRole(email, dbRole) {
      if (isSuperAdminEmail(email)) return 'admin';
      return dbRole === 'admin' ? 'admin' : 'semi_admin';
    }
    function moduleAllowed(moduleId, role) {
      return (ROLE_MODULES[role] || ROLE_MODULES.semi_admin).includes(moduleId);
    }

    function UsersAdminPanel({ sync }) {
      const [rows, setRows] = useState([]);
      const [loading, setLoading] = useState(true);
      const [error, setError] = useState(null);
      const [savingUid, setSavingUid] = useState(null);
      const [filter, setFilter] = useState('');

      const load = useCallback(async () => {
        if (!window.__fb || !window.__fbm) { setError('Firebase is not available.'); setLoading(false); return; }
        setLoading(true); setError(null);
        try {
          const { collection, getDocs } = window.__fbm;
          const snap = await getDocs(collection(window.__fb.db, 'userProfiles'));
          const list = snap.docs.map(d => {
            const data = d.data() || {};
            return {
              uid: d.id,
              email: data.email || '',
              dbRole: data.role || null,
              updatedAt: timestampMs(data.updatedAt),
              createdAt: timestampMs(data.createdAt),
            };
          });
          list.sort((a, b) => (a.email || '').localeCompare(b.email || ''));
          setRows(list);
        } catch (e) {
          console.warn('userProfiles list failed', e);
          setError(syncErrorMessage(e, 'Could not load users'));
        } finally {
          setLoading(false);
        }
      }, []);

      useEffect(() => { load(); }, [load]);

      const setRole = async (uid, nextRole) => {
        if (!window.__fb || !window.__fbm) return;
        setSavingUid(uid); setError(null);
        try {
          const { doc, setDoc, serverTimestamp, deleteField } = window.__fbm;
          const patch = {
            role: nextRole === 'semi_admin' ? deleteField() : nextRole,
            updatedAt: serverTimestamp(),
          };
          await setDoc(doc(window.__fb.db, 'userProfiles', uid), patch, { merge: true });
          setRows(rs => rs.map(r => r.uid === uid ? { ...r, dbRole: nextRole === 'semi_admin' ? null : nextRole } : r));
        } catch (e) {
          console.warn('role update failed', e);
          setError(syncErrorMessage(e, 'Could not save role'));
        } finally {
          setSavingUid(null);
        }
      };

      const formatDate = (ms) => ms ? new Date(ms).toLocaleString() : '—';
      const term = filter.trim().toLowerCase();
      const visible = term
        ? rows.filter(r => (r.email || '').toLowerCase().includes(term) || (r.uid || '').toLowerCase().includes(term))
        : rows;

      return (
        <div className="space-y-6">
          <div>
            <SectionLabel hint={`${rows.length} signed-up users · admins see every module · semi-admins see Whitelist SMS + Editor`}>User access</SectionLabel>
            <div className="flex flex-wrap items-center gap-2 mb-3">
              <Input placeholder="Filter by email or UID…" value={filter} onChange={e => setFilter(e.target.value)} className="max-w-xs" />
              <Btn variant="ghost" size="sm" onClick={load} disabled={loading}>{loading ? 'Loading…' : 'Refresh'}</Btn>
            </div>

            {error && (
              <div className="rounded-md border border-red-500/30 bg-red-500/5 px-3 py-2 text-[12px] text-red-300 mb-3">
                {error}
              </div>
            )}

            <div className="rounded-lg border border-neutral-800 overflow-hidden">
              <table className="w-full text-sm">
                <thead className="bg-neutral-950/60 text-[10px] uppercase tracking-[0.2em] text-neutral-500">
                  <tr>
                    <th className="text-left px-4 py-2.5 font-medium">Email</th>
                    <th className="text-left px-4 py-2.5 font-medium">Role</th>
                    <th className="text-left px-4 py-2.5 font-medium hidden md:table-cell">Last updated</th>
                    <th className="text-left px-4 py-2.5 font-medium hidden lg:table-cell">UID</th>
                  </tr>
                </thead>
                <tbody>
                  {loading && (
                    <tr><td colSpan={4} className="px-4 py-6 text-center text-neutral-500 text-xs">Loading users…</td></tr>
                  )}
                  {!loading && visible.length === 0 && (
                    <tr><td colSpan={4} className="px-4 py-6 text-center text-neutral-500 text-xs">
                      {rows.length === 0 ? 'No user profiles yet. Users appear here after they sign in for the first time.' : 'No users match this filter.'}
                    </td></tr>
                  )}
                  {!loading && visible.map((r) => {
                    const isSelf = r.uid === sync.uid;
                    const isSuper = isSuperAdminEmail(r.email);
                    const effective = effectiveRole(r.email, r.dbRole);
                    const selectValue = r.dbRole === 'admin' ? 'admin' : 'semi_admin';
                    return (
                      <tr key={r.uid} className="border-t border-neutral-900">
                        <td className="px-4 py-3 align-middle">
                          <div className="flex items-center gap-2 min-w-0">
                            <span className="truncate" title={r.email}>{r.email || <span className="text-neutral-600">(no email)</span>}</span>
                            {isSelf && <span className="shrink-0 text-[10px] uppercase tracking-wide text-blue-300 bg-blue-500/10 border border-blue-500/30 rounded px-1.5 py-0.5">you</span>}
                            {isSuper && <span className="shrink-0 text-[10px] uppercase tracking-wide text-amber-300 bg-amber-500/10 border border-amber-500/30 rounded px-1.5 py-0.5" title="Listed in SUPER_ADMINS — always admin">super</span>}
                          </div>
                        </td>
                        <td className="px-4 py-3 align-middle">
                          {isSuper ? (
                            <span className="text-xs text-neutral-400">{ROLE_LABELS.admin} <span className="text-neutral-600">(locked)</span></span>
                          ) : (
                            <div className="flex items-center gap-2">
                              <Select
                                value={selectValue}
                                disabled={savingUid === r.uid || isSelf}
                                onChange={e => setRole(r.uid, e.target.value)}
                                className="max-w-[160px]">
                                <option value="semi_admin">{ROLE_LABELS.semi_admin}</option>
                                <option value="admin">{ROLE_LABELS.admin}</option>
                              </Select>
                              {savingUid === r.uid && <span className="loader text-neutral-400" />}
                              {isSelf && <span className="text-[10px] text-neutral-600">can't change own role</span>}
                            </div>
                          )}
                        </td>
                        <td className="px-4 py-3 align-middle hidden md:table-cell text-xs text-neutral-500">{formatDate(r.updatedAt || r.createdAt)}</td>
                        <td className="px-4 py-3 align-middle hidden lg:table-cell text-[11px] font-mono text-neutral-600 truncate max-w-[200px]" title={r.uid}>{r.uid}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          <div className="rounded-lg border border-neutral-800 bg-neutral-950/40 p-4 text-[12px] text-neutral-400 leading-relaxed">
            <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-2">How roles work</div>
            <ul className="space-y-1.5 list-disc pl-5">
              <li><span className="text-neutral-200">Admin</span> — full access: SIP / FCS, BMR VOIP, BMR SMS, Whitelist SMS, Editor, and this Users panel.</li>
              <li><span className="text-neutral-200">Semi-admin</span> — Whitelist SMS and Editor only. Default for every new sign-in.</li>
              <li>Emails listed in <code className="text-amber-300 bg-neutral-900 px-1 rounded">SUPER_ADMINS</code> (in <code className="text-neutral-300 bg-neutral-900 px-1 rounded">index.html</code>) are always admin and cannot be demoted from this UI.</li>
              <li>Role changes take effect the next time the user signs in or refreshes the page.</li>
            </ul>
            <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500 mt-4 mb-2">Firestore security rules</div>
            <p className="mb-2">For this to be safe in production, lock down <code className="text-neutral-300 bg-neutral-900 px-1 rounded">userProfiles</code> so only admins can edit roles:</p>
            <pre className="overflow-x-auto rounded bg-neutral-950 border border-neutral-900 p-3 text-[11px] font-mono text-neutral-300">{`match /userProfiles/{uid} {
  function isAdmin() {
    return request.auth != null
      && get(/databases/$(database)/documents/userProfiles/$(request.auth.uid)).data.role == 'admin';
  }
  allow read:  if request.auth != null && (request.auth.uid == uid || isAdmin());
  allow create: if request.auth.uid == uid
    && !('role' in request.resource.data);
  allow update: if isAdmin()
    || (request.auth.uid == uid
        && request.resource.data.role == resource.data.role);
}

// Shared Whitelist SMS test number database — any signed-in user (admin or
// semi-admin) may read and write so semi-admins can add numbers too.
match /shared/whitelistSmsTestNumbers {
  allow read, write: if request.auth != null;
}`}</pre>
          </div>
        </div>
      );
    }

    function ModuleSwitcher({ modules, current, onSelect }) {
      const containerRef = useRef(null);
      const measureRef = useRef(null);
      const moreRef = useRef(null);
      const [visibleCount, setVisibleCount] = useState(modules.length);
      const [open, setOpen] = useState(false);

      const recalculate = useCallback(() => {
        const container = containerRef.current;
        const strip = measureRef.current;
        if (!container || !strip) return;
        // Measure the available width from the parent so the segmented bar
        // can shrink/grow correctly — its own clientWidth only reflects
        // what its current children occupy.
        const parent = container.parentElement;
        const parentStyle = parent ? window.getComputedStyle(parent) : null;
        const parentPad = parentStyle
          ? (parseFloat(parentStyle.paddingLeft) || 0) + (parseFloat(parentStyle.paddingRight) || 0)
          : 0;
        const available = parent
          ? Math.max(0, parent.clientWidth - parentPad - 8)
          : container.clientWidth;
        const buttons = Array.from(strip.children);
        if (!buttons.length || !available) return;
        const widths = buttons.map(b => b.offsetWidth);
        const gap = 4;
        const padding = 4;
        const moreWidth = 78;
        const allWidth = widths.reduce((s, w) => s + w, 0) + Math.max(0, widths.length - 1) * gap + padding;
        if (allWidth <= available) {
          setVisibleCount(modules.length);
          return;
        }
        let used = padding + moreWidth + gap;
        let count = 0;
        for (let i = 0; i < widths.length; i++) {
          const next = used + widths[i] + (count > 0 ? gap : 0);
          if (next > available) break;
          used = next;
          count++;
        }
        setVisibleCount(Math.max(1, count));
      }, [modules]);

      useLayoutEffect(() => { recalculate(); }, [recalculate]);

      useEffect(() => {
        const container = containerRef.current;
        if (!container) return;
        // Observe the parent element — the segmented bar itself only resizes
        // when its children change, so watching it would miss viewport
        // resizes when the visible-count is already accurate. The parent
        // (app-header-controls) reflects the actual available width.
        const target = container.parentElement || container;
        const ro = new ResizeObserver(() => recalculate());
        ro.observe(target);
        window.addEventListener('resize', recalculate);
        return () => {
          ro.disconnect();
          window.removeEventListener('resize', recalculate);
        };
      }, [recalculate]);

      useEffect(() => {
        if (!open) return;
        const closeOutside = (e) => {
          if (!moreRef.current?.contains(e.target)) setOpen(false);
        };
        const closeOnEscape = (e) => {
          if (e.key === 'Escape') setOpen(false);
        };
        window.addEventListener('pointerdown', closeOutside);
        window.addEventListener('keydown', closeOnEscape);
        return () => {
          window.removeEventListener('pointerdown', closeOutside);
          window.removeEventListener('keydown', closeOnEscape);
        };
      }, [open]);

      const safeCount = Math.min(visibleCount, modules.length);
      const visible = modules.slice(0, safeCount);
      const overflow = modules.slice(safeCount);
      const overflowHasActive = overflow.some(m => current === m.id);

      const baseBtn = 'px-3 py-1 text-[11px] font-semibold uppercase tracking-wider rounded transition-colors whitespace-nowrap';
      const activeCls = 'app-segment-active bg-blue-500/20 text-blue-200 border border-blue-500/40';
      const mutedCls = 'app-segment-muted text-neutral-500 hover:text-neutral-200';

      return (
        <div ref={containerRef} className="app-segmented relative flex items-center gap-1 rounded-md border border-neutral-800 bg-[#1a1a1d] p-0.5">
          <div ref={measureRef} aria-hidden="true"
            style={{ position: 'fixed', left: '-99999px', top: 0, display: 'flex', gap: '4px', pointerEvents: 'none', visibility: 'hidden' }}>
            {modules.map(m => (
              <button key={`m-${m.id}`} type="button" tabIndex={-1} className={`${baseBtn} ${mutedCls}`}>{m.label}</button>
            ))}
          </div>

          {visible.map(m => {
            const isActive = current === m.id || (m.id === 'sip_fcs' && !current);
            return (
              <button key={m.id} type="button" onClick={() => onSelect(m.id)}
                className={`${baseBtn} ${isActive ? activeCls : mutedCls}`}>
                {m.label}
              </button>
            );
          })}

          {overflow.length > 0 && (
            <div ref={moreRef} className="relative">
              <button type="button" onClick={() => setOpen(v => !v)} aria-expanded={open} aria-haspopup="menu" aria-label="More modules"
                className={`${baseBtn} inline-flex items-center gap-1 ${overflowHasActive ? activeCls : mutedCls}`}>
                More
                <svg width="10" height="10" viewBox="0 0 10 10" fill="none" aria-hidden="true">
                  <path d="M2 4l3 3 3-3" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </button>
              {open && (
                <div role="menu"
                  className="app-segment-more-menu absolute right-0 top-full z-30 mt-1 min-w-[170px] rounded-md border border-neutral-800 bg-[#232327] p-1 shadow-2xl">
                  {overflow.map(m => {
                    const isActive = current === m.id;
                    return (
                      <button key={m.id} type="button" role="menuitem"
                        onClick={() => { onSelect(m.id); setOpen(false); }}
                        className={`app-segment-more-item ${isActive ? 'is-active bg-blue-500/20 text-blue-200' : 'text-neutral-400 hover:bg-neutral-900 hover:text-neutral-100'} block w-full text-left px-3 py-1.5 text-[11px] font-semibold uppercase tracking-wider rounded transition-colors whitespace-nowrap`}>
                        {m.label}
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
          )}
        </div>
      );
    }

    function App() {
      const [state, setStateRaw] = useState(() => {
        try {
          const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
          if (saved && Array.isArray(saved.sheets)) {
            return normalizeAppState(saved, { migrateLegacyLabels: true });
          }
        } catch {}
        return normalizeAppState(DEFAULT_STATE, { migrateLegacyLabels: true });
      });
      const userEditRef = useRef(false);
      const setState = useCallback((next) => {
        setStateRaw(prev => {
          const resolved = typeof next === 'function' ? next(prev) : next;
          if (resolved === prev) return prev;
          userEditRef.current = true;
          return normalizeAppState({ ...resolved, localUpdatedAt: Date.now() });
        });
      }, []);
      const [busy, setBusy] = useState(false);
      const [toast, setToast] = useState(null);
      const [sync, setSync] = useState({
        status: window.__fb ? 'connecting' : 'offline',
        uid: null,
        email: null,
        role: 'semi_admin',
        message: window.__fb ? null : 'Firebase scripts did not load, so cloud sync is unavailable. Local browser saving still works.',
      });
      const [dialog, setDialog] = useState(null);
      // Google Sheets sync — transient connection status only (NOT persisted;
      // the OAuth token lives in memory inside googleSheetsSync). Durable sheet
      // IDs live in state.googleSheets.sheetIds.
      const [googleConn, setGoogleConn] = useState(() => ({
        connected: false,
        email: '',
        busyModule: null,   // moduleId (or '__connect__') currently working
        lastSync: {},       // { [moduleId]: { at, ok, fileId } }
        error: null,
      }));
      const [editorUndoStack, setEditorUndoStack] = useState([]);
      const hydratedRef = useRef(false);
      const localSaveTimerRef = useRef(null);
      const saveTimerRef = useRef(null);
      const lastSavedJsonRef = useRef('');
      // Shared Whitelist SMS test number database. Lives in a single Firestore
      // doc (shared/whitelistSmsTestNumbers) so every signed-in user — admin
      // and semi-admin alike — sees and edits the same list. Both roles can
      // add, edit, and delete; per-device selection state still lives in the
      // per-user wl.selectedNumberIds.
      const [sharedTestNumbers, setSharedTestNumbersRaw] = useState(() => DEFAULT_WL_SMS_STATE.testNumbers.map(n => ({ ...n })));
      const [sharedTestNumbersReady, setSharedTestNumbersReady] = useState(false);
      const sharedMigrationDoneRef = useRef(false);
      const stateRef = useRef(null);

      const upsertUserProfile = useCallback(async (user) => {
        if (!window.__fb || !window.__fbm || !user) return null;
        const { db } = window.__fb;
        const { doc, getDoc, setDoc, serverTimestamp } = window.__fbm;
        // Hard cap each Firestore call. If rules deny or the network stalls,
        // the SDK can sit on a get/set indefinitely and the caller (sign-in)
        // never gets to flip status off "connecting" — that's the "LOADING…"
        // forever symptom for semi-admins whose profile read is slow/denied.
        const withTimeout = (p, label) => Promise.race([
          p,
          new Promise((_, rej) => setTimeout(() => rej(new Error(label + ' timed out')), 10000)),
        ]);
        try {
          const ref = doc(db, 'userProfiles', user.uid);
          const snap = await withTimeout(getDoc(ref), 'userProfile read');
          const now = serverTimestamp();
          if (!snap.exists()) {
            // First-time: write email + createdAt only. Never set `role` here —
            // role must be assigned by an admin via the Users panel.
            try {
              await withTimeout(
                setDoc(ref, { email: user.email || '', createdAt: now, updatedAt: now }, { merge: true }),
                'userProfile create'
              );
            } catch (e) {
              // Profile create failing shouldn't block sign-in. Worst case the
              // user shows up in the Users panel only after a later write.
              console.warn('userProfile create failed', e);
            }
            return null;
          }
          // Keep email up to date in case the user changed it in Firebase Auth.
          const patch = { updatedAt: now };
          if ((snap.data() || {}).email !== user.email) patch.email = user.email || '';
          if (Object.keys(patch).length > 1) {
            try {
              await withTimeout(setDoc(ref, patch, { merge: true }), 'userProfile update');
            } catch (e) {
              console.warn('userProfile update failed', e);
            }
          }
          return (snap.data() || {}).role || null;
        } catch (e) {
          console.warn('userProfile upsert failed', e);
          return null;
        }
      }, []);

      const hydrateUserState = useCallback(async (user) => {
        if (!window.__fb || !user) return false;
        const { auth, db } = window.__fb;
        // Drop stale results once the user has signed out (or switched accounts)
        // in the time it took the Firestore round-trip to finish. Without this,
        // a late-arriving response would write the old uid back into sync state
        // and hide the AuthModal — leaving a black screen with no sign-in form.
        const isStale = () => auth.currentUser?.uid !== user.uid;
        // If the modular Firestore failed to load, we still want to recognise
        // the signed-in user so the AccountChip + sign-out button render —
        // just go straight to offline mode without trying to read cloud state.
        if (!window.__fbm || !db) {
          const role = effectiveRole(user.email, null);
          setSync({
            status: 'offline',
            uid: user.uid,
            email: user.email,
            role,
            message: 'Cloud sync unavailable — Firestore SDK failed to load on this machine. Auth and local data still work.',
          });
          hydratedRef.current = true;
          return false;
        }
        const { doc, getDoc } = window.__fbm;
        const dbRole = await upsertUserProfile(user);
        if (isStale()) return false;
        const role = effectiveRole(user.email, dbRole);
        setSync({ status: 'loading', uid: user.uid, email: user.email, role, message: null });
        try {
          if (isStale()) return false;
          const ref = doc(db, 'users', user.uid, 'state', 'main');
          // Hard cap the network read. If it hangs (e.g. rules path that the
          // SDK keeps retrying), we'd otherwise be stuck on "Loading…" forever.
          const snap = await Promise.race([
            getDoc(ref),
            new Promise((_, rej) => setTimeout(() => rej(new Error('Firestore read timed out after 15s')), 15000)),
          ]);
          if (isStale()) return false;
          if (snap.exists()) {
            const remote = withStateOwner(
              normalizeAppState(mergeLocalSheetActiveFlags(firestoreDesanitize(snap.data())), { migrateLegacyLabels: true }),
              user
            );
            if (remote && Array.isArray(remote.sheets)) {
              setStateRaw(local => {
                const current = normalizeAppState(local);
                const ownerUid = stateOwnerUid(current);
                const keepLocal = ownerUid === user.uid && stateEditedAt(current) > stateEditedAt(remote);
                if (keepLocal) userEditRef.current = true;
                else userEditRef.current = false;
                return withStateOwner({ ...(keepLocal ? current : remote), tab: current.tab }, user);
              });
            }
          } else {
            setStateRaw(local => {
              const current = normalizeAppState(local);
              const ownerUid = stateOwnerUid(current);
              const next = ownerUid && ownerUid !== user.uid
                ? { ...DEFAULT_STATE, theme: current.theme || DEFAULT_STATE.theme, tab: current.tab || DEFAULT_STATE.tab }
                : current;
              userEditRef.current = true;
              return withStateOwner({ ...next, localUpdatedAt: Date.now() }, user);
            });
          }
          hydratedRef.current = true;
          setSync({ status: 'synced', uid: user.uid, email: user.email, role, message: null });
          return true;
        } catch (e) {
          // Log the structured error so we can tell rules-deny from timeout
          // from offline. role/uid included to correlate with the deployed
          // Firestore rules when triaging.
          console.warn('Firestore hydrate failed', {
            code: e?.code, message: e?.message, role, uid: user.uid, error: e,
          });
          if (isStale()) return false;
          hydratedRef.current = true;
          setSync({
            status: 'offline',
            uid: user.uid,
            email: user.email,
            role,
            message: syncErrorMessage(e, 'Cloud sync read failed'),
          });
          return false;
        }
      }, [upsertUserProfile]);

      const retrySync = useCallback(async () => {
        if (!window.__fb) {
          setSync({ status: 'offline', uid: null, email: null, role: 'semi_admin', message: 'Firebase scripts did not load, so cloud sync is unavailable. Local browser saving still works.' });
          return;
        }
        const user = window.__fb.auth.currentUser;
        if (!user || user.isAnonymous) {
          setSync({ status: 'signed-out', uid: null, email: null, role: 'semi_admin', message: null });
          return;
        }
        await hydrateUserState(user);
      }, [hydrateUserState]);

      useEffect(() => {
        __dialogHandler = (req) => setDialog(req);
        return () => { __dialogHandler = null; };
      }, []);
      const resolveDialog = (value) => {
        const d = dialog;
        setDialog(null);
        if (d) d.resolve(value);
      };

      // localStorage cache (debounced so large BMR client edits stay responsive)
      useEffect(() => {
        if (localSaveTimerRef.current) clearTimeout(localSaveTimerRef.current);
        localSaveTimerRef.current = setTimeout(() => {
          try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch {}
        }, 180);
        return () => { if (localSaveTimerRef.current) clearTimeout(localSaveTimerRef.current); };
      }, [state]);

      // Keep a ref to current state so async effects can read it without
      // forcing themselves to depend on the whole state object.
      useEffect(() => { stateRef.current = state; }, [state]);

      // One-shot loader for the shared test-numbers doc. The previous
      // implementation used onSnapshot for real-time updates, but Firestore's
      // watch_change.ts hits "INTERNAL ASSERTION FAILED: Unexpected state" on
      // some users' networks/machines — even with experimentalForceLongPolling
      // — and once that fires the entire SDK is unreliable for the session.
      // A plain get() doesn't go through the watch target lifecycle, so it
      // sidesteps the bug entirely. Trade-off: no live updates from other
      // devices, but the data refreshes on sign-in and after every save.
      const loadSharedTestNumbers = useCallback(async () => {
        if (!window.__fb || !window.__fbm || !sync.uid) return;
        const { db } = window.__fb;
        const { doc, getDoc, setDoc, serverTimestamp } = window.__fbm;
        const ref = doc(db, SHARED_WL_SMS_NUMBERS_DOC[0], SHARED_WL_SMS_NUMBERS_DOC[1]);
        try {
          const snap = await getDoc(ref);
          if (!snap.exists()) {
            // Doc missing — try to seed with defaults (best effort; harmless
            // if it fails for permissions, the user just sees built-ins).
            try {
              await setDoc(ref, {
                items: DEFAULT_WL_SMS_STATE.testNumbers.map(n => ({ ...n })),
                updatedAt: serverTimestamp(),
                updatedBy: sync.uid,
              });
              setSharedTestNumbersRaw(DEFAULT_WL_SMS_STATE.testNumbers.map(n => ({ ...n })));
            } catch (seedErr) {
              console.warn('shared testNumbers seed failed', {
                code: seedErr?.code, message: seedErr?.message, uid: sync.uid, error: seedErr,
              });
            }
            setSharedTestNumbersReady(true);
            return;
          }
          const raw = (snap.data() || {}).items;
          const items = Array.isArray(raw) ? firestoreDesanitize(raw) : [];
          setSharedTestNumbersRaw(items);
          setSharedTestNumbersReady(true);
        } catch (err) {
          console.warn('shared testNumbers load failed', {
            code: err?.code, message: err?.message, uid: sync.uid, error: err,
          });
          // Don't block the UI — show built-in defaults and let the user retry.
          setSharedTestNumbersReady(true);
        }
      }, [sync.uid]);

      // Push an updated shared test number list to Firestore (and update the
      // local copy optimistically). All signed-in users — admin and semi-admin
      // — may call this; Firestore rules must permit it.
      const setSharedTestNumbers = useCallback(async (nextItems) => {
        const safeItems = Array.isArray(nextItems) ? nextItems : [];
        setSharedTestNumbersRaw(safeItems);
        if (!window.__fb || !window.__fbm || !sync.uid) return;
        try {
          const { db } = window.__fb;
          const { doc, setDoc, serverTimestamp } = window.__fbm;
          await setDoc(doc(db, SHARED_WL_SMS_NUMBERS_DOC[0], SHARED_WL_SMS_NUMBERS_DOC[1]), {
            items: firestoreSanitize(safeItems),
            updatedAt: serverTimestamp(),
            updatedBy: sync.uid,
          });
        } catch (e) {
          // Surface the actual Firestore error code (e.g. permission-denied) in
          // both the console and the toast — a bare "check rules" message hides
          // whether it's a rules deny, a network error, or something else.
          console.warn('shared testNumbers save failed', {
            code: e?.code, message: e?.message, role: sync.role, uid: sync.uid, error: e,
          });
          const code = e?.code ? ` (${e.code})` : '';
          setToast({
            type: 'err',
            msg: `Saving the shared number database failed${code}. ${e?.code === 'permission-denied' ? 'Firestore rules denied the write — verify the deployed rules allow signed-in writes to /shared/whitelistSmsTestNumbers.' : 'Check the browser console for details.'}`,
          });
          setTimeout(() => setToast(null), 8000);
        }
      }, [sync.uid, sync.role]);

      // Load the shared test numbers once on sign-in (replaces the realtime
      // listener). Sign-out resets to built-in defaults.
      useEffect(() => {
        if (!window.__fb || !sync.uid) {
          setSharedTestNumbersRaw(DEFAULT_WL_SMS_STATE.testNumbers.map(n => ({ ...n })));
          setSharedTestNumbersReady(false);
          sharedMigrationDoneRef.current = false;
          return;
        }
        loadSharedTestNumbers();
      }, [sync.uid, loadSharedTestNumbers]);

      // One-time migration: once shared list is ready and the user's state has
      // been hydrated, drain any per-user testNumbers (from legacy state) into
      // the shared list, dedupe by digits+network, then clear the per-user
      // copy and remap selectedNumberIds onto shared IDs. Runs once per
      // session (sharedMigrationDoneRef).
      useEffect(() => {
        if (!window.__fb || !sync.uid) return;
        if (!sharedTestNumbersReady) return;
        if (sharedMigrationDoneRef.current) return;
        const personal = state?.whitelistSms?.testNumbers || [];
        if (!personal.length) {
          sharedMigrationDoneRef.current = true;
          return;
        }
        sharedMigrationDoneRef.current = true;

        const items = sharedTestNumbers;
        const sharedByKey = new Map(items.map(n => [wlSmsNumberKey(n), n.id]));
        const sharedIds = new Set(items.map(n => n.id));
        const toAdd = [];
        const idMap = new Map();
        for (const n of personal) {
          const k = wlSmsNumberKey(n);
          if (!k.split('|')[0]) continue;
          if (sharedByKey.has(k)) {
            idMap.set(n.id, sharedByKey.get(k));
            continue;
          }
          const newId = sharedIds.has(n.id) ? wlSmsId('tn') : n.id;
          const entry = { id: newId, label: n.label || n.number, number: n.number, network: n.network };
          toAdd.push(entry);
          sharedIds.add(newId);
          sharedByKey.set(k, newId);
          idMap.set(n.id, newId);
        }

        if (toAdd.length && window.__fbm) {
          const merged = [...items, ...toAdd];
          const { db } = window.__fb;
          const { doc, setDoc, serverTimestamp } = window.__fbm;
          setDoc(doc(db, SHARED_WL_SMS_NUMBERS_DOC[0], SHARED_WL_SMS_NUMBERS_DOC[1]), {
            items: firestoreSanitize(merged),
            updatedAt: serverTimestamp(),
            updatedBy: sync.uid,
          }).catch(e => console.warn('shared testNumbers migration write failed', e));
        }

        // Clear personal testNumbers and remap selectedNumberIds. setStateRaw
        // (not setState) so this migration isn't tagged as a user edit and
        // doesn't fight cloud sync over a state that's just being normalized.
        setStateRaw(prev => {
          const oldWl = prev.whitelistSms || {};
          const remappedSelected = (oldWl.selectedNumberIds || [])
            .map(id => idMap.get(id) || id);
          return {
            ...prev,
            whitelistSms: {
              ...oldWl,
              testNumbers: [],
              selectedNumberIds: Array.from(new Set(remappedSelected)),
            },
          };
        });
      }, [sync.uid, sharedTestNumbersReady, sharedTestNumbers, state]);

      // Apply theme class to <html>
      useEffect(() => {
        const root = document.documentElement;
        if ((state.theme || 'dark') === 'light') root.classList.add('light');
        else root.classList.remove('light');
      }, [state.theme]);

      // Firebase: wait for email/password sign-in, then hydrate from Firestore
      useEffect(() => {
        if (!window.__fb || !window.__fbm) return;
        const { auth } = window.__fb;
        const { onAuthStateChanged, signOut } = window.__fbm;
        const unsub = onAuthStateChanged(auth, async (user) => {
          if (!user) {
            hydratedRef.current = false;
            setSync({ status: 'signed-out', uid: null, email: null, role: 'semi_admin', message: null });
            return;
          }
          if (user.isAnonymous) {
            try { await signOut(auth); } catch {}
            return;
          }
          await hydrateUserState(user);
        });
        return () => unsub();
      }, [hydrateUserState]);

      // Firebase: debounced write-back on state change (only after hydration)
      useEffect(() => {
        if (!window.__fb || !window.__fbm || !hydratedRef.current || !sync.uid) return;
        if (!userEditRef.current) return;
        const ownerUid = stateOwnerUid(state);
        if (ownerUid && ownerUid !== sync.uid) {
          console.warn('Skipped cloud save because local state belongs to a different user.', { ownerUid, syncUid: sync.uid });
          userEditRef.current = false;
          return;
        }
        const { tab, ...rest } = state;
        const persist = { ...rest, stateOwnerUid: sync.uid, stateOwnerEmail: sync.email || '' };
        const snapshotJson = JSON.stringify(persist);
        if (snapshotJson === lastSavedJsonRef.current) return;
        if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
        setSync(s => s.status === 'saving' ? s : { ...s, status: 'saving', message: null });
        saveTimerRef.current = setTimeout(async () => {
          const localUpdatedAt = state.localUpdatedAt || stateEditedAt(state) || Date.now();
          try {
            const { db } = window.__fb;
            const { doc, setDoc, serverTimestamp } = window.__fbm;
            const safe = firestoreSanitize({ ...persist, localUpdatedAt });
            await setDoc(
              doc(db, 'users', sync.uid, 'state', 'main'),
              { ...safe, updatedAt: serverTimestamp() },
              { merge: false }
            );
            lastSavedJsonRef.current = snapshotJson;
            if ((state.localUpdatedAt || stateEditedAt(state) || localUpdatedAt) === localUpdatedAt) userEditRef.current = false;
            setSync(s => ({ ...s, status: 'synced', message: null }));
          } catch (e) {
            console.warn('Firestore save failed', e);
            setSync(s => ({ ...s, status: 'offline', message: syncErrorMessage(e, 'Cloud sync save failed') }));
          }
        }, 400);
        return () => { if (saveTimerRef.current) clearTimeout(saveTimerRef.current); };
      }, [state, sync.uid]);

      const selected = state.sheets.find(s => s.id === state.selectedSheetId) || state.sheets[0];

      const updateSelectedSheet = (next) => setState(s => ({ ...s, sheets: s.sheets.map(x => x.id === next.id ? next : x) }));
      const deleteSelected = async () => {
        if (!selected) return;
        const ok = await confirmDialog({
          title: 'Delete sheet?',
          message: `"${selected.name}" and all of its configuration will be removed. This can't be undone.`,
          confirmText: 'Delete',
          tone: 'danger',
        });
        if (!ok) return;
        setState(s => {
          const remaining = s.sheets.filter(x => x.id !== selected.id);
          return { ...s, sheets: remaining, selectedSheetId: remaining[0]?.id || null };
        });
      };
      const duplicateSelected = () => {
        if (!selected) return;
        const copy = { ...JSON.parse(JSON.stringify(selected)), id: newId(), name: selected.name + ' (copy)' };
        setState(s => ({ ...s, sheets: [...s.sheets, copy], selectedSheetId: copy.id }));
      };

      // Push the stored Client ID into the sync helper on load and whenever it
      // changes (e.g. after Firestore hydrate or an in-app edit). Falls back to
      // the hard-coded constant when the field is empty.
      useEffect(() => {
        googleSheetsSync.setClientId(state.googleSheets?.clientId || GOOGLE_OAUTH_CLIENT_ID);
      }, [state.googleSheets?.clientId]);

      // Save a Client ID entered in the sidebar field. Resets any live
      // connection since a new ID means a different OAuth app.
      const setGoogleClientId = useCallback((id) => {
        const v = String(id || '').trim();
        googleSheetsSync.setClientId(v || GOOGLE_OAUTH_CLIENT_ID);
        googleSheetsSync.disconnect();
        setGoogleConn(c => ({ ...c, connected: false, email: '', busyModule: null, error: null }));
        setState(s => ({ ...s, googleSheets: { ...(s.googleSheets || {}), clientId: v } }));
      }, []);

      // Connect Google Drive (interactive account/consent popup).
      const connectGoogle = useCallback(async () => {
        if (!googleSheetsSync.isConfigured()) {
          setToast({ type: 'err', msg: 'Set GOOGLE_OAUTH_CLIENT_ID in the config to enable Google Sheets sync.' });
          setTimeout(() => setToast(null), 5000);
          return;
        }
        setGoogleConn(c => ({ ...c, busyModule: '__connect__', error: null }));
        try {
          const { email } = await googleSheetsSync.connect();
          setGoogleConn(c => ({ ...c, connected: true, email: email || '', busyModule: null }));
        } catch (e) {
          console.warn('Google connect failed', e);
          setGoogleConn(c => ({ ...c, busyModule: null, error: e.message || String(e) }));
          setToast({ type: 'err', msg: 'Google connect failed: ' + (e.message || e) });
          setTimeout(() => setToast(null), 6000);
        }
      }, []);

      const disconnectGoogle = useCallback(() => {
        googleSheetsSync.disconnect();
        setGoogleConn(c => ({ ...c, connected: false, email: '', busyModule: null, error: null }));
      }, []);

      // Build the module's workbook and push it into its Google Sheet, creating
      // the sheet on first sync and updating it (same id/url) afterwards. On a
      // deleted/inaccessible sheet (404/403) it recreates and re-stores the id.
      const syncModuleToSheets = useCallback(async (moduleId, { interactive = false, confirmOverwrite = false } = {}) => {
        const conf = GOOGLE_SYNC_MODULES[moduleId];
        if (!conf) return;
        if (!googleSheetsSync.isConfigured()) {
          setToast({ type: 'err', msg: 'Set GOOGLE_OAUTH_CLIENT_ID in the config to enable Google Sheets sync.' });
          setTimeout(() => setToast(null), 5000);
          return;
        }
        // One-way sync: the app is the master, so every sync REPLACES the whole
        // sheet. Confirm before overwriting an existing sheet (skipped on the
        // first sync, and on the silent Generate-triggered sync).
        const preId = (stateRef.current || DEFAULT_STATE).googleSheets?.sheetIds?.[moduleId] || '';
        if (confirmOverwrite && preId) {
          const ok = await confirmDialog({
            title: 'Overwrite Google Sheet?',
            message: 'This replaces the entire contents of the linked Google Sheet with the current app data. Any changes made directly in the sheet will be lost — the app is the source of truth.',
            confirmText: 'Overwrite & sync',
            tone: 'danger',
          });
          if (!ok) return;
        }
        setGoogleConn(c => ({ ...c, busyModule: moduleId, error: null }));
        try {
          await googleSheetsSync.getToken({ interactive });
          const latest = stateRef.current || DEFAULT_STATE;
          const { blob, sheets } = await conf.build(latest);
          const existingId = latest.googleSheets?.sheetIds?.[moduleId] || '';
          let fileId = existingId;
          if (existingId) {
            try {
              await googleSheetsSync.updateSheet(existingId, blob);
            } catch (e) {
              if (e && (e.status === 404 || e.status === 403)) {
                fileId = await googleSheetsSync.createSheet(conf.sheetName, blob);
              } else { throw e; }
            }
          } else {
            fileId = await googleSheetsSync.createSheet(conf.sheetName, blob);
          }
          if (fileId && fileId !== existingId) {
            setState(s => ({ ...s, googleSheets: { ...s.googleSheets, sheetIds: { ...s.googleSheets.sheetIds, [moduleId]: fileId } } }));
          }
          // Drive re-pads the converted sheet to a 1000-row grid on every upload,
          // so trim the empty tail back to the totals row after each sync. This
          // is best-effort: the data is already synced, so a trim failure must
          // not fail the sync — only surface the one actionable case (the Sheets
          // API isn't enabled for the OAuth project).
          let trimHint = '';
          try {
            await googleSheetsSync.trimSheetGrids(fileId, sheets);
          } catch (e) {
            console.warn('Trimming synced sheet grids failed', e);
            if (/Google Sheets API has not been used|SERVICE_DISABLED|has not been enabled|accessNotConfigured/i.test(e.message || '')) {
              trimHint = 'Synced, but empty rows below the totals row could not be trimmed — enable the Google Sheets API for your OAuth project, then sync again.';
            }
          }
          setGoogleConn(c => ({
            ...c, connected: true, email: googleSheetsSync.email() || c.email, busyModule: null,
            lastSync: { ...c.lastSync, [moduleId]: { at: Date.now(), ok: true, fileId } },
          }));
          setToast({ type: trimHint ? 'err' : 'ok', msg: trimHint || 'Synced to Google Sheets' });
          setTimeout(() => setToast(null), trimHint ? 8000 : 4000);
          return fileId;
        } catch (e) {
          console.warn('Google Sheets sync failed', e);
          setGoogleConn(c => ({ ...c, busyModule: null, error: e.message || String(e) }));
          setToast({ type: 'err', msg: 'Google Sheets sync failed: ' + (e.message || e) });
          setTimeout(() => setToast(null), 6000);
          throw e;
        }
      }, []);

      const showToast = useCallback((type, msg) => {
        setToast({ type, msg });
        setTimeout(() => setToast(null), 4000);
      }, []);

      const onGenerate = async () => {
        setBusy(true);
        try {
          if (state.module === 'editor') {
            const fn = await generateEditorWorkbook(state);
            setToast({ type: 'ok', msg: `Generated ${fn}` });
          } else if (state.module === 'bmr') {
            const fn = await generateBmrWorkbook(state);
            setToast({ type: 'ok', msg: `Generated ${fn}` });
          } else if (state.module === 'bmr_sms') {
            const fn = await generateBmrSmsWorkbook(state);
            setToast({ type: 'ok', msg: `Generated ${fn}` });
          } else if (state.module === 'whitelist_sms') {
            const fn = await generateWlSmsWorkbook(state, sharedTestNumbers);
            setToast({ type: 'ok', msg: `Generated ${fn}` });
          } else if (state.module === 'rca') {
            const fn = await generateRcaDocx(state);
            setToast({ type: 'ok', msg: `Generated ${fn}` });
          } else {
            await generateWorkbook(state);
            setToast({ type: 'ok', msg: `Generated ${MONTHS[state.month]}_${state.year}_SIP_FCS_Hourly_Record.xlsx` });
          }
          // After a successful download, mirror to Google Sheets when already
          // connected (SIP/FCS, BMR VOIP, BMR SMS only). Silent — never opens a
          // popup here, and a sync error can't undo the successful download.
          if (googleConn.connected && GOOGLE_SYNC_MODULES[state.module]) {
            try { await syncModuleToSheets(state.module, { interactive: false }); }
            catch { /* toast already surfaced by syncModuleToSheets */ }
          }
        } catch (e) {
          console.error(e);
          setToast({ type: 'err', msg: 'Generation failed: ' + (e.message || e) });
        } finally {
          setBusy(false);
          setTimeout(() => setToast(null), 4000);
        }
      };

      const activeCount = state.sheets.filter(s => s.active).length;
      const days = daysInMonth(state.year, state.month);
      const activeRules = state.rules.filter(r => r.enabled).length;

      // Bundle passed to the SIP/FCS, BMR VOIP and BMR SMS sidebars so their
      // GoogleSheetSync panel can set the Client ID, connect and sync.
      const googleClientId = String(state.googleSheets?.clientId || '');
      const googleConfigured = !!(googleClientId.trim() || GOOGLE_OAUTH_CLIENT_ID);
      const googleSyncProps = {
        conn: googleConn, configured: googleConfigured, clientId: googleClientId,
        onSetClientId: setGoogleClientId,
        connect: connectGoogle, disconnect: disconnectGoogle, sync: syncModuleToSheets,
        sheetUrl: googleSheetsSync.sheetUrl,
      };

      const TABS = [
        { id: 'config',  label: 'Config' },
        { id: 'rules',   label: 'Rules' },
        { id: 'preview', label: 'Preview' },
        { id: 'notes',   label: 'Notes' },
      ];

      const BMR_TABS = [
        { id: 'clients', label: 'Clients' },
        { id: 'rules',   label: 'Rules' },
        { id: 'preview', label: 'Preview' },
        { id: 'notes',   label: 'Notes' },
      ];

      const EDITOR_TABS = [
        { id: 'table', label: 'Table' },
        { id: 'rules', label: 'Rules' },
      ];

      const ALL_MODULES = [
        { id: 'sip_fcs', label: 'SIP / FCS' },
        { id: 'bmr',     label: 'BMR VOIP' },
        { id: 'bmr_sms', label: 'BMR SMS' },
        { id: 'whitelist_sms', label: 'Whitelist SMS' },
        { id: 'editor',  label: 'Editor' },
        { id: 'rca',     label: 'RCA' },
        { id: 'image',   label: 'Image Editor', adminOnly: true },
        // Recorder data still lives in the shared/recorder* docs, but the tab
        // is admin-only — signed-out and semi-admin users never see it.
        { id: 'recorder', label: 'Recorder', adminOnly: true },
        { id: 'users',   label: 'Users', adminOnly: true },
      ];
      const role = sync.role || 'semi_admin';
      const isAdmin = role === 'admin';
      // Only enforce role-based filtering once we know who the user is. Before
      // sign-in (or when Firebase is unavailable) show everything so refreshes
      // don't bounce people off their last-active module.
      const roleEnforced = !!window.__fb && !!sync.uid;
      const MODULES = roleEnforced
        ? ALL_MODULES.filter(m => m.adminOnly ? isAdmin : moduleAllowed(m.id, role))
        : ALL_MODULES.filter(m => !m.adminOnly);

      useEffect(() => {
        if (!roleEnforced) return;
        const current = state.module || 'sip_fcs';
        const allowedIds = MODULES.map(m => m.id);
        if (allowedIds.length && !allowedIds.includes(current)) {
          setState(s => ({ ...s, module: allowedIds[0] }));
        }
      }, [roleEnforced, role, state.module]);

      const isEditor = state.module === 'editor';
      const isBmr = state.module === 'bmr';
      const isBmrSms = state.module === 'bmr_sms';
      const isWlSms = state.module === 'whitelist_sms';
      const isRca = state.module === 'rca';
      const isImage = state.module === 'image';
      const isRecorder = state.module === 'recorder';
      const isUsers = state.module === 'users';
      const whitelistSms = wlSmsNormalizeState(state.whitelistSms);
      const setWhitelistSms = useCallback((next) => setState(s => {
        const prev = wlSmsNormalizeState(s.whitelistSms);
        return { ...s, whitelistSms: typeof next === 'function' ? next(prev) : next };
      }), []);
      const whitelistSmsTab = state.whitelistSmsTab || 'content';
      const bmr = state.bmr || DEFAULT_BMR_STATE;
      const setBmr = useCallback((next) => setState(s => {
        const prevBmr = s.bmr || DEFAULT_BMR_STATE;
        return { ...s, bmr: typeof next === 'function' ? next(prevBmr) : next };
      }), []);
      const bmrTab = state.bmrTab || 'clients';
      const bmrSms = state.bmrSms || DEFAULT_BMR_SMS_STATE;
      const setBmrSms = useCallback((next) => setState(s => {
        const prevSms = s.bmrSms || DEFAULT_BMR_SMS_STATE;
        return { ...s, bmrSms: typeof next === 'function' ? next(prevSms) : next };
      }), []);
      const bmrSmsTab = state.bmrSmsTab || 'clients';
      const editor = { ...DEFAULT_EDITOR_STATE, ...(state.editor || {}), cells: editorNormalizeGrid(state.editor?.cells || DEFAULT_EDITOR_STATE.cells) };
      const pushEditorUndo = useCallback((snapshot) => {
        const previous = editorSnapshot(snapshot);
        setEditorUndoStack(stack => {
          const last = stack[stack.length - 1];
          if (last && editorStatesEqual(last, previous)) return stack;
          return [...stack, previous].slice(-60);
        });
      }, []);
      const setEditor = useCallback((next) => {
        const prevEditor = editorSnapshot(editor);
        const nextEditor = editorSnapshot(typeof next === 'function' ? next(prevEditor) : next);
        if (editorStatesEqual(prevEditor, nextEditor)) return;
        pushEditorUndo(prevEditor);
        setState(s => ({ ...s, editor: nextEditor }));
      }, [editor, pushEditorUndo]);
      const editorTab = state.editorTab || 'table';
      const undoEditor = useCallback(() => {
        const previous = editorUndoStack[editorUndoStack.length - 1];
        if (!previous) return;
        setState(s => ({ ...s, editor: editorSnapshot(previous) }));
        setEditorUndoStack(stack => stack.slice(0, -1));
      }, [editorUndoStack]);
      const resetEditor = async () => {
        const ok = await confirmDialog({
          title: 'Reset Editor?',
          message: 'This will clear the pasted table, row colors, header colors, and Editor rules so you can start fresh.',
          confirmText: 'Reset',
          tone: 'danger',
        });
        if (!ok) return;
        setEditor({ ...DEFAULT_EDITOR_STATE });
        setState(s => ({ ...s, editorTab: 'table' }));
      };

      const ModuleSwitch = (
        <ModuleSwitcher
          modules={MODULES}
          current={state.module}
          onSelect={(id) => setState(s => ({ ...s, module: id }))}
        />
      );

      if (isImage) {
        return (
          <div className="flex min-h-screen text-neutral-100">
            <ImageEditorModule
              state={state} setState={setState} sync={sync} onRetrySync={retrySync}
              moduleSwitch={ModuleSwitch} onToast={showToast} />
            {toast && (
              <div className={`fixed bottom-6 right-6 rounded-lg border px-4 py-3 text-sm backdrop-blur shadow-xl anim-toast ${toast.type === 'ok' ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-200' : 'border-red-500/40 bg-red-500/10 text-red-200'}`}>
                {toast.msg}
              </div>
            )}
            {window.__fb && !sync.uid && sync.status !== 'connecting' && <AuthModal />}
            <Dialog dialog={dialog} onResolve={resolveDialog} />
          </div>
        );
      }

      if (isRecorder) {
        return (
          <div className="flex min-h-screen text-neutral-100">
            <RecorderModule
              state={state} setState={setState} sync={sync} onRetrySync={retrySync}
              moduleSwitch={ModuleSwitch} onToast={showToast} confirmDialog={confirmDialog}
              gsheets={googleSyncProps} />
            {toast && (
              <div className={`fixed bottom-6 right-6 rounded-lg border px-4 py-3 text-sm backdrop-blur shadow-xl anim-toast ${toast.type === 'ok' ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-200' : 'border-red-500/40 bg-red-500/10 text-red-200'}`}>
                {toast.msg}
              </div>
            )}
            {window.__fb && !sync.uid && sync.status !== 'connecting' && <AuthModal />}
            <Dialog dialog={dialog} onResolve={resolveDialog} />
          </div>
        );
      }

      if (isUsers) {
        return (
          <div className="flex min-h-screen text-neutral-100">
            <Sidebar state={state} setState={setState} onGenerate={onGenerate} busy={busy} sync={sync} onRetrySync={retrySync} gsheets={googleSyncProps} />
            <main className="flex-1 min-w-0 flex flex-col">
              <header className="app-topbar border-b border-neutral-900 px-4 sm:px-6 lg:px-8 py-4 sm:py-5 flex flex-wrap items-start justify-between gap-4 sticky top-0 bg-[#1c1c1f]/95 backdrop-blur z-10">
                <div className="app-title-block min-w-0">
                  <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-1">Module</div>
                  <h2 className="truncate text-[22px] font-bold tracking-tight">Users</h2>
                </div>
                <div className="app-header-controls flex flex-wrap items-center gap-2 sm:gap-3">
                  {ModuleSwitch}
                </div>
              </header>
              <div className="p-8 max-w-5xl w-full">
                <div className="anim-fade-in">
                  <UsersAdminPanel sync={sync} />
                </div>
              </div>
              <footer className="mt-auto border-t border-neutral-900 px-8 py-4 flex items-center justify-between text-[11px] text-neutral-600">
                <span className="font-mono">collection · userProfiles</span>
                <span>Signed in as {sync.email || '—'} ({ROLE_LABELS[role] || role})</span>
              </footer>
            </main>
            {toast && (
              <div className={`fixed bottom-6 right-6 rounded-lg border px-4 py-3 text-sm backdrop-blur shadow-xl anim-toast ${toast.type === 'ok' ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-200' : 'border-red-500/40 bg-red-500/10 text-red-200'}`}>
                {toast.msg}
              </div>
            )}
            {window.__fb && !sync.uid && sync.status !== 'connecting' && <AuthModal />}
            <Dialog dialog={dialog} onResolve={resolveDialog} />
          </div>
        );
      }

      if (isEditor) {
        const editorGrid = editorNormalizeGrid(editor.cells);
        const activeEditorRules = (editor.rules || []).filter(rule => rule.enabled !== false).length;
        return (
          <div className="flex min-h-screen text-neutral-100">
            <EditorSidebar state={state} setState={setState} onGenerate={onGenerate} busy={busy} sync={sync} onRetrySync={retrySync} />
            <main className="flex-1 min-w-0 flex flex-col">
              <header className="app-topbar border-b border-neutral-900 px-4 sm:px-6 lg:px-8 py-4 sm:py-5 flex flex-wrap items-start justify-between gap-4 sticky top-0 bg-[#1c1c1f]/95 backdrop-blur z-10">
                <div className="app-title-block min-w-0">
                  <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-1">Module</div>
                  <h2 className="truncate text-[22px] font-bold tracking-tight">Editor / {editor.name || 'Pasted Table'}</h2>
                </div>
                <div className="app-pill-row hidden xl:flex items-center gap-2">
                  <Pill>{editorGrid.length} rows</Pill>
                  <Pill>{editorGrid[0].length} columns</Pill>
                  <Pill tone={activeEditorRules ? 'accent' : 'default'}>{activeEditorRules} rules</Pill>
                </div>
                <div className="app-header-controls flex flex-wrap items-center gap-2 sm:gap-3">
                  {ModuleSwitch}
                  <Btn variant="ghost" size="sm" onClick={undoEditor} disabled={!editorUndoStack.length} title="Undo last Editor change"><IconUndo /> Undo</Btn>
                  <Btn variant="ghost" size="sm" onClick={resetEditor}>Reset</Btn>
                  <nav className="app-tab-nav flex flex-wrap items-center gap-0.5 border border-neutral-900 rounded-md p-0.5 bg-[#1a1a1d]">
                    {EDITOR_TABS.map(t => (
                      <button key={t.id} onClick={() => setState(s => ({ ...s, editorTab: t.id }))}
                        className={`px-3 py-1.5 text-xs font-medium rounded transition-colors ${editorTab === t.id ? 'app-tab-active bg-neutral-800 text-neutral-100' : 'app-tab-muted text-neutral-500 hover:text-neutral-200'}`}>
                        {t.label}
                      </button>
                    ))}
                  </nav>
                </div>
              </header>
              <div className="p-8 max-w-[1500px] w-full">
                <div key={editorTab} className="anim-fade-in">
                  {editorTab === 'table' && (
                    <div className="space-y-4">
                      <EditorPastePanel editor={editor} onChange={setEditor} />
                      <EditorHeaderSettings editor={editor} onChange={setEditor} />
                      <EditorTemplatesPanel editor={editor} onChange={setEditor} sync={sync} />
                      <EditorTable editor={editor} onChange={setEditor} />
                    </div>
                  )}
                  {editorTab === 'rules' && (
                    <div className="space-y-4">
                      <EditorRulesPanel editor={editor} onChange={setEditor} />
                      <EditorTable editor={editor} onChange={setEditor} />
                    </div>
                  )}
                </div>
              </div>
              <footer className="mt-auto border-t border-neutral-900 px-8 py-4 flex items-center justify-between text-[11px] text-neutral-600">
                <span className="font-mono">localStorage / key {STORAGE_KEY}</span>
                <span>{editorGrid.length} rows / {editorGrid[0].length} columns / {activeEditorRules} active rules</span>
              </footer>
            </main>
            {toast && (
              <div className={`fixed bottom-6 right-6 rounded-lg border px-4 py-3 text-sm backdrop-blur shadow-xl anim-toast ${toast.type === 'ok' ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-200' : 'border-red-500/40 bg-red-500/10 text-red-200'}`}>
                {toast.msg}
              </div>
            )}
            {window.__fb && !sync.uid && sync.status !== 'connecting' && <AuthModal />}
            <Dialog dialog={dialog} onResolve={resolveDialog} />
          </div>
        );
      }

      if (isBmr) {
        const visibleClients = (bmr.clients || []).filter(c => !c.hidden).length;
        const bmrClientIds = new Set((bmr.clients || []).map(c => c.id));
        const activeBmrRules = ((bmr.targetRules || []).filter(r => r.enabled && bmrClientIds.has(r.clientId)).length) + ((bmr.usageRules || []).filter(r => r.enabled).length);
        return (
          <div className="flex min-h-screen text-neutral-100">
            <BmrSidebar state={state} setState={setState} onGenerate={onGenerate} busy={busy} sync={sync} onRetrySync={retrySync} gsheets={googleSyncProps} />
            <main className="flex-1 min-w-0 flex flex-col">
              <header className="app-topbar border-b border-neutral-900 px-4 sm:px-6 lg:px-8 py-4 sm:py-5 flex flex-wrap items-start justify-between gap-4 sticky top-0 bg-[#1c1c1f]/95 backdrop-blur z-10">
                <div className="app-title-block">
                  <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-1">Module</div>
                  <h2 className="text-[22px] font-bold tracking-tight">BMR VOIP · {bmrTodayString()}</h2>
                </div>
                <div className="app-pill-row hidden xl:flex items-center gap-2">
                  <Pill>{visibleClients} clients</Pill>
                  <Pill>{(bmr.includeDay !== false ? 1 : 0) + (bmr.includeNight !== false ? 1 : 0)} shifts</Pill>
                  <Pill tone={activeBmrRules > 0 ? 'accent' : 'default'}>{activeBmrRules} rules</Pill>
                </div>
                <div className="app-header-controls flex flex-wrap items-center gap-2 sm:gap-3">
                  {ModuleSwitch}
                  <nav className="app-tab-nav flex flex-wrap items-center gap-0.5 border border-neutral-900 rounded-md p-0.5 bg-[#1a1a1d]">
                    {BMR_TABS.map(t => (
                      <button key={t.id} onClick={() => setState(s => ({ ...s, bmrTab: t.id }))}
                        className={`px-3 py-1.5 text-xs font-medium rounded transition-colors ${bmrTab === t.id ? 'app-tab-active bg-neutral-800 text-neutral-100' : 'app-tab-muted text-neutral-500 hover:text-neutral-200'}`}>
                        {t.label}
                      </button>
                    ))}
                  </nav>
                </div>
              </header>

              <div className="p-8 max-w-6xl w-full">
                <div key={bmrTab} className="anim-fade-in">
                  {bmrTab === 'clients' && (
                    <div className="space-y-6">
                      <SectionLabel hint="Drag to reorder · insert at row # · hide rows and columns">Clients</SectionLabel>
                      <BmrClientsEditor bmr={bmr} onChange={setBmr} />
                    </div>
                  )}
                  {bmrTab === 'rules' && (
                    <BmrRulesPanel bmr={bmr} onChange={setBmr} />
                  )}
                  {bmrTab === 'preview' && (
                    <div className="space-y-3">
                      <SectionLabel hint="Live preview · first few clients × first 3 timeslots">Structure preview</SectionLabel>
                      <BmrPreview bmr={bmr} />
                    </div>
                  )}
                  {bmrTab === 'notes' && (
                    <div>
                      <SectionLabel hint="Free-form notes carried with the BMR config">Notes</SectionLabel>
                      <Textarea rows={12} value={bmr.notes || ''} onChange={e => setBmr({ ...bmr, notes: e.target.value })}
                        placeholder={`Notes for BMR…\n\ne.g. Added APNTEL with debit-on-zero rule, hid CYN ACOM for the week.`} />
                    </div>
                  )}
                </div>
              </div>

              <footer className="mt-auto border-t border-neutral-900 px-8 py-4 flex items-center justify-between text-[11px] text-neutral-600">
                <span className="font-mono">localStorage · key {STORAGE_KEY}</span>
                <span>{visibleClients} visible clients · {activeBmrRules} active rules · {bmrTodayString()}</span>
              </footer>
            </main>

            {toast && (
              <div className={`fixed bottom-6 right-6 rounded-lg border px-4 py-3 text-sm backdrop-blur shadow-xl anim-toast ${toast.type === 'ok' ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-200' : 'border-red-500/40 bg-red-500/10 text-red-200'}`}>
                {toast.msg}
              </div>
            )}
            {window.__fb && !sync.uid && sync.status !== 'connecting' && <AuthModal />}
            <Dialog dialog={dialog} onResolve={resolveDialog} />
          </div>
        );
      }

      if (isBmrSms) {
        const visibleRetail = (bmrSms.retailClients || []).filter(c => !c.hidden).length;
        const visibleWholesale = (bmrSms.wholesaleClients || []).filter(c => !c.hidden).length;
        const smsClientIds = new Set([...(bmrSms.retailClients || []), ...(bmrSms.wholesaleClients || [])].map(c => c.id));
        const activeBmrSmsRules = ((bmrSms.targetRules || []).filter(r => r.enabled && smsClientIds.has(r.clientId)).length)
          + ((bmrSms.overdraftRules || []).filter(r => r.enabled).length)
          + ((bmrSms.usageRules || []).filter(r => r.enabled).length);
        const shiftCount = (bmrSms.includeDay !== false ? 1 : 0) + (bmrSms.includeNight !== false ? 1 : 0);
        const sheetCount = Math.max(1, shiftCount) * 2;
        return (
          <div className="flex min-h-screen text-neutral-100">
            <BmrSmsSidebar state={state} setState={setState} onGenerate={onGenerate} busy={busy} sync={sync} onRetrySync={retrySync} gsheets={googleSyncProps} />
            <main className="flex-1 min-w-0 flex flex-col">
              <header className="app-topbar border-b border-neutral-900 px-4 sm:px-6 lg:px-8 py-4 sm:py-5 flex flex-wrap items-start justify-between gap-4 sticky top-0 bg-[#1c1c1f]/95 backdrop-blur z-10">
                <div className="app-title-block">
                  <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-1">Module</div>
                  <h2 className="text-[22px] font-bold tracking-tight">BMR SMS · {bmrTodayString()}</h2>
                </div>
                <div className="app-pill-row hidden xl:flex items-center gap-2">
                  <Pill>{visibleRetail} RES</Pill>
                  <Pill>{visibleWholesale} WHS</Pill>
                  <Pill>{sheetCount} sheets</Pill>
                  <Pill tone={activeBmrSmsRules > 0 ? 'accent' : 'default'}>{activeBmrSmsRules} rules</Pill>
                </div>
                <div className="app-header-controls flex flex-wrap items-center gap-2 sm:gap-3">
                  {ModuleSwitch}
                  <nav className="app-tab-nav flex flex-wrap items-center gap-0.5 border border-neutral-900 rounded-md p-0.5 bg-[#1a1a1d]">
                    {BMR_TABS.map(t => (
                      <button key={t.id} onClick={() => setState(s => ({ ...s, bmrSmsTab: t.id }))}
                        className={`px-3 py-1.5 text-xs font-medium rounded transition-colors ${bmrSmsTab === t.id ? 'app-tab-active bg-neutral-800 text-neutral-100' : 'app-tab-muted text-neutral-500 hover:text-neutral-200'}`}>
                        {t.label}
                      </button>
                    ))}
                  </nav>
                </div>
              </header>

              <div className="p-8 max-w-6xl w-full">
                <div key={bmrSmsTab} className="anim-fade-in">
                  {bmrSmsTab === 'clients' && (
                    <div className="space-y-6">
                      <SectionLabel hint="Retail and Wholesale clients are edited separately">Clients</SectionLabel>
                      <BmrSmsClientsEditor sms={bmrSms} onChange={setBmrSms} />
                    </div>
                  )}
                  {bmrSmsTab === 'rules' && (
                    <BmrSmsRulesPanel sms={bmrSms} onChange={setBmrSms} />
                  )}
                  {bmrSmsTab === 'preview' && (
                    <div className="space-y-3">
                      <SectionLabel hint="Live preview · Retail and Wholesale structures">Structure preview</SectionLabel>
                      <BmrSmsPreview sms={bmrSms} />
                    </div>
                  )}
                  {bmrSmsTab === 'notes' && (
                    <div>
                      <SectionLabel hint="Free-form notes carried with the BMR SMS config">Notes</SectionLabel>
                      <Textarea rows={12} value={bmrSms.notes || ''} onChange={e => setBmrSms({ ...bmrSms, notes: e.target.value })}
                        placeholder={`Notes for BMR SMS...\n\ne.g. Added SMS client, hid old WHS row, adjusted balance rule.`} />
                    </div>
                  )}
                </div>
              </div>

              <footer className="mt-auto border-t border-neutral-900 px-8 py-4 flex items-center justify-between text-[11px] text-neutral-600">
                <span className="font-mono">localStorage · key {STORAGE_KEY}</span>
                <span>{visibleRetail} RES · {visibleWholesale} WHS · {activeBmrSmsRules} active rules · {bmrTodayString()}</span>
              </footer>
            </main>

            {toast && (
              <div className={`fixed bottom-6 right-6 rounded-lg border px-4 py-3 text-sm backdrop-blur shadow-xl anim-toast ${toast.type === 'ok' ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-200' : 'border-red-500/40 bg-red-500/10 text-red-200'}`}>
                {toast.msg}
              </div>
            )}
            {window.__fb && !sync.uid && sync.status !== 'connecting' && <AuthModal />}
            <Dialog dialog={dialog} onResolve={resolveDialog} />
          </div>
        );
      }

      if (isWlSms) {
        const wl = whitelistSms;
        const setWl = setWhitelistSms;
        const numbers = wlSmsEffectiveNumbers(wl, sharedTestNumbers);
        const rowCount = (wl.contents || []).filter(s => s && s.trim()).length * numbers.length;
        const WL_TABS = [
          { id: 'content', label: 'Content' },
          { id: 'utf',     label: 'UTF Check' },
          { id: 'numbers', label: 'Numbers' },
          { id: 'preview', label: 'Preview' },
          { id: 'notes',   label: 'Notes' },
        ];
        return (
          <div className="flex min-h-screen text-neutral-100">
            <WlSmsSidebar state={state} setState={setState} onGenerate={onGenerate} busy={busy} sync={sync} onRetrySync={retrySync} sharedTestNumbers={sharedTestNumbers} />
            <main className="flex-1 min-w-0 flex flex-col">
              <header className="app-topbar border-b border-neutral-900 px-4 sm:px-6 lg:px-8 py-4 sm:py-5 flex flex-wrap items-start justify-between gap-4 sticky top-0 bg-[#1c1c1f]/95 backdrop-blur z-10">
                <div className="app-title-block">
                  <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-1">Module</div>
                  <h2 className="text-[22px] font-bold tracking-tight">Whitelist SMS{wl.senderId ? ` · ${wl.senderId}` : ''}</h2>
                </div>
                <div className="app-pill-row hidden xl:flex items-center gap-2">
                  <Pill>{(wl.contents || []).length} contents</Pill>
                  <Pill>{numbers.length} numbers</Pill>
                  <Pill tone={rowCount ? 'accent' : 'default'}>{rowCount} rows</Pill>
                </div>
                <div className="app-header-controls flex flex-wrap items-center gap-2 sm:gap-3">
                  {ModuleSwitch}
                  <nav className="app-tab-nav flex flex-wrap items-center gap-0.5 border border-neutral-900 rounded-md p-0.5 bg-[#1a1a1d]">
                    {WL_TABS.map(t => (
                      <button key={t.id} onClick={() => setState(s => ({ ...s, whitelistSmsTab: t.id }))}
                        className={`px-3 py-1.5 text-xs font-medium rounded transition-colors ${whitelistSmsTab === t.id ? 'app-tab-active bg-neutral-800 text-neutral-100' : 'app-tab-muted text-neutral-500 hover:text-neutral-200'}`}>
                        {t.label}
                      </button>
                    ))}
                  </nav>
                </div>
              </header>

              <div className="p-8 max-w-6xl w-full">
                <div key={whitelistSmsTab} className="anim-fade-in">
                  {whitelistSmsTab === 'content' && (
                    <div className="space-y-6">
                      <SectionLabel hint="Paste content separated by a dash line — each block becomes one row">Content</SectionLabel>
                      <WlSmsContentPanel wl={wl} onChange={setWl} />
                    </div>
                  )}
                  {whitelistSmsTab === 'utf' && (
                    <div className="space-y-6">
                      <SectionLabel hint="Characters outside the GSM-7 alphabet force UTF (UCS-2) — 70 chars per SMS part instead of 160">Special character finder</SectionLabel>
                      <WlSmsUtfCheckPanel wl={wl} onChange={setWl} />
                    </div>
                  )}
                  {whitelistSmsTab === 'numbers' && (
                    <div className="space-y-6">
                      <SectionLabel hint="Pick the test numbers to include in the output">Test numbers</SectionLabel>
                      <WlSmsNumbersPanel wl={wl} onChange={setWl} sharedTestNumbers={sharedTestNumbers} onChangeSharedTestNumbers={setSharedTestNumbers} sharedReady={sharedTestNumbersReady} />
                    </div>
                  )}
                  {whitelistSmsTab === 'preview' && (
                    <div className="space-y-3">
                      <SectionLabel hint="Same shape as the generated xlsx">Preview</SectionLabel>
                      <WlSmsPreview wl={wl} sharedTestNumbers={sharedTestNumbers} />
                    </div>
                  )}
                  {whitelistSmsTab === 'notes' && (
                    <div>
                      <SectionLabel hint="Free-form notes carried with the Whitelist SMS config">Notes</SectionLabel>
                      <Textarea rows={12} value={wl.notes || ''} onChange={e => setWl({ ...wl, notes: e.target.value })}
                        placeholder={`Notes for whitelist runs...\n\ne.g. Sender ID approval ticket, vendor name, request date.`} />
                    </div>
                  )}
                </div>
              </div>

              <footer className="mt-auto border-t border-neutral-900 px-8 py-4 flex items-center justify-between text-[11px] text-neutral-600">
                <span className="font-mono">localStorage · key {STORAGE_KEY}</span>
                <span>{(wl.contents || []).length} contents · {numbers.length} numbers · {rowCount} rows · {bmrTodayString()}</span>
              </footer>
            </main>

            {toast && (
              <div className={`fixed bottom-6 right-6 rounded-lg border px-4 py-3 text-sm backdrop-blur shadow-xl anim-toast ${toast.type === 'ok' ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-200' : 'border-red-500/40 bg-red-500/10 text-red-200'}`}>
                {toast.msg}
              </div>
            )}
            {window.__fb && !sync.uid && sync.status !== 'connecting' && <AuthModal />}
            <Dialog dialog={dialog} onResolve={resolveDialog} />
          </div>
        );
      }

      if (isRca) {
        const rca = rcaNormalizeState(state.rca);
        const enabledSections = RCA_SECTION_DEFS.filter(def => rca.sections[def.key]?.show !== false).length;
        const signatories = rcaResolveSignatories(state, rca.signatoryIds);
        return (
          <div className="flex min-h-screen text-neutral-100">
            <RcaSidebar state={state} setState={setState} onGenerateDocx={onGenerate} busy={busy} sync={sync} onRetrySync={retrySync} />
            <main className="flex-1 min-w-0 flex flex-col">
              <header className="app-topbar border-b border-neutral-900 px-4 sm:px-6 lg:px-8 py-4 sm:py-5 flex flex-wrap items-start justify-between gap-4 sticky top-0 bg-[#1c1c1f]/95 backdrop-blur z-10">
                <div className="app-title-block">
                  <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-1">Module</div>
                  <h2 className="text-[22px] font-bold tracking-tight">RCA · {rca.customerCid || 'New Incident Report'}</h2>
                </div>
                <div className="app-pill-row hidden xl:flex items-center gap-2">
                  <Pill>{enabledSections} sections</Pill>
                  <Pill tone={signatories.length ? 'accent' : 'muted'}>{signatories.length} signers</Pill>
                  <Pill>{rca.font} · {rca.fontSize}pt</Pill>
                </div>
                <div className="app-header-controls flex flex-wrap items-center gap-2 sm:gap-3">
                  {ModuleSwitch}
                </div>
              </header>

              <div className="p-4 sm:p-6 lg:p-8 w-full max-w-[1700px]">
                <div className="anim-fade-in grid gap-6 xl:grid-cols-[minmax(0,1fr)_minmax(420px,560px)]">
                  <div>
                    <RcaForm state={state} setState={setState} sync={sync} />
                  </div>
                  <div className="xl:sticky xl:top-[88px] xl:self-start xl:max-h-[calc(100vh-110px)] xl:overflow-y-auto">
                    <SectionLabel hint="Updates live as you edit">Preview</SectionLabel>
                    <RcaPreview state={state} />
                  </div>
                </div>
              </div>

              <footer className="mt-auto border-t border-neutral-900 px-8 py-4 flex items-center justify-between text-[11px] text-neutral-600">
                <span className="font-mono">RCA · {rcaFileNameStem(rca)}</span>
                <span>{enabledSections} sections · {signatories.length} signers</span>
              </footer>
            </main>
            {toast && (
              <div className={`fixed bottom-6 right-6 rounded-lg border px-4 py-3 text-sm backdrop-blur shadow-xl anim-toast ${toast.type === 'ok' ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-200' : 'border-red-500/40 bg-red-500/10 text-red-200'}`}>
                {toast.msg}
              </div>
            )}
            {window.__fb && !sync.uid && sync.status !== 'connecting' && <AuthModal />}
            <Dialog dialog={dialog} onResolve={resolveDialog} />
          </div>
        );
      }

      return (
        <div className="flex min-h-screen text-neutral-100">
          <Sidebar state={state} setState={setState} onGenerate={onGenerate} busy={busy} sync={sync} onRetrySync={retrySync} gsheets={googleSyncProps} />
          <UploadedSidebar state={state} setState={setState} />
          <main className="flex-1 min-w-0 flex flex-col">
            {/* Top bar */}
            <header className="app-topbar border-b border-neutral-900 px-4 sm:px-6 lg:px-8 py-4 sm:py-5 flex flex-wrap items-start justify-between gap-4 sticky top-0 bg-[#1c1c1f]/95 backdrop-blur z-10">
              <div className="app-title-block">
                <div className="text-[10px] uppercase tracking-[0.2em] text-neutral-500 mb-1">Report</div>
                <h2 className="text-[22px] font-bold tracking-tight">{MONTHS[state.month]} {state.year}</h2>
              </div>
              <div className="app-pill-row hidden xl:flex items-center gap-2">
                <Pill>{activeCount} sheets</Pill>
                <Pill>{days} days</Pill>
                <Pill tone={activeRules > 0 ? 'accent' : 'default'}>{activeRules} rules</Pill>
              </div>
              <div className="app-header-controls flex flex-wrap items-center gap-2 sm:gap-3">
                {ModuleSwitch}
                <nav className="app-tab-nav flex flex-wrap items-center gap-0.5 border border-neutral-900 rounded-md p-0.5 bg-[#1a1a1d]">
                  {TABS.map(t => (
                    <button key={t.id} onClick={() => setState(s => ({ ...s, tab: t.id }))}
                      className={`px-3 py-1.5 text-xs font-medium rounded transition-colors ${state.tab === t.id ? 'app-tab-active bg-neutral-800 text-neutral-100' : 'app-tab-muted text-neutral-500 hover:text-neutral-200'}`}>
                      {t.label}
                    </button>
                  ))}
                </nav>
              </div>
            </header>

            {/* Body */}
            <div className="p-8 max-w-6xl w-full">
              <div key={state.tab} className="anim-fade-in">
                {state.tab === 'config' && (
                  <div className="space-y-12">
                    <SheetEditor sheet={selected} onChange={updateSelectedSheet} onDelete={deleteSelected} onDuplicate={duplicateSelected} />
                    <div>
                      <SectionLabel hint="Save & share full configurations">Templates</SectionLabel>
                      <TemplateBar state={state} setState={setState} sync={sync} />
                    </div>
                  </div>
                )}
                {state.tab === 'rules' && (
                  <div>
                    <SectionLabel hint="Formula-based · Google-Sheets-compatible">Formatting rules</SectionLabel>
                    <FormattingRules rules={state.rules} sheets={state.sheets} onChange={(r) => setState(s => ({ ...s, rules: r }))} />
                  </div>
                )}
                {state.tab === 'preview' && (
                  <div>
                    <SectionLabel hint={selected ? `Showing first 2 days of ${selected.name}` : ''}>Structure preview</SectionLabel>
                    <Preview sheet={selected} year={state.year} month={state.month} />
                    {selected?.layout === 'hourly' && (
                      <p className="text-[11px] text-neutral-600 mt-3 font-mono">Generated sheet will contain {days} day blocks × {visibleItems(selected.metrics).length} metric rows = {days * visibleItems(selected.metrics).length} data rows{hourlyHasRowSeparator(selected) ? ` + ${Math.max(0, days - 1)} separator rows` : ''}</p>
                    )}
                  </div>
                )}
                {state.tab === 'notes' && (
                  <div>
                    <SectionLabel hint="Saved per generation · optionally exported as CHANGELOG sheet">Monthly notes</SectionLabel>
                    <Textarea rows={12} value={state.changelog} onChange={e => setState(s => ({ ...s, changelog: e.target.value }))}
                      placeholder={`Notes for ${MONTHS[state.month]} ${state.year}…\n\ne.g. Added NEWCLIENT, UNOBANK now only 6AM–10PM, removed old SBC2.`} />
                  </div>
                )}
              </div>
            </div>

            <footer className="mt-auto border-t border-neutral-900 px-8 py-4 flex items-center justify-between text-[11px] text-neutral-600">
              <span className="font-mono">localStorage · key {STORAGE_KEY}</span>
              <span>{activeCount} active sheets · {activeRules} active rules · {days} days</span>
            </footer>
          </main>

          {toast && (
            <div className={`fixed bottom-6 right-6 rounded-lg border px-4 py-3 text-sm backdrop-blur shadow-xl anim-toast ${toast.type === 'ok' ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-200' : 'border-red-500/40 bg-red-500/10 text-red-200'}`}>
              {toast.msg}
            </div>
          )}

          {window.__fb && !sync.uid && sync.status !== 'connecting' && <AuthModal />}

          <Dialog dialog={dialog} onResolve={resolveDialog} />
        </div>
      );
    }

    // Error boundary — if Firestore (or anything else) throws during render,
    // React 18 unmounts the whole tree by default, producing the "black screen
    // after sign-out" symptom. This catches the error and shows a recovery UI
    // with a Reload button instead of nothing.
    class AppErrorBoundary extends React.Component {
      constructor(props) {
        super(props);
        this.state = { error: null };
      }
      static getDerivedStateFromError(error) {
        return { error };
      }
      componentDidCatch(error, info) {
        console.error('App crashed:', error, info);
      }
      render() {
        if (this.state.error) {
          const msg = String(this.state.error?.message || this.state.error || 'Unknown error');
          return (
            <div className="min-h-screen flex items-center justify-center p-8 bg-[#1c1c1f] text-neutral-100">
              <div className="max-w-lg w-full border border-red-500/40 bg-red-500/5 rounded-xl p-6">
                <div className="text-red-300 text-sm font-semibold mb-2">Something went wrong</div>
                <pre className="text-[11px] text-neutral-400 font-mono whitespace-pre-wrap break-words mb-4">{msg}</pre>
                <button
                  onClick={() => location.reload()}
                  className="px-3 py-1.5 text-xs font-semibold rounded bg-blue-500/20 text-blue-200 border border-blue-500/40 hover:bg-blue-500/30">
                  Reload
                </button>
              </div>
            </div>
          );
        }
        return this.props.children;
      }
    }

    // Surface unhandled Firestore SDK assertion errors. They originate inside
    // the SDK's promise queue and React never sees them — but they corrupt the
    // SDK's internal state, so we at least log a clear hint instead of leaving
    // the user staring at a half-broken app.
    window.addEventListener('unhandledrejection', (e) => {
      const msg = String(e.reason?.message || e.reason || '');
      if (msg.includes('INTERNAL ASSERTION FAILED')) {
        console.error('[firestore] Internal SDK assertion — Firestore state is now unreliable until reload. Reason:', e.reason);
      }
    });

    // Wait for the async firebase init (modular initializeFirestore) before
    // mounting React, otherwise App boots with window.__fb still undefined
    // and lands in offline mode for the entire session.
    (async () => {
      try {
        if (window.__firebaseInitPromise) await window.__firebaseInitPromise;
      } catch (e) {
        console.warn('Firebase init promise rejected:', e);
      }
      ReactDOM.createRoot(document.getElementById('root')).render(
        <AppErrorBoundary><App /></AppErrorBoundary>
      );
    })();
  