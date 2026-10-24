/*
 * Pulse engine: reads an M365 Pulse-style .xlsx workbook and builds a standalone HTML report.
 * No AI, no network, no dependencies. Runs in a modern browser and in Node 18+.
 *
 * Sheets are recognised by the columns they contain, never by sheet name, so renamed,
 * reordered, added or missing sheets are handled. Edit CONFIG and SIGNATURES to tune it.
 */
(function (root) {
  'use strict';

  // ---------------------------------------------------------------- settings
  const CONFIG = {
    staleDays: 90,            // no sign-in for this many days = stale account
    expiryCriticalDays: 7,    // app credential expiring within N days = critical
    expiryWarnDays: 30,       // app credential expiring within N days = high
    mfaCriticalBelow: 50,     // % of accounts with MFA
    mfaHighBelow: 90,
    maxGlobalAdmins: 5,
    maxCheckSheetRows: 40,    // small sheets with a status column become "tenant setting checks"
    topRiskControls: 8,
  };

  // A sheet matches a section when every group in `need` has at least one column present.
  // Aliases are compared after lower-casing and stripping everything except a-z and 0-9.
  const SIGNATURES = {
    score:       { need: [['currentscore'], ['maxscore'], ['date', 'createddatetime', 'scoredate']] },
    risks:       { need: [['title', 'controltitle', 'controlname'], ['result', 'status'], ['service', 'workload'], ['maxscore']] },
    actions:     { need: [['displayname', 'title', 'recommendation'], ['maxscore'], ['currentscore']], not: ['date'] },
    licenses:    { need: [['skupartnumber', 'sku', 'skuname'], ['totalunits', 'prepaidunits', 'enabledunits'], ['consumedunits', 'assignedunits']] },
    users:       { need: [['upn', 'userprincipalname'], ['mfaregistered', 'ismfaregistered', 'mfastatus', 'mfaregistration']] },
    apps:        { need: [['applicationname', 'appname', 'displayname'], ['ownercount', 'owners'], ['secretdaysleft', 'secretexpirydate', 'certdaysleft', 'certexpirydate']] },
    cap:         { need: [['policyname', 'displayname', 'name'], ['state'], ['grantcontrols', 'grant', 'grantcontrol']] },
    capBaseline: { need: [['check', 'control'], ['status'], ['requirement', 'enforcedby']] },
    guestGroups: { need: [['userprincipalname', 'upn'], ['groupname', 'groups']], prefer: ['groupdescription', 'dayssincesignin', 'invitationstate', 'rolename'], not: ['assignmenttype'] },
    admins:      { need: [['userprincipalname', 'upn'], ['adminroles', 'roles', 'directoryroles']] },
    gaAudit:     { need: [['assignmenttype'], ['principaltype'], ['userprincipalname', 'displayname']] },
    riskyUsers:  { need: [['risklevel'], ['riskstate'], ['userprincipalname', 'userdisplayname']] },
    mfaSettings: { need: [['mfatextmessageallowed', 'smsallowed', 'textmessageallowed'], ['allowapppasswords', 'mfaphonecallallowed', 'phonecallallowed']] },
    tenant:      { need: [['tenantdisplayname', 'tenantname', 'organizationname']] },
    domain:      { need: [['defaultdomain'], ['isverified', 'authenticationtype']] },
  };

  const SECTION_LABELS = {
    score: 'Secure Score history', risks: 'Secure Score controls', actions: 'Secure Score recommendations',
    licenses: 'Licenses', users: 'Users and sign-in methods', apps: 'App registrations', cap: 'Conditional Access policies',
    capBaseline: 'Conditional Access baseline', guestGroups: 'Guest group memberships', admins: 'Admin roles',
    gaAudit: 'Global Admin assignments', riskyUsers: 'Risky users', mfaSettings: 'MFA service settings',
    tenant: 'Tenant details', domain: 'Domains',
  };

  const SKU_NAMES = {
    SPE_E5: 'Microsoft 365 E5', SPE_E3: 'Microsoft 365 E3', ENTERPRISEPACK: 'Office 365 E3', ENTERPRISEPREMIUM: 'Office 365 E5',
    SPB: 'Microsoft 365 Business Premium', O365_BUSINESS_PREMIUM: 'Microsoft 365 Business Standard',
    O365_BUSINESS_ESSENTIALS: 'Microsoft 365 Business Basic', FLOW_FREE: 'Power Automate Free',
    Microsoft_Entra_Suite: 'Microsoft Entra Suite', AAD_PREMIUM: 'Entra ID P1', AAD_PREMIUM_P2: 'Entra ID P2',
    EMS: 'EMS E3', EMSPREMIUM: 'EMS E5', POWER_BI_STANDARD: 'Power BI (free)', TEAMS_EXPLORATORY: 'Teams Exploratory',
  };

  const SERVICE_NAMES = {
    AzureAD: 'Entra ID', EXO: 'Exchange Online', MDO: 'Defender for Office', SPO: 'SharePoint & OneDrive',
    'MS Teams': 'Teams', Teams: 'Teams', MIP: 'Purview', Admincenter: 'Admin center', FORMS: 'Forms', SWAY: 'Sway',
  };

  // ---------------------------------------------------------------- small helpers
  const norm = s => String(s ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const num = v => { const m = String(v ?? '').replace(/,/g, '').match(/-?\d+(\.\d+)?/); return m ? parseFloat(m[0]) : NaN; };
  const isNone = v => /^(|none|n\/a|na|null|not configured|notconfigured|-|\[\]|0 items)$/i.test(String(v ?? '').trim());
  const plural = (n, w, p) => `${n} ${n === 1 ? w : (p || w + 's')}`;
  const fmt = n => (Math.round(n * 10) / 10).toLocaleString('en-US');
  const listNames = (arr, max = 8) => arr.length <= max ? arr.join(', ') : `${arr.slice(0, max).join(', ')} and ${arr.length - max} more`;
  const sevRank = { crit: 0, high: 1, med: 2, info: 3, ok: 4 };

  // pass / fail / info from the many spellings exports use ("NOT-COMPLIANT", "? No - ...", "Risk Detected", "Yes")
  function classify(v) {
    const s = String(v ?? '').trim().replace(/^[^a-z0-9]+/i, '');
    if (/^(not|non)[\s_-]*compliant|^no\b|^fail|^risk|^false\b|^warning|^missing/i.test(s)) return 'fail';
    if (/^compliant|^yes\b|^pass|^true\b|^good|^enabled\b|^ok\b|^configured\b/i.test(s)) return 'pass';
    return 'info';
  }

  // ---------------------------------------------------------------- xlsx reader
  const td = new TextDecoder('utf-8');
  const decodeXml = s => s.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|quot|apos|amp);/gi, (m, e) => {
    if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    return { lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' }[e.toLowerCase()];
  });
  const textRuns = xml => decodeXml([...xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(m => m[1]).join(''));
  const attrs = s => Object.fromEntries([...s.matchAll(/([\w:.-]+)="([^"]*)"/g)].map(m => [m[1], decodeXml(m[2])]));

  function listZip(b) {
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    let eocd = -1;
    for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error('This file is not a valid .xlsx workbook.');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const out = new Map();
    for (let i = 0; i < count && dv.getUint32(p, true) === 0x02014b50; i++) {
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const name = td.decode(b.subarray(p + 46, p + 46 + nlen)).replace(/\\/g, '/');
      out.set(name.toLowerCase(), { method: dv.getUint16(p + 10, true), csize: dv.getUint32(p + 20, true), off: dv.getUint32(p + 42, true) });
      p += 46 + nlen + xlen + clen;
    }
    return out;
  }

  async function inflate(b, e) {
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    const start = e.off + 30 + dv.getUint16(e.off + 26, true) + dv.getUint16(e.off + 28, true);
    const data = b.subarray(start, start + e.csize);
    if (e.method === 0) return data;
    if (e.method !== 8) throw new Error('Unsupported compression in workbook.');
    const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  const colIndex = ref => { let n = 0; for (const ch of ref) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };

  function parseRows(xml, shared, maxRows = Infinity) {
    const rows = [];
    const re = /<row\b[^>]*>([\s\S]*?)<\/row>/g;
    let m;
    while ((m = re.exec(xml)) && rows.length < maxRows) {
      const cells = [];
      for (const c of m[1].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const a = attrs(c[1]), body = c[2] || '';
        const v = (/<v>([\s\S]*?)<\/v>/.exec(body) || [])[1];
        let val = '';
        if (a.t === 's') val = v != null ? shared[+v] ?? '' : '';
        else if (a.t === 'inlineStr') val = textRuns(body);
        else if (a.t === 'b') val = v === '1' ? 'TRUE' : 'FALSE';
        else if (a.t === 'e') val = '';
        else val = v != null ? decodeXml(v) : '';
        const ref = /^([A-Z]+)/.exec(a.r || '');
        cells[ref ? colIndex(ref[1]) : cells.length] = val;
      }
      rows.push(cells);
    }
    return rows;
  }

  async function readWorkbook(input) {
    const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
    const zip = listZip(bytes);
    const read = async path => { const e = zip.get(path.toLowerCase()); return e ? td.decode(await inflate(bytes, e)) : null; };

    const wbXml = await read('xl/workbook.xml');
    if (!wbXml) throw new Error('No workbook found inside this file. Is it an Excel .xlsx file?');
    const relsXml = (await read('xl/_rels/workbook.xml.rels')) || '';
    const rels = {};
    for (const m of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) { const a = attrs(m[1]); rels[a.Id] = a.Target; }
    const ssXml = (await read('xl/sharedStrings.xml')) || '';
    const shared = [...ssXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map(m => textRuns(m[1]));

    const sheets = [];
    for (const m of wbXml.matchAll(/<sheet\b([^>]*?)\/?>/g)) {
      const a = attrs(m[1]);
      const rid = Object.entries(a).find(([k]) => /(^|:)id$/.test(k));
      let target = rid && rels[rid[1]];
      if (!target) continue;
      target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      const xml = await read(target);
      if (!xml) continue;
      // header = first row with at least two filled cells
      const top = parseRows(xml, shared, 5);
      const hIdx = top.findIndex(r => r.filter(v => v != null && String(v).trim() !== '').length >= 2);
      if (hIdx < 0) { sheets.push({ name: a.name, headers: [], rowCount: 0, rows: () => [] }); continue; }
      const headers = Array.from(top[hIdx], v => String(v ?? '').trim());
      let rowCount = -1, pos = 0;
      while ((pos = xml.indexOf('<row', pos)) !== -1) { rowCount++; pos += 4; }
      rowCount = Math.max(0, rowCount - hIdx);
      let cache = null;
      const sheet = {
        name: a.name, headers, rowCount,
        cols: new Map(headers.map(h => [norm(h), h]).filter(([k]) => k)),
        rows() {
          if (!cache) {
            cache = parseRows(xml, shared).slice(hIdx + 1)
              .filter(r => r.some(v => v != null && String(v).trim() !== ''))
              .map(r => Object.fromEntries(headers.map((h, i) => [h, r[i] == null ? '' : String(r[i])]).filter(([h]) => h)));
          }
          return cache;
        },
      };
      sheets.push(sheet);
    }
    return { sheets };
  }

  // ---------------------------------------------------------------- sheet detection
  const col = (sheet, aliases) => { for (const a of aliases) if (sheet.cols.has(a)) return sheet.cols.get(a); return null; };

  function detect(wb) {
    const found = {};
    for (const [key, sig] of Object.entries(SIGNATURES)) {
      let best = null;
      for (const s of wb.sheets) {
        if (!s.cols || !s.rowCount) continue;
        if (!sig.need.every(g => g.some(a => s.cols.has(a)))) continue;
        if (sig.not && sig.not.some(a => s.cols.has(a))) continue;
        const extra = sig.need.reduce((n, g) => n + g.filter(a => s.cols.has(a)).length, 0) + (sig.prefer || []).filter(a => s.cols.has(a)).length;
        if (!best || extra > best.extra || (extra === best.extra && s.rowCount > best.sheet.rowCount)) best = { sheet: s, extra };
      }
      if (best) found[key] = best.sheet;
    }
    return found;
  }

  // Generic pass/fail sheets not claimed by a named section.
  function detectChecks(wb, used) {
    const groups = [];
    for (const s of wb.sheets) {
      if (!s.cols || !s.rowCount || used.has(s) || s.rowCount > CONFIG.maxCheckSheetRows) continue;
      const statusCol = col(s, ['compliant', 'compliance', 'result', 'status']);
      if (!statusCol) continue;
      const titleCol = col(s, ['controltitle', 'control', 'check', 'attribute', 'setting', 'displayname', 'description', 'service', 'name']);
      const detailCol = col(s, ['finding', 'findings', 'details']);
      const items = s.rows().map(r => {
        const status = r[statusCol];
        const state = classify(status);
        let title = titleCol ? r[titleCol] : '';
        if (!title || title.length > 120) title = prettySheet(s.name);
        else if (/^[a-z]+[A-Z]\w*$/.test(title)) title = `${prettySheet(s.name)}: ${title}`;
        let detail = detailCol ? r[detailCol] : '';
        if (!detail) { const rest = String(status).split(/\s+-\s+/).slice(1).join(' - '); if (rest) detail = rest; }
        if (!detail) {
          detail = Object.entries(r)
            .filter(([k, v]) => k !== statusCol && k !== titleCol && !/description|^id$|mail|phone/i.test(k) && String(v).length && String(v).length <= 60)
            .slice(0, 2).map(([k, v]) => `${k}: ${v}`).join(' · ');
        }
        return { title, detail, state, source: s.name };
      }).filter(i => i.state !== 'info');
      if (items.length) groups.push({ sheet: s, items });
    }
    return groups;
  }

  const prettySheet = n => String(n).replace(/^\d+_+/, '').replace(/_+/g, ' ').replace(/([a-z])([A-Z])/g, '$1 $2').replace(/\s+/g, ' ').replace(/\bTen[a-z]*nt\b/gi, 'Tenant').trim();

  // ---------------------------------------------------------------- dates
  function parseDate(v) {
    const s = String(v ?? '').trim();
    let m;
    if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/))) { const y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return new Date(Date.UTC(y, +m[2] - 1, +m[1])); }
    if (/^\d{5}(\.\d+)?$/.test(s)) return new Date(Date.UTC(1899, 11, 30) + parseFloat(s) * 86400000);
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }
  const fmtDate = d => d ? d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' }) : '';

  // ---------------------------------------------------------------- analysis
  function analyse(wb, fileName) {
    const S = detect(wb);
    const used = new Set(Object.values(S));
    const findings = [];
    const add = (sev, area, title, detail, source) => findings.push({ sev, area, title, detail, source });
    const model = { S, findings, fileName, stats: {} };

    // Tenant name and report date
    const t = S.tenant && S.tenant.rows()[0];
    const lic0 = S.licenses && S.licenses.rows()[0];
    const fileTenant = (String(fileName || '').match(/Report_([^_]+)_/i) || [])[1];
    model.tenant = (t && t[col(S.tenant, ['tenantdisplayname', 'tenantname', 'organizationname'])]) ||
      (lic0 && lic0[col(S.licenses, ['accountname'])] ) || fileTenant || 'Microsoft 365 tenant';
    model.domain = S.domain ? S.domain.rows()[0][col(S.domain, ['defaultdomain'])] : '';

    // Secure Score
    if (S.score) {
      const s = S.score;
      const cD = col(s, ['date', 'createddatetime', 'scoredate']), cC = col(s, ['currentscore']), cM = col(s, ['maxscore']), cP = col(s, ['scorepercent', 'percent']);
      let pts = s.rows().map((r, i) => {
        const cur = num(r[cC]), max = num(r[cM]);
        const pct = cP && !isNaN(num(r[cP])) ? num(r[cP]) : (cur / max) * 100;
        return { i, date: parseDate(r[cD]), raw: r[cD], cur, max, pct };
      }).filter(p => !isNaN(p.pct));
      if (pts.length && pts.every(p => p.date)) pts.sort((a, b) => b.date - a.date);
      if (pts.length) {
        const latest = pts[0], oldest = pts[pts.length - 1];
        model.score = { latest, oldest, series: pts.slice().reverse(), days: pts.length };
        if (latest.pct < 30) add('crit', 'Secure Score', `Secure Score is ${fmt(latest.pct)}%`, `${fmt(latest.cur)} of ${fmt(latest.max)} points.`, s.name);
        else if (latest.pct < 60) add('high', 'Secure Score', `Secure Score is ${fmt(latest.pct)}%`, `${fmt(latest.cur)} of ${fmt(latest.max)} points. ${Math.abs(latest.pct - oldest.pct) < 1 ? 'Almost no change' : (latest.pct > oldest.pct ? 'Up ' : 'Down ') + fmt(Math.abs(latest.pct - oldest.pct)) + ' percentage points'} across the ${plural(pts.length, 'day')} of history.`, s.name);
      }
    }
    const reportDate = (String(fileName || '').match(/(\d{4}-\d{2}-\d{2})/) || [])[1];
    model.reportDate = reportDate ? fmtDate(parseDate(reportDate)) : (model.score && model.score.latest.date ? fmtDate(model.score.latest.date) : fmtDate(new Date()));

    // Secure Score controls (risks)
    if (S.risks) {
      const s = S.risks;
      const cT = col(s, ['title', 'controltitle', 'controlname']), cR = col(s, ['result', 'status']), cS = col(s, ['service', 'workload']), cM = col(s, ['maxscore']);
      const all = s.rows().map(r => ({ title: r[cT], state: classify(r[cR]), service: r[cS] || 'Other', max: num(r[cM]) || 0 }));
      const family = svc => /mdatp|endpoint|^mde$/i.test(svc) ? 'endpoint' : /azure ?atp|^mdi$|defender for identity/i.test(svc) ? 'identity'
        : /^mda_|mcas|appg|cloud ?app/i.test(svc) ? 'cloudapps' : 'core';
      const open = all.filter(x => x.state === 'fail');
      const fam = { endpoint: 0, identity: 0, cloudapps: 0, core: 0 };
      open.forEach(x => fam[family(x.service)]++);
      const core = open.filter(x => family(x.service) === 'core');
      const byService = {};
      core.forEach(x => (byService[x.service] = byService[x.service] || []).push(x));
      model.risks = {
        total: all.length, open: open.length, passing: all.filter(x => x.state === 'pass').length, fam,
        services: Object.entries(byService).map(([svc, items]) => ({ name: SERVICE_NAMES[svc] || svc, items: items.sort((a, b) => b.max - a.max) }))
          .sort((a, b) => b.items.length - a.items.length),
        top: core.slice().sort((a, b) => b.max - a.max).slice(0, CONFIG.topRiskControls),
      };
    }

    // Secure Score recommendations (actions)
    if (S.actions) {
      const s = S.actions;
      const cN = col(s, ['displayname', 'title', 'recommendation']), cM = col(s, ['maxscore']), cC = col(s, ['currentscore']), cR = col(s, ['result', 'status']);
      model.actions = s.rows().map(r => {
        const max = num(r[cM]) || 0, cur = Math.max(0, num(r[cC]) || 0);
        return { title: String(r[cN]).trim(), max, cur, gap: Math.max(0, max - cur), state: cR ? classify(r[cR]) : (cur >= max ? 'pass' : 'fail') };
      }).sort((a, b) => b.gap - a.gap);
    }

    // Users and MFA
    let userIndex = new Map();
    if (S.users) {
      const s = S.users;
      const c = {
        name: col(s, ['user', 'displayname', 'name']), upn: col(s, ['upn', 'userprincipalname']), type: col(s, ['usertype']),
        admin: col(s, ['adminaccount', 'isadmin', 'admin']), mfa: col(s, ['mfaregistered', 'ismfaregistered', 'mfastatus', 'mfaregistration']),
        days: col(s, ['dayssincesignin', 'dayssincelastsignin', 'dayssincesuccessfulsignin']), last: col(s, ['lastsignin', 'lastsignindatetime', 'lastsuccessfulsignin']),
        methods: col(s, ['authmethods', 'methodsregistered', 'authenticationmethods']), state: col(s, ['state', 'accountenabled', 'accountstate']),
      };
      const users = s.rows().map(r => {
        const upn = r[c.upn] || '';
        const name = (c.name && r[c.name]) || upn;
        const mfaN = norm(r[c.mfa]);
        const daysRaw = c.days ? r[c.days] : '';
        const lastRaw = c.last ? r[c.last] : '';
        return {
          name, upn,
          guest: /guest/i.test(c.type ? r[c.type] : '') || /#EXT#/i.test(upn),
          admin: c.admin ? ['admin', 'yes', 'true', '1'].includes(norm(r[c.admin])) : false,
          mfa: ['mfaregistered', 'registered', 'yes', 'true', '1', 'enabled', 'enforced'].includes(mfaN),
          service: /sync|synchroni[sz]ation|service account|^svc[-_]/i.test(name + ' ' + upn),
          days: String(daysRaw).trim() === '' ? NaN : num(daysRaw),
          never: isNone(lastRaw) && String(daysRaw).trim() === '',
          weakMethod: c.methods ? /\bsms\b|voice|phone ?call|text message/i.test(r[c.methods]) : false,
          disabled: c.state ? /disabled|false|blocked/i.test(r[c.state]) : false,
        };
      });
      users.forEach(u => userIndex.set(u.upn.toLowerCase(), u));
      const people = users.filter(u => !u.service);
      const withMfa = people.filter(u => u.mfa);
      const pct = people.length ? (withMfa.length / people.length) * 100 : 100;
      model.users = { all: users, people, withMfa: withMfa.length, pct };
      const noMfa = people.filter(u => !u.mfa).map(u => u.name);
      if (pct < CONFIG.mfaCriticalBelow) add('crit', 'Identity', `MFA coverage is ${Math.round(pct)}%`, `Only ${withMfa.length} of ${people.length} accounts have MFA registered. Without MFA: ${listNames(noMfa)}.`, s.name);
      else if (pct < CONFIG.mfaHighBelow) add('high', 'Identity', `MFA coverage is ${Math.round(pct)}%`, `${withMfa.length} of ${people.length} accounts have MFA registered. Without MFA: ${listNames(noMfa)}.`, s.name);

      const adminsNoMfa = people.filter(u => u.admin && !u.mfa).map(u => u.name);
      if (adminsNoMfa.length) add('crit', 'Identity', `${plural(adminsNoMfa.length, 'admin account')} without MFA`, `${listNames(adminsNoMfa)}. Admin accounts should always have strong MFA.`, s.name);
      const adminsWeak = people.filter(u => u.admin && u.weakMethod).map(u => u.name);
      if (adminsWeak.length) add('high', 'Identity', 'Admins use SMS or voice for MFA', `${listNames(adminsWeak)} ${adminsWeak.length === 1 ? 'has' : 'have'} SMS or voice registered. Move admins to phishing-resistant methods such as passkeys or FIDO2 keys.`, s.name);

      const staleGuests = people.filter(u => u.guest && !u.disabled && (u.days > CONFIG.staleDays || u.never));
      const staleMembers = people.filter(u => !u.guest && !u.disabled && (u.days > CONFIG.staleDays || u.never));
      model.users.staleGuests = staleGuests;
      if (staleMembers.length) add('med', 'Identity', `${plural(staleMembers.length, 'member account')} stale or never used`, `No sign-in for over ${CONFIG.staleDays} days, or never signed in: ${listNames(staleMembers.map(u => u.name))}. Disable or remove accounts nobody uses.`, s.name);
    }

    // Guests in groups
    if (S.guestGroups) {
      const s = S.guestGroups;
      const cU = col(s, ['userprincipalname', 'upn']), cG = col(s, ['groupname', 'groups']), cN = col(s, ['displayname', 'name']), cD = col(s, ['dayssincesignin']);
      const inGroups = s.rows().filter(r => !isNone(r[cG]) && (/#EXT#/i.test(r[cU]) || (userIndex.get(String(r[cU]).toLowerCase()) || {}).guest));
      const stale = inGroups.filter(r => {
        const u = userIndex.get(String(r[cU]).toLowerCase());
        const d = cD ? num(r[cD]) : u ? u.days : NaN;
        return d > CONFIG.staleDays || (u && u.never);
      });
      if (stale.length) add('high', 'Guests', `${plural(stale.length, 'stale guest')} still in groups`, stale.map(r => `${r[cN] || r[cU]} is in ${r[cG]}`).join('; ') + `. No sign-in for over ${CONFIG.staleDays} days.`, s.name);
      else if (inGroups.length) add('info', 'Guests', `${plural(inGroups.length, 'guest')} in groups`, inGroups.map(r => `${r[cN] || r[cU]} is in ${r[cG]}`).join('; ') + '.', s.name);
    } else if (model.users && model.users.staleGuests.length) {
      add('med', 'Guests', `${plural(model.users.staleGuests.length, 'stale guest account')}`, `${listNames(model.users.staleGuests.map(u => u.name))}. Remove guests who no longer need access.`, S.users.name);
    }

    // Global admins
    let gas = null;
    if (S.gaAudit) {
      const s = S.gaAudit;
      const cU = col(s, ['userprincipalname']), cN = col(s, ['displayname']), cA = col(s, ['assignmenttype']), cP = col(s, ['principaltype']);
      gas = s.rows().filter(r => !cP || /user/i.test(r[cP])).map(r => ({ upn: r[cU] || '', name: r[cN] || r[cU], eligible: /eligible/i.test(r[cA]) }));
    } else if (S.admins) {
      const s = S.admins;
      const cU = col(s, ['userprincipalname', 'upn']), cN = col(s, ['displayname', 'name']), cR = col(s, ['adminroles', 'roles', 'directoryroles']);
      gas = s.rows().filter(r => /global administrator/i.test(r[cR])).map(r => ({ upn: r[cU] || '', name: r[cN] || r[cU], eligible: false }));
    }
    if (gas) {
      const active = gas.filter(g => !g.eligible);
      model.globalAdmins = { total: gas.length, active: active.length };
      const src = (S.gaAudit || S.admins).name;
      if (active.length > CONFIG.maxGlobalAdmins) add('med', 'Admins', `${active.length} permanent Global Admins`, `Microsoft recommends fewer than ${CONFIG.maxGlobalAdmins}. Use PIM for just-in-time access.`, src);
      if (active.length < 2) add('med', 'Admins', 'Fewer than two Global Admins', 'Keep at least two, including an emergency access account.', src);
      const dormant = active.filter(g => { const u = userIndex.get(g.upn.toLowerCase()); return u && (u.days > CONFIG.staleDays || u.never); });
      if (dormant.length) add('high', 'Admins', `${plural(dormant.length, 'dormant Global Admin')}`, `${listNames(dormant.map(g => `${g.name} (${g.upn})`))} ${dormant.length === 1 ? 'has' : 'have'} not signed in for over ${CONFIG.staleDays} days. Remove unused admins, or make them documented break-glass accounts excluded from CA.`, src);
    }

    // Risky users
    if (S.riskyUsers) {
      const s = S.riskyUsers;
      const cU = col(s, ['userdisplayname', 'userprincipalname']), cS = col(s, ['riskstate']), cL = col(s, ['risklevel']);
      const atRisk = s.rows().filter(r => /atrisk|confirmedcompromised/i.test(norm(r[cS])));
      if (atRisk.length) add(atRisk.some(r => /high|compromised/i.test(r[cL] + r[cS])) ? 'crit' : 'high', 'Identity', `${plural(atRisk.length, 'user')} flagged as at risk`,
        atRisk.map(r => `${r[cU]} (${r[cL]} risk)`).join(', ') + '. Investigate and remediate in Identity Protection.', s.name);
    }

    // MFA service settings
    if (S.mfaSettings) {
      const s = S.mfaSettings, r = s.rows()[0] || {};
      const on = aliases => { const k = col(s, aliases); return k ? /^(1|true|yes|enabled)$/i.test(String(r[k]).trim()) : false; };
      const weak = [on(['mfatextmessageallowed', 'smsallowed', 'textmessageallowed']) && 'text message', on(['mfaphonecallallowed', 'phonecallallowed']) && 'phone call'].filter(Boolean);
      if (weak.length) add('med', 'Identity', `${weak.join(' and ')} MFA allowed tenant-wide`, 'These methods can be intercepted or SIM-swapped. Prefer the Authenticator app, passkeys or FIDO2 keys.', s.name);
      if (on(['allowapppasswords'])) add('med', 'Identity', 'App passwords are allowed', 'App passwords bypass MFA. Turn them off unless a legacy app still needs them.', s.name);
    }

    // Conditional Access
    if (S.cap) model.cap = analyseCA(S.cap, add);
    if (S.capBaseline) {
      const s = S.capBaseline;
      const cC = col(s, ['check', 'control']), cS = col(s, ['status']);
      model.capBaseline = s.rows().map(r => ({ title: r[cC], state: classify(r[cS]) })).filter(x => x.state !== 'info');
    }

    // App registrations
    if (S.apps) model.apps = analyseApps(S.apps, add);

    // Licenses
    if (S.licenses) {
      const s = S.licenses;
      const cK = col(s, ['skupartnumber', 'sku', 'skuname']), cT = col(s, ['totalunits', 'prepaidunits', 'enabledunits']), cC = col(s, ['consumedunits', 'assignedunits']);
      model.licenses = s.rows().map(r => ({ sku: r[cK], name: SKU_NAMES[r[cK]] || r[cK], total: num(r[cT]) || 0, used: num(r[cC]) || 0 }));
    }

    // Generic tenant setting checks
    model.checks = detectChecks(wb, used);
    model.checks.forEach(g => used.add(g.sheet));

    model.mapping = Object.keys(SIGNATURES).map(k => ({ key: k, label: SECTION_LABELS[k], sheet: S[k] ? S[k].name : null }));
    model.unused = wb.sheets.filter(s => !used.has(s)).map(s => s.name);
    findings.sort((a, b) => sevRank[a.sev] - sevRank[b.sev]);
    return model;
  }

  function analyseCA(s, add) {
    const policies = s.rows().map(row => {
      const f = {};
      for (const [k, v] of Object.entries(row)) f[norm(k)] = v;
      // exports that pack settings as "Key : Value" lines inside a cell
      for (const v of Object.values(row)) for (const line of String(v).split(/\r?\n/)) {
        const m = line.match(/^\s*([A-Za-z][^:]{1,40}?)\s*:\s*(.*)$/);
        if (m && !(norm(m[1]) in f)) f[norm(m[1])] = m[2].trim();
      }
      const get = (...keys) => { for (const k of keys) if (k in f) return f[k]; return undefined; };
      const state = String(get('state') || '');
      return {
        name: get('policyname', 'displayname', 'name'),
        enabled: /^(enabled|on)$/i.test(state.trim()),
        reportOnly: /report/i.test(state),
        includeUsers: get('includeusers'), excludeUsers: get('excludeusers'), excludeGroups: get('excludegroups'),
        includeGroups: get('includegroups'), includeRoles: get('includeroles'), includeGuests: get('includeguests', 'includeguestsorexternalusers'),
        apps: get('includeapps', 'applications', 'includeapplications', 'cloudapps'),
        userActions: get('useractions'), authContexts: get('authcontexts', 'authenticationcontext', 'authenticationcontexts'),
        userRisk: get('userrisk', 'userrisklevels'), signInRisk: get('signinrisk', 'signinrisklevels'),
        authFlows: get('authenticationflows', 'authflows'),
        authStrength: get('authenticationstrength', 'authstrength'),
        grant: get('builtincontrols', 'grantcontrols', 'grant'),
      };
    });
    const live = policies.filter(p => p.enabled);
    const src = s.name;

    const appsKnown = policies.some(p => p.apps !== undefined);
    if (appsKnown) {
      const noTarget = live.filter(p => isNone(p.apps) && isNone(p.userActions) && isNone(p.authContexts));
      if (noTarget.length) {
        const someHaveApps = policies.some(p => !isNone(p.apps));
        add(someHaveApps ? 'crit' : 'high', 'Conditional Access', `${plural(noTarget.length, 'enabled CA policy', 'enabled CA policies')} ${noTarget.length === 1 ? 'targets' : 'target'} no cloud apps`,
          `${listNames(noTarget.map(p => p.name))}. A policy with no target resources never applies. Confirm in the Entra portal, because some exports leave this field empty.`, src);
      }
    }
    const noUsers = live.filter(p => p.includeUsers !== undefined && isNone(p.includeUsers) && isNone(p.includeGroups) && isNone(p.includeRoles) && isNone(p.includeGuests));
    if (noUsers.length) add('high', 'Conditional Access', `${plural(noUsers.length, 'CA policy', 'CA policies')} ${noUsers.length === 1 ? 'targets' : 'target'} no users`, `${listNames(noUsers.map(p => p.name))} ${noUsers.length === 1 ? 'includes' : 'include'} no users, groups or roles, so ${noUsers.length === 1 ? 'it never applies' : 'they never apply'}.`, src);

    const allNoExcl = live.filter(p => /all/i.test(p.includeUsers || '') && isNone(p.excludeUsers) && isNone(p.excludeGroups));
    if (allNoExcl.length) add('high', 'Conditional Access', 'No break-glass exclusion on all-user policies', `${listNames(allNoExcl.map(p => p.name))} cover all users with no exclusions. One mistake could lock every admin out. Exclude an emergency access account.`, src);

    const excl = live.filter(p => !isNone(p.excludeUsers));
    if (excl.length) add('med', 'Conditional Access', 'Review CA user exclusions', excl.map(p => `${p.name} excludes ${p.excludeUsers}`).join('; ') + '. Make sure each exclusion is a documented emergency account.', src);

    if (!live.some(p => !isNone(p.authStrength))) add('high', 'Conditional Access', 'No policy requires an authentication strength', 'Admins should be required to use phishing-resistant MFA through an authentication strength.', src);
    if (live.length && !live.some(p => !isNone(p.userRisk))) add('high', 'Conditional Access', 'No user-risk policy', 'Add a policy that requires a secure password change when user risk is high. This needs Entra ID P2.', src);
    if (live.length && !live.some(p => !isNone(p.signInRisk))) add('high', 'Conditional Access', 'No sign-in risk policy', 'Add a policy that requires MFA for medium and high sign-in risk. This needs Entra ID P2.', src);
    if (live.length && !live.some(p => /devicecode/i.test(norm(p.authFlows)))) add('med', 'Conditional Access', 'Device code flow is not blocked', 'Device code phishing is common. Block the device code flow for everyone who does not need it.', src);

    return { policies, enabled: live.length };
  }

  function analyseApps(s, add) {
    const c = {
      name: col(s, ['applicationname', 'appname', 'displayname']), state: col(s, ['state', 'accountenabled']), owners: col(s, ['ownercount', 'owners']),
      secret: col(s, ['secretdaysleft']), cert: col(s, ['certdaysleft']), saml: col(s, ['samlcertdaysleft']), expired: col(s, ['expiredcredentials']),
      tenant: col(s, ['tenanttype', 'signinaudience']), status: col(s, ['status', 'compliance']), findings: col(s, ['findings', 'finding', 'issues']),
    };
    const minDays = v => { const s2 = String(v ?? ''); if (/expired/i.test(s2)) return -1; const ns = (s2.match(/-?\d+/g) || []).map(Number); return ns.length ? Math.min(...ns) : Infinity; };
    const apps = s.rows().map(r => {
      const owners = c.owners ? (isNaN(num(r[c.owners])) ? (isNone(r[c.owners]) || /no owners/i.test(r[c.owners]) ? 0 : 1) : num(r[c.owners])) : NaN;
      const days = Math.min(minDays(c.secret && r[c.secret]), minDays(c.cert && r[c.cert]), minDays(c.saml && r[c.saml]));
      const expired = (c.expired && num(r[c.expired]) > 0) || days < 0;
      const disabled = c.state ? /disabled|false/i.test(r[c.state]) : false;
      const multi = c.tenant ? /multi|personal|common/i.test(r[c.tenant]) : false;
      const issues = [];
      if (days >= 0 && days <= CONFIG.expiryWarnDays) issues.push(`Credential expires in ${plural(days, 'day')}`);
      if (expired) issues.push('Expired credential not removed');
      if (owners === 0) issues.push('No owner');
      if (multi) issues.push('Multi-tenant');
      if (c.findings && !isNone(r[c.findings])) String(r[c.findings]).split(/;\s*/).forEach(x => { if (x && !issues.some(i => norm(i).includes(norm(x).slice(0, 8)))) issues.push(x); });
      let sev = 'ok';
      if (days >= 0 && days <= CONFIG.expiryCriticalDays) sev = 'crit';
      else if (expired || (days >= 0 && days <= CONFIG.expiryWarnDays) || /non.?compliant/i.test(c.status ? r[c.status] : '')) sev = 'high';
      else if (owners === 0 || multi) sev = 'med';
      return { name: r[c.name], disabled, owners, days, expired, multi, issues, sev };
    }).sort((a, b) => sevRank[a.sev] - sevRank[b.sev]);

    const src = s.name;
    const expiring = apps.filter(a => a.days >= 0 && a.days <= CONFIG.expiryWarnDays);
    expiring.forEach(a => add(a.days <= CONFIG.expiryCriticalDays ? 'crit' : 'high', 'Apps', `${a.name} credential expires in ${plural(a.days, 'day')}`, 'Rotate the secret or certificate before it expires, or remove it if the app is no longer used.', src));
    const expired = apps.filter(a => a.expired);
    if (expired.length) add('high', 'Apps', `${plural(expired.length, 'app')} still holding expired credentials`, `${listNames(expired.map(a => a.name))}. Remove expired secrets and certificates.`, src);
    const ownerless = apps.filter(a => a.owners === 0);
    if (ownerless.length) add('med', 'Apps', `${ownerless.length} of ${apps.length} app registrations have no owner`, `${listNames(ownerless.map(a => a.name))}. Ownerless apps are hard to audit and easy to forget.`, src);
    const multi = apps.filter(a => a.multi && !a.disabled);
    if (multi.length) add('med', 'Apps', `${plural(multi.length, 'multi-tenant app')}`, `${listNames(multi.map(a => a.name))}. Confirm that sign-ins from other tenants are intended.`, src);
    return apps;
  }

  // ---------------------------------------------------------------- report HTML
  const SEV_LABEL = { crit: 'Critical', high: 'High', med: 'Medium', info: 'Info', ok: 'OK' };

  function sparkline(series) {
    const W = 320, H = 70, P = 6;
    const vals = series.map(p => p.pct);
    let lo = Math.min(...vals), hi = Math.max(...vals);
    if (hi - lo < 1) { lo -= 1; hi += 1; }
    const x = i => P + (series.length === 1 ? (W - 2 * P) / 2 : (i * (W - 2 * P)) / (series.length - 1));
    const y = v => H - P - ((v - lo) / (hi - lo)) * (H - 2 * P);
    const line = series.map((p, i) => `${x(i).toFixed(1)},${y(p.pct).toFixed(1)}`).join(' ');
    const last = series[series.length - 1];
    return `<svg class="spark" viewBox="0 0 ${W} ${H}" role="img" aria-label="Secure Score trend">
      <polygon points="${x(0).toFixed(1)},${H - P} ${line} ${x(series.length - 1).toFixed(1)},${H - P}" fill="var(--accent)" fill-opacity=".14"/>
      <polyline points="${line}" fill="none" stroke="var(--accent)" stroke-width="2.5" stroke-linejoin="round"/>
      <circle cx="${x(series.length - 1).toFixed(1)}" cy="${y(last.pct).toFixed(1)}" r="4.5" fill="var(--accent-2)"/></svg>`;
  }

  function render(m) {
    const f = m.findings;
    const counts = { crit: 0, high: 0, med: 0 };
    f.forEach(x => { if (x.sev in counts) counts[x.sev]++; });
    const score = m.score && m.score.latest;
    const ringDash = score ? ((score.pct / 100) * 515).toFixed(0) : 0;
    const checkItems = m.checks.flatMap(g => g.items);
    const checksFailing = checkItems.filter(i => i.state === 'fail').length;

    const stats = [];
    if (score) stats.push(['var(--accent)', `${fmt(score.pct)}%`, `Secure Score (${fmt(score.cur)} / ${fmt(score.max)})`]);
    stats.push(['var(--crit)', String(counts.crit), `critical ${counts.crit === 1 ? 'finding' : 'findings'}, plus ${counts.high} high`]);
    if (m.users) stats.push(['var(--high)', `${m.users.withMfa} / ${m.users.people.length}`, 'accounts registered for MFA']);
    if (m.risks) stats.push(['var(--accent-2)', String(m.risks.open), `of ${m.risks.total} Secure Score controls open`]);
    else if (checkItems.length) stats.push(['var(--accent-2)', `${checksFailing} / ${checkItems.length}`, 'tenant setting checks failing']);

    const sections = [];

    // Findings
    sections.push(`<section><div class="sec-head"><div class="eyebrow">Findings</div><h2>${f.filter(x => x.sev !== 'info').length ? 'What needs attention' : 'No major findings'}</h2>
      <p>Sorted by severity. Each finding names the sheet it came from.</p></div>
      <div class="findings">${f.map(x => `<div class="finding ${x.sev}"><span class="tag">${SEV_LABEL[x.sev]}</span>
        <div class="f-body"><h3>${esc(x.title)}</h3><p>${esc(x.detail)}</p><small>${esc(x.area)} · ${esc(x.source)}</small></div></div>`).join('')}</div></section>`);

    // Risk breakdown
    if (m.risks && m.risks.open) {
      const r = m.risks, segs = [
        ['var(--sky)', r.fam.endpoint, 'Defender for Endpoint', 'Device hardening checks'],
        ['var(--accent)', r.fam.identity, 'Defender for Identity', 'On-prem Active Directory'],
        ['var(--accent-2)', r.fam.cloudapps, 'Defender for Cloud Apps', 'Third-party SaaS connectors'],
        ['var(--crit)', r.fam.core, 'Core M365', 'Entra, Exchange, SharePoint, Teams, Purview'],
        ['var(--ok)', r.passing, 'Compliant', 'Controls already passing'],
      ].filter(s => s[1] > 0);
      sections.push(`<section><div class="sec-head"><div class="eyebrow">Secure Score controls</div><h2>Where the ${r.open} open controls come from</h2>
        <p>Endpoint, on-prem identity and SaaS connector checks only matter if you use those workloads. ${r.fam.core} open controls cover core Microsoft 365.</p></div>
        <div class="bar" role="img" aria-label="${segs.map(s => `${s[2]} ${s[1]}`).join(', ')}">${segs.map(s => `<div style="flex:${s[1]};background:${s[0]}">${s[1]}</div>`).join('')}</div>
        <div class="legend">${segs.map(s => `<div><i style="background:${s[0]}"></i><span>${s[2]} <small>${s[3]}</small></span></div>`).join('')}</div></section>`);
    }

    // Users + CA baseline / policies
    const panels = [];
    if (m.users) {
      panels.push(`<div class="panel"><h3>MFA registration by account</h3><div class="sub">${m.users.withMfa} of ${m.users.people.length} accounts · ${esc(m.S.users.name)}</div>
        <div class="people">${m.users.all.map(u => `<div class="person ${u.service ? 'med' : u.mfa ? 'ok' : 'crit'}" title="${esc(u.upn)}"><i></i><span>${esc(u.name)}${u.admin ? ' (admin)' : ''}${u.guest ? ' (guest)' : ''}</span></div>`).join('')}</div>
        <div class="key"><span class="ok"><i></i>MFA registered</span><span class="crit"><i></i>No MFA</span><span class="med"><i></i>Service account</span></div></div>`);
    }
    if (m.capBaseline && m.capBaseline.length) {
      const pass = m.capBaseline.filter(x => x.state === 'pass').length;
      panels.push(`<div class="panel"><h3>Conditional Access baseline</h3><div class="sub">${pass} of ${m.capBaseline.length} checks pass · ${esc(m.S.capBaseline.name)}</div>
        <ul class="checks">${m.capBaseline.slice().sort((a, b) => (a.state === 'pass' ? 0 : 1) - (b.state === 'pass' ? 0 : 1)).map(x => `<li class="${x.state === 'pass' ? 'ok' : 'crit'}"><b>${x.state === 'pass' ? 'PASS' : 'FAIL'}</b>${esc(x.title)}</li>`).join('')}</ul></div>`);
    } else if (m.cap) {
      panels.push(`<div class="panel"><h3>Conditional Access policies</h3><div class="sub">${m.cap.enabled} of ${m.cap.policies.length} enabled · ${esc(m.S.cap.name)}</div>
        <ul class="checks">${m.cap.policies.map(p => `<li class="${p.enabled ? 'ok' : p.reportOnly ? 'med' : 'crit'}"><b>${p.enabled ? 'ON' : p.reportOnly ? 'REPORT' : 'OFF'}</b>${esc(p.name)}</li>`).join('')}</ul></div>`);
    }
    if (panels.length) sections.push(`<section><div class="cols${panels.length === 1 ? ' one' : ''}">${panels.join('')}</div></section>`);

    // Apps
    if (m.apps && m.apps.length) {
      const ownerless = m.apps.filter(a => a.owners === 0).length;
      const pill = { crit: 'Act now', high: 'Fix', med: 'Review', ok: 'OK' };
      sections.push(`<section><div class="sec-head"><div class="eyebrow">App registrations</div><h2>${ownerless} of ${m.apps.length} apps have no owner</h2>
        <p>Source: ${esc(m.S.apps.name)}.</p></div>
        <div class="tbl"><table><thead><tr><th>Application</th><th>State</th><th class="num">Owners</th><th>Issues</th><th>Status</th></tr></thead><tbody>
        ${m.apps.map(a => `<tr><td>${esc(a.name)}</td><td>${a.disabled ? 'Disabled' : 'Enabled'}</td><td class="num">${isNaN(a.owners) ? '–' : a.owners}</td><td>${esc(a.issues.join('; ') || 'None')}</td><td><span class="pill ${a.sev}">${pill[a.sev]}</span></td></tr>`).join('')}
        </tbody></table></div></section>`);
    }

    // Core services
    if (m.risks && m.risks.services.length) {
      sections.push(`<section><div class="sec-head"><div class="eyebrow">Email, collaboration &amp; data</div><h2>Open controls by Microsoft 365 service</h2>
        <p>Top items by Secure Score points.</p></div>
        <div class="svc">${m.risks.services.map(s => `<div class="panel"><h3>${esc(s.name)}<span class="count">${s.items.length} open</span></h3>
          <ul>${s.items.slice(0, 5).map(i => `<li>${esc(i.title)}</li>`).join('')}</ul></div>`).join('')}</div></section>`);
    }

    // Quick wins
    const wins = (m.risks ? m.risks.top.map(x => ({ title: x.title, pts: x.max })) : [])
      .concat((m.actions || []).filter(a => a.gap > 0 && !(m.risks || { top: [] }).top.some(t => norm(t.title).slice(0, 18) === norm(a.title).slice(0, 18))).map(a => ({ title: a.title, pts: a.gap })))
      .sort((a, b) => b.pts - a.pts).slice(0, 10);
    if (wins.length) {
      const maxPts = Math.max(...wins.map(w => w.pts));
      const total = wins.reduce((n, w) => n + w.pts, 0);
      sections.push(`<section><div class="sec-head"><div class="eyebrow">Biggest Secure Score gains</div><h2>About ${fmt(total)} points in the top ${wins.length} items</h2>
        <p>Points available for each open control or recommendation. Some overlap, so treat the total as an estimate.</p></div>
        <div class="wins">${wins.map(w => `<div class="win"><span>${esc(w.title)}</span><div class="track"><div style="width:${((w.pts / maxPts) * 100).toFixed(0)}%"></div></div><b class="mono">${fmt(w.pts)}</b></div>`).join('')}</div></section>`);
    }

    // Tenant setting checks
    if (checkItems.length) {
      const fails = checkItems.filter(i => i.state === 'fail'), passes = checkItems.filter(i => i.state === 'pass');
      sections.push(`<section><div class="sec-head"><div class="eyebrow">Tenant setting checks</div><h2>${fails.length} of ${checkItems.length} settings need attention</h2>
        <p>Collected from ${plural(m.checks.length, 'sheet')} with a pass/fail column.</p></div>
        <div class="tbl"><table><thead><tr><th>Setting</th><th>Detail</th><th>Result</th></tr></thead><tbody>
        ${fails.concat(passes).map(i => `<tr><td>${esc(i.title)}<div class="src">${esc(i.source)}</div></td><td>${esc(i.detail)}</td><td><span class="pill ${i.state === 'pass' ? 'ok' : 'high'}">${i.state === 'pass' ? 'Pass' : 'Fail'}</span></td></tr>`).join('')}
        </tbody></table></div></section>`);
    }

    // Score trend + licenses
    const side = [];
    if (m.score && m.score.series.length > 1) {
      const s = m.score;
      side.push(`<div class="panel"><h3>Secure Score trend</h3><div class="sub">${s.days} data points · ${esc(m.S.score.name)}</div>${sparkline(s.series)}
        <div class="trend-meta"><span>${esc(fmtDate(s.oldest.date) || s.oldest.raw)}: <b>${fmt(s.oldest.pct)}%</b></span><span>${esc(fmtDate(s.latest.date) || s.latest.raw)}: <b>${fmt(s.latest.pct)}%</b></span></div></div>`);
    }
    if (m.licenses && m.licenses.length) {
      side.push(`<div class="panel"><h3>Licenses</h3><div class="sub">${esc(m.S.licenses.name)}</div>
        <div class="tbl flat"><table><thead><tr><th>Product</th><th class="num">Assigned</th><th class="num">Total</th></tr></thead><tbody>
        ${m.licenses.map(l => `<tr><td>${esc(l.name)}</td><td class="num">${l.used}</td><td class="num">${l.total}</td></tr>`).join('')}</tbody></table></div></div>`);
    }
    if (side.length) sections.push(`<section><div class="cols${side.length === 1 ? ' one' : ''}">${side.join('')}</div></section>`);

    // What's working
    const good = [];
    if (m.capBaseline) m.capBaseline.filter(x => x.state === 'pass').forEach(x => good.push(x.title));
    checkItems.filter(i => i.state === 'pass').forEach(i => good.push(i.title));
    (m.actions || []).filter(a => a.state === 'pass').forEach(a => good.push(a.title));
    if (good.length) sections.push(`<section><div class="sec-head"><div class="eyebrow">Already in place</div><h2>What's working</h2></div>
      <div class="good">${[...new Set(good)].slice(0, 24).map(g => `<span>${esc(g)}</span>`).join('')}</div></section>`);

    const missing = m.mapping.filter(x => !x.sheet).map(x => x.label);
    const matched = m.mapping.filter(x => x.sheet);

    return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(m.tenant)} Security Pulse</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,700;12..96,800&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>${REPORT_CSS}</style></head><body>
<header class="hero"><div class="wrap">
  <div><div class="eyebrow">Microsoft 365 posture review</div><h1>${esc(m.tenant)} Security Pulse</h1>
  <p class="lede">Automatic findings from the M365 Pulse Report workbook. ${counts.crit ? `${plural(counts.crit, 'critical finding')} and ${counts.high} high need attention first.` : counts.high ? `${plural(counts.high, 'high finding')} need attention first.` : 'No critical or high findings.'}</p>
  <div class="meta"><span>Report date ${esc(m.reportDate)}</span><span>${plural(m.sheetCount, 'sheet')} read</span>${m.domain ? `<span>${esc(m.domain)}</span>` : ''}</div></div>
  ${score ? `<svg class="ring" viewBox="0 0 200 200" role="img" aria-label="Secure Score ${fmt(score.pct)} percent">
    <circle cx="100" cy="100" r="82" fill="none" stroke="rgba(255,255,255,.2)" stroke-width="18"/>
    <circle cx="100" cy="100" r="82" fill="none" stroke="#ffd166" stroke-width="18" stroke-linecap="round" stroke-dasharray="${ringDash} 515" transform="rotate(-90 100 100)"/>
    <text x="100" y="98" text-anchor="middle" fill="#fff" font-size="40" font-weight="800">${fmt(score.pct)}%</text>
    <text x="100" y="124" text-anchor="middle" fill="rgba(255,255,255,.8)" font-size="13" font-family="IBM Plex Mono, monospace">${fmt(score.cur)} / ${fmt(score.max)} pts</text></svg>` : ''}
</div></header>
<main class="wrap">
  <div class="stats">${stats.map(s => `<div class="stat" style="--c:${s[0]}"><b>${esc(s[1])}</b><span>${esc(s[2])}</span></div>`).join('')}</div>
  ${sections.join('\n')}
</main>
<footer class="wrap">
  <p>Generated ${esc(fmtDate(new Date()))} from <b>${esc(m.fileName || 'workbook')}</b> by Pulse Report Builder. No data left this computer.</p>
  <p>Sheets used: ${matched.map(x => `${esc(x.label)} (${esc(x.sheet)})`).join(', ')}${m.checks.length ? `, plus ${plural(m.checks.length, 'settings sheet')}` : ''}.</p>
  ${missing.length ? `<p>Not found in this workbook: ${esc(missing.join(', '))}.</p>` : ''}
</footer></body></html>`;
  }

  const REPORT_CSS = `
:root{--bg:#f5f4fb;--surface:#fff;--ink:#1d1b3a;--muted:#5f5c7d;--line:#e2dff0;--accent:#5b3fd6;--accent-2:#e2448a;--sky:#1f9bd1;
--crit:#e0413a;--crit-bg:#fde8e6;--high:#ef7d1a;--high-bg:#fdf0e2;--med:#b88a00;--med-bg:#fbf4d9;--ok:#159a6b;--ok-bg:#e1f5ec;--info:#1f7fb8;--info-bg:#e3f1fa;
--hero-1:#3a1fa8;--hero-2:#8b2fc9;--hero-3:#e2448a;
--display:"Bricolage Grotesque",ui-sans-serif,system-ui,sans-serif;--body:"IBM Plex Sans",ui-sans-serif,system-ui,"Segoe UI",sans-serif;--mono:"IBM Plex Mono",ui-monospace,Consolas,monospace}
@media (prefers-color-scheme:dark){:root{--bg:#13122a;--surface:#1d1b3a;--ink:#ecebf8;--muted:#a6a3c6;--line:#34315c;--accent:#9d87ff;--accent-2:#ff6fab;--sky:#4fc0f0;
--crit:#ff6b63;--crit-bg:#3d1c25;--high:#ffa04d;--high-bg:#3a2619;--med:#f0c53a;--med-bg:#352e14;--ok:#3fd09a;--ok-bg:#15332b;--info:#6cc4f5;--info-bg:#16304a;
--hero-1:#2a1680;--hero-2:#6a22a3;--hero-3:#b8336f;color-scheme:dark}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:var(--body);font-size:15px;line-height:1.55}
.wrap{max-width:1080px;margin:0 auto;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
h1,h2,h3{font-family:var(--display);text-wrap:balance;margin:0;line-height:1.15}
h2{font-size:1.6rem;font-weight:800}h3{font-size:1.05rem;font-weight:700}p{margin:0}
.mono{font-family:var(--mono);font-variant-numeric:tabular-nums}
.eyebrow{font-family:var(--mono);font-size:.75rem;letter-spacing:.12em;text-transform:uppercase;color:var(--muted)}
.hero{background:linear-gradient(120deg,var(--hero-1),var(--hero-2) 55%,var(--hero-3));color:#fff}
.hero .wrap{padding-block:40px 48px;display:grid;grid-template-columns:1fr auto;gap:32px;align-items:center}
.hero .eyebrow{color:rgba(255,255,255,.75)}
.hero h1{font-size:clamp(2rem,5vw,3.2rem);font-weight:800;margin:8px 0 12px}
.lede{max-width:58ch;color:rgba(255,255,255,.88)}
.meta{display:flex;flex-wrap:wrap;gap:8px;margin-top:18px}
.meta span{font-family:var(--mono);font-size:.78rem;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.25);padding:4px 10px;border-radius:999px}
.ring{width:190px;height:190px}.ring text{font-family:var(--display)}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-top:-26px;position:relative}
.stat{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:16px 18px;border-top:5px solid var(--c)}
.stat b{display:block;font-family:var(--display);font-size:2rem;font-weight:800;color:var(--c);line-height:1.1}
.stat span{font-size:.85rem;color:var(--muted)}
section{padding-block:40px 8px}
.sec-head{display:flex;flex-direction:column;gap:6px;margin-bottom:18px}.sec-head p{color:var(--muted);max-width:68ch}
.bar{display:flex;height:34px;border-radius:10px;overflow:hidden;border:1px solid var(--line)}
.bar div{display:flex;align-items:center;justify-content:center;color:#fff;font-family:var(--mono);font-size:.78rem;min-width:0;overflow:hidden;white-space:nowrap}
.legend{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:10px;margin-top:14px}
.legend div{display:flex;gap:10px;align-items:flex-start;font-size:.88rem}
.legend i{width:14px;height:14px;border-radius:4px;flex:none;margin-top:3px}.legend small{display:block;color:var(--muted)}
.findings{display:grid;gap:12px}
.finding{background:var(--surface);border:1px solid var(--line);border-left:6px solid var(--c);border-radius:12px;padding:16px 18px;display:flex;gap:14px;align-items:flex-start}
.finding .tag{font-family:var(--mono);font-size:.7rem;font-weight:500;letter-spacing:.08em;text-transform:uppercase;color:var(--c);background:var(--cb);padding:4px 8px;border-radius:6px;white-space:nowrap;min-width:72px;text-align:center}
.f-body{display:grid;gap:4px;min-width:0}.f-body p{color:var(--muted);font-size:.92rem;overflow-wrap:anywhere}.f-body small{font-family:var(--mono);font-size:.72rem;color:var(--muted)}
.crit{--c:var(--crit);--cb:var(--crit-bg)}.high{--c:var(--high);--cb:var(--high-bg)}.med{--c:var(--med);--cb:var(--med-bg)}.ok{--c:var(--ok);--cb:var(--ok-bg)}.info{--c:var(--info);--cb:var(--info-bg)}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:16px}.cols.one{grid-template-columns:1fr}
.panel{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:20px;min-width:0}
.panel h3{margin-bottom:4px}.panel .sub{color:var(--muted);font-size:.85rem;margin-bottom:14px}
.people{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:8px}
.person{display:flex;align-items:center;gap:8px;font-size:.82rem;padding:7px 9px;border-radius:8px;background:var(--cb);min-width:0}
.person i,.key i{width:10px;height:10px;border-radius:50%;background:var(--c);flex:none;display:inline-block}
.person span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.key{display:flex;flex-wrap:wrap;gap:14px;margin-top:12px;font-size:.8rem;color:var(--muted)}.key span{display:flex;align-items:center;gap:6px}
.checks{display:grid;gap:6px;list-style:none;margin:0;padding:0}
.checks li{display:flex;align-items:center;gap:10px;font-size:.88rem;padding:6px 10px;border-radius:8px;background:var(--cb)}
.checks li b{font-family:var(--mono);font-size:.7rem;color:var(--c);min-width:58px;letter-spacing:.06em}
.tbl{overflow-x:auto;border:1px solid var(--line);border-radius:14px;background:var(--surface)}
.tbl.flat{border:0;border-radius:0}.tbl.flat table{min-width:0}
table{border-collapse:collapse;width:100%;min-width:640px;font-size:.88rem}
th,td{text-align:left;padding:10px 14px;border-bottom:1px solid var(--line);vertical-align:top}
th{font-family:var(--mono);font-size:.72rem;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);font-weight:500;background:var(--bg)}
tr:last-child td{border-bottom:0}.num{text-align:right;font-family:var(--mono);font-variant-numeric:tabular-nums}
td .src{font-family:var(--mono);font-size:.7rem;color:var(--muted);margin-top:2px}
.pill{display:inline-block;font-family:var(--mono);font-size:.72rem;padding:2px 8px;border-radius:999px;color:var(--c);background:var(--cb);white-space:nowrap}
.svc{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px}
.svc ul{margin:0;padding-left:18px;color:var(--muted);font-size:.88rem;display:grid;gap:4px}
.svc .count{font-family:var(--mono);font-size:.78rem;color:var(--crit);margin-left:8px;font-weight:500}
.wins{display:grid;gap:10px}
.win{display:grid;grid-template-columns:1.2fr 2fr auto;gap:14px;align-items:center;background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:12px 16px;font-size:.9rem}
.win .track{height:12px;border-radius:999px;background:var(--bg);overflow:hidden;border:1px solid var(--line)}
.win .track div{height:100%;border-radius:999px;background:linear-gradient(90deg,var(--accent),var(--accent-2))}
.win b{font-family:var(--display);font-size:1.2rem;color:var(--accent);min-width:2.5ch;text-align:right}
.spark{width:100%;height:auto;max-width:100%;display:block}
.trend-meta{display:flex;justify-content:space-between;font-size:.82rem;color:var(--muted);margin-top:8px;font-family:var(--mono)}.trend-meta b{color:var(--ink)}
.good{display:flex;flex-wrap:wrap;gap:8px}.good span{font-size:.85rem;padding:6px 12px;border-radius:999px;background:var(--ok-bg);color:var(--ok);font-weight:500}
footer{padding-block:40px 48px;color:var(--muted);font-size:.82rem;display:grid;gap:6px}
@media (max-width:760px){.hero .wrap{grid-template-columns:1fr}.ring{width:160px;height:160px}.cols{grid-template-columns:1fr}
.win{grid-template-columns:1fr auto}.win .track{grid-column:1/-1;grid-row:2}.finding{flex-direction:column;gap:8px}}`;

  // ---------------------------------------------------------------- public API
  async function buildReport(input, fileName) {
    const wb = await readWorkbook(input);
    const model = analyse(wb, fileName);
    model.sheetCount = wb.sheets.length;
    return { html: render(model), model };
  }

  const api = { buildReport, readWorkbook, analyse, render, CONFIG, SIGNATURES };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PulseEngine = api;
})(typeof self !== 'undefined' ? self : this);
