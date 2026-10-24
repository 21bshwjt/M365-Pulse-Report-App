#!/usr/bin/env node
// Usage: node cli.js <workbook.xlsx> [more.xlsx ...] [-o output.html | -d output-folder]
const fs = require('fs');
const path = require('path');
const { buildReport } = require('./pulse-engine.js');

async function main() {
  const args = process.argv.slice(2);
  const inputs = [], opts = {};
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '-o' || args[i] === '--out') opts.out = args[++i];
    else if (args[i] === '-d' || args[i] === '--dir') opts.dir = args[++i];
    else inputs.push(args[i]);
  }
  if (!inputs.length) {
    console.log('Usage: node cli.js <workbook.xlsx> [more.xlsx ...] [-o output.html | -d output-folder]');
    process.exit(1);
  }
  for (const input of inputs) {
    const name = path.basename(input);
    const { html, model } = await buildReport(fs.readFileSync(input), name);
    const out = opts.out && inputs.length === 1 ? opts.out
      : path.join(opts.dir || path.dirname(input), name.replace(/\.xlsx?$/i, '') + '.html');
    fs.writeFileSync(out, html);
    const sev = s => model.findings.filter(f => f.sev === s).length;
    console.log(`\n${name} -> ${out}`);
    console.log(`  Tenant: ${model.tenant}   Findings: ${sev('crit')} critical, ${sev('high')} high, ${sev('med')} medium`);
    for (const m of model.mapping) console.log(`  ${m.sheet ? 'ok  ' : '--  '}${m.label.padEnd(30)} ${m.sheet || 'not found'}`);
    if (model.checks.length) console.log(`  ok  Tenant setting checks        ${model.checks.map(c => c.sheet.name).join(', ')}`);
  }
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
