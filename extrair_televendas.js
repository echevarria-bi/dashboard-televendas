var XLSX = require('C:\\Users\\mar\\AppData\\Local\\Temp\\opencode\\node_modules\\xlsx');
var fs = require('fs');

var BASE_DIR = 'C:\\Users\\mar\\OneDrive - SPADER DISTRIBUIDORA DE ALIMENTOS L\\Área de Trabalho\\';
var BASE = BASE_DIR + '_bases\\base_8026_2026.xlsx';
var OUT_DIR = BASE_DIR + 'dashboards\\televendas\\';

// Vendedores ativos (CODUSUR)
var ACTIVE = ['1596', '1464', '1211', '1429', '9886', '1624'];
// Todos os RCAs que já apareceram (para dados_mensal)
var ALL_RCAS = ['1596', '1464', '1211', '1429', '9886', '1624', '1527', '1571', '1573'];

function serialToDate(s) {
  if (!s || s < 60) return null;
  return new Date(Math.round((s - 25569) * 86400000));
}

console.log('Lendo base_8026_2026.xlsx...');
var wb = XLSX.readFile(BASE);
var ws = wb.Sheets['Plan1'];
var raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
console.log('Total linhas: ' + raw.length);

// ============================================================
// 1. Dados mensais (todos os meses) → dados_mensal.js
// ============================================================
console.log('\n[1/2] Extraindo dados mensais...');

var meses = {};
for (var m = 1; m <= 12; m++) {
  meses[m] = {};
  ALL_RCAS.forEach(function (c) { meses[m][c] = { fat: 0, clis: {} }; });
}

var globalClis = {};
var activeClis = {};
var now = new Date();
var currentMonth = now.getUTCMonth() + 1;
var currentYear = now.getUTCFullYear();

for (var ri = 1; ri < raw.length; ri++) {
  var r = raw[ri];
  if (!r || !r[27]) continue;
  var c = String(r[27]).trim();
  if (ALL_RCAS.indexOf(c) < 0) continue;
  var dt = serialToDate(parseFloat(r[2]));
  if (!dt || dt.getUTCFullYear() !== currentYear) continue;
  var mes = dt.getUTCMonth() + 1;
  if (mes < 1 || mes > currentMonth) continue;
  var fat = parseFloat(r[34]) || (parseFloat(r[9]) || 0) + (parseFloat(r[10]) || 0);
  var rec = meses[mes][c];
  rec.fat += fat;
  if (fat >= 1) {
    rec.clis[String(r[11])] = 1;
  }
  globalClis[String(r[11])] = 1;
  if (ACTIVE.indexOf(c) >= 0) activeClis[String(r[11])] = 1;
}

// Build DADOS_MES object
var dadosMes = {};
for (var m = 1; m <= currentMonth; m++) {
  dadosMes[m] = {};
  ALL_RCAS.forEach(function (c) {
    var rec = meses[m][c];
    dadosMes[m][c] = { fat: Math.round(rec.fat * 100) / 100, cli: Object.keys(rec.clis).length };
  });
}
dadosMes.totalCli = Object.keys(globalClis).length;
dadosMes.activeCli = Object.keys(activeClis).length;

var jsOut = 'var DADOS_MES=' + JSON.stringify(dadosMes) + ';';
fs.writeFileSync(OUT_DIR + 'dados_mensal.js', jsOut, 'utf8');
console.log('dados_mensal.js salvo (' + jsOut.length + ' bytes)');

// Print summary
for (var m = 1; m <= currentMonth; m++) {
  var nome = ['', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'][m];
  var totalFat = 0;
  ALL_RCAS.forEach(function (c) { totalFat += dadosMes[m][c].fat; });
  console.log('  ' + nome + ': R$ ' + totalFat.toFixed(2));
}

// ============================================================
// 2. Dados do mês atual → dados_julho.js (ou dados_mes.js)
// ============================================================
console.log('\n[2/2] Extraindo dados do mês atual...');

var rows = [];
for (var ri = 1; ri < raw.length; ri++) {
  var r = raw[ri];
  if (!r || !r[0]) continue;
  var codusur = String(r[27]).trim();
  if (ACTIVE.indexOf(codusur) < 0) continue;

  var serial = parseFloat(r[2]);
  var dt = serialToDate(serial);
  if (!dt) continue;
  if (dt.getUTCFullYear() !== currentYear || dt.getUTCMonth() !== (currentMonth - 1)) continue;

  rows.push({
    data: r[2],
    dataStr: dt.toLocaleDateString('pt-BR', { timeZone: 'UTC' }),
    codcli: r[11],
    cliente: r[12],
    ramo: r[16],
    depto: r[18],
    codusur: codusur,
    rca: r[28],
    supervisor: r[30],
    codfornec: String(r[24]).trim(),
    fat: parseFloat(r[34]) || (parseFloat(r[9]) || 0) + (parseFloat(r[10]) || 0)
  });
}

var jsMonth = 'var DADOS = ' + JSON.stringify(rows.map(function (r) {
  return { data: r.data, codcli: r.codcli, cliente: r.cliente, codusur: r.codusur, rca: r.rca, codfornec: r.codfornec, depto: r.depto, fat: r.fat };
})) + ';';
fs.writeFileSync(OUT_DIR + 'dados_julho.js', jsMonth, 'utf8');
console.log('dados_julho.js salvo (' + rows.length + ' pedidos, ' + jsMonth.length + ' bytes)');

// Also save JSON for reference
var jsonOut = {
  mes: nome + '/' + currentYear,
  filtro: { codusur: ACTIVE },
  rows: rows
};
fs.writeFileSync(OUT_DIR + 'dados_julho.json', JSON.stringify(jsonOut, null, 2));
console.log('dados_julho.json salvo');

console.log('\n=== Resumo ===');
console.log('Total pedidos mês atual: ' + rows.length);
var totalFat = rows.reduce(function (s, r) { return s + r.fat; }, 0);
console.log('Faturamento total: R$ ' + totalFat.toFixed(2));
console.log('Clientes únicos: ' + Object.keys(rows.reduce(function (s, r) { s[r.codcli] = 1; return s; }, {})).length);
