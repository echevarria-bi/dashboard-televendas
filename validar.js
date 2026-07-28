var XLSX=require('xlsx');
var BASEDIR='C:\\Users\\mar\\OneDrive - SPADER DISTRIBUIDORA DE ALIMENTOS L\\Área de Trabalho';
var wb=XLSX.readFile(BASEDIR+'\\_bases\\base_8026_2026.xlsx');
var ws=wb.Sheets['Plan1'];
var raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
console.log('Linhas totais base: '+raw.length);

function serialToDate(s){if(!s||s<60)return null;return new Date(Math.round((s-25569)*86400000));}

var ALL_RCAS=['1596','1464','1211','1429','9886','1624','1571','1527','1573'];
var METAS={
  '1596':{nome:'Ariane',regiao:'POA',fat:145000,pos:95},
  '1464':{nome:'Camila',regiao:'SUL',fat:127328,pos:116},
  '1211':{nome:'Cristielen',regiao:'VALE DOS SINOS',fat:117386,pos:116},
  '1429':{nome:'Natália',regiao:'KEY ACCOUNT',fat:350000,pos:10},
  '9886':{nome:'Tatiana',regiao:'SERRA',fat:145000,pos:100},
  '1624':{nome:'Tatiana',regiao:'LITORAL',fat:95000,pos:90}
};
var METAS_POLPA={
  '1596':{nome:'Ariane',regiao:'POA',fat:36500,pos:62},
  '1464':{nome:'Camila',regiao:'SUL',fat:40000,pos:98},
  '1211':{nome:'Cristielen',regiao:'VALE DOS SINOS',fat:75000,pos:103},
  '1429':{nome:'Natália',regiao:'KEY ACCOUNT',fat:175000,pos:0},
  '9886':{nome:'Tatiana',regiao:'SERRA',fat:87500,pos:143},
  '1624':{nome:'Tatiana',regiao:'LITORAL',fat:0,pos:0}
};

var fmt=function(v){return 'R$ '+v.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g,'.');};
var errs=0;
function check(label,ext,esperado,eps){
  eps=eps||0.1;
  if(esperado===null||esperado===undefined)return true;
  var diff=Math.abs(ext-esperado);
  if(diff>eps){
    console.log('  ❌ '+label+': extraído='+ext.toFixed(2)+' esperado='+esperado.toFixed(2)+' diff='+diff.toFixed(2));
    errs++;
    return false;
  }
  return true;
}

console.log('\n==============================================');
console.log('VALIDAÇÃO CRUZADA: BASE EXCEL vs EXTRAÇÃO');
console.log('==============================================');

// =============================================
// 1. Extrair dados Julho direto da base
// =============================================
var julRows=[];
var allRows=[];
for(var ri=1;ri<raw.length;ri++){
  var r=raw[ri];if(!r||!r[27])continue;
  var c=String(r[27]).trim();
  if(ALL_RCAS.indexOf(c)<0)continue;
  var dt=serialToDate(parseFloat(r[2]));
  if(!dt||dt.getUTCFullYear()!==2026)continue;
  var mes=dt.getUTCMonth()+1;
  var fat=parseFloat(r[34])||(parseFloat(r[9])||0)+(parseFloat(r[10])||0);
  if(isNaN(fat))fat=0;
  var codfornec=String(r[24]||'').trim();
  var row={codusur:c,rca:r[28],codcli:String(r[11]),cliente:r[12],fat:fat,mes:mes,dt:dt,dthr:parseFloat(r[2]),codfornec:codfornec};
  allRows.push(row);
  if(mes===7)julRows.push(row);
}

console.log('\n--- 1. CONTAGEM DE REGISTROS ---');
check('Registros Julho (base)',julRows.length,2371,0);
check('Registros Jan-Jul (base)',allRows.length,null,0);
var polpaJul=julRows.filter(function(r){return r.codfornec==='10828';});
check('Registros Polpanorte Julho',polpaJul.length,null,0);

// =============================================
// 2. Validar dados_mensal.js (todos os meses)
// =============================================
console.log('\n--- 2. VALIDAÇÃO dados_mensal.js vs BASE EXCEL ---');
var dadosMes={};
for(var m=1;m<=7;m++){
  dadosMes[m]={};
  ALL_RCAS.forEach(function(c){dadosMes[m][c]={fat:0,cli:{}};});
}
allRows.forEach(function(r){
  var rec=dadosMes[r.mes][r.codusur];
  if(!rec)return;
  rec.fat+=r.fat;
  if(r.fat>=1)rec.cli[r.codcli]=1;
});

// dados_mensal.js (hardcoded from file)
var dm={"1":{"1211":{"fat":0,"cli":0},"1429":{"fat":384085.69,"cli":153},"1464":{"fat":0,"cli":0},"1527":{"fat":212857.5,"cli":89},"1571":{"fat":390788.74,"cli":182},"1573":{"fat":187546.22,"cli":106},"1596":{"fat":0,"cli":0},"1624":{"fat":0,"cli":0},"9886":{"fat":161255.57,"cli":90}},"2":{"1211":{"fat":0,"cli":0},"1429":{"fat":285810.86,"cli":147},"1464":{"fat":0,"cli":0},"1527":{"fat":146954.52,"cli":84},"1571":{"fat":368639.95,"cli":192},"1573":{"fat":123268.94,"cli":91},"1596":{"fat":4705.97,"cli":3},"1624":{"fat":0,"cli":0},"9886":{"fat":176221.05,"cli":101}},"3":{"1211":{"fat":0,"cli":0},"1429":{"fat":371010.73,"cli":136},"1464":{"fat":0,"cli":0},"1527":{"fat":0,"cli":0},"1571":{"fat":304654.76,"cli":132},"1573":{"fat":131818.49,"cli":66},"1596":{"fat":181248.28,"cli":66},"1624":{"fat":0,"cli":0},"9886":{"fat":133398.79,"cli":76}},"4":{"1211":{"fat":0,"cli":0},"1429":{"fat":209669.06,"cli":124},"1464":{"fat":100007.04,"cli":63},"1527":{"fat":0,"cli":0},"1571":{"fat":234307.32,"cli":121},"1573":{"fat":400.36,"cli":1},"1596":{"fat":121767.7,"cli":62},"1624":{"fat":0,"cli":0},"9886":{"fat":158770.71,"cli":100}},"5":{"1211":{"fat":62863.94,"cli":52},"1429":{"fat":170101.69,"cli":102},"1464":{"fat":161974.55,"cli":120},"1527":{"fat":0,"cli":0},"1571":{"fat":-150.48,"cli":0},"1573":{"fat":0,"cli":0},"1596":{"fat":153592.91,"cli":69},"1624":{"fat":0,"cli":0},"9886":{"fat":101331.85,"cli":83}},"6":{"1211":{"fat":82172.88,"cli":65},"1429":{"fat":275800.86,"cli":109},"1464":{"fat":175506.76,"cli":115},"1527":{"fat":0,"cli":0},"1571":{"fat":0,"cli":0},"1573":{"fat":0,"cli":0},"1596":{"fat":120568.18,"cli":69},"1624":{"fat":413.6,"cli":1},"9886":{"fat":123576.67,"cli":92}},"7":{"1211":{"fat":132941.69,"cli":88},"1429":{"fat":291271.14,"cli":11},"1464":{"fat":135481.98,"cli":98},"1527":{"fat":0,"cli":0},"1571":{"fat":0,"cli":0},"1573":{"fat":0,"cli":0},"1596":{"fat":152133.6,"cli":53},"1624":{"fat":70745.99,"cli":50},"9886":{"fat":123935.37,"cli":83}}};

var mNames=['Jan','Fev','Mar','Abr','Mai','Jun','Jul'];
var diffCount=0;
for(var m=1;m<=7;m++){
  ALL_RCAS.forEach(function(c){
    var extF=Math.round(dadosMes[m][c].fat*100)/100;
    var extC=Object.keys(dadosMes[m][c].cli).length;
    var jsF=dm[m][c]?dm[m][c].fat:0;
    var jsC=dm[m][c]?dm[m][c].cli:0;
    var dF=Math.abs(extF-jsF);
    var dC=Math.abs(extC-jsC);
    if(dF>0.1||dC>0){
      console.log('  ❌ Mês '+mNames[m-1]+' Codusur '+c+': FAT extraído='+extF.toFixed(2)+' js='+jsF.toFixed(2)+' diff='+dF.toFixed(2)+' | CLI extraído='+extC+' js='+jsC+' diff='+dC);
      diffCount++;
    }
  });
}
if(diffCount===0)console.log('  ✅ dados_mensal.js 100% compatível com base Excel (todos os meses, todos os CODUSURs)');
else console.log('  ⚠️ '+diffCount+' diferenças encontradas');

// =============================================
// 3. Validar dados_polpa_mensal.js
// =============================================
console.log('\n--- 3. VALIDAÇÃO dados_polpa_mensal.js vs BASE EXCEL ---');
var polpaMeses={};
for(var m=1;m<=12;m++){
  polpaMeses[m]={};
  ALL_RCAS.forEach(function(c){polpaMeses[m][c]={fat:0,cli:{}};});
}
allRows.forEach(function(r){
  if(r.codfornec!=='10828')return;
  var rec=polpaMeses[r.mes][r.codusur];
  if(!rec)return;
  rec.fat+=r.fat;
  if(r.fat>=1)rec.cli[r.codcli]=1;
});

var pm={"1":{"1211":{"fat":0,"cli":0},"1429":{"fat":280666.17,"cli":117},"1464":{"fat":0,"cli":0},"1527":{"fat":73430.2,"cli":63},"1571":{"fat":188168.73,"cli":136},"1573":{"fat":147867.27,"cli":93},"1596":{"fat":0,"cli":0},"1624":{"fat":0,"cli":0},"9886":{"fat":82377.13,"cli":57}},"2":{"1211":{"fat":0,"cli":0},"1429":{"fat":240940.43,"cli":113},"1464":{"fat":0,"cli":0},"1527":{"fat":73436.87,"cli":61},"1571":{"fat":176935.03,"cli":149},"1573":{"fat":94791.22,"cli":75},"1596":{"fat":1690,"cli":1},"1624":{"fat":0,"cli":0},"9886":{"fat":115819.92,"cli":68}},"3":{"1211":{"fat":0,"cli":0},"1429":{"fat":285924.27,"cli":97},"1464":{"fat":0,"cli":0},"1527":{"fat":0,"cli":0},"1571":{"fat":98540.33,"cli":91},"1573":{"fat":95857.56,"cli":49},"1596":{"fat":72397.46,"cli":40},"1624":{"fat":0,"cli":0},"9886":{"fat":77347.71,"cli":48}},"4":{"1211":{"fat":0,"cli":0},"1429":{"fat":90779.71,"cli":83},"1464":{"fat":69959.92,"cli":48},"1527":{"fat":0,"cli":0},"1571":{"fat":65999.22,"cli":75},"1573":{"fat":179.8,"cli":1},"1596":{"fat":37917.91,"cli":39},"1624":{"fat":0,"cli":0},"9886":{"fat":85915.56,"cli":63}},"5":{"1211":{"fat":34257.5,"cli":38},"1429":{"fat":109837.77,"cli":57},"1464":{"fat":29591.78,"cli":64},"1527":{"fat":0,"cli":0},"1571":{"fat":0,"cli":0},"1573":{"fat":0,"cli":0},"1596":{"fat":22263.94,"cli":32},"1624":{"fat":0,"cli":0},"9886":{"fat":35173.71,"cli":43}},"6":{"1211":{"fat":46146.88,"cli":56},"1429":{"fat":215984.09,"cli":92},"1464":{"fat":43384.68,"cli":88},"1527":{"fat":0,"cli":0},"1571":{"fat":0,"cli":0},"1573":{"fat":0,"cli":0},"1596":{"fat":43802.29,"cli":52},"1624":{"fat":0,"cli":0},"9886":{"fat":46394.94,"cli":73}},"7":{"1211":{"fat":74264.2,"cli":64},"1429":{"fat":99723.04,"cli":1},"1464":{"fat":45265.51,"cli":71},"1527":{"fat":0,"cli":0},"1571":{"fat":0,"cli":0},"1573":{"fat":0,"cli":0},"1596":{"fat":41679.17,"cli":25},"1624":{"fat":45620.77,"cli":37},"9886":{"fat":52398.51,"cli":49}}};

var polpaDiff=0;
for(var m=1;m<=7;m++){
  ALL_RCAS.forEach(function(c){
    var extF=Math.round(polpaMeses[m][c].fat*100)/100;
    var extC=Object.keys(polpaMeses[m][c].cli).length;
    var jsF=pm[m][c]?pm[m][c].fat:0;
    var jsC=pm[m][c]?pm[m][c].cli:0;
    var dF=Math.abs(extF-jsF);
    var dC=Math.abs(extC-jsC);
    if(dF>0.1||dC>0){
      console.log('  ❌ Mês '+mNames[m-1]+' Codusur '+c+': FAT extraído='+extF.toFixed(2)+' js='+jsF.toFixed(2)+' diff='+dF.toFixed(2)+' | CLI extraído='+extC+' js='+jsC+' diff='+dC);
      polpaDiff++;
    }
  });
}
if(polpaDiff===0)console.log('  ✅ dados_polpa_mensal.js 100% compatível com base Excel');
else console.log('  ⚠️ '+polpaDiff+' diferenças encontradas');

// =============================================
// 4. Validar dados_julho.json
// =============================================
console.log('\n--- 4. VALIDAÇÃO dados_julho.json vs BASE EXCEL ---');
var dj=require('./dados_julho.json');
var djRows=dj.rows;

// Compare count
check('dados_julho.json rows count',djRows.length,julRows.length,0);

// Compare fat totals
var djFat=djRows.reduce(function(s,r){return s+r.fat;},0);
var extFat=julRows.reduce(function(s,r){return s+r.fat;},0);
check('dados_julho.json fat total',djFat,extFat,0.02);

// Compare per codusur
var djByCod={};djRows.forEach(function(r){if(!djByCod[r.codusur])djByCod[r.codusur]={fat:0,cli:{}};djByCod[r.codusur].fat+=r.fat;if(r.fat>=1)djByCod[r.codusur].cli[r.codcli]=1;});
var extByCod={};julRows.forEach(function(r){if(!extByCod[r.codusur])extByCod[r.codusur]={fat:0,cli:{}};extByCod[r.codusur].fat+=r.fat;if(r.fat>=1)extByCod[r.codusur].cli[r.codcli]=1;});

ALL_RCAS.forEach(function(c){
  var djF=djByCod[c]?djByCod[c].fat:0;
  var exF=extByCod[c]?extByCod[c].fat:0;
  var djC=djByCod[c]?Object.keys(djByCod[c].cli).length:0;
  var exC=extByCod[c]?Object.keys(extByCod[c].cli).length:0;
  check('Julho '+c+' FAT',djF,exF,0.02);
  check('Julho '+c+' CLI',djC,exC,0);
});

// Compare codfornec
var djPolpa=djRows.filter(function(r){return String(r.codfornec||'').trim()==='10828';});
var extPolpa=julRows.filter(function(r){return r.codfornec==='10828';});
check('Polpanorte registros Julho',djPolpa.length,extPolpa.length,0);
var djPF=djPolpa.reduce(function(s,r){return s+r.fat;},0);
var extPF=extPolpa.reduce(function(s,r){return s+r.fat;},0);
check('Polpanorte fat Julho',djPF,extPF,0.02);

// =============================================
// 5. Dashboard GERAL - Big Numbers
// =============================================
console.log('\n--- 5. DASHBOARD GERAL - BIG NUMBERS (base Excel) ---');

var fatBru=julRows.reduce(function(s,r){return s+r.fat;},0);
var posJul=julRows.filter(function(r){return r.fat>=1;});
var totPos=posJul.length;
var cliPos={};
posJul.forEach(function(r){cliPos[r.codcli]=(cliPos[r.codcli]||0)+r.fat;});
var qtdCliPos=Object.keys(cliPos).length;

var metaFatG=0,metaPosG=0;
Object.keys(METAS).forEach(function(k){metaFatG+=METAS[k].fat||0;metaPosG+=METAS[k].pos||0;});
var pfPctG=metaFatG>0?fatBru/metaFatG*100:0;
var ppPctG=metaPosG>0?qtdCliPos/metaPosG*100:0;
var dc=21,du=23,pmul=du/dc;
var fatProjG=Math.round(fatBru*pmul);

console.log('Faturamento bruto Julho:   '+fmt(fatBru));
console.log('Meta FAT:                  '+fmt(metaFatG));
console.log('% FAT:                     '+pfPctG.toFixed(1)+'%');
console.log('Projeção FAT:              '+fmt(fatProjG)+' ('+fatProjG/metaFatG*100+'%)');
console.log('');
console.log('Positivações (totPos):     '+totPos);
console.log('Clientes únicos:           '+qtdCliPos);
console.log('Meta POS:                  '+metaPosG);
console.log('% POS:                     '+ppPctG.toFixed(1)+'%');
console.log('');
console.log('Ticket médio:              '+fmt(fatBru/totPos));

// Per seller from base
console.log('\n--- MATRIZ POR VENDEDOR (base Excel Julho) ---');
var vends={};
julRows.forEach(function(r){
  if(!vends[r.codusur])vends[r.codusur]={codusur:r.codusur,fat:0,clientes:{}};
  vends[r.codusur].fat+=r.fat;
  if(!vends[r.codusur].clientes[r.codcli])vends[r.codusur].clientes[r.codcli]={fat:0,pos:0};
  vends[r.codusur].clientes[r.codcli].fat+=r.fat;
});
posJul.forEach(function(r){
  if(vends[r.codusur]&&vends[r.codusur].clientes[r.codcli])vends[r.codusur].clientes[r.codcli].pos++;
});

ALL_RCAS.forEach(function(c){
  var v=vends[c];if(!v)return;
  var mt=METAS[c]||{};
  var qc=Object.keys(v.clientes).length;
  var posC=Object.keys(v.clientes).filter(function(k){return v.clientes[k].pos>0;}).length;
  var pf=mt.fat>0?(v.fat/mt.fat*100):0;
  var pp=mt.pos>0?(posC/mt.pos*100):0;
  var vProj=Math.round(v.fat*pmul);
  var tk=qc>0?(v.fat/qc):0;
  console.log('  '+c+' '+((mt.nome||'?').padEnd(12))+
    ' Fat:'+fmt(v.fat).padStart(10)+
    ' Meta:'+fmt(mt.fat||0).padStart(10)+
    ' %F:'+pf.toFixed(1).padStart(6)+
    ' Proj:'+fmt(vProj).padStart(10)+
    ' Pos:'+posC+'/'+(mt.pos||'-')+
    ' %P:'+pp.toFixed(1).padStart(6)+
    ' Tick:'+fmt(tk).padStart(10));
});

// =============================================
// 6. Dashboard POLPANORTE - Big Numbers
// =============================================
console.log('\n--- 6. DASHBOARD POLPANORTE - BIG NUMBERS (base Excel) ---');

var fatBruP=polpaJul.reduce(function(s,r){return s+r.fat;},0);
var posP=polpaJul.filter(function(r){return r.fat>=1;});
var cliPosP={};
posP.forEach(function(r){cliPosP[r.codcli]=(cliPosP[r.codcli]||0)+r.fat;});
var qtdCliPosP=Object.keys(cliPosP).length;
var totPosP=posP.length;

var metaFatP=0,metaPosP=0;
Object.keys(METAS_POLPA).forEach(function(k){metaFatP+=METAS_POLPA[k].fat||0;metaPosP+=METAS_POLPA[k].pos||0;});
var pfPctP=metaFatP>0?fatBruP/metaFatP*100:0;
var ppPctP=metaPosP>0?qtdCliPosP/metaPosP*100:0;
var fatProjP=Math.round(fatBruP*pmul);

console.log('Faturamento Polpanorte:    '+fmt(fatBruP));
console.log('Meta FAT Polpanorte:       '+fmt(metaFatP));
console.log('% FAT Polpanorte:          '+pfPctP.toFixed(1)+'%');
console.log('Projeção FAT Polpanorte:   '+fmt(fatProjP)+' ('+fatProjP/metaFatP*100+'%)');
console.log('');
console.log('Clientes únicos Polpanorte:'+qtdCliPosP);
console.log('Meta POS Polpanorte:       '+metaPosP);
console.log('% POS Polpanorte:          '+ppPctP.toFixed(1)+'%');
console.log('');
console.log('Ticket médio Polpanorte:   '+fmt(totPosP>0?fatBruP/totPosP:0));

// Per seller Polpanorte
console.log('\n--- MATRIZ POLPANORTE POR VENDEDORA (base Excel Julho) ---');
var vp={};
polpaJul.forEach(function(r){
  if(!vp[r.codusur])vp[r.codusur]={codusur:r.codusur,fat:0,clientes:{}};
  vp[r.codusur].fat+=r.fat;
  if(!vp[r.codusur].clientes[r.codcli])vp[r.codusur].clientes[r.codcli]={fat:0,pos:0};
  vp[r.codusur].clientes[r.codcli].fat+=r.fat;
});
posP.forEach(function(r){
  if(vp[r.codusur]&&vp[r.codusur].clientes[r.codcli])vp[r.codusur].clientes[r.codcli].pos++;
});

ALL_RCAS.forEach(function(c){
  var v=vp[c];
  var mt=METAS_POLPA[c]||{};
  var fat=v?v.fat:0;
  var qc=v?Object.keys(v.clientes).length:0;
  var posC=v?Object.keys(v.clientes).filter(function(k){return v.clientes[k].pos>0;}).length:0;
  var pf=mt.fat>0?(fat/mt.fat*100):0;
  var pp=mt.pos>0?(posC/mt.pos*100):0;
  console.log('  '+c+' '+((mt.nome||'?').padEnd(12))+
    ' Fat:'+fmt(fat).padStart(10)+
    ' Meta:'+fmt(mt.fat||0).padStart(10)+
    ' %F:'+pf.toFixed(1).padStart(6)+
    ' Pos:'+posC+'/'+(mt.pos||'-')+
    ' %P:'+pp.toFixed(1).padStart(6));
});

// =============================================
// 7. Cross-check: valores dos arquivos vs cálculo direto
// =============================================
console.log('\n--- 7. CROSS-CHECK: ARQUIVOS vs CÁLCULO DIRETO ---');

// dados_mensal fat totals per month
for(var m=1;m<=7;m++){
  var extMT=0;
  ALL_RCAS.forEach(function(c){extMT+=dadosMes[m][c].fat;});
  var jsMT=0;
  ALL_RCAS.forEach(function(c){jsMT+=dm[m][c]?dm[m][c].fat:0;});
  check('Fat total mês '+mNames[m-1],extMT,jsMT,0.02);
}

// dados_polpa fat totals per month
for(var m=1;m<=7;m++){
  var extPT=0;
  ALL_RCAS.forEach(function(c){extPT+=polpaMeses[m][c].fat;});
  var jsPT=0;
  ALL_RCAS.forEach(function(c){jsPT+=pm[m][c]?pm[m][c].fat:0;});
  check('Polpa fat mês '+mNames[m-1],extPT,jsPT,0.02);
}

// July fat: all suppliers = dados_mensal[7] total?
var julAll=0;ALL_RCAS.forEach(function(c){julAll+=dadosMes[7][c].fat;});
var julPolpaAll=0;ALL_RCAS.forEach(function(c){julPolpaAll+=polpaMeses[7][c].fat;});
console.log('\nFat Julho GERAL (dados_mensal[7]):    '+fmt(julAll));
console.log('Fat Julho POLPANORTE (polpa[7]):      '+fmt(julPolpaAll));
console.log('Fat Julho TOTAL (base direta):        '+fmt(fatBru));
console.log('Diferença GERAL vs base direta:       R$ '+(julAll-fatBru).toFixed(2));
console.log('Diferença POLPA vs base direta:       R$ '+(julPolpaAll-fatBruP).toFixed(2));

// =============================================
// 8. Análise por Região - TOTAIS
// =============================================
console.log('\n--- 8. ANÁLISE POR REGIÃO (base Excel) ---');
var REGIAO_MES={
  'Reg. Metropolitana':{1:['1527'],2:['1527','1596'],3:['1596'],4:['1596'],5:['1596'],6:['1596'],7:['1596']},
  'Reg. Sul':{1:['1571'],2:['1571'],3:['1571'],4:['1571'],5:['1464','1571'],6:['1464'],7:['1464']},
  'Serra':{1:['9886'],2:['9886'],3:['9886'],4:['9886'],5:['9886'],6:['9886'],7:['9886']},
  'Litoral':{1:['1573'],2:['1573'],3:['1573','1464'],4:['1464','1573'],5:['1211'],6:['1211','1624'],7:['1624']},
  'Vale dos Sinos':{1:['1429'],2:['1429'],3:['1429'],4:['1429'],5:['1429'],6:['1429'],7:['1211']},
  'Key Account':{1:[],2:[],3:[],4:[],5:[],6:[],7:['1429']}
};
var regions=Object.keys(REGIAO_MES);
regions.forEach(function(reg){
  var rFTot=0;
  for(var m=1;m<=7;m++){
    var cods=REGIAO_MES[reg][m]||[];
    var rF=0;
    cods.forEach(function(c){if(dadosMes[m]&&dadosMes[m][c])rF+=dadosMes[m][c].fat;});
    rFTot+=rF;
  }
  console.log('  '+reg.padEnd(22)+'Fat Jan-Jul: '+fmt(rFTot).padStart(12));
});

var gRTot=0;
regions.forEach(function(reg){
  for(var m=1;m<=7;m++){
    var cods=REGIAO_MES[reg][m]||[];
    cods.forEach(function(c){if(dadosMes[m]&&dadosMes[m][c])gRTot+=dadosMes[m][c].fat;});
  }
});
console.log('  TOTAL REGIÕES:             '+fmt(gRTot));
console.log('  TOTAL BASE:                '+fmt(allRows.reduce(function(s,r){return s+r.fat;},0)));
console.log('  Diferença:                 R$ '+(gRTot-allRows.reduce(function(s,r){return s+r.fat;},0)).toFixed(2));

// =============================================
// 9. Validação cruzada: vendedor julho fat dados_mensal vs dados_julho
// =============================================
console.log('\n--- 9. FAT JULHO POR VENDEDOR: dados_mensal[7] vs dados_julho.json ---');
ALL_RCAS.forEach(function(c){
  var dmF=dm[7][c]?dm[7][c].fat:0;
  var djF=djByCod[c]?djByCod[c].fat:0;
  var extF=extByCod[c]?extByCod[c].fat:0;
  var d1=Math.abs(dmF-djF);
  var d2=Math.abs(dmF-extF);
  var status=(d1<0.1&&d2<0.1)?'✅':'❌';
  console.log('  '+status+' '+c+' '+((METAS[c]||{}).nome||'?').padEnd(12)+
    ' dados_mensal: '+fmt(dmF).padStart(12)+
    ' dados_julho: '+fmt(djF).padStart(12)+
    ' base_direta: '+fmt(extF).padStart(12)+
    ' diff1: R$'+d1.toFixed(2).padStart(8)+
    ' diff2: R$'+d2.toFixed(2).padStart(8));
});

// =============================================
// 10. VALIDAÇÃO TOTAL
// =============================================
console.log('\n==============================================');
if(errs===0)console.log('✅ TODAS AS VALIDAÇÕES PASSARAM - NUMEROS CORRETOS');
else console.log('❌ '+errs+' VALIDAÇÕES FALHARAM');
console.log('==============================================');
