"""
motor_reportes.py - Motor de reportes de avance DDH (KPDI / MREDP 2026), Área de Geología.

A partir del Excel de control de avances genera:
  - Dashboard interactivo (HTML)
  - Reporte gerencial resumido (Word)
  - Presentación de avances (PPT)

Se usa igual desde Google Colab y desde la app de escritorio:
    from motor_reportes import generar_todo
    generar_todo("Control_de_Avances.xlsx", fecha=None, carpeta="salida")
"""
# ---------------------------------------------------------------- plantilla del dashboard
TEMPLATE = r'''<title>Avance DDH MREDP 2026</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@500&display=swap">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
/* Layout: cabecera con corte de datos, franja de KPI, luego secciones de lectura (perforación, control de datos y muestras, alertas, leyes) y el detalle por DDH al final. */
:root{
  --bg:#F2F5F8; --surface:#FFFFFF; --ink:#15212D; --ink2:#53626F; --line:#DCE3EA; --soft:#EAF0F6;
  --accent:#1F5E99;
  --s-recv:#A7C9EB; --s-send:#6E97BD; --s-done:#1F5E99; --meta:#D9D9D9;
  --good:#2A7B58; --warn:#A8690F; --crit:#B3372F;
  --good-bg:#E4F3EC; --warn-bg:#FBF0DC; --crit-bg:#FBE6E3;
  --font:'IBM Plex Sans',system-ui,-apple-system,'Segoe UI',sans-serif;
  --mono:'IBM Plex Mono',ui-monospace,Menlo,Consolas,monospace;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --bg:#0D1319; --surface:#151D26; --ink:#E5ECF3; --ink2:#9BA9B7; --line:#25303C; --soft:#1B2530;
    --accent:#6FA8E0;
    --s-recv:#476A90; --s-send:#6F9BC8; --s-done:#A9D0F5; --meta:#2F3A47;
    --good:#5CC49A; --warn:#E3B45A; --crit:#F0877E;
    --good-bg:#16302A; --warn-bg:#33291A; --crit-bg:#37201F;
    color-scheme:dark;
  }
}
:root[data-theme="dark"]{
  --bg:#0D1319; --surface:#151D26; --ink:#E5ECF3; --ink2:#9BA9B7; --line:#25303C; --soft:#1B2530;
  --accent:#6FA8E0;
  --s-recv:#476A90; --s-send:#6F9BC8; --s-done:#A9D0F5; --meta:#2F3A47;
  --good:#5CC49A; --warn:#E3B45A; --crit:#F0877E;
  --good-bg:#16302A; --warn-bg:#33291A; --crit-bg:#37201F;
  color-scheme:dark;
}
*{box-sizing:border-box}
body{background:var(--bg);color:var(--ink);font-family:var(--font);font-size:14px;line-height:1.45;
  padding-block:0 48px;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
.wrap{max-width:1240px;margin:0 auto}
h1,h2,h3{margin:0;font-weight:600;text-wrap:balance}
.num{font-variant-numeric:tabular-nums}
.mono{font-family:var(--mono)}
.label{font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:var(--ink2);font-weight:500}

header.top{display:flex;flex-wrap:wrap;gap:12px 32px;justify-content:space-between;align-items:flex-end;padding-block:28px 20px;border-bottom:1px solid var(--line)}
header.top h1{font-size:26px;letter-spacing:-.01em}
header.top p{margin:4px 0 0;color:var(--ink2)}
.stamp{text-align:right;color:var(--ink2);font-size:12px}
.stamp b{color:var(--ink);font-weight:500}

nav.jump{display:flex;flex-wrap:wrap;gap:6px;padding-block:14px}
nav.jump a{color:var(--ink2);text-decoration:none;font-size:13px;padding:5px 11px;border:1px solid var(--line);border-radius:999px;background:var(--surface)}
nav.jump a:hover{color:var(--ink);border-color:var(--accent)}

.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:12px;margin-top:4px}
.kpi{background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:16px 18px;min-width:0}
.kpi .v{font-size:32px;font-weight:600;letter-spacing:-.02em;margin-top:6px;line-height:1.1}
.kpi .v small{font-size:15px;font-weight:400;color:var(--ink2);letter-spacing:0}
.kpi .s{color:var(--ink2);font-size:12.5px;margin-top:6px}
.meter{height:6px;background:var(--soft);border-radius:3px;margin-top:12px;overflow:hidden}
.meter i{display:block;height:100%;background:var(--accent)}

.seg{display:flex;height:10px;border-radius:5px;overflow:hidden;margin-top:12px;background:var(--soft)}
.seg i{display:block;height:100%}
.seg .c{background:var(--s-done)} .seg .p{background:var(--warn)} .seg .n{background:var(--meta)}
.legend-inline{display:flex;flex-wrap:wrap;gap:4px 14px;margin-top:8px;font-size:12px;color:var(--ink2)}
.dot{display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:5px;vertical-align:-1px}

section{margin-top:36px}
.sec-head{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:baseline;gap:6px 20px;margin-bottom:12px}
.sec-head h2{font-size:18px}
.sec-head p{margin:0;color:var(--ink2);font-size:13px;max-width:62ch}
.grid{display:grid;gap:14px}
.g2{grid-template-columns:repeat(auto-fit,minmax(min(100%,420px),1fr))}
.g3{grid-template-columns:repeat(auto-fit,minmax(min(100%,360px),1fr))}
.card{background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:16px 18px;min-width:0}
.card h3{font-size:14px;margin-bottom:2px}
.card .sub{color:var(--ink2);font-size:12.5px;margin-bottom:10px}
.chart{position:relative;width:100%}
.note{color:var(--ink2);font-size:12px;margin:10px 0 0}
.total-block{border-top:1px solid var(--line);margin-top:2px;padding-top:6px}
.total-block .label{font-size:10px;margin-left:2px}

.alerts{display:grid;gap:8px}
.alert{display:grid;grid-template-columns:auto 1fr;gap:2px 12px;padding:12px 14px;border-radius:6px;background:var(--surface);border:1px solid var(--line);border-left-width:4px}
.alert.crit{border-left-color:var(--crit)} .alert.warn{border-left-color:var(--warn)} .alert.info{border-left-color:var(--s-send)}
.pill{font-size:10.5px;letter-spacing:.07em;text-transform:uppercase;font-weight:600;padding:2px 8px;border-radius:999px;align-self:start;white-space:nowrap}
.pill.crit{background:var(--crit-bg);color:var(--crit)} .pill.warn{background:var(--warn-bg);color:var(--warn)} .pill.info{background:var(--soft);color:var(--ink2)}
.pill.ok{background:var(--good-bg);color:var(--good)}
.alert b{font-weight:600} .alert span.d{grid-column:2;color:var(--ink2);font-size:13px}

.seg-ctl{display:inline-flex;flex-wrap:wrap;gap:4px}
.seg-ctl button,.chip{font:inherit;font-size:12.5px;color:var(--ink2);background:var(--surface);border:1px solid var(--line);border-radius:999px;padding:5px 12px;cursor:pointer}
.seg-ctl button[aria-pressed="true"],.chip[aria-pressed="true"]{background:var(--accent);border-color:var(--accent);color:var(--surface)}
.seg-ctl button:hover,.chip:hover{border-color:var(--accent)}
button:focus-visible,select:focus-visible,input:focus-visible,a:focus-visible,th[tabindex]:focus-visible{outline:2px solid var(--accent);outline-offset:2px}

.tools{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:12px}
.tools input,.tools select{font:inherit;font-size:13px;padding:6px 10px;background:var(--surface);color:var(--ink);border:1px solid var(--line);border-radius:6px}
.tools input{min-width:200px}
.tablewrap{overflow-x:auto;border:1px solid var(--line);border-radius:8px;background:var(--surface);max-height:560px;overflow-y:auto}
table{border-collapse:collapse;width:100%;min-width:860px}
th,td{padding:8px 12px;text-align:left;border-bottom:1px solid var(--line);white-space:nowrap}
th{position:sticky;top:0;background:var(--surface);font-size:11px;letter-spacing:.07em;text-transform:uppercase;color:var(--ink2);font-weight:500;cursor:pointer;user-select:none;z-index:1}
th.r,td.r{text-align:right}
tbody tr:hover{background:var(--soft)}
td.id{font-family:var(--mono);font-size:12.5px}
td.muted{color:var(--ink2)}
.bar{display:inline-block;width:90px;height:6px;background:var(--soft);border-radius:3px;vertical-align:middle;margin-right:8px;overflow:hidden}
.bar i{display:block;height:100%;background:var(--accent)}
.bar i.over{background:var(--warn)}
.dots{display:inline-flex;gap:3px}
.dots i{width:8px;height:8px;border-radius:2px;background:var(--soft);border:1px solid var(--line)}
.dots i.on{background:var(--s-done);border-color:var(--s-done)}
.count{color:var(--ink2);font-size:12.5px;margin-left:auto}

.leyes-empty{display:grid;gap:8px;color:var(--ink2)}
.leyes-empty code{font-family:var(--mono);font-size:12px;background:var(--soft);padding:2px 6px;border-radius:4px;color:var(--ink)}
.stat-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin-bottom:12px}
.stat-row div{background:var(--soft);border-radius:6px;padding:10px 12px}
.stat-row .v{font-size:20px;font-weight:600}

footer{margin-top:40px;color:var(--ink2);font-size:12px}
@media (max-width:560px){
  header.top h1{font-size:22px} .stamp{text-align:left}
  .kpi .v{font-size:28px}
}
@media (prefers-reduced-motion:no-preference){ .meter i,.seg i{transition:width .5s ease} }
</style>

<div class="wrap">
  <header class="top">
    <div>
      <div class="label">KPDI · Programa MREDP 2026</div>
      <h1>Avance de perforación diamantina</h1>
      <p id="subtitle"></p>
    </div>
    <div class="stamp"><div>Archivo: <b id="st-file"></b></div><div>Procesado: <b id="st-date"></b></div></div>
  </header>
  <nav class="jump" aria-label="Secciones">
    <a href="#perforacion">Perforación</a><a href="#control">Datos y muestras</a><a href="#alertas">Alertas</a><a href="#leyes">Leyes</a><a href="#detalle">Detalle por DDH</a>
  </nav>

  <div class="kpis" id="kpis"></div>

  <section id="perforacion">
    <div class="sec-head"><h2>Perforación</h2><p>Metros ejecutados por equipo y recorrido del testigo por cada etapa de recepción y logueo.</p></div>
    <div class="grid g2">
      <div class="card"><h3>Metros por equipo</h3><div class="sub">Ejecutado frente a programado de los DDH asignados</div>
        <div class="chart" id="wrap-eq" style="height:300px"><canvas id="c-eq" role="img" aria-label="Metros ejecutados y programados por equipo"></canvas></div>
        <div class="total-block"><div class="label">Total del programa</div>
          <div class="chart" style="height:92px"><canvas id="c-ddh" role="img" aria-label="DDH completados frente a programados, con metraje"></canvas></div></div>
        <div class="legend-inline" style="justify-content:center"><span><span class="dot" style="background:var(--s-done)"></span>Ejecutado</span><span><span class="dot" style="background:var(--meta)"></span>Programado</span></div>
        <p class="note" id="eq-note"></p></div>
      <div class="card"><h3>Recorrido por etapa</h3><div class="sub">Metros que superaron cada etapa, sobre el total perforado</div>
        <div class="chart" style="height:300px"><canvas id="c-et" role="img" aria-label="Metros por etapa de procesamiento"></canvas></div>
        <p class="note" id="et-note"></p></div>
    </div>
  </section>

  <section id="control">
    <div class="sec-head"><h2>Control de avance de datos y muestras</h2><p>Cada barra se mide contra los 131 DDH programados (meta). Los porcentajes salen de la hoja Data.</p></div>
    <div class="grid g2">
      <div class="card"><h3>Datos validados</h3><div class="sub">Collar, Survey y Assay: KPDI frente a MCP</div>
        <div class="chart" style="height:300px"><canvas id="c-val" role="img" aria-label="Validación de Collar, Survey y Assay"></canvas></div></div>
      <div class="card"><h3>Corte y muestreo</h3><div class="sub">Ejecutado frente a DDH disponibles</div>
        <div class="chart" style="height:300px"><canvas id="c-cm" role="img" aria-label="Corte y muestreo ejecutado"></canvas></div></div>
      <div class="card" style="grid-column:1/-1"><h3>Laboratorio</h3><div class="sub">Preparado, enviado y recibido por tipo de muestra</div>
        <div class="chart" style="height:400px"><canvas id="c-lab" role="img" aria-label="Preparado, enviado y recibido por tipo de muestra"></canvas></div>
</div>
    </div>
  </section>

  <section id="alertas">
    <div class="sec-head"><h2>Alertas y cuellos de botella</h2><p>Generadas desde los datos en cada corrida, ordenadas por severidad.</p></div>
    <div class="alerts" id="alerts"></div>
  </section>

  <section id="leyes">
    <div class="sec-head"><h2>Resultados de leyes</h2><p id="leyes-p"></p></div>
    <div id="leyes-body"></div>
  </section>

  <section id="detalle">
    <div class="sec-head"><h2>Detalle por DDH</h2><p>Filtra, ordena y revisa el estado de cada sondaje.</p></div>
    <div class="tools">
      <div class="seg-ctl" id="f-estado" role="group" aria-label="Filtrar por estado"></div>
      <select id="f-eq" aria-label="Filtrar por equipo"></select>
      <input id="f-q" type="search" placeholder="Buscar DDH o plataforma" aria-label="Buscar DDH o plataforma">
      <span class="count num" id="f-count"></span>
    </div>
    <div class="tablewrap"><table><thead><tr id="thead"></tr></thead><tbody id="tbody"></tbody></table></div>
    <p class="note">Etapas (de izquierda a derecha): Collar, Survey, Assay, Logueo, Corte, Muestreo, Lotes recibidos, Densidad recibida, XRD recibido, Check Assay recibido. Cuadro lleno = registrado.</p>
  </section>

  <footer>Generado con build_dashboard.py desde el Excel de control. Los conteos se validan contra el resumen del mismo archivo en cada corrida.</footer>
</div>

<script>
const D = window.__DATA__ = /*__DATA__*/null;
const $ = id => document.getElementById(id);
const fmt = (n, d=0) => Number(n).toLocaleString('es-PE', {minimumFractionDigits:d, maximumFractionDigits:d});
const pct = (a, b, d=0) => b ? (100*a/b).toFixed(d)+'%' : '0%';
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

/* ---------- cabecera y KPI ---------- */
const K = D.kpi, M = K.m, N = K.ddh;
$('st-file').textContent = D.meta.archivo; $('st-date').textContent = D.meta.generado;
$('subtitle').textContent = `${N.total} sondajes programados · ${fmt(M.prog)} m · ${N.completado} completados, ${N.en_proceso} en proceso, ${N.pendiente} pendientes`;

const tiles = [
  `<div class="kpi"><div class="label">DDH completados</div><div class="v num">${N.completado} <small>/ ${N.total}</small></div>
    <div class="seg" role="img" aria-label="Completados, en proceso y pendientes"><i class="c" style="width:${100*N.completado/N.total}%"></i><i class="p" style="width:${100*N.en_proceso/N.total}%"></i><i class="n" style="width:${100*N.pendiente/N.total}%"></i></div>
    <div class="legend-inline"><span><span class="dot" style="background:var(--s-done)"></span>${N.completado} completados</span><span><span class="dot" style="background:var(--warn)"></span>${N.en_proceso} en proceso</span><span><span class="dot" style="background:var(--meta)"></span>${N.pendiente} pendientes</span></div></div>`,
  `<div class="kpi"><div class="label">Metros perforados</div><div class="v num">${fmt(M.ejec)} <small>m</small></div>
    <div class="meter"><i style="width:${100*M.ejec/M.prog}%"></i></div>
    <div class="s num">${pct(M.ejec,M.prog,1)} de ${fmt(M.prog)} m programados</div></div>`,
  `<div class="kpi"><div class="label">Cumplimiento en DDH iniciados</div><div class="v num">${pct(M.ejec,M.prog_iniciados,1)}</div>
    <div class="s num">${fmt(M.ejec)} m perforados sobre ${fmt(M.prog_iniciados)} m programados de los DDH ya iniciados</div></div>`,
  `<div class="kpi"><div class="label">Metros por perforar</div><div class="v num">${fmt(M.pend)} <small>m</small></div>
    <div class="s num">${N.pendiente} DDH pendientes · ${pct(M.pend,M.prog,1)} del metraje global programado</div></div>`,
];
$('kpis').innerHTML = tiles.join('');

/* ---------- gráficos ---------- */
const endLabels = {
  id:'endLabels',
  afterDatasetsDraw(chart){
    const {ctx} = chart; ctx.save();
    ctx.font = `500 12px ${css('--font')}`; ctx.textBaseline = 'middle'; ctx.fillStyle = css('--ink');
    chart.data.datasets.forEach((ds, i) => {
      const meta = chart.getDatasetMeta(i); if (meta.hidden) return;
      meta.data.forEach((bar, j) => { const t = ds.lab?.[j]; if (t != null) ctx.fillText(t, bar.x + 6, bar.y); });
    });
    ctx.restore();
  }
};
let charts = [];
function hbar(id, labels, sets, max, top){
  const c = new Chart($(id), {
    type:'bar', plugins:[endLabels],
    data:{labels, datasets:sets.map(s => ({label:s.label, data:s.data, lab:s.lab, tip:s.tip, backgroundColor:css(s.color), borderWidth:0, barPercentage:.92, categoryPercentage:.82}))},
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false, animation:{duration:300},
      layout:{padding:{right:8}},
      scales:{
        x:{min:0, max, ...(top?{position:'top'}:{}), ticks:{color:css('--ink2'), ...(max>200?{maxTicksLimit:6,maxRotation:0,callback:v=>fmt(v)}:{stepSize:20}), font:{family:css('--font')}}, grid:{color:css('--line')}, border:{display:false}},
        y:{afterFit:top?(s=>{s.width=150}):undefined, ticks:{color:css('--ink'), font:{family:css('--font'), size:13}}, grid:{display:false}, border:{color:css('--line')}}
      },
      plugins:{
        legend:{display:!top, position:'bottom', labels:{color:css('--ink2'), boxWidth:10, boxHeight:10, usePointStyle:true, pointStyle:'rect', font:{family:css('--font')}}},
        tooltip:{callbacks:{label:ctx => `${ctx.dataset.label}: ${ctx.dataset.tip?.[ctx.dataIndex] ?? fmt(ctx.parsed.x)}`}}
      }
    }
  });
  charts.push(c); return c;
}
const CT = D.control;
const lab1 = (n, t, d=0) => pct(n, t, d);
const withTip = (arr, t) => arr.map(n => `${fmt(n)} de ${t} (${pct(n,t,1)})`);

function drawCharts(){
  charts.forEach(c => c.destroy()); charts = [];

  /* equipos */
  const eq = D.equipos;
  $('wrap-eq').style.height = (110 + eq.length*58) + 'px';
  hbar('c-eq', eq.map(e => `${e.equipo} (${e.ddh})`), [
    {label:'Ejecutado', color:'--s-done', data:eq.map(e=>e.ejec), lab:eq.map(e=>`${fmt(e.ejec)} m · ${pct(e.ejec,e.prog)}`), tip:eq.map(e=>`${fmt(e.ejec)} m (${pct(e.ejec,e.prog,1)} de lo programado)`)},
    {label:'Programado', color:'--meta', data:eq.map(e=>e.prog), lab:eq.map(e=>`${fmt(e.prog)} m`), tip:eq.map(e=>`${fmt(e.prog)} m`)}
  ], Math.ceil(Math.max(...eq.map(e=>e.prog))*1.32/1000)*1000, true);
  const dd = [[`Ejecutado (${N.completado})`, N.completado, '--s-done', `${pct(N.completado,N.total,1)} DDH · ${fmt(M.ejec)} m (${pct(M.ejec,M.prog,1)})`], [`Programado (${N.total})`, N.total, '--meta', `${fmt(M.prog)} m`]];
  charts.push(new Chart($('c-ddh'), {type:'bar', plugins:[endLabels], data:{labels:dd.map(d=>d[0]), datasets:[{data:dd.map(d=>d[1]), lab:dd.map(d=>d[3]), backgroundColor:dd.map(d=>css(d[2])), borderWidth:0, barPercentage:.85, categoryPercentage:.95}]},
    options:{indexAxis:'y', responsive:true, maintainAspectRatio:false, animation:{duration:300}, layout:{padding:{right:8}},
      scales:{x:{display:false, min:0, max:Math.round(N.total*1.75)}, y:{afterFit:s=>{s.width=150}, ticks:{color:css('--ink'), font:{family:css('--font'), size:13}}, grid:{display:false}, border:{display:false}}},
      plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>`${fmt(c.parsed.x)} DDH`}}}}}));
  $('eq-note').textContent = `Entre paréntesis, DDH por equipo. Pendiente por perforar, sin equipo asignado: ${N.pendiente} DDH y ${fmt(M.pend)} m (${pct(M.pend,M.prog,1)} del metraje global programado).`;

  /* etapas */
  const et = D.etapas, base = et[0].m;
  hbar('c-et', et.map(e=>e.label), [
    {label:'Metros', color:'--s-done', data:et.map(e=>e.m), lab:et.map(e=>`${fmt(e.m)} m · ${pct(e.m,base,1)}`), tip:et.map(e=>`${fmt(e.m)} m (${pct(e.m,base,1)} de lo perforado)`)}
  ], Math.ceil(base*1.38/1000)*1000).options.plugins.legend.display = false;
  charts[charts.length-1].update();
  const gap = base - et[et.length-1].m;
  $('et-note').textContent = `Hay ${fmt(gap)} m perforados sin logueo geológico (${pct(gap,base,1)} del total).`;

  /* validación (referencia: gráfico 3) */
  const v = [...CT.validacion].reverse();
  hbar('c-val', v.map(x=>x.item), [
    {label:'Validado_MCP', color:'--s-send', data:v.map(x=>x.mcp), lab:v.map(x=>lab1(x.mcp,x.meta,1)), tip:withTip(v.map(x=>x.mcp),131)},
    {label:'Validado_KPDI', color:'--s-done', data:v.map(x=>x.kpdi), lab:v.map(x=>lab1(x.kpdi,x.meta,1)), tip:withTip(v.map(x=>x.kpdi),131)},
    {label:'Programado (Meta)', color:'--meta', data:v.map(x=>x.meta), lab:v.map(x=>String(x.meta)), tip:v.map(x=>String(x.meta))}
  ], 150);

  /* corte y muestreo (referencia: gráfico 2) */
  const cm = [...CT.corte_muestreo].reverse();
  hbar('c-cm', cm.map(x=>x.item), [
    {label:'Ejecutado', color:'--s-send', data:cm.map(x=>x.ejecutado), lab:cm.map(x=>lab1(x.ejecutado,x.meta)), tip:withTip(cm.map(x=>x.ejecutado),131)},
    {label:'Disponible', color:'--s-done', data:cm.map(x=>x.disponible), lab:cm.map(x=>lab1(x.disponible,x.meta)), tip:withTip(cm.map(x=>x.disponible),131)},
    {label:'Programado (Meta)', color:'--meta', data:cm.map(x=>x.meta), lab:cm.map(x=>String(x.meta)), tip:cm.map(x=>String(x.meta))}
  ], 150);

  /* laboratorio (referencia: gráfico 1) */
  const lb = [...CT.laboratorio].reverse();
  hbar('c-lab', lb.map(x=>x.item), [
    {label:'Recepción', color:'--s-recv', data:lb.map(x=>x.recepcion), lab:lb.map(x=>lab1(x.recepcion,x.meta)), tip:withTip(lb.map(x=>x.recepcion),131)},
    {label:'Enviado', color:'--s-send', data:lb.map(x=>x.enviado), lab:lb.map(x=>lab1(x.enviado,x.meta)), tip:withTip(lb.map(x=>x.enviado),131)},
    {label:'Ejecutado /Disponible', color:'--s-done', data:lb.map(x=>x.ejecutado), lab:lb.map(x=>lab1(x.ejecutado,x.meta)), tip:withTip(lb.map(x=>x.ejecutado),131)},
    {label:'Programado (Meta)', color:'--meta', data:lb.map(x=>x.meta), lab:lb.map(x=>String(x.meta)), tip:lb.map(x=>String(x.meta))}
  ], 150);

  if (D.leyes) drawLeyes();
}

/* ---------- alertas ---------- */
const SEV = {crit:'Crítica', warn:'Atención', info:'Revisar'};
const order = {crit:0, warn:1, info:2};
$('alerts').innerHTML = [...D.alertas].sort((a,b)=>order[a.sev]-order[b.sev]).map(a =>
  `<div class="alert ${a.sev}"><span class="pill ${a.sev}">${SEV[a.sev]}</span><b>${esc(a.title)}</b><span class="d">${esc(a.detail)}</span></div>`).join('');

/* ---------- leyes ---------- */
let leyEl = null;
function leyesUI(){
  const L = D.leyes;
  if (!L){
    $('leyes-p').textContent = 'El Excel de control no incluye leyes, solo el conteo de DDH con assay validado.';
    $('leyes-body').innerHTML = `<div class="card leyes-empty"><div><b style="color:var(--ink)">Sin datos de leyes en este archivo.</b> ${CT.validacion[2].kpdi} de ${D.meta.total} DDH tienen assay validado por KPDI (${pct(CT.validacion[2].kpdi,D.meta.total,1)}).</div>
      <div>Para activar esta sección agrega una hoja <code>Leyes</code> con las columnas <code>DDHID</code>, <code>Desde</code>, <code>Hasta</code> y una columna por elemento (<code>Zn</code>, <code>Pb</code>, <code>Ag</code>, <code>Cu</code>, <code>Au</code>) y vuelve a ejecutar el script.</div></div>`;
    return;
  }
  $('leyes-p').textContent = 'Leyes ponderadas por longitud de muestra, por elemento y por DDH.';
  leyEl = L.elements[0];
  $('leyes-body').innerHTML = `<div class="card"><div class="seg-ctl" id="ley-sel" role="group" aria-label="Elemento">${L.elements.map(e=>`<button aria-pressed="${e===leyEl}" data-el="${esc(e)}">${esc(e)}</button>`).join('')}</div>
    <div id="ley-stats" class="stat-row" style="margin-top:12px"></div>
    <div class="chart" style="height:420px"><canvas id="c-ley" role="img" aria-label="DDH con mayor ley ponderada"></canvas></div></div>`;
  $('ley-sel').addEventListener('click', e => { const b = e.target.closest('button'); if(!b) return; leyEl = b.dataset.el;
    document.querySelectorAll('#ley-sel button').forEach(x=>x.setAttribute('aria-pressed', x===b)); drawCharts(); });
}
function drawLeyes(){
  const L = D.leyes, s = L.stats[leyEl], top = L.holes[leyEl].slice(0,15);
  $('ley-stats').innerHTML = [['Ley media',fmt(s.mean,2)],['Mediana',fmt(s.p50,2)],['Máxima',fmt(s.max,2)],['Muestras',fmt(s.n)],['Metros',fmt(s.m)]]
    .map(([k,v])=>`<div><div class="label">${k}</div><div class="v num">${v}</div></div>`).join('');
  const mx = Math.max(...top.map(h=>h.grade));
  const c = hbar('c-ley', top.map(h=>h.id), [{label:`${leyEl} ponderado`, color:'--s-done', data:top.map(h=>h.grade), lab:top.map(h=>fmt(h.grade,2)), tip:top.map(h=>`${fmt(h.grade,2)} en ${fmt(h.len)} m`)}], Math.ceil(mx*1.2));
  c.options.plugins.legend.display = false; c.update();
}

/* ---------- tabla por DDH ---------- */
const COLS = [
  {k:'id', t:'DDH', f:h=>h.id || 'Sin ID'},
  {k:'plat', t:'Plataforma'},
  {k:'eq', t:'Equipo', f:h=>h.eq || ''},
  {k:'estado', t:'Estado'},
  {k:'prog', t:'Prog. (m)', r:1},
  {k:'ejec', t:'Ejec. (m)', r:1},
  {k:'pct', t:'Avance', r:0},
  {k:'etapas', t:'Etapas', nosort:1},
];
const STG = [['collar','Collar'],['survey','Survey'],['assay','Assay'],['log','Logueo'],['corte','Corte'],['muestreo','Muestreo'],['lotes','Lotes recibidos'],['dens','Densidad recibida'],['xrd','XRD recibido'],['check','Check Assay recibido']];
const ST = {state:'Todos', eq:'Todos', q:'', key:'estado', dir:1};
const estOrder = {'En Proceso':0,'Completado':1,'Pendiente':2};

$('thead').innerHTML = COLS.map(c=>`<th class="${c.r?'r':''}" data-k="${c.k}" ${c.nosort?'':'tabindex="0"'}>${c.t}</th>`).join('');
$('f-estado').innerHTML = ['Todos','Completado','En Proceso','Pendiente'].map(s=>`<button aria-pressed="${s==='Todos'}" data-s="${s}">${s==='En Proceso'?'En proceso':s}</button>`).join('');
const eqs = ['Todos', ...new Set(D.holes.map(h=>h.eq).filter(Boolean))].sort((a,b)=>a==='Todos'?-1:b==='Todos'?1:a.localeCompare(b));
$('f-eq').innerHTML = eqs.map(e=>`<option>${e}</option>`).join('');

function rows(){
  const q = ST.q.trim().toLowerCase();
  let r = D.holes.filter(h => (ST.state==='Todos'||h.estado===ST.state) && (ST.eq==='Todos'||h.eq===ST.eq) &&
    (!q || (h.id||'').toLowerCase().includes(q) || (h.plat||'').toLowerCase().includes(q)));
  const k = ST.key;
  r.sort((a,b)=>{
    let x, y;
    if (k==='estado'){ x=estOrder[a.estado]; y=estOrder[b.estado]; if (x===y){ x=a.id||'~'; y=b.id||'~'; } }
    else { x=a[k]??''; y=b[k]??''; }
    return (x<y?-1:x>y?1:0)*ST.dir;
  });
  return r;
}
function pill(e){ const m={'Completado':'ok','En Proceso':'warn','Pendiente':'info'}[e]; return `<span class="pill ${m}">${e==='En Proceso'?'En proceso':e}</span>`; }
function drawTable(){
  const r = rows();
  $('f-count').textContent = `${r.length} de ${D.holes.length} DDH`;
  $('tbody').innerHTML = r.map(h => {
    const w = Math.min(100,h.pct), over = h.pct>105;
    const dots = STG.map(([k,l])=>`<i class="${h.f[k]?'on':''}" title="${l}: ${h.f[k]?'sí':'no'}"></i>`).join('');
    return `<tr><td class="id">${h.id?esc(h.id):'<span style="color:var(--ink2)">Sin ID</span>'}</td><td>${esc(h.plat)}</td><td class="${h.eq?'':'muted'}">${h.eq?esc(h.eq):'Sin asignar'}</td>
      <td>${pill(h.estado)}</td><td class="r num">${fmt(h.prog,1)}</td><td class="r num">${fmt(h.ejec,1)}</td>
      <td class="r num"><span class="bar"><i class="${over?'over':''}" style="width:${w}%"></i></span>${h.pct}%</td><td><span class="dots">${dots}</span></td></tr>`;
  }).join('') || '<tr><td colspan="8" class="muted">Ningún DDH coincide con el filtro.</td></tr>';
}
$('f-estado').addEventListener('click', e=>{ const b=e.target.closest('button'); if(!b) return; ST.state=b.dataset.s;
  document.querySelectorAll('#f-estado button').forEach(x=>x.setAttribute('aria-pressed', x===b)); drawTable(); });
$('f-eq').addEventListener('change', e=>{ ST.eq=e.target.value; drawTable(); });
$('f-q').addEventListener('input', e=>{ ST.q=e.target.value; drawTable(); });
function sortBy(th){ const k=th.dataset.k; if (!k||k==='etapas') return; ST.dir = ST.key===k ? -ST.dir : 1; ST.key=k; drawTable(); }
$('thead').addEventListener('click', e=>{ const th=e.target.closest('th'); if(th) sortBy(th); });
$('thead').addEventListener('keydown', e=>{ if(e.key==='Enter'||e.key===' '){ const th=e.target.closest('th'); if(th){ e.preventDefault(); sortBy(th); } } });

/* ---------- arranque y cambio de tema ---------- */
leyesUI(); drawTable(); drawCharts();
let t; const redraw = ()=>{ clearTimeout(t); t=setTimeout(drawCharts, 60); };
matchMedia('(prefers-color-scheme: dark)').addEventListener?.('change', redraw);
new MutationObserver(redraw).observe(document.documentElement, {attributes:true, attributeFilter:['data-theme']});
</script>
'''

# ---------------------------------------------------------------- procesamiento de datos
import datetime as dt
import json
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd


STAGE_COLS = ["EJECUTADO", "RECEPCION", "RECUPERADO", "FOTOGRAFIADO", "GEOTECNICO", "GEOLOGICO"]
STAGE_LABELS = ["Perforado", "Recepción", "Recuperado", "Fotografiado", "Geotécnico", "Geológico"]

# flags por DDH: clave corta -> columna del Excel
FLAGS = {
    "collar": "Collar_KPDI", "survey": "Survey_KPDI", "assay": "Assay_KPDI",
    "log": "DDH_Logueado", "corte": "Corte Ejecutado", "muestreo": "Muestreo Ejecutado",
    "lotes": "Lotes_Recepción", "dens": "Densidad_Recepeción",
    "xrd": "XRD_Recepción", "check": "Check Assay_Recepción",
}


def num(x, nd=1):
    return round(float(x), nd)


def load(path):
    df = pd.read_excel(path, sheet_name="Data").iloc[:, :36]
    for c in STAGE_COLS + ["PROGRAMADO", "PCT"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["EQUIPO"] = df["EQUIPO"].where(df["EQUIPO"].notna(), None)
    df["DDHID"] = df["DDHID"].where(df["DDHID"].astype(str).str.strip().ne("") & df["DDHID"].notna(), None)
    return df


def count(df, col):
    return int(df[col].notna().sum()) if col in df.columns else 0


def holes_payload(df):
    out = []
    for _, r in df.iterrows():
        h = {
            "id": r["DDHID"],
            "plat": r["PLATAFORMA"],
            "eq": r["EQUIPO"],
            "estado": r["ESTADO"],
            "prog": num(r["PROGRAMADO"]),
            "ejec": num(r["EJECUTADO"]),
            "geol": num(r["GEOLOGICO"]),
            "pct": int(r["PCT"]),
            "f": {k: int(pd.notna(r[c])) for k, c in FLAGS.items()},
        }
        out.append(h)
    return out


def alerts(df, ctl, total):
    a = []
    comp = df[df["ESTADO"] == "Completado"]

    v = {x["item"]: x for x in ctl["validacion"]}
    a.append({
        "sev": "crit",
        "title": "Validación MCP sin avance",
        "detail": f"0 de {total} DDH validados por MCP en Collar, Survey y Assay. "
                  f"KPDI va en {v['Collar']['kpdi']} / {v['Survey']['kpdi']} / {v['Assay']['kpdi']} "
                  f"({v['Collar']['kpdi']/total:.0%} / {v['Survey']['kpdi']/total:.0%} / {v['Assay']['kpdi']/total:.0%}).",
    })

    lab = {x["item"]: x for x in ctl["laboratorio"]}
    x = lab["XRD"]
    a.append({
        "sev": "crit",
        "title": "XRD: casi sin resultados recibidos",
        "detail": f"{x['ejecutado']} muestras preparadas, {x['enviado']} enviadas y {x['recepcion']} recibidas: "
                  f"{x['ejecutado']-x['enviado']} por enviar y {x['enviado']-x['recepcion']} en laboratorio.",
    })
    x = lab["Check Assay"]
    a.append({
        "sev": "crit",
        "title": "Check Assay retrasado",
        "detail": f"{x['ejecutado']} preparados, {x['enviado']} enviados y {x['recepcion']} recibidos: "
                  f"{x['ejecutado']-x['enviado']} por enviar.",
    })
    x = ctl["qaqc"]
    a.append({
        "sev": "warn",
        "title": "QAQC sin envío registrado",
        "detail": f"{x['ejecutado']} lotes QAQC preparados; el Excel no registra envíos ni recepciones.",
    })
    x = lab["Lotes"]
    a.append({
        "sev": "warn",
        "title": "Lotes en laboratorio sin recepción",
        "detail": f"{x['ejecutado']} preparados, {x['enviado']} enviados y {x['recepcion']} recibidos: "
                  f"{x['ejecutado']-x['enviado']} por enviar y {x['enviado']-x['recepcion']} esperando resultados.",
    })
    x = lab["Densidad"]
    if x["ejecutado"] > x["enviado"]:
        a.append({
            "sev": "info",
            "title": "Densidad con muestras por enviar",
            "detail": f"{x['ejecutado']} preparadas y {x['enviado']} enviadas; lo enviado ya fue recibido en su totalidad.",
        })

    geo = comp[comp["GEOLOGICO"] < comp["EJECUTADO"] - 1]
    if len(geo):
        ids = ", ".join(str(i) for i in geo["DDHID"].tolist())
        a.append({
            "sev": "warn",
            "title": "Logueo geológico incompleto en DDH completados",
            "detail": f"{len(geo)} DDH con {num((geo['EJECUTADO']-geo['GEOLOGICO']).sum())} m sin logueo geológico: {ids}.",
        })

    short = comp[comp["PCT"] < 100]
    if len(short):
        a.append({
            "sev": "info",
            "title": "DDH completados bajo el metraje programado",
            "detail": f"{len(short)} DDH cerrados con menos metros que los programados "
                      f"(déficit total {num((short['PROGRAMADO']-short['EJECUTADO']).sum())} m). Revisar si son cierres por objetivo o pérdidas.",
        })

    over = df[df["PCT"] > 105]
    for _, r in over.iterrows():
        a.append({
            "sev": "info",
            "title": "Metraje muy por encima de lo programado",
            "detail": f"{r['DDHID']} ({r['PLATAFORMA']}): {num(r['EJECUTADO'])} m ejecutados vs {num(r['PROGRAMADO'])} m programados ({int(r['PCT'])}%). Confirmar el dato.",
        })

    mismatch = comp[comp["RECEPCION"] < comp["EJECUTADO"] * 0.5]
    for _, r in mismatch.iterrows():
        a.append({
            "sev": "info",
            "title": "Estado «Completado» inconsistente",
            "detail": f"{r['DDHID']}: {num(r['EJECUTADO'])} m perforados y solo {num(r['RECEPCION'])} m recibidos.",
        })
    return a


def leyes_payload(path):
    xl = pd.ExcelFile(path)
    if "Leyes" not in xl.sheet_names:
        return None
    ly = xl.parse("Leyes")
    ly.columns = [str(c).strip() for c in ly.columns]
    need = {"DDHID", "Desde", "Hasta"}
    if not need.issubset(ly.columns):
        return None
    els = [c for c in ly.columns if c not in need and pd.api.types.is_numeric_dtype(ly[c])]
    ly["L"] = (pd.to_numeric(ly["Hasta"], errors="coerce") - pd.to_numeric(ly["Desde"], errors="coerce")).clip(lower=0)
    stats, holes = {}, {}
    for el in els:
        s = ly[["DDHID", "L", el]].dropna()
        s = s[s["L"] > 0]
        if s.empty:
            continue
        w = np.average(s[el], weights=s["L"])
        stats[el] = {"n": int(len(s)), "mean": num(w, 3), "p50": num(s[el].median(), 3), "max": num(s[el].max(), 3),
                     "m": num(s["L"].sum())}
        g = s.groupby("DDHID").apply(lambda t: pd.Series({"grade": np.average(t[el], weights=t["L"]), "len": t["L"].sum()}),
                                     include_groups=False).reset_index()
        g = g.sort_values("grade", ascending=False)
        holes[el] = [{"id": str(r["DDHID"]), "grade": num(r["grade"], 3), "len": num(r["len"])} for _, r in g.iterrows()]
    return {"elements": list(stats), "stats": stats, "holes": holes} if stats else None


def build(path):
    df = load(path)
    total = int(df["Programa"].sum()) if "Programa" in df.columns else len(df)

    estado = df["ESTADO"].value_counts()
    prog_tot, ejec_tot = df["PROGRAMADO"].sum(), df["EJECUTADO"].sum()
    iniciados = df[df["ESTADO"].isin(["Completado", "En Proceso"])]

    kpi = {
        "ddh": {"total": total, "completado": int(estado.get("Completado", 0)),
                "en_proceso": int(estado.get("En Proceso", 0)), "pendiente": int(estado.get("Pendiente", 0))},
        "m": {"prog": num(prog_tot), "ejec": num(ejec_tot), "pend": num(df.loc[df["ESTADO"] == "Pendiente", "PROGRAMADO"].sum()),
              "prog_iniciados": num(iniciados["PROGRAMADO"].sum())},
    }

    eq = df[df["EQUIPO"].notna()].groupby("EQUIPO").agg(ddh=("DDHID", "size"), prog=("PROGRAMADO", "sum"), ejec=("EJECUTADO", "sum"))
    equipos = [{"equipo": k, "ddh": int(r.ddh), "prog": num(r.prog), "ejec": num(r.ejec)}
               for k, r in eq.sort_values("ejec", ascending=False).iterrows()]

    etapas = [{"k": c, "label": l, "m": num(df[c].sum())} for c, l in zip(STAGE_COLS, STAGE_LABELS)]

    c = lambda col: count(df, col)
    ctl = {
        "validacion": [
            {"item": "Collar", "meta": total, "kpdi": c("Collar_KPDI"), "mcp": c("Collar_MCP")},
            {"item": "Survey", "meta": total, "kpdi": c("Survey_KPDI"), "mcp": c("Survey_MCP")},
            {"item": "Assay", "meta": total, "kpdi": c("Assay_KPDI"), "mcp": c("Assay_MCP")},
        ],
        "corte_muestreo": [
            {"item": "Corte", "meta": total, "disponible": c("DDH_Logueado"), "ejecutado": c("Corte Ejecutado")},
            {"item": "Muestreo", "meta": total, "disponible": c("DDH_Logueado"), "ejecutado": c("Muestreo Ejecutado")},
        ],
        "laboratorio": [
            {"item": "XRD", "meta": total, "ejecutado": c("XRD_Preparado"), "enviado": c("XRD_Enviado"), "recepcion": c("XRD_Recepción")},
            {"item": "Lotes", "meta": total, "ejecutado": c("Lotes_Preparados"), "enviado": c("Lotes_Enviados"), "recepcion": c("Lotes_Recepción")},
            {"item": "Densidad", "meta": total, "ejecutado": c("Densidad_Preparados"), "enviado": c("Densidad_Enviada"), "recepcion": c("Densidad_Recepeción")},
            {"item": "Check Assay", "meta": total, "ejecutado": c("Check Assay_Preparado"), "enviado": c("Check Assay_Enviado"), "recepcion": c("Check Assay_Recepción")},
        ],
        "qaqc": {"ejecutado": c("QAQC")},
    }

    return {
        "meta": {"archivo": Path(path).name, "generado": dt.datetime.now().strftime("%d-%m-%Y %H:%M"), "total": total},
        "kpi": kpi, "equipos": equipos, "etapas": etapas, "control": ctl,
        "holes": holes_payload(df), "alertas": alerts(df, ctl, total), "leyes": leyes_payload(path),
    }


def check(path, data):
    """Compara los conteos recalculados con el bloque resumen AL:AM de la hoja Data."""
    ws = openpyxl.load_workbook(path, data_only=True)["Data"]
    df = load(path)
    bad = 0
    print("Validación contra el resumen del Excel (Data!AL:AM):")
    for r in range(1, 60):
        lab, val = ws[f"AL{r}"].value, ws[f"AM{r}"].value
        if lab is None:
            continue
        if lab == "Total":
            mine = data["meta"]["total"]
        elif lab in df.columns:
            mine = count(df, lab)
        else:
            continue
        ok = int(val or 0) == mine
        bad += not ok
        print(f"  {'OK ' if ok else 'ERR'} {lab:<24} excel={val!s:<5} calculado={mine}")
    print("Todo coincide." if not bad else f"{bad} diferencia(s): revisar el Excel.")
    return bad




# ---------------------------------------------------------------- reportes
# Celda 3: reportes (gráficos, Word gerencial y PPT)
import base64
import io
from decimal import ROUND_HALF_UP, Decimal
from datetime import datetime, timedelta, timezone

import matplotlib
import matplotlib.ticker
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt as PPt

# ----------------------------------------------------------------- parámetros
LOGO_B64 = "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"          # logo tomado de tu PPT de ejemplo
TITULO = "REPORTE AVANCES MREDP 2026"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]

# Texto de "Plan de acción" por lámina. Déjalo vacío para que salga "*" como en tu ejemplo,
# o escribe tus acciones, por ejemplo: PLAN_ACCION = {"validacion": ["Cerrar validación MCP de Collar y Survey"]}
PLAN_ACCION = {}               # claves: perforacion, validacion, envios

C = {"recv": "#A7C9EB", "send": "#6E97BD", "done": "#1F5E99", "meta": "#D9D9D9",
     "ink": "#595959", "grid": "#D9D9D9", "blue_hex": "1F5E99"}


def fecha_reporte(fecha=None):
    """Fecha del reporte: la indicada o la de hoy en Lima."""
    if fecha is None:
        return datetime.now(timezone(timedelta(hours=-5))).date()   # Lima: UTC-5, sin horario de verano
    return fecha if not isinstance(fecha, str) else datetime.strptime(fecha, "%d/%m/%Y").date()


def fecha_larga(f):
    return f"{f.day} de {MESES[f.month-1]} de {f.year}"


def logo_bytes():
    return base64.b64decode(LOGO_B64)


# ------------------------------------------------------------------- gráficos
def _num(n, d=0):
    q = Decimal(1).scaleb(-d)
    return f"{Decimal(str(n)).quantize(q, rounding=ROUND_HALF_UP):,.{d}f}"


def _pct(n, t, d=0):
    return f"{100*n/t:.{d}f}%" if t else "0%"


def grafico_barras(cats, series, titulo, xmax, size=(6.2, 4.4), tick=None, leyenda_cols=4, resumen=None):
    """Barras horizontales agrupadas al estilo de los gráficos de Excel de control.
    cats: etiquetas (la primera queda arriba). series: [(nombre, color, valores, etiquetas)].
    resumen: (etiqueta_ejecutado, valor, etiqueta_programado, valor, texto_ejecutado, texto_programado)
    -> bloque de dos barras justo después de la última categoría, con escala propia."""
    cats, series = list(cats)[::-1], [(n, c, list(v)[::-1], list(l)[::-1]) for n, c, v, l in series]
    n = len(series)
    h = 0.8 / n
    if resumen:
        fig, (ax, axp) = plt.subplots(2, 1, figsize=size, dpi=200,
                                      gridspec_kw={"height_ratios": [len(cats) * 1.0, 0.7], "hspace": 0.16})
    else:
        fig, ax = plt.subplots(figsize=size, dpi=200)
    for k, (nombre, color, vals, labs) in enumerate(series):
        ys = [i + 0.4 - h * (k + 0.5) for i in range(len(cats))]
        ax.barh(ys, vals, height=h * 0.98, color=color, label=nombre, zorder=3)
        for y, v, t in zip(ys, vals, labs):
            ax.text(v + xmax * 0.012, y, t, va="center", ha="left", fontsize=7.5, color=C["ink"], zorder=4,
                    bbox=dict(boxstyle="square,pad=0.2", fc="white", ec="#BFBFBF", lw=0.6))
    ax.set_yticks(range(len(cats)))
    ax.set_yticklabels(cats, fontsize=8.5, color=C["ink"])
    ax.set_xlim(0, xmax)
    if tick:
        ax.set_xticks(range(0, int(xmax) + 1, tick))
    else:
        ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.tick_params(axis="x", labelsize=7.5, colors=C["ink"], length=0)
    ax.tick_params(axis="y", length=0)
    if resumen:
        ax.xaxis.tick_top()
    ax.grid(axis="x", color=C["grid"], lw=0.6, zorder=0)
    for s in ax.spines.values():
        s.set_visible(False)
    ax.spines["left"].set_visible(True)
    ax.spines["left"].set_color("#BFBFBF")
    ax.set_ylim(-0.5, len(cats) - 0.5)
    ax.set_title(titulo, fontsize=9.5, color=C["ink"], pad=20 if resumen else 8)
    handles, labels = ax.get_legend_handles_labels()
    leg = dict(ncol=leyenda_cols, frameon=False, fontsize=7.5, labelcolor=C["ink"], handlelength=0.9,
               handleheight=0.9, columnspacing=1.2)
    if resumen:
        le, ve, lp, vp, te, tp = resumen
        axp.barh([1, 0], [ve, vp], height=0.78, color=[C["done"], C["meta"]], zorder=3)
        xm = vp * 1.6
        for y, v, t in ((1, ve, te), (0, vp, tp)):
            axp.text(v + xm * 0.012, y, t, va="center", ha="left", fontsize=7.5, color=C["ink"], zorder=4,
                     bbox=dict(boxstyle="square,pad=0.2", fc="white", ec="#BFBFBF", lw=0.6))
        axp.set_xlim(0, xm)
        axp.set_title("TOTAL DEL PROGRAMA", loc="left", fontsize=7, color="#7F7F7F", pad=3)
        axp.spines["top"].set_visible(True)
        axp.spines["top"].set_color("#BFBFBF")
        axp.set_ylim(-0.5, 1.5)
        axp.set_yticks([1, 0])
        axp.set_yticklabels([f"{le} ({ve})", f"{lp} ({vp})"], fontsize=8.5, color=C["ink"])
        axp.set_xticks([])
        axp.tick_params(axis="y", length=0)
        for k_, s in axp.spines.items():
            if k_ != "top":
                s.set_visible(False)
        fig.legend(handles, labels, loc="lower center", bbox_to_anchor=(0.5, 0.01), **leg)
        fig.subplots_adjust(left=0.27, right=0.96, top=0.88, bottom=0.09)
    else:
        ax.legend(handles, labels, loc="upper center", bbox_to_anchor=(0.5, -0.07), **leg)
        fig.tight_layout()
    fig.patch.set_edgecolor("#E3E3E3")
    fig.patch.set_linewidth(1)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


def hacer_graficos(D):
    CT, T = D["control"], D["meta"]["total"]
    g = {}
    v = CT["validacion"][::-1]          # Assay arriba, como en el ejemplo
    g["validacion"] = grafico_barras(
        [x["item"] for x in v],
        [("Validado_MCP", C["send"], [x["mcp"] for x in v], [_pct(x["mcp"], x["meta"], 1) for x in v]),
         ("Validado_KPDI", C["done"], [x["kpdi"] for x in v], [_pct(x["kpdi"], x["meta"], 1) for x in v]),
         ("Programado (Meta)", C["meta"], [x["meta"] for x in v], [str(x["meta"]) for x in v])],
        "Control de Avance de Datos y Muestras (KPI vs Programado)", 150, tick=20, leyenda_cols=3)
    cm = CT["corte_muestreo"][::-1]          # Muestreo arriba, como en el ejemplo
    g["corte"] = grafico_barras(
        [x["item"] for x in cm],
        [("Ejecutado", C["send"], [x["ejecutado"] for x in cm], [_pct(x["ejecutado"], x["meta"], 1) for x in cm]),
         ("Disponible", C["done"], [x["disponible"] for x in cm], [_pct(x["disponible"], x["meta"], 1) for x in cm]),
         ("Programado (Meta)", C["meta"], [x["meta"] for x in cm], [str(x["meta"]) for x in cm])],
        "Control de Avance de Datos y Muestras (KPI vs Programado)", 150, tick=20, leyenda_cols=3)
    lb = CT["laboratorio"][::-1]             # Check Assay arriba, como en el ejemplo
    g["lab"] = grafico_barras(
        [x["item"] for x in lb],
        [("Recepción", C["recv"], [x["recepcion"] for x in lb], [_pct(x["recepcion"], x["meta"]) for x in lb]),
         ("Enviado", C["send"], [x["enviado"] for x in lb], [_pct(x["enviado"], x["meta"]) for x in lb]),
         ("Ejecutado /Disponible", C["done"], [x["ejecutado"] for x in lb], [_pct(x["ejecutado"], x["meta"]) for x in lb]),
         ("Programado (Meta)", C["meta"], [x["meta"] for x in lb], [str(x["meta"]) for x in lb])],
        "Control de Avance de Datos y Muestras (KPI vs Programado)", 150, size=(6.8, 4.9), tick=20, leyenda_cols=4)
    eq = D["equipos"]
    mx = max(e["prog"] for e in eq) * 1.35
    g["equipos"] = grafico_barras(
        [f"{e['equipo']} ({e['ddh']})" for e in eq],
        [("Ejecutado", C["done"], [e["ejec"] for e in eq], [f"{_num(e['ejec'])} m · {_pct(e['ejec'], e['prog'])}" for e in eq]),
         ("Programado", C["meta"], [e["prog"] for e in eq], [f"{_num(e['prog'])} m" for e in eq])],
        "Metros por equipo: ejecutado vs programado", mx, size=(6.2, 5.2), leyenda_cols=2,
        resumen=("Ejecutado", D["kpi"]["ddh"]["completado"], "Programado", D["kpi"]["ddh"]["total"],
                 f"{_pct(D['kpi']['ddh']['completado'], D['kpi']['ddh']['total'], 1)} DDH · {_num(D['kpi']['m']['ejec'])} m ({_pct(D['kpi']['m']['ejec'], D['kpi']['m']['prog'], 1)})",
                 f"{_num(D['kpi']['m']['prog'])} m"))
    et = D["etapas"]
    base = et[0]["m"]
    g["etapas"] = grafico_barras(
        [e["label"] for e in et],
        [("Metros", C["done"], [e["m"] for e in et], [f"{_num(e['m'])} m · {_pct(e['m'], base, 1)}" for e in et])],
        "Metros que superaron cada etapa", base * 1.42, leyenda_cols=1)
    return g


# ------------------------------------------------------------------- textos
def textos(D):
    N, M, CT = D["kpi"]["ddh"], D["kpi"]["m"], D["control"]
    v = {x["item"]: x for x in CT["validacion"]}
    cm = {x["item"]: x for x in CT["corte_muestreo"]}
    lb = {x["item"]: x for x in CT["laboratorio"]}
    T = D["meta"]["total"]
    base, geo = D["etapas"][0]["m"], D["etapas"][-1]["m"]
    t = {}
    t["perforacion"] = [
        f"{N['completado']} de {N['total']} DDH completados ({_pct(N['completado'], N['total'], 1)}); {N['en_proceso']} en proceso y {N['pendiente']} pendientes.",
        f"{_num(M['ejec'])} m perforados de {_num(M['prog'])} m programados ({_pct(M['ejec'], M['prog'], 1)}); cumplimiento de {_pct(M['ejec'], M['prog_iniciados'], 1)} en los DDH iniciados.",
        f"Pendiente por perforar: {N['pendiente']} sondajes y {_num(M['pend'])} m ({_pct(M['pend'], M['prog'], 1)} del metraje global programado); sin equipo asignado.",
        f"{_num(base - geo)} m perforados sin logueo geológico ({_pct(base - geo, base, 1)}).",
    ]
    t["validacion"] = [
        f"KPDI: Collar {_pct(v['Collar']['kpdi'], T, 1)}, Survey {_pct(v['Survey']['kpdi'], T, 1)}, Assay {_pct(v['Assay']['kpdi'], T, 1)}. MCP sin validaciones ({v['Collar']['mcp']} / {v['Survey']['mcp']} / {v['Assay']['mcp']}).",
        f"Corte ejecutado en {cm['Corte']['ejecutado']} DDH ({_pct(cm['Corte']['ejecutado'], T, 1)}) y muestreo en {cm['Muestreo']['ejecutado']} ({_pct(cm['Muestreo']['ejecutado'], T, 1)}), sobre {cm['Corte']['disponible']} DDH disponibles.",
    ]
    t["envios"] = [
        f"XRD: {lb['XRD']['ejecutado']} preparadas, {lb['XRD']['enviado']} enviadas, {lb['XRD']['recepcion']} recibidas.",
        f"Lotes: {lb['Lotes']['ejecutado']} / {lb['Lotes']['enviado']} / {lb['Lotes']['recepcion']} (preparados / enviados / recibidos).",
        f"Densidad: {lb['Densidad']['ejecutado']} / {lb['Densidad']['enviado']} / {lb['Densidad']['recepcion']}. Check Assay: {lb['Check Assay']['ejecutado']} / {lb['Check Assay']['enviado']} / {lb['Check Assay']['recepcion']}.",
    ]
    return t


# ----------------------------------------------------------------------- PPT
def _tb(slide, x, y, w, h, text, size=14, italic=False, bold=False, color="000000", align=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.size, r.font.italic, r.font.bold, r.font.name = PPt(size), italic, bold, "Calibri"
    r.font.color.rgb = PRGB.from_string(color)
    if align:
        p.alignment = align
    return box


def _bloque(slide, x, y, w, titulo, lineas, size=11):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(1.2))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = titulo
    r.font.size, r.font.italic, r.font.name = PPt(14), True, "Calibri"
    for ln in (lineas or ["*"]):
        q = tf.add_paragraph()
        q.space_before = PPt(3)
        rr = q.add_run()
        rr.text = ln if ln == "*" else "• " + ln
        rr.font.size, rr.font.name = PPt(size), "Calibri"
        rr.font.color.rgb = PRGB.from_string("404040")


def _marco(prs, titulo):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(io.BytesIO(logo_bytes()), Inches(12.15), Inches(0), Inches(1.2), Inches(0.82))
    s.shapes.add_picture(io.BytesIO(logo_bytes()), Inches(0), Inches(6.29), Inches(1.75), Inches(1.2))
    ln = s.shapes.add_connector(1, Inches(0), Inches(1.02), Inches(13.22), Inches(1.02))
    ln.line.color.rgb = PRGB.from_string("7F7F7F")
    ln.line.width = PPt(0.75)
    _tb(s, 0.2, 0.34, 11, 0.5, titulo, size=24, italic=True)
    return s


def _img(slide, buf, x, y, w=None, h=None):
    buf.seek(0)
    slide.shapes.add_picture(buf, Inches(x), Inches(y), Inches(w) if w else None, Inches(h) if h else None)


def hacer_ppt(D, graf, f, ruta):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    tx = textos(D)
    plan = lambda k: PLAN_ACCION.get(k)

    # 1. Portada
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(io.BytesIO(logo_bytes()), Inches(12.15), Inches(0), Inches(1.2), Inches(0.82))
    s.shapes.add_picture(io.BytesIO(logo_bytes()), Inches(0), Inches(6.29), Inches(1.75), Inches(1.2))
    _tb(s, 1.56, 1.5, 10.0, 2.61, TITULO, size=44, italic=True, align=PP_ALIGN.CENTER)
    _tb(s, 1.56, 3.94, 10.0, 0.6, "Perforación diamantina KPDI · Control de datos y muestras", size=18, italic=True, color="595959", align=PP_ALIGN.CENTER)
    _tb(s, 9.59, 6.55, 3.0, 0.4, f.strftime("%d/%m/%Y"), size=14, color="595959", align=PP_ALIGN.RIGHT)

    # 2. Avance de perforación
    s = _marco(prs, "Avance de perforación")
    _img(s, graf["equipos"], 0.27, 1.3, h=4.0)
    _img(s, graf["etapas"], 6.9, 1.3, h=4.0)
    _bloque(s, 2.0, 5.45, 5.4, "Comentarios", tx["perforacion"], 10.5)
    _bloque(s, 8.2, 5.45, 4.6, "Plan de acción", plan("perforacion"), 10.5)

    # 3. Validación, corte y muestreo
    s = _marco(prs, "Trazabilidad validación , corte y muestreo")
    _img(s, graf["validacion"], 0.27, 1.3, h=4.0)
    _img(s, graf["corte"], 6.9, 1.3, h=4.0)
    _bloque(s, 2.0, 5.45, 5.4, "Comentarios", tx["validacion"], 10.5)
    _bloque(s, 8.2, 5.45, 4.6, "Plan de acción", plan("validacion"), 10.5)

    # 4. Envíos y resultados
    s = _marco(prs, "Trazabilidad Envíos , resultados")
    _img(s, graf["lab"], 1.0, 1.3, h=4.9)
    _bloque(s, 8.9, 1.6, 4.1, "Comentarios", tx["envios"], 11)
    _bloque(s, 8.9, 4.6, 4.1, "Plan de acción", plan("envios"), 11)

    prs.save(ruta)
    return ruta


# ---------------------------------------------------------------------- Word
def _shade(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), hex_)
    tcPr.append(sh)


def _borders(table, color="D9D9D9"):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{e}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), color)
        b.append(el)
    tblPr.append(b)


def _campo_pagina(par):
    r = par.add_run()
    for tipo, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if tipo:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tipo)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = txt
        r._r.append(el)


def _h(doc, texto):
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(12), Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(texto)
    r.bold, r.font.size = True, Pt(13)
    r.font.color.rgb = RGBColor.from_string(C["blue_hex"])
    pPr = p._p.get_or_add_pPr()
    bd = OxmlElement("w:pBdr")
    bt = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", "6"), ("space", "1"), ("color", "BFBFBF")):
        bt.set(qn(f"w:{k}"), v)
    bd.append(bt)
    pPr.append(bd)


def _bullets(doc, items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(2)
        p.add_run(it).font.size = Pt(10)


def _fig(doc, buf, ancho_cm):
    buf.seek(0)
    doc.add_picture(buf, width=Cm(ancho_cm))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER


def hacer_word(D, graf, f, ruta):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2)
    sec.top_margin, sec.bottom_margin = Cm(2.2), Cm(1.8)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Calibri", Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")

    # encabezado y pie
    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp.add_run().add_picture(io.BytesIO(logo_bytes()), height=Cm(1.1))
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run(f"Reporte gerencial MREDP 2026 · {f.strftime('%d/%m/%Y')} · Página ")
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
    _campo_pagina(fp)
    for rn in fp.runs:
        rn.font.size = Pt(8)

    # título
    p = doc.add_paragraph()
    r = p.add_run("Reporte gerencial de avances")
    r.bold, r.font.size = True, Pt(22)
    r.font.color.rgb = RGBColor.from_string(C["blue_hex"])
    p.paragraph_format.space_after = Pt(0)
    p = doc.add_paragraph()
    r = p.add_run(f"Perforación diamantina KPDI · Programa MREDP 2026\nFecha del reporte: {fecha_larga(f)}")
    r.font.size, r.font.color.rgb = Pt(10.5), RGBColor(0x59, 0x59, 0x59)

    N, M, CT = D["kpi"]["ddh"], D["kpi"]["m"], D["control"]
    tx = textos(D)

    # KPI
    _h(doc, "Resumen ejecutivo")
    kp = [("DDH completados", f"{N['completado']} / {N['total']}", _pct(N["completado"], N["total"], 1)),
          ("Metros perforados", f"{_num(M['ejec'])} m", f"{_pct(M['ejec'], M['prog'], 1)} de {_num(M['prog'])} m"),
          ("Cumplimiento DDH iniciados", _pct(M["ejec"], M["prog_iniciados"], 1), f"sobre {_num(M['prog_iniciados'])} m"),
          ("Metros por perforar", f"{_num(M['pend'])} m", f"{_pct(M['pend'], M['prog'], 1)} del metraje global · {N['pendiente']} DDH")]
    t = doc.add_table(rows=1, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c, (a, b, s) in zip(t.rows[0].cells, kp):
        _shade(c, "EAF0F6")
        c.paragraphs[0].add_run(a.upper()).font.size = Pt(7.5)
        q = c.add_paragraph()
        rr = q.add_run(b)
        rr.bold, rr.font.size = True, Pt(15)
        rr.font.color.rgb = RGBColor.from_string(C["blue_hex"])
        q2 = c.add_paragraph()
        q2.add_run(s).font.size = Pt(8)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    crit = [a for a in D["alertas"] if a["sev"] == "crit"]
    _bullets(doc, tx["perforacion"][:2] + [tx["perforacion"][2]] + [f"{len(crit)} alertas críticas: " + "; ".join(a["title"] for a in crit) + "."])

    # perforación
    _h(doc, "Avance de perforación")
    t = doc.add_table(rows=1, cols=2)
    for c, k in zip(t.rows[0].cells, ("equipos", "etapas")):
        graf[k].seek(0)
        c.paragraphs[0].add_run().add_picture(graf[k], width=Cm(8.4))
    _bullets(doc, tx["perforacion"][2:])

    # datos y muestras
    _h(doc, "Control de datos y muestras")
    t = doc.add_table(rows=1, cols=2)
    for c, k in zip(t.rows[0].cells, ("validacion", "corte")):
        graf[k].seek(0)
        c.paragraphs[0].add_run().add_picture(graf[k], width=Cm(8.4))
    _bullets(doc, tx["validacion"])
    _fig(doc, graf["lab"], 12.5)
    _bullets(doc, tx["envios"])

    # alertas
    _h(doc, "Alertas y cuellos de botella")
    col = {"crit": ("B3372F", "Crítica"), "warn": ("A8690F", "Atención"), "info": ("53626F", "Revisar")}
    orden = {"crit": 0, "warn": 1, "info": 2}
    al = sorted(D["alertas"], key=lambda a: orden[a["sev"]])
    t = doc.add_table(rows=1, cols=3)
    _borders(t)
    for c, h in zip(t.rows[0].cells, ("Nivel", "Tema", "Detalle")):
        _shade(c, "1F5E99")
        rr = c.paragraphs[0].add_run(h)
        rr.bold, rr.font.size, rr.font.color.rgb = True, Pt(9), RGBColor(255, 255, 255)
    for a in al:
        cells = t.add_row().cells
        rr = cells[0].paragraphs[0].add_run(col[a["sev"]][1])
        rr.bold, rr.font.size, rr.font.color.rgb = True, Pt(8.5), RGBColor.from_string(col[a["sev"]][0])
        cells[1].paragraphs[0].add_run(a["title"]).font.size = Pt(8.5)
        cells[2].paragraphs[0].add_run(a["detail"]).font.size = Pt(8.5)
    t.autofit = False
    for col_, w in zip(t.columns, (Cm(1.8), Cm(4.6), Cm(10.6))):
        col_.width = w
    for row in t.rows:
        for c, w in zip(row.cells, (Cm(1.8), Cm(4.6), Cm(10.6))):
            c.width = w

    # leyes
    _h(doc, "Resultados de leyes")
    L = D.get("leyes")
    if not L:
        doc.add_paragraph(f"El archivo de control no incluye leyes. {CT['validacion'][2]['kpdi']} de {D['meta']['total']} DDH "
                          f"({_pct(CT['validacion'][2]['kpdi'], D['meta']['total'], 1)}) tienen assay validado por KPDI.").runs[0].font.size = Pt(10)
    else:
        t = doc.add_table(rows=1, cols=6)
        _borders(t)
        for c, h in zip(t.rows[0].cells, ("Elemento", "Muestras", "Metros", "Ley media", "Mediana", "Máxima")):
            _shade(c, "1F5E99")
            rr = c.paragraphs[0].add_run(h)
            rr.bold, rr.font.size, rr.font.color.rgb = True, Pt(9), RGBColor(255, 255, 255)
        for el in L["elements"]:
            s = L["stats"][el]
            cells = t.add_row().cells
            for c, v in zip(cells, (el, _num(s["n"]), _num(s["m"]), f"{s['mean']:.2f}", f"{s['p50']:.2f}", f"{s['max']:.2f}")):
                c.paragraphs[0].add_run(v).font.size = Pt(9)
    doc.save(ruta)
    return ruta


# ------------------------------------------------------------- orquestador
def generar_todo(excel, fecha=None, carpeta=".", formatos=("html", "docx", "pptx")):
    """Genera dashboard HTML, reporte Word y PPT (o solo los formatos pedidos).
    La fecha del reporte va en los nombres de archivo y en el contenido. Devuelve {formato: ruta}."""
    f = fecha_reporte(fecha)
    carpeta = Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    data = build(excel)
    check(excel, data)
    iso = f.strftime("%Y-%m-%d")
    rutas = {
        "html": carpeta / f"Dashboard_DDH_MREDP_{iso}.html",
        "docx": carpeta / f"Reporte_Gerencial_MREDP_{iso}.docx",
        "pptx": carpeta / f"Reporte_Avances_MREDP_{iso}.pptx",
    }
    graf = hacer_graficos(data) if ("docx" in formatos or "pptx" in formatos) else None
    if "html" in formatos:
        payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
        rutas["html"].write_text(TEMPLATE.replace("/*__DATA__*/null", payload), encoding="utf-8")
    if "docx" in formatos:
        hacer_word(data, graf, f, rutas["docx"])
    if "pptx" in formatos:
        hacer_ppt(data, graf, f, rutas["pptx"])
    salida = {k: rutas[k] for k in formatos}
    for k, v in salida.items():
        print(f"{k:>5}: {v}")
    return salida
