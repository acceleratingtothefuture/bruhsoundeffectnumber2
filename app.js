/* ===== YEARS IN FOLDER ===== */
const YEARS = [2024, 2023];   // add 2025 next January

/***** CONSTANTS *****/
const COLORS = [
  '#000', '#e91e63', '#ff9800', '#ffeb3b', '#4caf50',
  '#00bcd4', '#9c27b0', '#f44336', '#3f51b5', '#2196f3', '#795548'
];

const STATUS_TYPES = ['Filed', 'Dismissed', 'Rejected', 'Open', 'Sentenced'];

const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug',
                     'Sep','Oct','Nov','Dec'];

/***** HOVER BAR PLUGIN *****/
const hoverBar = {
  id: 'hoverBar',
  afterDraw(c) {
    if (c.config.type !== 'line') return;
    const { ctx, tooltip, chartArea } = c;
    if (!tooltip._active?.length) return;
    const x = tooltip._active[0].element.x;
    ctx.save();
    ctx.fillStyle = 'rgba(0,0,0,.07)';
    ctx.fillRect(x - 18, chartArea.top, 36, chartArea.bottom - chartArea.top);
    ctx.restore();
  }
};
Chart.register(hoverBar);

/* ===== FILE LOCATION ===== */
// while you’re on Live Server, always use the local ./data/ folder
const FOLDER = './data/';                         // ← change this line

// later, when the files live in WordPress, swap it to:
// const FOLDER = '/wp-content/uploads/da-dashboard/';


let rows = [], charts = [], pieChart = null;

loadData(YEARS).then(() => {
  initDimension();
  build();
  initLargeChart();
});


/* find every cases_YYYY.xlsx that exists, newest first */
async function discoverYears() {
  const found = [];
  const thisYear = new Date().getFullYear();
  for (let y = thisYear; y >= 2015; y--) {
    const head = await fetch(`${FOLDER}cases_${y}.xlsx`, { method: 'HEAD' });
    if (head.ok) found.push(y);
    else if (found.length) break;               // stop at first gap
  }
  return found;
}

/* read both xlsx files per year, merge defendant info into each case row */
async function loadData(YEARS) {
  for (const y of YEARS) {
    const [bufCases, bufDefs] = await Promise.all([
      fetch(`${FOLDER}cases_${y}.xlsx`).then(r => r.arrayBuffer()),
      fetch(`${FOLDER}defendants_${y}.xlsx`).then(r => r.arrayBuffer())
    ]);

    const wbCases = XLSX.read(bufCases, { type: 'array' });
    const wbDefs  = XLSX.read(bufDefs,  { type: 'array' });

    const cases = XLSX.utils.sheet_to_json(
      wbCases.Sheets[wbCases.SheetNames[0]], { defval: '' }
    );
    const defs  = XLSX.utils.sheet_to_json(
      wbDefs.Sheets[wbDefs.SheetNames[0]], { defval: '' }
    );

    const byCase = {};
    defs.forEach(d => {
      const id = +d['Case ID'];
      if (!byCase[id]) byCase[id] = d;          // first match only
    });

    cases.forEach(c => {
      const id  = +c['Case ID'];
      const d   = byCase[id] || {};
      const r   = {
        case_id: id,
        date_da: c['Case Received By DA'],
        severity: c['Severity'],
        agency: c['Arresting Agency'],
        city: c['Location City'],
        status: c['Status'],
        sub_type: c['Case Sub Type'],
        days_to_file: +c['Days to file requested charges'] || 0,
        days_file_to_sent: +c['Days from charges filed to sentencing'] || 0,
        victim_case: c['Victim Case'] === 'TRUE',

        ethnicity: d['Ethnicity'] || '',
        gender: d['Gender'] || '',
        county_res: d['County Resident'] || '',
        age: +d['Defendant Age'] || null
      };

      let dt;
if (r.date_da instanceof Date) {          // SheetJS already parsed it
  dt = r.date_da;
} else {                                  // fallback to MM/DD/YYYY string
  const [mm, dd, yyyy] = String(r.date_da).split('/');
  dt = new Date(+yyyy, +mm - 1, +dd);     // month is 0-based
}
if (isNaN(dt)) return;                    // skip bad rows
               // skip garbage rows

r.ts       = dt.getTime();                // unix ms for quick max()
r.year     = dt.getFullYear();
r.month    = dt.getMonth() + 1;
r.quarter  = Math.floor(dt.getMonth() / 3) + 1;
r.age_group = r.age ? Math.floor(r.age / 10) * 10 + 's' : 'Unknown';

rows.push(r);

    });
  }
}

/***** CONTROLS *****/
['metric', 'range', 'dimension'].forEach(id =>
  document.getElementById(id).onchange = build
);
document.getElementById('pieToggle').onchange = build;

function initDimension() {
  const sel = document.getElementById('dimension');
  const ignore = ['case_id', 'date_da', 'year', 'month', 'quarter', `ts`];
  sel.innerHTML = Object.keys(rows[0])
    .filter(k => !ignore.includes(k))
    .map(k =>
      `<option value="${k}">${k.replace(/_/g,' ').replace(/\b\w/g,c=>c.toUpperCase())}</option>`
    ).join('');
}

/***** HELPERS *****/
const keyOf = (y,m,mode) =>
  mode === 'monthly'   ? `${y}-${m}`       :
  mode === 'quarterly' ? `${y}-Q${Math.ceil(m/3)}` :
  mode === 'annual'    ? String(y)          :
                         `${y}-${m}`;

const fmt = (v,isCount) => (v==null||Number.isNaN(v)) ? 'N/A'
                                                      : v + (isCount?' cases':'%');

function fadeColor(hex,a=.18){
  const n=parseInt(hex.slice(1),16);
  const r=(n>>16)&255,g=(n>>8)&255,b=n&255;
  return `rgba(${r},${g},${b},${a})`;
}

/***** BUILD DASHBOARD *****/
function build() {
  if (largeChart) {
    largeChart.data.datasets = [];
    largeChart.data.labels   = [];
    largeChart.update();
    document.getElementById('compareSection').style.display = 'none';
  }

  alasql('DROP TABLE IF EXISTS cases');
  alasql('CREATE TABLE cases');
  alasql('INSERT INTO cases SELECT * FROM ?', [rows]);

  const range     = document.getElementById('range').value;
  const dim       = document.getElementById('dimension').value;
  const metric    = document.getElementById('metric').value;
  const pieMode   = document.getElementById('pieToggle').checked &&
                    (metric === 'all_cases' || STATUS_TYPES.includes(metric));

  /* buckets */
  const buckets = [];
  if (range === 'last12') {
    const maxD = new Date(Math.max(...rows.map(r => r.ts)));

    for (let i=11;i>=0;i--) {
      const d   = new Date(maxD);
      d.setMonth(d.getMonth()-i);
      const y   = d.getFullYear();
      const m   = d.getMonth()+1;
      buckets.push({
        y,m,
        label:`${MONTH_NAMES[m-1]} '${String(y).slice(-2)}`,
        key:`${y}-${m}`
      });
    }
  } else if (range === 'monthly') {
  const years = [...new Set(rows.map(r => r.year))].sort((a, b) => a - b);
  years.forEach(year =>
    MONTH_NAMES.forEach((_, i) =>
      buckets.push({
        y: year,
        m: i + 1,
        label: `${MONTH_NAMES[i]} '${String(year).slice(-2)}`,
        key: `${year}-${i + 1}`
      })
    )
  );

} else if (range === 'quarterly') {
  const years = [...new Set(rows.map(r => r.year))].sort((a, b) => a - b);
  years.forEach(year =>
    [1, 2, 3, 4].forEach(q =>
      buckets.push({
        y: year,
        q,
        label: `Q${q} '${String(year).slice(-2)}`,
        key: `${year}-Q${q}`
      })
    )
  );

} else { /* annual */
  const years = [...new Set(rows.map(r => r.year))].sort((a, b) => a - b);
  years.forEach(year =>
    buckets.push({
      y: year,
      label: String(year),
      key: String(year)
    })
  );
}


  /* aggregates */
  const allCounts = {}, statusCounts = {}, groupAll = {}, groupStatus = {};

  rows.forEach(r=>{
    const key = keyOf(r.year,r.month,range);
    const g   = r[dim] || 'Unknown';

    allCounts[key]=(allCounts[key]||0)+1;
    (groupAll[g]??={})[key]=(groupAll[g][key]||0)+1;

    const s = r.status;
    (statusCounts[s]??={})[key]=(statusCounts[s][key]||0)+1;
    (groupStatus[s]??={});
    (groupStatus[s][g]??={});
    groupStatus[s][g][key]=(groupStatus[s][g][key]||0)+1;
  });

  /* pick the slice to plot */
  let bucketBase, groupBase;
  if (metric === 'all_cases') {
    bucketBase = allCounts;
    groupBase  = groupAll;
  } else {
    bucketBase = statusCounts[metric] || {};
    groupBase  = groupStatus[metric] || {};
  }

  if (pieMode) {
    const lineData = buckets.map(b=>bucketBase[b.key]||0);
    renderLinePie(buckets,lineData,groupBase,metric);
    return;
  }

  const datasets=[
    { label:'ALL', color:'#000',
      values:buckets.map(b=>bucketBase[b.key]||0) },
    ...Object.keys(groupBase).map((g,i)=>({
      label:g,
      color:COLORS[(i+1)%COLORS.length],
      values:buckets.map(b=>groupBase[g]?.[b.key]||0)
    }))
  ];

  render(datasets,buckets.map(b=>b.label),true);
}

/***** RENDER FUNCTIONS (unchanged) *****/
function render(datasets,labels,isCount){
  const grid=document.getElementById('chartGrid');
  grid.innerHTML='';
  charts.forEach(c=>c.destroy());
  charts=[];

  const first=labels[0],last=labels.at(-1);

  datasets.forEach((d,i)=>{
    const id=`c${i}`;
    grid.insertAdjacentHTML('beforeend',`
      <div class="chart-box">
        <div class="chart-head">
          <div class="chart-title">${escapeHtml(d.label)}</div>
          <div class="chart-month" id="m${i}"></div>
        </div>
        <div class="chart-number" id="v${i}">${fmt(d.values.at(-1),isCount)}</div>
        <div class="chart-canvas"><canvas id="${id}" width="280" height="100"></canvas></div>
        <div class="range-labels"><span>${first}</span><span>${last}</span></div>
        <label style="margin-top:8px;display:block;">
          <input type="checkbox" onchange="toggleLargeChart(${i})"> Compare
        </label>
      </div>`);

    const ctx=document.getElementById(id).getContext('2d');
    const chart=new Chart(ctx,{
      type:'line',
      data:{labels,datasets:[{
        label:d.label,data:d.values,
        borderColor:d.color,backgroundColor:d.color,
        tension:.18,pointRadius:0,pointHoverRadius:5
      }]},
      options:{
        responsive:false,animation:false,
        plugins:{legend:{display:false},tooltip:{enabled:false}},
        interaction:{mode:'nearest',axis:'x',intersect:false},
        scales:{x:{display:false},
                y:{beginAtZero:true,ticks:{callback:v=>Number.isInteger(v)?v:''}}},
        onHover:(e,els)=>els.length?hover(els[0].index,labels,isCount):clear(isCount)
      },
      plugins:[hoverBar]
    });
    charts.push(chart);
  });
}

function renderLinePie(buckets,lineData,groupCounts,metricName){
  const grid=document.getElementById('chartGrid');
  grid.innerHTML=`
    <div class="chart-box" style="flex:1 1 100%;">
      <div class="chart-head">
        <div class="chart-title">${escapeHtml(metricName)}</div>
        <div class="chart-month" id="lineMonth"></div>
      </div>
      <div class="chart-number" id="lineValue">${lineData.at(-1)} cases</div>
      <canvas id="lineMain" height="140"></canvas>
    </div>
    <div class="chart-box" style="flex:1 1 320px;">
      <div class="chart-head"><div class="chart-title">Breakdown</div></div>
      <div class="chart-number" id="sliceValue"></div>
      <canvas id="pieMain" height="140"></canvas>
    </div>`;

  const lineCtx=document.getElementById('lineMain').getContext('2d');
  const pieCtx=document.getElementById('pieMain').getContext('2d');
  const labels=buckets.map(b=>b.label);
  let origColors=[];

  new Chart(lineCtx,{
    type:'line',
    data:{
      labels,
      datasets:[{
        label:metricName,
        data:lineData,
        borderColor:'#000',backgroundColor:'#000',
        tension:.18,pointRadius:0,pointHoverRadius:5
      }]
    },
    options:{
      responsive:true,animation:false,
      plugins:{legend:{display:false},tooltip:{enabled:false}},
      interaction:{mode:'nearest',axis:'x',intersect:false},
      scales:{y:{beginAtZero:true}},
      onHover:(e,els)=>{
        if(!els.length) return;
        const idx=els[0].index;
        updatePie(idx);
        document.getElementById('lineValue').textContent=lineData[idx]+' cases';
        document.getElementById('lineMonth').textContent=labels[idx];
      }
    }
  });

  pieChart=new Chart(pieCtx,{
    type:'pie',
    data:{labels:[],datasets:[{data:[],backgroundColor:[]}]},
    options:{
      plugins:{legend:{position:'right'},tooltip:{enabled:false}},
      onHover:(e,els)=>{
        const box=document.getElementById('sliceValue');
        if(!els.length){
          pieChart.data.datasets[0].backgroundColor=origColors;
          pieChart.update();
          box.textContent='';
          box.style.color='#000';
          return;
        }
        const i=els[0].index;
        const lbl=pieChart.data.labels[i];
        const val=pieChart.data.datasets[0].data[i];
        pieChart.data.datasets[0].backgroundColor=
          origColors.map((c,idx)=>idx===i?c:fadeColor(c));
        pieChart.update();
        box.textContent=`${lbl}: ${val} cases`;
        box.style.color=origColors[i];
      }
    }
  });

  function updatePie(idx){
    const key=buckets[idx].key;
    const sliceLabels=[], sliceData=[], sliceColors=[];
    let colorIdx=1;
    Object.keys(groupCounts).forEach(g=>{
      const v=groupCounts[g]?.[key]||0;
      if(!v) return;
      sliceLabels.push(g);
      sliceData.push(v);
      sliceColors.push(COLORS[(colorIdx++)%COLORS.length]);
    });
    origColors=sliceColors.slice();
    pieChart.data.labels=sliceLabels;
    pieChart.data.datasets[0].data=sliceData;
    pieChart.data.datasets[0].backgroundColor=sliceColors;
    pieChart.update();
  }
  updatePie(buckets.length-1);
  document.getElementById('lineMonth').textContent=labels.at(-1);
}

/***** COMPARE CHART *****/
let largeChart=null;
function initLargeChart(){
  const ctx=document.getElementById('largeChart').getContext('2d');
  largeChart=new Chart(ctx,{
    type:'line',
    data:{labels:[],datasets:[]},
    options:{
      responsive:true,
      plugins:{legend:{position:'top'}},
      interaction:{mode:'nearest',axis:'x',intersect:false},
      scales:{y:{beginAtZero:true}}
    }
  });
}
function toggleLargeChart(index){
  const d=charts[index].data.datasets[0];
  const label=d.label;
  const existing=largeChart.data.datasets.find(ds=>ds.label===label);
  if(existing){
    largeChart.data.datasets=largeChart.data.datasets.filter(ds=>ds.label!==label);
  }else{
    largeChart.data.datasets.push({
      label,data:d.data,
      borderColor:d.borderColor,backgroundColor:d.borderColor,
      tension:.18,pointRadius:0,pointHoverRadius:4
    });
    if(!largeChart.data.labels.length){
      largeChart.data.labels=charts[index].data.labels;
    }
  }
  document.getElementById('compareSection').style.display=
    largeChart.data.datasets.length?'block':'none';
  largeChart.update();
  if(!largeChart.data.datasets.length){
    largeChart.data.labels=[];
  }
}

/***** HOVER HELPERS *****/
function hover(i,labels,isCount){
  charts.forEach((c,idx)=>{
    c.setActiveElements([{datasetIndex:0,index:i}]);
    c.update();
    const v=c.data.datasets[0].data[i];
    document.getElementById('v'+idx).textContent=fmt(v,isCount);
    document.getElementById('m'+idx).textContent=labels[i];
  });
}
function clear(isCount){
  charts.forEach((c,idx)=>{
    c.setActiveElements([]);
    c.update();
    const v=c.data.datasets[0].data.at(-1);
    document.getElementById('v'+idx).textContent=fmt(v,isCount);
    document.getElementById('m'+idx).textContent='';
  });
}

/* escape helper to kill XSS */
function escapeHtml(s){
  return String(s)
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;')
    .replaceAll("'",'&#39;');
}
