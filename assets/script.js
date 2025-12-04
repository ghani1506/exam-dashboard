
/* --- Updated script.js with correct Total Candidates logic --- */

let parsedRows = [];
let records = [];
let classChart = null;
let subjectChart = null;

const yearToFile = {
  "7": "assets/year7.xlsx",
  "8": "assets/year8.xlsx",
  "9": "assets/year9.xlsx",
  "10": "assets/year10.xlsx"
};

document.addEventListener('DOMContentLoaded', () => {
  const yearSelect = document.getElementById('yearSelect');
  const metricSelect = document.getElementById('metricSelect');
  const subjectSelect = document.getElementById('subjectSelect');

  yearSelect.addEventListener('change', () => {
    loadYearFile(yearSelect.value);
  });

  metricSelect.addEventListener('change', () => updateCharts());
  subjectSelect.addEventListener('change', () => updateCharts());

  loadYearFile(yearSelect.value);
});

function loadYearFile(yearValue) {
  const filePath = yearToFile[yearValue];
  const kpiWrapper = document.getElementById('kpiWrapper');
  kpiWrapper.innerHTML = '<div class="placeholder">Loading data for Year ' + yearValue + '...</div>';

  fetch(filePath)
    .then(resp => resp.arrayBuffer())
    .then(buf => {
      const data = new Uint8Array(buf);
      const wb = XLSX.read(data, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      parsedRows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
      records = extractYearSheetRecords(parsedRows);

      if (!records.length) {
        kpiWrapper.innerHTML = '<div class="placeholder">File loaded but contains no valid records.</div>';
        clearCharts();
        return;
      }

      initControls(records);
      updateKPIs(records);
      updateCharts();
    })
    .catch(err => {
      console.error(err);
      kpiWrapper.innerHTML = '<div class="placeholder">Failed to load file for Year ' + yearValue + '</div>';
    });
}

function clearCharts() {
  if (classChart) classChart.destroy();
  if (subjectChart) subjectChart.destroy();
}

function extractYearSheetRecords(rows) {
  const out = [];
  if (!rows || !rows.length) return out;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const c0 = (row[0] ?? "").toString().trim();
    const c2 = (row[2] ?? "").toString().trim();

    if (c0 === "Class" && c2 === "Total") {
      let subject = "";
      if (i - 2 >= 0) subject = (rows[i - 2][0] ?? "").toString().trim();

      let j = i + 1;
      while (j < rows.length) {
        const r = rows[j] || [];
        if (r.every(cell => !cell)) break;

        let className = (r[0] ?? r[1] ?? "").toString().trim();
        if (className) {
          out.push({
            subject,
            className,
            total: Number(r[2] || 0),
            Astar: Number(r[3] || 0),
            A2: Number(r[4] || 0),
            B3: Number(r[5] || 0),
            B4: Number(r[6] || 0),
            C5: Number(r[7] || 0),
            C6: Number(r[8] || 0),
            D7: Number(r[9] || 0),
            E8: Number(r[10] || 0),
            U: Number(r[11] || 0),
            total1_6: Number(r[12] || 0),
            pct1_6: Number(r[13] || 0),
            total1_8: Number(r[14] || 0),
            pct1_8: Number(r[15] || 0),
          });
        }
        j++;
      }
    }
  }
  return out;
}

function initControls(records) {
  const subjectSelect = document.getElementById('subjectSelect');
  const metricSelect = document.getElementById('metricSelect');

  subjectSelect.innerHTML = "";
  const subjects = [...new Set(records.map(r => r.subject))];

  subjects.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    subjectSelect.appendChild(opt);
  });

  subjectSelect.disabled = subjects.length === 0;
  metricSelect.disabled = subjects.length === 0;
}

function updateKPIs(records) {
  const kpiWrapper = document.getElementById('kpiWrapper');
  kpiWrapper.innerHTML = "";

  const overallRows = records.filter(r => r.className.toLowerCase() === "overall");

  let totalCandidates = 0;

  if (overallRows.length) {
    const totals = overallRows.map(r => r.total).filter(t => t > 0);
    if (totals.length) {
      const unique = [...new Set(totals)];
      totalCandidates = unique.length === 1 ? unique[0] : Math.max(...unique);
    }
  }

  if (!totalCandidates) {
    totalCandidates = overallRows.reduce((sum, r) => sum + r.total, 0);
  }

  const total1_6 = overallRows.reduce((sum, r) => sum + r.total1_6, 0);
  const total1_8 = overallRows.reduce((sum, r) => sum + r.total1_8, 0);

  const pct1_6 = totalCandidates ? (total1_6 / totalCandidates) * 100 : 0;
  const pct1_8 = totalCandidates ? (total1_8 / totalCandidates) * 100 : 0;

  const numSubjects = new Set(overallRows.map(r => r.subject)).size;

  const kpis = [
    { label: "Overall % 1–6", value: pct1_6.toFixed(1) + "%" },
    { label: "Overall % 1–8", value: pct1_8.toFixed(1) + "%" },
    { label: "Total Candidates", value: totalCandidates },
    { label: "No. of Subjects", value: numSubjects }
  ];

  kpis.forEach(k => {
    const div = document.createElement('div');
    div.className = "kpi";
    div.innerHTML = `
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value">${k.value}</div>
    `;
    kpiWrapper.appendChild(div);
  });
}

function updateCharts() {
  const subject = document.getElementById('subjectSelect').value;
  const metric = document.getElementById('metricSelect').value;

  drawClassChart(subject, metric);
  drawSubjectChart(metric);
}

function drawClassChart(subject, metric) {
  const ctx = document.getElementById('classChart').getContext('2d');

  const rows = records.filter(r => r.subject === subject && r.className.toLowerCase() !== "overall");

  const labels = rows.map(r => r.className);
  const values = rows.map(r => (metric === "pct1_6" ? r.pct1_6 : r.pct1_8));

  if (classChart) classChart.destroy();

  classChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ label: metric === "pct1_6" ? "% 1–6" : "% 1–8", data: values }]
    },
    options: {
      responsive: true,
      scales: { y: { beginAtZero: true, max: 100 } }
    }
  });
}

function drawSubjectChart(metric) {
  const ctx = document.getElementById('subjectChart').getContext('2d');

  const rows = records.filter(r => r.className.toLowerCase() === "overall");

  const labels = rows.map(r => r.subject);
  const values = rows.map(r => (metric === "pct1_6" ? r.pct1_6 : r.pct1_8));

  if (subjectChart) subjectChart.destroy();

  subjectChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ label: metric === "pct1_6" ? "Overall % 1–6" : "Overall % 1–8", data: values }]
    },
    options: {
      responsive: true,
      scales: { y: { beginAtZero: true, max: 100 } }
    }
  });
}
