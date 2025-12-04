
let parsedRows = [];
let records = [];
let classChart = null;
let subjectChart = null;

// Mapping of year value to embedded Excel filename
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

  // Initial load for default year (Year 7)
  loadYearFile(yearSelect.value);
});

function loadYearFile(yearValue) {
  const filePath = yearToFile[yearValue];
  const kpiWrapper = document.getElementById('kpiWrapper');
  kpiWrapper.innerHTML = '<div class="placeholder">Loading data for Year ' + yearValue + '...</div>';

  fetch(filePath)
    .then(resp => {
      if (!resp.ok) {
        throw new Error("Cannot load " + filePath + " (HTTP " + resp.status + ")");
      }
      return resp.arrayBuffer();
    })
    .then(buf => {
      const data = new Uint8Array(buf);
      const wb = XLSX.read(data, { type: 'array' });
      const firstSheetName = wb.SheetNames[0];
      const ws = wb.Sheets[firstSheetName];
      parsedRows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
      records = extractYearSheetRecords(parsedRows);

      if (!records.length) {
        kpiWrapper.innerHTML = '<div class="placeholder">File loaded but could not detect any subject/class records. Please check the layout.</div>';
        clearCharts();
        disableSelectors();
        return;
      }

      initControls(records);
      updateKPIs(records);
      updateCharts();
    })
    .catch(err => {
      console.error("Error loading year file:", err);
      kpiWrapper.innerHTML = '<div class="placeholder">Error loading file for this year. Please ensure it exists and matches the expected layout.</div>';
      clearCharts();
      disableSelectors();
    });
}

function clearCharts() {
  if (classChart) {
    classChart.destroy();
    classChart = null;
  }
  if (subjectChart) {
    subjectChart.destroy();
    subjectChart = null;
  }
}

function disableSelectors() {
  document.getElementById('subjectSelect').disabled = true;
  document.getElementById('metricSelect').disabled = true;
}

/**
 * Extract tidy records from the sheet.
 * Same structure as your YEAR 7 analysis:
 *   row i-2: SUBJECT NAME
 *   row i-1: Distinction/Credit/Pass/Fail labels
 *   row i:   Class | Taught By | Total | A* | A2 | B3 | B4 | C5 | C6 | D7 | E8 | U | Total 1-6 | %1-6 | Total 1-8 | %1-8
 *   rows i+1... : data rows, ending with blank row
 */
function extractYearSheetRecords(rows) {
  const out = [];
  if (!rows || !rows.length) return out;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const c0 = (row[0] ?? "").toString().trim();
    const c2 = (row[2] ?? "").toString().trim();

    if (c0 === "Class" && c2 === "Total") {
      let subject = "";
      if (i - 2 >= 0) {
        subject = (rows[i - 2][0] ?? "").toString().trim();
      }

      let j = i + 1;
      while (j < rows.length) {
        const r = rows[j] || [];
        const isBlank = r.every(cell => cell === null || cell === undefined || cell === "");
        if (isBlank) break;

        let className = (r[0] ?? r[1] ?? "").toString().trim();
        const total = toNumber(r[2]);
        const Astar = toNumber(r[3]);
        const A2 = toNumber(r[4]);
        const B3 = toNumber(r[5]);
        const B4 = toNumber(r[6]);
        const C5 = toNumber(r[7]);
        const C6 = toNumber(r[8]);
        const D7 = toNumber(r[9]);
        const E8 = toNumber(r[10]);
        const U = toNumber(r[11]);
        const total1_6 = toNumber(r[12]);
        const pct1_6 = toNumber(r[13]);
        const total1_8 = toNumber(r[14]);
        const pct1_8 = toNumber(r[15]);

        if (className) {
          out.push({
            subject,
            className,
            total,
            Astar,
            A2,
            B3,
            B4,
            C5,
            C6,
            D7,
            E8,
            U,
            total1_6,
            pct1_6,
            total1_8,
            pct1_8
          });
        }
        j++;
      }
    }
  }
  return out;
}

function toNumber(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function initControls(records) {
  const subjectSelect = document.getElementById('subjectSelect');
  const metricSelect = document.getElementById('metricSelect');

  const subjects = Array.from(new Set(records.map(r => r.subject).filter(s => s && s.length)));

  subjectSelect.innerHTML = "";
  subjects.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    subjectSelect.appendChild(opt);
  });

  subjectSelect.disabled = subjects.length === 0;
  metricSelect.disabled = false;

  if (!subjects.length) {
    const opt = document.createElement('option');
    opt.textContent = "No subjects found";
    subjectSelect.appendChild(opt);
    metricSelect.disabled = true;
  }
}

function updateKPIs(records) {
  const kpiWrapper = document.getElementById('kpiWrapper');
  kpiWrapper.innerHTML = "";

  if (!records || !records.length) {
    kpiWrapper.innerHTML = '<div class="placeholder">No data loaded.</div>';
    return;
  }

  const overallRows = records.filter(r => r.className && r.className.toLowerCase() === "overall");
  const source = overallRows.length ? overallRows : records;

  const totalCandidates = source.reduce((sum, r) => sum + (r.total || 0), 0);

  let weighted1_6 = 0;
  let weighted1_8 = 0;
  if (totalCandidates > 0) {
    weighted1_6 = source.reduce((sum, r) => sum + (r.pct1_6 || 0) * (r.total || 0), 0) / totalCandidates;
    weighted1_8 = source.reduce((sum, r) => sum + (r.pct1_8 || 0) * (r.total || 0), 0) / totalCandidates;
  }

  const numSubjects = new Set(source.map(r => r.subject)).size;

  const kpis = [
    { label: "Overall % 1–6", value: weighted1_6.toFixed(1) + "%" },
    { label: "Overall % 1–8", value: weighted1_8.toFixed(1) + "%" },
    { label: "Total Candidates", value: totalCandidates.toString() },
    { label: "No. of Subjects", value: numSubjects.toString() }
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
  if (!records || !records.length) return;
  const metric = document.getElementById('metricSelect').value;
  const subject = document.getElementById('subjectSelect').value;
  drawClassChart(subject, metric);
  drawSubjectChart(metric);
}

function drawClassChart(subject, metric) {
  const canvas = document.getElementById('classChart');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  const dataForSubject = records.filter(r => r.subject === subject && r.className && r.className.toLowerCase() !== "overall");

  const labels = dataForSubject.map(r => r.className);
  const values = dataForSubject.map(r => metric === "pct1_6" ? r.pct1_6 : r.pct1_8);

  if (classChart) {
    classChart.destroy();
    classChart = null;
  }

  classChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: metric === "pct1_6" ? "% 1–6" : "% 1–8",
        data: values
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          beginAtZero: true,
          max: 100,
          title: { display: true, text: "Percentage (%)" }
        },
        x: {
          title: { display: true, text: "Class" }
        }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              return ctx.parsed.y.toFixed(1) + "%";
            }
          }
        }
      }
    }
  });
}

function drawSubjectChart(metric) {
  const canvas = document.getElementById('subjectChart');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  let perSubject = {};
  records.forEach(r => {
    const key = r.subject || "";
    if (!key) return;

    const isOverall = r.className && r.className.toLowerCase() === "overall";
    if (isOverall) {
      perSubject[key] = metric === "pct1_6" ? r.pct1_6 : r.pct1_8;
    } else if (!(key in perSubject)) {
      const val = metric === "pct1_6" ? r.pct1_6 : r.pct1_8;
      perSubject[key] = val;
    }
  });

  const labels = Object.keys(perSubject);
  const values = labels.map(k => perSubject[k]);

  if (subjectChart) {
    subjectChart.destroy();
    subjectChart = null;
  }

  subjectChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: metric === "pct1_6" ? "Overall % 1–6" : "Overall % 1–8",
        data: values
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          beginAtZero: true,
          max: 100,
          title: { display: true, text: "Percentage (%)" }
        },
        x: {
          title: { display: true, text: "Subject" }
        }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              return ctx.parsed.y.toFixed(1) + "%";
            }
          }
        }
      }
    }
  });
}
