
let parsedRows = [];
let records = []; // flattened records per class+subject
let classChart = null;
let subjectChart = null;

document.getElementById('fileInput').addEventListener('change', handleFile);
document.getElementById('metricSelect').addEventListener('change', () => {
  updateCharts();
});
document.getElementById('subjectSelect').addEventListener('change', () => {
  updateCharts();
});

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = evt.target.result;
    const wb = XLSX.read(data, { type: 'binary' });

    // Take the first sheet by default (YEAR 7)
    const firstSheetName = wb.SheetNames[0];
    const ws = wb.Sheets[firstSheetName];

    // 2D array of rows
    parsedRows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
    records = extractYear7Records(parsedRows);

    if (!records.length) {
      alert("No valid data found. Please ensure you uploaded the correct YEAR 7 analysis file.");
      return;
    }

    initControls(records);
    updateKPIs(records);
    updateCharts();
  };
  reader.readAsBinaryString(file);
}

/**
 * Extract tidy records from the YEAR 7 sheet structure.
 * Each block looks like:
 *   row i-2: SUBJECT NAME (e.g. 'BAHASA MELAYU')
 *   row i-1: Distinction/Credit/Pass/Fail labels
 *   row i:   Class | Taught By | Total | A* | A2 | B3 | B4 | C5 | C6 | D7 | E8 | U | Total 1-6 | %1-6 | Total 1-8 | %1-8
 *   rows i+1... : data rows, ending with blank row
 */
function extractYear7Records(rows) {
  const out = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const c0 = (row[0] || "").toString().trim();
    const c2 = (row[2] || "").toString().trim();

    // Header row detection: first cell 'Class' and third cell 'Total'
    if (c0 === "Class" && c2 === "Total") {
      // Subject name is typically two rows above
      let subject = "";
      if (i - 2 >= 0) {
        subject = (rows[i - 2][0] || "").toString().trim();
      }

      // From i+1 downwards collect until a fully blank row
      let j = i + 1;
      while (j < rows.length) {
        const r = rows[j] || [];
        const isBlank = r.every(cell => cell === null || cell === undefined || cell === "");
        if (isBlank) break;

        const className = (r[0] || r[1] || "").toString().trim();
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

  // Unique subject list
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
  }
}

function updateKPIs(records) {
  const kpiWrapper = document.getElementById('kpiWrapper');
  kpiWrapper.innerHTML = "";

  if (!records.length) {
    kpiWrapper.innerHTML = '<div class="placeholder">No data loaded.</div>';
    return;
  }

  // Use 'Overall' rows per subject if available, else all rows
  const overallRows = records.filter(r => r.className.toLowerCase() === "overall");
  const source = overallRows.length ? overallRows : records;

  const totalCandidates = source.reduce((sum, r) => sum + (r.total || 0), 0);

  // Weighted averages of %1-6 and %1-8 (by total candidates)
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
  const metric = document.getElementById('metricSelect').value;
  const subject = document.getElementById('subjectSelect').value;

  if (!records.length) return;

  drawClassChart(subject, metric);
  drawSubjectChart(metric);
}

function drawClassChart(subject, metric) {
  const ctx = document.getElementById('classChart').getContext('2d');

  const dataForSubject = records.filter(r => r.subject === subject && r.className.toLowerCase() !== "overall");

  const labels = dataForSubject.map(r => r.className);
  const values = dataForSubject.map(r => metric === "pct1_6" ? r.pct1_6 : r.pct1_8);

  if (classChart) {
    classChart.destroy();
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
  const ctx = document.getElementById('subjectChart').getContext('2d');

  // Prefer 'Overall' rows for subject comparison
  let perSubject = {};
  records.forEach(r => {
    const key = r.subject || "";
    if (!key) return;

    const isOverall = r.className.toLowerCase() === "overall";
    if (isOverall) {
      perSubject[key] = metric === "pct1_6" ? r.pct1_6 : r.pct1_8;
    } else if (!(key in perSubject)) {
      // If no overall row, approximate by simple average
      const val = metric === "pct1_6" ? r.pct1_6 : r.pct1_8;
      perSubject[key] = val;
    }
  });

  const labels = Object.keys(perSubject);
  const values = labels.map(k => perSubject[k]);

  if (subjectChart) {
    subjectChart.destroy();
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
