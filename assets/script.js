/* --- script.js for multi-year exam dashboard (Years 7–10) ---
   - Loads embedded Excel files: assets/year7.xlsx, year8.xlsx, year9.xlsx, year10.xlsx
   - Total Candidates: based on 'Overall' totals (cohort size), not summed across subjects
   - Overall %1–6 and %1–8: SIMPLE AVERAGE of each subject's Overall %1–6 / %1–8
*/

let parsedRows = [];
let records = [];
let classChart = null;
let subjectChart = null;

// Map year selector to internal Excel filenames
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

  // Initial load (default year in the dropdown)
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
      const ws = wb.Sheets[wb.SheetNames[0]]; // first sheet
      parsedRows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
      records = extractYearSheetRecords(parsedRows);

      if (!records.length) {
        kpiWrapper.innerHTML = '<div class="placeholder">File loaded but no valid subject/class blocks were detected.</div>';
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
      kpiWrapper.innerHTML = '<div class="placeholder">Error loading data for this year. Check that the Excel file exists and matches the expected format.</div>';
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
 * Extract tidy records from a sheet with this pattern:
 *   row i-2: SUBJECT NAME
 *   row i-1: Distinction/Credit/Pass/Fail labels
 *   row i:   Class | Taught By | Total | A* | A2 | B3 | B4 | C5 | C6 | D7 | E8 | U | Total 1-6 | %1-6 | Total 1-8 | %1-8
 *   rows i+1... : data rows until a blank row
 */
function extractYearSheetRecords(rows) {
  const out = [];
  if (!rows || !rows.length) return out;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const c0 = (row[0] ?? "").toString().trim();
    const c2 = (row[2] ?? "").toString().trim();

    if (c0 === "Class" && c2 === "Total") {
      // subject label is typically two rows above
      let subject = "";
      if (i - 2 >= 0) {
        subject = (rows[i - 2][0] ?? "").toString().trim();
      }

      // collect detail rows
      let j = i + 1;
      while (j < rows.length) {
        const r = rows[j] || [];
        const isBlank = r.every(cell => cell === null || cell === undefined || cell === "");
        if (isBlank) break;

        const className = (r[0] ?? r[1] ?? "").toString().trim();
        if (className) {
          out.push({
            subject,
            className,
            total: toNumber(r[2]),
            Astar: toNumber(r[3]),
            A2: toNumber(r[4]),
            B3: toNumber(r[5]),
            B4: toNumber(r[6]),
            C5: toNumber(r[7]),
            C6: toNumber(r[8]),
            D7: toNumber(r[9]),
            E8: toNumber(r[10]),
            U: toNumber(r[11]),
            total1_6: toNumber(r[12]),
            pct1_6: toNumber(r[13]),
            total1_8: toNumber(r[14]),
            pct1_8: toNumber(r[15])
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
  metricSelect.disabled = subjects.length === 0;
}

/**
 * KPI logic:
 * - Total candidates:
 *     based on 'Overall' totals (cohort size), not sum across subjects.
 * - Overall %1–6 and %1–8:
 *     SIMPLE AVERAGE of each subject's "Overall %1–6" / "Overall %1–8".
 */
function updateKPIs(records) {
  const kpiWrapper = document.getElementById('kpiWrapper');
  kpiWrapper.innerHTML = "";

  if (!records || !records.length) {
    kpiWrapper.innerHTML = '<div class="placeholder">No data loaded.</div>';
    return;
  }

  // Subject-level "Overall" rows
  const overallRows = records.filter(
    r => r.className && r.className.toLowerCase() === "overall"
  );

  // ===== Total Candidates (cohort size) =====
  let totalCandidates = 0;
  if (overallRows.length) {
    const totals = overallRows
      .map(r => r.total || 0)
      .filter(t => t > 0);

    if (totals.length) {
      const unique = Array.from(new Set(totals));
      if (unique.length === 1) {
        totalCandidates = unique[0];      // all subjects agree on cohort size
      } else {
        totalCandidates = Math.max(...unique); // use max if they differ
      }
    }
  }
  // Fallback: if no overall rows or no usable totals
  if (!totalCandidates) {
    totalCandidates = overallRows.reduce((sum, r) => sum + (r.total || 0), 0);
  }

  // ===== Overall %1–6 and %1–8 (simple average across subjects) =====
  let overallPct1_6 = 0;
  let overallPct1_8 = 0;

  if (overallRows.length) {
    const avg1_6 =
      overallRows.reduce((sum, r) => sum + (r.pct1_6 || 0), 0) / overallRows.length;
    const avg1_8 =
      overallRows.reduce((sum, r) => sum + (r.pct1_8 || 0), 0) / overallRows.length;

    overallPct1_6 = avg1_6;
    overallPct1_8 = avg1_8;
  } else {
    // Fallback if no Overall rows: average across all rows
    const avg1_6 =
      records.reduce((sum, r) => sum + (r.pct1_6 || 0), 0) / records.length;
    const avg1_8 =
      records.reduce((sum, r) => sum + (r.pct1_8 || 0), 0) / records.length;

    overallPct1_6 = avg1_6;
    overallPct1_8 = avg1_8;
  }

  const numSubjects = new Set(overallRows.map(r => r.subject)).size || 0;

  const kpis = [
    { label: "Overall % 1–6", value: overallPct1_6.toFixed(1) + "%" },
    { label: "Overall % 1–8", value: overallPct1_8.toFixed(1) + "%" },
    { label: "Total Candidates", value: totalCandidates || "-" },
    { label: "No. of Subjects", value: numSubjects || "-" }
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
  const subject = document.getElementById('subjectSelect').value;
  const metric = document.getElementById('metricSelect').value;

  drawClassChart(subject, metric);
  drawSubjectChart(metric);
}

function drawClassChart(subject, metric) {
  const canvas = document.getElementById('classChart');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  const rows = records.filter(
    r => r.subject === subject &&
         r.className &&
         r.className.toLowerCase() !== "overall"
  );

  const labels = rows.map(r => r.className);
  const values = rows.map(r => (metric === "pct1_6" ? r.pct1_6 : r.pct1_8));

  if (classChart) {
    classChart.destroy();
    classChart = null;
  }

  classChart = new Chart(ctx, {
    type: "bar",
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
            label: ctx => ctx.parsed.y.toFixed(1) + "%"
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

  const overallRows = records.filter(
    r => r.className && r.className.toLowerCase() === "overall"
  );

  const labels = overallRows.map(r => r.subject);
  const values = overallRows.map(
    r => (metric === "pct1_6" ? r.pct1_6 : r.pct1_8)
  );

  if (subjectChart) {
    subjectChart.destroy();
    subjectChart = null;
  }

  subjectChart = new Chart(ctx, {
    type: "bar",
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
            label: ctx => ctx.parsed.y.toFixed(1) + "%"
          }
        }
      }
    }
  });
}
