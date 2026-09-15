/* ============================================================
   SMPAPHMASNA ASSESSMENT 2 DASHBOARD 2026
   Years 7–11

   Expected files:
   assets/y7.xlsx
   assets/y8.xlsx
   assets/y9.xlsx
   assets/y10.xlsx
   assets/y11.xlsx

   Notes:
   - Reads the FIRST worksheet in each workbook.
   - Total Candidates: based on subject "Overall" totals.
   - Overall %1–6 and %1–8: simple average across subjects.
============================================================ */

let parsedRows = [];
let records = [];

let classChart = null;
let subjectChart = null;


/* ============================================================
   YEAR → EXCEL FILE
============================================================ */

const yearToFile = {
  "7": "assets/y7.xlsx",
  "8": "assets/y8.xlsx",
  "9": "assets/y9.xlsx",
  "10": "assets/y10.xlsx",
  "11": "assets/y11.xlsx"
};


/* ============================================================
   INITIALISE DASHBOARD
============================================================ */

document.addEventListener("DOMContentLoaded", () => {

  const yearSelect = document.getElementById("yearSelect");
  const metricSelect = document.getElementById("metricSelect");
  const subjectSelect = document.getElementById("subjectSelect");

  if (!yearSelect || !metricSelect || !subjectSelect) {
    console.error("Dashboard controls could not be found.");
    return;
  }


  /* Change year */
  yearSelect.addEventListener("change", () => {
    loadYearFile(yearSelect.value);
  });


  /* Change metric */
  metricSelect.addEventListener("change", () => {
    updateCharts();
  });


  /* Change subject */
  subjectSelect.addEventListener("change", () => {
    updateCharts();
  });


  /* Initial load */
  loadYearFile(yearSelect.value);

});


/* ============================================================
   LOAD EXCEL FILE
============================================================ */

async function loadYearFile(yearValue) {

  const filePath = yearToFile[yearValue];

  const kpiWrapper = document.getElementById("kpiWrapper");

  if (!filePath) {

    console.error(`No Excel file configured for Year ${yearValue}`);

    kpiWrapper.innerHTML = `
      <div class="placeholder">
        No Excel file is configured for Year ${yearValue}.
      </div>
    `;

    clearCharts();
    disableSelectors();

    return;
  }


  kpiWrapper.innerHTML = `
    <div class="placeholder">
      Loading data for Year ${yearValue}...
    </div>
  `;


  try {

    /* --------------------------------------------------------
       Fetch Excel file
    -------------------------------------------------------- */

    const response = await fetch(filePath);

    if (!response.ok) {

      throw new Error(
        `Cannot load ${filePath} — HTTP ${response.status}`
      );

    }


    /* --------------------------------------------------------
       Read Excel workbook
    -------------------------------------------------------- */

    const buffer = await response.arrayBuffer();

    const data = new Uint8Array(buffer);

    const workbook = XLSX.read(data, {
      type: "array"
    });


    if (!workbook.SheetNames.length) {

      throw new Error(
        `The workbook ${filePath} contains no worksheets.`
      );

    }


    /* --------------------------------------------------------
       Read FIRST worksheet
    -------------------------------------------------------- */

    const firstSheetName = workbook.SheetNames[0];

    console.log(
      `Year ${yearValue}: reading worksheet "${firstSheetName}"`
    );


    const worksheet = workbook.Sheets[firstSheetName];


    parsedRows = XLSX.utils.sheet_to_json(
      worksheet,
      {
        header: 1,
        raw: true,
        defval: ""
      }
    );


    console.log(
      `Year ${yearValue}: ${parsedRows.length} Excel rows read.`
    );


    /* --------------------------------------------------------
       Extract subject/class blocks
    -------------------------------------------------------- */

    records = extractYearSheetRecords(parsedRows);


    console.log(
      `Year ${yearValue}: ${records.length} records extracted.`
    );


    if (!records.length) {

      kpiWrapper.innerHTML = `
        <div class="placeholder">
          <strong>Excel file loaded successfully, but no valid
          subject/class blocks were detected.</strong>
          <br><br>
          Check that each subject table contains:
          <br>
          Class | Taught By | Total | A* | A2 | B3 | B4 |
          C5 | C6 | D7 | E8 | U | Total 1–6 | %1–6 |
          Total 1–8 | %1–8
        </div>
      `;

      clearCharts();
      disableSelectors();

      return;
    }


    /* --------------------------------------------------------
       Initialise dashboard
    -------------------------------------------------------- */

    initControls(records);

    updateKPIs(records);

    updateCharts();

  }

  catch (error) {

    console.error(
      `Error loading Year ${yearValue}:`,
      error
    );


    kpiWrapper.innerHTML = `
      <div class="placeholder">
        <strong>Error loading Year ${yearValue} data.</strong>
        <br><br>
        ${error.message}
        <br><br>
        Expected file:
        <strong>${filePath}</strong>
      </div>
    `;


    records = [];
    parsedRows = [];

    clearCharts();
    disableSelectors();

  }

}


/* ============================================================
   CLEAR CHARTS
============================================================ */

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


/* ============================================================
   DISABLE CONTROLS
============================================================ */

function disableSelectors() {

  const subjectSelect =
    document.getElementById("subjectSelect");

  const metricSelect =
    document.getElementById("metricSelect");


  if (subjectSelect) {

    subjectSelect.disabled = true;

    subjectSelect.innerHTML =
      '<option value="">No data</option>';

  }


  if (metricSelect) {

    metricSelect.disabled = true;

  }

}


/* ============================================================
   EXTRACT RECORDS FROM EXCEL

   Expected pattern:

   SUBJECT NAME

   Distinction / Credit / Pass labels

   Class | Taught By | Total | A* | A2 | B3 | B4 |
   C5 | C6 | D7 | E8 | U |
   Total 1–6 | %1–6 |
   Total 1–8 | %1–8

============================================================ */

function extractYearSheetRecords(rows) {

  const output = [];


  if (!Array.isArray(rows) || rows.length === 0) {

    return output;

  }


  for (let i = 0; i < rows.length; i++) {

    const row = rows[i] || [];


    const column0 =
      String(row[0] ?? "").trim();


    const column2 =
      String(row[2] ?? "").trim();


    /* --------------------------------------------------------
       Detect table header
    -------------------------------------------------------- */

    if (
      column0.toLowerCase() === "class" &&
      column2.toLowerCase() === "total"
    ) {

      /* ------------------------------------------------------
         Find subject name
         Normally 2 rows above the "Class" header
      ------------------------------------------------------ */

      let subject = "";


      if (i - 2 >= 0) {

        subject =
          String(rows[i - 2]?.[0] ?? "").trim();

      }


      /*
         Fallback:
         search a few rows upward if the subject name
         was not found exactly two rows above.
      */

      if (!subject) {

        for (
          let searchRow = i - 1;
          searchRow >= Math.max(0, i - 5);
          searchRow--
        ) {

          const possibleSubject =
            String(
              rows[searchRow]?.[0] ?? ""
            ).trim();


          if (
            possibleSubject &&
            possibleSubject.toLowerCase() !== "class"
          ) {

            subject = possibleSubject;

            break;

          }

        }

      }


      if (!subject) {

        subject = "Unknown Subject";

      }


      /* ------------------------------------------------------
         Read class rows
      ------------------------------------------------------ */

      let j = i + 1;


      while (j < rows.length) {

        const dataRow = rows[j] || [];


        const isBlank =
          dataRow.every(cell => {

            return (
              cell === null ||
              cell === undefined ||
              String(cell).trim() === ""
            );

          });


        if (isBlank) {

          break;

        }


        const className =
          String(
            dataRow[0] ??
            dataRow[1] ??
            ""
          ).trim();


        /*
           Stop if another header begins
        */

        if (
          className.toLowerCase() === "class"
        ) {

          break;

        }


        if (className) {

          output.push({

            subject: subject,

            className: className,

            total: toNumber(dataRow[2]),

            Astar: toNumber(dataRow[3]),

            A2: toNumber(dataRow[4]),

            B3: toNumber(dataRow[5]),

            B4: toNumber(dataRow[6]),

            C5: toNumber(dataRow[7]),

            C6: toNumber(dataRow[8]),

            D7: toNumber(dataRow[9]),

            E8: toNumber(dataRow[10]),

            U: toNumber(dataRow[11]),

            total1_6: toNumber(dataRow[12]),

            pct1_6: normalisePercentage(
              dataRow[13]
            ),

            total1_8: toNumber(dataRow[14]),

            pct1_8: normalisePercentage(
              dataRow[15]
            )

          });

        }


        j++;

      }

    }

  }


  return output;

}


/* ============================================================
   CONVERT VALUE TO NUMBER
============================================================ */

function toNumber(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {

    return 0;

  }


  if (typeof value === "number") {

    return Number.isFinite(value)
      ? value
      : 0;

  }


  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/%/g, "")
    .trim();


  const number = Number(cleaned);


  return Number.isFinite(number)
    ? number
    : 0;

}


/* ============================================================
   NORMALISE PERCENTAGE

   Handles:
   54.3
   54.3%
   0.543
============================================================ */

function normalisePercentage(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {

    return 0;

  }


  let number;


  if (typeof value === "string") {

    number = Number(
      value
        .replace("%", "")
        .trim()
    );

  }

  else {

    number = Number(value);

  }


  if (!Number.isFinite(number)) {

    return 0;

  }


  /*
     Excel percentage stored as decimal.
     Example:
     0.543 → 54.3%
  */

  if (
    number > 0 &&
    number <= 1
  ) {

    number *= 100;

  }


  return number;

}


/* ============================================================
   INITIALISE SUBJECT + METRIC CONTROLS
============================================================ */

function initControls(dataRecords) {

  const subjectSelect =
    document.getElementById("subjectSelect");


  const metricSelect =
    document.getElementById("metricSelect");


  const subjects =
    Array.from(
      new Set(
        dataRecords
          .map(record => record.subject)
          .filter(subject =>
            subject &&
            subject.trim().length > 0
          )
      )
    );


  subjectSelect.innerHTML = "";


  subjects.forEach(subject => {

    const option =
      document.createElement("option");


    option.value = subject;

    option.textContent = subject;


    subjectSelect.appendChild(option);

  });


  const hasSubjects =
    subjects.length > 0;


  subjectSelect.disabled =
    !hasSubjects;


  metricSelect.disabled =
    !hasSubjects;

}


/* ============================================================
   UPDATE KPI CARDS
============================================================ */

function updateKPIs(dataRecords) {

  const kpiWrapper =
    document.getElementById("kpiWrapper");


  kpiWrapper.innerHTML = "";


  if (
    !dataRecords ||
    !dataRecords.length
  ) {

    kpiWrapper.innerHTML = `
      <div class="placeholder">
        No data loaded.
      </div>
    `;

    return;

  }


  /* --------------------------------------------------------
     Find Overall rows
  -------------------------------------------------------- */

  const overallRows =
    dataRecords.filter(record => {

      return (
        record.className &&
        record.className
          .toLowerCase()
          .trim() === "overall"
      );

    });


  /* --------------------------------------------------------
     Total Candidates
  -------------------------------------------------------- */

  let totalCandidates = 0;


  if (overallRows.length) {

    const totals =
      overallRows
        .map(record =>
          Number(record.total) || 0
        )
        .filter(total =>
          total > 0
        );


    if (totals.length) {

      /*
         Use maximum overall total.
         Useful when some subjects have fewer candidates.
      */

      totalCandidates =
        Math.max(...totals);

    }

  }


  /* --------------------------------------------------------
     Fallback total candidates
  -------------------------------------------------------- */

  if (!totalCandidates) {

    const classRows =
      dataRecords.filter(record => {

        return (
          record.className &&
          record.className
            .toLowerCase()
            .trim() !== "overall"
        );

      });


    /*
       Estimate using first subject
       rather than adding every subject.
    */

    if (classRows.length) {

      const firstSubject =
        classRows[0].subject;


      totalCandidates =
        classRows
          .filter(
            record =>
              record.subject === firstSubject
          )
          .reduce(
            (sum, record) =>
              sum +
              (Number(record.total) || 0),
            0
          );

    }

  }


  /* --------------------------------------------------------
     Overall percentages
  -------------------------------------------------------- */

  let overallPct1_6 = 0;

  let overallPct1_8 = 0;


  if (overallRows.length) {

    const valid1_6 =
      overallRows
        .map(record =>
          Number(record.pct1_6)
        )
        .filter(value =>
          Number.isFinite(value)
        );


    const valid1_8 =
      overallRows
        .map(record =>
          Number(record.pct1_8)
        )
        .filter(value =>
          Number.isFinite(value)
        );


    if (valid1_6.length) {

      overallPct1_6 =
        valid1_6.reduce(
          (sum, value) =>
            sum + value,
          0
        ) /
        valid1_6.length;

    }


    if (valid1_8.length) {

      overallPct1_8 =
        valid1_8.reduce(
          (sum, value) =>
            sum + value,
          0
        ) /
        valid1_8.length;

    }

  }


  /* --------------------------------------------------------
     Number of subjects
  -------------------------------------------------------- */

  const subjectSet =
    new Set(
      dataRecords
        .map(record => record.subject)
        .filter(Boolean)
    );


  const numberOfSubjects =
    subjectSet.size;


  /* --------------------------------------------------------
     KPI cards
  -------------------------------------------------------- */

  const kpis = [

    {
      label: "Overall % 1–6",
      value:
        overallRows.length
          ? overallPct1_6.toFixed(1) + "%"
          : "-"
    },

    {
      label: "Overall % 1–8",
      value:
        overallRows.length
          ? overallPct1_8.toFixed(1) + "%"
          : "-"
    },

    {
      label: "Total Candidates",
      value:
        totalCandidates || "-"
    },

    {
      label: "No. of Subjects",
      value:
        numberOfSubjects || "-"
    }

  ];


  kpis.forEach(kpi => {

    const div =
      document.createElement("div");


    div.className = "kpi";


    div.innerHTML = `
      <div class="kpi-label">
        ${escapeHTML(kpi.label)}
      </div>

      <div class="kpi-value">
        ${escapeHTML(String(kpi.value))}
      </div>
    `;


    kpiWrapper.appendChild(div);

  });

}


/* ============================================================
   UPDATE ALL CHARTS
============================================================ */

function updateCharts() {

  if (
    !records ||
    !records.length
  ) {

    clearCharts();

    return;

  }


  const subjectSelect =
    document.getElementById(
      "subjectSelect"
    );


  const metricSelect =
    document.getElementById(
      "metricSelect"
    );


  const subject =
    subjectSelect.value;


  const metric =
    metricSelect.value;


  drawClassChart(
    subject,
    metric
  );


  drawSubjectChart(
    metric
  );

}


/* ============================================================
   CLASS PERFORMANCE CHART
============================================================ */

function drawClassChart(
  subject,
  metric
) {

  const canvas =
    document.getElementById(
      "classChart"
    );


  if (!canvas) {

    return;

  }


  const context =
    canvas.getContext("2d");


  const rows =
    records.filter(record => {

      return (
        record.subject === subject &&
        record.className &&
        record.className
          .toLowerCase()
          .trim() !== "overall"
      );

    });


  const labels =
    rows.map(record =>
      record.className
    );


  const values =
    rows.map(record => {

      return (
        metric === "pct1_6"
          ? record.pct1_6
          : record.pct1_8
      );

    });


  if (classChart) {

    classChart.destroy();

    classChart = null;

  }


  classChart =
    new Chart(
      context,
      {

        type: "bar",

        data: {

          labels: labels,

          datasets: [

            {

              label:
                metric === "pct1_6"
                  ? "% 1–6"
                  : "% 1–8",

              data: values,

              backgroundColor:
                metric === "pct1_6"
                  ? "rgba(0, 64, 128, 0.75)"
                  : "rgba(0, 128, 96, 0.75)",

              borderColor:
                metric === "pct1_6"
                  ? "rgba(0, 64, 128, 1)"
                  : "rgba(0, 128, 96, 1)",

              borderWidth: 1

            }

          ]

        },


        options: {

          responsive: true,

          maintainAspectRatio: false,


          scales: {

            y: {

              beginAtZero: true,

              max: 100,

              title: {

                display: true,

                text: "Percentage (%)"

              }

            },


            x: {

              title: {

                display: true,

                text: "Class"

              }

            }

          },


          plugins: {

            legend: {

              display: false

            },


            tooltip: {

              callbacks: {

                label: context => {

                  return (
                    Number(
                      context.parsed.y
                    ).toFixed(1)
                    + "%"
                  );

                }

              }

            }

          }

        }

      }
    );

}


/* ============================================================
   OVERALL SUBJECT COMPARISON
============================================================ */

function drawSubjectChart(metric) {

  const canvas =
    document.getElementById(
      "subjectChart"
    );


  if (!canvas) {

    return;

  }


  const context =
    canvas.getContext("2d");


  const overallRows =
    records.filter(record => {

      return (
        record.className &&
        record.className
          .toLowerCase()
          .trim() === "overall"
      );

    });


  const labels =
    overallRows.map(
      record =>
        record.subject
    );


  const values =
    overallRows.map(record => {

      return (
        metric === "pct1_6"
          ? record.pct1_6
          : record.pct1_8
      );

    });


  if (subjectChart) {

    subjectChart.destroy();

    subjectChart = null;

  }


  subjectChart =
    new Chart(
      context,
      {

        type: "bar",

        data: {

          labels: labels,

          datasets: [

            {

              label:
                metric === "pct1_6"
                  ? "Overall % 1–6"
                  : "Overall % 1–8",

              data: values,

              backgroundColor:
                metric === "pct1_6"
                  ? "rgba(0, 64, 128, 0.75)"
                  : "rgba(0, 128, 96, 0.75)",

              borderColor:
                metric === "pct1_6"
                  ? "rgba(0, 64, 128, 1)"
                  : "rgba(0, 128, 96, 1)",

              borderWidth: 1

            }

          ]

        },


        options: {

          responsive: true,

          maintainAspectRatio: false,


          scales: {

            y: {

              beginAtZero: true,

              max: 100,

              title: {

                display: true,

                text: "Percentage (%)"

              }

            },


            x: {

              title: {

                display: true,

                text: "Subject"

              }

            }

          },


          plugins: {

            legend: {

              display: false

            },


            tooltip: {

              callbacks: {

                label: context => {

                  return (
                    Number(
                      context.parsed.y
                    ).toFixed(1)
                    + "%"
                  );

                }

              }

            }

          }

        }

      }
    );

}


/* ============================================================
   ESCAPE HTML
============================================================ */

function escapeHTML(value) {

  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

}
