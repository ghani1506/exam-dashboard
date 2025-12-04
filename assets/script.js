
document.getElementById('fileInput').addEventListener('change', handleFile);

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = evt.target.result;
    const wb = XLSX.read(data, {type:'binary'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws);

    // Expect a column named "Marks"
    const marks = json.map(r => Number(r.Marks));

    const ctx = document.getElementById('chart1').getContext('2d');
    new Chart(ctx, {
      type: 'bar',
      data: {
        labels: marks.map((v,i)=> "Student " + (i+1)),
        datasets: [{
          label: 'Marks',
          data: marks
        }]
      }
    });
  };
  reader.readAsBinaryString(file);
}
