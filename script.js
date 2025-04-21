fetch('links.xlsx')
  .then(res => res.arrayBuffer())
  .then(buffer => {
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    const container = document.getElementById('linksContainer');
    data.forEach((row, index) => {
      if (row.Link) {
        const card = document.createElement('div');
        card.className = 'link-card';
        card.innerHTML = `<p><a href="${row.Link}" target="_blank">Event Link ${index + 1}</a></p>`;
        container.appendChild(card);
      }
    });
  })
  .catch(err => console.error('Error loading or parsing Excel file:', err));
