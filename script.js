const fileUpload = document.getElementById('file-upload');
const fileName = document.getElementById('file-name');
const excelData = document.getElementById('excel-data');
const songsSection = document.getElementById('songs-section');
const dropZone = document.getElementById('drop-zone');
const startRowInput = document.getElementById('start-row');
const showNameInput = document.getElementById('show-name');
const distributionIdContainer = document.getElementById('distribution-id-inputs');
const submitBtn = document.getElementById('submit-btn');
const exportAllBtn = document.getElementById('export-all-btn');
const loader = document.getElementById('loader');

let workbookData = {};

// Handle file select
fileUpload.addEventListener('change', (event) => {
    handleFile(event.target.files[0]);
});

// Drag and drop events
dropZone.addEventListener('dragover', (event) => {
    event.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragleave');
});

dropZone.addEventListener('drop', (event) => {
    event.preventDefault();
    dropZone.classList.remove('dragover');
    handleFile(event.dataTransfer.files[0]);
});

submitBtn.addEventListener('click', () => {
    if (Object.keys(workbookData).length > 0) {
        songsSection.style.display = 'block';
        displayAllEpisodeData();
        exportAllBtn.style.display = 'block';
    } else {
        alert('Please select an Excel file first.');
    }
});

exportAllBtn.addEventListener('click', () => {
    const showName = showNameInput.value.trim();
    if (!showName) {
        alert('Please enter a Title name.');
        return;
    }

    const zip = new JSZip();
    const episodeRegex = /Episode #(\d+)/;

    Object.keys(workbookData).forEach(sheetName => {
        const match = sheetName.match(episodeRegex);
        if (match) {
            const episodeNumber = match[1];
            const sheetData = workbookData[sheetName];
            const distIdInput = document.getElementById(`dist-id-${episodeNumber}`);
            const distributionId = distIdInput ? distIdInput.value.trim() : '';
            const processedData = processSheetData(sheetData, distributionId);
            
            const filename = getFileName(showName, episodeNumber);
            const ws = XLSX.utils.aoa_to_sheet(processedData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Music Rights');
            
            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            zip.file(filename, wbout);
        }
    });

    zip.generateAsync({ type: "blob" }).then(function(content) {
        const zipFilename = `Music_Rights_Upload_${showName}_All_Episodes.zip`;
        saveAs(content, zipFilename);
    });
});

function handleFile(file) {
    if (file) {
        fileName.textContent = file.name;
        excelData.innerHTML = ''; // Clear previous data
        songsSection.style.display = 'none'; // Hide section until submit
        exportAllBtn.style.display = 'none';
        distributionIdContainer.innerHTML = ''; // Clear old inputs
        
        // Disable submit button and show loader
        submitBtn.disabled = true;
        loader.style.display = 'block';

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', sheetStubs: true });
            workbookData = {};
            const episodeRegex = /Episode #(\d+)/;

            workbook.SheetNames.forEach(sheetName => {
                const match = sheetName.match(episodeRegex);
                if (match) {
                    const episodeNumber = match[1];
                    const worksheet = workbook.Sheets[sheetName];
                    workbookData[sheetName] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    createDistributionIdInput(episodeNumber);
                }
            });

            // Re-enable submit button and hide loader
            submitBtn.disabled = false;
            loader.style.display = 'none';
        };
        reader.readAsArrayBuffer(file);
    } else {
        fileName.textContent = 'or drag and drop it here';
        excelData.innerHTML = '';
        songsSection.style.display = 'none';
        exportAllBtn.style.display = 'none';
        distributionIdContainer.innerHTML = '';
        workbookData = {};
    }
}

function createDistributionIdInput(episodeNumber) {
    const inputGroup = document.createElement('div');
    inputGroup.className = 'dist-id-input-group';

    const label = document.createElement('label');
    label.htmlFor = `dist-id-${episodeNumber}`;
    label.textContent = `Distribution ID for Episode #${episodeNumber}:`;
    inputGroup.appendChild(label);

    const input = document.createElement('input');
    input.type = 'text';
    input.id = `dist-id-${episodeNumber}`;
    input.className = 'distribution-id-input';
    inputGroup.appendChild(input);

    distributionIdContainer.appendChild(inputGroup);
}

function displayAllEpisodeData() {
    excelData.innerHTML = ''; // Clear previous display
    const episodeRegex = /Episode #(\d+)/;

    Object.keys(workbookData).forEach(sheetName => {
        const match = sheetName.match(episodeRegex);
        if (match) {
            const episodeNumber = match[1];
            const sheetData = workbookData[sheetName];
            
            const distIdInput = document.getElementById(`dist-id-${episodeNumber}`);
            const distributionId = distIdInput ? distIdInput.value.trim() : '';

            const processedData = processSheetData(sheetData, distributionId);

            const episodeContainer = document.createElement('div');
            episodeContainer.className = 'episode-container';

            const title = document.createElement('h2');
            title.textContent = `Episode #${episodeNumber}`;
            episodeContainer.appendChild(title);

            const tableHTML = generateTableForSheet(processedData);
            const tableContainer = document.createElement('div');
            tableContainer.innerHTML = tableHTML;
            episodeContainer.appendChild(tableContainer);

            const exportButton = document.createElement('button');
            exportButton.className = 'export-btn';
            exportButton.textContent = `Export Episode #${episodeNumber}`;
            exportButton.onclick = () => {
                exportSingleEpisode(processedData, episodeNumber);
            };
            episodeContainer.appendChild(exportButton);

            excelData.appendChild(episodeContainer);
        }
    });

    if (excelData.innerHTML === '') {
        excelData.innerHTML = '<p>No sheets found with the format "Episode #NNN".</p>';
    }
}

function generateTableForSheet(processedData) {
    let table = '<table>';
    table += `<thead><tr><th>DISTRIBUTION_ID</th><th>SONG_TITLE</th><th>ARTISTS</th><th>WRITERS</th></tr></thead>`;
    table += '<tbody>';
    // Skip header row for display
    for (let i = 1; i < processedData.length; i++) {
        const row = processedData[i];
        table += '<tr>';
        table += `<td>${row[0]}</td>`;
        table += `<td>${row[1]}</td>`;
        table += `<td>${row[2]}</td>`;
        table += `<td>${row[3]}</td>`;
        table += '</tr>';
    }
    table += '</tbody></table>';
    return table;
}

function processSheetData(data, distributionId) {
    const startRow = parseInt(startRowInput.value, 10);
    
    if (!startRow || startRow <= 0) {
        alert('Please enter a valid starting row number.');
        return [['DISTRIBUTION_ID', 'SONG_TITLE', 'ARTISTS', 'WRITERS']];
    }

    const processedData = [['DISTRIBUTION_ID', 'SONG_TITLE', 'ARTISTS', 'WRITERS']];
    const uniqueRows = new Set();
    const startIndex = startRow;

    for (let i = startIndex; i < data.length; i++) {
        const row = data[i];
        if (row && row.length > 0 && (row[0] || row[1])) {
            let firstCol = row[0] || '';
            let artists = row[1] || '';
            
            if (artists.toUpperCase() === 'N/A') {
                artists = '';
            }

            let songTitle = '';
            let writers = '';

            const splitRegex = /Written by/i;
            const match = firstCol.match(splitRegex);

            if (match) {
                const splitIndex = match.index;
                songTitle = firstCol.substring(0, splitIndex).replace(/"/g, '').trim();
                writers = firstCol.substring(splitIndex + match[0].length).trim();
            } else {
                songTitle = firstCol.replace(/"/g, '').trim();
            }

            const rowIdentifier = `${songTitle}|${artists}|${writers}`;

            if (!uniqueRows.has(rowIdentifier)) {
                uniqueRows.add(rowIdentifier);
                const newRow = [distributionId, songTitle, artists, writers];
                processedData.push(newRow);
            }
        }
    }
    return processedData;
}

function getFileName(showName, episodeNumber) {
    const date = new Date();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    const yy = String(date.getFullYear()).slice(-2);
    const formattedDate = `${mm}_${dd}_${yy}`;
    return `Music_Rights_Upload_${showName}_${episodeNumber}_${formattedDate}.xlsx`;
}

function exportSingleEpisode(data, episodeNumber) {
    const showName = showNameInput.value.trim();
    if (!showName) {
        alert('Please enter a Title name.');
        return;
    }
    const filename = getFileName(showName, episodeNumber);
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Music Rights');
    XLSX.writeFile(wb, filename);
}