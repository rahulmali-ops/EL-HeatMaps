// ==================== POSITIVE MODE (LEFT) ====================
const wrapperPos = document.getElementById('mapWrapperPos');
const tooltipPos = document.getElementById('tooltipPos');
const selectorPos = document.getElementById('dataSelectorPos');
const fileInputPos = document.getElementById('excelUploadPos');
const statusPos = document.getElementById('statusPos');
const targetInputPos = document.getElementById('fileCountTargetPos');

let globalDataPos = [];
let myChartPos = null;
const rowLetters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"];

fileInputPos.addEventListener('change', async (e) => {
    const files = Array.from(e.target.files);
    const targetCount = parseInt(targetInputPos.value);

    if (files.length !== targetCount) {
        statusPos.innerText = `❌ Error: Please upload exactly ${targetCount} files.`;
        return;
    }

    statusPos.innerText = "Processing files...";
    
    try {
        const allFilesData = [];
        for (const file of files) {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            allFilesData.push(XLSX.utils.sheet_to_json(firstSheet));
        }

        // Filter data for STATION = "Positive"
        const positiveData = allFilesData.map(fileData => 
            fileData.filter(row => row.STATION === "Positive")
        );

        // Check if positive data exists
        const hasPositiveData = positiveData.some(data => data.length > 0);
        
        if (!hasPositiveData) {
            statusPos.innerText = `❌ Error: No "Positive" station data found in uploaded files.`;
            wrapperPos.innerHTML = '<div style="padding: 20px; text-align: center; color: #ef4444;">No Positive Station Data</div>';
            return;
        }

        globalDataPos = calculateWeldAveragesWithHistory(positiveData);
        statusPos.innerText = `✓ Loaded ${files.length} file(s) - Positive data (${globalDataPos.length} records)`;
        renderMapPos();
    } catch (err) {
        statusPos.innerText = "❌ Error processing files.";
        console.error(err);
    }
});

function calculateWeldAveragesWithHistory(filesData) {
    // Find the longest dataset to use as template
    const baseFile = filesData.reduce((longest, current) => 
        current.length > longest.length ? current : longest
    );

    if (baseFile.length === 0) return [];

    const numFiles = filesData.length;

    return baseFile.map((row, rowIndex) => {
        let averagedRow = { ...row };
        
        const params = ['AF', 'BF', 'Resistance', 'PeakCurrent', 'AverageCurrent', 'PeakVoltage', 'AverageVoltage', 'Power'];
        const suffixes = ['1', '2'];

        params.forEach(param => {
            suffixes.forEach(num => {
                let key = "";
                if (param === 'AF') key = `STNLOADVALUEAFTERFORCER2CELL${num}`;
                else if (param === 'BF') key = `STNLOADVALUEBEFOREFORCER2CELL${num}`;
                else if (param === 'Resistance') key = `STNR2RESISTANCE${num}mohms`;
                else if (param === 'PeakCurrent') key = `STNR2PEAKCURRENT${num}KA`;
                else if (param === 'AverageCurrent') key = `STNR2AVERAGECURRENT${num}KA`;
                else if (param === 'PeakVoltage') key = `STNR2PEAKVOLTAGE${num}V`;
                else if (param === 'AverageVoltage') key = `STNR2AVERAGEVOLTAGE${num}V`;
                else if (param === 'Power') key = `STNR2POWER${num}KW`;

                let sum = 0;
                let history = [];
                let validCount = 0;

                filesData.forEach(fileData => {
                    if (fileData[rowIndex]) {
                        const val = parseFloat(fileData[rowIndex][key]) || 0;
                        sum += val;
                        history.push(val);
                        validCount++;
                    }
                });
                
                averagedRow[key] = validCount > 0 ? sum / validCount : 0;
                averagedRow[`${key}_history`] = history;
            });
        });

        return averagedRow;
    });
}

function getColor(val, min, max) {
    const range = max - min || 1;
    const normalized = (val - min) / range;
    if (normalized > 0.8) return '#ff0000';
    if (normalized > 0.6) return '#ffff00';
    if (normalized > 0.4) return '#2cff96';
    if (normalized > 0.2) return '#0098ff';
    return '#96005a';
}

function renderMapPos() {
    if (globalDataPos.length === 0) return;
    const selectedParam = selectorPos.value;
    wrapperPos.innerHTML = '';

    const getField = (num) => {
        const map = {
            'AF': `STNLOADVALUEAFTERFORCER2CELL${num}`,
            'BF': `STNLOADVALUEBEFOREFORCER2CELL${num}`,
            'Resistance': `STNR2RESISTANCE${num}mohms`,
            'PeakCurrent': `STNR2PEAKCURRENT${num}KA`,
            'AverageCurrent': `STNR2AVERAGECURRENT${num}KA`,
            'PeakVoltage': `STNR2PEAKVOLTAGE${num}V`,
            'AverageVoltage': `STNR2AVERAGEVOLTAGE${num}V`,
            'Power': `STNR2POWER${num}KW`
        };
        return map[selectedParam];
    };

    let allValues = [];
    globalDataPos.forEach(entry => {
        allValues.push(parseFloat(entry[getField(1)]) || 0, parseFloat(entry[getField(2)]) || 0);
    });
    
    const min = Math.min(...allValues);
    const max = Math.max(...allValues);
    document.getElementById('maxLabelPos').innerText = max.toFixed(2);
    document.getElementById('minLabelPos').innerText = min.toFixed(2);

    // 16 columns × 11 rows = 176 cells
    globalDataPos.forEach((entry, index) => {
        if (index >= 176) return;
        const col = index % 16;
        const row = Math.floor(index / 16);
        const letter = rowLetters[row];

        const cell = document.createElement('div');
        cell.className = 'weld-cell';

        const weldSet = [
            { type: 'L1', val: parseFloat(entry[getField(1)]) || 0, label: 'L1', history: entry[`${getField(1)}_history`] },
            { type: 'L2', val: parseFloat(entry[getField(2)]) || 0, label: 'L2', history: entry[`${getField(2)}_history`] }
        ];

        weldSet.forEach(w => {
            const dot = document.createElement('div');
            dot.className = `weld weld-${w.type}`;
            dot.id = `weld-pos-${letter}-${col + 1}-${w.type}`;
            dot.style.backgroundColor = getColor(w.val, min, max);

            dot.onmouseover = (e) => {
                tooltipPos.style.opacity = 1;
                
                document.getElementById('tooltip-text-pos').innerHTML = `
                    <div style="margin-bottom: 8px; border-bottom: 1px solid #334155; padding-bottom: 4px;">
                        <strong>${w.label} - ${letter}${col+1}</strong>
                    </div>
                    Avg ${selectedParam}: <strong>${w.val.toFixed(4)}</strong>
                `;

                const ctx = document.getElementById('weldChartPos').getContext('2d');
                if (myChartPos) myChartPos.destroy();

                myChartPos = new Chart(ctx, {
                    type: 'line',
                    data: {
                        labels: w.history.map((_, i) => `F${i + 1}`),
                        datasets: [{
                            label: selectedParam,
                            data: w.history,
                            borderColor: '#38bdf8',
                            backgroundColor: 'rgba(56, 189, 248, 0.1)',
                            borderWidth: 2,
                            tension: 0.3,
                            fill: true,
                            pointRadius: 4,
                            pointBackgroundColor: '#38bdf8'
                        }]
                    },
                    options: {
                        responsive: false,
                        animation: { duration: 300 },
                        plugins: {
                            legend: { display: false },
                            tooltip: { enabled: true }
                        },
                        scales: {
                            y: {
                                grid: { color: 'rgba(255,255,255,0.1)' },
                                ticks: { color: '#94a3b8', font: { size: 9 } }
                            },
                            x: {
                                grid: { display: false },
                                ticks: { color: '#94a3b8', font: { size: 9 } }
                            }
                        }
                    }
                });
            };

            dot.onmousemove = (e) => {
                tooltipPos.style.left = (e.pageX + 15) + 'px';
                tooltipPos.style.top = (e.pageY - 50) + 'px';
            };

            dot.onmouseout = () => tooltipPos.style.opacity = 0;
            
            cell.appendChild(dot);
        });

        wrapperPos.appendChild(cell);
    });
}

selectorPos.addEventListener('change', renderMapPos);

// ==================== NEGATIVE MODE (RIGHT) ====================
const wrapperNeg = document.getElementById('mapWrapperNeg');
const tooltipNeg = document.getElementById('tooltipNeg');
const selectorNeg = document.getElementById('dataSelectorNeg');
const fileInputNeg = document.getElementById('excelUploadNeg');
const statusNeg = document.getElementById('statusNeg');
const targetInputNeg = document.getElementById('fileCountTargetNeg');

let globalDataNeg = [];
let myChartNeg = null;

fileInputNeg.addEventListener('change', async (e) => {
    const files = Array.from(e.target.files);
    const targetCount = parseInt(targetInputNeg.value);

    if (files.length !== targetCount) {
        statusNeg.innerText = `❌ Error: Please upload exactly ${targetCount} files.`;
        return;
    }

    statusNeg.innerText = "Processing files...";
    
    try {
        const allFilesData = [];
        for (const file of files) {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            allFilesData.push(XLSX.utils.sheet_to_json(firstSheet));
        }

        // Filter data for STATION = "Negative"
        const negativeData = allFilesData.map(fileData => 
            fileData.filter(row => row.STATION === "Negative")
        );

        // Check if negative data exists
        const hasNegativeData = negativeData.some(data => data.length > 0);
        
        if (!hasNegativeData) {
            statusNeg.innerText = `❌ Error: No "Negative" station data found in uploaded files.`;
            wrapperNeg.innerHTML = '<div style="padding: 20px; text-align: center; color: #ef4444;">No Negative Station Data</div>';
            return;
        }

        globalDataNeg = calculateWeldAveragesWithHistory(negativeData);
        statusNeg.innerText = `✓ Loaded ${files.length} file(s) - Negative data (${globalDataNeg.length} records)`;
        renderMapNeg();
    } catch (err) {
        statusNeg.innerText = "❌ Error processing files.";
        console.error(err);
    }
});

function renderMapNeg() {
    if (globalDataNeg.length === 0) return;
    const selectedParam = selectorNeg.value;
    wrapperNeg.innerHTML = '';

    const getFieldSuffix = (num) => {
        const map = {
            'AF': `STNLOADVALUEAFTERFORCER2CELL${num}`,
            'BF': `STNLOADVALUEBEFOREFORCER2CELL${num}`,
            'Resistance': `STNR2RESISTANCE${num}mohms`,
            'PeakCurrent': `STNR2PEAKCURRENT${num}KA`,
            'AverageCurrent': `STNR2AVERAGECURRENT${num}KA`,
            'PeakVoltage': `STNR2PEAKVOLTAGE${num}V`,
            'AverageVoltage': `STNR2AVERAGEVOLTAGE${num}V`,
            'Power': `STNR2POWER${num}KW`
        };
        return map[selectedParam];
    };

    let allValues = [];
    globalDataNeg.forEach(entry => {
        allValues.push(parseFloat(entry[getFieldSuffix(1)]) || 0, parseFloat(entry[getFieldSuffix(2)]) || 0);
    });
    
    const min = Math.min(...allValues);
    const max = Math.max(...allValues);
    document.getElementById('maxLabelNeg').innerText = max.toFixed(2);
    document.getElementById('minLabelNeg').innerText = min.toFixed(2);

    // 16 columns × 11 rows = 176 cells
    globalDataNeg.forEach((entry, index) => {
        if (index >= 176) return;
        const col = index % 16;
        const row = Math.floor(index / 16);
        const letter = rowLetters[row];

        const cell = document.createElement('div');
        cell.className = 'weld-cell';

        const weldSet = [
            { type: 'L1', val: parseFloat(entry[getFieldSuffix(1)]) || 0, label: 'L1', history: entry[`${getFieldSuffix(1)}_history`] },
            { type: 'L2', val: parseFloat(entry[getFieldSuffix(2)]) || 0, label: 'L2', history: entry[`${getFieldSuffix(2)}_history`] }
        ];

        weldSet.forEach(w => {
            const dot = document.createElement('div');
            dot.className = `weld weld-${w.type}`;
            dot.id = `weld-neg-${letter}-${col + 1}-${w.type}`;
            dot.style.backgroundColor = getColor(w.val, min, max);

            dot.onmouseover = (e) => {
                tooltipNeg.style.opacity = 1;
                
                document.getElementById('tooltip-text-neg').innerHTML = `
                    <div style="margin-bottom: 8px; border-bottom: 1px solid #334155; padding-bottom: 4px;">
                        <strong>${w.label} - ${letter}${col+1}</strong>
                    </div>
                    Avg ${selectedParam}: <strong>${w.val.toFixed(4)}</strong>
                `;

                const ctx = document.getElementById('weldChartNeg').getContext('2d');
                if (myChartNeg) myChartNeg.destroy();

                myChartNeg = new Chart(ctx, {
                    type: 'line',
                    data: {
                        labels: w.history.map((_, i) => `F${i + 1}`),
                        datasets: [{
                            label: selectedParam,
                            data: w.history,
                            borderColor: '#38bdf8',
                            backgroundColor: 'rgba(56, 189, 248, 0.1)',
                            borderWidth: 2,
                            tension: 0.3,
                            fill: true,
                            pointRadius: 4,
                            pointBackgroundColor: '#38bdf8'
                        }]
                    },
                    options: {
                        responsive: false,
                        animation: { duration: 300 },
                        plugins: {
                            legend: { display: false },
                            tooltip: { enabled: true }
                        },
                        scales: {
                            y: {
                                grid: { color: 'rgba(255,255,255,0.1)' },
                                ticks: { color: '#94a3b8', font: { size: 9 } }
                            },
                            x: {
                                grid: { display: false },
                                ticks: { color: '#94a3b8', font: { size: 9 } }
                            }
                        }
                    }
                });
            };

            dot.onmousemove = (e) => {
                tooltipNeg.style.left = (e.pageX + 15) + 'px';
                tooltipNeg.style.top = (e.pageY - 50) + 'px';
            };

            dot.onmouseout = () => {
                tooltipNeg.style.opacity = 0;
            };
            
            cell.appendChild(dot);
        });

        wrapperNeg.appendChild(cell);
    });
}

selectorNeg.addEventListener('change', renderMapNeg);