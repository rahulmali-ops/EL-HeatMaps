const wrapper = document.getElementById('mapWrapper');
const colHeadersContainer = document.getElementById('colHeaders');
const rowLabelsContainer = document.getElementById('rowLabels');
const tooltip = document.getElementById('tooltip');
const tooltipText = document.getElementById('tooltip-text');
const selector = document.getElementById('dataSelector');
const fileInput = document.getElementById('excelUpload');
const statusText = document.getElementById('status');
const targetInput = document.getElementById('fileCountTarget');
const statsContainer = document.getElementById('statsContainer');
const statsPanelDiv = document.getElementById('statsPanel');

let globalData = [];
let myChart = null;
let rawCombinedData = [];
const POSITION_COLUMNS = 16;
const POSITION_ROWS = 11;

/**
 * Convert the physical position shown in image (4) to the weld point number.
 * Position column 1 is on the right side, 16 is on the left side.
 * Position row 1 is at the bottom, 11 is at the top.
 */
function gridPositionToWeldPoint(positionColumn, positionRow) {
    const baseValue = 171 - (positionColumn - 1) * POSITION_ROWS;
    return positionRow % 2 === 1
        ? baseValue + ((positionRow - 1) / 2)
        : baseValue - (positionRow / 2);
}

fileInput.addEventListener('change', async (e) => {
    const files = Array.from(e.target.files);
    const targetCount = parseInt(targetInput.value);

    if (files.length !== targetCount) {
        statusText.innerText = `❌ Error: Please upload exactly ${targetCount} files. (Selected: ${files.length})`;
        return;
    }

    statusText.innerText = "⏳ Processing files and combining all batteries...";

    try {
        const allFilesData = [];
        
        for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
            const file = files[fileIndex];
            console.log(`Processing file ${fileIndex + 1}: ${file.name}`);
            
            try {
                const data = await file.arrayBuffer();
                const workbook = XLSX.read(data, { type: 'array' });
                
                if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                    throw new Error(`File "${file.name}" has no sheets`);
                }
                
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const sheetData = XLSX.utils.sheet_to_json(firstSheet);
                
                if (!sheetData || sheetData.length === 0) {
                    throw new Error(`File "${file.name}" has no data in first sheet`);
                }
                
                console.log(`File ${fileIndex + 1} loaded: ${sheetData.length} rows`);
                allFilesData.push(sheetData);
                
            } catch (fileError) {
                console.error(`Error reading file ${fileIndex + 1}:`, fileError);
                statusText.innerText = `❌ Error reading file "${file.name}": ${fileError.message}`;
                return;
            }
        }

        const combinedData = [];
        allFilesData.forEach((fileData, index) => {
            console.log(`Combining file ${index + 1}: ${fileData.length} records`);
            combinedData.push(...fileData);
        });

        rawCombinedData = combinedData;
        console.log(`Total combined records: ${combinedData.length}`);

        // Validate data structure
        if (combinedData.length > 0) {
            const firstRow = combinedData[0];
            console.log('First row keys:', Object.keys(firstRow));
            
            // Check for required columns
            const requiredColumns = ['STATIONWELDPOINT', 'STNBINNUMBER', 'STNR1PEAKCURRENT1KA'];
            const missingColumns = requiredColumns.filter(col => !(col in firstRow));
            
            if (missingColumns.length > 0) {
                throw new Error(`Missing required columns: ${missingColumns.join(', ')}`);
            }
        }

        const binNumbers = [...new Set(combinedData.map(row => row.STNBINNUMBER))];
        console.log(`Unique batteries: ${binNumbers.length}`, binNumbers);

        globalData = calculateWeldAverages(combinedData);
        console.log(`Averaged data points: ${Object.keys(globalData).filter(k => globalData[k] !== null).length}`);

        statusText.innerText = `✅ Loaded ${files.length} file(s) | Combined ${binNumbers.length} battery(ies) | Total ${Object.values(globalData).filter(d => d !== null).length} records`;
        renderMap();
        updateStatistics();
        statsContainer.style.display = 'block';
        
    } catch (err) {
        console.error('Error processing files:', err);
        statusText.innerText = `❌ Error processing files: ${err.message}`;
    }
});

function calculateWeldAverages(combinedData) {
    const positionMap = {};

    combinedData.forEach((row, index) => {
        const weldPoint = parseInt(row.STATIONWELDPOINT);
        
        if (isNaN(weldPoint) || weldPoint < 1 || weldPoint > 176) {
            console.warn(`Row ${index}: Invalid weld point: ${row.STATIONWELDPOINT}`);
            return;
        }

        if (!positionMap[weldPoint]) {
            positionMap[weldPoint] = [];
        }
        positionMap[weldPoint].push(row);
    });

    console.log(`Position map created with ${Object.keys(positionMap).length} unique positions`);

    const averagedData = {};

    for (let weldPoint = 1; weldPoint <= 176; weldPoint++) {
        const rowsAtPosition = positionMap[weldPoint] || [];

        if (rowsAtPosition.length === 0) {
            averagedData[weldPoint] = null;
            continue;
        }

        let averagedRow = { ...rowsAtPosition[0], STATIONWELDPOINT: weldPoint };

        const params = [
            { name: 'PeakCurrent', suffix1: 'STNR1PEAKCURRENT1KA', suffix2: 'STNR1PEAKCURRENT2KA' },
            { name: 'AverageCurrent', suffix1: 'STNR1AVERAGECURRENT1KA', suffix2: 'STNR1AVERAGECURRENT2KA' },
            { name: 'PeakVoltage', suffix1: 'STNR1PEAKVOLTAGE1V', suffix2: 'STNR1PEAKVOLTAGE2V' },
            { name: 'AverageVoltage', suffix1: 'STNR1AVERAGEVOLTAGE1KA', suffix2: 'STNR1AVERAGEVOLTAGE2KA' },
            { name: 'Power', suffix1: 'STNR1POWER1KW', suffix2: 'STNR1POWER2KW' },
            { name: 'Resistance', suffix1: 'STNR1RESISTANCE1mohms', suffix2: 'STNR1RESISTANCE2mohms' },
            { name: 'AF', suffix1: 'STNLOADVALUEAFTERFORCER1CELL1', suffix2: 'STNLOADVALUEAFTERFORCER1CELL2' },
            { name: 'BF', suffix1: 'STNLOADVALUEBEFOREFORCER1CELL1', suffix2: 'STNLOADVALUEBEFOREFORCER1CELL2' }
        ];

        params.forEach(param => {
            [param.suffix1, param.suffix2].forEach(key => {
                let sum = 0;
                let history = [];
                let batteryLabels = [];
                let validCount = 0;

                rowsAtPosition.forEach(row => {
                    const raw = row[key];
                    const val = raw !== undefined && raw !== null && raw !== "" ? parseFloat(raw) : NaN;
                    const binNumber = row.STNBINNUMBER || 'Unknown';
                    
                    if (!Number.isNaN(val) && isFinite(val)) {
                        sum += val;
                        history.push(val);
                        batteryLabels.push(binNumber);
                        validCount++;
                    }
                });

                averagedRow[key] = validCount > 0 ? (sum / validCount) : 0;
                averagedRow[`${key}_history`] = history;
                averagedRow[`${key}_labels`] = batteryLabels;
            });
        });

        averagedData[weldPoint] = averagedRow;
    }

    return averagedData;
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

function renderMap() {
    if (!globalData || Object.keys(globalData).length === 0) return;
    const selectedParam = selector.value;
    wrapper.innerHTML = '';
    colHeadersContainer.innerHTML = '';
    rowLabelsContainer.innerHTML = '';

    const getFields = (param) => {
        const map = {
            'PeakCurrent': { field1: 'STNR1PEAKCURRENT1KA', field2: 'STNR1PEAKCURRENT2KA' },
            'AverageCurrent': { field1: 'STNR1AVERAGECURRENT1KA', field2: 'STNR1AVERAGECURRENT2KA' },
            'PeakVoltage': { field1: 'STNR1PEAKVOLTAGE1V', field2: 'STNR1PEAKVOLTAGE2V' },
            'AverageVoltage': { field1: 'STNR1AVERAGEVOLTAGE1KA', field2: 'STNR1AVERAGEVOLTAGE2KA' },
            'Power': { field1: 'STNR1POWER1KW', field2: 'STNR1POWER2KW' },
            'Resistance': { field1: 'STNR1RESISTANCE1mohms', field2: 'STNR1RESISTANCE2mohms' },
            'AF': { field1: 'STNLOADVALUEAFTERFORCER1CELL1', field2: 'STNLOADVALUEAFTERFORCER1CELL2' },
            'BF': { field1: 'STNLOADVALUEBEFOREFORCER1CELL1', field2: 'STNLOADVALUEBEFOREFORCER1CELL2' }
        };
        return map[param];
    };

    const fields = getFields(selectedParam);

    let allValues = [];
    Object.values(globalData).forEach(entry => {
        if (!entry) return;
        const a = parseFloat(entry[fields.field1]);
        const b = parseFloat(entry[fields.field2]);
        if (!Number.isNaN(a) && isFinite(a) && a !== 0) allValues.push(a);
        if (!Number.isNaN(b) && isFinite(b) && b !== 0) allValues.push(b);
    });

    const min = allValues.length ? Math.min(...allValues) : 0;
    const max = allValues.length ? Math.max(...allValues) : 0;
    document.getElementById('maxLabel').innerText = allValues.length ? max.toFixed(2) : 'N/A';
    document.getElementById('minLabel').innerText = allValues.length ? min.toFixed(2) : 'N/A';

    // Position columns: left side of the image is 16, right side is 1.
    for (let positionColumn = POSITION_COLUMNS; positionColumn >= 1; positionColumn--) {
        const header = document.createElement('div');
        header.className = 'col-header';
        header.innerText = positionColumn;
        colHeadersContainer.appendChild(header);
    }

    // Position rows: top of the image is 11, bottom is 1.
    for (let positionRow = POSITION_ROWS; positionRow >= 1; positionRow--) {
        const label = document.createElement('div');
        label.className = 'row-label';
        label.innerText = positionRow;
        rowLabelsContainer.appendChild(label);
    }

    // Create cells in the same visual order as image (4).
    for (let positionRow = POSITION_ROWS; positionRow >= 1; positionRow--) {
        for (let positionColumn = POSITION_COLUMNS; positionColumn >= 1; positionColumn--) {
            const weldPoint = gridPositionToWeldPoint(positionColumn, positionRow);
            const entry = globalData[weldPoint];

            const cell = document.createElement('div');
            cell.className = 'weld-cell';

            if (entry && entry !== null) {
                const weldSet = [
                    { 
                        type: 'L1', 
                        val: parseFloat(entry[fields.field1]) || 0, 
                        label: 'L1', 
                        history: entry[`${fields.field1}_history`] || [],
                        batteryLabels: entry[`${fields.field1}_labels`] || []
                    },
                    { 
                        type: 'L2', 
                        val: parseFloat(entry[fields.field2]) || 0, 
                        label: 'L2', 
                        history: entry[`${fields.field2}_history`] || [],
                        batteryLabels: entry[`${fields.field2}_labels`] || []
                    }
                ];

                weldSet.forEach(w => {
                    const dot = document.createElement('div');
                    dot.className = `weld weld-${w.type}`;
                    dot.id = `weld-${positionColumn}-${positionRow}-${w.type}`;
                    dot.style.backgroundColor = getColor(w.val, min, max);

                    dot.onmouseover = (e) => {
                        showEnhancedTooltip(e, w, positionColumn, positionRow, weldPoint, selectedParam);
                    };
                    
                    dot.onmousemove = (e) => {
                        tooltip.style.left = '50%';
                        tooltip.style.top = '50%';
                        tooltip.style.transform = 'translate(-50%, -50%)';
                    };
                    
                    dot.onmouseout = () => {
                        tooltip.style.opacity = 0;
                        if (myChart) { 
                            myChart.destroy(); 
                            myChart = null; 
                        }
                    };
                    
                    cell.appendChild(dot);
                });
            } else {
                cell.innerText = '';
                cell.style.backgroundColor = '#1a1a1a';
                cell.style.border = '1px dashed #333';
            }

            wrapper.appendChild(cell);
        }
    }
}

function showEnhancedTooltip(e, weldData, positionColumn, positionRow, weldPoint, selectedParam) {
    tooltip.style.opacity = 1;
    tooltip.style.position = 'fixed';
    tooltip.style.left = '50%';
    tooltip.style.top = '50%';
    tooltip.style.transform = 'translate(-50%, -50%)';
    tooltip.style.width = '80vw';
    tooltip.style.maxWidth = '1200px';
    tooltip.style.height = '70vh';
    tooltip.style.maxHeight = '800px';
    tooltip.style.zIndex = '10000';
    
    const hist = Array.isArray(weldData.history) ? weldData.history : [];
    const labels = Array.isArray(weldData.batteryLabels) ? weldData.batteryLabels : [];
    
    // Calculate statistics
    const validData = hist.filter(v => !Number.isNaN(v) && isFinite(v) && v !== 0);
    const mean = validData.length > 0 ? validData.reduce((a, b) => a + b, 0) / validData.length : 0;
    const minVal = validData.length > 0 ? Math.min(...validData) : 0;
    const maxVal = validData.length > 0 ? Math.max(...validData) : 0;
    const stdDev = validData.length > 0 ? Math.sqrt(validData.reduce((sq, n) => sq + Math.pow(n - mean, 2), 0) / validData.length) : 0;
    
    tooltipText.innerHTML = `
        <div style="padding: 20px; height: 100%; display: flex; flex-direction: column;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; border-bottom: 2px solid #28a745; padding-bottom: 10px;">
                <div>
                    <h2 style="margin: 0; color: #28a745; font-size: 24px;">Position: ${positionColumn}/${positionRow} (WP: ${weldPoint}) | Weld: ${weldData.label}</h2>
                    <p style="margin: 5px 0 0 0; color: #94a3b8; font-size: 14px;">${selectedParam} Analysis</p>
                </div>
                <div style="text-align: right;">
                    <div style="font-size: 14px; color: #94a3b8;">Batteries: <strong style="color: #28a745;">${validData.length}</strong></div>
                    <div style="font-size: 14px; color: #94a3b8;">Average: <strong style="color: #28a745;">${mean.toFixed(4)}</strong></div>
                </div>
            </div>
            
            <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 15px;">
                <div style="background: rgba(40, 167, 69, 0.1); padding: 10px; border-radius: 6px; border-left: 3px solid #28a745;">
                    <div style="font-size: 11px; color: #94a3b8;">Mean</div>
                    <div style="font-size: 18px; font-weight: bold; color: #28a745;">${mean.toFixed(4)}</div>
                </div>
                <div style="background: rgba(32, 201, 151, 0.1); padding: 10px; border-radius: 6px; border-left: 3px solid #20c997;">
                    <div style="font-size: 11px; color: #94a3b8;">Std Dev</div>
                    <div style="font-size: 18px; font-weight: bold; color: #20c997;">${stdDev.toFixed(4)}</div>
                </div>
                <div style="background: rgba(13, 110, 253, 0.1); padding: 10px; border-radius: 6px; border-left: 3px solid #0d6efd;">
                    <div style="font-size: 11px; color: #94a3b8;">Min</div>
                    <div style="font-size: 18px; font-weight: bold; color: #0d6efd;">${minVal.toFixed(4)}</div>
                </div>
                <div style="background: rgba(220, 53, 69, 0.1); padding: 10px; border-radius: 6px; border-left: 3px solid #dc3545;">
                    <div style="font-size: 11px; color: #94a3b8;">Max</div>
                    <div style="font-size: 18px; font-weight: bold; color: #dc3545;">${maxVal.toFixed(4)}</div>
                </div>
            </div>
            
            <div style="flex: 1; position: relative; min-height: 400px;">
                <canvas id="weldChart"></canvas>
            </div>
        </div>
    `;

    const ctx = document.getElementById('weldChart').getContext('2d');
    if (myChart) myChart.destroy();

    if (validData.length > 0) {
        myChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: labels.length > 0 ? labels : hist.map((_, i) => `Battery ${i + 1}`),
                datasets: [{
                    label: selectedParam,
                    data: hist,
                    borderColor: '#28a745',
                    backgroundColor: 'rgba(40, 167, 69, 0.1)',
                    borderWidth: 3,
                    tension: 0.4,
                    fill: true,
                    pointRadius: 6,
                    pointHoverRadius: 8,
                    pointBackgroundColor: '#28a745',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    pointHoverBackgroundColor: '#fff',
                    pointHoverBorderColor: '#28a745',
                    pointHoverBorderWidth: 3
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: { duration: 500 },
                plugins: {
                    legend: { 
                        display: true,
                        position: 'top',
                        labels: {
                            color: '#f8fafc',
                            font: { size: 14, weight: 'bold' },
                            padding: 15
                        }
                    },
                    tooltip: { 
                        enabled: true,
                        backgroundColor: 'rgba(15, 23, 42, 0.95)',
                        titleColor: '#28a745',
                        bodyColor: '#f8fafc',
                        borderColor: '#28a745',
                        borderWidth: 2,
                        padding: 12,
                        displayColors: true,
                        callbacks: {
                            title: function(context) {
                                return `Battery: ${context[0].label}`;
                            },
                            label: function(context) {
                                return `${selectedParam}: ${context.parsed.y.toFixed(4)}`;
                            },
                            afterLabel: function(context) {
                                const diff = context.parsed.y - mean;
                                const sign = diff >= 0 ? '+' : '';
                                return `Deviation: ${sign}${diff.toFixed(4)} (${sign}${((diff/mean)*100).toFixed(2)}%)`;
                            }
                        }
                    }
                },
                scales: {
                    y: { 
                        grid: { 
                            color: 'rgba(255,255,255,0.1)',
                            lineWidth: 1
                        },
                        border: {
                            color: '#94a3b8',
                            width: 2
                        },
                        ticks: { 
                            color: '#f8fafc', 
                            font: { size: 14, weight: 'bold' },
                            padding: 10,
                            callback: function(value) {
                                return value.toFixed(3);
                            }
                        },
                        title: {
                            display: true,
                            text: selectedParam,
                            color: '#28a745',
                            font: { size: 16, weight: 'bold' },
                            padding: 10
                        }
                    },
                    x: { 
                        grid: { 
                            color: 'rgba(255,255,255,0.05)',
                            lineWidth: 1
                        },
                        border: {
                            color: '#94a3b8',
                            width: 2
                        },
                        ticks: { 
                            color: '#f8fafc', 
                            font: { size: 12, weight: 'bold' },
                            padding: 8,
                            maxRotation: 45,
                            minRotation: 45
                        },
                        title: {
                            display: true,
                            text: 'Battery ID',
                            color: '#28a745',
                            font: { size: 16, weight: 'bold' },
                            padding: 10
                        }
                    }
                },
                interaction: {
                    intersect: false,
                    mode: 'index'
                }
            },
            plugins: [{
                id: 'chartAreaBorder',
                beforeDraw(chart) {
                    const { ctx, chartArea: { left, top, width, height } } = chart;
                    ctx.save();
                    ctx.strokeStyle = '#28a745';
                    ctx.lineWidth = 3;
                    ctx.strokeRect(left, top, width, height);
                    ctx.restore();
                }
            }]
        });
    } else {
        ctx.fillStyle = '#94a3b8';
        ctx.font = '20px Inter';
        ctx.textAlign = 'center';
        ctx.fillText('No valid data available', ctx.canvas.width / 2, ctx.canvas.height / 2);
    }
}

function updateStatistics() {
    const selectedParam = selector.value;
    const paramData = extractParameterData(rawCombinedData, selectedParam);
    const stats = calculateStatistics(paramData);
    displayAdvancedStatistics(stats, selectedParam);
}

function extractParameterData(data, param) {
    const values = [];
    
    const fieldMap = {
        'PeakCurrent': ['STNR1PEAKCURRENT1KA', 'STNR1PEAKCURRENT2KA'],
        'AverageCurrent': ['STNR1AVERAGECURRENT1KA', 'STNR1AVERAGECURRENT2KA'],
        'PeakVoltage': ['STNR1PEAKVOLTAGE1V', 'STNR1PEAKVOLTAGE2V'],
        'AverageVoltage': ['STNR1AVERAGEVOLTAGE1KA', 'STNR1AVERAGEVOLTAGE2KA'],
        'Power': ['STNR1POWER1KW', 'STNR1POWER2KW'],
        'Resistance': ['STNR1RESISTANCE1mohms', 'STNR1RESISTANCE2mohms'],
        'AF': ['STNLOADVALUEAFTERFORCER1CELL1', 'STNLOADVALUEAFTERFORCER1CELL2'],
        'BF': ['STNLOADVALUEBEFOREFORCER1CELL1', 'STNLOADVALUEBEFOREFORCER1CELL2']
    };

    const fields = fieldMap[param] || [];

    data.forEach(row => {
        fields.forEach(field => {
            const val = parseFloat(row[field]);
            if (!Number.isNaN(val) && isFinite(val) && val !== 0) {
                values.push(val);
            }
        });
    });

    return values.sort((a, b) => a - b);
}

function displayAdvancedStatistics(stats, paramName) {
    let html = `
        <h2 style="margin-top: 0; color: #28a745; margin-bottom: 20px;">📊 Advanced Statistics - ${paramName}</h2>
        
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px;">
    `;

    html += createStatCard('Mean', stats.mean.toFixed(4), '#28a745');
    html += createStatCard('Median', stats.median.toFixed(4), '#20c997');
    html += createStatCard('Mode', stats.mode.toFixed(4), '#0dcaf0');
    html += createStatCard('Std Dev', stats.stdDev.toFixed(4), '#0d6efd');
    html += createStatCard('Variance', stats.variance.toFixed(4), '#6f42c1');
    html += createStatCard('Range', stats.range.toFixed(4), '#fd7e14');
    html += createStatCard('Min', stats.min.toFixed(4), '#dc3545');
    html += createStatCard('Max', stats.max.toFixed(4), '#dc3545');
    html += createStatCard('Q1', stats.q1.toFixed(4), '#6c757d');
    html += createStatCard('Q3', stats.q3.toFixed(4), '#6c757d');
    html += createStatCard('IQR', stats.iqr.toFixed(4), '#6c757d');
    html += createStatCard('Skewness', stats.skewness.toFixed(4), '#17a2b8');
    html += createStatCard('Kurtosis', stats.kurtosis.toFixed(4), '#17a2b8');
    html += createStatCard('CV', (stats.cv * 100).toFixed(2) + '%', '#198754');
    html += createStatCard('Count', stats.count, '#6c757d');

    html += `
        </div>
        
        <div style="margin-top: 20px; display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 15px;">
            <div style="padding: 15px; background: rgba(40, 167, 69, 0.1); border-left: 4px solid #28a745; border-radius: 6px;">
                <h3 style="margin-top: 0; color: #28a745; font-size: 14px;">📈 Distribution Analysis</h3>
                <p style="margin: 5px 0; font-size: 12px;"><strong>Type:</strong> ${stats.distributionType}</p>
                <p style="margin: 5px 0; font-size: 12px;"><strong>Outliers:</strong> ${stats.outliers.count} (${stats.outliers.percentage.toFixed(2)}%)</p>
                <p style="margin: 5px 0; font-size: 12px;"><strong>Quality:</strong> ${stats.dataQuality}</p>
            </div>

            <div style="padding: 15px; background: rgba(32, 201, 151, 0.1); border-left: 4px solid #20c997; border-radius: 6px;">
                <h3 style="margin-top: 0; color: #20c997; font-size: 14px;">✅ Summary</h3>
                <p style="margin: 5px 0; font-size: 12px;"><strong>MAD:</strong> ${stats.meanAbsoluteDeviation.toFixed(4)}</p>
                <p style="margin: 5px 0; font-size: 12px;"><strong>Std Error:</strong> ${stats.standardError.toFixed(4)}</p>
                <p style="margin: 5px 0; font-size: 12px;"><strong>Span:</strong> ${stats.min.toFixed(4)} → ${stats.max.toFixed(4)}</p>
            </div>
        </div>
    `;

    statsPanelDiv.innerHTML = html;
}

function createStatCard(label, value, color) {
    return `
        <div style="
            padding: 12px;
            background: rgba(0, 0, 0, 0.2);
            border: 1px solid ${color}33;
            border-radius: 6px;
            border-left: 3px solid ${color};
            text-align: center;
        ">
            <div style="font-size: 11px; color: #94a3b8; margin-bottom: 4px;">${label}</div>
            <div style="font-size: 16px; font-weight: bold; color: ${color};">${value}</div>
        </div>
    `;
}

function calculateStatistics(values) {
    if (values.length === 0) return {
        count: 0, mean: 0, median: 0, mode: 0, stdDev: 0, variance: 0,
        min: 0, max: 0, range: 0, q1: 0, q3: 0, iqr: 0, skewness: 0,
        kurtosis: 0, cv: 0, outliers: { count: 0, percentage: 0 },
        meanAbsoluteDeviation: 0, standardError: 0,
        distributionType: 'N/A', dataQuality: 'N/A'
    };

    const count = values.length;
    const sum = values.reduce((a, b) => a + b, 0);
    const mean = sum / count;

    const median = values.length % 2 === 0 
        ? (values[values.length / 2 - 1] + values[values.length / 2]) / 2 
        : values[Math.floor(values.length / 2)];

    const frequency = {};
    values.forEach(val => {
        frequency[val] = (frequency[val] || 0) + 1;
    });
    const mode = parseFloat(Object.keys(frequency).reduce((a, b) => 
        frequency[a] > frequency[b] ? a : b
    ));

    const squaredDiffs = values.map(val => Math.pow(val - mean, 2));
    const variance = squaredDiffs.reduce((a, b) => a + b, 0) / count;
    const stdDev = Math.sqrt(variance);

    const min = Math.min(...values);
    const max = Math.max(...values);
    const range = max - min;

    const q1Index = Math.floor(count * 0.25);
    const q3Index = Math.floor(count * 0.75);
    const q1 = values[q1Index];
    const q3 = values[q3Index];
    const iqr = q3 - q1;

    const lowerBound = q1 - 1.5 * iqr;
    const upperBound = q3 + 1.5 * iqr;
    const outliers = values.filter(v => v < lowerBound || v > upperBound);
    const outliersPercentage = (outliers.length / count) * 100;

    const cubedDiffs = values.map(val => Math.pow(val - mean, 3));
    const skewness = (cubedDiffs.reduce((a, b) => a + b, 0) / count) / Math.pow(stdDev, 3);

    const fourthDiffs = values.map(val => Math.pow(val - mean, 4));
    const kurtosis = ((fourthDiffs.reduce((a, b) => a + b, 0) / count) / Math.pow(stdDev, 4)) - 3;

    const cv = stdDev / mean;

    const mad = values.reduce((sum, val) => sum + Math.abs(val - mean), 0) / count;

    const standardError = stdDev / Math.sqrt(count);

    let distributionType = 'Normal';
    if (Math.abs(skewness) > 1) distributionType = 'Highly Skewed';
    else if (Math.abs(skewness) > 0.5) distributionType = 'Moderately Skewed';
    if (kurtosis > 1) distributionType += ' (Heavy Tails)';
    else if (kurtosis < -1) distributionType += ' (Light Tails)';

    let dataQuality = '✅ Excellent';
    if (outliersPercentage > 10) dataQuality = '⚠️ Fair (High Outliers)';
    else if (outliersPercentage > 5) dataQuality = '⚠️ Good (Some Outliers)';

    return {
        count,
        mean,
        median,
        mode,
        stdDev,
        variance,
        min,
        max,
        range,
        q1,
        q3,
        iqr,
        skewness,
        kurtosis,
        cv,
        outliers: { count: outliers.length, percentage: outliersPercentage },
        meanAbsoluteDeviation: mad,
        standardError,
        distributionType,
        dataQuality
    };
}

selector.addEventListener('change', () => {
    renderMap();
    updateStatistics();
});
