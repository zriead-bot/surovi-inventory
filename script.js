// script.js - Surovi Agro Industries Stock Dashboard
// Complete script with enhanced Excel parsing and export functionality

document.addEventListener('DOMContentLoaded', function() {
    // Set current year in footer
    document.getElementById('currentYear').textContent = new Date().getFullYear();
    
    // DOM Elements
    const fileInput = document.getElementById('fileInput');
    const browseButton = document.getElementById('browseButton');
    const dropArea = document.getElementById('dropArea');
    const fileInfo = document.getElementById('fileInfo');
    const fileName = document.getElementById('fileName');
    const fileSize = document.getElementById('fileSize');
    const reportDate = document.getElementById('reportDate');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const exportButton = document.getElementById('exportButton');
    const tableBody = document.getElementById('tableBody');
    const stockTable = document.getElementById('stockTable');
    const tablePlaceholder = document.getElementById('tablePlaceholder');
    const productCount = document.getElementById('productCount');
    
    // Stats elements
    const totalProducts = document.getElementById('totalProducts');
    const totalUnits = document.getElementById('totalUnits');
    const lowStockItems = document.getElementById('lowStockItems');
    const depotCount = document.getElementById('depotCount');
    
    // Control elements
    const lowStockThreshold = document.getElementById('lowStockThreshold');
    const locationFilter = document.getElementById('locationFilter');
    const productFilter = document.getElementById('productFilter');
    
    // Data storage
    let stockData = [];
    let filteredData = [];
    let currentSortColumn = 'product';
    let currentSortDirection = 'asc';
    let currentFileName = '';
    
    // Event Listeners
    browseButton.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);
    exportButton.addEventListener('click', exportToExcel);
    
    lowStockThreshold.addEventListener('input', applyFilters);
    locationFilter.addEventListener('change', applyFilters);
    productFilter.addEventListener('input', applyFilters);
    
    // Drag and drop functionality
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, unhighlight, false);
    });
    
    function highlight() {
        dropArea.style.backgroundColor = '#e8f5e9';
        dropArea.style.borderColor = '#388E3C';
    }
    
    function unhighlight() {
        dropArea.style.backgroundColor = '#f9f9f9';
        dropArea.style.borderColor = '#4CAF50';
    }
    
    dropArea.addEventListener('drop', handleDrop, false);
    
    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length > 0) {
            fileInput.files = files;
            handleFileSelect({ target: fileInput });
        }
    }
    
    // Table sorting
    document.querySelectorAll('th[data-sort]').forEach(th => {
        th.addEventListener('click', () => {
            const column = th.getAttribute('data-sort');
            
            if (currentSortColumn === column) {
                currentSortDirection = currentSortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                currentSortColumn = column;
                currentSortDirection = 'asc';
            }
            
            sortTable();
            renderTable();
            
            // Update sort indicators
            document.querySelectorAll('th i').forEach(icon => {
                icon.className = 'fas fa-sort';
            });
            
            const sortIcon = th.querySelector('i');
            if (sortIcon) {
                sortIcon.className = currentSortDirection === 'asc' 
                    ? 'fas fa-sort-up' 
                    : 'fas fa-sort-down';
            }
        });
    });
    
    // File handling
    function handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        currentFileName = file.name;
        
        // Show file info
        fileName.textContent = file.name;
        fileSize.textContent = formatFileSize(file.size);
        
        // Extract date from filename
        const dateMatch = file.name.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\.\s-]*(\d{4})/i);
        if (dateMatch) {
            const month = dateMatch[1];
            const year = dateMatch[2];
            reportDate.textContent = `${month} ${year}`;
        } else {
            // Try to extract from any date pattern
            const anyDateMatch = file.name.match(/\d{2}[\.\-\/]\d{2}[\.\-\/]\d{4}/) || 
                               file.name.match(/\d{4}[\.\-\/]\d{2}[\.\-\/]\d{2}/);
            if (anyDateMatch) {
                reportDate.textContent = anyDateMatch[0];
            } else {
                reportDate.textContent = 'Not detected';
            }
        }
        
        fileInfo.classList.add('show');
        
        // Show loading indicator
        loadingIndicator.classList.add('show');
        
        // Process the file
        processExcelFile(file);
    }
    
    function processExcelFile(file) {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: 'binary' });
                
                // Get the first sheet (assuming data is in the first sheet)
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert to JSON with raw values (to get formulas)
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
                
                console.log('Excel data loaded, rows:', jsonData.length);
                
                // Parse the data based on Surovi Excel structure
                parseStockData(jsonData);
                
                // Hide loading indicator
                setTimeout(() => {
                    loadingIndicator.classList.remove('show');
                }, 500);
                
                // Enable export button
                exportButton.disabled = false;
                
                // Apply initial filters and render
                applyFilters();
                
            } catch (error) {
                console.error('Error processing Excel file:', error);
                alert('Error processing the Excel file. Please make sure it follows the Surovi stock format.');
                loadingIndicator.classList.remove('show');
            }
        };
        
        reader.onerror = function() {
            alert('Error reading the file. Please try again.');
            loadingIndicator.classList.remove('show');
        };
        
        reader.readAsBinaryString(file);
    }
    
    // Parse Surovi Excel Stock Data
    function parseStockData(data) {
        stockData = [];
        
        console.log("Parsing Excel data with", data.length, "rows");
        
        // First, try to find the data table by looking for column headers
        let dataStartRow = findDataStartRow(data);
        
        if (dataStartRow === -1) {
            // Fallback: start from row 6 (0-indexed) as in original structure
            dataStartRow = 6;
            console.log("Using default start row:", dataStartRow);
        }
        
        // Parse each row from the starting point
        for (let i = dataStartRow; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length < 10) continue;
            
            // Try to extract product information
            const productInfo = extractProductInfo(row, i);
            
            if (productInfo) {
                stockData.push(productInfo);
            }
        }
        
        console.log(`Successfully parsed ${stockData.length} products`);
        
        // If no products found, try alternative parsing
        if (stockData.length === 0) {
            console.log("No products found with standard parsing, trying alternative method...");
            parseAlternativeStructure(data);
        }
    }
    
    // Find where the actual data starts by looking for headers
    function findDataStartRow(data) {
        for (let i = 0; i < Math.min(15, data.length); i++) {
            const row = data[i];
            if (row && row.length > 5) {
                const rowText = row.join(' ').toLowerCase();
                
                // Look for header indicators in your Excel
                if ((rowText.includes('dhaka') && rowText.includes('jhenaidah') && 
                     rowText.includes('bogra') && rowText.includes('rangpur')) ||
                    (rowText.includes('product') && rowText.includes('name') && 
                     rowText.includes('pack') && rowText.includes('size'))) {
                    
                    console.log(`Found headers at row ${i} (${rowText.substring(0, 50)}...)`);
                    return i + 1; // Data starts after header row
                }
                
                // Also check for depot names in individual cells
                if (row[4] && row[4].toString().toLowerCase().includes('dhaka') &&
                    row[5] && row[5].toString().toLowerCase().includes('jhenaidah')) {
                    console.log(`Found depot headers at row ${i}`);
                    return i + 1;
                }
            }
        }
        return -1;
    }
    
    // Extract product information from a row
    function extractProductInfo(row, rowIndex) {
        // Check multiple columns for product name
        let productName = '';
        let productNameCol = -1;
        
        // Try column B (index 1) first, then others
        const potentialNameCols = [1, 0, 2, 3];
        
        for (const col of potentialNameCols) {
            if (row[col] && typeof row[col] === 'string') {
                const name = row[col].toString().trim();
                
                // Skip if it's a header, empty, or not a real product name
                if (name && name.length > 1 && 
                    !name.includes('Portfolio') && 
                    !name.includes('Product Name') &&
                    !name.includes('Depot FG') &&
                    !name.includes('Factory') &&
                    !name.includes('Total') &&
                    !name.includes('Sub') &&
                    name !== 'Surovi Agro Industries Ltd.' &&
                    !name.startsWith('National Stock')) {
                    
                    productName = name;
                    productNameCol = col;
                    break;
                }
            }
        }
        
        if (!productName) {
            return null;
        }
        
        // Extract pack size - usually in column C (index 2)
        const packSize = row[2] ? row[2].toString().trim() : '';
        
        // Extract stock values based on Surovi Excel structure
        // Default mapping for standard Surovi format:
        // E=Dhaka(4), F=Jhenaidah(5), G=Bogra(6), H=Rangpur(7), I=Chattagram(8)
        // J=Factory FG(9), K=Dam/Ex Factory FG(10)
        
        const dhaka = parseExcelStockValue(row[4], 'Dhaka', rowIndex, productName);
        const jhenaidah = parseExcelStockValue(row[5], 'Jhenaidah', rowIndex, productName);
        const bogra = parseExcelStockValue(row[6], 'Bogra', rowIndex, productName);
        const rangpur = parseExcelStockValue(row[7], 'Rangpur', rowIndex, productName);
        const chattagram = parseExcelStockValue(row[8], 'Chattagram', rowIndex, productName);
        
        // Try alternative column mappings if main ones are empty
        let factoryFG = parseExcelStockValue(row[9], 'Factory', rowIndex, productName);
        let damFG = parseExcelStockValue(row[10], 'Dam', rowIndex, productName);
        
        // If factory/dam not in expected columns, check nearby columns
        if (factoryFG === 0 && row.length > 12) {
            factoryFG = parseExcelStockValue(row[11], 'Factory Alt', rowIndex, productName);
        }
        
        // Calculate totals
        const totalDepot = dhaka + jhenaidah + bogra + rangpur + chattagram;
        const total = totalDepot + factoryFG + damFG;
        
        // Also check if there's a total column in the Excel
        let excelTotal = 0;
        if (row[11] && (typeof row[11] === 'number' || (typeof row[11] === 'string' && row[11].match(/\d+/)))) {
            excelTotal = parseExcelStockValue(row[11], 'Excel Total', rowIndex, productName);
        }
        
        // Use the larger of calculated total or Excel total
        const finalTotal = Math.max(total, excelTotal);
        
        // Create product object
        return {
            name: productName,
            packSize: packSize,
            portfolio: row[0] ? row[0].toString().trim() : '',
            dhaka: dhaka,
            jhenaidah: jhenaidah,
            bogra: bogra,
            rangpur: rangpur,
            chattagram: chattagram,
            factoryFG: factoryFG,
            damFG: damFG,
            totalDepot: totalDepot,
            total: finalTotal,
            rowIndex: rowIndex,
            rawData: row.slice(0, 15) // Store first 15 columns for debugging
        };
    }
    
    // Parse Excel values including formulas
    function parseExcelStockValue(value, location, rowIndex, productName) {
        if (value === undefined || value === null || value === '') {
            return 0;
        }
        
        // If it's already a number
        if (typeof value === 'number') {
            return Math.max(0, value);
        }
        
        // If it's a string
        if (typeof value === 'string') {
            const str = value.toString().trim();
            
            // Empty string
            if (str === '' || str === '-' || str === 'N/A') {
                return 0;
            }
            
            // Check if it's a formula (starts with =)
            if (str.startsWith('=')) {
                try {
                    // Handle Surovi-style formulas like: =400-240-40+200-40-40-120-80-40
                    let formula = str.substring(1);
                    
                    // Remove any quotes or extra characters
                    formula = formula.replace(/["']/g, '');
                    
                    // Check if it's a complex calculation
                    if (formula.includes('-') || formula.includes('+')) {
                        // Replace sequences like 400-240-40 with proper arithmetic
                        // First, handle subtraction chains
                        formula = formula.replace(/(\d+)-(\d+)-(\d+)/g, '$1 - $2 - $3');
                        formula = formula.replace(/(\d+)-(\d+)/g, '$1 - $2');
                        // Then handle addition
                        formula = formula.replace(/(\d+)\+(\d+)/g, '$1 + $2');
                        
                        // Remove any non-arithmetic characters
                        formula = formula.replace(/[^0-9\.\+\-\s]/g, '');
                        
                        // Evaluate safely
                        try {
                            const result = eval(formula);
                            return Math.max(0, Math.round(result));
                        } catch (evalError) {
                            console.warn(`Could not evaluate formula for ${productName} ${location}: ${formula}`, evalError);
                        }
                    }
                    
                    // If formula evaluation failed, try to extract numbers
                    const numbers = str.match(/\d+/g);
                    if (numbers && numbers.length > 0) {
                        // Try to calculate if it looks like a sequence of operations
                        if (numbers.length >= 3 && str.includes('-') && str.includes('+')) {
                            // Take the first number as base
                            let total = parseInt(numbers[0]);
                            let subtractMode = true;
                            
                            for (let i = 1; i < numbers.length; i++) {
                                if (subtractMode) {
                                    total -= parseInt(numbers[i]);
                                } else {
                                    total += parseInt(numbers[i]);
                                }
                                // Alternate if there are + and - in the string
                                if (str.includes('+') && str.includes('-')) {
                                    subtractMode = !subtractMode;
                                }
                            }
                            return Math.max(0, total);
                        }
                        return parseInt(numbers[0]);
                    }
                } catch (error) {
                    console.warn(`Error parsing formula for ${productName} ${location}: ${str}`, error);
                }
            }
            
            // Try to parse as a regular number
            const num = parseFloat(str.replace(/[^\d\.\-]/g, ''));
            if (!isNaN(num)) {
                return Math.max(0, num);
            }
        }
        
        return 0;
    }
    
    // Alternative parsing if standard method fails
    function parseAlternativeStructure(data) {
        console.log("Using alternative parsing method");
        
        // Simple approach: look for rows with product names and numbers
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length < 5) continue;
            
            // Find cells that might be product names
            for (let col = 0; col < row.length; col++) {
                const cell = row[col];
                if (cell && typeof cell === 'string') {
                    const str = cell.trim();
                    
                    // Check if this looks like a product name
                    if (str.length > 2 && str.length < 50 && 
                        !str.includes('Portfolio') &&
                        !str.includes('Product Name') &&
                        !str.includes('Depot') &&
                        !str.includes('Factory') &&
                        !str.includes('Total') &&
                        !str.match(/^[0-9\.\-\+=]+$/) && // Not just numbers/formulas
                        !str.includes('Surovi') &&
                        !str.includes('National Stock')) {
                        
                        // This might be a product name
                        const productName = str;
                        
                        // Look for stock values in subsequent columns
                        const stockValues = [];
                        for (let stockCol = col + 1; stockCol < Math.min(col + 10, row.length); stockCol++) {
                            const stockVal = parseExcelStockValue(row[stockCol], `Col${stockCol}`, i, productName);
                            if (stockVal > 0 || (row[stockCol] && typeof row[stockCol] === 'string' && row[stockCol].includes('='))) {
                                stockValues.push(stockVal);
                            }
                        }
                        
                        // Need at least 2 stock values to be useful
                        if (stockValues.length >= 2) {
                            const product = {
                                name: productName,
                                packSize: '',
                                portfolio: '',
                                dhaka: stockValues.length > 0 ? stockValues[0] : 0,
                                jhenaidah: stockValues.length > 1 ? stockValues[1] : 0,
                                bogra: stockValues.length > 2 ? stockValues[2] : 0,
                                rangpur: stockValues.length > 3 ? stockValues[3] : 0,
                                chattagram: stockValues.length > 4 ? stockValues[4] : 0,
                                factoryFG: 0,
                                damFG: 0,
                                totalDepot: stockValues.reduce((a, b) => a + b, 0),
                                total: stockValues.reduce((a, b) => a + b, 0),
                                rowIndex: i,
                                rawData: []
                            };
                            
                            stockData.push(product);
                            break; // Move to next row
                        }
                    }
                }
            }
        }
        
        console.log(`Alternative parsing found ${stockData.length} products`);
    }
    
    // Apply filters to the data
    function applyFilters() {
        if (stockData.length === 0) {
            console.log("No data to filter");
            return;
        }
        
        const threshold = parseInt(lowStockThreshold.value) || 50;
        const selectedLocation = locationFilter.value;
        const searchTerm = productFilter.value.toLowerCase();
        
        console.log(`Applying filters: threshold=${threshold}, location=${selectedLocation}, search="${searchTerm}"`);
        
        // Start with all data
        filteredData = [...stockData];
        
        // Filter by product name if search term exists
        if (searchTerm) {
            filteredData = filteredData.filter(product => 
                product.name.toLowerCase().includes(searchTerm) ||
                (product.packSize && product.packSize.toLowerCase().includes(searchTerm)) ||
                (product.portfolio && product.portfolio.toLowerCase().includes(searchTerm))
            );
        }
        
        // Filter by location if not "all"
        if (selectedLocation !== 'all') {
            const locationKey = selectedLocation.toLowerCase();
            filteredData = filteredData.filter(product => {
                if (locationKey === 'dhaka') return product.dhaka > 0;
                if (locationKey === 'jhenaidah') return product.jhenaidah > 0;
                if (locationKey === 'bogra') return product.bogra > 0;
                if (locationKey === 'rangpur') return product.rangpur > 0;
                if (locationKey === 'chattagram') return product.chattagram > 0;
                return true;
            });
        }
        
        // Sort the filtered data
        sortTable();
        
        // Update statistics
        updateStatistics();
        
        // Render the table
        renderTable();
    }
    
    // Sort the table data
    function sortTable() {
        if (filteredData.length === 0) return;
        
        filteredData.sort((a, b) => {
            let aValue, bValue;
            
            // Determine what to sort by
            switch (currentSortColumn) {
                case 'product':
                    aValue = a.name.toLowerCase();
                    bValue = b.name.toLowerCase();
                    break;
                case 'dhaka':
                    aValue = a.dhaka;
                    bValue = b.dhaka;
                    break;
                case 'jhenaidah':
                    aValue = a.jhenaidah;
                    bValue = b.jhenaidah;
                    break;
                case 'bogra':
                    aValue = a.bogra;
                    bValue = b.bogra;
                    break;
                case 'rangpur':
                    aValue = a.rangpur;
                    bValue = b.rangpur;
                    break;
                case 'chattagram':
                    aValue = a.chattagram;
                    bValue = b.chattagram;
                    break;
                case 'total':
                    aValue = a.total;
                    bValue = b.total;
                    break;
                default:
                    aValue = a.name.toLowerCase();
                    bValue = b.name.toLowerCase();
            }
            
            // Handle string vs number comparison
            if (typeof aValue === 'string' && typeof bValue === 'string') {
                return currentSortDirection === 'asc' 
                    ? aValue.localeCompare(bValue) 
                    : bValue.localeCompare(aValue);
            } else {
                // Numeric comparison
                return currentSortDirection === 'asc' 
                    ? aValue - bValue 
                    : bValue - aValue;
            }
        });
    }
    
    // Render the table with data
    function renderTable() {
        if (filteredData.length === 0) {
            stockTable.style.display = 'none';
            tablePlaceholder.style.display = 'block';
            productCount.textContent = '0';
            return;
        }
        
        const threshold = parseInt(lowStockThreshold.value) || 50;
        const selectedLocation = locationFilter.value;
        
        stockTable.style.display = 'table';
        tablePlaceholder.style.display = 'none';
        productCount.textContent = filteredData.length;
        
        // Clear the table
        tableBody.innerHTML = '';
        
        // Populate the table
        filteredData.forEach(product => {
            const row = document.createElement('tr');
            
            // Check if product is low stock
            const isLowStock = (
                product.dhaka < threshold ||
                product.jhenaidah < threshold ||
                product.bogra < threshold ||
                product.rangpur < threshold ||
                product.chattagram < threshold
            );
            
            if (isLowStock) {
                row.classList.add('low-stock-row');
            }
            
            // Determine status based on filters
            let status = 'Normal';
            let statusClass = 'normal';
            
            if (selectedLocation !== 'all') {
                const locationStock = product[selectedLocation.toLowerCase()];
                if (locationStock < threshold) {
                    status = 'Low Stock';
                    statusClass = 'low';
                }
            } else if (isLowStock) {
                status = 'Low Stock';
                statusClass = 'low';
            }
            
            // Create table row HTML
            row.innerHTML = `
                <td>
                    <div class="product-name">${escapeHtml(product.name)}</div>
                    ${product.packSize ? `<small class="pack-size">${escapeHtml(product.packSize)}</small>` : ''}
                    ${product.portfolio ? `<small class="portfolio">${escapeHtml(product.portfolio)}</small>` : ''}
                </td>
                <td class="stock-cell ${product.dhaka < threshold ? 'low' : ''}">
                    ${formatNumber(product.dhaka)}
                    ${product.dhaka < threshold ? '<i class="fas fa-exclamation-circle low-icon"></i>' : ''}
                </td>
                <td class="stock-cell ${product.jhenaidah < threshold ? 'low' : ''}">
                    ${formatNumber(product.jhenaidah)}
                    ${product.jhenaidah < threshold ? '<i class="fas fa-exclamation-circle low-icon"></i>' : ''}
                </td>
                <td class="stock-cell ${product.bogra < threshold ? 'low' : ''}">
                    ${formatNumber(product.bogra)}
                    ${product.bogra < threshold ? '<i class="fas fa-exclamation-circle low-icon"></i>' : ''}
                </td>
                <td class="stock-cell ${product.rangpur < threshold ? 'low' : ''}">
                    ${formatNumber(product.rangpur)}
                    ${product.rangpur < threshold ? '<i class="fas fa-exclamation-circle low-icon"></i>' : ''}
                </td>
                <td class="stock-cell ${product.chattagram < threshold ? 'low' : ''}">
                    ${formatNumber(product.chattagram)}
                    ${product.chattagram < threshold ? '<i class="fas fa-exclamation-circle low-icon"></i>' : ''}
                </td>
                <td class="stock-cell ${product.total < threshold ? 'low' : ''}">
                    <strong>${formatNumber(product.total)}</strong>
                    ${product.totalDepot > 0 && product.totalDepot !== product.total ? 
                        `<br><small>Depot: ${formatNumber(product.totalDepot)}</small>` : ''}
                    ${product.factoryFG > 0 ? `<br><small>Factory: ${formatNumber(product.factoryFG)}</small>` : ''}
                </td>
                <td>
                    <span class="status ${statusClass}">${status}</span>
                    ${product.damFG > 0 ? `<br><small>Dam: ${formatNumber(product.damFG)}</small>` : ''}
                </td>
            `;
            
            tableBody.appendChild(row);
        });
    }
    
    // Update statistics display
    function updateStatistics() {
        if (filteredData.length === 0) {
            totalProducts.textContent = '0';
            totalUnits.textContent = '0';
            lowStockItems.textContent = '0';
            depotCount.textContent = '5';
            return;
        }
        
        const threshold = parseInt(lowStockThreshold.value) || 50;
        
        // Update total products
        totalProducts.textContent = filteredData.length;
        
        // Calculate total units and low stock items
        let totalUnitsValue = 0;
        let lowStockCount = 0;
        
        filteredData.forEach(product => {
            totalUnitsValue += product.total;
            
            // Check if product is low stock in any location
            if (
                product.dhaka < threshold ||
                product.jhenaidah < threshold ||
                product.bogra < threshold ||
                product.rangpur < threshold ||
                product.chattagram < threshold
            ) {
                lowStockCount++;
            }
        });
        
        totalUnits.textContent = formatNumber(totalUnitsValue);
        lowStockItems.textContent = lowStockCount;
        
        // Count unique depots with stock
        const depotsWithStock = new Set();
        filteredData.forEach(product => {
            if (product.dhaka > 0) depotsWithStock.add('Dhaka');
            if (product.jhenaidah > 0) depotsWithStock.add('Jhenaidah');
            if (product.bogra > 0) depotsWithStock.add('Bogra');
            if (product.rangpur > 0) depotsWithStock.add('Rangpur');
            if (product.chattagram > 0) depotsWithStock.add('Chattagram');
        });
        
        depotCount.textContent = depotsWithStock.size;
    }
    
    // Export data to Excel
    function exportToExcel() {
        if (filteredData.length === 0) {
            alert('No data to export. Please upload and process a file first.');
            return;
        }
        
        try {
            // Create worksheet data
            const wsData = [
                ['Surovi Agro Industries Ltd. - Stock Export'],
                [`Exported: ${new Date().toLocaleString()}`],
                [`Source File: ${currentFileName || 'Unknown'}`],
                [`Low Stock Threshold: ${lowStockThreshold.value}`],
                [],
                ['Product Name', 'Pack Size', 'Portfolio', 'Dhaka', 'Jhenaidah', 'Bogra', 'Rangpur', 'Chattagram', 'Factory FG', 'Dam FG', 'Total Depot', 'Grand Total', 'Status']
            ];
            
            const threshold = parseInt(lowStockThreshold.value) || 50;
            
            // Add data rows
            filteredData.forEach(product => {
                let status = 'Normal';
                
                if (
                    product.dhaka < threshold ||
                    product.jhenaidah < threshold ||
                    product.bogra < threshold ||
                    product.rangpur < threshold ||
                    product.chattagram < threshold
                ) {
                    status = 'Low Stock';
                }
                
                wsData.push([
                    product.name,
                    product.packSize || '',
                    product.portfolio || '',
                    product.dhaka,
                    product.jhenaidah,
                    product.bogra,
                    product.rangpur,
                    product.chattagram,
                    product.factoryFG,
                    product.damFG,
                    product.totalDepot,
                    product.total,
                    status
                ]);
            });
            
            // Add summary row
            wsData.push([]);
            wsData.push(['Summary', '', '', '', '', '', '', '', '', '', '', '', '']);
            
            const totalProducts = filteredData.length;
            const totalUnits = filteredData.reduce((sum, p) => sum + p.total, 0);
            const lowStockCount = filteredData.filter(p => 
                p.dhaka < threshold || p.jhenaidah < threshold || p.bogra < threshold || 
                p.rangpur < threshold || p.chattagram < threshold
            ).length;
            
            wsData.push(['Total Products', totalProducts]);
            wsData.push(['Total Units', totalUnits]);
            wsData.push(['Low Stock Items', lowStockCount]);
            wsData.push(['Export Date', new Date().toLocaleDateString()]);
            
            // Create worksheet
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            
            // Set column widths
            const wscols = [
                {wch: 30}, // Product Name
                {wch: 10}, // Pack Size
                {wch: 15}, // Portfolio
                {wch: 10}, // Dhaka
                {wch: 12}, // Jhenaidah
                {wch: 10}, // Bogra
                {wch: 10}, // Rangpur
                {wch: 12}, // Chattagram
                {wch: 12}, // Factory FG
                {wch: 10}, // Dam FG
                {wch: 12}, // Total Depot
                {wch: 12}, // Grand Total
                {wch: 12}  // Status
            ];
            ws['!cols'] = wscols;
            
            // Create workbook
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Stock Summary');
            
            // Generate filename
            const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
            const filename = `Surovi_Stock_Export_${dateStr}.xlsx`;
            
            // Export file
            XLSX.writeFile(wb, filename);
            
            console.log(`Exported ${filteredData.length} products to ${filename}`);
            
        } catch (error) {
            console.error('Error exporting to Excel:', error);
            alert('Error exporting data to Excel. Please try again.');
        }
    }
    
    // Helper functions
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
    
    function formatNumber(num) {
        if (num >= 1000000) {
            return (num / 1000000).toFixed(1).replace(/\.0$/, '') + 'M';
        } else if (num >= 1000) {
            return (num / 1000).toFixed(1).replace(/\.0$/, '') + 'K';
        } else {
            return num.toString();
        }
    }
    
    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    // Initialize with sample data for testing (remove in production)
    // Uncomment the next line to test without uploading a file
    // initializeWithSampleData();
    
    function initializeWithSampleData() {
        console.log("Initializing with sample data for testing");
        
        // Sample data matching Surovi products
        stockData = [
            { name: 'Prosper 10gm', packSize: '10gm', portfolio: 'Insecticides', dhaka: 120, jhenaidah: 85, bogra: 200, rangpur: 65, chattagram: 110, factoryFG: 50, damFG: 30, totalDepot: 580, total: 660, rowIndex: 1 },
            { name: 'Moto 20ml', packSize: '20ml', portfolio: 'Insecticides', dhaka: 45, jhenaidah: 30, bogra: 60, rangpur: 25, chattagram: 40, factoryFG: 20, damFG: 10, totalDepot: 200, total: 230, rowIndex: 2 },
            { name: 'Current 70 WP', packSize: '500gm', portfolio: 'Insecticides', dhaka: 320, jhenaidah: 180, bogra: 420, rangpur: 150, chattagram: 210, factoryFG: 100, damFG: 50, totalDepot: 1280, total: 1430, rowIndex: 3 },
            { name: 'Ratol 500gm', packSize: '500gm', portfolio: 'Rodenticides', dhaka: 15, jhenaidah: 8, bogra: 22, rangpur: 5, chattagram: 10, factoryFG: 0, damFG: 0, totalDepot: 60, total: 60, rowIndex: 4 },
            { name: 'Averast 5 SG', packSize: '16gm', portfolio: 'Insecticides', dhaka: 180, jhenaidah: 95, bogra: 220, rangpur: 80, chattagram: 130, factoryFG: 60, damFG: 40, totalDepot: 705, total: 805, rowIndex: 5 },
            { name: 'Avision 30ml', packSize: '30ml', portfolio: 'Insecticides', dhaka: 75, jhenaidah: 40, bogra: 90, rangpur: 35, chattagram: 55, factoryFG: 25, damFG: 15, totalDepot: 295, total: 335, rowIndex: 6 },
            { name: 'Filtap 50 SP', packSize: '100gm', portfolio: 'Insecticides', dhaka: 220, jhenaidah: 110, bogra: 250, rangpur: 95, chattagram: 140, factoryFG: 70, damFG: 30, totalDepot: 815, total: 915, rowIndex: 7 },
            { name: 'Anitrozine 80 WDG', packSize: '100gm', portfolio: 'Insecticides', dhaka: 28, jhenaidah: 12, bogra: 35, rangpur: 8, chattagram: 18, factoryFG: 5, damFG: 3, totalDepot: 101, total: 109, rowIndex: 8 },
            { name: 'Astrum 60 WDG', packSize: '100gm', portfolio: 'Insecticides', dhaka: 310, jhenaidah: 195, bogra: 380, rangpur: 145, chattagram: 225, factoryFG: 120, damFG: 60, totalDepot: 1255, total: 1435, rowIndex: 9 },
            { name: 'Audi 6 WG', packSize: '4gm', portfolio: 'Insecticides', dhaka: 42, jhenaidah: 20, bogra: 55, rangpur: 15, chattagram: 28, factoryFG: 12, damFG: 8, totalDepot: 160, total: 180, rowIndex: 10 },
            { name: 'Rota2.5 EC', packSize: '1000ml', portfolio: 'Insecticides', dhaka: 85, jhenaidah: 45, bogra: 110, rangpur: 35, chattagram: 65, factoryFG: 30, damFG: 15, totalDepot: 340, total: 385, rowIndex: 11 },
            { name: 'Agent 505 EC', packSize: '1000ml', portfolio: 'Insecticides', dhaka: 150, jhenaidah: 80, bogra: 190, rangpur: 60, chattagram: 100, factoryFG: 50, damFG: 25, totalDepot: 580, total: 655, rowIndex: 12 },
            { name: 'Iprotin 1.8 EC', packSize: '1000ml', portfolio: 'Insecticides', dhaka: 95, jhenaidah: 50, bogra: 120, rangpur: 40, chattagram: 75, factoryFG: 35, damFG: 20, totalDepot: 380, total: 435, rowIndex: 13 },
            { name: 'Bactrol 20 WP', packSize: '250gm', portfolio: 'Fungicides', dhaka: 280, jhenaidah: 150, bogra: 350, rangpur: 120, chattagram: 200, factoryFG: 100, damFG: 50, totalDepot: 1100, total: 1250, rowIndex: 14 },
            { name: 'Sinozeb 80 WP', packSize: '1000gm', portfolio: 'Fungicides', dhaka: 180, jhenaidah: 95, bogra: 220, rangpur: 80, chattagram: 130, factoryFG: 60, damFG: 40, totalDepot: 705, total: 805, rowIndex: 15 }
        ];
        
        fileInfo.classList.add('show');
        fileName.textContent = 'Sample_Stock_Data.xlsx';
        fileSize.textContent = '45 KB';
        reportDate.textContent = 'Jan 2026';
        
        exportButton.disabled = false;
        applyFilters();
        
        console.log("Sample data loaded with", stockData.length, "products");
    }
});
