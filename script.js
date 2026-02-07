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
    
    // Control elements
    const lowStockThreshold = document.getElementById('lowStockThreshold');
    const locationFilter = document.getElementById('locationFilter');
    const productFilter = document.getElementById('productFilter');
    
    // Data storage
    let stockData = [];
    let filteredData = [];
    let currentSortColumn = 'product';
    let currentSortDirection = 'asc';
    
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
        
        // Show file info
        fileName.textContent = file.name;
        fileSize.textContent = formatFileSize(file.size);
        
        // Extract date from filename
        const dateMatch = file.name.match(/Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec\s+\d{4}/);
        if (dateMatch) {
            reportDate.textContent = dateMatch[0];
        } else {
            reportDate.textContent = 'Not detected';
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
                
                // Get the first sheet
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert to JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                // Parse the data (this is a simplified parser - you may need to adjust based on your Excel structure)
                parseStockData(jsonData);
                
                // Hide loading indicator
                loadingIndicator.classList.remove('show');
                
                // Enable export button
                exportButton.disabled = false;
                
                // Apply initial filters and render
                applyFilters();
                
            } catch (error) {
                console.error('Error processing Excel file:', error);
                alert('Error processing the Excel file. Please make sure it follows the correct format.');
                loadingIndicator.classList.remove('show');
            }
        };
        
        reader.onerror = function() {
            alert('Error reading the file.');
            loadingIndicator.classList.remove('show');
        };
        
        reader.readAsBinaryString(file);
    }
    
    function parseStockData(data) {
        stockData = [];
        
        // This is a simplified parser - you'll need to adjust it based on your actual Excel structure
        // Based on your sample data, I'm assuming the structure
        
        // Skip header rows and parse data rows
        for (let i = 6; i < data.length; i++) {
            const row = data[i];
            
            // Check if this row contains product data
            if (row && row.length >= 6 && row[1] && typeof row[1] === 'string') {
                const product = {
                    name: row[1].trim(),
                    dhaka: parseStockValue(row[4]),
                    jhenaidah: parseStockValue(row[5]),
                    bogra: parseStockValue(row[6]),
                    rangpur: parseStockValue(row[7]),
                    chattagram: parseStockValue(row[8]),
                    total: 0
                };
                
                // Calculate total
                product.total = product.dhaka + product.jhenaidah + product.bogra + product.rangpur + product.chattagram;
                
                // Only add if product has a name and at least some stock
                if (product.name && product.name.length > 0 && (product.total > 0 || product.name !== 'Portfolio')) {
                    stockData.push(product);
                }
            }
        }
        
        console.log(`Parsed ${stockData.length} products from Excel file`);
    }
    
    function parseStockValue(value) {
        if (value === undefined || value === null || value === '') return 0;
        
        // If it's already a number
        if (typeof value === 'number') return Math.max(0, value);
        
        // If it's a string that might contain a formula or number
        if (typeof value === 'string') {
            // Try to extract numbers from strings like "=100-50+20"
            const numberMatch = value.match(/[-+]?\d*\.?\d+/g);
            if (numberMatch) {
                // Calculate if it looks like a formula
                if (value.includes('=') && value.includes('-') && value.includes('+')) {
                    try {
                        // Remove the = sign and evaluate the expression
                        const expr = value.replace('=', '').replace(/(\d+)-(\d+)/g, '$1 - $2').replace(/(\d+)\+(\d+)/g, '$1 + $2');
                        return Math.max(0, eval(expr));
                    } catch {
                        // If evaluation fails, use the first number found
                        return Math.max(0, parseFloat(numberMatch[0]));
                    }
                } else {
                    return Math.max(0, parseFloat(numberMatch[0]));
                }
            }
        }
        
        return 0;
    }
    
    function applyFilters() {
        if (stockData.length === 0) return;
        
        const threshold = parseInt(lowStockThreshold.value) || 50;
        const selectedLocation = locationFilter.value;
        const searchTerm = productFilter.value.toLowerCase();
        
        filteredData = stockData.filter(product => {
            // Filter by product name
            if (searchTerm && !product.name.toLowerCase().includes(searchTerm)) {
                return false;
            }
            
            return true;
        });
        
        // Sort the filtered data
        sortTable();
        
        // Update statistics
        updateStatistics();
        
        // Render the table
        renderTable();
    }
    
    function sortTable() {
        filteredData.sort((a, b) => {
            let aValue, bValue;
            
            if (currentSortColumn === 'product') {
                aValue = a.name.toLowerCase();
                bValue = b.name.toLowerCase();
            } else if (currentSortColumn === 'total') {
                aValue = a.total;
                bValue = b.total;
            } else {
                aValue = a[currentSortColumn];
                bValue = b[currentSortColumn];
            }
            
            if (currentSortDirection === 'asc') {
                return aValue > bValue ? 1 : -1;
            } else {
                return aValue < bValue ? 1 : -1;
            }
        });
    }
    
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
            
            // Check if product is low stock in any location
            const isLowStock = (
                product.dhaka < threshold ||
                product.jhenaidah < threshold ||
                product.bogra < threshold ||
                product.rangpur < threshold ||
                product.chattagram < threshold ||
                product.total < threshold
            );
            
            if (isLowStock) {
                row.classList.add('low-stock-row');
            }
            
            // Check status based on location filter
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
            
            row.innerHTML = `
                <td>${product.name}</td>
                <td class="stock-cell ${product.dhaka < threshold ? 'low' : ''}">${formatNumber(product.dhaka)}</td>
                <td class="stock-cell ${product.jhenaidah < threshold ? 'low' : ''}">${formatNumber(product.jhenaidah)}</td>
                <td class="stock-cell ${product.bogra < threshold ? 'low' : ''}">${formatNumber(product.bogra)}</td>
                <td class="stock-cell ${product.rangpur < threshold ? 'low' : ''}">${formatNumber(product.rangpur)}</td>
                <td class="stock-cell ${product.chattagram < threshold ? 'low' : ''}">${formatNumber(product.chattagram)}</td>
                <td class="stock-cell ${product.total < threshold ? 'low' : ''}">${formatNumber(product.total)}</td>
                <td><span class="${statusClass}">${status}</span></td>
            `;
            
            tableBody.appendChild(row);
        });
    }
    
    function updateStatistics() {
        const threshold = parseInt(lowStockThreshold.value) || 50;
        
        // Update total products
        totalProducts.textContent = filteredData.length;
        
        // Calculate total units and low stock items
        let units = 0;
        let lowStockCount = 0;
        
        filteredData.forEach(product => {
            units += product.total;
            
            // Check if product is low stock
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
        
        totalUnits.textContent = formatNumber(units);
        lowStockItems.textContent = lowStockCount;
    }
    
    function exportToExcel() {
        if (filteredData.length === 0) {
            alert('No data to export. Please upload and process a file first.');
            return;
        }
        
        try {
            // Create a new worksheet
            const wsData = [
                ['Product Name', 'Dhaka', 'Jhenaidah', 'Bogra', 'Rangpur', 'Chattagram', 'Total', 'Status']
            ];
            
            const threshold = parseInt(lowStockThreshold.value) || 50;
            
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
                    product.dhaka,
                    product.jhenaidah,
                    product.bogra,
                    product.rangpur,
                    product.chattagram,
                    product.total,
                    status
                ]);
            });
            
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            
            // Create a new workbook
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Stock Summary');
            
            // Generate filename with date
            const dateStr = new Date().toISOString().slice(0, 10);
            const filename = `Surovi_Stock_Export_${dateStr}.xlsx`;
            
            // Export the file
            XLSX.writeFile(wb, filename);
            
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
            return (num / 1000000).toFixed(1) + 'M';
        } else if (num >= 1000) {
            return (num / 1000).toFixed(1) + 'K';
        } else {
            return num.toString();
        }
    }
    
    // Sample data for testing (remove in production)
    // Uncomment the next line to test with sample data without uploading a file
    // loadSampleData();
    
    function loadSampleData() {
        // Sample data for testing
        stockData = [
            { name: 'Prosper 10gm', dhaka: 120, jhenaidah: 85, bogra: 200, rangpur: 65, chattagram: 110, total: 580 },
            { name: 'Moto 20ml', dhaka: 45, jhenaidah: 30, bogra: 60, rangpur: 25, chattagram: 40, total: 200 },
            { name: 'Current 70 WP', dhaka: 320, jhenaidah: 180, bogra: 420, rangpur: 150, chattagram: 210, total: 1280 },
            { name: 'Ratol 500gm', dhaka: 15, jhenaidah: 8, bogra: 22, rangpur: 5, chattagram: 10, total: 60 },
            { name: 'Averast 5 SG', dhaka: 180, jhenaidah: 95, bogra: 220, rangpur: 80, chattagram: 130, total: 705 },
            { name: 'Avision 30ml', dhaka: 75, jhenaidah: 40, bogra: 90, rangpur: 35, chattagram: 55, total: 295 },
            { name: 'Filtap 50 SP', dhaka: 220, jhenaidah: 110, bogra: 250, rangpur: 95, chattagram: 140, total: 815 },
            { name: 'Anitrozine 80 WDG', dhaka: 28, jhenaidah: 12, bogra: 35, rangpur: 8, chattagram: 18, total: 101 },
            { name: 'Astrum 60 WDG', dhaka: 310, jhenaidah: 195, bogra: 380, rangpur: 145, chattagram: 225, total: 1255 },
            { name: 'Audi 6 WG', dhaka: 42, jhenaidah: 20, bogra: 55, rangpur: 15, chattagram: 28, total: 160 }
        ];
        
        fileInfo.classList.add('show');
        fileName.textContent = 'Sample_Data.xlsx';
        fileSize.textContent = '25 KB';
        reportDate.textContent = 'Jan 2026';
        
        exportButton.disabled = false;
        applyFilters();
    }
});
