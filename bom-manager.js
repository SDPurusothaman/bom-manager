// BOM Manager - Walmart Sourcing
// JavaScript functionality for BOM input, export, and email

// Row counters for each section
const rowCounters = {
    partComponent: 0,
    packagingLabeling: 0,
    overheadsCosts: 0
};

// Department and Category data (from Walmart Cat Group.xlsx)
const departmentCategories = {
    '16': [
        { value: 'garden-tools-watering', name: 'GARDEN TOOLS AND WATERING' },
        { value: 'grill-accessories-fuel', name: 'GRILL ACCESSORIES AND FUEL' },
        { value: 'outdoor-cooking', name: 'OUTDOOR COOKING' },
        { value: 'outdoor-power-equipment', name: 'OUTDOOR POWER EQUIPMENT' },
        { value: 'patio-decor', name: 'PATIO DECOR' },
        { value: 'patio-furniture', name: 'PATIO FURNITURE' },
        { value: 'planters-bird', name: 'PLANTERS AND BIRD' },
        { value: 'pool-salt', name: 'POOL AND SALT' }
    ],
    '56': [
        { value: 'annuals-vegetables', name: 'ANNUALS AND VEGETABLES' },
        { value: 'cvp', name: 'CVP' },
        { value: 'gardening', name: 'GARDENING' },
        { value: 'indoor-plants', name: 'INDOOR PLANTS' },
        { value: 'landscaping', name: 'LANDSCAPING' },
        { value: 'landscaping-plants', name: 'LANDSCAPING PLANTS' },
        { value: 'lawn-garden-chemicals', name: 'LAWN AND GARDEN CHEMICALS' },
        { value: 'other', name: 'OTHER' },
        { value: 'pr-annuals', name: 'PR ANNUALS' },
        { value: 'pr-everyday-gift-deco-bulbs', name: 'PR EVERYDAY GIFT DECO AND BULBS' },
        { value: 'pr-landscaping', name: 'PR LANDSCAPING' },
        { value: 'pr-perenials-l3', name: 'PR PERENIALS L3' },
        { value: 'pr-seed-gardening', name: 'PR SEED GARDENING' },
        { value: 'pr-vegetables-l3', name: 'PR VEGETABLES L3' },
        { value: 'trash-delete', name: 'TRASH DELETE' },
        { value: 'unpub-seasonal-plants', name: 'UNPUB SEASONAL PLANTS' },
        { value: 'unpub-tropicals-ferns', name: 'UNPUB TROPICALS AND FERNS' }
    ]
};

// Default rows for Packaging/Labeling section (from BOM template)
const defaultPackagingRows = [
    { part: 'Retail Packaging', material: '', spec: '' },
    { part: 'AD Labels', material: '', spec: '' },
    { part: 'User Manual/ On product labels', material: '', spec: '' },
    { part: 'Shipping Carton', material: '', spec: '' },
    { part: 'Slip Sheet (if applicable)', material: '', spec: '' },
    { part: 'Others (not included above)', material: '', spec: '' }
];

// Default rows for Overhead + Other Costs section (from BOM template)
// Note: Profit moved to last position
const defaultOverheadRows = [
    { part: 'Labor Cost', remark: '' },
    { part: 'Logistics', remark: 'Including inland/outbound transportation to port' },
    { part: 'Overhead / Administration', remark: '' },
    { part: 'License fee/ Royalty', remark: '' },
    { part: 'After Sales Service (if applicable)', remark: 'Call center/ Warranty' },
    { part: 'WM Store Displays (if applicable)', remark: 'Displays/ Demo Cost' },
    { part: 'Total Tooling/Fixture cost', remark: '' },
    { part: 'Total Testing/product qualification', remark: '' },
    { part: 'Pre Sales Service', remark: '' },
    { part: 'Total Certification (Compliance) cost', remark: '' },
    { part: 'Profit', remark: '' }
];

// Currency exchange rates (to USD)
const currencyRates = {
    'USD': { rate: 1, symbol: '$', name: 'US Dollar' },
    'CNY': { rate: 7.25, symbol: '¥', name: 'Chinese Yuan (RMB)' },
    'INR': { rate: 83.12, symbol: '₹', name: 'Indian Rupee' },
    'VND': { rate: 24500, symbol: '₫', name: 'Vietnamese Dong' },
    'KHR': { rate: 4100, symbol: '៛', name: 'Cambodian Riel' }
};

// Country options for dropdown
const countries = [
    'China', 'Thailand', 'Vietnam', 'India', 'Indonesia', 'Malaysia',
    'Philippines', 'Bangladesh', 'Mexico', 'USA', 'Taiwan', 'South Korea',
    'Japan', 'Pakistan', 'Sri Lanka', 'Cambodia', 'Turkey', 'Other'
];

// Unit of measure options
const units = ['pcs', 'kg', 'g', 'yard', 'm', 't', 'lb', 'oz', 'set', 'pair', 'roll'];

// Product image data
let productImageData = null;

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => {
    // Set today's date as default
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('bomDate').value = today;
    
    // Add initial row to Part/Component section
    addRow('partComponent');
    
    // Add default rows to Packaging/Labeling section
    addDefaultPackagingRows();
    
    // Add default rows to Overheads section
    addDefaultOverheadRows();
    
    // Update email subject with product description
    document.getElementById('productDescription').addEventListener('input', (e) => {
        document.getElementById('emailSubject').value = `BOM - ${e.target.value}`;
    });
});

// Add default packaging rows from template
function addDefaultPackagingRows() {
    defaultPackagingRows.forEach(item => {
        addRow('packagingLabeling');
        const row = document.getElementById(`packagingLabeling-row-${rowCounters.packagingLabeling}`);
        row.querySelector('[name="partsComponents"]').value = item.part;
        row.querySelector('[name="commodityMaterials"]').value = item.material || '';
        row.querySelector('[name="keySpec"]').value = item.spec || '';
        if (item.remark) {
            row.querySelector('[name="remark"]').value = item.remark;
        }
    });
}

// Add default overhead rows from template
function addDefaultOverheadRows() {
    defaultOverheadRows.forEach(item => {
        addRow('overheadsCosts');
        const row = document.getElementById(`overheadsCosts-row-${rowCounters.overheadsCosts}`);
        row.querySelector('[name="partsComponents"]').value = item.part;
        if (item.remark) {
            row.querySelector('[name="remark"]').value = item.remark;
        }
    });
}

// Update categories based on department selection
function updateCategories() {
    const deptSelect = document.getElementById('departmentNumber');
    const catSelect = document.getElementById('category');
    const deptValue = deptSelect.value;
    
    // Clear existing options
    catSelect.innerHTML = '<option value="">Select Category...</option>';
    
    if (deptValue && departmentCategories[deptValue]) {
        departmentCategories[deptValue].forEach(cat => {
            const option = document.createElement('option');
            option.value = cat.value;
            option.textContent = cat.name;
            catSelect.appendChild(option);
        });
    }
}

// Update exchange rate based on currency selection
function updateExchangeRate() {
    const currencySelect = document.getElementById('currency');
    const exchangeRateInput = document.getElementById('exchangeRate');
    const selectedCurrency = currencySelect.value;
    
    if (currencyRates[selectedCurrency]) {
        exchangeRateInput.value = currencyRates[selectedCurrency].rate;
    }
    
    updateAllTotals();
}

// Image preview functions
function previewImage(event) {
    const file = event.target.files[0];
    if (file) {
        // Check file size (5MB limit)
        if (file.size > 5 * 1024 * 1024) {
            showNotification('Image file size must be less than 5MB', 'error');
            return;
        }
        
        const reader = new FileReader();
        reader.onload = function(e) {
            productImageData = e.target.result;
            document.getElementById('imagePreview').src = productImageData;
            document.getElementById('imagePreviewContainer').classList.remove('hidden');
        };
        reader.readAsDataURL(file);
    }
}

function clearImage() {
    document.getElementById('productImage').value = '';
    document.getElementById('imagePreview').src = '';
    document.getElementById('imagePreviewContainer').classList.add('hidden');
    productImageData = null;
}

// Add a new row to a specific section
function addRow(section) {
    rowCounters[section]++;
    const rowId = `${section}-row-${rowCounters[section]}`;
    const tbody = document.getElementById(`${section}Body`);
    
    const row = document.createElement('tr');
    row.className = 'hover:bg-walmart-blue-10 transition-colors animate-fade-in';
    row.id = rowId;
    row.dataset.section = section;
    
    // Section C (Overheads) has a simplified structure
    if (section === 'overheadsCosts') {
        row.innerHTML = `
            <td class="px-3 py-2 text-walmart-gray-160 font-medium row-number"></td>
            <td class="px-3 py-2">
                <input type="text" name="partsComponents" placeholder="Cost item" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm">
            </td>
            <td class="px-3 py-2">
                <select name="countryOfOrigin" class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm bg-white">
                    <option value="">Select...</option>
                    ${countries.map(c => `<option value="${c}">${c}</option>`).join('')}
                </select>
            </td>
            <td class="px-3 py-2">
                <input type="number" name="usageQty" step="0.001" min="0" placeholder="0" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm text-right"
                    onchange="calculateOverheadCost('${rowId}')">
            </td>
            <td class="px-3 py-2">
                <select name="unitOfMeasure" class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm bg-white">
                    ${units.map(u => `<option value="${u}">${u}</option>`).join('')}
                </select>
            </td>
            <td class="px-3 py-2">
                <input type="number" name="unitCost" step="0.0001" min="0" placeholder="0.00" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm text-right"
                    onchange="calculateOverheadCost('${rowId}')">
            </td>
            <td class="px-3 py-2">
                <input type="text" name="remark" placeholder="Notes" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm">
            </td>
            <td class="px-3 py-2">
                <button onclick="deleteRow('${rowId}', '${section}')" 
                    class="p-1.5 text-walmart-red-100 hover:bg-walmart-red-10 rounded transition-colors"
                    title="Delete row">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                            d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/>
                    </svg>
                </button>
            </td>
        `;
    } else {
        // Sections A and B have full structure
        row.innerHTML = `
            <td class="px-3 py-2 text-walmart-gray-160 font-medium row-number"></td>
            <td class="px-3 py-2">
                <input type="text" name="partsComponents" placeholder="Part name" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm">
            </td>
            <td class="px-3 py-2">
                <input type="text" name="commodityMaterials" placeholder="Material" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm">
            </td>
            <td class="px-3 py-2">
                <input type="text" name="keySpec" placeholder="Specifications" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm">
            </td>
            <td class="px-3 py-2">
                <input type="text" name="brand" placeholder="Brand" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm">
            </td>
            <td class="px-3 py-2">
                <input type="text" name="modelNumber" placeholder="Model #" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm">
            </td>
            <td class="px-3 py-2">
                <input type="text" name="upstreamSupplier" placeholder="Supplier" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm">
            </td>
            <td class="px-3 py-2">
                <select name="countryOfOrigin" class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm bg-white">
                    <option value="">Select...</option>
                    ${countries.map(c => `<option value="${c}">${c}</option>`).join('')}
                </select>
            </td>
            <td class="px-3 py-2">
                <input type="number" name="usageQty" step="0.001" min="0" placeholder="0" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm text-right"
                    onchange="calculateExtendedCost('${rowId}')">
            </td>
            <td class="px-3 py-2">
                <select name="unitOfMeasure" class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm bg-white">
                    ${units.map(u => `<option value="${u}">${u}</option>`).join('')}
                </select>
            </td>
            <td class="px-3 py-2">
                <input type="number" name="unitCost" step="0.01" min="0" placeholder="0.00" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm text-right"
                    onchange="calculateExtendedCost('${rowId}')">
            </td>
            <td class="px-3 py-2">
                <input type="text" name="extendedCost" readonly 
                    class="w-full px-2 py-1.5 bg-walmart-gray-10 border border-walmart-gray-50 rounded text-sm text-right font-medium text-walmart-blue-100">
            </td>
            <td class="px-3 py-2">
                <input type="text" name="remark" placeholder="Notes" 
                    class="w-full px-2 py-1.5 border border-walmart-gray-50 rounded text-sm">
            </td>
            <td class="px-3 py-2">
                <button onclick="deleteRow('${rowId}', '${section}')" 
                    class="p-1.5 text-walmart-red-100 hover:bg-walmart-red-10 rounded transition-colors"
                    title="Delete row">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                            d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/>
                    </svg>
                </button>
            </td>
        `;
    }
    
    tbody.appendChild(row);
    updateRowNumbers(section);
}

// Calculate cost for overhead rows (no extended cost field)
function calculateOverheadCost(rowId) {
    const row = document.getElementById(rowId);
    if (!row) return;
    
    const section = row.dataset.section;
    updateSectionTotal(section);
    updateGrandTotal();
}

// Delete a row
function deleteRow(rowId, section) {
    const row = document.getElementById(rowId);
    if (row) {
        row.remove();
        updateRowNumbers(section);
        updateSectionTotal(section);
        updateGrandTotal();
    }
}

// Update row numbers for a section
function updateRowNumbers(section) {
    const rows = document.querySelectorAll(`#${section}Body tr`);
    rows.forEach((row, index) => {
        const numberCell = row.querySelector('.row-number');
        if (numberCell) {
            numberCell.textContent = index + 1;
        }
    });
}

// Calculate extended cost for a row
function calculateExtendedCost(rowId) {
    const row = document.getElementById(rowId);
    if (!row) return;
    
    const usageQty = parseFloat(row.querySelector('[name="usageQty"]').value) || 0;
    const unitCost = parseFloat(row.querySelector('[name="unitCost"]').value) || 0;
    const extendedCost = usageQty * unitCost;
    
    row.querySelector('[name="extendedCost"]').value = extendedCost.toFixed(4);
    
    const section = row.dataset.section;
    updateSectionTotal(section);
    updateGrandTotal();
}

// Update section total
function updateSectionTotal(section) {
    const rows = document.querySelectorAll(`#${section}Body tr`);
    let total = 0;
    
    rows.forEach(row => {
        if (section === 'overheadsCosts') {
            // Section C uses unitCost directly (no extended cost calculation)
            const unitCost = parseFloat(row.querySelector('[name="unitCost"]').value) || 0;
            total += unitCost;
        } else {
            // Sections A and B use extendedCost
            const extendedCost = parseFloat(row.querySelector('[name="extendedCost"]').value) || 0;
            total += extendedCost;
        }
    });
    
    document.getElementById(`${section}Total`).textContent = `$${total.toFixed(2)}`;
    return total;
}

// Update all section totals
function updateAllTotals() {
    updateSectionTotal('partComponent');
    updateSectionTotal('packagingLabeling');
    updateSectionTotal('overheadsCosts');
    updateGrandTotal();
}

// Update grand total
function updateGrandTotal() {
    const partTotal = updateSectionTotal('partComponent');
    const packagingTotal = updateSectionTotal('packagingLabeling');
    const overheadsTotal = updateSectionTotal('overheadsCosts');
    
    const grandTotal = partTotal + packagingTotal + overheadsTotal;
    document.getElementById('grandTotal').textContent = `$${grandTotal.toFixed(2)}`;
    
    // Show converted amount
    const currency = document.getElementById('currency').value;
    const rate = parseFloat(document.getElementById('exchangeRate').value) || 1;
    
    if (currency !== 'USD') {
        const converted = grandTotal * rate;
        const symbol = currencyRates[currency]?.symbol || '';
        document.getElementById('grandTotalConverted').textContent = 
            `≈ ${symbol}${converted.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})} ${currency}`;
    } else {
        document.getElementById('grandTotalConverted').textContent = '';
    }
}

// Collect all BOM data
function collectBOMData() {
    const deptSelect = document.getElementById('departmentNumber');
    const catSelect = document.getElementById('category');
    
    const header = {
        date: document.getElementById('bomDate').value,
        supplierName: document.getElementById('supplierName').value,
        productDescription: document.getElementById('productDescription').value,
        departmentNumber: deptSelect.value,
        departmentName: deptSelect.options[deptSelect.selectedIndex]?.text || '',
        category: catSelect.value,
        categoryName: catSelect.options[catSelect.selectedIndex]?.text || '',
        currency: document.getElementById('currency').value,
        exchangeRate: document.getElementById('exchangeRate').value,
        selectedLevel: document.querySelector('input[name="levelSelection"]:checked')?.value || '',
        productImage: productImageData
    };
    
    const sections = {};
    ['partComponent', 'packagingLabeling'].forEach(section => {
        sections[section] = [];
        document.querySelectorAll(`#${section}Body tr`).forEach((row, index) => {
            sections[section].push({
                rowNum: index + 1,
                partsComponents: row.querySelector('[name="partsComponents"]').value,
                commodityMaterials: row.querySelector('[name="commodityMaterials"]').value,
                keySpec: row.querySelector('[name="keySpec"]').value,
                brand: row.querySelector('[name="brand"]').value,
                modelNumber: row.querySelector('[name="modelNumber"]').value,
                upstreamSupplier: row.querySelector('[name="upstreamSupplier"]').value,
                countryOfOrigin: row.querySelector('[name="countryOfOrigin"]').value,
                usageQty: row.querySelector('[name="usageQty"]').value,
                unitOfMeasure: row.querySelector('[name="unitOfMeasure"]').value,
                unitCost: row.querySelector('[name="unitCost"]').value,
                extendedCost: row.querySelector('[name="extendedCost"]').value,
                remark: row.querySelector('[name="remark"]').value
            });
        });
    });
    
    // Section C (Overheads) has simplified structure
    sections['overheadsCosts'] = [];
    document.querySelectorAll('#overheadsCostsBody tr').forEach((row, index) => {
        sections['overheadsCosts'].push({
            rowNum: index + 1,
            partsComponents: row.querySelector('[name="partsComponents"]').value,
            countryOfOrigin: row.querySelector('[name="countryOfOrigin"]').value,
            usageQty: row.querySelector('[name="usageQty"]').value,
            unitOfMeasure: row.querySelector('[name="unitOfMeasure"]').value,
            unitCost: row.querySelector('[name="unitCost"]').value,
            remark: row.querySelector('[name="remark"]').value
        });
    });
    
    return { header, sections };
}

// Export to Excel
function exportToExcel() {
    const data = collectBOMData();
    const wb = XLSX.utils.book_new();
    
    // Prepare worksheet data
    const wsData = [
        // Header information
        [`Date: ${data.header.date}`, `Product Description: ${data.header.productDescription}`, `Exchange Rate: ${data.header.exchangeRate}`],
        [`Supplier: ${data.header.supplierName}`, `Department: ${data.header.departmentName}`, `Currency: ${data.header.currency}`],
        [`Category: ${data.header.categoryName}`, `Level: ${data.header.selectedLevel || 'None'}`, ''],
        [], // Empty row
    ];
    
    // Column headers for Sections A and B
    const columnHeadersAB = [
        '', 'Parts/Components', 'Commodity Materials', 'Key Spec',
        'Brand', 'Model#', 'Upstream Supplier', 'Country of Origin',
        'Usage (QTY)', 'Unit', 'Unit Cost', 'Extended Cost', 'Remark'
    ];
    
    // Column headers for Section C (simplified)
    const columnHeadersC = [
        '', 'Cost Item', 'Country of Origin', 'Usage (QTY)', 'Unit', 'Unit Cost', 'Remark'
    ];
    
    // Section names for Column B headers
    const sectionNames = {
        partComponent: 'Part / Component',
        packagingLabeling: 'Packaging / Labeling',
        overheadsCosts: 'Overheads and Other Costs'
    };
    
    let sectionTotals = {};
    
    // Add Sections A and B (full structure)
    ['partComponent', 'packagingLabeling'].forEach(section => {
        wsData.push(['', sectionNames[section], '', '', '', '', '', '', '', '', '', '', '']);
        wsData.push(columnHeadersAB);
        
        let sectionTotal = 0;
        
        data.sections[section].forEach(row => {
            const extCost = parseFloat(row.extendedCost) || 0;
            sectionTotal += extCost;
            
            wsData.push([
                row.rowNum,
                row.partsComponents,
                row.commodityMaterials,
                row.keySpec,
                row.brand,
                row.modelNumber,
                row.upstreamSupplier,
                row.countryOfOrigin,
                parseFloat(row.usageQty) || 0,
                row.unitOfMeasure,
                parseFloat(row.unitCost) || 0,
                extCost,
                row.remark
            ]);
        });
        
        wsData.push(['', '', '', '', '', '', '', '', '', '', 'Subtotal:', sectionTotal.toFixed(2), '']);
        wsData.push([]);
        
        sectionTotals[section] = sectionTotal;
    });
    
    // Add Section C (simplified structure)
    wsData.push(['', sectionNames.overheadsCosts, '', '', '', '', '']);
    wsData.push(columnHeadersC);
    
    let overheadTotal = 0;
    data.sections.overheadsCosts.forEach(row => {
        const cost = parseFloat(row.unitCost) || 0;
        overheadTotal += cost;
        
        wsData.push([
            row.rowNum,
            row.partsComponents,
            row.countryOfOrigin,
            parseFloat(row.usageQty) || 0,
            row.unitOfMeasure,
            cost,
            row.remark
        ]);
    });
    
    wsData.push(['', '', '', '', 'Subtotal:', overheadTotal.toFixed(2), '']);
    wsData.push([]);
    sectionTotals.overheadsCosts = overheadTotal;
    
    // Grand total
    const grandTotal = Object.values(sectionTotals).reduce((a, b) => a + b, 0);
    wsData.push(['', '', '', '', '', 'GRAND TOTAL:', grandTotal.toFixed(2)]);
    
    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // Set column widths
    ws['!cols'] = [
        { wch: 5 },   // #
        { wch: 25 },  // Parts/Components
        { wch: 20 },  // Commodity Materials
        { wch: 30 },  // Key Spec
        { wch: 15 },  // Brand
        { wch: 15 },  // Model#
        { wch: 20 },  // Upstream Supplier
        { wch: 15 },  // Country of Origin
        { wch: 12 },  // Usage (QTY)
        { wch: 10 },  // Unit
        { wch: 12 },  // Unit Cost
        { wch: 15 },  // Extended Cost
        { wch: 20 }   // Remark
    ];
    
    // Add worksheet to workbook
    const sheetName = (data.header.productDescription || 'BOM').slice(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    
    // Generate filename
    const filename = `BOM_${(data.header.productDescription || 'Export').replace(/[^a-zA-Z0-9]/g, '_')}_${data.header.date}.xlsx`;
    
    // Download file
    XLSX.writeFile(wb, filename);
    
    showNotification('Excel file exported successfully!', 'success');
}

// Email modal functions
function emailBOM() {
    const data = collectBOMData();
    
    // Pre-fill subject
    document.getElementById('emailSubject').value = `BOM - ${data.header.productDescription}`;
    
    // Calculate totals
    let grandTotal = 0;
    let totalItems = 0;
    Object.values(data.sections).forEach(section => {
        section.forEach(row => {
            grandTotal += parseFloat(row.extendedCost) || 0;
            totalItems++;
        });
    });
    
    // Pre-fill message
    document.getElementById('emailMessage').value = 
        `Please find attached the Bill of Materials for ${data.header.productDescription}.\n\n` +
        `Supplier: ${data.header.supplierName}\n` +
        `Department: ${data.header.departmentName}\n` +
        `Category: ${data.header.categoryName}\n` +
        `Date: ${data.header.date}\n` +
        `Currency: ${data.header.currency}\n` +
        `Total Items: ${totalItems}\n` +
        `Grand Total: $${grandTotal.toFixed(2)}\n\n` +
        `Best regards`;
    
    // Show modal
    document.getElementById('emailModal').classList.remove('hidden');
    document.getElementById('emailModal').classList.add('flex');
}

function closeEmailModal() {
    document.getElementById('emailModal').classList.add('hidden');
    document.getElementById('emailModal').classList.remove('flex');
}

function sendEmail() {
    const recipient = document.getElementById('recipientEmail').value;
    const subject = document.getElementById('emailSubject').value;
    const message = document.getElementById('emailMessage').value;
    
    if (!recipient) {
        showNotification('Please enter a recipient email address', 'error');
        return;
    }
    
    // First export the Excel file
    exportToExcel();
    
    // Create mailto link
    const mailtoLink = `mailto:${encodeURIComponent(recipient)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(message + '\n\n[Please attach the downloaded Excel file]')}`;
    
    window.location.href = mailtoLink;
    
    closeEmailModal();
    showNotification('Email client opened. Please attach the downloaded Excel file.', 'success');
}

// Clear all data
function clearAll() {
    if (!confirm('Are you sure you want to clear all data? This cannot be undone.')) {
        return;
    }
    
    // Reset header fields
    document.getElementById('supplierName').value = '';
    document.getElementById('productDescription').value = '';
    document.getElementById('departmentNumber').value = '';
    document.getElementById('category').innerHTML = '<option value="">Select Department First...</option>';
    document.getElementById('currency').value = 'USD';
    document.getElementById('exchangeRate').value = '1';
    
    // Clear level selection (radio buttons)
    const levelRadios = document.querySelectorAll('input[name="levelSelection"]');
    levelRadios.forEach(radio => radio.checked = false);
    
    // Clear image
    clearImage();
    
    // Clear all sections
    ['partComponent', 'packagingLabeling', 'overheadsCosts'].forEach(section => {
        document.getElementById(`${section}Body`).innerHTML = '';
        rowCounters[section] = 0;
    });
    
    // Add initial row to Part/Component
    addRow('partComponent');
    
    // Add default rows to Packaging/Labeling and Overheads
    addDefaultPackagingRows();
    addDefaultOverheadRows();
    
    updateAllTotals();
    showNotification('All data cleared', 'success');
}

// Load sample data
function loadSampleData() {
    // Clear existing data
    ['partComponent', 'packagingLabeling', 'overheadsCosts'].forEach(section => {
        document.getElementById(`${section}Body`).innerHTML = '';
        rowCounters[section] = 0;
    });
    
    // Set header
    document.getElementById('supplierName').value = 'Thai Garden Tools Co., Ltd.';
    document.getElementById('productDescription').value = 'EG FG Bow Rake';
    document.getElementById('departmentNumber').value = '16';
    updateCategories();
    document.getElementById('category').value = 'garden-tools-watering';
    document.getElementById('currency').value = 'USD';
    document.getElementById('exchangeRate').value = '1';
    document.getElementById('levelL3').checked = true;  // Select L3 radio
    
    // Sample Part/Component data
    const partData = [
        { part: 'Blade with socket', material: 'Steel', spec: '14" x 3.30" & 0.14"', country: 'Thailand', qty: 0.536, unit: 't', cost: 2.6082 },
        { part: 'Collar', material: 'Chrome', spec: '1.56" x 1.29" & 0.02"', country: 'Thailand', qty: 0.021, unit: 't', cost: 12.021 },
        { part: 'Handle', material: 'Fiberglass', spec: '53.15" x 1.18"', country: 'China', qty: 1.27, unit: 'm', cost: 1.62 },
        { part: 'Grip', material: 'TPR', spec: '7.87" & 1.36"', country: 'Thailand', qty: 0.048, unit: 't', cost: 4.375 }
    ];
    
    partData.forEach(item => {
        addRow('partComponent');
        const row = document.getElementById(`partComponent-row-${rowCounters.partComponent}`);
        row.querySelector('[name="partsComponents"]').value = item.part;
        row.querySelector('[name="commodityMaterials"]').value = item.material;
        row.querySelector('[name="keySpec"]').value = item.spec;
        row.querySelector('[name="countryOfOrigin"]').value = item.country;
        row.querySelector('[name="usageQty"]').value = item.qty;
        row.querySelector('[name="unitOfMeasure"]').value = item.unit;
        row.querySelector('[name="unitCost"]').value = item.cost;
        calculateExtendedCost(`partComponent-row-${rowCounters.partComponent}`);
    });
    
    // Add default Packaging/Labeling rows from template with sample data
    const packagingSampleData = {
        'Retail Packaging': { material: 'pp label, T0.08mm', country: 'China', qty: 1, cost: 0.032 },
        'AD Labels': { material: '80g adhesive paper label', country: 'Thailand', qty: 1, cost: 0.0557 },
        'Shipping Carton': { material: 'pp poly bag', country: 'Thailand', qty: 1, cost: 0.0928 },
        'Others (not included above)': { material: 'paper sleeve', country: 'Thailand', qty: 1, cost: 0.0059 }
    };
    
    defaultPackagingRows.forEach(item => {
        addRow('packagingLabeling');
        const row = document.getElementById(`packagingLabeling-row-${rowCounters.packagingLabeling}`);
        row.querySelector('[name="partsComponents"]').value = item.part;
        row.querySelector('[name="commodityMaterials"]').value = item.material || '';
        row.querySelector('[name="keySpec"]').value = item.spec || '';
        if (item.remark) row.querySelector('[name="remark"]').value = item.remark;
        
        // Fill in sample data if available
        const sampleData = packagingSampleData[item.part];
        if (sampleData) {
            row.querySelector('[name="commodityMaterials"]').value = sampleData.material || '';
            row.querySelector('[name="countryOfOrigin"]').value = sampleData.country || '';
            row.querySelector('[name="usageQty"]').value = sampleData.qty || '';
            row.querySelector('[name="unitOfMeasure"]').value = 'pcs';
            row.querySelector('[name="unitCost"]').value = sampleData.cost || '';
            calculateExtendedCost(`packagingLabeling-row-${rowCounters.packagingLabeling}`);
        }
    });
    
    // Add default Overhead rows from template with sample data
    const overheadSampleData = {
        'Labor Cost': { cost: 0.3168 },
        'Logistics': { cost: 0.198 },
        'Overhead / Administration': { cost: 0.315 },
        'Profit': { cost: 0.2411 }
    };
    
    defaultOverheadRows.forEach(item => {
        addRow('overheadsCosts');
        const row = document.getElementById(`overheadsCosts-row-${rowCounters.overheadsCosts}`);
        row.querySelector('[name="partsComponents"]').value = item.part;
        if (item.remark) row.querySelector('[name="remark"]').value = item.remark;
        
        // Fill in sample data if available
        const sampleData = overheadSampleData[item.part];
        if (sampleData) {
            row.querySelector('[name="unitCost"]').value = sampleData.cost || '';
            calculateOverheadCost(`overheadsCosts-row-${rowCounters.overheadsCosts}`);
        }
    });
    
    showNotification('Sample data loaded', 'success');
}

// Show notification
function showNotification(message, type = 'success') {
    const existing = document.querySelector('.notification');
    if (existing) existing.remove();
    
    const colors = {
        success: 'bg-walmart-green-100',
        error: 'bg-walmart-red-100',
        warning: 'bg-walmart-spark-100 text-walmart-gray-160'
    };
    
    const notification = document.createElement('div');
    notification.className = `notification fixed top-4 right-4 ${colors[type]} text-white px-6 py-3 rounded-lg shadow-lg z-50 animate-fade-in font-medium`;
    notification.textContent = message;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.remove();
    }, 3000);
}
