/* global Excel */

// DataService - Handles Excel data reading and financial table structure analysis
class DataService {
    constructor() {
        this.CONFIG = {
            MAX_ROWS: 200,
            MAX_COLS: 30,
        };
        this.currentWorksheetData = null;
    }

    // Enhanced Excel data reading with better structure understanding
    async readCurrentWorksheetDataEnhanced(shouldUseSelection = true) {
        return new Promise((resolve, reject) => {
            Excel.run(async (context) => {
                try {
                    const worksheet = context.workbook.worksheets.getActiveWorksheet();
                    worksheet.load('name');
                    
                    let range;
                    let isSelection = false;
                    
                    if (shouldUseSelection) {
                        // Try to use selected range
                        const selectedRange = context.workbook.getSelectedRange();
                        selectedRange.load(['values', 'formulas', 'rowCount', 'columnCount', 'address']);
                        await context.sync();
                        
                        // Check if selection is valid (more than just a single empty cell)
                        if (selectedRange.rowCount > 1 || selectedRange.columnCount > 1 || 
                            (selectedRange.values[0][0] !== null && selectedRange.values[0][0] !== "")) {
                            range = selectedRange;
                            isSelection = true;
                        }
                    }
                    
                    // If not using selection or selection was invalid, use fixed range
                    if (!range) {
                        // Simple fixed range: first 200 rows × first 30 columns (A1:AD200)
                        range = worksheet.getRange('A1:AD200');
                        range.load(['values', 'formulas', 'rowCount', 'columnCount', 'address']);
                        await context.sync();
                        
                        console.log(`Using fixed range for non-selected mode: ${range.address}`);
                    }
                    
                    if (!range || range.rowCount === 0 || range.columnCount === 0) {
                        this.currentWorksheetData = {
                            worksheetName: worksheet.name,
                            structuredData: [],
                            summary: 'The current worksheet appears to be empty.'
                        };
                        resolve(this.currentWorksheetData);
                        return;
                    }
                    
                    // Use the range data directly - no truncation needed for fixed range
                    const rawValues = range.values;
                    const finalAddress = range.address;
                    const maxRows = range.rowCount;
                    const maxCols = range.columnCount;
                    
                    console.log(`Processing range: ${finalAddress} (${maxRows} rows × ${maxCols} columns)`);
                    
                    // Enhanced data structure analysis
                    const structuredData = this.analyzeFinancialTableStructure(rawValues);
                    
                    this.currentWorksheetData = {
                        worksheetName: worksheet.name,
                        address: finalAddress,
                        totalRows: range.rowCount,
                        totalCols: range.columnCount,
                        displayRows: maxRows,
                        displayCols: maxCols,
                        rawData: rawValues,
                        structuredData: structuredData,
                        isSelection: isSelection,
                        summary: `${isSelection ? 'Selected range' : 'Worksheet'} "${isSelection ? finalAddress : worksheet.name}" contains ${range.rowCount} rows and ${range.columnCount} columns.${
                            range.rowCount > this.CONFIG.MAX_ROWS || range.columnCount > this.CONFIG.MAX_COLS 
                            ? ` Showing first ${maxRows} rows and ${maxCols} columns for analysis.` 
                            : ''
                        }`
                    };
                    
                    resolve(this.currentWorksheetData);
                    
                } catch (error) {
                    reject(new Error('Failed to read Excel data: ' + error.message));
                }
            });
        });
    }

    // Enhanced financial table structure analysis
    analyzeFinancialTableStructure(rawValues) {
        if (!rawValues || rawValues.length === 0) {
            return { type: 'empty', headers: [], dataRows: [], keyRows: [] };
        }
        
        const result = {
            type: 'financial_table',
            headers: [],
            columnHeaders: [],
            dataRows: [],
            keyRows: [],
            totalRows: [],
            quarterlyData: [],
            rowLabels: []
        };
        
        // Clean and process data
        const cleanData = rawValues.map(row => 
            row.map(cell => {
                if (cell === null || cell === undefined) return '';
                if (typeof cell === 'string') return cell.trim();
                return cell;
            })
        );
        
        // Find potential header rows (usually first few rows with text)
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(5, cleanData.length); i++) {
            const row = cleanData[i];
            const textCellCount = row.filter(cell => 
                typeof cell === 'string' && cell.length > 0 && isNaN(cell)
            ).length;
            
            // Look for quarterly patterns like "1Q22", "2Q22", etc.
            const quarterlyPattern = row.filter(cell => 
                typeof cell === 'string' && /^[1-4]Q\d{2}$/i.test(cell)
            ).length;
            
            if (quarterlyPattern >= 2 || (textCellCount >= 3 && headerRowIndex === -1)) {
                headerRowIndex = i;
                result.columnHeaders = row.slice();
                break;
            }
        }
        
        // If no clear header found, use first row
        if (headerRowIndex === -1) {
            headerRowIndex = 0;
            result.columnHeaders = cleanData[0].slice();
        }
        
        // Process data rows
        for (let i = headerRowIndex + 1; i < cleanData.length; i++) {
            const row = cleanData[i];
            const rowLabel = row[0]; // First column is usually the row label
            
            if (!rowLabel || rowLabel === '') continue;
            
            const rowData = {
                rowIndex: i,
                label: rowLabel,
                values: row.slice(1), // Exclude the label column
                isTotal: this.isLikelyTotalRow(rowLabel),
                isSubtotal: this.isLikelySubtotalRow(rowLabel),
                category: this.categorizeRowLabel(rowLabel)
            };
            
            result.dataRows.push(rowData);
            result.rowLabels.push(rowLabel);
            
            // Identify key rows (totals, revenue, etc.)
            if (rowData.isTotal || /revenue|sales|income/i.test(rowLabel)) {
                result.keyRows.push(rowData);
            }
            
            if (rowData.isTotal) {
                result.totalRows.push(rowData);
            }
            
            // Extract quarterly data if headers contain quarters
            if (result.columnHeaders.some(h => /^[1-4]Q\d{2}$/i.test(String(h)))) {
                const quarterlyRow = this.extractQuarterlyData(rowData, result.columnHeaders);
                if (quarterlyRow.quarters.length > 0) {
                    result.quarterlyData.push(quarterlyRow);
                }
            }
        }
        
        return result;
    }

    // Helper functions for table structure analysis
    isLikelyTotalRow(label) {
        if (typeof label !== 'string') return false;
        const totalKeywords = /^(total|sum|grand total|net|aggregate)/i;
        return totalKeywords.test(label.trim());
    }

    isLikelySubtotalRow(label) {
        if (typeof label !== 'string') return false;
        const subtotalKeywords = /subtotal|sub-total|sub total/i;
        return subtotalKeywords.test(label.trim());
    }

    categorizeRowLabel(label) {
        if (typeof label !== 'string') return 'data';
        
        // Return the actual label for dynamic AI interpretation
        // The AI will understand the context and meaning of each row
        return label.trim();
    }

    extractQuarterlyData(rowData, headers) {
        const quarters = [];
        
        headers.forEach((header, index) => {
            if (typeof header === 'string' && /^[1-4]Q\d{2}$/i.test(header)) {
                const value = rowData.values[index - 1]; // -1 because values excludes label column
                if (value !== null && value !== undefined && value !== '') {
                    quarters.push({
                        quarter: header,
                        value: parseFloat(value) || value,
                        columnIndex: index
                    });
                }
            }
        });
        
        return {
            label: rowData.label,
            labelType: rowData.category, // Dynamic label for AI interpretation
            quarters: quarters,
            isTotal: rowData.isTotal
        };
    }

    // Utility functions for Excel column manipulation
    getColumnLetter(columnNumber) {
        let columnLetter = '';
        while (columnNumber > 0) {
            const remainder = (columnNumber - 1) % 26;
            columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
            columnNumber = Math.floor((columnNumber - 1) / 26);
        }
        return columnLetter;
    }

    getColumnNumber(columnLetter) {
        let columnNumber = 0;
        for (let i = 0; i < columnLetter.length; i++) {
            columnNumber = columnNumber * 26 + (columnLetter.charCodeAt(i) - 64);
        }
        return columnNumber;
    }

    // Getter for current worksheet data
    getCurrentWorksheetData() {
        return this.currentWorksheetData;
    }

    // Clear current data
    clearCurrentData() {
        this.currentWorksheetData = null;
    }
}

// Export for use in other modules
if (typeof window !== 'undefined') {
    window.DataService = DataService;
}