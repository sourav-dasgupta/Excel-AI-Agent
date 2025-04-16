/* global Excel */

export interface ExcelData {
  range: string;
  values: any[][];
  hasHeaders: boolean;
  columnCount: number;
  rowCount: number;
}

export async function extractDataFromSelection(): Promise<ExcelData | null> {
  try {
    return await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'values', 'rowCount', 'columnCount']);
      
      await context.sync();
      
      // Attempt to determine if first row contains headers
      const hasHeaders = determineIfHasHeaders(range.values);
      
      return {
        range: range.address,
        values: range.values,
        hasHeaders,
        columnCount: range.columnCount,
        rowCount: range.rowCount
      };
    });
  } catch (error) {
    console.error("Error extracting data from Excel:", error);
    return null;
  }
}

export async function insertFormula(formula: string, targetAddress?: string): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      const range = targetAddress 
        ? context.workbook.worksheets.getActiveWorksheet().getRange(targetAddress)
        : context.workbook.getSelectedRange();
        
      range.formulas = [[formula]];
      await context.sync();
      return true;
    });
  } catch (error) {
    console.error("Error inserting formula:", error);
    return false;
  }
}

export async function createChart(chartType: Excel.ChartType, dataRange: string, title?: string): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(dataRange);
      
      const chart = sheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);
      
      if (title) {
        chart.title.text = title;
      }
      
      chart.setPosition("A15", "H30");
      await context.sync();
      return true;
    });
  } catch (error) {
    console.error("Error creating chart:", error);
    return false;
  }
}

export async function createPivotTable(
  dataRange: string, 
  destinationSheet: string = "PivotTable", 
  rowFields?: string[], 
  columnFields?: string[], 
  valueFields?: string[]
): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      console.log("🚀 Starting pivot table creation process...");
      console.log(`📊 Data range: ${dataRange}, Destination: ${destinationSheet}`);
      
      // Get the source data range
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const sourceRange = sheet.getRange(dataRange);
      sourceRange.load("values, rowCount, columnCount");
      await context.sync();
      
      // Create new worksheet using our specialized function
      const pivotSheet = await createWorksheet(destinationSheet);
      
      if (!pivotSheet) {
        console.error("❌ Failed to create pivot worksheet");
        return false;
      }
      
      // Load needed properties
      pivotSheet.load("name, id, position");
      await context.sync();
      
      console.log(`📊 Source data: ${sourceRange.rowCount} rows x ${sourceRange.columnCount} columns`);
      console.log(`📝 Pivot sheet name: ${pivotSheet.name}, Position: ${pivotSheet.position}`);
      
      // Define destination for the pivot table (cell A3)
      const pivotDestination = pivotSheet.getRange("A3");
      
      // Create the pivot table with robust error handling
      try {
        console.log("🛠️ Creating pivot table...");
        const pivotTable = pivotSheet.pivotTables.add(
          "MyPivotTable", // Name
          sourceRange,    // Source
          pivotDestination // Destination
        );
        
        // Load the newly created pivot table
        pivotTable.load("name");
        await context.sync();
        console.log(`✅ Pivot table created: ${pivotTable.name}`);
        
        // Extract headers from data
        const headerRow = sourceRange.values[0];
        console.log("📋 Headers:", headerRow);
        
        // If no fields specified, use smart defaults
        if (!rowFields && !columnFields && !valueFields) {
          console.log("🧠 No fields specified, using smart defaults");
          
          // Find potential row fields (text columns)
          const potentialRowFields = [];
          // Find potential value fields (numeric columns)
          const potentialValueFields = [];
          
          // Analyze first few rows to find text vs numeric columns
          const sampleSize = Math.min(5, sourceRange.values.length);
          for (let i = 0; i < headerRow.length; i++) {
            let numericCount = 0;
            let textCount = 0;
            
            for (let row = 1; row < sampleSize; row++) {
              if (row < sourceRange.values.length) {
                const val = sourceRange.values[row][i];
                if (typeof val === 'number') {
                  numericCount++;
                } else if (typeof val === 'string' && !isNaN(Number(val))) {
                  numericCount++;
                } else {
                  textCount++;
                }
              }
            }
            
            if (numericCount > textCount) {
              potentialValueFields.push(i);
            } else {
              potentialRowFields.push(i);
            }
          }
          
          console.log("🔍 Auto-detected row fields:", potentialRowFields.map(i => headerRow[i]));
          console.log("🔍 Auto-detected value fields:", potentialValueFields.map(i => headerRow[i]));
          
          // Use first text column as row, first date column (if any) as column
          // and first numeric column as value
          if (potentialRowFields.length > 0) {
            // Add first row field (limit to 2 max)
            const maxRows = Math.min(2, potentialRowFields.length);
            for (let i = 0; i < maxRows; i++) {
              const fieldIndex = potentialRowFields[i];
              console.log(`➕ Adding row field: ${headerRow[fieldIndex]}`);
              pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(fieldIndex));
            }
          }
          
          // Add all numeric columns as values
          if (potentialValueFields.length > 0) {
            const maxValues = Math.min(3, potentialValueFields.length);
            for (let i = 0; i < maxValues; i++) {
              const fieldIndex = potentialValueFields[i];
              console.log(`➕ Adding value field: ${headerRow[fieldIndex]}`);
              const dataHierarchy = pivotTable.dataHierarchies.add(
                pivotTable.hierarchies.getItem(fieldIndex)
              );
              dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
            }
          }
        } else {
          // Add specified fields
          
          // Add row fields
          if (rowFields && rowFields.length > 0) {
            rowFields.forEach(field => {
              const columnIndex = headerRow.findIndex(h => 
                String(h).toLowerCase() === field.toLowerCase());
              if (columnIndex >= 0) {
                console.log(`➕ Adding row field: ${field} (column ${columnIndex})`);
                pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(columnIndex));
              }
            });
          }
          
          // Add column fields
          if (columnFields && columnFields.length > 0) {
            columnFields.forEach(field => {
              const columnIndex = headerRow.findIndex(h => 
                String(h).toLowerCase() === field.toLowerCase());
              if (columnIndex >= 0) {
                console.log(`➕ Adding column field: ${field} (column ${columnIndex})`);
                pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(columnIndex));
              }
            });
          }
          
          // Add data fields
          if (valueFields && valueFields.length > 0) {
            valueFields.forEach(field => {
              const columnIndex = headerRow.findIndex(h => 
                String(h).toLowerCase() === field.toLowerCase());
              if (columnIndex >= 0) {
                console.log(`➕ Adding data field: ${field} (column ${columnIndex})`);
                const dataHierarchy = pivotTable.dataHierarchies.add(
                  pivotTable.hierarchies.getItem(columnIndex)
                );
                dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
              }
            });
          }
        }
        
        // Final sync to apply pivot table structure
        await context.sync();
        console.log("✅ Pivot table created successfully!");
        return true;
      } catch (pivotError) {
        console.error("❌ Error creating pivot table:", pivotError);
        
        // FALLBACK: Create a simple table on the new sheet
        try {
          console.log("🔄 Falling back to simple table creation");
          
          // Copy the source data to the destination worksheet
          const valuesRange = pivotSheet.getRange("A1");
          valuesRange.values = sourceRange.values;
          
          // Format as a table
          const table = pivotSheet.tables.add(
            valuesRange.getResizedRange(sourceRange.rowCount - 1, sourceRange.columnCount - 1),
            true
          );
          table.name = "DataTable";
          
          await context.sync();
          console.log("✅ Created a formatted table with the data");
          return true;
        } catch (tableError) {
          console.error("❌ Error creating table fallback:", tableError);
          return false;
        }
      }
    });
  } catch (error) {
    console.error("❌ Outer error in createPivotTable:", error);
    return false;
  }
}

export async function analyzeDataForPivotSuggestions(range: string): Promise<{
  numericColumns: string[];
  textColumns: string[];
  dateColumns: string[];
}> {
  try {
    return await Excel.run(async (context) => {
      const rangeObj = context.workbook.getSelectedRange();
      rangeObj.load('values');
      await context.sync();
      
      const values = rangeObj.values;
      const numericColumns: string[] = [];
      const textColumns: string[] = [];
      const dateColumns: string[] = [];
      
      // Assuming the first row contains headers
      const headers = values[0];
      
      // Analyze each column
      for (let col = 0; col < headers.length; col++) {
        let numericCount = 0;
        let textCount = 0;
        let dateCount = 0;
        
        // Check values in each column (skip header row)
        for (let row = 1; row < values.length; row++) {
          const value = values[row][col];
          
          if (value === null || value === undefined || value === '') {
            continue; // Skip empty cells
          }
          
          if (typeof value === 'number') {
            numericCount++;
          } else if (value instanceof Date || !isNaN(Date.parse(value))) {
            dateCount++;
          } else {
            textCount++;
          }
        }
        
        // Determine column type based on majority
        const total = numericCount + textCount + dateCount;
        if (total === 0) continue;
        
        if (numericCount / total > 0.7) {
          numericColumns.push(headers[col]);
        } else if (dateCount / total > 0.7) {
          dateColumns.push(headers[col]);
        } else {
          textColumns.push(headers[col]);
        }
      }
      
      return { numericColumns, textColumns, dateColumns };
    });
  } catch (error) {
    console.error("Error analyzing data:", error);
    return { numericColumns: [], textColumns: [], dateColumns: [] };
  }
}

export async function createSamplePivotTable(dataRange: string): Promise<boolean> {
  try {
    console.log("🎯 createSamplePivotTable called with range:", dataRange);
    
    // Try standard pivot table first
    const success = await createPivotTable(dataRange);
    
    if (success) {
      console.log("✅ Standard pivot table created successfully");
      return true;
    }
    
    console.log("⚠️ Standard pivot table failed, trying simple alternative");
    
    // If standard method fails, try the simple alternative
    const simplePivotSuccess = await createSimplePivotTable(dataRange);
    
    if (simplePivotSuccess) {
      console.log("✅ Simple alternative method succeeded");
      return true;
    }
    
    console.error("❌ All pivot table creation methods failed");
    return false;
  } catch (error) {
    console.error("❌ Error in createSamplePivotTable:", error);
    return false;
  }
}

export async function calculateAndInsertSum(targetCell?: string): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      console.log("🟢 calculateAndInsertSum START");
      
      // Get the selected range
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load('values, address, rowIndex, columnIndex, rowCount');
      
      // CRITICAL: Ensure sync occurs before processing
      await context.sync();
      
      console.log("🟢 Selected range:", selectedRange.address);
      
      // Calculate the sum
      let sum = 0;
      for (let row = 0; row < selectedRange.values.length; row++) {
        for (let col = 0; col < selectedRange.values[row].length; col++) {
          const val = selectedRange.values[row][col];
          if (typeof val === 'number') {
            sum += val;
          } else if (typeof val === 'string') {
            // Try to convert string to number
            const num = parseFloat(val);
            if (!isNaN(num)) {
              sum += num;
            }
          }
        }
      }
      
      console.log("💰 Calculated sum:", sum);
      
      // CRITICAL: Get active worksheet explicitly
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Get target cell - either specified or below the selection
      let targetRange;
      if (targetCell) {
        targetRange = worksheet.getRange(targetCell);
        console.log(`📍 Using specified target cell: ${targetCell}`);
      } else {
        // Create a target cell below the selected range
        const targetRow = selectedRange.rowIndex + selectedRange.rowCount;
        const targetCol = selectedRange.columnIndex;
        
        console.log(`📍 Using calculated target cell: Row ${targetRow}, Col ${targetCol}`);
        targetRange = worksheet.getCell(targetRow, targetCol);
      }
      
      // Insert the sum with explicit type
      targetRange.values = [[sum]];
      targetRange.numberFormat = [["0.00"]];
      
      // IMPORTANT: Ensure changes are synchronized to Excel
      await context.sync();
      
      console.log(`✅ Sum ${sum} written to cell`);
      return true;
    });
  } catch (error) {
    console.error("❌ Error calculating and inserting sum:", error);
    return false;
  }
}

export async function applyFormulaToSelection(formula: string): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      // Get the selected range
      const selectedRange = context.workbook.getSelectedRange();
      
      // If formula doesn't start with =, add it
      if (!formula.startsWith('=')) {
        formula = '=' + formula;
      }
      
      // Apply the formula
      selectedRange.formulas = [[formula]];
      
      await context.sync();
      return true;
    });
  } catch (error) {
    console.error("Error applying formula:", error);
    return false;
  }
}

export async function sumSpecificColumns(columnNames: string[]): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      console.log("🟢 sumSpecificColumns START with columns:", columnNames);
      
      // Get the selected range first to determine the data
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load('values, rowCount, columnCount, address, rowIndex, columnIndex');
      
      // CRITICAL: Ensure sync occurs before processing data
      await context.sync();
      
      console.log("🟢 Selected range:", selectedRange.address);
      console.log("🟢 Selected values:", selectedRange.values);
      console.log("🟢 Row index:", selectedRange.rowIndex, "Column index:", selectedRange.columnIndex);
      
      const values = selectedRange.values;
      if (!values || values.length <= 1) {
        console.error("❌ Not enough data (only 1 row or empty)");
        return false;
      }
      
      // Assume first row has headers
      const headers = values[0];
      console.log("🟢 Headers found:", headers);
      
      const results: { [key: string]: number } = {};
      
      // Case-insensitive column matching
      for (const colName of columnNames) {
        console.log(`🔍 Looking for column: ${colName}`);
        
        // Try exact match first
        let colIndex = headers.findIndex(h => 
          h?.toString().toLowerCase() === colName.toLowerCase());
        
        // If not found, try contains match
        if (colIndex < 0) {
          colIndex = headers.findIndex(h => 
            h?.toString().toLowerCase().includes(colName.toLowerCase()));
        }
        
        console.log(`${colIndex >= 0 ? '✅' : '❌'} Column "${colName}" found at index: ${colIndex}`);
        
        if (colIndex >= 0) {
          // Found the column, calculate sum
          let sum = 0;
          for (let row = 1; row < values.length; row++) {
            const val = values[row][colIndex];
            
            if (typeof val === 'number') {
              sum += val;
            } else if (typeof val === 'string') {
              const num = parseFloat(val);
              if (!isNaN(num)) sum += num;
            }
          }
          results[colName] = sum;
          console.log(`💰 Sum for ${colName}: ${sum}`);
          
          // CRITICAL: Get current worksheet explicitly
          const worksheet = context.workbook.worksheets.getActiveWorksheet();
          
          // Calculate target cell address
          const targetRowIndex = selectedRange.rowIndex + values.length;
          const targetColIndex = selectedRange.columnIndex + colIndex;
          
          console.log(`📍 Target cell: Row ${targetRowIndex}, Col ${targetColIndex}`);
          
          // Get cell by index rather than relative to range
          const targetCell = worksheet.getCell(targetRowIndex, targetColIndex);
          
          // Set value with explicit type
          targetCell.values = [[sum]];
          targetCell.numberFormat = [["0.00"]];
          
          // IMPORTANT: Flush changes after each cell update
          await context.sync();
          
          console.log(`✅ Value ${sum} written to cell at row ${targetRowIndex}, col ${targetColIndex}`);
        }
      }
      
      // Final sync to ensure all changes are applied
      await context.sync();
      console.log("✅ Sum operation completed, results:", results);
      return Object.keys(results).length > 0;
    });
  } catch (error) {
    console.error("❌ Error summing specific columns:", error);
    return false;
  }
}

export async function writeValueToCell(value: any, targetCell: string): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      console.log(`Writing value ${value} to cell ${targetCell}`);
      
      const range = context.workbook.worksheets.getActiveWorksheet().getRange(targetCell);
      range.values = [[value]];
      
      // Format as number if it is one
      if (typeof value === 'number' || !isNaN(parseFloat(value))) {
        range.numberFormat = [["0.00"]];
      }
      
      await context.sync();
      console.log(`Successfully wrote value to ${targetCell}`);
      return true;
    });
  } catch (error) {
    console.error(`Error writing to cell ${targetCell}:`, error);
    return false;
  }
}

export async function sumIntoNextRow(): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      // Get the selected range
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load('address, values, rowIndex, columnIndex, rowCount, columnCount');
      await context.sync();
      
      console.log("Selected range for next row sum:", selectedRange.address);
      
      // Calculate sums for each column
      const values = selectedRange.values;
      const columnSums = new Array(values[0].length).fill(0);
      
      for (let col = 0; col < values[0].length; col++) {
        for (let row = 0; row < values.length; row++) {
          const val = values[row][col];
          if (typeof val === 'number') {
            columnSums[col] += val;
          } else if (typeof val === 'string') {
            const num = parseFloat(val);
            if (!isNaN(num)) columnSums[col] += num;
          }
        }
      }
      
      console.log("Column sums:", columnSums);
      
      // Insert sums in the next row
      const targetRow = selectedRange.rowIndex + selectedRange.rowCount;
      
      for (let col = 0; col < columnSums.length; col++) {
        const targetCell = context.workbook.worksheets.getActiveWorksheet().getCell(
          targetRow, 
          selectedRange.columnIndex + col
        );
        
        // Only write non-zero values or columns that had a sum
        if (columnSums[col] !== 0) {
          targetCell.values = [[columnSums[col]]];
          targetCell.numberFormat = [["0.00"]];
        }
      }
      
      await context.sync();
      console.log("Sums written to next row");
      return true;
    });
  } catch (error) {
    console.error("Error summing into next row:", error);
    return false;
  }
}

// Helper functions
function determineIfHasHeaders(values: any[][]): boolean {
  if (values.length <= 1) return false;
  
  // Check if first row has different data types or patterns than other rows
  const firstRow = values[0];
  const secondRow = values[1];
  
  // Simple heuristic: check if first row elements are strings and second row elements are numbers
  let firstRowStringCount = 0;
  let secondRowNumberCount = 0;
  
  for (let i = 0; i < firstRow.length; i++) {
    if (typeof firstRow[i] === 'string') firstRowStringCount++;
    if (typeof secondRow[i] === 'number') secondRowNumberCount++;
  }
  
  return firstRowStringCount > secondRowNumberCount;
}

export async function testExcelConnection(): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      console.log("TESTING Excel connection...");
      
      // Get active worksheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      
      // Create test cell in A1
      const testCell = sheet.getRange("A1");
      const testValue = "Excel connection test: " + new Date().toLocaleTimeString();
      console.log("Writing test value:", testValue);
      
      // Write test value
      testCell.values = [[testValue]];
      
      // Use flush to force immediate sync
      await context.sync();
      
      // Read it back to confirm
      testCell.load("values");
      await context.sync();
      
      console.log("Value read back:", testCell.values[0][0]);
      
      return true;
    });
  } catch (error) {
    console.error("Excel Connection Test FAILED:", error);
    return false;
  }
}

export async function listAllWorksheets(): Promise<string[]> {
  try {
    return await Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name");
      
      await context.sync();
      
      const worksheetNames = worksheets.items.map(sheet => sheet.name);
      console.log("📚 All worksheets:", worksheetNames);
      
      return worksheetNames;
    });
  } catch (error) {
    console.error("❌ Error listing worksheets:", error);
    return [];
  }
}

export async function createWorksheet(name: string): Promise<Excel.Worksheet | null> {
  try {
    return await Excel.run(async (context) => {
      console.log(`🔶 Attempting to create worksheet: "${name}"`);
      
      // First, check if worksheet with this name already exists
      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name");
      await context.sync();
      
      const existingNames = worksheets.items.map(sheet => sheet.name);
      console.log(`🔶 Existing worksheets: ${existingNames.join(', ')}`);
      
      let newSheet: Excel.Worksheet;
      
      // Try different methods to create the worksheet
      try {
        if (existingNames.includes(name)) {
          console.log(`🔶 Worksheet "${name}" already exists, using it`);
          newSheet = worksheets.getItem(name);
        } else {
          console.log(`🔶 Creating new worksheet: "${name}"`);
          newSheet = worksheets.add(name);
        }
      } catch (error1) {
        console.error(`❌ First worksheet creation attempt failed: ${error1}`);
        
        try {
          // Try with a unique name
          const uniqueName = `${name}_${Date.now()}`;
          console.log(`🔶 Trying with unique name: "${uniqueName}"`);
          newSheet = worksheets.add(uniqueName);
        } catch (error2) {
          console.error(`❌ Second worksheet creation attempt failed: ${error2}`);
          
          // Last resort: Try to add at specific position
          try {
            console.log(`🔶 Trying to add worksheet at position 0`);
            newSheet = worksheets.add(name, "Before", 0);
          } catch (error3) {
            console.error(`❌ All worksheet creation attempts failed: ${error3}`);
            return null;
          }
        }
      }
      
      // Make sure it's visible
      newSheet.activate();
      newSheet.load("name, position, visibility");
      
      // Wait for worksheet to be fully created
      await context.sync();
      
      console.log(`✅ Successfully created worksheet "${newSheet.name}" at position ${newSheet.position}`);
      console.log(`✅ Worksheet visibility: ${newSheet.visibility}`);
      
      return newSheet;
    });
  } catch (outerError) {
    console.error(`❌ Outer error in createWorksheet: ${outerError}`);
    return null;
  }
}

export async function createSimplePivotTable(dataRange: string): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      console.log("🔄 Starting simple pivot table alternative...");
      
      // Get original data
      const sourceSheet = context.workbook.worksheets.getActiveWorksheet();
      const sourceRange = sourceSheet.getRange(dataRange);
      sourceRange.load("values, address, rowCount, columnCount");
      await context.sync();
      
      console.log(`📊 Source data range: ${sourceRange.address}`);
      
      // Create a data copy in a temporary array
      const data = sourceRange.values;
      if (!data || data.length === 0) {
        console.error("❌ No data found in selection");
        return false;
      }
      
      // Try a completely different approach to worksheet creation
      // This approach uses the lower-level API
      let pivotSheet: Excel.Worksheet;
      try {
        console.log("🔄 Creating new sheet with alternate method");
        // Add sheet to end
        pivotSheet = context.workbook.worksheets.add();
        pivotSheet.load("name");
        await context.sync();
        
        // Rename after creation
        pivotSheet.name = "PivotData";
        await context.sync();
        
        console.log(`✅ Created sheet: ${pivotSheet.name}`);
      } catch (sheetError) {
        console.error("❌ Failed to create new sheet:", sheetError);
        return false;
      }
      
      try {
        // Copy data to new sheet
        console.log("📋 Copying data to new sheet");
        const targetRange = pivotSheet.getRange("A1");
        targetRange.values = data;
        
        // Add a title
        const titleCell = pivotSheet.getRange("A1").getOffsetRange(-1, 0);
        titleCell.values = [["Data Summary"]];
        titleCell.format.font.bold = true;
        titleCell.format.font.size = 14;
        
        await context.sync();
        
        // Activate the new sheet to make it visible
        pivotSheet.activate();
        await context.sync();
        
        console.log("✅ Data copied to new sheet");
        return true;
      } catch (copyError) {
        console.error("❌ Failed to copy data:", copyError);
        return false;
      }
    });
  } catch (outerError) {
    console.error("❌ Outer error in createSimplePivotTable:", outerError);
    return false;
  }
} 