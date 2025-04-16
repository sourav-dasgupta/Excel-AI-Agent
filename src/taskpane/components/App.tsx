import * as React from 'react';
import { useEffect, useState } from 'react';
import { TextField, PrimaryButton, Stack, Text, Spinner, MessageBar, MessageBarType } from '@fluentui/react';
import { getOpenAIResponse } from '../services/openai-service';
import { extractDataFromSelection, insertFormula, createChart, createPivotTable, createSamplePivotTable, calculateAndInsertSum, applyFormulaToSelection, sumSpecificColumns, writeValueToCell, sumIntoNextRow, testExcelConnection, listAllWorksheets, createWorksheet } from '../services/excel-service';

export interface Message {
  role: 'user' | 'assistant';
  content: string;
}

interface AIResponse {
  message: string;
  action?: {
    type: 'formula' | 'chart' | 'pivotTable' | 'function';
    params: any;
  };
}

export const App: React.FC = () => {
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  
  useEffect(() => {
    console.log("App component mounted");
    try {
      // Initialize with welcome message
      setMessages([{
        role: 'assistant',
        content: 'Hello! I\'m your Excel Assistant. Select some cells and ask me anything about Excel formulas, charts, or data analysis.'
      }]);
      console.log("Initial message set");
    } catch (err) {
      console.error("Error in useEffect:", err);
    }
  }, []);

  const handleSend = async () => {
    console.log("handleSend triggered");
    if (!input.trim()) return;
    
    try {
      setIsLoading(true);
      setError(null);
      
      // Add user message
      const userMessage = { role: 'user', content: input };
      setMessages(prev => [...prev, userMessage]);
      setInput('');
      
      // Get selected data from Excel
      const selectedData = await extractDataFromSelection();
      console.log("Selected data:", selectedData);
      
      // Detect task intent BEFORE getting AI response
      let actionPerformed = false;
      
      // Check for sum requests
      if (/\b(sum|add|total|calculate)\b.*?\b(column|row|cell|range|value|amount|cost|price|number|data)\b/i.test(input) && selectedData) {
        console.log("Detected sum request");
        try {
          const success = await calculateAndInsertSum();
          if (success) {
            actionPerformed = true;
            setMessages(prev => [...prev, { 
              role: 'assistant', 
              content: "I've calculated the sum of the selected cells and placed it in the cell below the selection."
            }]);
          }
        } catch (actionError) {
          console.error("Error performing sum action:", actionError);
        }
      }
      
      // Check for specific column sum requests
      else if ((/\b(sum|add|total)\b.*?\b(column|columns)\b/i.test(input) || 
                /\b(sum|add|total)\b.*?\b(amount|cost|price|value|total)\b/i.test(input)) && 
                selectedData) {
        console.log("Detected column sum request");
        try {
          // Try to match column names from the input
          const columnsToSum: string[] = [];
          
          // Extract column names mentioned in the input
          const columnMatches = input.match(/\b(sum|add|total)\b.*?\b(column|columns)?\s+([\w\s,]+)/i);
          let columnText = columnMatches ? columnMatches[3] : '';
          
          // Extract column names specifically mentioned
          const specificColumnPattern = /\b(amount|cost|price|value|total|sales|revenue|profit|quantity)\b/gi;
          let match;
          while ((match = specificColumnPattern.exec(input)) !== null) {
            columnsToSum.push(match[0]);
          }
          
          // If specific columns were found in the input, use those
          if (columnsToSum.length === 0 && columnText) {
            columnsToSum.push(...columnText
              .split(/,|and/)
              .map(col => col.trim())
              .filter(col => col.length > 0));
          }
          
          // If still no columns found, try to match headers from the data
          if (columnsToSum.length === 0 && selectedData && selectedData.values.length > 0) {
            const headers = selectedData.values[0];
            
            // If there are only a few columns, sum them all
            if (headers.length <= 3) {
              columnsToSum.push(...headers);
            } else {
              // Look for numeric columns
              for (let colIndex = 0; colIndex < headers.length; colIndex++) {
                const header = headers[colIndex];
                let hasNumbers = false;
                
                // Check if this column has numeric values
                for (let row = 1; row < selectedData.values.length; row++) {
                  const value = selectedData.values[row][colIndex];
                  if (typeof value === 'number' || (typeof value === 'string' && !isNaN(parseFloat(value)))) {
                    hasNumbers = true;
                    break;
                  }
                }
                
                if (hasNumbers) {
                  columnsToSum.push(header);
                }
              }
            }
          }
          
          console.log("Columns to sum:", columnsToSum);
          
          if (columnsToSum.length > 0) {
            const success = await sumSpecificColumns(columnsToSum);
            if (success) {
              actionPerformed = true;
              setMessages(prev => [...prev, { 
                role: 'assistant', 
                content: `I've calculated the sum for the following columns: ${columnsToSum.join(', ')}. The results have been placed in the cells below each column.`
              }]);
            } else {
              // Fallback to simple sum if column-specific sum failed
              const success = await calculateAndInsertSum();
              if (success) {
                actionPerformed = true;
                setMessages(prev => [...prev, { 
                  role: 'assistant', 
                  content: "I've calculated the sum of all selected cells and placed it in the cell below the selection."
                }]);
              }
            }
          }
        } catch (actionError) {
          console.error("Error performing column sum:", actionError);
        }
      }
      
      // Check for pivot table requests
      else if (/\b(create|make|build|generate)\b.*\b(pivot|pivot\s*table)\b/i.test(input) && selectedData) {
        console.log("Detected pivot table request");
        try {
          const success = await createSamplePivotTable(selectedData.range);
          if (success) {
            actionPerformed = true;
            setMessages(prev => [...prev, { 
              role: 'assistant', 
              content: "I've created a pivot table based on your selected data in a new worksheet named 'PivotTable'."
            }]);
          }
        } catch (actionError) {
          console.error("Error creating pivot table:", actionError);
        }
      }
      
      // Check for chart requests
      else if (/\b(create|make|build|generate|plot)\b.*\b(chart|graph|plot)\b/i.test(input) && selectedData) {
        console.log("Detected chart request");
        try {
          const chartType = getChartTypeFromInput(input);
          const success = await createChart(chartType, selectedData.range, "Chart from selected data");
          if (success) {
            actionPerformed = true;
            setMessages(prev => [...prev, { 
              role: 'assistant', 
              content: `I've created a ${getChartTypeName(chartType)} chart using your selected data.`
            }]);
          }
        } catch (actionError) {
          console.error("Error creating chart:", actionError);
        }
      }
      
      // Check for formula requests
      else if (/\b(apply|use|add|create)\b.*\b(formula|function)\b/i.test(input) && selectedData) {
        console.log("Detected formula request");
        try {
          // Extract the formula from the input
          const formulaMatch = input.match(/=([A-Z]+\([^)]+\))/i) || 
                               input.match(/=([A-Z0-9+\-*/^()&<>=" ]+)/i);
          
          if (formulaMatch) {
            const formula = formulaMatch[0];
            const success = await applyFormulaToSelection(formula);
            if (success) {
              actionPerformed = true;
              setMessages(prev => [...prev, { 
                role: 'assistant', 
                content: `I've applied the formula ${formula} to your selected cells.`
              }]);
            }
          }
        } catch (actionError) {
          console.error("Error applying formula:", actionError);
        }
      }
      
      // Special case for "amount and cost" which is common in your example
      if (input.toLowerCase().includes('amount') && 
          input.toLowerCase().includes('cost') && 
          selectedData) {
        try {
          // Direct approach for this common scenario
          let amountSum = 0;
          let costSum = 0;
          let amountCol = -1;
          let costCol = -1;
          
          // Find the amount and cost columns
          const headers = selectedData.values[0];
          for (let i = 0; i < headers.length; i++) {
            const header = String(headers[i]).toLowerCase();
            if (header.includes('amount')) amountCol = i;
            if (header.includes('cost')) costCol = i;
          }
          
          // Calculate sums
          if (amountCol >= 0) {
            for (let row = 1; row < selectedData.values.length; row++) {
              const val = selectedData.values[row][amountCol];
              if (typeof val === 'number') amountSum += val;
              else if (typeof val === 'string') {
                const num = parseFloat(val);
                if (!isNaN(num)) amountSum += num;
              }
            }
          }
          
          if (costCol >= 0) {
            for (let row = 1; row < selectedData.values.length; row++) {
              const val = selectedData.values[row][costCol];
              if (typeof val === 'number') costSum += val;
              else if (typeof val === 'string') {
                const num = parseFloat(val);
                if (!isNaN(num)) costSum += num;
              }
            }
          }
          
          // Write to specific cells
          const targetRow = selectedData.rowCount + 1;
          let success1 = false, success2 = false;
          
          if (amountCol >= 0) {
            const cellAddress = `${String.fromCharCode(65 + amountCol)}${targetRow}`;
            success1 = await writeValueToCell(amountSum, cellAddress);
          }
          
          if (costCol >= 0) {
            const cellAddress = `${String.fromCharCode(65 + costCol)}${targetRow}`;
            success2 = await writeValueToCell(costSum, cellAddress);
          }
          
          if (success1 || success2) {
            actionPerformed = true;
            setMessages(prev => [...prev, { 
              role: 'assistant', 
              content: `I've calculated the sums: Amount = ${amountSum}, Cost = ${costSum}. The results are now in row ${targetRow}.`
            }]);
          }
        } catch (err) {
          console.error("Error in amount/cost special case:", err);
        }
      }
      
      // Check for 'next row' requests
      else if (/\b(next|below|following)\b.*?\b(row|line)\b/i.test(input) || 
               input.toLowerCase().includes("next row") ||
               input.toLowerCase().includes("row below")) {
        console.log("🔍 Detected 'next row' request");
        try {
          setMessages(prev => [...prev, { 
            role: 'assistant', 
            content: "I'll calculate the sums and put them in the next row..."
          }]);
          
          const success = await sumIntoNextRow();
          
          if (success) {
            actionPerformed = true;
            setMessages(prev => [...prev, { 
              role: 'assistant', 
              content: "✅ I've calculated the sums for each column and placed them in the row below your selection."
            }]);
          } else {
            setMessages(prev => [...prev, { 
              role: 'assistant', 
              content: "❌ Sorry, I couldn't calculate the sums. Please make sure you've selected a range with numeric data."
            }]);
          }
        } catch (actionError) {
          console.error("Error performing next row sum:", actionError);
        }
      }
      
      // Only get AI response if no action was performed
      if (!actionPerformed) {
        console.log("No action performed, getting AI response");
        const response = await getOpenAIResponse(messages, input, selectedData);
        setMessages(prev => [...prev, { role: 'assistant', content: response }]);
      }
    } catch (err) {
      console.error("Detailed error in handleSend:", err);
      setError(`Error: ${err.message || 'Something went wrong'}`);
    } finally {
      setIsLoading(false);
    }
  };

  // Helper function to determine chart type
  const getChartTypeFromInput = (input: string): Excel.ChartType => {
    const input_lower = input.toLowerCase();
    
    if (input_lower.includes("bar chart") || input_lower.includes("bar graph")) {
      return Excel.ChartType.barClustered;
    } else if (input_lower.includes("line chart") || input_lower.includes("line graph")) {
      return Excel.ChartType.line;
    } else if (input_lower.includes("pie chart")) {
      return Excel.ChartType.pie;
    } else if (input_lower.includes("scatter") || input_lower.includes("scatter plot")) {
      return Excel.ChartType.xyscatter;
    } else {
      // Default to column chart
      return Excel.ChartType.columnClustered;
    }
  };

  // Helper function to get friendly chart type name
  const getChartTypeName = (chartType: Excel.ChartType): string => {
    switch (chartType) {
      case Excel.ChartType.barClustered: return "bar";
      case Excel.ChartType.line: return "line";
      case Excel.ChartType.pie: return "pie";
      case Excel.ChartType.xyscatter: return "scatter";
      case Excel.ChartType.columnClustered: return "column";
      default: return "column";
    }
  };

  const detectActionRequest = (input: string): boolean => {
    const actionKeywords = [
      'sum', 'add', 'total', 'calculate', 'compute',
      'create pivot', 'make pivot', 'create a pivot', 
      'create chart', 'make chart', 'plot', 'graph',
      'set cell', 'update cell', 'fill cell'
    ];
    
    return actionKeywords.some(keyword => 
      input.toLowerCase().includes(keyword)
    );
  };

  const executeAction = async (action: any, selectedData?: any) => {
    console.log("Executing action:", action);
    
    switch (action.type) {
      case 'formula':
        // Insert formula into selected cell or specified cell
        await insertFormula(action.params.formula, action.params.targetCell);
        break;
        
      case 'chart':
        // Create chart with selected data
        await createChart(
          action.params.chartType || Excel.ChartType.columnClustered,
          action.params.dataRange || selectedData?.range,
          action.params.title
        );
        break;
        
      case 'pivotTable':
        // Create pivot table
        await createPivotTable(
          action.params.dataRange || selectedData?.range,
          action.params.destinationSheet || "PivotTable",
          action.params.rowFields,
          action.params.columnFields,
          action.params.valueFields
        );
        break;
        
      case 'function':
        // Execute an Excel function
        if (action.params.function === 'SUM') {
          // Calculate sum and place in target
          if (selectedData) {
            // Flatten the 2D array and filter out non-numbers
            const numbers = selectedData.values.flat().filter(val => typeof val === 'number');
            const sum = numbers.reduce((a, b) => a + b, 0);
            
            // Insert the sum into the target cell
            await insertFormula(sum.toString(), action.params.targetCell);
          }
        }
        break;
    }
  };

  return (
    <Stack tokens={{ childrenGap: 10, padding: 10 }}>
      <Text variant="xLarge">Excel AI Agent</Text>
      
      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(null)}>
          {error}
        </MessageBar>
      )}
      
      <Stack tokens={{ childrenGap: 8 }} style={{ height: '300px', overflowY: 'auto', border: '1px solid #EDEBE9', padding: '8px' }}>
        {messages.map((msg, index) => (
          <Stack 
            key={index}
            styles={{
              root: {
                backgroundColor: msg.role === 'user' ? '#E8F1FB' : '#F3F2F1',
                padding: 10,
                borderRadius: 4,
                maxWidth: '80%',
                alignSelf: msg.role === 'user' ? 'flex-end' : 'flex-start'
              }
            }}
          >
            <Text>{msg.content}</Text>
          </Stack>
        ))}
      </Stack>
      
      <div style={{ height: '1px', backgroundColor: '#EDEBE9', margin: '10px 0' }}></div>
      
      <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { padding: '10px', border: '1px solid #EDEBE9', borderRadius: '4px' }}}>
        <TextField 
          placeholder="Ask about your Excel data..." 
          value={input}
          onChange={(_, newValue) => setInput(newValue || '')}
          disabled={isLoading}
          styles={{ 
            root: { 
              width: '100%', 
              minHeight: '40px' 
            },
            fieldGroup: {
              height: '40px'
            }
          }}
          onKeyDown={(e) => e.key === 'Enter' && !e.shiftKey && handleSend()}
        />
        <PrimaryButton 
          onClick={handleSend} 
          disabled={isLoading || !input.trim()}
          styles={{
            root: {
              minWidth: '80px',
              height: '40px'
            }
          }}
        >
          {isLoading ? <Spinner /> : 'Send'}
        </PrimaryButton>
      </Stack>
      
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton 
          onClick={async () => {
            const success = await testExcelConnection();
            setMessages(prev => [...prev, { 
              role: 'assistant', 
              content: success ? "Excel connection test successful!" : "Excel connection test failed!"
            }]);
          }}
        >
          Test Excel
        </PrimaryButton>
        <PrimaryButton 
          onClick={async () => {
            try {
              // Test writing to a specific cell
              const success = await writeValueToCell(123.45, "C1");
              setMessages(prev => [...prev, { 
                role: 'assistant', 
                content: success ? "Successfully wrote 123.45 to cell C1!" : "Failed to write to Excel!"
              }]);
            } catch (e) {
              console.error("Debug button error:", e);
              setMessages(prev => [...prev, { 
                role: 'assistant', 
                content: "Error: " + e.message
              }]);
            }
          }}
        >
          Write Test Value
        </PrimaryButton>
      </Stack>
      
      <PrimaryButton 
        onClick={async () => {
          try {
            const worksheets = await listAllWorksheets();
            setMessages(prev => [...prev, { 
              role: 'assistant', 
              content: `Worksheets in this workbook: ${worksheets.join(', ')}`
            }]);
          } catch (e) {
            console.error("Worksheet listing error:", e);
            setMessages(prev => [...prev, { 
              role: 'assistant', 
              content: "Error listing worksheets: " + e.message
            }]);
          }
        }}
      >
        List Worksheets
      </PrimaryButton>
      
      <PrimaryButton 
        onClick={async () => {
          try {
            setMessages(prev => [...prev, { 
              role: 'assistant', 
              content: "Attempting to create a test worksheet..."
            }]);
            
            // Create a test worksheet and report success/failure
            const worksheet = await createWorksheet("TestSheet" + Date.now());
            
            if (worksheet) {
              setMessages(prev => [...prev, { 
                role: 'assistant', 
                content: `Successfully created worksheet: ${worksheet.name}`
              }]);
            } else {
              setMessages(prev => [...prev, { 
                role: 'assistant', 
                content: "Failed to create worksheet. Check console for errors."
              }]);
            }
          } catch (e) {
            console.error("Worksheet creation error:", e);
            setMessages(prev => [...prev, { 
              role: 'assistant', 
              content: "Error creating worksheet: " + e.message
            }]);
          }
        }}
      >
        Create Test Sheet
      </PrimaryButton>
    </Stack>
  );
}; 