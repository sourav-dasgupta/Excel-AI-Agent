import axios from 'axios';
import { Message } from '../components/App';
import { ExcelData } from './excel-service';

// Use environment variable for API key
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || '';
const API_URL = "https://api.openai.com/v1/chat/completions";

// If running in development, handle config differently
const getApiKey = () => {
  // Try environment variable first (from webpack)
  if (process.env.OPENAI_API_KEY) {
    return process.env.OPENAI_API_KEY;
  }
  
  // Fallback to global config (if used)
  if (window.ExcelAIConfig?.apiKey) {
    return window.ExcelAIConfig.apiKey;
  }
  
  // No API key found
  console.error("No API key found. Please set OPENAI_API_KEY environment variable.");
  return '';
};

// Use the function to get API key
const apiKey = getApiKey();

export async function getOpenAIResponse(
  messageHistory: Message[], 
  currentQuery: string,
  selectedData?: ExcelData | null,
  isActionRequest: boolean = false
): Promise<string> {
  try {
    // Add this check before making the API call
    if (!apiKey) {
      console.error("OpenAI API key is missing");
      return "Error: API key configuration is missing. Please check the setup.";
    }
    
    // Create system message with Excel assistant context
    const systemMessage = {
      role: "system",
      content: `You are an Excel AI assistant that EXECUTES ACTIONS directly rather than explaining how to do them.
      
      IMPORTANT RULES:
      1. Keep responses under 100 words
      2. When asked to perform an action like summing cells, creating pivot tables, or making charts, reply with CONFIRMATION that you've done it, not instructions.
      3. NEVER give manual steps or formulas unless specifically asked for "how to" instructions
      4. Assume all actions are automatically executed by the assistant, not manually by the user
      
      Example good response: "I've summed the values in the Amount column and placed the result in cell B7."
      Example bad response: "To sum the values, you can use the formula =SUM(B2:B6) in cell B7."
      
      ${selectedData ? 'The user has selected Excel data for you to work with.' : 'Ask the user to select some data first to perform operations on it.'}`
    };
    
    // Format the selected data if available
    let dataContext = "";
    if (selectedData) {
      dataContext = `
      Selected Range: ${selectedData.range}
      Data (${selectedData.rowCount} rows × ${selectedData.columnCount} columns):
      ${JSON.stringify(selectedData.values)}
      ${selectedData.hasHeaders ? 'First row appears to contain headers.' : ''}
      `;
    }
    
    // Update the current query with the data context
    const enhancedQuery = selectedData 
      ? `${currentQuery}\n\nSelected Excel Data:\n${dataContext}`
      : currentQuery;
    
    // Combine all messages for context
    const messages = [
      systemMessage,
      ...messageHistory.slice(-10), // Keep only last 10 messages for context window
      { role: "user", content: enhancedQuery }
    ];
    
    const response = await axios.post(
      API_URL,
      {
        model: "gpt-4-turbo",
        messages,
        temperature: 0.5,
        max_tokens: 800
      },
      {
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${apiKey}`
        }
      }
    );
    
    return response.data.choices[0].message.content;
  } catch (error) {
    console.error("OpenAI API error:", error);
    // More detailed error message
    return `Error: ${error.response?.data?.error?.message || error.message || "Failed to get AI response"}`;
  }
} 