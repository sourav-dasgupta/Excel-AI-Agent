# Excel AI Agent

An intelligent Excel Add-in that provides AI assistance for Excel tasks including formulas, pivot tables, charts, data analysis, and VBA programming through a conversational interface.

## Overview

Excel AI Agent integrates OpenAI's GPT models directly into Excel, allowing users to interact with their spreadsheet data through natural language. Simply select cells, ask questions, and get intelligent assistance directly within Excel.

## Features

### Core Capabilities

- **AI-Powered Chat Interface**: Interact with your Excel data using natural language
- **Data Selection & Analysis**: Select ranges and get AI insights on your data
- **Formula Assistance**: Get help with complex Excel formulas
- **Chart Creation**: Generate appropriate charts based on your selected data
- **Pivot Table Generation**: Create and modify pivot tables through natural language
- **Data Modeling Support**: Get assistance with data model design and relationships
- **Macros & VBA**: Generate VBA code snippets for automation tasks

### Typical Use Cases

- Ask questions about your data: "What's the trend in these sales figures?"
- Request formula help: "Create a formula to calculate the growth rate between columns B and C"
- Generate visualizations: "Create a chart showing the distribution of values in this range"
- Build pivot tables: "Create a pivot table summarizing sales by region and quarter"
- Automate tasks: "Write a VBA macro to format all dates in column A"

## Technical Architecture

- **Frontend**: React with Fluent UI (Microsoft's design system)
- **Add-in Technology**: Office.js API for Excel integration
- **AI Integration**: OpenAI API (GPT-4)
- **Data Processing**: Local data extraction and processing within Excel

## Setup Instructions

### Prerequisites

- Microsoft Excel (desktop or online)
- Node.js and npm
- OpenAI API key

### Installation for Development

1. Clone the repository:
   ```
   git clone [repository-url]
   cd excel-ai-agent
   ```

2. Install dependencies:
   ```
   npm install
   ```

3. Set up OpenAI API key:
   Create a `.env` file in the root directory:
   ```
   OPENAI_API_KEY=your_api_key_here
   ```

4. Start the development server:
   ```

## Environment Setup

This project uses environment variables for configuration:

1. Copy `.env.sample` to a new file named `.env`
2. Replace the placeholder values with your actual API keys
3. Never commit your `.env` file to version control

```bash
# Copy the sample env file
cp .env.sample .env

# Edit the file with your favorite editor
nano .env
```