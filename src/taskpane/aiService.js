// Import Gemini API
import { GoogleGenerativeAI } from '@google/generative-ai';

// AIService - Handles Gemini AI integration, prompts, and conversation management
class AIService {
    constructor() {
        this.CONFIG = {
            // Get API key from environment variable
            GEMINI_API_KEY: process.env.GEMINI_API_KEY,
            GEMINI_MODEL: 'gemini-2.5-flash',
            MAX_HISTORY: 10
        };
        
        this.genAI = null;
        this.model = null;
        this.conversationHistory = [];
        this.chatSession = null;
    }

    // Initialize Gemini AI with chat session
    initializeGemini() {
        try {
            if (!this.CONFIG.GEMINI_API_KEY || this.CONFIG.GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY_HERE') {
                throw new Error('Please configure your Gemini API key in the .env file');
            }
            
            this.genAI = new GoogleGenerativeAI(this.CONFIG.GEMINI_API_KEY);
            this.model = this.genAI.getGenerativeModel({ model: this.CONFIG.GEMINI_MODEL });
            
            this.conversationHistory = [];
            this.chatSession = null;
            
            console.log('Gemini AI initialized with model:', this.CONFIG.GEMINI_MODEL);
            return true;
        } catch (error) {
            console.error('Error initializing Gemini AI:', error);
            throw new Error('Failed to initialize AI service. Please check your API key.');
        }
    }

    // Ask Gemini with enhanced context
    async askGeminiWithContext(userQuestion, currentWorksheetData) {
        if (!this.model) {
            throw new Error('Gemini AI not initialized');
        }
        
        if (!currentWorksheetData) {
            throw new Error('No Excel data available');
        }
        
        try {
            if (!this.chatSession) {
                const systemPrompt = this.createEnhancedSystemPrompt();
                this.chatSession = this.model.startChat({
                    history: [],
                    generationConfig: {
                        maxOutputTokens: 2048,
                        temperature: 0.7,
                    }
                });
                
                await this.chatSession.sendMessage(systemPrompt);
            }
            
            const contextualPrompt = this.createEnhancedContextualPrompt(userQuestion, currentWorksheetData);
            
            const result = await this.chatSession.sendMessage(contextualPrompt);
            const response = await result.response;
            const text = response.text();
            
            if (!text || text.trim().length === 0) {
                throw new Error('Empty response from AI');
            }
            
            return text.trim();
            
        } catch (error) {
            console.error('Gemini API error:', error);
            
            if (error.message.includes('chat') || error.message.includes('session')) {
                this.chatSession = null;
                throw new Error('Chat session reset. Please try your question again.');
            }
            
            if (error.message.includes('API key')) {
                throw new Error('Invalid API key. Please check your Gemini API configuration.');
            } else if (error.message.includes('quota') || error.message.includes('limit')) {
                throw new Error('API quota exceeded. Please try again later.');
            } else {
                throw new Error('AI service error: ' + error.message);
            }
        }
    }

    // Enhanced system prompt
    createEnhancedSystemPrompt() {
        return `You are an advanced Excel Financial Data Assistant with enhanced understanding of complex table structures.

ENHANCED CAPABILITIES:
1. **Row-wise data understanding** - I can read row labels and understand what each row represents
2. **Financial table structure** - I recognize revenue lines, totals, subtotals, and categories
3. **Quarterly data analysis** - I can identify and analyze quarterly patterns (1Q22, 2Q22, etc.)
4. **Proper data relationships** - I understand how row labels relate to column values

ANALYSIS APPROACH:
- Always identify the row label first, then find the corresponding values
- For financial metrics questions, look for rows with matching or similar labels
- Recognize diverse financial categories (revenue, costs, segments, regions, products, etc.)
- Understand time-series progression and trends (quarterly, monthly, yearly)
- Provide specific numbers from the exact rows requested
- Adapt to various financial statement formats and business structures

CONVERSATION MEMORY:
- Maintain context across conversations
- Reference previous analysis when relevant
- Build on earlier insights

Respond with "Enhanced analysis ready!" to confirm.`;
    }

    // Enhanced contextual prompt
    createEnhancedContextualPrompt(userQuestion, currentWorksheetData) {
        const data = currentWorksheetData;
        
        let prompt = `ENHANCED EXCEL DATA ANALYSIS:
Worksheet: ${data.worksheetName}
Range: ${data.address} (${data.totalRows} rows × ${data.totalCols} columns)

`;

        // Add structured data analysis
        if (data.structuredData) {
            const struct = data.structuredData;
            
            prompt += `TABLE STRUCTURE:
- Type: ${struct.type}
- Column Headers: ${struct.columnHeaders.filter(h => h && h !== '').join(' | ')}
- Total Data Rows: ${struct.dataRows.length}
- Key Rows Found: ${struct.keyRows.length}
- Total/Summary Rows: ${struct.totalRows.length}

`;

            // Add quarterly data if available
            if (struct.quarterlyData.length > 0) {
                prompt += `QUARTERLY DATA IDENTIFIED:\n`;
                struct.quarterlyData.forEach(qData => {
                    if (qData.quarters.length > 0) {
                        prompt += `${qData.label}: `;
                        const quarterValues = qData.quarters.map(q => `${q.quarter}=${q.value}`).join(', ');
                        prompt += quarterValues + '\n';
                    }
                });
                prompt += '\n';
            }
            
            // Add key rows with their values
            if (struct.keyRows.length > 0) {
                prompt += `KEY FINANCIAL ROWS:\n`;
                struct.keyRows.forEach(row => {
                    prompt += `"${row.label}"${row.isTotal ? ' [TOTAL]' : ''}: `;
                    const nonEmptyValues = row.values.filter(v => v !== null && v !== undefined && v !== '');
                    prompt += nonEmptyValues.slice(0, 10).join(', '); // Show first 10 values
                    if (nonEmptyValues.length > 10) prompt += `, ... (${nonEmptyValues.length - 10} more)`;
                    prompt += '\n';
                });
                prompt += '\n';
            }
            
            // Add sample of other data rows
            if (struct.dataRows.length > struct.keyRows.length) {
                prompt += `OTHER DATA ROWS (sample):\n`;
                const otherRows = struct.dataRows.filter(row => !struct.keyRows.includes(row)).slice(0, 5);
                otherRows.forEach(row => {
                    prompt += `"${row.label}": `;
                    const nonEmptyValues = row.values.filter(v => v !== null && v !== undefined && v !== '');
                    prompt += nonEmptyValues.slice(0, 5).join(', ');
                    if (nonEmptyValues.length > 5) prompt += `, ... (${nonEmptyValues.length - 5} more)`;
                    prompt += '\n';
                });
                prompt += '\n';
            }
        }
        
        // Add conversation context
        if (this.conversationHistory.length > 2) {
            prompt += `RECENT CONVERSATION:\n`;
            const recentHistory = this.conversationHistory.slice(-4);
            recentHistory.forEach((msg) => {
                if (msg.role === 'user') {
                    prompt += `User: "${msg.content}"\n`;
                } else {
                    const truncated = msg.content.length > 100 ? msg.content.substring(0, 100) + '...' : msg.content;
                    prompt += `Assistant: "${truncated}"\n`;
                }
            });
            prompt += '\n';
        }
        
        prompt += `CURRENT USER QUESTION: "${userQuestion}"

IMPORTANT INSTRUCTIONS:
1. When user asks about specific metrics, find the exact row with matching or similar labels
2. For time-period analysis, look at the time-series data patterns in headers and values
3. Use the specific row labels and values provided - match user queries to actual row names
4. Don't average or aggregate data unless specifically asked - provide actual values from identified rows
5. Consider conversation context and adapt to the specific business/financial domain
6. Recognize various financial statement formats (P&L, Balance Sheet, Cash Flow, etc.)

Please provide a detailed analysis based on this enhanced data structure.`;

        return prompt;
    }

    // Conversation history management
    addToConversationHistory(role, content) {
        this.conversationHistory.push({
            role: role,
            content: content,
            timestamp: new Date().toISOString()
        });
        
        // Trim history if it gets too long
        if (this.conversationHistory.length > this.CONFIG.MAX_HISTORY * 2) {
            this.conversationHistory = this.conversationHistory.slice(-this.CONFIG.MAX_HISTORY * 2);
        }
    }

    // Clear conversation and start fresh
    clearConversation() {
        this.conversationHistory = [];
        this.chatSession = null;
        console.log('AI conversation history cleared');
    }

    // Get conversation history length
    getConversationLength() {
        return Math.floor(this.conversationHistory.length / 2);
    }

    // Get conversation history
    getConversationHistory() {
        return this.conversationHistory;
    }

    // Check if AI is initialized
    isInitialized() {
        return this.model !== null;
    }
}

// Export for use in other modules
if (typeof window !== 'undefined') {
    window.AIService = AIService;
}