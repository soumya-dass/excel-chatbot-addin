/* global console, Excel, Office */

import './taskpane.css'; // Import CSS directly into JS

// Import the services directly (for Webpack bundling)
import './dataService.js';
import './aiService.js';
import './uiService.js';

// Enhanced Excel Data Assistant - Main Orchestrator
// Uses DataService, AIService, and UIService for clean separation of concerns

// Service instances
let dataService;
let aiService;
let uiService;

// Initialize Office Add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeServices();
        console.log('Enhanced Excel Chatbot initialized successfully');
    }
});

// Initialize all services and set up the application
function initializeServices() {
    try {
        // Now you instantiate them directly (no 'window.' prefix)
        dataService = new DataService();
        aiService = new AIService();
        uiService = new UIService();
        
        // ... rest of initialization ...
    } catch (error) {
        console.error('Error initializing services:', error);
        if (uiService) {
            uiService.showError('Failed to initialize application: ' + error.message);
        }
    }
}

// Main message handling orchestration
async function handleSendMessage() {
    const message = uiService.getUserInput();
    if (!message) return;
    
    // Clear input and disable controls
    uiService.clearUserInput();
    uiService.setControlsEnabled(false);
    
    // Add user message to chat
    uiService.addChatMessage(message, true);
    
    // Add to conversation history
    aiService.addToConversationHistory('user', message);
    
    // Show loading state
    uiService.showLoading(true);
    uiService.updateStatus('Reading Excel data with enhanced structure analysis...');
    
    try {
        // Read Excel data
        const shouldUseSelection = uiService.getUseSelectionState();
        const worksheetData = await dataService.readCurrentWorksheetDataEnhanced(shouldUseSelection);
        
        // Process with AI
        uiService.updateStatus('Analyzing with AI using improved data structure...');
        const response = await aiService.askGeminiWithContext(message, worksheetData);
        
        // Add AI response to conversation history
        aiService.addToConversationHistory('assistant', response);
        
        // Display response
        const conversationLength = aiService.getConversationLength();
        uiService.addChatMessage(response, false, false, conversationLength);
        uiService.updateStatus(`Enhanced analysis complete (${conversationLength} exchanges) - ready for follow-up questions`);
        
    } catch (error) {
        console.error('Error processing message:', error);
        uiService.addChatMessage('Sorry, I encountered an error processing your request. Please try again.', false, true);
        uiService.updateStatus('Error occurred - conversation context maintained');
    } finally {
        // Re-enable controls
        uiService.showLoading(false);
        uiService.setControlsEnabled(true);
        uiService.focusUserInput();
    }
}

// Handle clear conversation
function handleClearConversation() {
    // Clear AI conversation
    aiService.clearConversation();
    
    // Clear UI
    uiService.clearChatMessages();
    uiService.showWelcomeMessage();
    uiService.updateStatus('Conversation cleared - ready for new questions');
    
    // Clear data service cache
    dataService.clearCurrentData();
    
    console.log('Conversation cleared');
}