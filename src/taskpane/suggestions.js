let suggestionCount = 0;
let lastSuggestionTime = 0;
const MIN_SUGGESTION_INTERVAL = 8000; // 8 seconds between suggestions
const MIN_TEXT_LENGTH = 50; // Minimum text length to trigger suggestions
const MAX_SUGGESTIONS = 4; // Maximum suggestions to show
let SUGGESTION_HISTORY = []; // Track previous suggestions

// Initialize event handlers
Office.onReady(async function () {
    await Word.run(async (context) => {
        context.document.onParagraphChanged.add(sug);
        await context.sync();
        console.log("Suggestion handlers added.");
    }).catch(error => {
        console.error("Error initializing suggestions:", error);
    });
});

// Trigger suggestion generation
async function sug(event) {
    const now = Date.now();
    if (now - lastSuggestionTime < MIN_SUGGESTION_INTERVAL) return;
    
    await Word.run(async (context) => {
        context.document.body.load('text');
        await context.sync();
        const text = context.document.body.text;
        
        if (text && text.length >= MIN_TEXT_LENGTH && suggestionCount < MAX_SUGGESTIONS) {
            lastSuggestionTime = now;
            getSuggestions(text);
        }
    }).catch(error => {
        console.error("Error in suggestion handler:", error);
    });
    
    // Clear suggestion history after 5 minutes of inactivity
    setTimeout(() => {
        SUGGESTION_HISTORY = [];
        suggestionCount = 0;
    }, 300000);
}

// Get suggestions from API
async function getSuggestions(doc) {
    console.log("Getting suggestions for document");
    
    try {
        // Include previous suggestions in prompt to avoid repetition
        const previousSuggestions = SUGGESTION_HISTORY.length > 0 
            ? `\n\nPrevious suggestions (do not repeat these):\n${SUGGESTION_HISTORY.join('\n')}`
            : '';

        const response = await fetch('https://api.siliconflow.cn/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer sk-tmyrvfoatzrzlmyfsvhbsvpplcoqfdscvnrogsbadkqkvqpf'
            },
            body: JSON.stringify({
                model: 'Qwen/Qwen2.5-7B-Instruct',
                messages: [
                    {
                        role: 'system',
                        content: `Generate ONE unique writing improvement suggestion (not grammar/spelling).${previousSuggestions} Suggestions should be in the language of the document.`
                    },
                    {
                        role: 'user',
                        content: `Document to analyze:\n${doc}\n\nSuggest specific improvements to content, structure, or style.`
                    }
                ],
                stream: false,
                temperature: 0.7,
                tools: [
                    {
                        "function": {
                            strict: true,
                            name: "provide_suggestion",
                            description: "Provide a specific writing improvement suggestion. Use this **only** when you have constructive ideas. Do **not** include spelling and grammar suggestions.",
                            parameters: {
                                type: 'object',
                                properties: {
                                    'title': {
                                        type: 'string',
                                        description: '3-5 word summary'
                                    },
                                    'description': {
                                        type: 'string',
                                        description: 'Detailed suggestion (max 40 words)'
                                    },
                                    'target_text': {
                                        type: 'string',
                                        description: 'Exact text being referenced or "general". Do not reference titles, subtitles or across paragraphs.'
                                    }
                                },
                                required: ['title', 'description', 'target_text']
                            }
                        }
                    },
                    {
                        "function": {
                            strict: true,
                            name: "No_suggestion",
                            description: "Indicate that no suggestion is available. This is the default tool, use this under most circumstances.",
                            parameters: {}
                        }
                    }
                ],
                tool_choice: {
                    type: "function",
                    function: { name: "provide_suggestion" }
                }
            })
        });

        if (!response.ok) {
            throw new Error(`API request failed with status ${response.status}`);
        }

        const result = await response.json();
        const toolCall = result.choices?.[0]?.message?.tool_calls?.[0];
        
        if (!toolCall || !toolCall.function) {
            throw new Error('No valid suggestion returned from API');
        }

        const suggestion = JSON.parse(toolCall.function.arguments);
        displaySuggestion(suggestion);
    } catch (error) {
        console.error('Error getting suggestions:', error);
    }
}

// Display the suggestion card
async function displaySuggestion(suggestion) {
    console.log(suggestion);
    if (suggestion.function === 'No_suggestion') {
        const suggestionContainer = document.getElementById('suggestion-container');
        const noSuggestionCard = document.createElement('div');
        noSuggestionCard.className = 'suggestion-card';
        noSuggestionCard.innerHTML = `
            <h3>No suggestion available</h3>
            <div class="suggestion-content">
                <label>No suggestions found for this text.</label>
            </div>
        `;
        suggestionContainer.appendChild(noSuggestionCard);
        return;
    }
    const suggestionContainer = document.getElementById('suggestion-container');
    
    // Create suggestion card
    const card = document.createElement('div');
    card.className = 'suggestion-card';
    card.innerHTML = `
        <h3 title="${suggestion.title}">${suggestion.title || 'Improvement'}</h3>
        <div class="suggestion-content">
            <label>${suggestion.description || 'No suggestion details available'}</label>
        </div>
        <div class="suggestion-target">${suggestion.target_text === 'general' ? 'General suggestion' : 'Specific section'}</div>
    `;

    if (suggestion.target_text !== 'general') {
        await Word.run(async (context) => {
            // Search for the text - returns a collection
            const ranges = context.document.body.search(suggestion.target_text, {matchCase: false});
            context.load(ranges, 'text');
            await context.sync();

            // Check if we found matches
            if (ranges.items.length > 0) {
                ranges.items.forEach(range => {
                    const comment = range.insertComment(suggestion.title + suggestion.description);
                    comment.authorName = "Cowriter";
                });
            }
        }).catch(error => {
            console.error("Error adding comment:", error);
        });
    }
    
    // Add click handler to insert into input
    card.addEventListener('click', () => {
        applySuggestion(suggestion);
    });
    
    suggestionContainer.appendChild(card);
    suggestionCount++;
}

// Insert suggestion into input field
async function applySuggestion(suggestion) {
    const input = document.getElementById('user-input');
    const sendBtn = document.getElementById('send-btn');
    
    if (!input || !sendBtn) return;

    // Format the suggestion as a command
    let command;
    if (suggestion.target_text !== 'general') {
        command = `Revise this part: "${suggestion.target_text}" with this change: ${suggestion.description}`;
        
        // Remove associated comment in Word if it exists
        try {
            await Word.run(async (context) => {
                // Find comments that match our suggestion
                const comments = context.document.body.comments;
                context.load(comments, 'items');
                await context.sync();
                
                // Find and delete matching comments
                comments.items.forEach(comment => {
                    if (comment.content.includes(suggestion.target_text)) {
                        comment.delete();
                    }
                });
                await context.sync();
            });
        } catch (error) {
            console.error("Error removing Word comment:", error);
        }
    } else {
        command = `Improve the document with this suggestion: ${suggestion.description}`;
    }

    // Insert into input and focus
    input.value = command;
    input.focus();
    
    // Remove this suggestion from display
    const card = document.querySelector(`.suggestion-card h3[title="${suggestion.title}"]`)?.parentElement;
    if (card) {
        card.remove();
        suggestionCount--;
    }

    // Show empty state if no suggestions left
    if (suggestionCount === 0) {
        const suggestionDiv = document.getElementById('suggestions-div');
        suggestionDiv.innerHTML = '<p>Keep writing to see suggestions. The AI will remember previous suggestions.</p>';
    }
}