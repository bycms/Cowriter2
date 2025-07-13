const apiKey = 'sk-tmyrvfoatzrzlmyfsvhbsvpplcoqfdscvnrogsbadkqkvqpf';

// ================= INTENT FUNCTION =================
async function intent(q) {
    await Word.run(async (context) => {
        let p = context.document.body.paragraphs.getFirst();
        p.load("text");
        await context.sync();
        console.log(p.text);
    });
    let response = await fetch('https://api.siliconflow.cn/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${apiKey}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            model: 'THUDM/GLM-4-9B-0414',
            messages: [{
                role: 'user',
                content: `Here is a query from user: ${q}. Judge whether user is going to\n1) Directly ask a question,\n2) Write a passage from scratch,\n3) Edit a passage,\n4) Continue writing. You are **required** to use one of the functions.`
            }],
            tools: [
                {
                    "function": {
                        strict: true,
                        name: "Directly ask",
                        description: `Use this function when user is\n- Mentioning only a noun\n- Asking a usual question.`,
                        parameters: {
                            type: 'object',
                            properties: {
                                'Need_doc_reference': {
                                    type: 'bool',
                                    description: 'Whether the user\'s question requires existing passage as a reference to be answered.'
                                }
                            },
                            required: ['Need_doc_reference']
                        }
                    },
                    type: "function"
                },
                {
                    "function": {
                        strict: true,
                        name: "Draft a passage",
                        description: `Use this function when user wants to \n- Start a new passage, article, story, or essay\n. Keywords include "write" "compose" "draft" "turn... into a passage" etc.`,
                        parameters: {
                            type: 'object',
                            properties: {
                                'Word_range': {
                                    type: 'int',
                                    description: 'The word count of user\'s expected passage. Return 0 if not mentioned.'
                                },
                                'Passage_query': {
                                    type: 'string',
                                    description: 'The general requirements for passage, e.g. Title, Topic, Tone... This string must be in the language of user\'s request, unless user mentioned desired language.'
                                }
                            },
                            required: ['Word_range', 'Passage_query']
                        }
                    }
                },
                {
                    "function": {
                        strict: true,
                        name: "Refine a passage or paragraph",
                        description: `Use this function when user is\n- Providing an exisiting passage or paragraph and suggesting improvements\n- Dissatisfied with a passage or paragraph.\n- Add or remove things from an existing piece of text.`,
                        parameters: {
                            type: 'object',
                            properties: {
                                'Refinement_query': {
                                    type: 'string',
                                    description: 'The general requirements for editing existing passage, e.g. Tone, Length, Format...'
                                }
                            },
                            required: ['Refinement_query']
                        }
                    }
                },
                {
                    "function": {
                        strict: true,
                        name: "Continue writing",
                        description: `Use this function when user wants to continue writing at the end of the document, e.g. "Continue", "Go on", "Write more", "Add more".`,
                    }
                }
            ]
        })
    });

    if (!response.ok) {
        throw new Error(`API request failed with status ${response.status}`);
    }
    const data = await response.json();
    // Parse function call result
    const toolCall = data.choices?.[0]?.message?.tool_calls?.[0];
    if (!toolCall || !toolCall.function) {
        throw new Error('No function call detected in LLM response.');
    }
    let functionName = toolCall.function.name;
    let args = {};
    if (toolCall.function.arguments) {
        try {
            args = JSON.parse(toolCall.function.arguments);
        } catch (e) {
            // fallback: arguments may already be an object
            args = toolCall.function.arguments;
        }
    }
    return { functionName, arguments: args };
}

async function handleDirectAsk(query, needContext, displayCallback) {
    // Call the LLM for a direct answer with streaming
    if (needContext) {
        await Word.run(async (context) => {
            let body = context.document.body;
            body.load('text');
            await context.sync();
            query += ` Here's the passage for reference: ${body.text}`;
        })
    }
    const response = await fetch('https://api.siliconflow.cn/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${apiKey}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            model: 'Qwen/Qwen2.5-7B-Instruct',
            messages: [
                { role: 'user', content: query }
            ],
            stream: true
        })
    });
    if (!response.ok) {
        throw new Error(`API request failed with status ${response.status}`);
    }
    // Streaming response - ONLY display in chat, don't insert into Word
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let fullResponse = '';
    let buffer = '';
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop();
        for (const line of lines) {
            if (line.trim() === '') continue;
            if (line.startsWith('data: ')) {
                const data = line.substring(6);
                if (data === '[DONE]') {
                    displayCallback(md(fullResponse));
                    return;
                }
                try {
                    const parsed = JSON.parse(data);
                    if (parsed.choices && parsed.choices[0].delta && parsed.choices[0].delta.content) {
                        const content = parsed.choices[0].delta.content;
                        fullResponse += content;
                        displayCallback(md(fullResponse));
                    }
                } catch (e) {
                    console.error('Error parsing JSON:', e);
                }
            }
        }
    }
}

// ================= HANDLE DRAFT PASSAGE =================
async function handleDraftPassage(query, wordRange, displayCallback) {
    // Prepare chat history for context
    let historyForModel = messageHistory.map(msg => ({ role: msg.role, content: msg.content }));
    // Add user request as last message
    historyForModel.push({ role: 'user', content: query + `
1. LENGTH: ${wordRange} words (${wordRange ? 'strictly adhere' : 'aim for 200-800 words'})
2. OUTPUT REQUIREMENTS:
   - Return ONLY the generated passage in **HTML** format with inline styles
   - Use explicit style attributes like: style="color:#2b579a; font-size:14pt; font-family:'Times New Roman'"
   - Example: 
        <h1 style="font-size:18pt; color:#1a365d; font-family:'Times New Roman'">Market Analysis</h1>
        <p style="font-size:12pt; color:#111111; font-family:'Times New Roman'">The global economy <strong>shows steady growth</strong>... </p>
        <h2 style="font-size:16pt; color:#2b579a; font-family:'Times New Roman'">Key Trends</h2>
        <p style="font-size:12pt; color:#111111; font-family:'Times New Roman'">
          1. <span style="color:#d47500">Technology</span><br>
          2. <span style="color:#00724d">Healthcare</span>
        </p>

3. STRUCTURE RULES:
   - Headings: 
     - h1: 16-20pt | Bold | Dark color (e.g. #1a365d)
     - h2: 14-16pt | Semi-bold | Slightly lighter (e.g. #2b579a)
   - Paragraphs: 
     - 11-13pt font size 
     - Line spacing: 1.15
     - Default black (#111111) unless accent needed
   - Fonts:
     - Western: 'Times New Roman' (primary), 'Arial' (accent)
     - Chinese: 'Microsoft YaHei'
     - Code: 'Consolas'

4. STYLING CONSTRAINTS:
   - MAXIMUM:
     - 2 accent colors for all except headings (1 recommended)
     - 3 font variations (size/weight/color combinations)
   - NEVER use:
     - Background colors
     - Underlines (except hyperlinks)
     - More than 2 heading levels
     - More than 3 subheadings
     - Font sizes <10pt or >24pt

5. LANGUAGE & TONE:
   - Match the user's request language exactly
   - Maintain professional academic/business formatting
   - Style annotations must be hidden in final document
            ` });
    const response = await fetch('https://api.siliconflow.cn/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${apiKey}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            model: 'Qwen/Qwen2.5-7B-Instruct',
            messages: historyForModel,
            stream: true
        })
    });
    if (!response.ok) {
        throw new Error(`API request failed with status ${response.status}`);
    }
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let fullPassage = '';
    let buffer = '';
    let htmlBuffer = '';
    let isFirstInsert = true;
    
    // Define block-level HTML tags
    const blockTags = ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'div', 'pre', 'blockquote'];
    const blockCloseRegex = new RegExp(`</(${blockTags.join('|')})>`, 'gi');
    
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop();
        for (const line of lines) {
            if (line.trim() === '') continue;
            if (line.startsWith('data: ')) {
                const data = line.substring(6);
                if (data === '[DONE]') {
                    // Insert any remaining HTML
                    if (htmlBuffer.trim()) {
                        try {
                            await Word.run(async (context) => {
                                let range;
                                if (isFirstInsert) {
                                    range = context.document.getSelection();
                                    range.insertHtml(htmlBuffer, Word.InsertLocation.replace);
                                    isFirstInsert = false;
                                } else {
                                    range = context.document.body;
                                    range.insertHtml('<br/>' + htmlBuffer, Word.InsertLocation.end);
                                }
                                await context.sync();
                            });
                        } catch (e) {
                            console.error('Word API error (final HTML insert):', e);
                        }
                    }
                    displayCallback('<span style="color:green">Passage inserted into document. Please review and confirm.</span>');
                    // Append generated passage to chat history as system message
                    messageHistory.push({ role: 'system', content: fullPassage });
                    return;
                }
                try {
                    const parsed = JSON.parse(data);
                    if (parsed.choices && parsed.choices[0].delta && parsed.choices[0].delta.content) {
                        const content = parsed.choices[0].delta.content;
                        fullPassage += content;
                        htmlBuffer += content;
                        
                        // Find all complete blocks
                        let lastIndex = 0;
                        let match;
                        while ((match = blockCloseRegex.exec(htmlBuffer)) !== null) {
                            const endPos = match.index + match[0].length;
                            const blockContent = htmlBuffer.substring(lastIndex, endPos);
                            lastIndex = endPos;
                            
                            try {
                                await Word.run(async (context) => {
                                    let range;
                                    if (isFirstInsert) {
                                        range = context.document.getSelection();
                                        range.insertHtml(blockContent, Word.InsertLocation.replace);
                                        isFirstInsert = false;
                                    } else {
                                        range = context.document.body;
                                        range.insertHtml('<br/>' + blockContent, Word.InsertLocation.end);
                                    }
                                    await context.sync();
                                });
                            } catch (e) {
                                console.error('Word API error (HTML insert):', e);
                            }
                        }
                        
                        // Keep the remaining buffer
                        htmlBuffer = htmlBuffer.substring(lastIndex);
                    }
                } catch (e) {
                    console.error('Error parsing JSON:', e);
                }
            }
        }
    }
}

// ================= HANDLE REFINE PASSAGE =================
async function handleRefinePassage(refinementQuery, displayCallback) {
    // Get the selected text or default to the whole document
    let selectedText = '';
    let selectionIsWholeDoc = false;
    await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load('text');
        await context.sync();
        selectedText = range.text;
        if (!selectedText) {
            context.document.body.load('text');
            await context.sync();
            selectedText = context.document.body.text;
            selectionIsWholeDoc = true;
        }
    });
    // Call the LLM to refine the passage or selected paragraph, with styling requirements
    const refinePrompt = `REFINEMENT DIRECTIVES:

1. CONTENT PRESERVATION:
   - Maintain original core meaning and length (±10%) unless explicitly requested otherwise
   - Preserve all key facts, names, and technical terms

2. OUTPUT REQUIREMENTS:
   - Output whole passage if ORIGINAL TEXT is a complete passage; otherwise output only the refined selected paragraph
   - Return ONLY the refined part in **HTML** format with inline styles
   - Use style attributes: style="color:#2b579a; font-size:14pt; font-family:'Times New Roman'"
   - COMPLETE output must match original scope (paragraph/section/document)
   - NEVER output placeholders or partial content

3. STYLING RULES:
   - Hierarchy:
     • h1: 16-20pt | Bold | Dark primary color (#1a365d)
     • h2: 14-16pt | Semi-bold | Secondary color (#2b579a)
     • Body: 11-13pt | #111111 (black) default
   - Fonts:
     • Western: 'Times New Roman' (primary), 'Arial' (accent)
     • Chinese: 'Microsoft YaHei'
   - Constraints:
     • MAX 1 accent color
     • NO background colors or underlines
     • Font sizes strictly 10-24pt range

4. REFINEMENT PRIORITIES:
   [1] Correctness → [2] Clarity → [3] Conciseness → [4] Style
   - Highlight major changes with ††explanation†† when non-obvious

5. LANGUAGE:
   - Match original text language precisely
   - Maintain consistent terminology
   - Preserve regional variants (e.g. British vs American English)

ORIGINAL TEXT (preserve formatting markers if present):
"""
${selectedText}
"""

REFINEMENT REQUEST:
"${refinementQuery}"

OUTPUT FORMAT EXAMPLE:
<h1 style="font-size:18pt; color:#1a365d">Revised Section</h1>
`;
    const response = await fetch('https://api.siliconflow.cn/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${apiKey}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            model: 'Qwen/Qwen2.5-7B-Instruct',
            messages: [
                { role: 'user', content: refinePrompt }
            ],
            stream: true
        })
    });
    if (!response.ok) {
        throw new Error(`API request failed with status ${response.status}`);
    }
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let fullPassage = '';
    let buffer = '';
    let htmlBuffer = '';  // Buffer for HTML content
    let isFirstInsert = true;
    
    // Define block-level HTML tags
    const blockTags = ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'div', 'pre', 'blockquote'];
    const blockCloseRegex = new RegExp(`</(${blockTags.join('|')})>`, 'gi');

    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop();
        for (const line of lines) {
            if (line.trim() === '') continue;
            if (line.startsWith('data: ')) {
                const data = line.substring(6);
                if (data === '[DONE]') {
                    // Insert any remaining HTML
                    if (htmlBuffer.trim()) {
                        try {
                            await Word.run(async (context) => {
                                let range;
                                if (isFirstInsert) {
                                    range = context.document.body;
                                    range.load('text');
                                    await context.sync();
                                    range.insertHtml(htmlBuffer, Word.InsertLocation.replace);
                                    isFirstInsert = false;
                                } else {
                                    range = context.document.body;
                                    range.insertHtml('<br/>' + htmlBuffer, Word.InsertLocation.end);
                                }
                                await context.sync();
                            });
                        } catch (e) {
                            console.error('Word API error (final HTML insert):', e);
                        }
                    }
                    displayCallback('<span style="color:green">Passage inserted into document. Please review and confirm.</span>');
                    messageHistory.push({ role: 'system', content: fullPassage });
                    return;
                }
                try {
                    const parsed = JSON.parse(data);
                    if (parsed.choices && parsed.choices[0].delta && parsed.choices[0].delta.content) {
                        const content = parsed.choices[0].delta.content;
                        fullPassage += content;
                        htmlBuffer += content;

                        // Find all complete blocks
                        let lastIndex = 0;
                        let match;
                        while ((match = blockCloseRegex.exec(htmlBuffer)) !== null) {
                            const endPos = match.index + match[0].length;
                            const blockContent = htmlBuffer.substring(lastIndex, endPos);
                            lastIndex = endPos;

                            try {
                                await Word.run(async (context) => {
                                    let range;
                                    if (isFirstInsert) {
                                        range = context.document.body;
                                        range.load('text');
                                        await context.sync();
                                        range.insertHtml(blockContent, Word.InsertLocation.replace);
                                        isFirstInsert = false;
                                    } else {
                                        range = context.document.body;
                                        range.insertHtml('<br/>' + blockContent, Word.InsertLocation.end);
                                    }
                                    await context.sync();
                                });
                            } catch (e) {
                                console.error('Word API error (HTML insert):', e);
                            }
                        }

                        // Keep the remaining buffer
                        htmlBuffer = htmlBuffer.substring(lastIndex);
                    }
                } catch (e) {
                    console.error('Error parsing JSON:', e);
                }
            }
        }
    }
}

// ================= HANDLE CONTINUE WRITING =================
async function handleContinueWriting(continueQuery, displayCallback) {
    // Prepare chat history for context
    let historyForModel = messageHistory.map(msg => ({ role: msg.role, content: msg.content }));
    historyForModel.push({ role: 'user', content: continueQuery });
    let docText = '';
    await Word.run(async (context) => {
        const body = context.document.body;
        body.load('text');
        await context.sync();
        docText = body.text;
    });
    // Call the LLM to continue writing at the end
    const response = await fetch('https://api.siliconflow.cn/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${apiKey}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            model: 'Qwen/Qwen2.5-7B-Instruct',
            messages: historyForModel.concat([{
                role: 'system',
                content: `Continue writing at the end of the following document.\n**This is original document content: **${docText}\nRequirements: ${continueQuery}\n- Output **only** your generated content in HTML, connecting the last word of existing document.\n- Write in the language of user's request if not specified.\n- Use clean, professional styling with minimal colors\n- Default to black text (#111111) unless specified.\n- Use only ONE accent color if needed.\n- Set paragraph font size to 12pt.\n- Title: 18pt bold, Subheadings: 14pt bold.\n- Do **not** use lists unless explicitly requested.\n- No more than 2 subtitles unless requested.\n- Add <br/> between <li>s.\n- Never use background colors.\n- Default fonts are: Times New Roman for Latin, Microsoft Yahei for Chinese.`
            }]),
            stream: true
        })
    });
    if (!response.ok) {
        throw new Error(`API request failed with status ${response.status}`);
    }
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let fullContent = '';
    let buffer = '';
    let htmlBuffer = '';
    
    // Define block-level HTML tags
    const blockTags = ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'div', 'pre', 'blockquote'];
    const blockCloseRegex = new RegExp(`</(${blockTags.join('|')})>`, 'gi');
    
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop();
        for (const line of lines) {
            if (line.trim() === '') continue;
            if (line.startsWith('data: ')) {
                const data = line.substring(6);
                if (data === '[DONE]') {
                    // Insert any remaining HTML at the end
                    if (htmlBuffer.trim()) {
                        try {
                            await Word.run(async (context) => {
                                const body = context.document.body;
                                body.insertHtml(htmlBuffer, Word.InsertLocation.end);
                                await context.sync();
                            });
                        } catch (e) {
                            console.error('Word API error (final HTML insert):', e);
                        }
                    }
                    displayCallback('<span style="color:green">Content continued at the end of the document.</span>');
                    // Append generated passage to chat history as system message
                    messageHistory.push({ role: 'system', content: fullContent });
                    return;
                }
                try {
                    const parsed = JSON.parse(data);
                    if (parsed.choices && parsed.choices[0].delta && parsed.choices[0].delta.content) {
                        const content = parsed.choices[0].delta.content;
                        fullContent += content;
                        htmlBuffer += content;
                        
                        // Find all complete blocks
                        let lastIndex = 0;
                        let match;
                        while ((match = blockCloseRegex.exec(htmlBuffer)) !== null) {
                            const endPos = match.index + match[0].length;
                            const blockContent = htmlBuffer.substring(lastIndex, endPos);
                            lastIndex = endPos;
                            
                            try {
                                await Word.run(async (context) => {
                                    const body = context.document.body;
                                    body.insertHtml(blockContent, Word.InsertLocation.end);
                                    await context.sync();
                                });
                            } catch (e) {
                                console.error('Word API error (HTML insert):', e);
                            }
                        }
                        
                        // Keep the remaining buffer
                        htmlBuffer = htmlBuffer.substring(lastIndex);
                    }
                } catch (e) {
                    console.error('Error parsing JSON:', e);
                }
            }
        }
    }
}

// ================= CHAT HISTORY AND MARKDOWN =================
let messageHistory = [];

// ================= RENDER MESSAGES =================
function renderMessages(chatMessagesElem) {
    chatMessagesElem.innerHTML = '';
    for (const msg of messageHistory) {
        const msgDiv = document.createElement('div');
        msgDiv.className = msg.role === 'user' ? 'chat-message user-message' : 'chat-message ai-message';
        msgDiv.innerHTML = msg.content;
        chatMessagesElem.appendChild(msgDiv);
    }
    chatMessagesElem.scrollTop = chatMessagesElem.scrollHeight;
}

function updateLastAIMessage(content, chatMessagesElem) {
    if (!messageHistory.length || messageHistory[messageHistory.length-1].role !== 'assistant') {
        messageHistory.push({ role: 'assistant', content: '' });
    }
    messageHistory[messageHistory.length-1].content = content;
    renderMessages(chatMessagesElem);
}

// ================= CALL AI STREAMING (from chatFlow.js) =================
async function callAIStreaming(q, request_type, onChunkReceived) {
    try {
        // Prepare messages based on request type
        let messages;
        if (request_type === 'conversation') {
            // For conversations, use the full message history
            messages = [
                ...messageHistory.map(msg => ({ role: msg.role, content: msg.content })),
                { role: 'user', content: q }
            ];
        } else {
            // For other request types (to be implemented later), use just the current query
            messages = [{ role: 'user', content: q }];
        }
        const response = await fetch('https://api.siliconflow.cn/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${apiKey}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                model: 'Qwen/Qwen2.5-7B-Instruct',
                messages: messages,
                stream: true
            })
        });
        if (!response.ok) {
            throw new Error(`API request failed with status ${response.status}`);
        }
        // Process the streaming response
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let fullResponse = '';
        let buffer = '';
        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop(); // Keep incomplete line in buffer
            for (const line of lines) {
                if (line.trim() === '') continue;
                if (line.startsWith('data: ')) {
                    const data = line.substring(6);
                    if (data === '[DONE]') {
                        // Add the complete response to history
                        messageHistory.push({ role: 'assistant', content: fullResponse });
                        return;
                    }
                    try {
                        const parsed = JSON.parse(data);
                        if (parsed.choices && parsed.choices[0].delta && parsed.choices[0].delta.content) {
                            const content = parsed.choices[0].delta.content;
                            fullResponse += content;
                            onChunkReceived(fullResponse);
                        }
                    } catch (e) {
                        console.error('Error parsing JSON:', e);
                    }
                }
            }
        }
    } catch (error) {
        console.error('Error in callAIStreaming:', error);
        throw error;
    }
}

// Office Add-in Chat UI for Microsoft Word
(function() {
    // Wait for Office to initialize
    Office.onReady()
        .then(function() {
            console.log("Office.js is ready for Word");
            initializeChatUI();
        })
        .catch(function(error) {
            console.error("Office.js initialization error:", error);
            document.getElementById("root").innerHTML = `
                <div class="error-message">
                    Failed to initialize Word add-in. Please try reloading this add-in.
                </div>
            `;
        });

    function initializeChatUI() {
        const root = document.getElementById("root");
        if (!root) {
            console.error("Root element not found");
            return;
        }

        // Initialize markdown-it if not already done
        if (typeof markdownit !== 'undefined' && !markdown) {
            markdown = new markdownit();
        }
        // Create UI structure
        root.innerHTML = `
            <div id="chat-div" class="chat-div">
                <div class="chat-messages" id="chat-messages"></div>
                <div class="loader" id="loader" style="display:none;">
                  <span class="dot"></span><span class="dot"></span><span class="dot"></span>
                </div>
            </div>
            <div id="input-div" class="input-div">
                <div id="selection-indicator"></div>
                <textarea id="user-input" placeholder="Ask me anything..." class="user-input"></textarea>
                <button id="send-btn" class="send-btn">Send</button>
            </div>
        `;

        // Get DOM elements after innerHTML is set
        const input = root.querySelector('#user-input');
        const sendBtn = root.querySelector('#send-btn');
        const chatMessages = root.querySelector('#chat-messages');
        const loader = root.querySelector('#loader');

        function setLoading(isLoading) {
            if (isLoading) {
                loader.style.display = '';
                sendBtn.disabled = true;
                input.disabled = true;
                sendBtn.textContent = 'Processing...';
            } else {
                loader.style.display = 'none';
                sendBtn.disabled = false;
                input.disabled = false;
                sendBtn.textContent = 'Send';
            }
        }

        // Helper for phase messages
        let phaseMsgId = null;
        function showPhase(msg) {
            // Remove previous phase message
            if (phaseMsgId !== null) {
                const prev = document.getElementById(phaseMsgId);
                if (prev) prev.remove();
            }
            phaseMsgId = 'phase-' + Date.now();
            const div = document.createElement('div');
            div.id = phaseMsgId;
            div.className = 'chat-message ai-message';
            div.style.color = 'grey';
            div.innerText = msg;
            chatMessages.appendChild(div);
            chatMessages.scrollTop = chatMessages.scrollHeight;
        }
        function removePhase() {
            if (phaseMsgId !== null) {
                const prev = document.getElementById(phaseMsgId);
                if (prev) prev.remove();
                phaseMsgId = null;
            }
        }

        async function handleSendMessage() {
            const value = input.value.trim();
            if (!value) return;
            setLoading(true);
            // Remove any unfinished assistant message before sending a new user message
            if (messageHistory.length && messageHistory[messageHistory.length-1].role === 'assistant' && !messageHistory[messageHistory.length-1].content.trim()) {
                messageHistory.pop();
            }
            // Add user message to history and display
            messageHistory.push({ role: 'user', content: value });
            renderMessages(chatMessages);
            input.value = '';
            try {
                showPhase('Deciding what to do');
                const intentResult = await intent(value);
                console.log(intentResult);
                // intentResult should contain functionName and arguments
                let functionName = intentResult.functionName;
                let args = intentResult.arguments || {};
                removePhase();
                // Route to the correct handler
                if (functionName === 'Directly ask') {
                    showPhase('Generating an answer');
                    await handleDirectAsk(value, args.Need_doc_reference, (result) => {
                        removePhase();
                        updateLastAIMessage(result, chatMessages);
                    });
                } else if (functionName === 'Draft a passage') {
                    showPhase('Generating your passage');
                    await handleDraftPassage(args.Passage_query || value, args.Word_range ? "about" + args.Word_range + " words" : "suitable length (above 200 words, below 800 words)", (result) => {
                        removePhase();
                        updateLastAIMessage(result, chatMessages);
                    });
                } else if (functionName === 'Refine a passage or paragraph') {
                    showPhase('Editing passage');
                    await handleRefinePassage(args.Refinement_query || value, (result) => {
                        removePhase();
                        updateLastAIMessage(result, chatMessages);
                    });
                } else if (functionName === 'Continue writing') {
                    showPhase('Continuing writing');
                    await handleContinueWriting(args.Continue_query || value, (result) => {
                        removePhase();
                        updateLastAIMessage(result, chatMessages);
                    });
                } else {
                    await handleDirectAsk(value, true, (result) => {
                        removePhase();
                        updateLastAIMessage(result, chatMessages);
                    });
                }
                showPhase('Finishing up');
                setTimeout(removePhase, 600);
            } catch (error) {
                removePhase();
                updateLastAIMessage(`<p style=\"color: red;\">Error: ${error.message}</p>`, chatMessages);
                console.error("AI call error:", error);
            } finally {
                setLoading(false);
            }
        }

        sendBtn.addEventListener('click', handleSendMessage);
        input.addEventListener('keydown', function(e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                handleSendMessage();
            }
        });

        input.addEventListener('input', function() {
            sendBtn.disabled = !input.value.trim();
        });
        sendBtn.disabled = true;

        // Initial welcome message
        if (messageHistory.length === 0) {
            messageHistory.push({ role: 'assistant', content: 'Welcome to Cowriter. Start by applying a suggestion or asking a question.' });
        }

        renderMessages(chatMessages);
    }

})();

const intervalId = setInterval(a, 5000);

// Define the function to be called periodically
function a() {
    Word.run(function (context) {
        // Get the current selection
        var selection = context.document.getSelection();
        selection.load("text");

        return context.sync()
            .then(function () {
                const indicator = document.getElementById('selection-indicator');
                if (selection.text) {
                    indicator.innerHTML = "<h3>You selected</h3><p>" + selection.text + "</p>";
                    indicator.classList.add('visible');
                }
                else {
                    indicator.classList.remove('visible');
                    indicator.textContent = '';
                }
            });
    }).catch(function (error) {
        console.log("Error in function a(): " + JSON.stringify(error));
    });
}

// ================= TOOL FUNCTION: CONVERT STYLED MARKDOWN TO HTML =================
function md_paragraph(s) {
    if (!s) return '';
    
    // Preserve all original newlines first
    const lines = s.split('\n');
    let output = [];

    for (let line of lines) {
        line = line.trim();
        if (!line) {
            output.push('<br>'); // Preserve empty lines as breaks
            continue;
        }

        // Process style annotations {color:blue; font-size:14pt}
        let processedLine = line.replace(/([^{}\n]*?)\s*\{([^}]+)\}/g, (_, text, style) => {
            // Clean and validate styles
            const validStyles = style.split(';').filter(part => {
                const [prop, value] = part.split(':').map(s => s.trim());
                return prop && value && /^(color|font-size|font-family|font-weight|line-height)/.test(prop);
            }).join(';');
            return validStyles ? `<span style="${validStyles}">${text}</span>` : text;
        });

        // Escape HTML (except within tags)
        processedLine = processedLine.replace(/[&<>](?![^<]*>)/g, m => 
            ({'&':'&amp;','<':'&lt;','>':'&gt;'}[m]));

        // Process headings (preserve original spacing)
        if (/^#+\s/.test(line)) {
            const level = Math.min((line.match(/#/g) || []).length, 6);
            const headingText = line.replace(/^#+\s*/, '').replace(/\s*\{[^}]*\}$/, '');
            processedLine = `<h${level}>${headingText}</h${level}>`;
        }
        // Process lists (preserve indentation)
        else if (/^(\s*)[-*+] /.test(line)) {
            const indent = (line.match(/^\s*/) || [''])[0];
            processedLine = indent + '<li>' + line.replace(/^(\s*)[-*+] /, '') + '</li>';
        }
        // Process other block elements
        else if (/^(?:```|>)/.test(line)) {
            // No transformation for code blocks/blockquotes (handled later)
        }
        // Default paragraph handling (preserve leading/trailing spaces)
        else {
            processedLine = line.replace(/^(\s*)(.*?)(\s*)$/, (_, lead, content, trail) => {
                return lead + '<p>' + content + '</p>' + trail;
            });
        }

        output.push(processedLine);
    }

    // Reconstruct with proper spacing
    return output.join('\n');
}

// Full markdown renderer for UI display
function md(s) {
    if (!s) return '';
    
    // First preserve all newlines and spaces
    const paragraphs = s.split(/\n\n+/);
    let output = [];
    
    for (let para of paragraphs) {
        // Process style annotations
        let processedPara = para.replace(/([^{}\n]*?)\s*\{([^}]+)\}/g, (_, text, style) => {
            return `<span style="${style}">${text}</span>`;
        });

        // Escape HTML
        processedPara = processedPara.replace(/[&<>]/g, m => 
            ({'&':'&amp;','<':'&lt;','>':'&gt;'}[m]));

        // Process block elements with spacing preservation
        if (/^#+\s/.test(processedPara)) {
            const level = Math.min((processedPara.match(/#/g) || []).length, 6);
            const headingText = processedPara.replace(/^#+\s*/, '').replace(/\s*\{[^}]*\}$/, '');
            output.push(`<h${level}>${headingText}</h${level}>`);
        }
        else if (/^```/.test(processedPara)) {
            output.push('<pre><code>' + processedPara.replace(/^```|```$/g, '') + '</code></pre>');
        }
        else if (/^> /.test(processedPara)) {
            output.push('<blockquote>' + processedPara.replace(/^> /gm, '') + '</blockquote>');
        }
        else if (/^[-*+] /.test(processedPara)) {
            output.push('<ul>' + processedPara.replace(/^[-*+] /gm, '<li>').replace(/\n/g, '</li><li>') + '</li></ul>');
        }
        else if (/^\d+\. /.test(processedPara)) {
            output.push('<ol>' + processedPara.replace(/^\d+\. /gm, '<li>').replace(/\n/g, '</li><li>') + '</li></ol>');
        }
        else {
            // Preserve internal line breaks
            processedPara = processedPara.replace(/\n/g, '<br>');
            output.push(`<p>${processedPara}</p>`);
        }
    }

    return output.join('\n\n');
}