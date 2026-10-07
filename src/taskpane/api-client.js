console.log("🔥 api-client.js LOADED SUCCESSFULLY!");

// Initialize conversation history on the global window object
window.conversationHistory = [
    {
        role: "system",
        content: (
            typeof window.SYSTEM_PROMPT !== "undefined"
            ? window.SYSTEM_PROMPT
            : ""
        )
    }
];

// Call Azure OpenAI Chat Completions
window.callAzureAI = async function(displayPrompt, apiPrompt) {
    if (typeof window.appendMessage === "function") {
        window.appendMessage("user", displayPrompt);
    }

    // Update system prompt if it was assigned after initial script execution
    if (
        window.conversationHistory.length > 0
        && !window.conversationHistory[0].content
        && window.SYSTEM_PROMPT
    ) {
        window.conversationHistory[0].content = window.SYSTEM_PROMPT;
    }

    window.conversationHistory.push({
        role: "user",
        content: apiPrompt
    });

    const config = window.CONFIG || {};

    if (
        !config.endpoint
        || !config.deployment
        || !config.apiVersion
    ) {
        const errorMsg = (
            "Configuration error: window.CONFIG is missing "
            +"or incomplete."
        );

        console.error(errorMsg);

        if (typeof window.appendMessage === "function") {
            window.appendMessage("assistant", errorMsg);
        }

        return;
    }

    const baseEndpoint = config.endpoint.replace(/\/+$/, "");

    const url = (
        `${baseEndpoint}/openai/deployments/` +
        `${config.deployment}/chat/completions` +
        `?api-version=${config.apiVersion}`
    );

    console.log("Azure OpenAI Request URL:", url);

    try {
        if (typeof window.getAccessToken !== "function") {
            throw new Error(
                "Function getAccessToken is not defined on window."
            );
        }

        const accessToken = await window.getAccessToken();

        const requestPayload = {
            messages: window.conversationHistory,
            response_format: {
                type: "json_object"
            },
            reasoning_effort: "low",
            max_completion_tokens: 10000
        };

        const response = await fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${accessToken}`
            },
            body: JSON.stringify(requestPayload)
        });

        if (!response.ok) {
            const errorText = await response.text();

            if (response.status === 401) {
                throw new Error(
                    "401 Unauthorized: The Microsoft Entra token "
                   + "is missing, expired, or has the wrong audience."
                );
            }

            if (response.status === 403) {
                throw new Error(
                    "403 Forbidden: The signed-in user does not have "
                    +"permission to access the Azure OpenAI resource."
                );
            }

            if (response.status === 429) {
                const retryAfter = (
                    response.headers.get("Retry-After") || "a few"
                );

                throw new Error(
                    `Rate limit exceeded. Please retry after ` +
                    `${retryAfter} seconds.`
                );
            }

            throw new Error(
                `API Error (${response.status}): ${errorText}`
            );
        }

        const data = await response.json();

        if (
            !data.choices
            || !Array.isArray(data.choices)
            || !data.choices[0]
            || !data.choices[0].message
        ) {
            throw new Error(
                "Azure OpenAI returned an unexpected response format."
            );
        }

        let replyString = data.choices[0].message.content;

        window.conversationHistory.push({
            role: "assistant",
            content: replyString
        });

        replyString = replyString
            .trim()
            .replace(/^```(?:json)?\s*/i, "")
            .replace(/\s*```$/, "");

        const aiResponse = JSON.parse(replyString);

        if (typeof window.appendMessage === "function") {
            window.appendMessage(
                "assistant",
                aiResponse.chat_message,
                aiResponse.action
            );
        }

        // Execute all requested Word document actions sequentially
        if (Array.isArray(aiResponse.actions)) {
            for (const act of aiResponse.actions) {
                if (
                    act.type === "format_text"
                    && act.target_text
                    && typeof window.formatDocumentSubstring === "function"
                ) {
                    await window.formatDocumentSubstring(
                        act.target_text,
                        act.format
                    );
                } else if (
                    act.type === "replace_all"
                    && act.action_text
                    && typeof window.replaceEntireDocument === "function"
                ) {
                    await window.replaceEntireDocument(
                        act.action_text
                    );
                } else if (
                    act.type === "replace_text"
                    && act.target_text
                    && typeof window.replaceDocumentSubstring === "function"
                ) {
                    await window.replaceDocumentSubstring(
                        act.target_text,
                        act.action_text || ""
                    );
                } else if (
                    act.type === "insert_end"
                    && act.action_text
                    && typeof window.insertTextAtEnd === "function"
                ) {
                    await window.insertTextAtEnd(
                        act.action_text
                    );
                }
            }
        }
    } catch (error) {
        console.error("Azure OpenAI Fetch error:", error);

        if (typeof window.appendMessage === "function") {
            window.appendMessage(
                "assistant",
                "Communication/parsing error: " +
                `${error.message}\n\nEndpoint URL:\n${url}`
            );
        }
    }
};

// Fill all matching Word form tags from the agent JSON response
window.fillFormFromJson = async function(agentResponse) {
    try {
        const formData = (
            typeof agentResponse === "string"
            ? JSON.parse(agentResponse)
            : agentResponse
        );

        if (
            !formData
            || typeof formData !== "object"
            || Array.isArray(formData)
        ) {
            throw new Error(
                "Agent response is not a valid JSON object."
            );
        }

        for (const [key, value] of Object.entries(formData)) {
            const targetTag = `{{${key}}}`;

            if (typeof value !== "string") {
                console.error(
                    `Invalid value for tag "${key}".`
                );
                continue;
            }

            if (
                typeof window.replaceAllFormTags === "function"
            ) {
                await window.replaceAllFormTags(
                    targetTag,
                    value
                );
            } else {
                throw new Error(
                    "Function replaceAllFormTags is not defined on window."
                );
            }
        }
    } catch (error) {
        console.error("Form filling error:", error);
        throw error;
    }
};

// Run the selected Microsoft Foundry Agent
window.runFoundryAgent = async function(
    accessToken,
    config,
    agent,
    userInput
) {
    if (!accessToken) {
        throw new Error("Access token is missing.");
    }

    if (!config || !config.endpoint) {
        throw new Error("Foundry Agent endpoint is missing.");
    }

    if (!config.apiVersion) {
        throw new Error("Foundry Agent API version is missing.");
    }

    if (!agent || !agent.name) {
        throw new Error("Foundry Agent name is missing.");
    }

    if (!userInput || !userInput.trim()) {
        throw new Error("Agent input is empty.");
    }

    const baseEndpoint = config.endpoint.replace(/\/+$/, "");

    const url = (
        `${baseEndpoint}/agents/${encodeURIComponent(agent.name)}` +
        `/endpoint/protocols/openai/responses` +
        `?api-version=${config.apiVersion}`
    );

    const requestPayload = {
        input: [
            {
                role: "user",
                content: userInput.trim()
            }
        ]
    };

    console.log("Foundry Agent URL:", url);
    console.log("Foundry Agent:", agent.name);
    console.log("Foundry Agent Version:", agent.version || "default");

    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${accessToken}`
        },
        body: JSON.stringify(requestPayload)
    });

    if (!response.ok) {
        const errorText = await response.text();

        console.error(
            "Foundry Agent HTTP status:",
            response.status
        );

        console.error(
            "Foundry Agent error:",
            errorText
        );

        if (response.status === 401) {
            throw new Error(
                "401 Unauthorized: The Microsoft Entra token "
                +"is missing, expired, or has the wrong audience."
            );
        }

        if (response.status === 403) {
            throw new Error(
                "403 Forbidden: The signed-in user does not have "
                +"permission to invoke this Foundry Agent."
            );
        }

        if (response.status === 404) {
            throw new Error(
                "404 Not Found: The Foundry Agent endpoint, "
                +"project, or agent name may be incorrect."
            );
        }

        if (response.status === 429) {
            const retryAfter = (
                response.headers.get("Retry-After") || "a few"
            );

            throw new Error(
                `Rate limit exceeded. Please retry after ` +
                `${retryAfter} seconds.`
            );
        }

        throw new Error(
            `Agent API Error (${response.status}): ${errorText}`
        );
    }

    const data = await response.json();

    console.log("Raw Foundry Agent response:", data);

    let outputText = "";

    if (typeof data.output_text === "string") {
        outputText = data.output_text;
    } else if (Array.isArray(data.output)) {
        for (const outputItem of data.output) {
            if (!Array.isArray(outputItem.content)) {
                continue;
            }

            for (const contentItem of outputItem.content) {
                if (
                    contentItem.type === "output_text"
                    && typeof contentItem.text === "string"
                ) {
                    outputText += contentItem.text;
                }
            }
        }
    }

    outputText = outputText.trim();

    if (!outputText) {
        throw new Error(
            "The agent returned an empty response."
        );
    }

    outputText = outputText
        .replace(/^```(?:json)?\s*/i, "")
        .replace(/\s*```$/, "")
        .trim();

    try {
        JSON.parse(outputText);
    } catch (error) {
        console.error(
            "Agent returned invalid JSON:",
            outputText
        );

        throw new Error(
            "The agent response is not valid JSON."
        );
    }

    return outputText;
};

// Send user input to the selected Foundry Agent
// and fill the active Word document from the returned JSON.
window.processAgentRequest = async function() {
    const sourceText = document.getElementById("source-text");
    const runAgentBtn = document.getElementById("run-agent-btn");
    const agentStatus = document.getElementById("agent-status");
    const agentSelector = document.getElementById("agent-selector");
    const agentConfig = window.AGENT_CONFIG || {};

    if (!sourceText || !runAgentBtn || !agentStatus) {
        console.error("Agent UI elements are missing.");
        return;
    }

    const userInput = sourceText.value.trim();

    if (!userInput) {
        agentStatus.textContent =
            "יש להזין פרטים לפני הפעלת הסוכן.";
        return;
    }

    if (!agentConfig.endpoint) {
        console.error("AGENT_CONFIG.endpoint is missing.");
        agentStatus.textContent = "כתובת הסוכן חסרה.";
        return;
    }

    if (!agentConfig.apiVersion) {
        console.error("AGENT_CONFIG.apiVersion is missing.");
        agentStatus.textContent = "גרסת ה־API של הסוכן חסרה.";
        return;
    }

    if (!agentConfig.agents) {
        console.error("AGENT_CONFIG.agents is missing.");
        agentStatus.textContent = "הגדרת הסוכנים חסרה.";
        return;
    }

    if (!agentSelector) {
        console.error("Agent selector element is missing.");
        agentStatus.textContent = "בחירת הסוכן אינה זמינה.";
        return;
    }

    const agentId = agentSelector.value;

    if (!agentId) {
        agentStatus.textContent = "יש לבחור תבנית מסמך.";
        return;
    }

    const agent = agentConfig.agents[agentId];

    if (!agent) {
        console.error(
            `Agent configuration not found: ${agentId}`
        );

        agentStatus.textContent =
            "הגדרת הסוכן שנבחר אינה קיימת.";

        return;
    }

    runAgentBtn.disabled = true;
    agentStatus.textContent = "מפעיל סוכן...";

    try {
        if (typeof window.getAccessToken !== "function") {
            throw new Error(
                "Function getAccessToken is not defined on window."
            );
        }

        if (typeof window.fillFormFromJson !== "function") {
            throw new Error(
                "Function fillFormFromJson is not defined on window."
            );
        }

        if (typeof window.runFoundryAgent !== "function") {
            throw new Error(
                "Function runFoundryAgent is not defined on window."
            );
        }

        const accessToken = await window.getAccessToken();

        const agentOutput = await window.runFoundryAgent(
            accessToken,
            agentConfig,
            agent,
            userInput
        );

        console.log(
            "Validated Foundry Agent JSON:",
            agentOutput
        );

        await window.fillFormFromJson(agentOutput);

        agentStatus.textContent = "המסמך מולא בהצלחה.";
    } catch (error) {
        console.error(
            "Agent processing error:",
            error
        );

        agentStatus.textContent =
            `שגיאה: ${error.message}`;
    } finally {
        runAgentBtn.disabled = false;
    }
};