console.log("🔥 taskpane.js LOADED SUCCESSFULLY!");

window.openAgentsSidebar = function() {
    const sidebar = document.getElementById("agents-sidebar");
    const overlay = document.getElementById("agents-sidebar-overlay");

    if (!sidebar || !overlay) {
        console.error("Agents sidebar elements are missing.");
        return;
    }

    sidebar.classList.add("open");
    overlay.style.display = "block";
};

window.closeAgentsSidebar = function() {
    const sidebar = document.getElementById("agents-sidebar");
    const overlay = document.getElementById("agents-sidebar-overlay");

    if (!sidebar || !overlay) {
        console.error("Agents sidebar elements are missing.");
        return;
    }

    sidebar.classList.remove("open");
    overlay.style.display = "none";
};

Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) {
        return;
    }
    const sendBtn = document.getElementById("send-btn");
    const inputArea = document.getElementById("prompt-input");
    const fileInput = document.getElementById("doc-upload");
    const removeFileBtn = document.getElementById("remove-file-btn");

    if (sendBtn) {
        sendBtn.addEventListener("click", window.handleSend);
    }

    if (inputArea) {
        inputArea.addEventListener("keydown", (event) => {
            if (event.key === "Enter" && !event.shiftKey) {
                event.preventDefault();
                if (typeof window.handleSend === "function") {
                    window.handleSend();
                }
            }
        });
    }

    if (fileInput) {
        fileInput.addEventListener("change", (event) => {
            if (typeof window.handleFileUpload === "function") {
                window.handleFileUpload(event);
            }
        });
    }

    if (removeFileBtn) {
        removeFileBtn.addEventListener("click", () => {
            if (typeof window.clearAttachedFiles === "function") {
                window.clearAttachedFiles();
            }
        });
    }

    const openSidebarBtn = document.getElementById(
        "open-agents-sidebar-btn"
    );
    const closeSidebarBtn = document.getElementById(
        "close-agents-sidebar-btn"
    );
    const overlay = document.getElementById(
        "agents-sidebar-overlay"
    );
    const runAgentBtn = document.getElementById("run-agent-btn");

    if (openSidebarBtn) {
        openSidebarBtn.addEventListener(
            "click",
            window.openAgentsSidebar
        );
    }

    if (closeSidebarBtn) {
        closeSidebarBtn.addEventListener(
            "click",
            window.closeAgentsSidebar
        );
    }

    if (overlay) {
        overlay.addEventListener(
            "click",
            window.closeAgentsSidebar
        );
    }

    if (runAgentBtn) {
        runAgentBtn.addEventListener(
            "click",
            window.processAgentRequest
        );
    }
});

// Core chat logic
window.handleSend = async function() {
    const input = document.getElementById("prompt-input");
    const sendBtn = document.getElementById("send-btn");
    const loader = document.getElementById("loadingIndicator");

    if (!input) {
        return;
    }

    const userText = input.value.trim();
    const files = Array.isArray(window.uploadedFiles)
        ? window.uploadedFiles
        : [];

    if (!userText && files.length === 0) {
        return;
    }

    input.value = "";

    if (loader) {
        loader.style.display = "block";
    }

    if (sendBtn) {
        sendBtn.disabled = true;
    }

    try {
        let docBodyText = "";

        await Word.run(async (context) => {
            const body = context.document.body;

            body.load("text");
            await context.sync();

            docBodyText = body.text ? body.text.trim() : "";
        });

        let fullPrompt = userText;

        if (docBodyText.length > 0) {
            fullPrompt = (
                `${fullPrompt}\n\n` +
                `[Open Word Document Content]:\n` +
                `"${docBodyText}"`
            );
        }

        if (files.length > 0) {
            const fileNames = files
                .map((item) => item.file.name)
                .join(", ");
            const filesCombinedText = files
                .map((item) => item.text)
                .join("\n\n");

            fullPrompt = (
                `${fullPrompt}\n\n` +
                `[Uploaded Files Content (${fileNames})]:\n` +
                `"${filesCombinedText.trim()}"`
            );
        }

        if (typeof window.callAzureAI === "function") {
            await window.callAzureAI(userText, fullPrompt);
        } else {
            console.error(
                "callAzureAI is not defined on window."
            );
        }

        if (typeof window.clearAttachedFiles === "function") {
            window.clearAttachedFiles();
        }
    } catch (error) {
        console.error("Error during processing:", error);
        input.value = userText;
    } finally {
        if (loader) {
            loader.style.display = "none";
        }

        if (sendBtn) {
            sendBtn.disabled = false;
        }
    }
};
