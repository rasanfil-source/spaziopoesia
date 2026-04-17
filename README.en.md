# 🖋️ Spaziopoesia: Gmail AI Autoresponder

<p align="center">
  <a href="README.md"><b>🇮🇹 Italiano</b></a> | 
  <a href="README.en.md"><b>🇬🇧 English</b></a>
</p>

---

**Spaziopoesia** is an advanced Gmail autoresponder system powered by **Google Apps Script** and **Gemini AI**. It is designed to automatically handle communications related to poetry contests, information requests, and poem submissions, ensuring fast and context-aware responses.

## 🚀 How It Works

The system operates as an automated pipeline that follows these steps:

1.  **Time Trigger**: Every 5 minutes (configurable), the system activates automatically.
2.  **Resource Loading**: It reads instructions (Knowledge Base) and configurations from a dedicated Google Sheet. It uses a sophisticated caching system to optimize performance.
3.  **Email Analysis**: It identifies unread emails, ignoring spam and blacklisted contacts.
4.  **Memory Retrieval**: Using the `MemoryService`, the bot "remembers" previous interactions in the same thread to maintain a coherent conversation.
5.  **Prompt Engineering**: The `PromptEngine` builds a detailed request for Gemini, including thread context, the Knowledge Base, and any attachments (including OCR).
6.  **Generation & Validation**: Gemini generates the response, which is then validated by the `ResponseValidator` to ensure quality and consistency.
7.  **Sending & Labeling**: The response is sent, and the email is archived with a label (e.g., `IA`, `Verifica`, or `Errore`).

## ✨ Key Features

-   🤖 **Multimodal AI**: Support for analyzing attachments (images, PDFs, Office documents) via Gemini Vision.
-   📚 **Dynamic Knowledge Base**: Easily updateable instructions through a Google Sheet ("Istruzioni").
-   🧠 **Contextual Memory**: Thread management to avoid repetitive responses and maintain the flow of conversation.
-   🛡️ **Semantic Validation**: Scoring system to filter out potentially incorrect or out-of-context responses.
-   🚦 **Rate Limiting**: Intelligent management of API quotas to avoid blocks.
-   📊 **Metrics**: Daily export of usage statistics to Google Sheets.

## 🛠️ Project Structure

The code is divided into specialized modules:

-   `gas_main.js`: Main orchestrator and trigger management.
-   `gas_email_processor.js`: Email batch processing logic.
-   `gas_gemini_service.js`: Interface with Gemini APIs.
-   `gas_prompt_engine.js`: AI prompt construction.
-   `gas_memory_service.js`: Sheet-based conversation history management.
-   `gas_config.js`: Centralized configuration (API keys, models, limits).
-   `gas_setup_ui.js`: Interface for initial trigger configuration.

## ⚙️ Quick Setup

1.  **Script Properties**: Set `GEMINI_API_KEY`, `SPREADSHEET_ID`, and `METRICS_SHEET_ID`.
2.  **Google Sheets**: Create a sheet with tabs `Istruzioni`, `Sostituzioni`, `Controllo`, `ConversationMemory`, and `DailyMetrics`.
3.  **Triggers**: Run the `setupAllTriggers()` function from the `gas_main.js` file to activate the system.

---

<p align="center">
  Created with ❤️ for automated poetry management.
</p>
