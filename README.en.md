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
7.  **Sending & Labeling**: The response is sent and the message is labeled (e.g., `IA`, `Verifica`, or `Errore`) without being automatically archived.

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

The generation chain follows the GMAIL PARROCCHIA/AG project:
`gemini-3.8-flash` (primary, backup) → `gemini-3.6-flash` (primary) →
`gemini-flash-latest` (primary, backup) → `gemini-3.5-flash-lite` (primary, backup).
Backup attempts are skipped when no backup key is configured.
Quick checks, classification and semantic validation prioritize `gemini-3.5-flash-lite`,
then `gemini-flash-lite-latest`, the primary Flash model and `gemini-flash-latest`, according to quota availability.
Internal keys `flash-3.7` and `flash-3.7-backup` remain stable to preserve quota counters;
the `name` field identifies the actual API model. Existing operational limits are retained and should be checked against the project's AI Studio quotas.
Context caches are automatically recreated when the model changes.
See the [official Gemini 3.8 Flash migration guide](https://ai.google.dev/gemini-api/docs/generate-content/latest-model).

1.  **Script Properties**: Set `GEMINI_API_KEY`, `SPREADSHEET_ID`, and `METRICS_SHEET_ID`.
2.  **Google Sheets**: Create a sheet with tabs `Istruzioni`, `Sostituzioni`, `Controllo`, `ConversationMemory`, and `DailyMetrics`.
3.  **Triggers**: Run the `setupAllTriggers()` function from the `gas_main.js` file to activate the system.

---

<p align="center">
  Created with ❤️ for automated poetry management.
</p>
