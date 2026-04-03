# Implementation Plan for GAS Project (Spaziopoesia)

## Overview
This plan addresses the **critical bugs**, **security hardening**, **performance improvements**, and **test coverage** across the Google Apps Script (GAS) codebase. The work is organized by file, prioritized by severity, and includes suggested implementation steps, testing strategies, and rollout considerations.

---

## 1. `gas_config.js`
### Critical Bugs
- **B1 – Hard‑coded alias**: Replace static `KNOWN_ALIASES` with a lazy‑loaded value from Script Properties.
- **B2 – Hard‑coded admin email**: Move `ADMIN_EMAIL` to Script Properties.
### Improvements
- **M1 – Validate `GEMINI_MODELS`**: Add a check that the object exists and is non‑empty.
### Implementation Steps
1. Add helper `getKnownAliases()` returning the property value or default.
2. Add helper `getAdminEmail()` reading the property.
3. Update any references to the old constants.
4. Insert validation in `validateConfig()`.
5. Document the new properties in the project README.
### Testing
- Unit test that missing property falls back to default.
- Test that an empty `GEMINI_MODELS` triggers an error.
---

## 2. `gas_logger.js`
### Critical Bugs
- **B1 – LogLevel validation**: Guard against undefined levels.
- **B2 – Rate‑limit race condition**: Use `LockService` around notification check/set.
### Implementation Steps
1. Refactor `_log` to compute `numericLevel` and early‑return on undefined or below min.
2. Wrap the notification block with `LockService.getScriptLock().tryLock(3000)`; ensure lock release.
3. Add fallback if lock acquisition fails (skip notification).
### Testing
- Simulate concurrent calls (mock lock) and verify only one email is sent.
- Verify logs respect the configured `minLevel`.
---

## 3. `gas_prompt_context.js`
### Critical Bugs
- **B1 – `_safeStringify` fallback**: Return generic `[object Object]` contrary to comment.
- **B2 – Hallucination risk length check**: Uses `.length` on possibly non‑string objects.
### Implementation Steps
1. Replace fallback with safe `String(value)` and a sentinel for non‑serializable.
2. Clarify length check: if `knowledgeBaseMeta?.length` is defined, use it; otherwise compute character count of `knowledgeBase` (e.g., `JSON.stringify(knowledgeBase).length`).
### Testing
- Unit tests for objects that throw on `JSON.stringify`.
- Verify length‑based guard works for both meta and array.
---

## 4. `gas_rate_limiter.js`
### Critical Bugs
- **B1 – Mutating shared `strategies`**: Clone before modification.
- **B2 – Token count includes empty strings**.
- **B3 – Over‑engineered midnight calculation**.
### Improvements
- **M1 – Shadow token counter for `skipRateLimit` paths**.
### Implementation Steps
1. Replace `const taskStrategies = this.strategies;` with a shallow copy via `Object.assign({}, this.strategies)`.
2. Update `_estimateTokens` to `text.trim().split(/\s+/).filter(Boolean).length`.
3. Simplify `_getNextResetTime` using direct arithmetic as described.
4. Add a secondary counter (e.g., `this.shadowTokenUsage`) that increments even when `skipRateLimit` is true.
### Testing
- Verify strategies are not mutated across instances.
- Test token estimation with leading/trailing spaces.
- Validate reset time matches Pacific midnight.
---

## 5. `gas_memory_service.js`
### Critical Bugs
- **B1 – Topic matching with underscores**: Use regex for word stems.
- **B2 – GC not triggered on reads**.
- **B3 – `cleanOldEntries` deletes rows with empty `threadId`**.
- **B5 – `_normalizeTopicKey` returns empty string**.
### Implementation Steps
1. Replace simple `includes` with a regex that matches word boundaries and common stems.
2. Increment `_opCount` in `_getFromCache` and trigger GC when threshold reached.
3. Add guard `if (!data[i][0]) continue;` before deletion.
4. Filter out empty topics before using them in memory.
### Testing
- Simulate cache reads/writes to ensure GC runs.
- Verify that rows without threadId are preserved.
---

## 6. `gas_classifier.js`
### Critical Bugs
- **B1 – Over‑aggressive out‑of‑office detection for "malattia"**.
- **B2 – Unicode property escape compatibility**.
- **B3 – Missing "gentile" greeting detection**.
### Implementation Steps
1. Refine regex to require context words (e.g., `fuori` or `in malattia fino`).
2. Replace `\p{L}\p{N}` with a safe ASCII‑plus‑accented character class.
3. Extend greeting list to include `gentile`.
### Testing
- Unit tests for various email bodies containing the word "malattia" in legitimate contexts.
---

## 7. `gas_error_types.js`
### Critical Bugs
- **B1 – Over‑broad `500` detection**.
- **B2 – Missing handling for `ECONNREFUSED` / `ENOTFOUND`**.
### Implementation Steps
1. Replace simple `includes` with regex `/\b(status\s+)?500\b|internal\s+server\s+error/i`.
2. Extend error type list to include the two DNS/connection errors.
### Testing
- Mock error messages and verify correct classification.
---

## 8. `gas_main.js`
### Critical Bugs
- **B1 – Unused `_parseStrictHour`**.
- **B2 – Formula error cells included in KB text**.
- **B3 – Insufficient locking around resource loading**.
### Implementation Steps
1. Remove dead code or integrate into hour‑parsing logic if needed.
2. Filter out cells whose string starts with `#` before formatting.
3. Add internal lock inside `_loadResourcesInternal` (e.g., `LockService.getScriptLock()`).
### Testing
- Verify that KB text no longer contains `#REF!` etc.
- Simulate concurrent triggers to ensure only one loads resources.
---

## 9. `gas_request_classifier.js`
### Critical Bugs
- **B1 – Unused `externalHint` parameter**.
- **B2 – Hard‑coded email in hint string**.
- **B3 – Inconsistent parameter order**.
### Implementation Steps
1. Remove `externalHint` or actually use it in classification.
2. Move email address to `CONFIG` or Script Properties and reference it.
3. Align method signatures: either keep both but document, or unify to `(subject, body)`.
### Testing
- Ensure `classify` works with/without hint.
---

## 10. `gas_response_validator.js`
### Critical Bugs
- **B1 – Dead `_rimuoviThinkingLeak`**.
- **B2 – Dependency on `Utilities.formatDate` in tests**.
- **B3 – Auto‑healing hides recurring issues**.
- **B4 – Over‑broad forbidden phrases**.
### Implementation Steps
1. Invoke `_rimuoviThinkingLeak` inside `_perfezionamentoAutomatico` when relevant.
2. Add fallback/mock for `Utilities.formatDate` in unit tests to return hour when needed.
3. Log every auto‑correction with details (model, tokens, etc.).
4. Refine regex for `forse`/`probabilmente` with word‑boundary checks.
### Testing
- Add tests for thinking‑leak removal.
- Verify that forbidden phrase detection respects boundaries.
---

## 11. `gas_prompt_engine.js`
### Critical Bugs
- **B1 – Duplicate `_renderKnowledgeBase`** (lost anti‑hallucination rule).
- **B2 – ReDoS vulnerable address regex**.
- **B3 – Misuse of `workingAttachmentsContext`**.
### Implementation Steps
1. Consolidate the two definitions into a single robust implementation that includes the “NON‑INVENTARE” rule.
2. Replace the complex regex with a safer pattern using a limited character class and length cap.
3. Refine attachment‑context check to verify actual poetry criteria (e.g., `category === 'poem_submission' && attachmentBlobs.length > 0`).
### Testing
- Unit test that the knowledge‑base rendering includes the rule.
- Benchmark the new address regex against long strings.
---

## 12. `gas_unit_tests.js`
### Critical Bugs
- **B1 – Inaccurate `Utilities.formatDate` mock**.
- **B2 – Ineffective lock‑management test**.
- **B3 – Missing tests for `computeSalutationMode` & `computeResponseDelay`**.
### Implementation Steps
1. Replace mock with a more faithful implementation handling hour (`HH`) and date formats.
2. Enhance lock test to assert lock state after operation.
3. Add dedicated test groups for the two pure functions.
### Testing
- Run full suite; ensure coverage > 90% for critical modules.
---

## 13. `gas_gemini_service.js`
### Critical Bugs
- **B1 – Backup‑key retry logic mis‑ordered**.
- **B2 – `_resolveLanguage` safety‑grade undefined**.
### Implementation Steps
1. Refactor `_quickCheckWithModel` to separate primary and backup attempts as described.
2. Provide default `localSafetyGrade = 0` in `_resolveLanguage`.
### Testing
- Simulate primary failure and verify backup is used exactly once.
---

## 14. `gas_email_processor.js`
### Critical Bugs
- **B1 – Fallback‑Lite may be blocked by rate limiter**.
- **B2 – ReDoS in `_detectProvidedTopics` (already addressed in prompt engine)**.
- **B3 – `processedCount` does not include skipped threads**.
- **B5 – `_normalizeTopicKey` can produce empty keys**.
### Implementation Steps
1. Set `skipRateLimit: true` for the fallback‑lite strategy and use backup key if available.
2. Ensure `processedCount` increments for all processed threads, optionally adding a `skippedCount` metric.
3. Filter out empty topics before using them.
### Testing
- Verify that fallback‑lite runs even when primary quota is exhausted.
- Check that counters reflect total iterations.
---

## 15. `gas_gmail_service.js`
### Critical Bugs
- **B1 – `substring` on mutated `result` may mis‑align**.
- **B2 – `sanitizeUrl` lacks handling for protocol‑relative URLs**.
- **B3 – Emoji conversion breaks ZWJ sequences**.
### Implementation Steps
1. Operate on the original text or recompute positions after each mutation.
2. Extend `sanitizeUrl` to accept URLs starting with `//` and prepend `https:`.
3. Adjust emoji conversion to skip code point `0x200D` (ZWJ).
### Testing
- Unit tests for URL sanitization with protocol‑relative URLs.
- Verify emoji sequences remain intact after conversion.
---

## 16. `gas_setup_ui.js`
### Critical Bugs
- **B1 – Duplicate range protections**.
- **B2 – Aggressive `;` → `,` replacement in formulas**.
### Implementation Steps
1. Before adding a protection, remove any existing one for the same range.
2. Refine locale fallback to replace only delimiter semicolons outside string literals (use a simple parser or regex that respects quotes).
### Testing
- Ensure a second call to `setupControlloSheet` does not increase protection count.
- Verify formulas with embedded semicolons remain correct.
---

## 17. General Tasks
1. **Documentation**: Update README with new configuration properties and security considerations.
2. **CI Integration**: Add a lint step (ESLint) and ensure all new code passes.
3. **Deployment**: After changes, run `clasp push` and verify the script works in a test environment before production rollout.
4. **Monitoring**: Add logging for rate‑limit events and backup‑key usage to aid future debugging.

---

**Next Steps**
- Prioritize critical bugs (files 1, 2, 4, 8, 10, 11, 13, 14).
- Implement fixes in small batches, run unit tests after each batch.
- Once all critical bugs are resolved, proceed with improvements and test enhancements.

*Please review this plan and let me know if any priorities should be adjusted or if additional constraints exist.*
