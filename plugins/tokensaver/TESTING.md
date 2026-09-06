# TokenSaver verification — 6 September 2026

- Eight Python unit tests passed for both transcript formats, stale/fresh/threshold decisions, malformed records, missing timestamps, platform SessionStart behaviour, and private bounded handoffs. Default-boundary tests cover Claude at 60 minutes and Codex at 30 minutes (allowed at the boundary, blocked one second later), a warm 20-minute Codex context, the size threshold, and a TTL override.
- Claude plugin manifest and Codex plugin validator passed.
- Claude Code live UserPromptSubmit test: release guard received a synthetic 60,000-token stale transcript via a test-only adapter. Host returned blocked, zero turns, zero usage and $0 cost.
- Codex CLI 0.146.0 live UserPromptSubmit test: release guard received the same synthetic stale condition. Host returned turn.completed with zero input, cached input, cache-write input, output and reasoning tokens. A deliberately unavailable localhost model provider avoided paid generation; model-catalog lookup errors occurred before the hook, so this test does not claim zero background network activity. Hook trust was bypassed only for this reviewed test invocation.
- Earlier Claude Desktop local Code prototype test: exact-trigger and size/age-condition prompts were visibly blocked; ordinary control proceeded. Transcript contained no assistant responses for blocked prompts. Release packaging was separately CLI-tested; it has not been installed globally into Desktop.

Codex resume SessionStart handling is unit-tested; live Codex coverage is UserPromptSubmit. Actual provider cache expiry and automatic new-session creation are not covered. Lifetime defaults remain heuristics.

No private transcripts, handoff contents, account information, or local machine paths are included in these published test records.
