# TokenSaver

A local hook and companion skill that stop an old, large conversation before another expensive input-context request, where the host supports blocking hooks. Requires Python 3. No dependencies, network calls, or model calls are used by the checker.

Built by [Freshwater Futures](https://freshwaterfutures.com). We combine behavioural science and AI engineering to help organisations turn AI inertia into AI impact.

Claude Code uses `UserPromptSubmit`; its `SessionStart` event cannot block and produces no output. Codex uses `SessionStart` on resume plus `UserPromptSubmit`. Platform detection uses Codex's `PLUGIN_ROOT`; tests and custom launchers can pass `codex` or `claude` explicitly to `scripts/guard.py`.

The checker reads at most the last 8 MiB of the transcript. Claude context size is the latest assistant usage's input plus cache-read plus cache-creation input tokens. Codex uses `event_msg.token_count.info.last_token_usage.input_tokens`, not cumulative session tokens. It compares the associated timestamp with a configured assumed cache lifetime.

| Environment variable | Default |
| --- | --- |
| `TOKENSAVER_THRESHOLD_TOKENS` | `50000` |
| `TOKENSAVER_CLAUDE_TTL_SECONDS` | `3600` |
| `TOKENSAVER_CODEX_TTL_SECONDS` | `1800` |

## Intended use and timing assumptions

TokenSaver is intended for **subscription usage in the main conversation**: a reminder to start fresh instead of accidentally submitting a prompt to a very large thread after a long break. Defaults assume Claude Code is within its subscription's included usage and Codex uses GPT-5.6 or later. They are not tuned for API billing, usage-credit overages, third-party providers, subagents, or background requests. These are usage assumptions, not automatically detected restrictions.

- **Claude: 60 minutes.** Claude Code requests a one-hour TTL for the main conversation within included subscription usage. API keys, usage credits and most subagent requests default to five minutes. See [Claude Code cache lifetime](https://code.claude.com/docs/en/prompt-caching#cache-lifetime) and [Anthropic prompt caching](https://platform.claude.com/docs/en/build-with-claude/prompt-caching).
- **Codex: 30 minutes.** OpenAI documents a 30-minute minimum cache lifetime for GPT-5.6 and later, refreshed on write or reuse. We use the end of that minimum window as a practical point to guard against a costly return; entries can remain cached longer. This is not a separate guarantee for every Codex subscription backend. See [OpenAI cache lifetime](https://developers.openai.com/api/docs/guides/prompt-caching#cache-lifetime) and [OpenAI's GPT-5.6 builder's guide](https://openai.com/index/builders-guide-to-gpt-5-6/).

These defaults are configurable estimates, not confirmed expiry times. The checker uses the latest usage-record timestamp as a proxy for cache activity. Anthropic measures TTL from request start, so response generation and transcript recording can make our trigger late. Prefix changes can cause misses earlier, and retention or reuse elsewhere can keep prefixes warm longer. The checker does not query live cache state or detect authentication, billing mode, model, or effective provider TTL.

The trigger requires size at or above 50,000 input tokens and age strictly greater than the configured lifetime. Override the environment variables for other usage patterns (for example `TOKENSAVER_CLAUDE_TTL_SECONDS=300` for a five-minute policy).

On a block, TokenSaver creates a private temporary directory (0700) containing a per-session handoff file (0600), outside the repository. It includes at most six recent user/assistant text entries (4,000 characters each) and 4,000 characters of the pending prompt. Start a fresh session and ask it to read that file. Original sessions are preserved. Automatic session creation is not implemented. Temporary handoffs may contain sensitive conversation text; delete them when no longer needed.

Missing, malformed, or unavailable telemetry allows work to continue. Tail parsing can miss older usage records. Hook errors/timeouts or unsupported host versions can also allow work to continue; this is a cost-saving aid, not a billing guarantee. Hooks cannot prevent local transcript loading or unrelated host background requests. A fresh session may still have large system/project instructions.

Run tests:

```sh
python3 -m unittest discover -s plugins/tokensaver/tests -v
```

Tests use synthetic transcripts; no private session transcripts are shipped.

## Install

Claude Code (including local Code sessions in Desktop): add marketplace `fwfutures/skills`, then install `tokensaver@fwf-public-marketplace` from the plugin menu. Restart the session to load hooks.

Codex:

```sh
codex plugin marketplace add fwfutures/skills
codex plugin add tokensaver@fwf-public-marketplace
```

Start a new session and review/trust the hook definitions. Plugin installation alone does not grant hook trust. Skill-only installers do not activate hooks.

See [TESTING.md](TESTING.md) for measured coverage and limitations.
