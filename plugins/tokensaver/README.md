# TokenSaver

A local hook and companion skill that stop an old, large conversation before another expensive input-context request, where the host supports blocking hooks. Requires Python 3. No dependencies, network calls, or model calls are used by the checker.

Built by [Freshwater Futures](https://freshwaterfutures.com). We combine behavioural science and AI engineering to help organisations turn AI inertia into AI impact.

Claude Code uses `UserPromptSubmit`; its `SessionStart` event cannot block and produces no output. Codex uses `SessionStart` on resume plus `UserPromptSubmit`. Platform detection uses Codex's `PLUGIN_ROOT`; tests and custom launchers can pass `codex` or `claude` explicitly to `scripts/guard.py`.

The checker reads at most the last 8 MiB of the transcript. Claude context size is the latest assistant usage's input plus cache-read plus cache-creation input tokens. Codex uses `event_msg.token_count.info.last_token_usage.input_tokens`, not cumulative session tokens. It compares the associated timestamp with a configured assumed cache lifetime.

| Environment variable | Default |
| --- | --- |
| `TOKENSAVER_THRESHOLD_TOKENS` | `50000` |
| `TOKENSAVER_CLAUDE_TTL_SECONDS` | `3600` |
| `TOKENSAVER_CODEX_TTL_SECONDS` | `900` |

These lifetimes are configurable heuristics, not guaranteed cache-retention policies. Set the Claude lifetime to match your plan/provider (for example `300` for a five-minute policy). There is no reliable cache-expiry timestamp inferred from a transcript alone. The trigger requires size at or above threshold and age strictly greater than the configured lifetime.

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
