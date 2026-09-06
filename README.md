# Freshwater Futures Skills

Public agent skills for Claude Code and Codex.

## Install

### Local development

Use the repo installer so skills are symlinked from your checkout and stay editable in place:

```bash
./scripts/install-skills.sh
```

### Published package

```bash
npx skills add fwfutures/skills -g
```

## TokenSaver

[TokenSaver](plugins/tokensaver/README.md) is a Claude Code and Codex plugin that blocks large stale conversation turns and saves a bounded local handoff. Install the plugin (skill-only installation does not enable its hooks).

Claude Code:

```text
/plugin marketplace add fwfutures/skills
/plugin install tokensaver@fwf-public-marketplace
```

Codex: add `fwfutures/skills` as a plugin marketplace, install `tokensaver`, and review/trust its hooks. See the plugin README for configuration and test coverage.

## Skills

| Skill | Description |
|-------|-------------|
| hello-world | Tells the user a joke to brighten their day |
| fresh-auth | Unified secure access to Office, email, calendar, Microsoft Graph, and Notion via auth.freshhub.ai |

## License

MIT
