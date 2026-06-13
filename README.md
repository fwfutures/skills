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

## Skills

| Skill | Description |
|-------|-------------|
| hello-world | Tells the user a joke to brighten their day |
| fresh-auth | Unified secure access to Office, email, calendar, Microsoft Graph, and Notion via auth.freshhub.ai |
| hermes-tweet | Installs and operates Hermes Tweet for X/Twitter search, reads, and gated actions through Hermes Agent |

## License

MIT
