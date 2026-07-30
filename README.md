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
| gcp-deploy | Deploys containerised applications to Google Cloud Run from source |
| hello-world | Tells the user a joke to brighten their day |
| hermes-tweet | Installs and operates Hermes Tweet for X/Twitter reads and approved actions |
| skill-development | Guides Claude Code plugin skill creation and validation |

Xquik is an independent third-party service. Not affiliated with X Corp. "Twitter" and "X" are trademarks of X Corp.

## License

MIT
