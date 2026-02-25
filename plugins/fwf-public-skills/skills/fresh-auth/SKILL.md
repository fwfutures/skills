---
name: fresh-auth
description: This skill should be used when the user asks to access secured work data through "office", "email", "calendar", "microsoft graph", or "notion" using auth.freshhub.ai and the Freshwater secure auth proxy.
---

# Fresh Auth Workspace CLI

Use this skill as the single entry point for secure, identity-aware access to Microsoft 365 (Graph) and Notion.

## Use this skill for

- Drive operations, OneDrive file access, and share links through `office-cli`.
- Outlook inbox, search, send, and reply actions through `office-cli`.
- Calendar lookup and scheduling visibility through `office-cli`.
- People lookups through `office-cli`.
- Notion database search, query, page read/write, and markdown conversion through `notion-query`.
- Any request that mixes Office and Notion data in one workflow.

## Security model

- Use the Auth Service Proxy at `https://auth.freshhub.ai` for Microsoft Graph actions.
- Use `office-cli` with agent-session grants, never raw OAuth tokens.
- Keep `~/.config/office-cli/agent-session` at secure mode (600).
- Store secrets in environment variables.
- Use `NOTION_API_KEY` for Notion API requests.
- Use `OPENROUTER_API_KEY` for Office PDF/image conversion.
- Use `AUTH_SERVICE_URL` only when overriding the default proxy endpoint.

## Runtime prerequisites

- Install Bun for `scripts/office-cli.ts`.
- Install `jq` and `curl` for `scripts/notion-query.sh`.
- Run commands from this skill folder or export explicit paths.

```bash
# Resolve skill location for common installers
export FRESH_AUTH_DIR="${HOME}/.agents/skills/fresh-auth"
[ -d "$FRESH_AUTH_DIR" ] || export FRESH_AUTH_DIR="${HOME}/.codex/skills/fresh-auth"

export OFFICE_CLI="${FRESH_AUTH_DIR}/scripts/office-cli.ts"
export NOTION_CLI="${FRESH_AUTH_DIR}/scripts/notion-query.sh"
export AUTH_SERVICE_URL="https://auth.freshhub.ai"

# Quick command discovery
[ -f "$OFFICE_CLI" ] && bun "$OFFICE_CLI" status
[ -x "$NOTION_CLI" ] && "$NOTION_CLI" me
```

## Bundled scripts

- `scripts/office-cli.ts` for Microsoft Graph-backed Drive, Mail, Calendar, and People actions.
- `scripts/notion-query.sh` for direct Notion API read/write workflows.

## Office + Graph: canonical flow

Follow this flow when granting access for Graph-backed tools.

```bash
# Register and create grants
bun "$OFFICE_CLI" login
bun "$OFFICE_CLI" request drive
bun "$OFFICE_CLI" request mail
bun "$OFFICE_CLI" request cal

# Verify active grants
bun "$OFFICE_CLI" status
```

## Command map: Office CLI

## Drive / Graph storage

```bash
bun "$OFFICE_CLI" drive list
bun "$OFFICE_CLI" drive list "/Documents"
bun "$OFFICE_CLI" drive search "Quarterly report"
bun "$OFFICE_CLI" drive download <file-id> out.docx
bun "$OFFICE_CLI" drive content <file-id>
bun "$OFFICE_CLI" drive convert <file-id> --output=notes.md
bun "$OFFICE_CLI" drive share <file-id> --type edit
bun "$OFFICE_CLI" drive share <file-id> --anyone
bun "$OFFICE_CLI" drive permissions <file-id>
bun "$OFFICE_CLI" drive unshare <file-id> <permission-id>
```

## Mail / Email

```bash
bun "$OFFICE_CLI" mail inbox
bun "$OFFICE_CLI" mail inbox --count 50
bun "$OFFICE_CLI" mail unread
bun "$OFFICE_CLI" mail search "team update"
bun "$OFFICE_CLI" mail read <message-id>
bun "$OFFICE_CLI" mail send --to "teammate@example.com" --subject "Brief" --body "Thanks for the update"
bun "$OFFICE_CLI" mail send --to "brad" --subject "Quick check" --body "Approved" --yes
bun "$OFFICE_CLI" mail reply <message-id> --body "Got it."
bun "$OFFICE_CLI" mail reply-all <message-id> --body "Thanks everyone."
```

## Calendar

```bash
bun "$OFFICE_CLI" cal today
bun "$OFFICE_CLI" cal tomorrow
bun "$OFFICE_CLI" cal events --days 14
bun "$OFFICE_CLI" cal events --full
```

## People lookup (Graph contact helper)

```bash
bun "$OFFICE_CLI" people "brad"
bun "$OFFICE_CLI" people "brad" --verbose
```

## Notion command map

```bash
$NOTION_CLI me
$NOTION_CLI find-db "my database"
$NOTION_CLI search "my database"
$NOTION_CLI get-db <database-id>
$NOTION_CLI query-db <database-id>
$NOTION_CLI get-page <page-id>
$NOTION_CLI get-markdown <page-id>
$NOTION_CLI create <database-id> "Title" -p "Status=In progress" -p "Priority=High"
$NOTION_CLI update <page-id> -p "Status=Done"
$NOTION_CLI set-body <page-id> -
$NOTION_CLI append-body <page-id> -
```

Use `find-db` first when the database ID is unknown. It returns database `id`, `title`, and `url` so the ID can be copied directly into `get-db`, `query-db`, or `create`.

## Notion backlog helper

```bash
# Optional: enable shortcuts for a specific Notion backlog database
export NOTION_BACKLOG_DB_ID="<database-id>"

$NOTION_CLI backlog
$NOTION_CLI backlog "In Progress"
$NOTION_CLI create-backlog "New task"
```

## Multi-tool patterns

- Run `people` first, then `mail send --to <resolved email>` for safer identity resolution.
- Pull a Notion task with `search` or `query-db`, then append context with `append-body`.
- Convert a meeting PDF in Drive to markdown with `drive convert`, then store notes in Notion via `append-body`.

## Error handling

- `no_agent_session`: run `bun "$OFFICE_CLI" login`.
- `no_grant`: run `bun "$OFFICE_CLI" request <drive|mail|cal>`.
- `token expired`: run `bun "$OFFICE_CLI" status` and follow the returned re-authorisation URL.
- `NOTION_API_KEY` missing: set variable and rerun the failing Notion command.
- `NOTION_BACKLOG_DB_ID` missing: set variable or call generic `query-db`/`create` commands instead of backlog shortcuts.
- Microsoft account not linked: follow the URL output by Graph proxy responses.

## Public publication checks

- Keep `NOTION_API_KEY` and graph proxy URLs configurable by environment.
- Do not embed API keys or session IDs in skill outputs.
- Keep all commands pointed at `https://auth.freshhub.ai` by default.
- Include both CLIs under this skill's `scripts/` folder for self-contained installation.
- Mention both Microsoft 365 and Notion capabilities in onboarding docs because this is a unified user-facing access path.
