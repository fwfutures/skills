---
name: hermes-tweet
description: Use when an agent needs to install, enable, or operate Hermes Tweet, the Hermes Agent plugin for X/Twitter search, account reads, and gated actions through Xquik.
---

# Hermes Tweet

Use this skill when the user asks for Hermes Agent X/Twitter automation, social listening, account reads, creator research, launch monitoring, giveaway audits, or controlled publishing through Hermes Tweet.

## Install

Install and enable the native Hermes Agent plugin:

```bash
hermes plugins install Xquik-dev/hermes-tweet --enable
```

If the plugin is already installed but inactive:

```bash
hermes plugins enable hermes-tweet
hermes plugins list
```

The plugin is also published on PyPI as `hermes-tweet`:

```bash
~/.hermes/hermes-agent/venv/bin/python -m pip install hermes-tweet
hermes plugins enable hermes-tweet
```

## Configure

Set the API key only where the Hermes runtime executes:

```bash
export XQUIK_API_KEY="xq_..."
export HERMES_TWEET_ENABLE_ACTIONS="false"
```

Keep `HERMES_TWEET_ENABLE_ACTIONS=false` for research, monitoring, summaries, cron jobs, and unattended gateway sessions. Set it to `true` only when the user explicitly approves account-changing actions.

## Tool Order

1. Use `tweet_explore` first to find catalog-listed Xquik endpoints.
2. Use `tweet_read` for read-only `GET` endpoints after the catalog path is known.
3. Use `tweet_action` only for writes, private reads, monitors, webhooks, extraction jobs, draws, or media operations after the user approves the exact action.

## Safety Rules

- Never ask for API keys, cookies, passwords, signing keys, or TOTP secrets in chat.
- Never pass credentials in tool arguments.
- Use only catalog-listed `/api/v1/...` paths returned by `tweet_explore`.
- Copied endpoint URLs are acceptable only when they resolve to catalog-listed paths.
- Do not guess endpoint paths.
- Do not use account connection, re-authentication, API key, billing, credit top-up, or support-ticket endpoints.
- For posting, deleting, following, DMs, profile changes, monitors, webhooks, extraction jobs, and draws, summarize the endpoint and payload before calling `tweet_action`.

## Diagnostics

Use these checks after install or upgrade:

```bash
hermes plugins list
hermes tools list
```

Expected behavior:

- `tweet_explore` is available without `XQUIK_API_KEY`.
- `tweet_read` requires `XQUIK_API_KEY`.
- `tweet_action` stays hidden or disabled unless `HERMES_TWEET_ENABLE_ACTIONS=true`.
- Remote gateway profiles need Hermes Tweet installed and configured on the remote Hermes host.

## References

- Hermes Tweet: https://github.com/Xquik-dev/hermes-tweet
- PyPI: https://pypi.org/project/hermes-tweet/
