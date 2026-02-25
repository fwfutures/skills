#!/usr/bin/env node
/**
 * Notion CLI - Access Notion through auth service proxy (OAuth via auth.freshhub.ai)
 */

import { existsSync, readFileSync, writeFileSync, mkdirSync, unlinkSync } from "fs";
import { dirname, resolve, join } from "path";
import { fileURLToPath } from "url";
import { homedir, hostname } from "os";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

function loadDotEnv(envPath) {
  if (!existsSync(envPath)) return;
  const content = readFileSync(envPath, "utf8");
  for (const rawLine of content.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const eq = line.indexOf("=");
    if (eq <= 0) continue;
    const key = line.slice(0, eq).trim();
    let value = line.slice(eq + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    if (!(key in process.env)) process.env[key] = value;
  }
}

loadDotEnv(resolve(__dirname, "..", ".env"));

const RAW_AUTH_SERVICE_URL =
  process.env.AUTH_SERVICE_URL ||
  process.env.AUTH_SERVICE ||
  "https://auth.freshhub.ai";
const AUTH_SERVICE_URL = RAW_AUTH_SERVICE_URL.replace(/\/+$/, "").replace(
  /\/api$/,
  ""
);
const NOTION_PROXY_BASE = `${AUTH_SERVICE_URL}/api/proxy/notion`;
const NOTION_STATUS_URL = `${AUTH_SERVICE_URL}/api/proxy/status/notion`;
const SHARED_AGENT_SESSION_FILE =
  process.env.FRESH_AUTH_AGENT_SESSION_FILE ||
  join(homedir(), ".config", "fresh-auth", "agent-session");
const LEGACY_AGENT_SESSION_FILES = [
  join(homedir(), ".config", "office-cli", "agent-session"),
];
const AUTO_REQUEST_POLL_MS = 2000;
const AUTO_REQUEST_ENABLED = process.env.OFFICE_AUTO_REQUEST !== "0";
const NOTION_SCOPES = ["read", "write"];
const NOTION_GRANT_DURATION = "1h";
const NOTION_VERSION = process.env.NOTION_API_VERSION || "2022-06-28";
const NOTION_BACKLOG_DB_ID = process.env.NOTION_BACKLOG_DB_ID || "";

function printJson(value) {
  console.log(JSON.stringify(value, null, 2));
}

function errorAndExit(message, code = 1) {
  console.error(message);
  process.exit(code);
}

function splitFirst(input, separator) {
  const idx = input.indexOf(separator);
  if (idx < 0) return [input, ""];
  return [input.slice(0, idx), input.slice(idx + separator.length)];
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getAgentSession() {
  const candidates = [SHARED_AGENT_SESSION_FILE, ...LEGACY_AGENT_SESSION_FILES];
  try {
    for (const sessionFile of candidates) {
      if (!existsSync(sessionFile)) continue;
      const sessionRaw = readFileSync(sessionFile, "utf8").trim();
      if (!sessionRaw) continue;

      try {
        const parsed = JSON.parse(sessionRaw);
        const normalized =
          parsed.agentSessionId || parsed.agentSession || parsed.session;
        if (normalized?.trim()) return normalized.trim();
      } catch {
        return sessionRaw;
      }
    }
    return null;
  } catch {
    return null;
  }
}

function saveAgentSession(agentSessionId) {
  mkdirSync(dirname(SHARED_AGENT_SESSION_FILE), { recursive: true });
  writeFileSync(SHARED_AGENT_SESSION_FILE, `${agentSessionId.trim()}\n`, { mode: 0o600 });
}

function clearAgentSession() {
  try {
    for (const sessionFile of [
      SHARED_AGENT_SESSION_FILE,
      ...LEGACY_AGENT_SESSION_FILES,
    ]) {
      if (existsSync(sessionFile)) {
        unlinkSync(sessionFile);
      }
    }
  } catch {
    // Ignore clear failures
  }
}

async function readErrorBody(response) {
  const text = await response.text();
  if (!text) return { data: null, text: "" };
  try {
    return { data: JSON.parse(text), text };
  } catch {
    return { data: null, text };
  }
}

function normalizeAuthError(data) {
  const details =
    data?.details && typeof data.details === "object" ? data.details : {};
  return { ...(data || {}), ...details };
}

function describeAuthError(data) {
  const error = normalizeAuthError(data);

  if (error.error === "no_agent_session") {
    console.error("No agent session. Run 'node notion-query.js login' first.");
    return;
  }

  if (
    ["no_grant", "grant_expired", "single_use_exhausted"].includes(
      error.error || ""
    )
  ) {
    console.error("No active grant for notion.");
    console.error("");
    console.error("Request access by running:");
    console.error("  node notion-query.js request");
    return;
  }

  if (error.error === "no_oauth_token" && error.connectUrl) {
    console.error("Notion account not linked.");
    console.error("");
    console.error("Link account at:");
    console.error(`  ${error.connectUrl}`);
    return;
  }

  if (error.reauthorizeUrl) {
    console.error("Access token expired or missing scopes.");
    if (Array.isArray(error.missingScopes) && error.missingScopes.length > 0) {
      console.error(`Missing scopes: ${error.missingScopes.join(" ")}`);
    }
    console.error("");
    console.error("Re-authorize at:");
    console.error(`  ${error.reauthorizeUrl}`);
    return;
  }

  if (error.elevateUrl) {
    console.error("Additional authorization required.");
    console.error("");
    console.error("Approve at:");
    console.error(`  ${error.elevateUrl}`);
    return;
  }

  if (error?.message || error?.error) {
    console.error(error.message || error.error);
  }
}

function isAutoRequestError(data) {
  const error = normalizeAuthError(data);
  if (!error?.error) return false;
  return ["no_grant", "grant_expired", "single_use_exhausted"].includes(
    error.error
  );
}

async function requestGrant() {
  const agentSession = getAgentSession();
  if (!agentSession) {
    errorAndExit("No agent session. Run 'node notion-query.js login' first.");
  }

  const response = await fetch(`${AUTH_SERVICE_URL}/api/auth-request`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Agent-Session": agentSession,
    },
    body: JSON.stringify({
      service: "notion",
      scopes: NOTION_SCOPES,
      duration: NOTION_GRANT_DURATION,
    }),
  });

  if (!response.ok) {
    const { data, text } = await readErrorBody(response);
    if (data?.error || data?.message) {
      throw new Error(data.message || data.error);
    }
    throw new Error(`Failed to create auth request (${response.status}): ${text || response.statusText}`);
  }

  return response.json();
}

async function waitForGrantApproval(pollUrl, expiresAt) {
  const agentSession = getAgentSession();
  if (!agentSession) {
    errorAndExit("No agent session. Run 'node notion-query.js login' first.");
  }

  const deadline = expiresAt
    ? new Date(expiresAt).getTime()
    : Date.now() + 5 * 60 * 1000;
  let lastStatus = null;

  while (Date.now() < deadline) {
    const response = await fetch(pollUrl, {
      headers: { "X-Agent-Session": agentSession },
    });

    if (!response.ok) {
      const { data, text } = await readErrorBody(response);
      if (data) {
        describeAuthError(data);
        process.exit(1);
      }
      throw new Error(`Auth request poll failed (${response.status}): ${text || response.statusText}`);
    }

    const status = await response.json();
    if (status.status !== lastStatus) {
      lastStatus = status.status;
      if (status.status === "oauth_pending") {
        console.error("OAuth connection required. Complete it in your browser...");
      }
    }

    if (status.status === "approved") return;
    if (status.status === "denied") throw new Error(status.message || "Authorization denied.");
    if (status.status === "expired") throw new Error(status.message || "Authorization request expired.");

    await sleep(AUTO_REQUEST_POLL_MS);
  }

  throw new Error("Timed out waiting for authorization approval.");
}

async function ensureGrant() {
  if (!AUTO_REQUEST_ENABLED) return;
  const request = await requestGrant();
  if (request.autoApproved) {
    console.error("Access granted automatically (within policy).");
    return;
  }

  console.error("Authorization required for notion access.");
  console.error(`Approve at: ${request.approveUrl}`);
  console.error("Waiting for approval...");
  await waitForGrantApproval(request.pollUrl, request.expiresAt);
  console.error("Authorization granted. Retrying request...");
}

function isProbableNotionId(value) {
  return /^[0-9a-fA-F-]{32,36}$/.test(value);
}

function toPlainRichText(richText = []) {
  return richText.map((t) => t?.plain_text || "").join("");
}

function extractTitleFromProperties(properties = {}) {
  for (const [name, prop] of Object.entries(properties)) {
    if (prop?.type === "title" && Array.isArray(prop.title)) {
      return toPlainRichText(prop.title) || name || "Untitled";
    }
  }
  return "Untitled";
}

function findTitlePropertyName(properties = {}) {
  for (const [name, prop] of Object.entries(properties)) {
    if (prop?.type === "title") return name;
  }
  return "Name";
}

function parseBool(value) {
  const normalized = String(value).trim().toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes";
}

function flattenProperty(prop) {
  if (!prop || typeof prop !== "object") return null;
  switch (prop.type) {
    case "title":
      return toPlainRichText(prop.title);
    case "rich_text":
      return toPlainRichText(prop.rich_text);
    case "select":
      return prop.select?.name ?? null;
    case "status":
      return prop.status?.name ?? null;
    case "multi_select":
      return (prop.multi_select || []).map((x) => x?.name).filter(Boolean).join(", ");
    case "number":
      return prop.number ?? null;
    case "checkbox":
      return prop.checkbox ?? null;
    case "date":
      return prop.date?.start ?? null;
    case "url":
      return prop.url ?? null;
    case "email":
      return prop.email ?? null;
    case "phone_number":
      return prop.phone_number ?? null;
    case "people":
      return (prop.people || [])
        .map((x) => x?.name || x?.id)
        .filter(Boolean)
        .join(", ");
    case "relation":
      return (prop.relation || []).map((x) => x?.id).filter(Boolean).join(", ");
    case "formula": {
      const f = prop.formula || {};
      return f.string ?? f.number ?? f.boolean ?? f.date?.start ?? null;
    }
    case "rollup": {
      const r = prop.rollup || {};
      if (Array.isArray(r.array)) return `${r.array.length} items`;
      return r.number ?? null;
    }
    case "created_time":
      return prop.created_time ?? null;
    case "last_edited_time":
      return prop.last_edited_time ?? null;
    case "created_by":
      return prop.created_by?.name ?? prop.created_by?.id ?? null;
    case "last_edited_by":
      return prop.last_edited_by?.name ?? prop.last_edited_by?.id ?? null;
    case "files":
      return (prop.files || [])
        .map((f) => f?.name || f?.file?.url || f?.external?.url)
        .filter(Boolean)
        .join(", ");
    default:
      return null;
  }
}

function toNotionPropertyValue(type, rawValue) {
  const value = rawValue ?? "";
  switch (type) {
    case "title":
      return { title: [{ text: { content: value } }] };
    case "rich_text":
    case "text":
      return { rich_text: [{ text: { content: value } }] };
    case "select":
      return { select: { name: value } };
    case "status":
      return { status: { name: value } };
    case "multi_select": {
      const names = String(value)
        .split(",")
        .map((x) => x.trim())
        .filter(Boolean)
        .map((name) => ({ name }));
      return { multi_select: names };
    }
    case "number": {
      const n = Number(value);
      return { number: Number.isFinite(n) ? n : null };
    }
    case "checkbox":
      return { checkbox: parseBool(value) };
    case "date":
      return { date: { start: value } };
    case "url":
      return { url: value };
    case "email":
      return { email: value };
    case "phone_number":
      return { phone_number: value };
    case "relation": {
      const relation = String(value)
        .split(",")
        .map((x) => x.trim())
        .filter(Boolean)
        .map((id) => ({ id }));
      return { relation };
    }
    default:
      return { select: { name: value } };
  }
}

function parseExplicitTypedArg(arg) {
  const match = arg.match(
    /^--(select|status|text|number|checkbox|date|url|multi-select|email|phone|relation)-([^=]+)=(.*)$/
  );
  if (!match) return null;
  const [, rawType, name, value] = match;
  const mappedType =
    rawType === "multi-select"
      ? "multi_select"
      : rawType === "phone"
        ? "phone_number"
        : rawType;
  return { type: mappedType, name, value };
}

function printFindDbHint() {
  console.error("Need a database ID? Discover it with:");
  console.error('  notion-query.js find-db "keyword"');
  console.error("  notion-query.js find-db");
}

async function cmdLogin() {
  const initResponse = await fetch(`${AUTH_SERVICE_URL}/api/agent/init`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      agentName: "Fresh Auth CLI",
      deviceInfo: `notion-query.js ${process.platform} ${process.arch} ${hostname()}`,
    }),
  });

  if (!initResponse.ok) {
    const { data, text } = await readErrorBody(initResponse);
    throw new Error(data?.message || data?.error || `Registration init failed: ${text || initResponse.statusText}`);
  }

  const registration = await initResponse.json();
  if (registration.agentSessionId?.trim()) {
    saveAgentSession(registration.agentSessionId);
    console.error("Session saved (pending approval).");
  }
  const registrationId = registration.registrationId;
  const verifyUrl =
    registration.verifyUrl ||
    `${AUTH_SERVICE_URL}/agent/verify?code=${encodeURIComponent(registration.code || "")}`;
  const pollUrl =
    registration.pollUrl || `${AUTH_SERVICE_URL}/api/agent/poll/${registrationId}`;

  console.error("Approve registration at:");
  console.error(`  ${verifyUrl}`);
  console.error("Waiting for approval...");

  const deadline = Date.now() + ((registration.expiresIn || 300) * 1000);
  while (Date.now() < deadline) {
    const pollResponse = await fetch(pollUrl);
    if (!pollResponse.ok) {
      const { data, text } = await readErrorBody(pollResponse);
      throw new Error(data?.message || data?.error || `Registration poll failed: ${text || pollResponse.statusText}`);
    }

    const status = await pollResponse.json();
    if (status.agentSessionId?.trim()) {
      saveAgentSession(status.agentSessionId);
    }
    if (status.status === "approved" && status.agentSessionId) {
      console.error("Login successful.");
      return;
    }
    if (status.status === "denied") throw new Error(status.message || "Registration denied.");
    if (status.status === "expired") throw new Error(status.message || "Registration expired.");

    await sleep(AUTO_REQUEST_POLL_MS);
  }

  throw new Error("Timed out waiting for login approval.");
}

async function cmdLogout() {
  const session = getAgentSession();
  if (!session) {
    console.log("No active session.");
    return;
  }

  await fetch(`${AUTH_SERVICE_URL}/api/agent/session`, {
    method: "DELETE",
    headers: { "X-Agent-Session": session },
  }).catch(() => {});

  clearAgentSession();
  console.log("Session cleared.");
}

async function cmdStatus() {
  const session = getAgentSession();
  if (!session) {
    printJson({
      authService: AUTH_SERVICE_URL,
      authenticated: false,
      reason: "no_agent_session",
      next: "node notion-query.js login",
    });
    return;
  }

  const [agentRes, proxyRes] = await Promise.all([
    fetch(`${AUTH_SERVICE_URL}/api/agent/status`, {
      headers: { "X-Agent-Session": session },
    }),
    fetch(NOTION_STATUS_URL, {
      headers: { "X-Agent-Session": session },
    }),
  ]);

  const { data: agentData } = await readErrorBody(agentRes);
  const { data: proxyData } = await readErrorBody(proxyRes);

  printJson({
    authService: AUTH_SERVICE_URL,
    authenticated: agentRes.ok,
    sessionValid: agentRes.ok,
    user: agentRes.ok ? { name: agentData?.name, email: agentData?.email } : null,
    notionGrant: proxyRes.ok
      ? {
          hasGrant: !!proxyData?.hasGrant,
          expiresAt: proxyData?.expiresAt || null,
          scopes: proxyData?.scopes || [],
        }
      : {
          hasGrant: false,
          error: proxyData?.error || "status_unavailable",
        },
  });
}

async function cmdRequest() {
  const request = await requestGrant();
  if (request.autoApproved) {
    console.log("Access granted automatically (within policy).");
    return;
  }
  console.error("Approval required at:");
  console.error(`  ${request.approveUrl}`);
  console.error("Waiting for approval...");
  await waitForGrantApproval(request.pollUrl, request.expiresAt);
  console.log("Access granted.");
}

async function notionApi(method, endpoint, body, attempt = 0) {
  const agentSession = getAgentSession();
  if (!agentSession) {
    errorAndExit("No agent session. Run 'node notion-query.js login' first.");
  }

  const headers = {
    "X-Agent-Session": agentSession,
    "Notion-Version": NOTION_VERSION,
  };
  const init = { method, headers };
  if (body !== undefined) {
    headers["Content-Type"] = "application/json";
    init.body = JSON.stringify(body);
  }

  const response = await fetch(`${NOTION_PROXY_BASE}${endpoint}`, init);
  if (response.status === 401 || response.status === 403 || response.status === 429) {
    const { data, text } = await readErrorBody(response);
    if (attempt === 0 && AUTO_REQUEST_ENABLED && isAutoRequestError(data)) {
      await ensureGrant();
      return notionApi(method, endpoint, body, attempt + 1);
    }
    if (data) {
      describeAuthError(data);
      process.exit(1);
    }
    throw new Error(text || `Authentication failed (${response.status})`);
  }

  const { data, text } = await readErrorBody(response);
  if (!response.ok) {
    if (data) {
      const error = normalizeAuthError(data);
      if (
        error.error === "no_agent_session" ||
        error.error === "no_oauth_token" ||
        error.reauthorizeUrl ||
        error.elevateUrl
      ) {
        describeAuthError(data);
        process.exit(1);
      }
    }
    if (data?.error || data?.message) {
      throw new Error(data.message || data.error);
    }
    throw new Error(text || `Error: ${response.status} ${response.statusText}`);
  }

  if (data?.object === "error") {
    throw new Error(`Error: ${data.message || "Notion proxy error"}`);
  }

  return data || {};
}

async function listAllBlockChildren(blockId) {
  const results = [];
  let cursor = null;
  do {
    const query = new URLSearchParams({ page_size: "100" });
    if (cursor) query.set("start_cursor", cursor);
    const data = await notionApi(
      "GET",
      `/blocks/${blockId}/children?${query.toString()}`
    );
    results.push(...(data.results || []));
    cursor = data.has_more ? data.next_cursor : null;
  } while (cursor);
  return results;
}

async function appendChildren(parentId, children) {
  const chunks = [];
  for (let i = 0; i < children.length; i += 100) {
    chunks.push(children.slice(i, i + 100));
  }
  for (const chunk of chunks) {
    await notionApi("PATCH", `/blocks/${parentId}/children`, { children: chunk });
  }
}

function blockText(block, type) {
  const rich = block?.[type]?.rich_text || [];
  return toPlainRichText(rich);
}

function blockToMarkdown(block) {
  switch (block.type) {
    case "heading_1":
      return `# ${blockText(block, "heading_1")}`;
    case "heading_2":
      return `## ${blockText(block, "heading_2")}`;
    case "heading_3":
      return `### ${blockText(block, "heading_3")}`;
    case "paragraph":
      return blockText(block, "paragraph");
    case "bulleted_list_item":
      return `- ${blockText(block, "bulleted_list_item")}`;
    case "numbered_list_item":
      return `1. ${blockText(block, "numbered_list_item")}`;
    case "to_do":
      return `- [${block.to_do?.checked ? "x" : " "}] ${blockText(block, "to_do")}`;
    case "quote":
      return `> ${blockText(block, "quote")}`;
    case "code": {
      const language = block.code?.language || "";
      const body = blockText(block, "code");
      return `\`\`\`${language}\n${body}\n\`\`\``;
    }
    case "divider":
      return "---";
    case "callout": {
      const emoji = block.callout?.icon?.emoji || "💡";
      return `> ${emoji} ${blockText(block, "callout")}`;
    }
    default:
      return "";
  }
}

function richText(value) {
  if (!value) return [];
  return [{ type: "text", text: { content: value } }];
}

function markdownToBlocks(markdown) {
  const lines = markdown.replace(/\r\n/g, "\n").split("\n");
  const blocks = [];
  let inCode = false;
  let codeLang = "";
  let codeLines = [];

  for (const line of lines) {
    if (inCode) {
      if (line.startsWith("```")) {
        blocks.push({
          object: "block",
          type: "code",
          code: {
            rich_text: richText(codeLines.join("\n")),
            language: codeLang || "plain text",
          },
        });
        inCode = false;
        codeLang = "";
        codeLines = [];
      } else {
        codeLines.push(line);
      }
      continue;
    }

    if (line.startsWith("```")) {
      inCode = true;
      codeLang = line.slice(3).trim();
      continue;
    }

    const trimmed = line.trim();
    if (!trimmed) continue;

    if (trimmed === "---") {
      blocks.push({ object: "block", type: "divider", divider: {} });
      continue;
    }

    if (trimmed.startsWith("### ")) {
      blocks.push({
        object: "block",
        type: "heading_3",
        heading_3: { rich_text: richText(trimmed.slice(4)) },
      });
      continue;
    }

    if (trimmed.startsWith("## ")) {
      blocks.push({
        object: "block",
        type: "heading_2",
        heading_2: { rich_text: richText(trimmed.slice(3)) },
      });
      continue;
    }

    if (trimmed.startsWith("# ")) {
      blocks.push({
        object: "block",
        type: "heading_1",
        heading_1: { rich_text: richText(trimmed.slice(2)) },
      });
      continue;
    }

    const todo = trimmed.match(/^- \[([ xX])\] (.*)$/);
    if (todo) {
      blocks.push({
        object: "block",
        type: "to_do",
        to_do: {
          rich_text: richText(todo[2] || ""),
          checked: /x/i.test(todo[1]),
        },
      });
      continue;
    }

    if (/^[-*] /.test(trimmed)) {
      blocks.push({
        object: "block",
        type: "bulleted_list_item",
        bulleted_list_item: { rich_text: richText(trimmed.slice(2)) },
      });
      continue;
    }

    if (/^\d+\.\s+/.test(trimmed)) {
      const text = trimmed.replace(/^\d+\.\s+/, "");
      blocks.push({
        object: "block",
        type: "numbered_list_item",
        numbered_list_item: { rich_text: richText(text) },
      });
      continue;
    }

    if (trimmed.startsWith("> ")) {
      blocks.push({
        object: "block",
        type: "quote",
        quote: { rich_text: richText(trimmed.slice(2)) },
      });
      continue;
    }

    blocks.push({
      object: "block",
      type: "paragraph",
      paragraph: { rich_text: richText(trimmed) },
    });
  }

  if (inCode) {
    blocks.push({
      object: "block",
      type: "code",
      code: {
        rich_text: richText(codeLines.join("\n")),
        language: codeLang || "plain text",
      },
    });
  }

  return blocks;
}

async function readStdin() {
  const chunks = [];
  for await (const chunk of process.stdin) chunks.push(chunk);
  return Buffer.concat(chunks.map((c) => Buffer.from(c))).toString("utf8");
}

async function findDatabases(query = "") {
  const payload = {
    query,
    page_size: 100,
    filter: { property: "object", value: "database" },
  };
  const result = await notionApi("POST", "/search", payload);
  return (result.results || []).map((db) => ({
    id: db.id,
    title: toPlainRichText(db.title || []) || "Untitled",
    url: db.url,
  }));
}

async function cmdMe() {
  const me = await notionApi("GET", "/users/me");
  printJson({
    id: me.id,
    name: me.name,
    type: me.type,
    object: me.object,
  });
}

async function cmdSearch(query = "") {
  const result = await notionApi("POST", "/search", { query, page_size: 100 });
  const out = (result.results || []).map((item) => ({
    id: item.id,
    type: item.object,
    title:
      item.object === "database"
        ? toPlainRichText(item.title || []) || "Untitled"
        : extractTitleFromProperties(item.properties || {}),
    url: item.url,
  }));
  printJson(out);
}

async function cmdFindDb(query = "") {
  const out = await findDatabases(query);
  printJson(out);
}

function parseQueryDbFlags(args) {
  let raw = false;
  let filterProp = "";
  let filterValue = "";
  for (const arg of args) {
    if (arg === "--raw") raw = true;
    if (arg.startsWith("--filter-")) {
      const [left, value] = splitFirst(arg, "=");
      filterProp = left.replace(/^--filter-/, "");
      filterValue = value;
    }
  }
  return { raw, filterProp, filterValue };
}

function parsePropertyArgs(args, schema, titlePropertyName) {
  const props = {};
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i];
    if (arg === "-p") {
      const pair = args[i + 1];
      i += 1;
      if (!pair || !pair.includes("=")) continue;
      const [name, value] = splitFirst(pair, "=");
      const propertyType = schema?.properties?.[name]?.type || "select";
      props[name] = toNotionPropertyValue(propertyType, value);
      continue;
    }

    if (arg.startsWith("--title=")) {
      const value = arg.slice("--title=".length);
      props[titlePropertyName] = toNotionPropertyValue("title", value);
      continue;
    }

    const typed = parseExplicitTypedArg(arg);
    if (typed) {
      props[typed.name] = toNotionPropertyValue(typed.type, typed.value);
    }
  }
  return props;
}

async function cmdQueryDb(dbId, flags) {
  if (!dbId) {
    console.error("Usage: notion-query.js query-db <database_id> [--raw] [--filter-prop=value]");
    console.error("");
    printFindDbHint();
    process.exit(1);
  }
  if (!isProbableNotionId(dbId)) {
    console.error(`Error: '${dbId}' does not look like a Notion database ID.`);
    console.error(`Database matches for '${dbId}':`);
    printJson(await findDatabases(dbId));
    process.exit(1);
  }

  const result = await notionApi("POST", `/databases/${dbId}/query`, {});
  if (flags.raw) {
    printJson(result);
    return;
  }

  let rows = (result.results || []).map((page) => {
    const row = { id: page.id, url: page.url };
    for (const [key, value] of Object.entries(page.properties || {})) {
      row[key] = flattenProperty(value);
    }
    return row;
  });

  if (flags.filterProp && flags.filterValue) {
    rows = rows.filter(
      (row) => String(row[flags.filterProp] ?? "") === flags.filterValue
    );
  }

  printJson(rows);
}

async function cmdBacklog(roadmapFilter = "") {
  if (!NOTION_BACKLOG_DB_ID) {
    errorAndExit(
      "Error: NOTION_BACKLOG_DB_ID environment variable not set\nSet NOTION_BACKLOG_DB_ID to use backlog shortcuts."
    );
  }
  const result = await notionApi("POST", `/databases/${NOTION_BACKLOG_DB_ID}/query`, {});
  const items = (result.results || []).map((page) => ({
    title: extractTitleFromProperties(page.properties || {}),
    roadmap:
      page.properties?.Roadmap?.select?.name ||
      page.properties?.Roadmap?.status?.name ||
      "Unspecified",
    status:
      page.properties?.Status?.status?.name ||
      page.properties?.Status?.select?.name ||
      null,
  }));

  const filtered = roadmapFilter
    ? items.filter((item) => item.roadmap === roadmapFilter)
    : items;

  const grouped = new Map();
  for (const item of filtered) {
    if (!grouped.has(item.roadmap)) grouped.set(item.roadmap, []);
    grouped.get(item.roadmap).push({ title: item.title, status: item.status });
  }

  const out = [...grouped.entries()].map(([roadmap, groupedItems]) => ({
    roadmap,
    items: groupedItems,
  }));

  printJson(out);
}

async function cmdGetPage(pageId, raw = false) {
  if (!pageId) errorAndExit("Usage: notion-query.js get-page <page_id> [--raw]");
  const page = await notionApi("GET", `/pages/${pageId}`);
  if (raw) {
    printJson(page);
    return;
  }
  const out = { id: page.id, url: page.url, title: extractTitleFromProperties(page.properties || {}) };
  for (const [key, value] of Object.entries(page.properties || {})) {
    out[key] = flattenProperty(value);
  }
  printJson(out);
}

async function cmdGetBlocks(blockId, raw = false) {
  if (!blockId) errorAndExit("Usage: notion-query.js get-blocks <block_id> [--raw]");
  const blocks = await listAllBlockChildren(blockId);
  if (raw) {
    printJson({ results: blocks });
    return;
  }
  const out = blocks
    .map((block) => ({
      id: block.id,
      type: block.type,
      text: blockToMarkdown(block).replace(/^#+\s/, "").replace(/^- \[[ x]\]\s/, "").replace(/^- /, ""),
      checked: block.type === "to_do" ? Boolean(block.to_do?.checked) : null,
    }))
    .filter((x) => x.text);
  printJson(out);
}

async function cmdGetMarkdown(pageId) {
  if (!pageId) errorAndExit("Usage: notion-query.js get-markdown <page_id>");
  const blocks = await listAllBlockChildren(pageId);
  const lines = blocks.map(blockToMarkdown).filter(Boolean);
  console.log(lines.join("\n\n"));
}

async function cmdGetDb(dbId, raw = false) {
  if (!dbId) {
    console.error("Usage: notion-query.js get-db <database_id> [--raw]");
    console.error("");
    printFindDbHint();
    process.exit(1);
  }
  if (!isProbableNotionId(dbId)) {
    console.error(`Error: '${dbId}' does not look like a Notion database ID.`);
    console.error(`Database matches for '${dbId}':`);
    printJson(await findDatabases(dbId));
    process.exit(1);
  }

  const db = await notionApi("GET", `/databases/${dbId}`);
  if (raw) {
    printJson(db);
    return;
  }

  const properties = Object.entries(db.properties || {}).map(([name, value]) => ({
    name,
    type: value.type,
    options:
      value.select?.options?.map((o) => o.name) ||
      value.status?.options?.map((o) => o.name) ||
      value.multi_select?.options?.map((o) => o.name) ||
      null,
  }));

  printJson({
    id: db.id,
    title: toPlainRichText(db.title || []) || "Untitled",
    properties,
  });
}

async function cmdCreate(dbId, title, propArgs) {
  if (!dbId || !title) {
    console.error("Usage: notion-query.js create <db_id> <title> [-p NAME=VALUE]...");
    console.error("       notion-query.js create <db_id> <title> [--TYPE-NAME=VALUE]...");
    console.error("");
    console.error(
      "Property types: --select-NAME, --status-NAME, --text-NAME, --number-NAME, --checkbox-NAME, --date-NAME, --url-NAME"
    );
    console.error("");
    printFindDbHint();
    process.exit(1);
  }
  if (!isProbableNotionId(dbId)) {
    console.error(`Error: '${dbId}' does not look like a Notion database ID.`);
    console.error(`Database matches for '${dbId}':`);
    printJson(await findDatabases(dbId));
    process.exit(1);
  }

  const schema = await notionApi("GET", `/databases/${dbId}`);
  const titlePropertyName = findTitlePropertyName(schema.properties || {});
  const props = parsePropertyArgs(propArgs, schema, titlePropertyName);
  props[titlePropertyName] = toNotionPropertyValue("title", title).title
    ? { title: toNotionPropertyValue("title", title).title }
    : toNotionPropertyValue("title", title);

  const page = await notionApi("POST", "/pages", {
    parent: { database_id: dbId },
    properties: props,
  });

  printJson({
    id: page.id,
    url: page.url,
    title: extractTitleFromProperties(page.properties || {}),
  });
}

async function cmdCreateBacklog(title, args) {
  if (!title) {
    errorAndExit("Usage: notion-query.js create-backlog <title> [--roadmap=X] [--status=X]");
  }
  if (!NOTION_BACKLOG_DB_ID) {
    errorAndExit(
      "Error: NOTION_BACKLOG_DB_ID environment variable not set\nSet NOTION_BACKLOG_DB_ID to use backlog shortcuts."
    );
  }

  const translatedArgs = [];
  for (const arg of args) {
    if (arg.startsWith("--roadmap=")) {
      translatedArgs.push("-p", `Roadmap=${arg.slice("--roadmap=".length)}`);
    } else if (arg.startsWith("--status=")) {
      translatedArgs.push("-p", `Status=${arg.slice("--status=".length)}`);
    } else {
      translatedArgs.push(arg);
    }
  }

  await cmdCreate(NOTION_BACKLOG_DB_ID, title, translatedArgs);
}

async function cmdUpdate(pageId, propArgs) {
  if (!pageId) {
    console.error("Usage: notion-query.js update <page_id> [-p NAME=VALUE]...");
    console.error("       notion-query.js update <page_id> [--title=X] [--TYPE-NAME=VALUE]...");
    process.exit(1);
  }

  const pageInfo = await notionApi("GET", `/pages/${pageId}`);
  const parentDbId = pageInfo.parent?.database_id || "";
  let schema = null;
  if (parentDbId) {
    try {
      schema = await notionApi("GET", `/databases/${parentDbId}`);
    } catch {
      schema = null;
    }
  }

  const titlePropertyName = findTitlePropertyName(
    schema?.properties || pageInfo.properties || {}
  );
  const props = parsePropertyArgs(propArgs, schema, titlePropertyName);
  if (Object.keys(props).length === 0) {
    errorAndExit("No properties provided. Use -p NAME=VALUE or --TYPE-NAME=VALUE.");
  }

  const updated = await notionApi("PATCH", `/pages/${pageId}`, {
    properties: props,
  });
  printJson({
    id: updated.id,
    url: updated.url,
    title: extractTitleFromProperties(updated.properties || {}),
  });
}

async function cmdArchive(pageId) {
  if (!pageId) errorAndExit("Usage: notion-query.js archive <page_id>");
  const archived = await notionApi("PATCH", `/pages/${pageId}`, { archived: true });
  printJson({ id: archived.id, archived: archived.archived, url: archived.url });
}

async function cmdSetBody(pageId, markdownArg) {
  if (!pageId || !markdownArg) {
    errorAndExit(
      "Usage: notion-query.js set-body <page_id> <markdown>\n       echo 'markdown' | notion-query.js set-body <page_id> -"
    );
  }
  const markdown = markdownArg === "-" ? await readStdin() : markdownArg;
  const existing = await listAllBlockChildren(pageId);
  for (const block of existing) {
    await notionApi("PATCH", `/blocks/${block.id}`, { archived: true });
  }
  const children = markdownToBlocks(markdown);
  if (children.length > 0) await appendChildren(pageId, children);
  printJson({ id: pageId, replaced: true, blocks_written: children.length });
}

async function cmdAppendBody(pageId, markdownArg) {
  if (!pageId || !markdownArg) {
    errorAndExit(
      "Usage: notion-query.js append-body <page_id> <markdown>\n       echo 'markdown' | notion-query.js append-body <page_id> -"
    );
  }
  const markdown = markdownArg === "-" ? await readStdin() : markdownArg;
  const children = markdownToBlocks(markdown);
  if (children.length > 0) await appendChildren(pageId, children);
  printJson({ id: pageId, appended: true, blocks_written: children.length });
}

function printUsage() {
  console.log(`Notion CLI (OAuth via auth service, no API key required)

Usage: notion-query.js <command> [args]

AUTH Commands:
  login                        Register agent session
  logout                       Clear saved agent session
  status                       Check session and notion grant status
  request                      Request notion access grant

READ Commands:
  me                           Get bot info
  search [query]               Search pages/databases (compact)
  find-db [query]              Find databases and show database IDs
  query-db <db_id> [opts]      Query database (all properties)
  backlog [roadmap]            Query configured backlog DB (shortcut)
  get-page <page_id>           Get page properties (compact)
  get-blocks <block_id>        Get block text content (JSON)
  get-markdown <page_id>       Get page content as markdown
  get-db <db_id>               Get database schema (property types)

WRITE Commands:
  create <db_id> <title>       Create page with properties
  create-backlog <title>       Create item in configured backlog DB
  update <page_id>             Update page with properties
  archive <page_id>            Archive (trash) a page
  set-body <page_id> <md>      Replace page content with markdown
  append-body <page_id> <md>   Append markdown to page

Generic Property Options (for create/update):
  -p NAME=VALUE                Auto-detect type from schema
  --title=X                    Set page title
  --select-NAME=VALUE          Set select property
  --status-NAME=VALUE          Set status property
  --text-NAME=VALUE            Set rich_text property
  --number-NAME=VALUE          Set number property
  --checkbox-NAME=VALUE        Set checkbox (true/false)
  --date-NAME=VALUE            Set date (YYYY-MM-DD)
  --url-NAME=VALUE             Set URL property
  --multi-select-NAME=VALUE    Add to multi_select
  --email-NAME=VALUE           Set email property
  --phone-NAME=VALUE           Set phone number property
  --relation-NAME=ID1,ID2      Set relation property

Query Options:
  --raw                        Return full API response
  --filter-PROP=VALUE          Filter query-db rows by property

Environment:
  AUTH_SERVICE_URL             Auth service URL (default: https://auth.freshhub.ai)
  OFFICE_AUTO_REQUEST=0        Disable auto-request on missing grants
  NOTION_API_VERSION           Notion API version (default: 2022-06-28)
  NOTION_BACKLOG_DB_ID         Optional DB ID for backlog/create-backlog

Examples:
  notion-query.js login
  notion-query.js request
  notion-query.js find-db "Roadmap"
  notion-query.js get-db <db_id>
  notion-query.js query-db <db_id> --filter-Status="In progress"
  notion-query.js create <db_id> "New Task" -p "Status=In progress" -p "Priority=High"
`);
}

async function main() {
  const [cmd, ...rest] = process.argv.slice(2);
  switch (cmd) {
    case "login":
      await cmdLogin();
      return;
    case "logout":
      await cmdLogout();
      return;
    case "status":
      await cmdStatus();
      return;
    case "request":
      await cmdRequest();
      return;
    case "me":
      await cmdMe();
      return;
    case "search":
      await cmdSearch(rest[0] || "");
      return;
    case "find-db":
    case "find-dbs":
    case "list-dbs":
      await cmdFindDb(rest[0] || "");
      return;
    case "query-db": {
      const dbId = rest[0];
      const flags = parseQueryDbFlags(rest.slice(1));
      await cmdQueryDb(dbId, flags);
      return;
    }
    case "backlog":
      await cmdBacklog(rest[0] || "");
      return;
    case "get-page":
      await cmdGetPage(rest[0], rest[1] === "--raw");
      return;
    case "get-blocks":
      await cmdGetBlocks(rest[0], rest[1] === "--raw");
      return;
    case "get-markdown":
      await cmdGetMarkdown(rest[0]);
      return;
    case "get-db":
      await cmdGetDb(rest[0], rest[1] === "--raw");
      return;
    case "create":
      await cmdCreate(rest[0], rest[1], rest.slice(2));
      return;
    case "create-backlog":
      await cmdCreateBacklog(rest[0], rest.slice(1));
      return;
    case "update":
      await cmdUpdate(rest[0], rest.slice(1));
      return;
    case "archive":
      await cmdArchive(rest[0]);
      return;
    case "set-body":
      await cmdSetBody(rest[0], rest[1]);
      return;
    case "append-body":
      await cmdAppendBody(rest[0], rest[1]);
      return;
    case "-h":
    case "--help":
    case undefined:
      printUsage();
      return;
    default:
      console.error(`Unknown command: ${cmd}`);
      printUsage();
      process.exit(1);
  }
}

main().catch((error) => {
  console.error(error?.message || String(error));
  process.exit(1);
});
