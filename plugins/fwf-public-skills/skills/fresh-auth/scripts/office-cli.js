#!/usr/bin/env node
import { mkdir, writeFile, readFile, unlink, readdir } from "fs/promises";
import { dirname, join, extname } from "path";
import { existsSync } from "fs";
import { homedir, tmpdir, hostname } from "os";
import { spawn } from "child_process";
const RAW_AUTH_SERVICE_URL = process.env.AUTH_SERVICE_URL || "https://auth.freshhub.ai";
const AUTH_SERVICE_URL = RAW_AUTH_SERVICE_URL.replace(/\/+$/, "").replace(
  /\/api$/,
  ""
);
const API_BASE = `${AUTH_SERVICE_URL}/proxy/msgraph`;
const STATUS_BASE = `${AUTH_SERVICE_URL}/api/proxy/status`;
const DEFAULT_CONFIG_HOME = join(homedir(), ".config");
function normalizeSessionFile(value) {
  const trimmed = value?.trim();
  if (!trimmed) return null;
  if (trimmed === "~") return homedir();
  if (trimmed.startsWith("~/")) return join(homedir(), trimmed.slice(2));
  return trimmed;
}
function uniqueSessionFiles(values) {
  const unique = /* @__PURE__ */ new Set();
  for (const value of values) {
    if (!value) continue;
    unique.add(value);
  }
  return [...unique];
}
const XDG_CONFIG_HOME = normalizeSessionFile(process.env.XDG_CONFIG_HOME);
const CONFIG_HOME_CANDIDATES = uniqueSessionFiles([
  XDG_CONFIG_HOME,
  DEFAULT_CONFIG_HOME
]);
const SHARED_AGENT_SESSION_FILE = normalizeSessionFile(process.env.FRESH_AUTH_AGENT_SESSION_FILE) || join(CONFIG_HOME_CANDIDATES[0], "fresh-auth", "agent-session");
const LEGACY_AGENT_SESSION_FILES = uniqueSessionFiles([
  normalizeSessionFile(process.env.OFFICE_CLI_AGENT_SESSION_FILE),
  normalizeSessionFile(process.env.AGENT_SESSION_FILE),
  ...CONFIG_HOME_CANDIDATES.map((configHome) => join(configHome, "office-cli", "agent-session")),
  ...CONFIG_HOME_CANDIDATES.map((configHome) => join(configHome, "cal-cli", "agent-session")),
  ...CONFIG_HOME_CANDIDATES.map((configHome) => join(configHome, "onedrive-cli", "agent-session"))
]).filter((sessionFile) => sessionFile !== SHARED_AGENT_SESSION_FILE);
const AGENT_SESSION_FILES = [
  SHARED_AGENT_SESSION_FILE,
  ...LEGACY_AGENT_SESSION_FILES
];
const AUTO_REQUEST_POLL_MS = 2e3;
const AUTO_REQUEST_ENABLED = process.env.OFFICE_AUTO_REQUEST !== "0";
const SCOPES = {
  drive: ["Files.Read", "Files.ReadWrite.All"],
  mail: ["Mail.Read", "Mail.Send", "People.Read"],
  cal: ["Calendars.Read"]
};
const GRANT_DURATION = {
  drive: "30m",
  mail: "1h",
  cal: "1h"
};
async function getAgentSession() {
  try {
    for (const sessionFile of AGENT_SESSION_FILES) {
      if (!existsSync(sessionFile)) continue;
      const sessionRaw = (await readFile(sessionFile, "utf-8")).trim();
      if (!sessionRaw) continue;
      try {
        const parsed = JSON.parse(sessionRaw);
        const normalized = parsed.agentSessionId || parsed.agentSession || parsed.session;
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
async function saveAgentSession(agentSessionId) {
  const dir = dirname(SHARED_AGENT_SESSION_FILE);
  await mkdir(dir, { recursive: true });
  await writeFile(SHARED_AGENT_SESSION_FILE, `${agentSessionId.trim()}\n`, { mode: 384 });
}
async function clearAgentSession() {
  try {
    for (const sessionFile of AGENT_SESSION_FILES) {
      if (existsSync(sessionFile)) {
        await unlink(sessionFile);
      }
    }
  } catch {
  }
}
function buildRegistrationUrl(code) {
  return `${AUTH_SERVICE_URL}/agent/verify?code=${encodeURIComponent(code)}`;
}
function describeAuthError(data, service) {
  if (data.error === "no_agent_session") {
    console.error("No agent session. Run 'office-cli.js login' to register.");
    return;
  }
  if (["no_grant", "grant_expired", "single_use_exhausted"].includes(
    data.error || ""
  )) {
    console.error(`No active grant for ${service}.`);
    console.error("");
    console.error("Request access by running:");
    console.error(`  ./office-cli.js request ${service}`);
    return;
  }
  if (data.error === "no_oauth_token" && data.connectUrl) {
    console.error("Microsoft account not linked.");
    console.error("");
    console.error("To connect, visit:");
    console.error(`  ${data.connectUrl}`);
    return;
  }
  if (data.reauthorizeUrl) {
    console.error("Access token expired or missing scopes.");
    if (data.missingScopes?.length) {
      console.error(`Missing scopes: ${data.missingScopes.join(" ")}`);
    }
    console.error("");
    console.error("Re-authorise at:");
    console.error(`  ${data.reauthorizeUrl}`);
    return;
  }
  if (data.elevateUrl) {
    console.error(data.error || "Access not authorised.");
    console.error("");
    console.error("To grant access, visit:");
    console.error(`  ${data.elevateUrl}`);
    return;
  }
  console.error(data.message || data.error || "Authentication failed");
}
function isAutoRequestError(data) {
  if (!data?.error) return false;
  return ["no_grant", "grant_expired", "single_use_exhausted"].includes(
    data.error
  );
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
async function requestGrant(service) {
  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run './office-cli.js login' first.");
    process.exit(1);
  }
  const response = await fetch(`${AUTH_SERVICE_URL}/api/auth-request`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Agent-Session": agentSession
    },
    body: JSON.stringify({
      service: "msgraph",
      scopes: SCOPES[service],
      duration: GRANT_DURATION[service]
    })
  });
  if (!response.ok) {
    const { data, text } = await readErrorBody(response);
    if (data?.error || data?.message) {
      throw new Error(data.message || data.error);
    }
    throw new Error(
      `Failed to create auth request (${response.status}): ${text || response.statusText}`
    );
  }
  return response.json();
}
async function waitForGrantApproval(pollUrl, expiresAt) {
  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run './office-cli.js login' first.");
    process.exit(1);
  }
  const deadline = expiresAt ? new Date(expiresAt).getTime() : Date.now() + 5 * 60 * 1e3;
  let lastStatus = null;
  while (Date.now() < deadline) {
    const response = await fetch(pollUrl, {
      headers: { "X-Agent-Session": agentSession }
    });
    if (!response.ok) {
      const { data, text } = await readErrorBody(response);
      if (data) {
        describeAuthError(data, "msgraph");
        process.exit(1);
      }
      throw new Error(
        `Auth request poll failed (${response.status}): ${text || response.statusText}`
      );
    }
    const status = await response.json();
    if (status.status !== lastStatus) {
      lastStatus = status.status;
      if (status.status === "oauth_pending") {
        console.error(
          "OAuth connection required. Complete it in your browser..."
        );
      }
    }
    if (status.status === "approved") return;
    if (status.status === "denied")
      throw new Error(status.message || "Authorization denied.");
    if (status.status === "expired")
      throw new Error(status.message || "Authorization request expired.");
    await new Promise((resolve) => setTimeout(resolve, AUTO_REQUEST_POLL_MS));
  }
  throw new Error("Timed out waiting for authorization approval.");
}
let currentService = "drive";
async function ensureGrant() {
  if (!AUTO_REQUEST_ENABLED) return;
  const request = await requestGrant(currentService);
  if (request.autoApproved) {
    console.error("Access granted automatically (within policy).");
    return;
  }
  console.error(`Authorization required for ${currentService} access.`);
  console.error(`Approve at: ${request.approveUrl}`);
  console.error("Waiting for approval...");
  await waitForGrantApproval(request.pollUrl, request.expiresAt);
  console.error("Authorization granted. Retrying request...");
}
async function apiRequest(endpoint, options = {}, attempt = 0) {
  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run 'login' to create an agent session.");
    process.exit(1);
  }
  const url = `${API_BASE}${endpoint}`;
  const response = await fetch(url, {
    ...options,
    headers: {
      "X-Agent-Session": agentSession,
      "Content-Type": "application/json",
      ...options.headers
    },
    redirect: "manual"
  });
  if (response.status === 401 || response.status === 403 || response.status === 429) {
    const { data, text } = await readErrorBody(response);
    if (attempt === 0 && AUTO_REQUEST_ENABLED && isAutoRequestError(data)) {
      await ensureGrant();
      return apiRequest(endpoint, options, attempt + 1);
    }
    if (data) {
      describeAuthError(data, currentService);
      process.exit(1);
    }
    throw new Error(text || `Authentication failed (${response.status})`);
  }
  if (!response.ok) {
    const { data, text } = await readErrorBody(response);
    if (data?.error || data?.message) {
      throw new Error(data.message || data.error);
    }
    throw new Error(text || `API error (${response.status})`);
  }
  return response.json();
}
async function fetchGraphContent(endpoint, attempt = 0) {
  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run 'login' to create an agent session.");
    process.exit(1);
  }
  const response = await fetch(`${API_BASE}${endpoint}`, {
    headers: { "X-Agent-Session": agentSession },
    redirect: "manual"
  });
  if (response.status === 401 || response.status === 403 || response.status === 429) {
    const { data, text } = await readErrorBody(response);
    if (attempt === 0 && AUTO_REQUEST_ENABLED && isAutoRequestError(data)) {
      await ensureGrant();
      return fetchGraphContent(endpoint, attempt + 1);
    }
    if (data) {
      describeAuthError(data, currentService);
      process.exit(1);
    }
    throw new Error(text || `Authentication failed (${response.status})`);
  }
  return response;
}
function formatBytes(bytes) {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
}
function stripHtml(html) {
  if (!html) return "";
  return html.replace(/<style[\s\S]*?<\/style>/gi, "").replace(/<script[\s\S]*?<\/script>/gi, "").replace(/<\/?[^>]+>/g, "").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/\s+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}
function truncate(str, len) {
  if (str.length <= len) return str;
  return str.substring(0, len - 1) + "\u2026";
}
async function driveList(folderPath) {
  let endpoint = "/me/drive/root/children";
  if (folderPath) {
    const encodedPath = folderPath.split("/").map(encodeURIComponent).join("/");
    endpoint = `/me/drive/root:${encodedPath}:/children`;
  }
  const query = "?$select=id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,parentReference&$top=100";
  const result = await apiRequest(
    `${endpoint}${query}`
  );
  if (!result.value || result.value.length === 0) {
    console.log("No files found.");
    return;
  }
  console.log(`
Files in ${folderPath || "root"}:
`);
  console.log("\u2500".repeat(90));
  console.log(
    `${"Name".padEnd(40)} ${"Size".padEnd(10)} ${"Modified".padEnd(20)} ID`
  );
  console.log("\u2500".repeat(90));
  for (const item of result.value) {
    const isFolder = !!item.folder;
    const name = isFolder ? `\u{1F4C1} ${item.name}` : `\u{1F4C4} ${item.name}`;
    const size = item.size ? formatBytes(item.size) : "-";
    const modified = new Date(item.lastModifiedDateTime).toLocaleDateString();
    console.log(
      `${name.padEnd(40)} ${size.padEnd(10)} ${modified.padEnd(20)} ${item.id}`
    );
  }
}
async function driveSearch(query) {
  const endpoint = `/me/drive/root/search(q='${encodeURIComponent(query)}')`;
  const select = "?$select=id,name,size,lastModifiedDateTime,webUrl,file,folder,parentReference&$top=25";
  const result = await apiRequest(
    `${endpoint}${select}`
  );
  if (!result.value || result.value.length === 0) {
    console.log(`No files found matching "${query}".`);
    return;
  }
  const ids = result.value.map((item) => item.id);
  let pathMap = {};
  const chunks = [];
  for (let i = 0; i < ids.length; i += 20) {
    chunks.push(ids.slice(i, i + 20));
  }
  const batchResults = await Promise.all(
    chunks.map(async (chunk) => {
      try {
        const batchRequests = chunk.map((id, index) => ({
          id: String(index),
          method: "GET",
          url: `/me/drive/items/${id}?$select=id,name,parentReference`
        }));
        const batchResult = await apiRequest("/$batch", {
          method: "POST",
          body: JSON.stringify({ requests: batchRequests })
        });
        const paths = {};
        for (const resp of batchResult.responses || []) {
          if (resp.status === 200 && resp.body?.parentReference?.path) {
            const parsedIdx = parseInt(resp.id, 10);
            if (Number.isNaN(parsedIdx) || parsedIdx < 0 || parsedIdx >= chunk.length) continue;
            const itemId = chunk[parsedIdx];
            const rawPath = resp.body.parentReference.path;
            paths[itemId] = rawPath.replace(/^\/drive\/root:?/, "") || "/";
          }
        }
        return paths;
      } catch {
        return {};
      }
    })
  );
  for (const paths of batchResults) {
    Object.assign(pathMap, paths);
  }
  console.log(`
Search results for "${query}":
`);
  for (const item of result.value) {
    const icon = item.folder ? "\u{1F4C1}" : "\u{1F4C4}";
    const path = pathMap[item.id] || "/";
    console.log(`${icon} ${item.name}`);
    console.log(`   ID: ${item.id}`);
    console.log(`   Path: ${path}`);
    console.log(
      `   Modified: ${new Date(item.lastModifiedDateTime).toLocaleString()}`
    );
    if (item.size) console.log(`   Size: ${formatBytes(item.size)}`);
    console.log();
  }
}
async function driveDownload(itemId, outputPath) {
  const info = await apiRequest(`/me/drive/items/${itemId}`);
  if (info.folder) throw new Error("Cannot download a folder.");
  const fileName = outputPath || info.name;
  console.error(
    `Downloading ${info.name} (${formatBytes(info.size || 0)})...`
  );
  let response = await fetchGraphContent(`/me/drive/items/${itemId}/content`);
  if (response.status === 302 || response.status === 301) {
    const downloadUrl = response.headers.get("location");
    if (!downloadUrl) throw new Error("Download redirect missing location");
    response = await fetch(downloadUrl);
  }
  if (!response.ok) throw new Error(`Download failed (${response.status})`);
  const content = await response.arrayBuffer();
  const dir = dirname(fileName);
  if (dir && dir !== ".") await mkdir(dir, { recursive: true });
  await writeFile(fileName, Buffer.from(content));
  console.error(`Saved to: ${fileName}`);
  return fileName;
}
async function driveContent(itemId) {
  const info = await apiRequest(
    `/me/drive/items/${itemId}?$select=id,name,size,file,folder`
  );
  if (info.folder) throw new Error("Cannot get content of a folder");
  const textExtensions = [
    "txt",
    "md",
    "json",
    "xml",
    "csv",
    "html",
    "css",
    "js",
    "ts",
    "py",
    "sh",
    "yaml",
    "yml",
    "log",
    "ini",
    "conf"
  ];
  const ext = info.name.split(".").pop()?.toLowerCase() || "";
  const mimeType = info.file?.mimeType || "";
  const isText = textExtensions.includes(ext) || mimeType.startsWith("text/");
  if (!isText)
    throw new Error("Not a text file. Use 'drive download' instead.");
  let response = await fetchGraphContent(`/me/drive/items/${itemId}/content`);
  if (response.status === 302 || response.status === 301) {
    const downloadUrl = response.headers.get("location");
    if (!downloadUrl) throw new Error("Download redirect missing location");
    response = await fetch(downloadUrl);
  }
  if (!response.ok) throw new Error(`Content download failed (${response.status})`);
  console.log(await response.text());
}
async function driveInfo(itemId) {
  const item = await apiRequest(`/me/drive/items/${itemId}`);
  const rawPath = item.parentReference?.path || "";
  const path = rawPath.replace(/^\/drive\/root:?/, "") || "/";
  console.log("\nFile Information:");
  console.log("\u2500".repeat(50));
  console.log(`Name: ${item.name}`);
  console.log(`ID: ${item.id}`);
  console.log(`Type: ${item.folder ? "Folder" : "File"}`);
  if (item.file) console.log(`MIME Type: ${item.file.mimeType}`);
  if (item.size) console.log(`Size: ${formatBytes(item.size)}`);
  console.log(`Created: ${new Date(item.createdDateTime).toLocaleString()}`);
  console.log(
    `Modified: ${new Date(item.lastModifiedDateTime).toLocaleString()}`
  );
  console.log(`Web URL: ${item.webUrl}`);
  console.log(`Path: ${path}`);
}
async function driveShare(itemId, options) {
  if (options.anyone) {
    options.scope = options.scope || "anonymous";
    options.type = options.type || "edit";
    if (!options.expires) {
      const expiry = /* @__PURE__ */ new Date();
      expiry.setDate(expiry.getDate() + 7);
      options.expires = expiry.toISOString();
    }
  }
  if (options.email) {
    const role = options.role || "read";
    if (!["read", "write"].includes(role)) {
      throw new Error("Role must be 'read' or 'write'");
    }
    const result = await apiRequest(
      `/me/drive/items/${itemId}/invite`,
      {
        method: "POST",
        body: JSON.stringify({
          recipients: [{ email: options.email }],
          requireSignIn: true,
          sendInvitation: true,
          roles: [role],
          message: "Shared via Office CLI"
        })
      }
    );
    console.log(`
Shared with ${options.email} (${role} access)`);
    if (result.value?.[0]?.link?.webUrl) {
      console.log(`Link: ${result.value[0].link.webUrl}`);
    }
  } else {
    const type = options.type || "view";
    const scope = options.scope || "organization";
    if (!["view", "edit", "embed"].includes(type)) {
      throw new Error("Link type must be 'view', 'edit', or 'embed'");
    }
    if (!["anonymous", "organization", "users"].includes(scope)) {
      throw new Error("Scope must be 'anonymous', 'organization', or 'users'");
    }
    const body = { type, scope };
    if (options.expires) {
      body.expirationDateTime = options.expires.includes("T") ? options.expires : `${options.expires}T00:00:00Z`;
    }
    try {
      const result = await apiRequest(
        `/me/drive/items/${itemId}/createLink`,
        {
          method: "POST",
          body: JSON.stringify(body)
        }
      );
      console.log(`
Share link created (${type}, ${scope}):`);
      console.log(result.link.webUrl);
      if (options.expires) {
        const expiryDate = new Date(body.expirationDateTime);
        console.log(`Expires: ${expiryDate.toLocaleDateString("en-AU", { weekday: "short", year: "numeric", month: "short", day: "numeric" })}`);
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      if (scope === "anonymous" && (msg.includes("invalidRequest") || msg.includes("not allowed") || msg.includes("disabled"))) {
        console.error(`
Failed to create anonymous ${type} link.`);
        console.error("Anonymous sharing may be disabled by your SharePoint admin.");
        console.error("Admin needs to enable in: SharePoint Admin > Policies > Sharing > set to 'Anyone'");
        console.error("\nTry --scope organization instead for org-wide access.");
        process.exit(1);
      }
      throw err;
    }
  }
}
async function drivePermissions(itemId) {
  const info = await apiRequest(`/me/drive/items/${itemId}?$select=id,name`);
  const result = await apiRequest(
    `/me/drive/items/${itemId}/permissions`
  );
  if (!result.value || result.value.length === 0) {
    console.log(`No permissions found for ${info.name}.`);
    return;
  }
  console.log(`
Permissions for "${info.name}":
`);
  console.log("\u2500".repeat(80));
  for (const perm of result.value) {
    const roles = perm.roles.join(", ");
    if (perm.link) {
      console.log(`\u{1F517} Link (${perm.link.type}, ${perm.link.scope})`);
      console.log(`   URL: ${perm.link.webUrl}`);
      console.log(`   Roles: ${roles}`);
      console.log(`   Permission ID: ${perm.id}`);
    } else {
      const user = perm.grantedToV2?.user || perm.grantedToV2?.siteUser || perm.grantedTo?.user;
      const email = user?.email || perm.invitation?.email || "unknown";
      const name = user?.displayName || email;
      console.log(`\u{1F464} ${name}`);
      if (email !== name) console.log(`   Email: ${email}`);
      console.log(`   Roles: ${roles}`);
      console.log(`   Permission ID: ${perm.id}`);
    }
    console.log();
  }
}
async function driveUnshare(itemId, permissionId) {
  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run './office-cli.js login' first.");
    process.exit(1);
  }
  const response = await fetch(
    `${API_BASE}/me/drive/items/${itemId}/permissions/${permissionId}`,
    {
      method: "DELETE",
      headers: { "X-Agent-Session": agentSession }
    }
  );
  if (response.status === 204 || response.status === 200) {
    console.log(`Permission ${permissionId} removed.`);
    return;
  }
  if (response.status === 401 || response.status === 403) {
    const { data } = await readErrorBody(response);
    if (isAutoRequestError(data) && AUTO_REQUEST_ENABLED) {
      await ensureGrant();
      const retry = await fetch(
        `${API_BASE}/me/drive/items/${itemId}/permissions/${permissionId}`,
        {
          method: "DELETE",
          headers: { "X-Agent-Session": agentSession }
        }
      );
      if (retry.status === 204 || retry.status === 200) {
        console.log(`Permission ${permissionId} removed.`);
        return;
      }
    }
    if (data) {
      describeAuthError(data, "drive");
      process.exit(1);
    }
  }
  throw new Error(`Failed to remove permission (${response.status})`);
}
const MARKDOWN_CACHE_FOLDER = ".markdown";
async function driveConvert(itemId, options) {
  if (!options.force) {
    console.error("Checking OneDrive cache...");
    try {
      const encodedPath = `/${MARKDOWN_CACHE_FOLDER}`.split("/").map(encodeURIComponent).join("/");
      const result = await apiRequest(
        `/me/drive/root:${encodedPath}:/children?$select=id,name&$top=200`
      );
      const cached = result.value?.find((i) => i.name === `${itemId}.md`);
      if (cached) {
        console.error(`Found cached version: ${cached.name}`);
        let response2 = await fetchGraphContent(
          `/me/drive/items/${cached.id}/content`
        );
        if (response2.status === 302 || response2.status === 301) {
          const url = response2.headers.get("location");
          if (url) response2 = await fetch(url);
        }
        const markdown = await response2.text();
        if (options.output) {
          await writeFile(options.output, markdown);
          console.error(`Saved to: ${options.output}`);
        } else {
          console.log(markdown);
        }
        return;
      }
    } catch {
    }
    console.error("No cache found, converting...");
  }
  console.error("Downloading file...");
  const info = await apiRequest(
    `/me/drive/items/${itemId}?$select=id,name,size,file,folder`
  );
  if (info.folder) throw new Error("Cannot convert a folder");
  const tempDir = join(tmpdir(), "office-convert");
  await mkdir(tempDir, { recursive: true });
  const tempPath = join(tempDir, info.name);
  let response = await fetchGraphContent(`/me/drive/items/${itemId}/content`);
  if (response.status === 302 || response.status === 301) {
    const url = response.headers.get("location");
    if (url) response = await fetch(url);
  }
  if (!response.ok) throw new Error(`Download failed (${response.status})`);
  await writeFile(tempPath, Buffer.from(await response.arrayBuffer()));
  console.error(`Downloaded: ${info.name} (${formatBytes(info.size || 0)})`);
  try {
    const ext = extname(info.name).toLowerCase();
    const mimeType = info.file?.mimeType || getMimeType(ext);
    let images;
    let imageMimeType = "image/png";
    if (mimeType === "application/pdf" || ext === ".pdf") {
      console.error("Converting PDF pages to images...");
      images = await convertPdfToImages(tempPath);
      console.error(`Extracted ${images.length} page(s)`);
    } else if (mimeType.startsWith("image/") || [".png", ".jpg", ".jpeg", ".gif", ".webp"].includes(ext)) {
      images = [await imageToBase64(tempPath)];
      imageMimeType = mimeType.startsWith("image/") ? mimeType : getMimeType(ext);
    } else {
      throw new Error(
        `Unsupported file type: ${mimeType}. Supported: PDF, PNG, JPG, GIF, WebP`
      );
    }
    const apiKey = process.env.OPENROUTER_API_KEY;
    if (!apiKey) throw new Error("OPENROUTER_API_KEY environment variable required");
    console.error(`Converting ${images.length} page(s) with Gemini 3 Flash...`);
    const content = [];
    for (const base64 of images) {
      content.push({
        type: "image_url",
        image_url: { url: `data:${imageMimeType};base64,${base64}` }
      });
    }
    content.push({
      type: "text",
      text: `Convert this document "${info.name}" to well-formatted markdown.
Preserve structure: headings, lists, tables, formatting.
For multi-page documents, use horizontal rules (---) between pages.
Extract all text accurately. For tables, use markdown table syntax.
Do not include any commentary or explanations - just the converted content.`
    });
    const geminiResponse = await fetch(
      "https://openrouter.ai/api/v1/chat/completions",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          model: "google/gemini-3-flash-preview",
          messages: [{ role: "user", content }]
        })
      }
    );
    if (!geminiResponse.ok) {
      throw new Error(
        `OpenRouter API error (${geminiResponse.status}): ${await geminiResponse.text()}`
      );
    }
    const result = await geminiResponse.json();
    if (result.error) throw new Error(`OpenRouter error: ${result.error.message}`);
    const markdown = result.choices[0].message.content;
    const cachePath = `/${MARKDOWN_CACHE_FOLDER}/${itemId}.md`;
    const encodedCachePath = cachePath.split("/").map(encodeURIComponent).join("/");
    console.error(`Uploading cached markdown to OneDrive ${cachePath}...`);
    await apiRequest(`/me/drive/root:${encodedCachePath}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "text/plain" },
      body: markdown
    });
    console.error("Cached to OneDrive successfully.");
    if (options.output) {
      await writeFile(options.output, markdown);
      console.error(`Saved to: ${options.output}`);
    } else {
      console.log(markdown);
    }
  } finally {
    try {
      await unlink(tempPath);
    } catch {
    }
  }
}
function getMimeType(ext) {
  const mimeTypes = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".pdf": "application/pdf"
  };
  return mimeTypes[ext.toLowerCase()] || "application/octet-stream";
}
async function convertPdfToImages(pdfPath) {
  const outputDir = join(dirname(pdfPath), "pages");
  await mkdir(outputDir, { recursive: true });
  const outputPrefix = join(outputDir, "page");
  return new Promise((resolve, reject) => {
    const proc = spawn("pdftoppm", [
      "-png",
      "-r",
      "150",
      pdfPath,
      outputPrefix
    ]);
    let stderr = "";
    proc.stderr.on("data", (d) => stderr += d.toString());
    proc.on("close", async (code) => {
      if (code !== 0) {
        reject(new Error(`pdftoppm failed: ${stderr}`));
        return;
      }
      try {
        const files = await readdir(outputDir);
        const pngFiles = files.filter((f) => f.endsWith(".png")).sort();
        const images = [];
        for (const file of pngFiles) {
          const imgPath = join(outputDir, file);
          const imgBuffer = await readFile(imgPath);
          images.push(imgBuffer.toString("base64"));
          await unlink(imgPath);
        }
        resolve(images);
      } catch (err) {
        reject(err);
      }
    });
  });
}
async function imageToBase64(imagePath) {
  return (await readFile(imagePath)).toString("base64");
}
function formatEmailAddress(addr) {
  if (!addr) return "";
  const name = addr.name?.trim();
  const email = addr.address?.trim();
  if (name && email) return `${name} <${email}>`;
  return name || email || "";
}
async function mailInbox(count = 20, unreadOnly = false, folder = "inbox") {
  let filter = "";
  if (unreadOnly) filter = "&$filter=isRead eq false";
  const folderPath = folder === "inbox" ? `/me/messages` : `/me/mailFolders/${folder}/messages`;
  const result = await apiRequest(
    `${folderPath}?$select=id,subject,from,receivedDateTime,isRead,hasAttachments,importance,bodyPreview&$orderby=receivedDateTime desc&$top=${count}${filter}`
  );
  if (!result.value || result.value.length === 0) {
    console.log(unreadOnly ? "No unread messages." : "No messages found.");
    return;
  }
  const folderLabel = folder === "inbox" ? unreadOnly ? "Unread" : "Recent" : folder.charAt(0).toUpperCase() + folder.slice(1);
  console.log(
    `
${folderLabel} messages (${result.value.length}):
`
  );
  console.log("\u2500".repeat(95));
  console.log(
    `${"".padEnd(2)} ${"From".padEnd(25)} ${"Subject".padEnd(40)} ${"Date".padEnd(18)} ID`
  );
  console.log("\u2500".repeat(95));
  for (const msg of result.value) {
    const read = msg.isRead ? " " : "\u25CF";
    const from = truncate(
      formatEmailAddress(msg.from?.emailAddress),
      24
    );
    const subj = truncate(msg.subject || "(No subject)", 39);
    const date = new Date(msg.receivedDateTime).toLocaleDateString("en-AU", {
      day: "2-digit",
      month: "short",
      hour: "2-digit",
      minute: "2-digit"
    });
    const attach = msg.hasAttachments ? "\u{1F4CE}" : "";
    console.log(
      `${read} ${from.padEnd(25)} ${subj.padEnd(40)} ${date.padEnd(18)} ${msg.id.substring(0, 20)}\u2026`
    );
    if (attach) console.log(`  ${attach}`);
  }
  console.log(
    `
Use 'office-cli.js mail read <id>' to read a message.`
  );
}
async function mailSearch(query, count = 15, includeJunk = false) {
  const folderPath = includeJunk ? `/me/messages` : `/me/mailFolders/inbox/messages`;
  const result = await apiRequest(
    `${folderPath}?$search="${encodeURIComponent(query)}"&$select=id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview&$top=${count}`
  );
  if (!result.value || result.value.length === 0) {
    console.log(`No messages found matching "${query}".`);
    return;
  }
  console.log(`
Search results for "${query}" (${result.value.length}):
`);
  for (const msg of result.value) {
    const from = formatEmailAddress(msg.from?.emailAddress);
    const date = new Date(msg.receivedDateTime).toLocaleString("en-AU");
    const read = msg.isRead ? "" : " [UNREAD]";
    console.log(`${msg.isRead ? "\u{1F4E7}" : "\u{1F4EC}"} ${msg.subject || "(No subject)"}${read}`);
    console.log(`   From: ${from}`);
    console.log(`   Date: ${date}`);
    console.log(`   ID: ${msg.id}`);
    if (msg.bodyPreview) {
      console.log(`   Preview: ${truncate(msg.bodyPreview, 100)}`);
    }
    console.log();
  }
}
async function mailRead(messageId) {
  const msg = await apiRequest(
    `/me/messages/${messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,importance,body,webLink`,
    {
      headers: { Prefer: 'outlook.body-content-type="text"' }
    }
  );
  console.log("\n" + "\u2550".repeat(70));
  console.log(`Subject: ${msg.subject || "(No subject)"}`);
  console.log("\u2500".repeat(70));
  console.log(`From: ${formatEmailAddress(msg.from?.emailAddress)}`);
  if (msg.toRecipients?.length) {
    const to = msg.toRecipients.map((r) => formatEmailAddress(r.emailAddress)).join(", ");
    console.log(`To: ${to}`);
  }
  if (msg.ccRecipients?.length) {
    const cc = msg.ccRecipients.map((r) => formatEmailAddress(r.emailAddress)).join(", ");
    console.log(`CC: ${cc}`);
  }
  console.log(
    `Date: ${new Date(msg.receivedDateTime).toLocaleString("en-AU")}`
  );
  console.log(`Importance: ${msg.importance}`);
  if (msg.hasAttachments) console.log("Attachments: Yes \u{1F4CE}");
  console.log(`ID: ${msg.id}`);
  if (msg.webLink) console.log(`Web: ${msg.webLink}`);
  console.log("\u2500".repeat(70));
  if (msg.body?.content) {
    const bodyText = msg.body.contentType?.toLowerCase() === "html" ? stripHtml(msg.body.content) : msg.body.content;
    console.log(bodyText);
  } else {
    console.log("(No body content)");
  }
  console.log("\u2550".repeat(70));
}
async function peopleSearch(query, top = 10) {
  const result = await apiRequest(
    `/me/people?$search="${encodeURIComponent(query)}"&$select=displayName,givenName,surname,scoredEmailAddresses,personType,jobTitle,department,userPrincipalName&$top=${top}`,
    {
      headers: {
        "X-PeopleQuery-QuerySources": "Mailbox,Directory"
      }
    }
  );
  return result.value || [];
}
async function peopleList(query, verbose = false) {
  const people = await peopleSearch(query);
  if (people.length === 0) {
    console.log(`No results for "${query}".`);
    return;
  }
  console.log(`
People matching "${query}":
`);
  for (let i = 0; i < people.length; i++) {
    const p = people[i];
    const email = p.scoredEmailAddresses?.[0]?.address || p.userPrincipalName || "no email";
    const score = p.scoredEmailAddresses?.[0]?.relevanceScore;
    const type = p.personType?.subclass || "unknown";
    const typeLabel = type === "OrganizationUser" ? "org" : type === "PersonalContact" ? "contact" : type === "ImplicitContact" ? "recent" : type;
    console.log(`  ${i + 1}. ${p.displayName} <${email}>  (${typeLabel}${score ? `, relevance: ${score}` : ""})`);
    if (verbose) {
      if (p.jobTitle) console.log(`     Title: ${p.jobTitle}`);
      if (p.department) console.log(`     Dept: ${p.department}`);
      if (p.scoredEmailAddresses && p.scoredEmailAddresses.length > 1) {
        for (let j = 1; j < p.scoredEmailAddresses.length; j++) {
          console.log(`     Also: ${p.scoredEmailAddresses[j].address}`);
        }
      }
    }
  }
  console.log();
}
async function resolveRecipient(nameOrEmail) {
  if (validateEmail(nameOrEmail)) {
    return [{ address: nameOrEmail, displayName: nameOrEmail }];
  }
  const people = await peopleSearch(nameOrEmail, 5);
  return people.filter((p) => p.scoredEmailAddresses?.[0]?.address).map((p) => ({
    address: p.scoredEmailAddresses[0].address,
    displayName: p.displayName
  }));
}
async function resolveRecipients(input, autoConfirm = false) {
  const parts = input.split(",").map((s) => s.trim());
  const resolved = [];
  for (const part of parts) {
    if (validateEmail(part)) {
      resolved.push({ emailAddress: { address: part } });
      continue;
    }
    const matches = await resolveRecipient(part);
    if (matches.length === 0) {
      console.error(`No matches found for "${part}". Please provide a full email address.`);
      process.exit(1);
    }
    console.log(`
  Matches for "${part}":`);
    const shown = matches.slice(0, 5);
    for (let i = 0; i < shown.length; i++) {
      console.log(`    ${i + 1}. ${shown[i].displayName} <${shown[i].address}>`);
    }
    if (autoConfirm) {
      console.log(`  \u2192 Using: ${shown[0].displayName} <${shown[0].address}>`);
      resolved.push({ emailAddress: { address: shown[0].address } });
    } else {
      console.log(`
  Re-run with the exact email, or use --yes to auto-pick the top result.`);
      process.exit(1);
    }
  }
  return resolved;
}
function validateEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
}
function parseRecipients(csv) {
  return csv.split(",").map((email) => {
    const trimmed = email.trim();
    if (!validateEmail(trimmed)) {
      throw new Error(`Invalid email address: ${trimmed}`);
    }
    return { emailAddress: { address: trimmed } };
  });
}
async function mailSend(options) {
  const toRecipients = await resolveRecipients(options.to, options.autoConfirm);
  const payload = {
    message: {
      subject: options.subject,
      body: {
        contentType: options.isHtml ? "HTML" : "Text",
        content: options.body
      },
      toRecipients
    },
    saveToSentItems: true
  };
  if (options.cc) {
    payload.message.ccRecipients = await resolveRecipients(options.cc, options.autoConfirm);
  }
  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run './office-cli.js login' first.");
    process.exit(1);
  }
  const response = await fetch(`${API_BASE}/me/sendMail`, {
    method: "POST",
    headers: {
      "X-Agent-Session": agentSession,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });
  if (response.status === 202 || response.status === 200) {
    console.log(`Email sent to ${options.to}`);
    return;
  }
  if (response.status === 401 || response.status === 403) {
    const { data, text } = await readErrorBody(response);
    if (isAutoRequestError(data) && AUTO_REQUEST_ENABLED) {
      await ensureGrant();
      const retry = await fetch(`${API_BASE}/me/sendMail`, {
        method: "POST",
        headers: {
          "X-Agent-Session": agentSession,
          "Content-Type": "application/json"
        },
        body: JSON.stringify(payload)
      });
      if (retry.status === 202 || retry.status === 200) {
        console.log(`Email sent to ${options.to}`);
        return;
      }
      throw new Error(
        `Send failed after grant (${retry.status}): ${await retry.text()}`
      );
    }
    if (data) {
      describeAuthError(data, "mail");
      process.exit(1);
    }
    throw new Error(text || `Auth error (${response.status})`);
  }
  const { text: errText } = await readErrorBody(response);
  throw new Error(
    `Send failed (${response.status}): ${errText || response.statusText}`
  );
}
async function mailReply(messageId, comment, replyAll = false) {
  const endpoint = replyAll ? `/me/messages/${messageId}/replyAll` : `/me/messages/${messageId}/reply`;
  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run './office-cli.js login' first.");
    process.exit(1);
  }
  const response = await fetch(`${API_BASE}${endpoint}`, {
    method: "POST",
    headers: {
      "X-Agent-Session": agentSession,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ comment })
  });
  if (response.status === 202 || response.status === 200) {
    console.log(replyAll ? "Reply-all sent." : "Reply sent.");
    return;
  }
  if (response.status === 401 || response.status === 403) {
    const { data, text } = await readErrorBody(response);
    if (isAutoRequestError(data) && AUTO_REQUEST_ENABLED) {
      await ensureGrant();
      const retry = await fetch(`${API_BASE}${endpoint}`, {
        method: "POST",
        headers: {
          "X-Agent-Session": agentSession,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({ comment })
      });
      if (retry.status === 202 || retry.status === 200) {
        console.log(replyAll ? "Reply-all sent." : "Reply sent.");
        return;
      }
      throw new Error(
        `Reply failed after grant (${retry.status}): ${await retry.text()}`
      );
    }
    if (data) {
      describeAuthError(data, "mail");
      process.exit(1);
    }
    throw new Error(text || `Auth error (${response.status})`);
  }
  const { text: replyErrText } = await readErrorBody(response);
  throw new Error(
    `Reply failed (${response.status}): ${replyErrText || response.statusText}`
  );
}
function parseEventDate(dateStr, timeZone) {
  if (timeZone === "UTC" && !dateStr.endsWith("Z")) {
    return /* @__PURE__ */ new Date(dateStr + "Z");
  }
  return new Date(dateStr);
}
function formatTime(dateStr, isAllDay, timeZone) {
  if (isAllDay) return "All day";
  const date = parseEventDate(dateStr, timeZone);
  return date.toLocaleTimeString("en-AU", {
    hour: "numeric",
    minute: "2-digit",
    hour12: true
  });
}
function formatDateHeader(date) {
  const today = /* @__PURE__ */ new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  const dateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const todayOnly = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate()
  );
  const tomorrowOnly = new Date(
    tomorrow.getFullYear(),
    tomorrow.getMonth(),
    tomorrow.getDate()
  );
  const formatted = date.toLocaleDateString("en-AU", {
    weekday: "long",
    month: "long",
    day: "numeric"
  });
  if (dateOnly.getTime() === todayOnly.getTime())
    return `Today - ${formatted}`;
  if (dateOnly.getTime() === tomorrowOnly.getTime())
    return `Tomorrow - ${formatted}`;
  return formatted;
}
function getEventDateKey(event) {
  const startStr = event.start.dateTime || event.start.date || "";
  const date = parseEventDate(startStr, event.start.timeZone);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}
async function calEvents(days = 7, includeDetails = false) {
  const startDate = /* @__PURE__ */ new Date();
  startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(startDate);
  endDate.setDate(endDate.getDate() + days);
  console.error(
    `Fetching calendar events for the next ${days} day${days !== 1 ? "s" : ""}...`
  );
  const selectFields = includeDetails ? "subject,start,end,location,isAllDay,organizer,attendees,body" : "subject,start,end,location,isAllDay,organizer";
  const headers = {};
  if (includeDetails)
    headers["Prefer"] = 'outlook.body-content-type="text"';
  const result = await apiRequest(
    `/me/calendarView?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}&$select=${selectFields}&$orderby=start/dateTime&$top=50`,
    { headers }
  );
  const events = result.value || [];
  if (events.length === 0) {
    console.log("No events scheduled for this period.");
    return;
  }
  const eventsByDate = {};
  for (const event of events) {
    const key = getEventDateKey(event);
    if (!eventsByDate[key]) eventsByDate[key] = [];
    eventsByDate[key].push(event);
  }
  console.log("");
  for (const dateKey of Object.keys(eventsByDate).sort()) {
    const [year, month, day] = dateKey.split("-").map(Number);
    const date = new Date(year, month - 1, day);
    console.log(`## ${formatDateHeader(date)}
`);
    for (const event of eventsByDate[dateKey]) {
      const startTime = formatTime(
        event.start.dateTime || event.start.date || "",
        event.isAllDay,
        event.start.timeZone
      );
      const endTime = formatTime(
        event.end.dateTime || event.end.date || "",
        event.isAllDay,
        event.end.timeZone
      );
      const timeRange = event.isAllDay ? "All day" : `${startTime} - ${endTime}`;
      console.log(`  **${event.subject || "(No title)"}**`);
      console.log(`  ${timeRange}`);
      if (event.location?.displayName)
        console.log(`  Location: ${event.location.displayName}`);
      if (includeDetails) {
        const organizer = formatEmailAddress(event.organizer?.emailAddress);
        if (organizer) console.log(`  Organizer: ${organizer}`);
        const attendees = (event.attendees || []).map((a) => {
          const label = formatEmailAddress(a.emailAddress);
          return a.type ? `${label} (${a.type.toLowerCase()})` : label;
        }).filter(Boolean);
        if (attendees.length > 0) {
          console.log("  Participants:");
          for (const a of attendees) console.log(`    - ${a}`);
        }
        if (event.body?.content) {
          const bodyText = event.body.contentType?.toLowerCase() === "html" ? stripHtml(event.body.content) : event.body.content.trim();
          if (bodyText) {
            console.log("  Body:");
            for (const line of bodyText.split("\n"))
              console.log(`    ${line}`);
          }
        }
      }
      console.log("");
    }
  }
  console.log(
    `${events.length} event${events.length !== 1 ? "s" : ""} over ${days} day${days !== 1 ? "s" : ""}`
  );
}
async function calTomorrow(includeDetails = false) {
  const tomorrow = /* @__PURE__ */ new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  const dayAfter = new Date(tomorrow);
  dayAfter.setDate(dayAfter.getDate() + 1);
  console.error("Fetching tomorrow's events...");
  const selectFields = includeDetails ? "subject,start,end,location,isAllDay,organizer,attendees,body" : "subject,start,end,location,isAllDay,organizer";
  const headers = {};
  if (includeDetails)
    headers["Prefer"] = 'outlook.body-content-type="text"';
  const result = await apiRequest(
    `/me/calendarView?startDateTime=${tomorrow.toISOString()}&endDateTime=${dayAfter.toISOString()}&$select=${selectFields}&$orderby=start/dateTime&$top=50`,
    { headers }
  );
  const events = result.value || [];
  if (events.length === 0) {
    console.log("No events scheduled for tomorrow.");
    return;
  }
  console.log(`
## ${formatDateHeader(tomorrow)}
`);
  for (const event of events) {
    const startTime = formatTime(
      event.start.dateTime || event.start.date || "",
      event.isAllDay,
      event.start.timeZone
    );
    const endTime = formatTime(
      event.end.dateTime || event.end.date || "",
      event.isAllDay,
      event.end.timeZone
    );
    const timeRange = event.isAllDay ? "All day" : `${startTime} - ${endTime}`;
    console.log(`  **${event.subject || "(No title)"}**`);
    console.log(`  ${timeRange}`);
    if (event.location?.displayName)
      console.log(`  Location: ${event.location.displayName}`);
    if (includeDetails) {
      const organizer = formatEmailAddress(event.organizer?.emailAddress);
      if (organizer) console.log(`  Organizer: ${organizer}`);
      const attendees = (event.attendees || []).map((a) => {
        const label = formatEmailAddress(a.emailAddress);
        return a.type ? `${label} (${a.type.toLowerCase()})` : label;
      }).filter(Boolean);
      if (attendees.length > 0) {
        console.log("  Participants:");
        for (const a of attendees) console.log(`    - ${a}`);
      }
    }
    console.log("");
  }
  console.log(
    `${events.length} event${events.length !== 1 ? "s" : ""}`
  );
}
async function checkStatus(service) {
  const session = await getAgentSession();
  console.log("Office CLI Status\n");
  console.log(`Auth Service: ${AUTH_SERVICE_URL}`);
  console.log(`Primary session file: ${SHARED_AGENT_SESSION_FILE}`);
  if (LEGACY_AGENT_SESSION_FILES.length > 0) {
    console.log(
      `Fallback session files: ${LEGACY_AGENT_SESSION_FILES.join(", ")}`
    );
  }
  console.log(`Agent session saved: ${session ? "Yes" : "No"}`);
  if (!session) {
    console.log("\n\u274C Not registered");
    console.log("");
    console.log("To register:");
    console.log("1. Run: ./office-cli.js login");
    console.log("2. Follow the approval link shown");
    return;
  }
  console.log("\n\u2705 Agent session saved");
  try {
    const response = await fetch(`${STATUS_BASE}/msgraph`, {
      headers: { "X-Agent-Session": session }
    });
    if (!response.ok) {
      const data = await response.json();
      describeAuthError(data, "msgraph");
      return;
    }
    const status = await response.json();
    if (!status.hasGrant) {
      console.log("\n\u26A0\uFE0F  No active grant for Microsoft Graph");
      console.log("Run: ./office-cli.js request <drive|mail|cal>");
      return;
    }
    console.log("\n\u2705 Grant active for Microsoft Graph");
    if (status.scopes?.length) {
      console.log(`   Scopes: ${status.scopes.join(" ")}`);
      const activeScopes = new Set(status.scopes);
      const services = [
        { name: "Drive", scopes: SCOPES.drive },
        { name: "Mail", scopes: SCOPES.mail },
        { name: "Calendar", scopes: SCOPES.cal }
      ];
      for (const svc of services) {
        const covered = svc.scopes.every((s) => activeScopes.has(s));
        console.log(`   ${covered ? "\u2705" : "\u26A0\uFE0F "} ${svc.name}: ${svc.scopes.join(", ")}`);
      }
    }
    if (status.expiresAt) {
      console.log(
        `   Expires at: ${new Date(status.expiresAt).toLocaleString()}`
      );
    }
  } catch {
  }
}
async function doLogin(agentName) {
  const name = agentName || "Fresh Auth CLI";
  const deviceInfo = `${hostname()} (${process.platform})`;
  const response = await fetch(`${AUTH_SERVICE_URL}/api/agent/init`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ agentName: name, deviceInfo })
  });
  if (!response.ok) {
    throw new Error(
      `Failed to start agent registration: ${await response.text()}`
    );
  }
  const data = await response.json();
  if (data.agentSessionId?.trim()) {
    await saveAgentSession(data.agentSessionId);
    console.log("Agent session saved (pending approval).");
  }
  console.log("Agent registration started.");
  console.log(`Code: ${data.code}`);
  const verifyUrl = data.verifyUrl || buildRegistrationUrl(data.code);
  console.log(`Verify at: ${verifyUrl}`);
  console.log("");
  console.log("Waiting for approval...");
  const pollStart = Date.now();
  while (true) {
    const poll = await fetch(data.pollUrl);
    if (!poll.ok) throw new Error(`Registration poll failed: ${await poll.text()}`);
    const status = await poll.json();
    if (status.agentSessionId?.trim()) {
      await saveAgentSession(status.agentSessionId);
    }
    if (status.status === "approved" && status.agentSessionId) {
      console.log("Agent session saved.");
      console.log("Next: run './office-cli.js request <drive|mail|cal>' to request access.");
      break;
    }
    if (status.status === "denied")
      throw new Error(status.message || "Registration denied.");
    if (status.status === "expired")
      throw new Error("Registration expired. Run login again.");
    if (Date.now() - pollStart > 10 * 60 * 1e3)
      throw new Error("Timed out waiting for approval.");
    await new Promise((resolve) => setTimeout(resolve, 2e3));
  }
}
function printUsage() {
  console.log(`
Office CLI - Unified Microsoft 365 CLI via Auth Service Proxy

Uses agent sessions + grants. Tokens are managed by the auth
service for better security isolation.

Usage:
  office-cli.js <command> [options]
  office-cli.js <service> <command> [options]

Authentication:
  login                        Register agent session
  logout                       Clear saved agent session
  request <drive|mail|cal>     Request access grant for a service
  status                       Check session and grant status

Drive (OneDrive):
  drive list [folder-path]     List files in root or folder
  drive search <query>         Search for files
  drive download <id> [output] Download a file by ID
  drive content <id>           Get text content of a file
  drive info <id>              Get file metadata
  drive convert <id> [opts]    Convert PDF/image to markdown
  drive share <id>             Create shareable link (view, org-wide)
  drive share <id> --email <e> Share with specific person
  drive permissions <id>       List sharing permissions
  drive unshare <id> <perm-id> Remove a sharing permission

Mail (Outlook):
  mail inbox [--count N]       Show recent messages (default: 20)
  mail unread                  Show unread messages
  mail search <query>          Search messages
  mail read <id>               Read a message
  mail send --to <email> --subject <subj> --body <text>
  mail send --to <email> --subject <subj> --file <path>
  mail reply <id> --body <text>
  mail reply-all <id> --body <text>

Calendar:
  cal events [--days N]        Show upcoming events (default: 7 days)
  cal events --full            Include attendees and body
  cal today [--full]           Today's events
  cal tomorrow [--full]        Tomorrow's events

Drive Convert Options:
  --force                      Force re-conversion (ignore cache)
  --output=<path>              Save to local file

First-time Setup:
  1. ./office-cli.js login
  2. Approve agent registration in browser
  3. ./office-cli.js request drive   (for OneDrive)
     ./office-cli.js request mail    (for Outlook email)
     ./office-cli.js request cal     (for Calendar)

Environment:
  AUTH_SERVICE_URL              Auth service URL (default: https://auth.freshhub.ai)
  OPENROUTER_API_KEY            Required for drive convert command
  OFFICE_AUTO_REQUEST=0         Disable auto-request on missing grants
`);
}
function parseNamedArg(args, flag) {
  for (const arg of args) {
    if (arg.startsWith(`${flag}=`)) return arg.substring(flag.length + 1);
  }
  const idx = args.indexOf(flag);
  return idx !== -1 && args[idx + 1] ? args[idx + 1] : void 0;
}
async function main() {
  const args = process.argv.slice(2);
  if (args.length === 0 || args[0] === "help" || args[0] === "--help" || args[0] === "-h") {
    printUsage();
    process.exit(0);
  }
  const command = args[0];
  try {
    switch (command) {
      case "login":
      case "register":
        await doLogin(args[1]);
        return;
      case "logout":
      case "unregister":
        await clearAgentSession();
        console.log("Agent session cleared.");
        return;
      case "status":
        await checkStatus(args[1]);
        return;
      case "request": {
        const service = args[1];
        if (!service || !SCOPES[service]) {
          console.error("Usage: office-cli.js request <drive|mail|cal>");
          console.error("Services: drive, mail, cal");
          process.exit(1);
        }
        const data = await requestGrant(service);
        if (data.autoApproved) {
          console.log("Access granted automatically (within policy).");
        } else {
          console.log("Authorisation request created.");
          console.log(`Approve at: ${data.approveUrl}`);
          console.log("Once approved, try your command again.");
        }
        return;
      }
    }
    const subCommand = args[1];
    switch (command) {
      // ── Drive ──────────────────────────────────────────────
      case "drive": {
        currentService = "drive";
        if (!subCommand) {
          console.error("Usage: office-cli.js drive <list|search|download|content|info|convert>");
          process.exit(1);
        }
        switch (subCommand) {
          case "list":
            await driveList(args[2]);
            break;
          case "search":
            if (!args[2]) {
              console.error("Usage: office-cli.js drive search <query>");
              process.exit(1);
            }
            await driveSearch(args[2]);
            break;
          case "download":
            if (!args[2]) {
              console.error("Usage: office-cli.js drive download <id> [output]");
              process.exit(1);
            }
            await driveDownload(args[2], args[3]);
            break;
          case "content":
            if (!args[2]) {
              console.error("Usage: office-cli.js drive content <id>");
              process.exit(1);
            }
            await driveContent(args[2]);
            break;
          case "info":
            if (!args[2]) {
              console.error("Usage: office-cli.js drive info <id>");
              process.exit(1);
            }
            await driveInfo(args[2]);
            break;
          case "convert": {
            if (!args[2]) {
              console.error("Usage: office-cli.js drive convert <id> [--force] [--output=<path>]");
              process.exit(1);
            }
            const convertOpts = {};
            for (let i = 3; i < args.length; i++) {
              if (args[i] === "--force") convertOpts.force = true;
              else if (args[i].startsWith("--output="))
                convertOpts.output = args[i].substring("--output=".length);
            }
            await driveConvert(args[2], convertOpts);
            break;
          }
          case "share": {
            if (!args[2]) {
              console.error("Usage: office-cli.js drive share <id> [--anyone] [--expires <date>] [--email <email>] [--role read|write] [--type view|edit] [--scope organization|anonymous]");
              process.exit(1);
            }
            const shareOpts = {};
            shareOpts.anyone = args.slice(2).includes("--anyone");
            shareOpts.expires = parseNamedArg(args.slice(2), "--expires");
            shareOpts.email = parseNamedArg(args.slice(2), "--email");
            shareOpts.role = parseNamedArg(args.slice(2), "--role");
            shareOpts.type = parseNamedArg(args.slice(2), "--type");
            shareOpts.scope = parseNamedArg(args.slice(2), "--scope");
            await driveShare(args[2], shareOpts);
            break;
          }
          case "permissions": {
            if (!args[2]) {
              console.error("Usage: office-cli.js drive permissions <id>");
              process.exit(1);
            }
            await drivePermissions(args[2]);
            break;
          }
          case "unshare": {
            if (!args[2] || !args[3]) {
              console.error("Usage: office-cli.js drive unshare <item-id> <permission-id>");
              process.exit(1);
            }
            await driveUnshare(args[2], args[3]);
            break;
          }
          default:
            console.error(`Unknown drive command: ${subCommand}`);
            process.exit(1);
        }
        break;
      }
      // ── Mail ───────────────────────────────────────────────
      case "mail": {
        currentService = "mail";
        if (!subCommand) {
          console.error("Usage: office-cli.js mail <inbox|unread|junk|search|read|send|reply|reply-all>");
          process.exit(1);
        }
        switch (subCommand) {
          case "inbox": {
            const count = parseInt(parseNamedArg(args.slice(2), "--count") || "20", 10);
            await mailInbox(count, false, "inbox");
            break;
          }
          case "unread":
            await mailInbox(20, true, "inbox");
            break;
          case "junk": {
            const count = parseInt(parseNamedArg(args.slice(2), "--count") || "20", 10);
            await mailInbox(count, false, "junkemail");
            break;
          }
          case "search":
            if (!args[2]) {
              console.error("Usage: office-cli.js mail search <query> [--include-junk]");
              process.exit(1);
            }
            await mailSearch(args[2], 15, args.slice(3).includes("--include-junk"));
            break;
          case "read":
            if (!args[2]) {
              console.error("Usage: office-cli.js mail read <message-id>");
              process.exit(1);
            }
            await mailRead(args[2]);
            break;
          case "send": {
            const to = parseNamedArg(args.slice(2), "--to");
            const subject = parseNamedArg(args.slice(2), "--subject");
            const bodyText = parseNamedArg(args.slice(2), "--body");
            const file = parseNamedArg(args.slice(2), "--file");
            const cc = parseNamedArg(args.slice(2), "--cc");
            const htmlFlag = args.slice(2).includes("--html");
            const autoConfirm = args.slice(2).includes("--yes") || args.slice(2).includes("-y");
            if (!to || !subject) {
              console.error(
                "Usage: office-cli.js mail send --to <email> --subject <subject> --body <text>"
              );
              console.error(
                "   or: office-cli.js mail send --to <email> --subject <subject> --body <html> --html"
              );
              console.error(
                "   or: office-cli.js mail send --to <email> --subject <subject> --file <path>"
              );
              process.exit(1);
            }
            let body;
            let isHtml = htmlFlag;
            if (file) {
              if (!existsSync(file)) {
                console.error(`File not found: ${file}`);
                process.exit(1);
              }
              body = await readFile(file, "utf-8");
              if (!htmlFlag) isHtml = file.endsWith(".html") || file.endsWith(".htm");
            } else if (bodyText) {
              body = bodyText;
            } else {
              const chunks = [];
              for await (const chunk of Bun.stdin.stream()) {
                chunks.push(Buffer.from(chunk));
              }
              body = Buffer.concat(chunks).toString("utf-8");
            }
            if (!body.trim()) {
              console.error("No email body provided.");
              process.exit(1);
            }
            await mailSend({ to, subject, body, cc, isHtml, autoConfirm });
            break;
          }
          case "reply":
          case "reply-all": {
            if (!args[2]) {
              console.error(
                `Usage: office-cli.js mail ${subCommand} <message-id> --body <text>`
              );
              process.exit(1);
            }
            const replyBody = parseNamedArg(args.slice(2), "--body");
            if (!replyBody) {
              console.error("--body is required for reply");
              process.exit(1);
            }
            await mailReply(args[2], replyBody, subCommand === "reply-all");
            break;
          }
          default:
            console.error(`Unknown mail command: ${subCommand}`);
            process.exit(1);
        }
        break;
      }
      // ── Calendar ───────────────────────────────────────────
      case "cal": {
        currentService = "cal";
        if (!subCommand) {
          console.error("Usage: office-cli.js cal <events|today|tomorrow>");
          process.exit(1);
        }
        switch (subCommand) {
          case "events": {
            const daysStr = parseNamedArg(args.slice(2), "--days");
            const days = daysStr ? parseInt(daysStr, 10) || 7 : 7;
            const full = args.includes("--full") || args.includes("--details");
            await calEvents(days, full);
            break;
          }
          case "today": {
            const full = args.includes("--full") || args.includes("--details");
            await calEvents(1, full);
            break;
          }
          case "tomorrow": {
            const full = args.includes("--full") || args.includes("--details");
            await calTomorrow(full);
            break;
          }
          default:
            console.error(`Unknown cal command: ${subCommand}`);
            process.exit(1);
        }
        break;
      }
      case "people": {
        const query = args.slice(1).filter((a) => !a.startsWith("--")).join(" ");
        if (!query) {
          console.error("Usage: office-cli.js people <name>");
          process.exit(1);
        }
        const verbose = args.includes("--verbose") || args.includes("-v");
        await peopleList(query, verbose);
        break;
      }
      default:
        console.error(`Unknown command: ${command}`);
        printUsage();
        process.exit(1);
    }
  } catch (error) {
    console.error(
      `Error: ${error instanceof Error ? error.message : error}`
    );
    process.exit(1);
  }
}
main();
