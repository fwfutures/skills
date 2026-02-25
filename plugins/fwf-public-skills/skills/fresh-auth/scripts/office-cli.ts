#!/usr/bin/env bun
/**
 * Office CLI - Unified Microsoft 365 CLI via Auth Service Proxy
 *
 * Subcommands:
 *   drive   - OneDrive file operations (list, search, download, content, convert)
 *   mail    - Outlook email (inbox, search, read, send, reply)
 *   cal     - Calendar events (events, today, tomorrow)
 *
 * Uses agent session authentication - never sees raw OAuth tokens.
 * All token management is handled by the auth service.
 */

import { mkdir, writeFile, readFile, unlink, readdir } from "fs/promises";
import { dirname, join, basename, extname } from "path";
import { existsSync } from "fs";
import { homedir, tmpdir, hostname } from "os";
import { spawn } from "child_process";

// =============================================================================
// Configuration
// =============================================================================

const RAW_AUTH_SERVICE_URL =
  process.env.AUTH_SERVICE_URL || "https://auth.freshhub.ai";
const AUTH_SERVICE_URL = RAW_AUTH_SERVICE_URL.replace(/\/+$/, "").replace(
  /\/api$/,
  ""
);
const API_BASE = `${AUTH_SERVICE_URL}/proxy/msgraph`;
const STATUS_BASE = `${AUTH_SERVICE_URL}/api/proxy/status`;
const AGENT_SESSION_FILE = join(
  homedir(),
  ".config",
  "office-cli",
  "agent-session"
);
const AUTO_REQUEST_POLL_MS = 2000;
const AUTO_REQUEST_ENABLED = process.env.OFFICE_AUTO_REQUEST !== "0";

// Scope sets per service
const SCOPES = {
  drive: ["Files.Read", "Files.ReadWrite.All"],
  mail: ["Mail.Read", "Mail.Send", "People.Read"],
  cal: ["Calendars.Read"],
} as const;

type ServiceName = keyof typeof SCOPES;

const GRANT_DURATION: Record<ServiceName, string> = {
  drive: "30m",
  mail: "1h",
  cal: "1h",
};

// =============================================================================
// Shared Types
// =============================================================================

interface ApiResponse {
  error?: string;
  message?: string;
  requestUrl?: string;
  connectUrl?: string;
  reauthorizeUrl?: string;
  missingScopes?: string[];
  elevateUrl?: string;
  status?: string;
  value?: any[];
}

interface GrantStatusResponse {
  service: string;
  hasGrant: boolean;
  scopes?: string[];
  expiresAt?: string;
  remainingUses?: number;
}

interface AuthRequestResponse {
  requestId: string;
  status: string;
  pollUrl: string;
  approveUrl: string;
  expiresAt: string;
  expiresIn?: number;
  autoApproved?: boolean;
}

// =============================================================================
// Session Management
// =============================================================================

async function getAgentSession(): Promise<string | null> {
  try {
    if (!existsSync(AGENT_SESSION_FILE)) return null;
    const sessionRaw = (await readFile(AGENT_SESSION_FILE, "utf-8")).trim();
    if (!sessionRaw) return null;

    try {
      const parsed = JSON.parse(sessionRaw) as {
        agentSessionId?: string;
        agentSession?: string;
        session?: string;
      };
      const normalized =
        parsed.agentSessionId || parsed.agentSession || parsed.session;
      return normalized?.trim() || null;
    } catch {
      return sessionRaw;
    }
  } catch {
    return null;
  }
}

async function saveAgentSession(agentSessionId: string): Promise<void> {
  const dir = dirname(AGENT_SESSION_FILE);
  await mkdir(dir, { recursive: true });
  await writeFile(AGENT_SESSION_FILE, `${agentSessionId.trim()}\n`, {
    mode: 0o600,
  });
}

async function clearAgentSession(): Promise<void> {
  try {
    if (existsSync(AGENT_SESSION_FILE)) {
      await unlink(AGENT_SESSION_FILE);
    }
  } catch {
    // Ignore
  }
}

// =============================================================================
// Auth Helpers
// =============================================================================

function buildRegistrationUrl(code: string): string {
  return `${AUTH_SERVICE_URL}/agent/verify?code=${encodeURIComponent(code)}`;
}

function describeAuthError(data: ApiResponse, service: string): void {
  if (data.error === "no_agent_session") {
    console.error("No agent session. Run 'office-cli.ts login' to register.");
    return;
  }

  if (
    ["no_grant", "grant_expired", "single_use_exhausted"].includes(
      data.error || ""
    )
  ) {
    console.error(`No active grant for ${service}.`);
    console.error("");
    console.error("Request access by running:");
    console.error(`  ./office-cli.ts request ${service}`);
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

function isAutoRequestError(data: ApiResponse | null): boolean {
  if (!data?.error) return false;
  return ["no_grant", "grant_expired", "single_use_exhausted"].includes(
    data.error
  );
}

async function readErrorBody<T>(
  response: Response
): Promise<{ data: T | null; text: string }> {
  const text = await response.text();
  if (!text) return { data: null, text: "" };
  try {
    return { data: JSON.parse(text) as T, text };
  } catch {
    return { data: null, text };
  }
}

async function requestGrant(
  service: ServiceName
): Promise<AuthRequestResponse> {
  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run './office-cli.ts login' first.");
    process.exit(1);
  }

  const response = await fetch(`${AUTH_SERVICE_URL}/api/auth-request`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Agent-Session": agentSession,
    },
    body: JSON.stringify({
      service: "msgraph",
      scopes: SCOPES[service],
      duration: GRANT_DURATION[service],
    }),
  });

  if (!response.ok) {
    const { data, text } = await readErrorBody<ApiResponse>(response);
    if (data?.error || data?.message) {
      throw new Error(data.message || data.error);
    }
    throw new Error(
      `Failed to create auth request (${response.status}): ${text || response.statusText}`
    );
  }

  return response.json() as Promise<AuthRequestResponse>;
}

async function waitForGrantApproval(
  pollUrl: string,
  expiresAt?: string
): Promise<void> {
  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run './office-cli.ts login' first.");
    process.exit(1);
  }

  const deadline = expiresAt
    ? new Date(expiresAt).getTime()
    : Date.now() + 5 * 60 * 1000;
  let lastStatus: string | null = null;

  while (Date.now() < deadline) {
    const response = await fetch(pollUrl, {
      headers: { "X-Agent-Session": agentSession },
    });

    if (!response.ok) {
      const { data, text } = await readErrorBody<ApiResponse>(response);
      if (data) {
        describeAuthError(data, "msgraph");
        process.exit(1);
      }
      throw new Error(
        `Auth request poll failed (${response.status}): ${text || response.statusText}`
      );
    }

    const status = (await response.json()) as {
      status: string;
      message?: string;
    };

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

// Current service context for auto-request
let currentService: ServiceName = "drive";

async function ensureGrant(): Promise<void> {
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

async function apiRequest<T>(
  endpoint: string,
  options: RequestInit = {},
  attempt = 0
): Promise<T> {
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
      ...options.headers,
    },
    redirect: "manual",
  });

  if (
    response.status === 401 ||
    response.status === 403 ||
    response.status === 429
  ) {
    const { data, text } = await readErrorBody<ApiResponse>(response);
    if (attempt === 0 && AUTO_REQUEST_ENABLED && isAutoRequestError(data)) {
      await ensureGrant();
      return apiRequest<T>(endpoint, options, attempt + 1);
    }

    if (data) {
      describeAuthError(data, currentService);
      process.exit(1);
    }

    throw new Error(text || `Authentication failed (${response.status})`);
  }

  if (!response.ok) {
    const { data, text } = await readErrorBody<ApiResponse>(response);
    if (data?.error || data?.message) {
      throw new Error(data.message || data.error);
    }
    throw new Error(text || `API error (${response.status})`);
  }

  return response.json() as Promise<T>;
}

async function fetchGraphContent(
  endpoint: string,
  attempt = 0
): Promise<Response> {
  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run 'login' to create an agent session.");
    process.exit(1);
  }

  const response = await fetch(`${API_BASE}${endpoint}`, {
    headers: { "X-Agent-Session": agentSession },
    redirect: "manual",
  });

  if (
    response.status === 401 ||
    response.status === 403 ||
    response.status === 429
  ) {
    const { data, text } = await readErrorBody<ApiResponse>(response);
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

// =============================================================================
// Utilities
// =============================================================================

function formatBytes(bytes: number): string {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
}

function stripHtml(html: string): string {
  if (!html) return "";
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<\/?[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function truncate(str: string, len: number): string {
  if (str.length <= len) return str;
  return str.substring(0, len - 1) + "…";
}

// =============================================================================
// DRIVE: OneDrive Operations
// =============================================================================

interface DriveItem {
  id: string;
  name: string;
  size?: number;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl: string;
  file?: { mimeType: string };
  folder?: { childCount: number };
  parentReference?: { path: string };
}

async function driveList(folderPath?: string): Promise<void> {
  let endpoint = "/me/drive/root/children";
  if (folderPath) {
    const encodedPath = folderPath
      .split("/")
      .map(encodeURIComponent)
      .join("/");
    endpoint = `/me/drive/root:${encodedPath}:/children`;
  }
  const query =
    "?$select=id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,parentReference&$top=100";
  const result = await apiRequest<{ value: DriveItem[] }>(
    `${endpoint}${query}`
  );

  if (!result.value || result.value.length === 0) {
    console.log("No files found.");
    return;
  }

  console.log(`\nFiles in ${folderPath || "root"}:\n`);
  console.log("─".repeat(90));
  console.log(
    `${"Name".padEnd(40)} ${"Size".padEnd(10)} ${"Modified".padEnd(20)} ID`
  );
  console.log("─".repeat(90));

  for (const item of result.value) {
    const isFolder = !!item.folder;
    const name = isFolder ? `📁 ${item.name}` : `📄 ${item.name}`;
    const size = item.size ? formatBytes(item.size) : "-";
    const modified = new Date(item.lastModifiedDateTime).toLocaleDateString();
    console.log(
      `${name.padEnd(40)} ${size.padEnd(10)} ${modified.padEnd(20)} ${item.id}`
    );
  }
}

async function driveSearch(query: string): Promise<void> {
  const endpoint = `/me/drive/root/search(q='${encodeURIComponent(query)}')`;
  const select =
    "?$select=id,name,size,lastModifiedDateTime,webUrl,file,folder,parentReference&$top=25";
  const result = await apiRequest<{ value: DriveItem[] }>(
    `${endpoint}${select}`
  );

  if (!result.value || result.value.length === 0) {
    console.log(`No files found matching "${query}".`);
    return;
  }

  // Batch-fetch full paths
  const ids = result.value.map((item) => item.id);
  let pathMap: Record<string, string> = {};

  const chunks: string[][] = [];
  for (let i = 0; i < ids.length; i += 20) {
    chunks.push(ids.slice(i, i + 20));
  }

  const batchResults = await Promise.all(
    chunks.map(async (chunk) => {
      try {
        const batchRequests = chunk.map((id, index) => ({
          id: String(index),
          method: "GET",
          url: `/me/drive/items/${id}?$select=id,name,parentReference`,
        }));
        const batchResult = await apiRequest<{
          responses: Array<{ id: string; status: number; body?: any }>;
        }>("/$batch", {
          method: "POST",
          body: JSON.stringify({ requests: batchRequests }),
        });
        const paths: Record<string, string> = {};
        for (const resp of batchResult.responses || []) {
          if (resp.status === 200 && resp.body?.parentReference?.path) {
            const parsedIdx = parseInt(resp.id, 10);
            if (Number.isNaN(parsedIdx) || parsedIdx < 0 || parsedIdx >= chunk.length) continue;
            const itemId = chunk[parsedIdx];
            const rawPath = resp.body.parentReference.path as string;
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

  console.log(`\nSearch results for "${query}":\n`);

  for (const item of result.value) {
    const icon = item.folder ? "📁" : "📄";
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

async function driveDownload(
  itemId: string,
  outputPath?: string
): Promise<string> {
  const info = await apiRequest<DriveItem>(`/me/drive/items/${itemId}`);
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

async function driveContent(itemId: string): Promise<void> {
  const info = await apiRequest<DriveItem>(
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
    "conf",
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

async function driveInfo(itemId: string): Promise<void> {
  const item = await apiRequest<DriveItem>(`/me/drive/items/${itemId}`);
  const rawPath = item.parentReference?.path || "";
  const path = rawPath.replace(/^\/drive\/root:?/, "") || "/";

  console.log("\nFile Information:");
  console.log("─".repeat(50));
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

// Drive share / permissions

interface DrivePermission {
  id: string;
  roles: string[];
  link?: { type: string; scope: string; webUrl: string };
  grantedToV2?: {
    user?: { displayName?: string; email?: string };
    siteUser?: { displayName?: string; email?: string };
  };
  grantedTo?: { user?: { displayName?: string; email?: string } };
  invitation?: { email?: string };
}

async function driveShare(
  itemId: string,
  options: { email?: string; role?: string; type?: string; scope?: string; anyone?: boolean; expires?: string }
): Promise<void> {
  // --anyone flag: anonymous edit link with default 7-day expiry
  if (options.anyone) {
    options.scope = options.scope || "anonymous";
    options.type = options.type || "edit";
    if (!options.expires) {
      const expiry = new Date();
      expiry.setDate(expiry.getDate() + 7);
      options.expires = expiry.toISOString();
    }
  }

  if (options.email) {
    // Share with specific person via invite
    const role = options.role || "read";
    if (!["read", "write"].includes(role)) {
      throw new Error("Role must be 'read' or 'write'");
    }
    const result = await apiRequest<{ value: DrivePermission[] }>(
      `/me/drive/items/${itemId}/invite`,
      {
        method: "POST",
        body: JSON.stringify({
          recipients: [{ email: options.email }],
          requireSignIn: true,
          sendInvitation: true,
          roles: [role],
          message: "Shared via Office CLI",
        }),
      }
    );

    console.log(`\nShared with ${options.email} (${role} access)`);
    if (result.value?.[0]?.link?.webUrl) {
      console.log(`Link: ${result.value[0].link.webUrl}`);
    }
  } else {
    // Create a shareable link
    const type = options.type || "view";
    const scope = options.scope || "organization";

    if (!["view", "edit", "embed"].includes(type)) {
      throw new Error("Link type must be 'view', 'edit', or 'embed'");
    }
    if (!["anonymous", "organization", "users"].includes(scope)) {
      throw new Error("Scope must be 'anonymous', 'organization', or 'users'");
    }

    const body: Record<string, string> = { type, scope };
    if (options.expires) {
      body.expirationDateTime = options.expires.includes("T")
        ? options.expires
        : `${options.expires}T00:00:00Z`;
    }

    try {
      const result = await apiRequest<{ link: { webUrl: string; type: string; scope: string } }>(
        `/me/drive/items/${itemId}/createLink`,
        {
          method: "POST",
          body: JSON.stringify(body),
        }
      );

      console.log(`\nShare link created (${type}, ${scope}):`);
      console.log(result.link.webUrl);
      if (options.expires) {
        const expiryDate = new Date(body.expirationDateTime);
        console.log(`Expires: ${expiryDate.toLocaleDateString("en-AU", { weekday: "short", year: "numeric", month: "short", day: "numeric" })}`);
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      if (scope === "anonymous" && (msg.includes("invalidRequest") || msg.includes("not allowed") || msg.includes("disabled"))) {
        console.error(`\nFailed to create anonymous ${type} link.`);
        console.error("Anonymous sharing may be disabled by your SharePoint admin.");
        console.error("Admin needs to enable in: SharePoint Admin > Policies > Sharing > set to 'Anyone'");
        console.error("\nTry --scope organization instead for org-wide access.");
        process.exit(1);
      }
      throw err;
    }
  }
}

async function drivePermissions(itemId: string): Promise<void> {
  const info = await apiRequest<DriveItem>(`/me/drive/items/${itemId}?$select=id,name`);
  const result = await apiRequest<{ value: DrivePermission[] }>(
    `/me/drive/items/${itemId}/permissions`
  );

  if (!result.value || result.value.length === 0) {
    console.log(`No permissions found for ${info.name}.`);
    return;
  }

  console.log(`\nPermissions for "${info.name}":\n`);
  console.log("─".repeat(80));

  for (const perm of result.value) {
    const roles = perm.roles.join(", ");

    if (perm.link) {
      console.log(`🔗 Link (${perm.link.type}, ${perm.link.scope})`);
      console.log(`   URL: ${perm.link.webUrl}`);
      console.log(`   Roles: ${roles}`);
      console.log(`   Permission ID: ${perm.id}`);
    } else {
      const user =
        perm.grantedToV2?.user ||
        perm.grantedToV2?.siteUser ||
        perm.grantedTo?.user;
      const email = user?.email || perm.invitation?.email || "unknown";
      const name = user?.displayName || email;
      console.log(`👤 ${name}`);
      if (email !== name) console.log(`   Email: ${email}`);
      console.log(`   Roles: ${roles}`);
      console.log(`   Permission ID: ${perm.id}`);
    }
    console.log();
  }
}

async function driveUnshare(itemId: string, permissionId: string): Promise<void> {
  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run './office-cli.ts login' first.");
    process.exit(1);
  }

  const response = await fetch(
    `${API_BASE}/me/drive/items/${itemId}/permissions/${permissionId}`,
    {
      method: "DELETE",
      headers: { "X-Agent-Session": agentSession },
    }
  );

  if (response.status === 204 || response.status === 200) {
    console.log(`Permission ${permissionId} removed.`);
    return;
  }

  if (response.status === 401 || response.status === 403) {
    const { data } = await readErrorBody<ApiResponse>(response);
    if (isAutoRequestError(data) && AUTO_REQUEST_ENABLED) {
      await ensureGrant();
      const retry = await fetch(
        `${API_BASE}/me/drive/items/${itemId}/permissions/${permissionId}`,
        {
          method: "DELETE",
          headers: { "X-Agent-Session": agentSession },
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

// Drive convert (PDF/image to markdown)
const MARKDOWN_CACHE_FOLDER = ".markdown";

async function driveConvert(
  itemId: string,
  options: { force?: boolean; output?: string }
): Promise<void> {
  // Check cache
  if (!options.force) {
    console.error("Checking OneDrive cache...");
    try {
      const encodedPath = `/${MARKDOWN_CACHE_FOLDER}`
        .split("/")
        .map(encodeURIComponent)
        .join("/");
      const result = await apiRequest<{ value: DriveItem[] }>(
        `/me/drive/root:${encodedPath}:/children?$select=id,name&$top=200`
      );
      const cached = result.value?.find((i) => i.name === `${itemId}.md`);
      if (cached) {
        console.error(`Found cached version: ${cached.name}`);
        let response = await fetchGraphContent(
          `/me/drive/items/${cached.id}/content`
        );
        if (response.status === 302 || response.status === 301) {
          const url = response.headers.get("location");
          if (url) response = await fetch(url);
        }
        const markdown = await response.text();
        if (options.output) {
          await writeFile(options.output, markdown);
          console.error(`Saved to: ${options.output}`);
        } else {
          console.log(markdown);
        }
        return;
      }
    } catch {
      // Cache folder may not exist
    }
    console.error("No cache found, converting...");
  }

  // Download file
  console.error("Downloading file...");
  const info = await apiRequest<DriveItem>(
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
    let images: string[];
    let imageMimeType = "image/png";

    if (mimeType === "application/pdf" || ext === ".pdf") {
      console.error("Converting PDF pages to images...");
      images = await convertPdfToImages(tempPath);
      console.error(`Extracted ${images.length} page(s)`);
    } else if (
      mimeType.startsWith("image/") ||
      [".png", ".jpg", ".jpeg", ".gif", ".webp"].includes(ext)
    ) {
      images = [await imageToBase64(tempPath)];
      imageMimeType = mimeType.startsWith("image/")
        ? mimeType
        : getMimeType(ext);
    } else {
      throw new Error(
        `Unsupported file type: ${mimeType}. Supported: PDF, PNG, JPG, GIF, WebP`
      );
    }

    const apiKey = process.env.OPENROUTER_API_KEY;
    if (!apiKey) throw new Error("OPENROUTER_API_KEY environment variable required");

    console.error(`Converting ${images.length} page(s) with Gemini 3 Flash...`);
    const content: Array<{ type: string; text?: string; image_url?: { url: string } }> = [];
    for (const base64 of images) {
      content.push({
        type: "image_url",
        image_url: { url: `data:${imageMimeType};base64,${base64}` },
      });
    }
    content.push({
      type: "text",
      text: `Convert this document "${info.name}" to well-formatted markdown.\nPreserve structure: headings, lists, tables, formatting.\nFor multi-page documents, use horizontal rules (---) between pages.\nExtract all text accurately. For tables, use markdown table syntax.\nDo not include any commentary or explanations - just the converted content.`,
    });

    const geminiResponse = await fetch(
      "https://openrouter.ai/api/v1/chat/completions",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          model: "google/gemini-3-flash-preview",
          messages: [{ role: "user", content }],
        }),
      }
    );

    if (!geminiResponse.ok) {
      throw new Error(
        `OpenRouter API error (${geminiResponse.status}): ${await geminiResponse.text()}`
      );
    }

    const result = (await geminiResponse.json()) as {
      choices: Array<{ message: { content: string } }>;
      error?: { message: string };
    };
    if (result.error) throw new Error(`OpenRouter error: ${result.error.message}`);

    const markdown = result.choices[0].message.content;

    // Upload to cache
    const cachePath = `/${MARKDOWN_CACHE_FOLDER}/${itemId}.md`;
    const encodedCachePath = cachePath
      .split("/")
      .map(encodeURIComponent)
      .join("/");
    console.error(`Uploading cached markdown to OneDrive ${cachePath}...`);
    await apiRequest(`/me/drive/root:${encodedCachePath}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "text/plain" },
      body: markdown,
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
    } catch {}
  }
}

function getMimeType(ext: string): string {
  const mimeTypes: Record<string, string> = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".pdf": "application/pdf",
  };
  return mimeTypes[ext.toLowerCase()] || "application/octet-stream";
}

async function convertPdfToImages(pdfPath: string): Promise<string[]> {
  const outputDir = join(dirname(pdfPath), "pages");
  await mkdir(outputDir, { recursive: true });
  const outputPrefix = join(outputDir, "page");

  return new Promise((resolve, reject) => {
    const proc = spawn("pdftoppm", [
      "-png",
      "-r",
      "150",
      pdfPath,
      outputPrefix,
    ]);
    let stderr = "";
    proc.stderr.on("data", (d) => (stderr += d.toString()));
    proc.on("close", async (code) => {
      if (code !== 0) {
        reject(new Error(`pdftoppm failed: ${stderr}`));
        return;
      }
      try {
        const files = await readdir(outputDir);
        const pngFiles = files.filter((f) => f.endsWith(".png")).sort();
        const images: string[] = [];
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

async function imageToBase64(imagePath: string): Promise<string> {
  return (await readFile(imagePath)).toString("base64");
}

// =============================================================================
// MAIL: Outlook Email Operations
// =============================================================================

interface MailMessage {
  id: string;
  subject: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  toRecipients?: Array<{
    emailAddress?: { name?: string; address?: string };
  }>;
  ccRecipients?: Array<{
    emailAddress?: { name?: string; address?: string };
  }>;
  receivedDateTime: string;
  isRead: boolean;
  hasAttachments: boolean;
  importance: string;
  bodyPreview?: string;
  body?: { content?: string; contentType?: string };
  conversationId?: string;
  webLink?: string;
}

function formatEmailAddress(addr?: {
  name?: string;
  address?: string;
}): string {
  if (!addr) return "";
  const name = addr.name?.trim();
  const email = addr.address?.trim();
  if (name && email) return `${name} <${email}>`;
  return name || email || "";
}

async function mailInbox(count = 20, unreadOnly = false, folder = "inbox"): Promise<void> {
  let filter = "";
  if (unreadOnly) filter = "&$filter=isRead eq false";

  const folderPath = folder === "inbox"
    ? `/me/messages`
    : `/me/mailFolders/${folder}/messages`;

  const result = await apiRequest<{ value: MailMessage[] }>(
    `${folderPath}?$select=id,subject,from,receivedDateTime,isRead,hasAttachments,importance,bodyPreview&$orderby=receivedDateTime desc&$top=${count}${filter}`
  );

  if (!result.value || result.value.length === 0) {
    console.log(unreadOnly ? "No unread messages." : "No messages found.");
    return;
  }

  const folderLabel = folder === "inbox" ? (unreadOnly ? "Unread" : "Recent") : folder.charAt(0).toUpperCase() + folder.slice(1);
  console.log(
    `\n${folderLabel} messages (${result.value.length}):\n`
  );
  console.log("─".repeat(95));
  console.log(
    `${"".padEnd(2)} ${"From".padEnd(25)} ${"Subject".padEnd(40)} ${"Date".padEnd(18)} ID`
  );
  console.log("─".repeat(95));

  for (const msg of result.value) {
    const read = msg.isRead ? " " : "●";
    const from = truncate(
      formatEmailAddress(msg.from?.emailAddress),
      24
    );
    const subj = truncate(msg.subject || "(No subject)", 39);
    const date = new Date(msg.receivedDateTime).toLocaleDateString("en-AU", {
      day: "2-digit",
      month: "short",
      hour: "2-digit",
      minute: "2-digit",
    });
    const attach = msg.hasAttachments ? "📎" : "";

    console.log(
      `${read} ${from.padEnd(25)} ${subj.padEnd(40)} ${date.padEnd(18)} ${msg.id.substring(0, 20)}…`
    );
    if (attach) console.log(`  ${attach}`);
  }

  console.log(
    `\nUse 'office-cli.ts mail read <id>' to read a message.`
  );
}

async function mailSearch(query: string, count = 15, includeJunk = false): Promise<void> {
  // Search inbox folder by default to exclude junk/spam; use includeJunk for all folders
  const folderPath = includeJunk ? `/me/messages` : `/me/mailFolders/inbox/messages`;
  const result = await apiRequest<{ value: MailMessage[] }>(
    `${folderPath}?$search="${encodeURIComponent(query)}"&$select=id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview&$top=${count}`
  );

  if (!result.value || result.value.length === 0) {
    console.log(`No messages found matching "${query}".`);
    return;
  }

  console.log(`\nSearch results for "${query}" (${result.value.length}):\n`);

  for (const msg of result.value) {
    const from = formatEmailAddress(msg.from?.emailAddress);
    const date = new Date(msg.receivedDateTime).toLocaleString("en-AU");
    const read = msg.isRead ? "" : " [UNREAD]";

    console.log(`${msg.isRead ? "📧" : "📬"} ${msg.subject || "(No subject)"}${read}`);
    console.log(`   From: ${from}`);
    console.log(`   Date: ${date}`);
    console.log(`   ID: ${msg.id}`);
    if (msg.bodyPreview) {
      console.log(`   Preview: ${truncate(msg.bodyPreview, 100)}`);
    }
    console.log();
  }
}

async function mailRead(messageId: string): Promise<void> {
  const msg = await apiRequest<MailMessage>(
    `/me/messages/${messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,importance,body,webLink`,
    {
      headers: { Prefer: 'outlook.body-content-type="text"' },
    }
  );

  console.log("\n" + "═".repeat(70));
  console.log(`Subject: ${msg.subject || "(No subject)"}`);
  console.log("─".repeat(70));
  console.log(`From: ${formatEmailAddress(msg.from?.emailAddress)}`);

  if (msg.toRecipients?.length) {
    const to = msg.toRecipients
      .map((r) => formatEmailAddress(r.emailAddress))
      .join(", ");
    console.log(`To: ${to}`);
  }
  if (msg.ccRecipients?.length) {
    const cc = msg.ccRecipients
      .map((r) => formatEmailAddress(r.emailAddress))
      .join(", ");
    console.log(`CC: ${cc}`);
  }

  console.log(
    `Date: ${new Date(msg.receivedDateTime).toLocaleString("en-AU")}`
  );
  console.log(`Importance: ${msg.importance}`);
  if (msg.hasAttachments) console.log("Attachments: Yes 📎");
  console.log(`ID: ${msg.id}`);
  if (msg.webLink) console.log(`Web: ${msg.webLink}`);
  console.log("─".repeat(70));

  if (msg.body?.content) {
    const bodyText =
      msg.body.contentType?.toLowerCase() === "html"
        ? stripHtml(msg.body.content)
        : msg.body.content;
    console.log(bodyText);
  } else {
    console.log("(No body content)");
  }
  console.log("═".repeat(70));
}

// ============================================
// People API
// ============================================

interface PersonResult {
  displayName: string;
  givenName?: string;
  surname?: string;
  jobTitle?: string;
  department?: string;
  scoredEmailAddresses?: Array<{ address: string; relevanceScore?: number }>;
  personType?: { class?: string; subclass?: string };
  userPrincipalName?: string;
}

async function peopleSearch(query: string, top = 10): Promise<PersonResult[]> {
  const result = await apiRequest<{ value: PersonResult[] }>(
    `/me/people?$search="${encodeURIComponent(query)}"&$select=displayName,givenName,surname,scoredEmailAddresses,personType,jobTitle,department,userPrincipalName&$top=${top}`,
    {
      headers: {
        "X-PeopleQuery-QuerySources": "Mailbox,Directory",
      },
    }
  );
  return result.value || [];
}

async function peopleList(query: string, verbose = false): Promise<void> {
  const people = await peopleSearch(query);

  if (people.length === 0) {
    console.log(`No results for "${query}".`);
    return;
  }

  console.log(`\nPeople matching "${query}":\n`);

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

/**
 * Resolve a name or email to a list of email addresses.
 * If the input looks like an email, returns it directly.
 * Otherwise, searches the People API and returns matches.
 */
async function resolveRecipient(nameOrEmail: string): Promise<{ address: string; displayName: string }[]> {
  if (validateEmail(nameOrEmail)) {
    return [{ address: nameOrEmail, displayName: nameOrEmail }];
  }

  const people = await peopleSearch(nameOrEmail, 5);
  return people
    .filter((p) => p.scoredEmailAddresses?.[0]?.address)
    .map((p) => ({
      address: p.scoredEmailAddresses![0].address,
      displayName: p.displayName,
    }));
}

/**
 * Resolve a recipient string that may be a name or email.
 * For comma-separated values, resolves each independently.
 *
 * Without --yes: exits with a "did you mean?" message (safe for agents).
 * With --yes: auto-picks top result (agent has already confirmed with user).
 */
async function resolveRecipients(
  input: string,
  autoConfirm = false
): Promise<Array<{ emailAddress: { address: string } }>> {
  const parts = input.split(",").map((s) => s.trim());
  const resolved: Array<{ emailAddress: { address: string } }> = [];

  for (const part of parts) {
    if (validateEmail(part)) {
      resolved.push({ emailAddress: { address: part } });
      continue;
    }

    // It's a name - look up via People API
    const matches = await resolveRecipient(part);

    if (matches.length === 0) {
      console.error(`No matches found for "${part}". Please provide a full email address.`);
      process.exit(1);
    }

    // Show matches
    console.log(`\n  Matches for "${part}":`);
    const shown = matches.slice(0, 5);
    for (let i = 0; i < shown.length; i++) {
      console.log(`    ${i + 1}. ${shown[i].displayName} <${shown[i].address}>`);
    }

    if (autoConfirm) {
      // --yes mode: auto-pick top result
      console.log(`  → Using: ${shown[0].displayName} <${shown[0].address}>`);
      resolved.push({ emailAddress: { address: shown[0].address } });
    } else {
      // Safe mode: show matches and exit so caller can re-run with exact email
      console.log(`\n  Re-run with the exact email, or use --yes to auto-pick the top result.`);
      process.exit(1);
    }
  }

  return resolved;
}

function validateEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
}

function parseRecipients(csv: string): Array<{ emailAddress: { address: string } }> {
  return csv.split(",").map((email) => {
    const trimmed = email.trim();
    if (!validateEmail(trimmed)) {
      throw new Error(`Invalid email address: ${trimmed}`);
    }
    return { emailAddress: { address: trimmed } };
  });
}

async function mailSend(options: {
  to: string;
  subject: string;
  body: string;
  cc?: string;
  isHtml?: boolean;
  autoConfirm?: boolean;
}): Promise<void> {
  // Resolve names to email addresses via People API
  const toRecipients = await resolveRecipients(options.to, options.autoConfirm);

  const payload: any = {
    message: {
      subject: options.subject,
      body: {
        contentType: options.isHtml ? "HTML" : "Text",
        content: options.body,
      },
      toRecipients,
    },
    saveToSentItems: true,
  };

  if (options.cc) {
    payload.message.ccRecipients = await resolveRecipients(options.cc, options.autoConfirm);
  }

  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run './office-cli.ts login' first.");
    process.exit(1);
  }

  const response = await fetch(`${API_BASE}/me/sendMail`, {
    method: "POST",
    headers: {
      "X-Agent-Session": agentSession,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (response.status === 202 || response.status === 200) {
    console.log(`Email sent to ${options.to}`);
    return;
  }

  // Handle auth errors with auto-request
  if (response.status === 401 || response.status === 403) {
    const { data, text } = await readErrorBody<ApiResponse>(response);
    if (isAutoRequestError(data) && AUTO_REQUEST_ENABLED) {
      await ensureGrant();
      // Retry
      const retry = await fetch(`${API_BASE}/me/sendMail`, {
        method: "POST",
        headers: {
          "X-Agent-Session": agentSession,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
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

  const { text: errText } = await readErrorBody<ApiResponse>(response);
  throw new Error(
    `Send failed (${response.status}): ${errText || response.statusText}`
  );
}

async function mailReply(
  messageId: string,
  comment: string,
  replyAll = false
): Promise<void> {
  const endpoint = replyAll
    ? `/me/messages/${messageId}/replyAll`
    : `/me/messages/${messageId}/reply`;

  const agentSession = await getAgentSession();
  if (!agentSession) {
    console.error("Not registered. Run './office-cli.ts login' first.");
    process.exit(1);
  }

  const response = await fetch(`${API_BASE}${endpoint}`, {
    method: "POST",
    headers: {
      "X-Agent-Session": agentSession,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ comment }),
  });

  if (response.status === 202 || response.status === 200) {
    console.log(replyAll ? "Reply-all sent." : "Reply sent.");
    return;
  }

  if (response.status === 401 || response.status === 403) {
    const { data, text } = await readErrorBody<ApiResponse>(response);
    if (isAutoRequestError(data) && AUTO_REQUEST_ENABLED) {
      await ensureGrant();
      const retry = await fetch(`${API_BASE}${endpoint}`, {
        method: "POST",
        headers: {
          "X-Agent-Session": agentSession,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ comment }),
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

  const { text: replyErrText } = await readErrorBody<ApiResponse>(response);
  throw new Error(
    `Reply failed (${response.status}): ${replyErrText || response.statusText}`
  );
}

// =============================================================================
// CAL: Calendar Operations
// =============================================================================

interface CalendarEvent {
  subject: string;
  start: { dateTime?: string; date?: string; timeZone?: string };
  end: { dateTime?: string; date?: string; timeZone?: string };
  location?: { displayName?: string };
  isAllDay: boolean;
  organizer?: { emailAddress?: { name?: string; address?: string } };
  attendees?: Array<{
    emailAddress?: { name?: string; address?: string };
    type?: string;
  }>;
  body?: { content?: string; contentType?: string };
}

function parseEventDate(dateStr: string, timeZone?: string): Date {
  if (timeZone === "UTC" && !dateStr.endsWith("Z")) {
    return new Date(dateStr + "Z");
  }
  return new Date(dateStr);
}

function formatTime(
  dateStr: string,
  isAllDay: boolean,
  timeZone?: string
): string {
  if (isAllDay) return "All day";
  const date = parseEventDate(dateStr, timeZone);
  return date.toLocaleTimeString("en-AU", {
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  });
}

function formatDateHeader(date: Date): string {
  const today = new Date();
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
    day: "numeric",
  });

  if (dateOnly.getTime() === todayOnly.getTime())
    return `Today - ${formatted}`;
  if (dateOnly.getTime() === tomorrowOnly.getTime())
    return `Tomorrow - ${formatted}`;
  return formatted;
}

function getEventDateKey(event: CalendarEvent): string {
  const startStr = event.start.dateTime || event.start.date || "";
  const date = parseEventDate(startStr, event.start.timeZone);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

async function calEvents(days = 7, includeDetails = false): Promise<void> {
  const startDate = new Date();
  startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(startDate);
  endDate.setDate(endDate.getDate() + days);

  console.error(
    `Fetching calendar events for the next ${days} day${days !== 1 ? "s" : ""}...`
  );

  const selectFields = includeDetails
    ? "subject,start,end,location,isAllDay,organizer,attendees,body"
    : "subject,start,end,location,isAllDay,organizer";

  const headers: Record<string, string> = {};
  if (includeDetails)
    headers["Prefer"] = 'outlook.body-content-type="text"';

  const result = await apiRequest<{ value: CalendarEvent[] }>(
    `/me/calendarView?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}&$select=${selectFields}&$orderby=start/dateTime&$top=50`,
    { headers }
  );

  const events = result.value || [];
  if (events.length === 0) {
    console.log("No events scheduled for this period.");
    return;
  }

  // Group by date
  const eventsByDate: Record<string, CalendarEvent[]> = {};
  for (const event of events) {
    const key = getEventDateKey(event);
    if (!eventsByDate[key]) eventsByDate[key] = [];
    eventsByDate[key].push(event);
  }

  console.log("");
  for (const dateKey of Object.keys(eventsByDate).sort()) {
    const [year, month, day] = dateKey.split("-").map(Number);
    const date = new Date(year, month - 1, day);
    console.log(`## ${formatDateHeader(date)}\n`);

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
      const timeRange = event.isAllDay
        ? "All day"
        : `${startTime} - ${endTime}`;

      console.log(`  **${event.subject || "(No title)"}**`);
      console.log(`  ${timeRange}`);
      if (event.location?.displayName)
        console.log(`  Location: ${event.location.displayName}`);

      if (includeDetails) {
        const organizer = formatEmailAddress(event.organizer?.emailAddress);
        if (organizer) console.log(`  Organizer: ${organizer}`);

        const attendees = (event.attendees || [])
          .map((a) => {
            const label = formatEmailAddress(a.emailAddress);
            return a.type ? `${label} (${a.type.toLowerCase()})` : label;
          })
          .filter(Boolean);
        if (attendees.length > 0) {
          console.log("  Participants:");
          for (const a of attendees) console.log(`    - ${a}`);
        }

        if (event.body?.content) {
          const bodyText =
            event.body.contentType?.toLowerCase() === "html"
              ? stripHtml(event.body.content)
              : event.body.content.trim();
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

async function calTomorrow(includeDetails = false): Promise<void> {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  const dayAfter = new Date(tomorrow);
  dayAfter.setDate(dayAfter.getDate() + 1);

  console.error("Fetching tomorrow's events...");

  const selectFields = includeDetails
    ? "subject,start,end,location,isAllDay,organizer,attendees,body"
    : "subject,start,end,location,isAllDay,organizer";

  const headers: Record<string, string> = {};
  if (includeDetails)
    headers["Prefer"] = 'outlook.body-content-type="text"';

  const result = await apiRequest<{ value: CalendarEvent[] }>(
    `/me/calendarView?startDateTime=${tomorrow.toISOString()}&endDateTime=${dayAfter.toISOString()}&$select=${selectFields}&$orderby=start/dateTime&$top=50`,
    { headers }
  );

  const events = result.value || [];
  if (events.length === 0) {
    console.log("No events scheduled for tomorrow.");
    return;
  }

  console.log(`\n## ${formatDateHeader(tomorrow)}\n`);

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
    const timeRange = event.isAllDay
      ? "All day"
      : `${startTime} - ${endTime}`;

    console.log(`  **${event.subject || "(No title)"}**`);
    console.log(`  ${timeRange}`);
    if (event.location?.displayName)
      console.log(`  Location: ${event.location.displayName}`);

    if (includeDetails) {
      const organizer = formatEmailAddress(event.organizer?.emailAddress);
      if (organizer) console.log(`  Organizer: ${organizer}`);

      const attendees = (event.attendees || [])
        .map((a) => {
          const label = formatEmailAddress(a.emailAddress);
          return a.type ? `${label} (${a.type.toLowerCase()})` : label;
        })
        .filter(Boolean);
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

// =============================================================================
// Status Check (all services)
// =============================================================================

async function checkStatus(service?: string): Promise<void> {
  const session = await getAgentSession();

  console.log("Office CLI Status\n");
  console.log(`Auth Service: ${AUTH_SERVICE_URL}`);
  console.log(`Agent session file: ${AGENT_SESSION_FILE}`);
  console.log(`Agent session saved: ${session ? "Yes" : "No"}`);

  if (!session) {
    console.log("\n❌ Not registered");
    console.log("");
    console.log("To register:");
    console.log("1. Run: ./office-cli.ts login");
    console.log("2. Follow the approval link shown");
    return;
  }

  console.log("\n✅ Agent session saved");

  // Check grant status
  try {
    const response = await fetch(`${STATUS_BASE}/msgraph`, {
      headers: { "X-Agent-Session": session },
    });

    if (!response.ok) {
      const data = (await response.json()) as ApiResponse;
      describeAuthError(data, "msgraph");
      return;
    }

    const status = (await response.json()) as GrantStatusResponse;

    if (!status.hasGrant) {
      console.log("\n⚠️  No active grant for Microsoft Graph");
      console.log("Run: ./office-cli.ts request <drive|mail|cal>");
      return;
    }

    console.log("\n✅ Grant active for Microsoft Graph");
    if (status.scopes?.length) {
      console.log(`   Scopes: ${status.scopes.join(" ")}`);

      // Show which services are covered
      const activeScopes = new Set(status.scopes);
      const services: Array<{ name: string; scopes: readonly string[] }> = [
        { name: "Drive", scopes: SCOPES.drive },
        { name: "Mail", scopes: SCOPES.mail },
        { name: "Calendar", scopes: SCOPES.cal },
      ];
      for (const svc of services) {
        const covered = svc.scopes.every((s) => activeScopes.has(s));
        console.log(`   ${covered ? "✅" : "⚠️ "} ${svc.name}: ${svc.scopes.join(", ")}`);
      }
    }
    if (status.expiresAt) {
      console.log(
        `   Expires at: ${new Date(status.expiresAt).toLocaleString()}`
      );
    }
  } catch {
    // Handled above
  }
}

// =============================================================================
// Login / Registration
// =============================================================================

async function doLogin(agentName?: string): Promise<void> {
  const name = agentName || "Office CLI";
  const deviceInfo = `${hostname()} (${process.platform})`;

  const response = await fetch(`${AUTH_SERVICE_URL}/api/agent/init`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ agentName: name, deviceInfo }),
  });

  if (!response.ok) {
    throw new Error(
      `Failed to start agent registration: ${await response.text()}`
    );
  }

  const data = (await response.json()) as {
    registrationId: string;
    code: string;
    pollUrl: string;
    verifyUrl?: string;
    expiresIn: number;
    expiresAt: string;
  };

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

    const status = (await poll.json()) as {
      status: "pending" | "approved" | "denied" | "expired";
      agentSessionId?: string;
      message?: string;
    };

    if (status.status === "approved" && status.agentSessionId) {
      await saveAgentSession(status.agentSessionId);
      console.log("Agent session saved.");
      console.log("Next: run './office-cli.ts request <drive|mail|cal>' to request access.");
      break;
    }

    if (status.status === "denied")
      throw new Error(status.message || "Registration denied.");
    if (status.status === "expired")
      throw new Error("Registration expired. Run login again.");
    if (Date.now() - pollStart > 10 * 60 * 1000)
      throw new Error("Timed out waiting for approval.");

    await new Promise((resolve) => setTimeout(resolve, 2000));
  }
}

// =============================================================================
// CLI Interface
// =============================================================================

function printUsage(): void {
  console.log(`
Office CLI - Unified Microsoft 365 CLI via Auth Service Proxy

Uses agent sessions + grants. Tokens are managed by the auth
service for better security isolation.

Usage:
  office-cli.ts <command> [options]
  office-cli.ts <service> <command> [options]

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
  1. ./office-cli.ts login
  2. Approve agent registration in browser
  3. ./office-cli.ts request drive   (for OneDrive)
     ./office-cli.ts request mail    (for Outlook email)
     ./office-cli.ts request cal     (for Calendar)

Environment:
  AUTH_SERVICE_URL              Auth service URL (default: https://auth.freshhub.ai)
  OPENROUTER_API_KEY            Required for drive convert command
  OFFICE_AUTO_REQUEST=0         Disable auto-request on missing grants
`);
}

function parseNamedArg(args: string[], flag: string): string | undefined {
  // Handle --flag=value
  for (const arg of args) {
    if (arg.startsWith(`${flag}=`)) return arg.substring(flag.length + 1);
  }
  // Handle --flag value
  const idx = args.indexOf(flag);
  return idx !== -1 && args[idx + 1] ? args[idx + 1] : undefined;
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);

  if (
    args.length === 0 ||
    args[0] === "help" ||
    args[0] === "--help" ||
    args[0] === "-h"
  ) {
    printUsage();
    process.exit(0);
  }

  const command = args[0];

  try {
    // Top-level commands
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
        const service = args[1] as ServiceName;
        if (!service || !SCOPES[service]) {
          console.error("Usage: office-cli.ts request <drive|mail|cal>");
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

    // Service subcommands
    const subCommand = args[1];

    switch (command) {
      // ── Drive ──────────────────────────────────────────────
      case "drive": {
        currentService = "drive";
        if (!subCommand) {
          console.error("Usage: office-cli.ts drive <list|search|download|content|info|convert>");
          process.exit(1);
        }

        switch (subCommand) {
          case "list":
            await driveList(args[2]);
            break;
          case "search":
            if (!args[2]) {
              console.error("Usage: office-cli.ts drive search <query>");
              process.exit(1);
            }
            await driveSearch(args[2]);
            break;
          case "download":
            if (!args[2]) {
              console.error("Usage: office-cli.ts drive download <id> [output]");
              process.exit(1);
            }
            await driveDownload(args[2], args[3]);
            break;
          case "content":
            if (!args[2]) {
              console.error("Usage: office-cli.ts drive content <id>");
              process.exit(1);
            }
            await driveContent(args[2]);
            break;
          case "info":
            if (!args[2]) {
              console.error("Usage: office-cli.ts drive info <id>");
              process.exit(1);
            }
            await driveInfo(args[2]);
            break;
          case "convert": {
            if (!args[2]) {
              console.error("Usage: office-cli.ts drive convert <id> [--force] [--output=<path>]");
              process.exit(1);
            }
            const convertOpts: { force?: boolean; output?: string } = {};
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
              console.error("Usage: office-cli.ts drive share <id> [--anyone] [--expires <date>] [--email <email>] [--role read|write] [--type view|edit] [--scope organization|anonymous]");
              process.exit(1);
            }
            const shareOpts: { email?: string; role?: string; type?: string; scope?: string; anyone?: boolean; expires?: string } = {};
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
              console.error("Usage: office-cli.ts drive permissions <id>");
              process.exit(1);
            }
            await drivePermissions(args[2]);
            break;
          }
          case "unshare": {
            if (!args[2] || !args[3]) {
              console.error("Usage: office-cli.ts drive unshare <item-id> <permission-id>");
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
          console.error("Usage: office-cli.ts mail <inbox|unread|junk|search|read|send|reply|reply-all>");
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
              console.error("Usage: office-cli.ts mail search <query> [--include-junk]");
              process.exit(1);
            }
            await mailSearch(args[2], 15, args.slice(3).includes("--include-junk"));
            break;
          case "read":
            if (!args[2]) {
              console.error("Usage: office-cli.ts mail read <message-id>");
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
                "Usage: office-cli.ts mail send --to <email> --subject <subject> --body <text>"
              );
              console.error(
                "   or: office-cli.ts mail send --to <email> --subject <subject> --body <html> --html"
              );
              console.error(
                "   or: office-cli.ts mail send --to <email> --subject <subject> --file <path>"
              );
              process.exit(1);
            }

            let body: string;
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
              // Try stdin
              const chunks: Buffer[] = [];
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
                `Usage: office-cli.ts mail ${subCommand} <message-id> --body <text>`
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
          console.error("Usage: office-cli.ts cal <events|today|tomorrow>");
          process.exit(1);
        }

        switch (subCommand) {
          case "events": {
            const daysStr = parseNamedArg(args.slice(2), "--days");
            const days = daysStr ? parseInt(daysStr, 10) || 7 : 7;
            const full =
              args.includes("--full") || args.includes("--details");
            await calEvents(days, full);
            break;
          }
          case "today": {
            const full =
              args.includes("--full") || args.includes("--details");
            await calEvents(1, full);
            break;
          }
          case "tomorrow": {
            const full =
              args.includes("--full") || args.includes("--details");
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
          console.error("Usage: office-cli.ts people <name>");
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
