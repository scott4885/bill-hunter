/**
 * OAuth token manager for Microsoft Graph (Delegated permissions).
 *
 * First-time auth happens via scripts/auth-setup.ts (browser flow, one-time).
 * After that, this module handles silent refresh using the refresh_token.
 *
 * Tokens are persisted to .env locally and to Azure Function App settings in prod.
 */

import { promises as fs } from 'fs';
import path from 'path';

const TOKEN_ENDPOINT = (tenantId: string) =>
  `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

const GRAPH_SCOPES = [
  'Mail.Read',
  'Mail.ReadBasic',
  'Files.ReadWrite',
  'offline_access',
  'User.Read',
].join(' ');

export interface TokenSet {
  accessToken: string;
  refreshToken: string;
  expiresAt: number; // epoch ms
}

let cached: TokenSet | null = null;

export async function getAccessToken(): Promise<string> {
  if (cached && cached.expiresAt > Date.now() + 60_000) {
    return cached.accessToken;
  }

  const refreshToken = process.env.GRAPH_REFRESH_TOKEN;
  if (!refreshToken) {
    throw new Error(
      'No refresh token. Run `npm run auth:setup` for first-time authentication.',
    );
  }

  cached = await refreshAccessToken(refreshToken);
  await persistTokens(cached);
  return cached.accessToken;
}

export async function refreshAccessToken(refreshToken: string): Promise<TokenSet> {
  const tenantId = required('AZURE_TENANT_ID');
  const clientId = required('AZURE_CLIENT_ID');
  const clientSecret = required('AZURE_CLIENT_SECRET');

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    refresh_token: refreshToken,
    grant_type: 'refresh_token',
    scope: GRAPH_SCOPES,
  });

  const res = await fetch(TOKEN_ENDPOINT(tenantId), {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Token refresh failed: ${res.status} ${text}`);
  }

  const data = (await res.json()) as {
    access_token: string;
    refresh_token: string;
    expires_in: number;
  };

  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token, // MS rotates these; always use the latest
    expiresAt: Date.now() + data.expires_in * 1000,
  };
}

export async function exchangeCodeForTokens(
  code: string,
  redirectUri: string,
): Promise<TokenSet> {
  const tenantId = required('AZURE_TENANT_ID');
  const clientId = required('AZURE_CLIENT_ID');
  const clientSecret = required('AZURE_CLIENT_SECRET');

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    code,
    redirect_uri: redirectUri,
    grant_type: 'authorization_code',
    scope: GRAPH_SCOPES,
  });

  const res = await fetch(TOKEN_ENDPOINT(tenantId), {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });

  if (!res.ok) {
    throw new Error(`Code exchange failed: ${res.status} ${await res.text()}`);
  }

  const data = (await res.json()) as {
    access_token: string;
    refresh_token: string;
    expires_in: number;
  };

  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresAt: Date.now() + data.expires_in * 1000,
  };
}

async function persistTokens(tokens: TokenSet): Promise<void> {
  // In Azure Functions, write back to App Settings via the management API.
  // Locally, update .env file.
  if (process.env.AZURE_FUNCTIONS_ENVIRONMENT) {
    // Production path - implement via Azure REST or Key Vault
    // For now, rely on env vars being managed externally
    return;
  }

  const envPath = path.join(process.cwd(), '.env');
  try {
    let content = await fs.readFile(envPath, 'utf-8');
    content = upsertEnvVar(content, 'GRAPH_ACCESS_TOKEN', tokens.accessToken);
    content = upsertEnvVar(content, 'GRAPH_REFRESH_TOKEN', tokens.refreshToken);
    content = upsertEnvVar(content, 'TOKEN_EXPIRES_AT', String(tokens.expiresAt));
    await fs.writeFile(envPath, content);
  } catch {
    // .env doesn't exist yet during first-time setup
  }
}

function upsertEnvVar(content: string, key: string, value: string): string {
  const line = `${key}=${value}`;
  const re = new RegExp(`^${key}=.*$`, 'm');
  return re.test(content) ? content.replace(re, line) : `${content.trimEnd()}\n${line}\n`;
}

function required(name: string): string {
  const v = process.env[name];
  if (!v) throw new Error(`Missing env var: ${name}`);
  return v;
}

export function getAuthorizeUrl(redirectUri: string, state: string): string {
  const tenantId = required('AZURE_TENANT_ID');
  const clientId = required('AZURE_CLIENT_ID');
  const params = new URLSearchParams({
    client_id: clientId,
    response_type: 'code',
    redirect_uri: redirectUri,
    response_mode: 'query',
    scope: GRAPH_SCOPES,
    state,
    prompt: 'consent',
  });
  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${params}`;
}
