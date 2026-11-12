import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { LikeServiceError } from './ILikeService';

/** Escape a value for use inside an OData single-quoted string literal. */
export function odataString(value: string): string {
  return encodeURIComponent(value.replace(/'/g, "''"));
}

/** Pull the message out of a SharePoint OData error body (v3 or v4 JSON). */
function extractReason(body: string): string {
  try {
    const json = JSON.parse(body);
    const err = json['odata.error'] || json.error;
    const msg = err?.message;
    return (typeof msg === 'string' ? msg : msg?.value) || '';
  } catch {
    return '';
  }
}

async function ensureOk(response: SPHttpClientResponse, action: string): Promise<SPHttpClientResponse> {
  if (!response.ok) {
    let detail = '';
    try {
      detail = await response.text();
    } catch {
      // ignore
    }
    throw new LikeServiceError(
      `${action} failed (${response.status}) ${detail}`.trim(),
      response.status,
      extractReason(detail) || response.statusText || ''
    );
  }
  return response;
}

export async function getJson<T>(client: SPHttpClient, url: string): Promise<T> {
  const response = await client.get(url, SPHttpClient.configurations.v1);
  await ensureOk(response, `GET ${url}`);
  return (await response.json()) as T;
}

export async function post(
  client: SPHttpClient,
  url: string,
  body?: unknown,
  extraHeaders?: Record<string, string>
): Promise<SPHttpClientResponse> {
  const response = await client.post(url, SPHttpClient.configurations.v1, {
    headers: { 'Content-Type': 'application/json', ...extraHeaders },
    body: body === undefined ? undefined : JSON.stringify(body)
  });
  return ensureOk(response, `POST ${url}`);
}
