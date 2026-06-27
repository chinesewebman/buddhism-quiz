// Cloudflare Pages Function — HTTP Basic Auth middleware
// Set 2 env variables in Pages dashboard (Production):
//   BASIC_AUTH_USER       — the username
//   BASIC_AUTH_PASS_HASH  — SHA-256(salt + password) hex
// Both AUTH_USER and AUTH_PASS_HASH must be set; if either missing, returns 503
// (fail-closed) so the site is never accidentally public.

const SALT = "diamond-sutra-2026"; // baked into source — combined with env-supplied
                                  // password hash so a leaked repo alone is not enough
                                  // to forge a valid credential. For higher security,
                                  // move SALT to env too and rotate periodically.

async function sha256Hex(input: string): Promise<string> {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(input));
  return Array.from(new Uint8Array(buf))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

function unauthorized(reason: string): Response {
  return new Response(reason, {
    status: 401,
    headers: {
      "WWW-Authenticate": 'Basic realm="Diamond Sutra Quiz", charset="UTF-8"',
      "Content-Type": "text/plain; charset=utf-8",
      "Cache-Control": "no-store",
    },
  });
}

export const onRequest = async (context: {
  request: Request;
  env: { BASIC_AUTH_USER?: string; BASIC_AUTH_PASS_HASH?: string };
  next: () => Promise<Response>;
}): Promise<Response> => {
  const { request, env, next } = context;

  const expectedUser = env.BASIC_AUTH_USER;
  const expectedHash = env.BASIC_AUTH_PASS_HASH;

  // Fail-closed: if env not configured, refuse all access.
  if (!expectedUser || !expectedHash) {
    return new Response("Auth not configured on server.", { status: 503 });
  }

  const auth = request.headers.get("Authorization") || "";
  if (!auth.startsWith("Basic ")) {
    return unauthorized("Authentication required.");
  }

  // Decode credentials
  let decoded: string;
  try {
    decoded = atob(auth.slice(6));
  } catch {
    return unauthorized("Invalid auth header.");
  }
  const colonIdx = decoded.indexOf(":");
  if (colonIdx < 0) {
    return unauthorized("Malformed credentials.");
  }
  const user = decoded.slice(0, colonIdx);
  const pass = decoded.slice(colonIdx + 1);

  if (user !== expectedUser) {
    return unauthorized("Invalid credentials.");
  }

  // Constant-time-ish comparison via SHA-256
  const candidateHash = await sha256Hex(SALT + pass);
  if (candidateHash !== expectedHash) {
    return unauthorized("Invalid credentials.");
  }

  // Authenticated — set a marker header (useful for log inspection) and pass through
  const resp = await next();
  const newHeaders = new Headers(resp.headers);
  newHeaders.set("X-Auth-User", user);
  return new Response(resp.body, {
    status: resp.status,
    statusText: resp.statusText,
    headers: newHeaders,
  });
};
