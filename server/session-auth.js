import { createHmac, timingSafeEqual } from 'node:crypto';

const SESSION_SECONDS = 60 * 60 * 12;

function signature(value, secret) {
  return createHmac('sha256', secret).update(value).digest('base64url');
}

function equalSecrets(left, right) {
  const leftBuffer = Buffer.from(left);
  const rightBuffer = Buffer.from(right);
  return leftBuffer.length === rightBuffer.length && timingSafeEqual(leftBuffer, rightBuffer);
}

export function createSessionAuth({ enabled, password, secret }) {
  if (enabled && (!password || !secret)) {
    throw new Error('Production requires APP_PASSWORD and SESSION_SECRET before the client API can start.');
  }

  function readSession(request) {
    const cookie = request.headers.cookie?.split(';').map((part) => part.trim()).find((part) => part.startsWith('casework_session='));
    if (!cookie) return false;
    const token = cookie.slice('casework_session='.length);
    const separator = token.lastIndexOf('.');
    if (separator < 1) return false;
    const payload = token.slice(0, separator);
    const receivedSignature = token.slice(separator + 1);
    if (!equalSecrets(receivedSignature, signature(payload, secret))) return false;
    try {
      return JSON.parse(Buffer.from(payload, 'base64url').toString()).expiresAt > Date.now();
    } catch {
      return false;
    }
  }

  function createCookie() {
    const payload = Buffer.from(JSON.stringify({ expiresAt: Date.now() + SESSION_SECONDS * 1000 })).toString('base64url');
    const token = `${payload}.${signature(payload, secret)}`;
    const secure = enabled ? ' Secure;' : '';
    return `casework_session=${token}; HttpOnly; SameSite=Strict; Path=/; Max-Age=${SESSION_SECONDS};${secure}`;
  }

  return {
    enabled,
    isAuthenticated: enabled ? readSession : () => true,
    passwordMatches(candidate) {
      return equalSecrets(String(candidate || ''), String(password || ''));
    },
    setCookie: createCookie,
    clearCookie() {
      return `casework_session=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0;${enabled ? ' Secure;' : ''}`;
    },
  };
}
