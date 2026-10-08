// Seals a secret (a CRM access token) before it is stored, with a key made from the server secret.
import crypto from 'node:crypto';

const keyFrom = (secret) => crypto.createHash('sha256').update(`wos-sheets-seal:${secret}`).digest();

export function seal(text, secret) {
  const iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv('aes-256-gcm', keyFrom(secret), iv);
  const body = Buffer.concat([c.update(String(text), 'utf8'), c.final()]);
  return `v1.${Buffer.concat([iv, c.getAuthTag(), body]).toString('base64url')}`;
}

export function unseal(sealed, secret) {
  try {
    const raw = Buffer.from(String(sealed).slice(3), 'base64url');
    const d = crypto.createDecipheriv('aes-256-gcm', keyFrom(secret), raw.subarray(0, 12));
    d.setAuthTag(raw.subarray(12, 28));
    return Buffer.concat([d.update(raw.subarray(28)), d.final()]).toString('utf8');
  } catch { return null; }
}
