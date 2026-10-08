// Sending sign-in links. SMTP_URL (smtp://user:pass@host:587) with MAIL_FROM, or RESEND_API_KEY with
// MAIL_FROM (or EMAIL_FROM). Without either, links are written to the server log.
export async function createMailer(env = process.env) {
  const from = env.MAIL_FROM || env.EMAIL_FROM || 'Sheets <sheets@localhost>';
  if (env.SMTP_URL) {
    const { default: nodemailer } = await import('nodemailer');
    const t = nodemailer.createTransport(env.SMTP_URL);
    return { ready: true, async send(m) { await t.sendMail({ from, ...m }); } };
  }
  if (env.RESEND_API_KEY) {
    return {
      ready: true,
      async send(m) {
        const r = await fetch('https://api.resend.com/emails', { method: 'POST', headers: { authorization: `Bearer ${env.RESEND_API_KEY}`, 'content-type': 'application/json' }, body: JSON.stringify({ from, to: [m.to], subject: m.subject, text: m.text }) });
        if (!r.ok) console.error('sheets: email not sent:', r.status, (await r.text()).slice(0, 200));
      },
    };
  }
  return { ready: false, sent: [], async send(m) { this.sent.push(m); console.log(`sheets: no SMTP_URL or RESEND_API_KEY, so here is the email to ${m.to}:\n${m.text}`); } };
}
