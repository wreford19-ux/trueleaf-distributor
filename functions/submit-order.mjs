import { Resend } from "resend";

const esc = (x) => String(x ?? "").replace(/[<>&]/g, c => ({ "<":"&lt;", ">":"&gt;", "&":"&amp;" }[c]));

export default async (req) => {
  if (req.method !== "POST") return json({ ok:false, error:"Method not allowed" }, 405);

  let o;
  try { o = await req.json(); } catch { return json({ ok:false, error:"Bad request" }, 400); }

  const key = process.env.RESEND_API_KEY;
  if (!key) return json({ ok:false, error:"RESEND_API_KEY not set" }, 500);

  // FROM_EMAIL defaults to Resend's test sender (delivers only to the account owner).
  const FROM = process.env.FROM_EMAIL || "onboarding@resend.dev";
  // Internal recipients. In test-sender mode this MUST be the Resend account owner's address.
  const NOTIFY = (process.env.ORDER_NOTIFY_TO || "wreford19@gmail.com").split(",").map(s => s.trim()).filter(Boolean);
  const isTestSender = /resend\.dev/i.test(FROM);

  const resend = new Resend(key);
  const a = o.addr || {};
  const addrLine = [a.addr1, a.suburb, a.city, a.province, a.postal].filter(Boolean).map(esc).join(", ");
  const rows = (o.lines || []).map(l =>
    `<tr><td style="padding:4px 10px;font-family:monospace;color:#5c3d1e">${esc(l.code)}</td>
         <td style="padding:4px 10px">${esc(l.name)}</td>
         <td style="padding:4px 10px;text-align:right;font-family:monospace">${esc(l.qty)}</td></tr>`).join("");
  const table = `<table style="border-collapse:collapse;width:100%;margin-top:8px;font-size:14px">
    <thead><tr style="background:#f0e8cd;color:#5c3d1e">
      <th style="text-align:left;padding:6px 10px">Code</th>
      <th style="text-align:left;padding:6px 10px">Seed</th>
      <th style="text-align:right;padding:6px 10px">Qty</th></tr></thead>
    <tbody>${rows}</tbody></table>`;

  const internal = `
    <div style="font-family:system-ui,sans-serif;color:#2a2015;max-width:600px">
      <h2 style="color:#2d4a1e;margin:0 0 4px">New distributor order — ${esc(o.ref)}</h2>
      <p style="margin:0 0 12px;color:#8a7a5a">${esc(o.business)} · ${o.lineCount} lines · ${o.totalUnits} units</p>
      <p style="margin:0 0 4px"><strong>${esc(o.contact)}</strong> · ${esc(o.phone)} · ${esc(o.email)}</p>
      <p style="margin:0 0 12px">Deliver to: ${addrLine}</p>
      ${table}
    </div>`;

  const distributor = `
    <div style="font-family:system-ui,sans-serif;color:#2a2015;max-width:600px">
      <h2 style="color:#2d4a1e;margin:0 0 4px">Thank you, ${esc(o.business)} 🌱</h2>
      <p style="margin:0 0 12px">We've received your order <strong>${esc(o.ref)}</strong> and the TrueLeaf team is on it.
      You'll be contacted to confirm and arrange delivery.</p>
      ${table}
      <p style="margin:14px 0 0;color:#8a7a5a;font-size:13px">TrueLeaf Seed Co. · Seeds That Grow With You</p>
    </div>`;

  try {
    // 1) Internal order alert — always sent.
    await resend.emails.send({
      from: FROM,
      to: NOTIFY,
      reply_to: o.email || undefined,
      subject: `New order ${o.ref} — ${o.business} (${o.totalUnits} units)`,
      html: internal,
    });

    // 2) Distributor confirmation — only once a real (non-test) domain is verified,
    //    because the test sender cannot deliver to external addresses.
    if (!isTestSender && o.email) {
      await resend.emails.send({
        from: FROM,
        to: [o.email],
        subject: `We've received your order ${o.ref}`,
        html: distributor,
      });
    }
    return json({ ok:true, testSender:isTestSender });
  } catch (e) {
    return json({ ok:false, error:String(e && e.message || e) }, 500);
  }
};

function json(body, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { "content-type":"application/json" } });
}
