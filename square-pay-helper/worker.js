/**
 * Hot Diggity Dog – Square "Pay now" helper (Cloudflare Worker)
 *
 * The website sends the order here. This helper asks Square for a secure
 * checkout page for that exact order and sends the customer to it. After they
 * pay, the order shows up in the Square app as a PICKUP order.
 *
 * Settings to add in Cloudflare (Worker > Settings > Variables and Secrets):
 *   SQUARE_ACCESS_TOKEN  (type: Secret)  - Production access token from developer.squareup.com
 *   SQUARE_LOCATION_ID   (type: Text)    - the truck's Location ID from the same dashboard
 *   SITE_URL             (type: Text)    - the website address, e.g. https://www.hotdiggitydogtruck.com/
 *
 * If prices change on the website, update BASE_PRICES below to match.
 * Customers can never be charged LESS than these base prices.
 */
const SQUARE_VERSION = "2026-09-16";

const BASE_PRICES = {
  "Firecracker": 7,
  "South of the Border": 7,
  "Fiery Diggity": 7,
  "Sweet & Spicy": 7,
  "Spanish Dog": 7,
  "Chili Cheese Dog": 7,
  "Cali Dog": 7,
  "Everything Dog": 7,
  "Taco Dog": 7,
  "German Dog": 7,
  "Chicago Dog": 7,
  "BLT Dog": 7,
  "Street Dog": 7,
  "Any Hot Dog Combo": 10,
  "Crispy Chicken Sandwich Combo": 11,
  "BBQ Sandwich Combo": 11,
  "Mozzarella Sticks Combo": 12,
  "Any Dog + Lemon Shake-Up": 12,
  "Boneless Wings (6 pc)": 7,
  "Bone-In Wings (5 pc)": 10,
  "Chicken Drumsticks (3 pc)": 7,
  "Chicken Drumsticks (5 pc)": 9,
  "Chicken Tenders (3 pc)": 7,
  "Chicken Tenders (5 pc)": 9,
  "BBQ Pulled Pork or Chicken Sandwich": 8,
  "Crispy Chicken Sandwich": 8,
  "Corn Dog": 7,
  "Pork Parfait": 9,
  "Loaded Nachos": 10,
  "Regular Nachos": 7,
  "Chili Cheese Fries": 7,
  "Loaded Fries": 10,
  "Fries": 4,
  "Chips": 2,
  "Coleslaw": 2,
  "Mozzarella Sticks (6 pc)": 7,
  "Side of Cheese Sauce": 2,
  "Side of Chili Sauce": 2,
  "Drink": 2,
  "Lemon Shake-Up (24 oz)": 7,
  "Lemon Shake-Up Refill": 4,
  "Build Your Own Dog": 7,
  "Peanut Sauce Chicken": 10
};

export default {
  async fetch(req, env) {
    const origin = (env.SITE_URL || "*").replace(/^(https?:\/\/[^/]+).*$/, "$1");
    const cors = {
      "Access-Control-Allow-Origin": origin,
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    };
    if (req.method === "OPTIONS") return new Response(null, { headers: cors });
    if (req.method !== "POST") return json({ ok: true, message: "Hot Diggity Dog pay helper is running." }, 200, cors);

    try {
      if (!env.SQUARE_ACCESS_TOKEN || !env.SQUARE_LOCATION_ID) throw new Error("Helper is missing its Square settings.");
      const d = await req.json();
      const lines = Array.isArray(d.lines) ? d.lines.slice(0, 50) : [];
      if (!lines.length) throw new Error("The order is empty.");

      const line_items = lines.map((l) => {
        const base = BASE_PRICES[l.name];
        if (base == null) throw new Error("Unknown menu item: " + l.name);
        const unit = Math.max(Number(l.unit) || 0, base);          // never below the menu price
        const qty = Math.min(Math.max(parseInt(l.qty, 10) || 1, 1), 50);
        const item = {
          name: String(l.name).slice(0, 200),
          quantity: String(qty),
          base_price_money: { amount: Math.round(unit * 100), currency: "USD" },
        };
        if (l.detail) item.note = String(l.detail).slice(0, 500);
        return item;
      });

      const digits = String(d.phone || "").replace(/\D/g, "").slice(-10);
      const phone = digits.length === 10 ? "+1" + digits : undefined;
      const num = String(d.num || "").slice(0, 40);
      const site = env.SITE_URL || "";
      const custName = String(d.name || "Customer").slice(0, 60);
      const label = ("#" + num + " " + custName).slice(0, 30);   // what shows on the tablet
      const customerId = await findOrCreateCustomer(env, custName, phone).catch(() => null);

      const body = {
        idempotency_key: crypto.randomUUID(),
        order: {
          location_id: env.SQUARE_LOCATION_ID,
          reference_id: num,
          ticket_name: label,
          ...(customerId ? { customer_id: customerId } : {}),
          line_items,
          fulfillments: [{
            type: "PICKUP",
            state: "PROPOSED",
            pickup_details: {
              recipient: { display_name: label, ...(customerId ? { customer_id: customerId } : {}), ...(phone ? { phone_number: phone } : {}) },
              schedule_type: "ASAP",
              note: ("Website order #" + num + (d.notes ? " · " + d.notes : "")).slice(0, 500),
            },
          }],
        },
        checkout_options: site ? { redirect_url: site + (site.includes("?") ? "&" : "?") + "paid=" + encodeURIComponent(num) } : undefined,
        pre_populated_data: phone ? { buyer_phone_number: phone } : undefined,
        payment_note: "Hot Diggity Dog website order #" + num,
      };

      const r = await fetch("https://connect.squareup.com/v2/online-checkout/payment-links", {
        method: "POST",
        headers: {
          "Authorization": "Bearer " + env.SQUARE_ACCESS_TOKEN,
          "Square-Version": SQUARE_VERSION,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
      });
      const j = await r.json();
      if (!r.ok || !j.payment_link) {
        const msg = (j.errors && j.errors[0] && (j.errors[0].detail || j.errors[0].code)) || "Square error";
        return json({ ok: false, error: msg }, 502, cors);
      }
      return json({ ok: true, url: j.payment_link.url }, 200, cors);
    } catch (e) {
      return json({ ok: false, error: String((e && e.message) || e) }, 400, cors);
    }
  },
};

// Finds the customer in his Square customer list by phone, or adds them,
// so the order shows the customer's name on the tablet.
async function findOrCreateCustomer(env, name, phone) {
  const h = {
    "Authorization": "Bearer " + env.SQUARE_ACCESS_TOKEN,
    "Square-Version": SQUARE_VERSION,
    "Content-Type": "application/json",
  };
  if (phone) {
    const sr = await fetch("https://connect.squareup.com/v2/customers/search", {
      method: "POST", headers: h,
      body: JSON.stringify({ limit: 1, query: { filter: { phone_number: { exact: phone } } } }),
    });
    const sj = await sr.json();
    if (sj.customers && sj.customers[0]) return sj.customers[0].id;
  }
  const [given, ...rest] = name.trim().split(/\s+/);
  const cr = await fetch("https://connect.squareup.com/v2/customers", {
    method: "POST", headers: h,
    body: JSON.stringify({
      idempotency_key: crypto.randomUUID(),
      given_name: given || "Customer",
      ...(rest.length ? { family_name: rest.join(" ") } : {}),
      ...(phone ? { phone_number: phone } : {}),
      note: "Added from hotdiggitydogtruck.com order",
    }),
  });
  const cj = await cr.json();
  return cj.customer ? cj.customer.id : null;
}

function json(o, status, headers) {
  return new Response(JSON.stringify(o), { status, headers: { ...headers, "Content-Type": "application/json" } });
}
