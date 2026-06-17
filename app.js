"use strict";

const STORAGE_KEY = "invoice-generator-v1";

// ---- State ----
const FIELDS = [
  "fromName", "fromAddress", "fromEmail",
  "toName", "toAddress", "toEmail",
  "invoiceNumber", "date", "dueDate",
  "taxRate", "currency", "notes"
];

let state = {
  fields: {},
  items: [] // { description, qty, price }
};

// ---- DOM refs ----
const form = document.getElementById("invoiceForm");
const itemsBody = document.getElementById("itemsBody");
const preview = document.getElementById("preview");
const saveStatus = document.getElementById("saveStatus");

// ---- Helpers ----
function money(n, currency) {
  const v = Number.isFinite(n) ? n : 0;
  return currency + v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function num(v) {
  const n = parseFloat(v);
  return Number.isFinite(n) ? n : 0;
}

function lineAmount(item) {
  return num(item.qty) * num(item.price);
}

function computeTotals() {
  const subtotal = state.items.reduce((sum, it) => sum + lineAmount(it), 0);
  const taxRate = num(state.fields.taxRate);
  const tax = subtotal * (taxRate / 100);
  return { subtotal, tax, total: subtotal + tax, taxRate };
}

function escapeHtml(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

// ---- Persistence ----
let saveTimer = null;
function save() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      saveStatus.textContent = "Saved";
      setTimeout(() => { saveStatus.textContent = ""; }, 1200);
    } catch (e) {
      saveStatus.textContent = "Save failed";
    }
  }, 250);
}

function load() {
  let saved = null;
  try {
    saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
  } catch (e) { saved = null; }

  if (saved && typeof saved === "object") {
    state.fields = saved.fields || {};
    state.items = Array.isArray(saved.items) ? saved.items : [];
  }

  // Defaults on first run.
  if (!state.fields.currency) state.fields.currency = "$";
  if (state.fields.taxRate == null) state.fields.taxRate = 0;
  if (!state.fields.date) state.fields.date = new Date().toISOString().slice(0, 10);
  if (state.items.length === 0) {
    state.items.push({ description: "", qty: 1, price: 0 });
  }
}

// ---- Render: form inputs from state ----
function hydrateFields() {
  FIELDS.forEach(name => {
    const el = form.querySelector(`[data-field="${name}"]`);
    if (el && state.fields[name] != null) el.value = state.fields[name];
  });
}

// ---- Render: line items ----
function renderItems() {
  itemsBody.innerHTML = "";
  const currency = state.fields.currency || "$";

  state.items.forEach((item, i) => {
    const tr = document.createElement("tr");

    const tdDesc = document.createElement("td");
    tdDesc.className = "desc";
    const desc = document.createElement("input");
    desc.type = "text";
    desc.value = item.description || "";
    desc.placeholder = "Item description";
    desc.addEventListener("input", () => { state.items[i].description = desc.value; onItemChange(); });
    tdDesc.appendChild(desc);

    const tdQty = document.createElement("td");
    tdQty.className = "num";
    const qty = document.createElement("input");
    qty.type = "number"; qty.min = "0"; qty.step = "1";
    qty.value = item.qty;
    qty.addEventListener("input", () => { state.items[i].qty = qty.value; onItemChange(); });
    tdQty.appendChild(qty);

    const tdPrice = document.createElement("td");
    tdPrice.className = "num";
    const price = document.createElement("input");
    price.type = "number"; price.min = "0"; price.step = "0.01";
    price.value = item.price;
    price.addEventListener("input", () => { state.items[i].price = price.value; onItemChange(); });
    tdPrice.appendChild(price);

    const tdAmount = document.createElement("td");
    tdAmount.className = "num amount-cell";
    tdAmount.textContent = money(lineAmount(item), currency);

    const tdRm = document.createElement("td");
    tdRm.className = "rm";
    const rm = document.createElement("button");
    rm.type = "button";
    rm.className = "rm-btn";
    rm.textContent = "×";
    rm.title = "Remove line";
    rm.addEventListener("click", () => {
      state.items.splice(i, 1);
      if (state.items.length === 0) state.items.push({ description: "", qty: 1, price: 0 });
      renderItems(); renderPreview(); save();
    });
    tdRm.appendChild(rm);

    tr.append(tdDesc, tdQty, tdPrice, tdAmount, tdRm);
    itemsBody.appendChild(tr);
  });
}

function onItemChange() {
  // Update only the amount cells + preview without full re-render (keeps focus).
  const currency = state.fields.currency || "$";
  itemsBody.querySelectorAll(".amount-cell").forEach((cell, i) => {
    cell.textContent = money(lineAmount(state.items[i]), currency);
  });
  renderPreview();
  save();
}

// ---- Render: preview ----
function renderPreview() {
  const f = state.fields;
  const currency = f.currency || "$";
  const { subtotal, tax, total, taxRate } = computeTotals();

  const hasContent = f.fromName || f.toName || state.items.some(it => it.description);
  if (!hasContent) {
    preview.innerHTML = '<p class="pv-empty">Fill in the form to see a live preview of your invoice.</p>';
    return;
  }

  const itemRows = state.items
    .filter(it => it.description || num(it.qty) || num(it.price))
    .map(it => `
      <tr>
        <td>${escapeHtml(it.description)}</td>
        <td class="r">${num(it.qty)}</td>
        <td class="r">${money(num(it.price), currency)}</td>
        <td class="r">${money(lineAmount(it), currency)}</td>
      </tr>`).join("");

  preview.innerHTML = `
    <div class="pv-head">
      <div>
        <p class="pv-title">INVOICE</p>
        <div class="who">${escapeHtml(f.fromName) || "&nbsp;"}</div>
        <div class="addr">${escapeHtml(f.fromAddress)}</div>
        <div class="addr">${escapeHtml(f.fromEmail)}</div>
      </div>
      <div class="pv-meta">
        <div><strong>#</strong> ${escapeHtml(f.invoiceNumber) || "—"}</div>
        <div><strong>Date:</strong> ${escapeHtml(f.date) || "—"}</div>
        <div><strong>Due:</strong> ${escapeHtml(f.dueDate) || "—"}</div>
      </div>
    </div>

    <div class="pv-parties">
      <div>
        <h3>Bill To</h3>
        <div class="who">${escapeHtml(f.toName) || "—"}</div>
        <div class="addr">${escapeHtml(f.toAddress)}</div>
        <div class="addr">${escapeHtml(f.toEmail)}</div>
      </div>
    </div>

    <table class="pv-items">
      <thead>
        <tr><th>Description</th><th class="r">Qty</th><th class="r">Unit</th><th class="r">Amount</th></tr>
      </thead>
      <tbody>${itemRows || '<tr><td colspan="4" class="pv-empty">No line items yet.</td></tr>'}</tbody>
    </table>

    <div class="pv-totals">
      <div class="row"><span>Subtotal</span><span>${money(subtotal, currency)}</span></div>
      <div class="row"><span>Tax (${taxRate}%)</span><span>${money(tax, currency)}</span></div>
      <div class="row grand"><span>Total</span><span>${money(total, currency)}</span></div>
    </div>

    ${f.notes ? `<div class="pv-notes"><strong>Notes:</strong>\n${escapeHtml(f.notes)}</div>` : ""}
  `;
}

// ---- Wire up form field inputs ----
form.addEventListener("input", (e) => {
  const el = e.target.closest("[data-field]");
  if (!el) return;
  const name = el.getAttribute("data-field");
  state.fields[name] = el.value;
  if (name === "currency" || name === "taxRate") renderItems();
  renderPreview();
  save();
});

document.getElementById("addItem").addEventListener("click", () => {
  state.items.push({ description: "", qty: 1, price: 0 });
  renderItems(); renderPreview(); save();
});

document.getElementById("clearAll").addEventListener("click", () => {
  if (!confirm("Clear all invoice data? This cannot be undone.")) return;
  localStorage.removeItem(STORAGE_KEY);
  state = { fields: {}, items: [] };
  load();
  hydrateFields();
  renderItems();
  renderPreview();
});

// ---- Init ----
load();
hydrateFields();
renderItems();
renderPreview();

// Export buttons (handlers defined in export.js).
document.getElementById("exportPdf").addEventListener("click", () => window.exportPdf(state, computeTotals()));
document.getElementById("exportWord").addEventListener("click", () => window.exportWord(state, computeTotals()));
