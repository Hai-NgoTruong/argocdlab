"use strict";

// Shared formatting helpers (mirrors app.js money/num behavior).
function _num(v) { const n = parseFloat(v); return Number.isFinite(n) ? n : 0; }
function _money(n, c) {
  return c + (Number.isFinite(n) ? n : 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function _filename(state, ext) {
  const inv = (state.fields.invoiceNumber || "invoice").replace(/[^a-z0-9\-_]+/gi, "-");
  return `${inv}.${ext}`;
}
function _visibleItems(state) {
  return state.items.filter(it => it.description || _num(it.qty) || _num(it.price));
}

// =================== PDF (jsPDF + autotable) ===================
window.exportPdf = function (state, totals) {
  if (!window.jspdf || !window.jspdf.jsPDF) {
    alert("PDF library is still loading. Try again in a moment.");
    return;
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const f = state.fields;
  const c = f.currency || "$";
  const margin = 48;
  let y = margin;

  doc.setFontSize(26);
  doc.setFont("helvetica", "bold");
  doc.text("INVOICE", margin, y);

  // Invoice meta (right aligned)
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  const right = doc.internal.pageSize.getWidth() - margin;
  doc.text(`#  ${f.invoiceNumber || "—"}`, right, y - 14, { align: "right" });
  doc.text(`Date:  ${f.date || "—"}`, right, y, { align: "right" });
  doc.text(`Due:   ${f.dueDate || "—"}`, right, y + 14, { align: "right" });

  y += 36;

  // From / Bill To
  doc.setFont("helvetica", "bold"); doc.setFontSize(9);
  doc.text("FROM", margin, y);
  doc.text("BILL TO", right - 200, y);
  doc.setFont("helvetica", "normal"); doc.setFontSize(10);

  const fromLines = [f.fromName, f.fromAddress, f.fromEmail].filter(Boolean).join("\n").split("\n");
  const toLines = [f.toName, f.toAddress, f.toEmail].filter(Boolean).join("\n").split("\n");
  doc.text(fromLines.length ? fromLines : ["—"], margin, y + 14);
  doc.text(toLines.length ? toLines : ["—"], right - 200, y + 14);

  y += 14 + Math.max(fromLines.length, toLines.length) * 13 + 18;

  // Items table
  const body = _visibleItems(state).map(it => [
    it.description || "",
    String(_num(it.qty)),
    _money(_num(it.price), c),
    _money(_num(it.qty) * _num(it.price), c)
  ]);

  doc.autoTable({
    startY: y,
    head: [["Description", "Qty", "Unit price", "Amount"]],
    body: body.length ? body : [["No line items", "", "", ""]],
    margin: { left: margin, right: margin },
    headStyles: { fillColor: [37, 99, 235], halign: "left" },
    columnStyles: { 1: { halign: "right" }, 2: { halign: "right" }, 3: { halign: "right" } },
    styles: { fontSize: 10, cellPadding: 6 }
  });

  // Totals
  let ty = doc.lastAutoTable.finalY + 18;
  const labelX = right - 200;
  doc.setFontSize(10); doc.setFont("helvetica", "normal");
  doc.text("Subtotal", labelX, ty); doc.text(_money(totals.subtotal, c), right, ty, { align: "right" });
  ty += 16;
  doc.text(`Tax (${totals.taxRate}%)`, labelX, ty); doc.text(_money(totals.tax, c), right, ty, { align: "right" });
  ty += 6;
  doc.setLineWidth(1); doc.line(labelX, ty + 4, right, ty + 4);
  ty += 22;
  doc.setFont("helvetica", "bold"); doc.setFontSize(13);
  doc.text("Total", labelX, ty); doc.text(_money(totals.total, c), right, ty, { align: "right" });

  // Notes
  if (f.notes) {
    ty += 32;
    doc.setFont("helvetica", "bold"); doc.setFontSize(9);
    doc.text("NOTES", margin, ty);
    doc.setFont("helvetica", "normal"); doc.setFontSize(10);
    doc.text(doc.splitTextToSize(f.notes, right - margin), margin, ty + 14);
  }

  doc.save(_filename(state, "pdf"));
};

// =================== Word (docx) ===================
window.exportWord = function (state, totals) {
  if (!window.docx) {
    alert("Word library is still loading. Try again in a moment.");
    return;
  }
  const D = window.docx;
  const f = state.fields;
  const c = f.currency || "$";

  const para = (text, opts) => new D.Paragraph({ children: [new D.TextRun(Object.assign({ text: text || "" }, opts))] });

  function cell(text, opts = {}) {
    return new D.TableCell({
      children: [new D.Paragraph({
        alignment: opts.align || D.AlignmentType.LEFT,
        children: [new D.TextRun({ text: text == null ? "" : String(text), bold: !!opts.bold, color: opts.color })]
      })],
      shading: opts.fill ? { fill: opts.fill, type: D.ShadingType.CLEAR, color: "auto" } : undefined
    });
  }

  const headerRow = new D.TableRow({
    children: [
      cell("Description", { bold: true, color: "FFFFFF", fill: "2563EB" }),
      cell("Qty", { bold: true, color: "FFFFFF", fill: "2563EB", align: D.AlignmentType.RIGHT }),
      cell("Unit price", { bold: true, color: "FFFFFF", fill: "2563EB", align: D.AlignmentType.RIGHT }),
      cell("Amount", { bold: true, color: "FFFFFF", fill: "2563EB", align: D.AlignmentType.RIGHT })
    ]
  });

  const itemRows = _visibleItems(state).map(it => new D.TableRow({
    children: [
      cell(it.description),
      cell(String(_num(it.qty)), { align: D.AlignmentType.RIGHT }),
      cell(_money(_num(it.price), c), { align: D.AlignmentType.RIGHT }),
      cell(_money(_num(it.qty) * _num(it.price), c), { align: D.AlignmentType.RIGHT })
    ]
  }));

  const itemsTable = new D.Table({
    width: { size: 100, type: D.WidthType.PERCENTAGE },
    rows: [headerRow, ...(itemRows.length ? itemRows : [new D.TableRow({ children: [cell("No line items"), cell(""), cell(""), cell("")] })])]
  });

  function totalsLine(label, value, bold) {
    return new D.Paragraph({
      alignment: D.AlignmentType.RIGHT,
      children: [new D.TextRun({ text: `${label}:  ${value}`, bold: !!bold, size: bold ? 26 : 22 })]
    });
  }

  const children = [
    new D.Paragraph({ children: [new D.TextRun({ text: "INVOICE", bold: true, size: 48 })] }),
    para(""),
    para(`Invoice #: ${f.invoiceNumber || "—"}`),
    para(`Date: ${f.date || "—"}    Due: ${f.dueDate || "—"}`),
    para(""),
    new D.Paragraph({ children: [new D.TextRun({ text: "FROM", bold: true })] }),
    ...[f.fromName, f.fromAddress, f.fromEmail].filter(Boolean).flatMap(s => s.split("\n")).map(s => para(s)),
    para(""),
    new D.Paragraph({ children: [new D.TextRun({ text: "BILL TO", bold: true })] }),
    ...[f.toName, f.toAddress, f.toEmail].filter(Boolean).flatMap(s => s.split("\n")).map(s => para(s)),
    para(""),
    itemsTable,
    para(""),
    totalsLine("Subtotal", _money(totals.subtotal, c)),
    totalsLine(`Tax (${totals.taxRate}%)`, _money(totals.tax, c)),
    totalsLine("Total", _money(totals.total, c), true)
  ];

  if (f.notes) {
    children.push(para(""));
    children.push(new D.Paragraph({ children: [new D.TextRun({ text: "Notes", bold: true })] }));
    f.notes.split("\n").forEach(line => children.push(para(line)));
  }

  const docFile = new D.Document({ sections: [{ children }] });

  D.Packer.toBlob(docFile).then(blob => {
    if (window.saveAs) window.saveAs(blob, _filename(state, "docx"));
    else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = _filename(state, "docx");
      a.click();
      URL.revokeObjectURL(url);
    }
  });
};
