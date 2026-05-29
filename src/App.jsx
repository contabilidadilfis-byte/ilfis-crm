import { useState, useMemo, useRef } from "react";
import React from "react";

// ══════════════════════════════════════════════════════════════════
//  Motor de exportación Excel — genera .xlsx nativo sin librerías
//  Usa el formato XML SpreadsheetML (compatible con Excel / LibreOffice)
// ══════════════════════════════════════════════════════════════════

function escXml(v) {
  if (v === null || v === undefined) return "";
  return String(v)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function xlsCell(value, type = "String") {
  const v = escXml(value);
  if (type === "Number") return `<Cell><Data ss:Type="Number">${v}</Data></Cell>`;
  return `<Cell><Data ss:Type="String">${v}</Data></Cell>`;
}

function xlsHeader(label) {
  return `<Cell ss:StyleID="header"><Data ss:Type="String">${escXml(label)}</Data></Cell>`;
}

function buildXls(sheets) {
  const styleSheet = `
    <Styles>
      <Style ss:ID="header">
        <Font ss:Bold="1" ss:Color="#FFFFFF" ss:Size="11"/>
        <Interior ss:Color="#CC0000" ss:Pattern="Solid"/>
        <Alignment ss:Horizontal="Center"/>
      </Style>
      <Style ss:ID="subheader">
        <Font ss:Bold="1" ss:Size="10"/>
        <Interior ss:Color="#F1F5F9" ss:Pattern="Solid"/>
      </Style>
      <Style ss:ID="currency">
        <NumberFormat ss:Format="#,##0.00"/>
      </Style>
      <Style ss:ID="date">
        <NumberFormat ss:Format="DD/MM/YYYY"/>
      </Style>
    </Styles>`;

  const worksheets = sheets.map(({ name, headers, rows }) => {
    const headerRow = `<Row>${headers.map(h => xlsHeader(h)).join("")}</Row>`;
    const dataRows = rows.map(row =>
      `<Row>${row.map(([v, t]) => xlsCell(v, t)).join("")}</Row>`
    ).join("");
    return `
      <Worksheet ss:Name="${escXml(name)}">
        <Table>
          <Column ss:Width="140"/>
          ${headerRow}
          ${dataRows}
        </Table>
      </Worksheet>`;
  }).join("");

  return `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  ${styleSheet}
  ${worksheets}
</Workbook>`;
}

function downloadXls(filename, sheets) {
  const xml = buildXls(sheets);
  const blob = new Blob([xml], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ── Exportadores por sección ──────────────────────────────────────────────────

function exportClients(clients, label = "Clientes") {
  const now = new Date().toLocaleDateString("es-CL");
  const headers = ["Empresa", "Contacto", "Teléfono", "Correo", "País", "Producto", "Valor (USD)", "Asesor", "Etapa", "Días Activo", "Alerta", "Fecha Creación"];
  const rows = clients.map(c => {
    const days = Math.floor((new Date() - new Date(c.createdAt)) / 86400000);
    const alert = (c.stage !== "Ganado" && c.stage !== "Perdido")
      ? days >= 90 ? "🔴 Expirado" : days >= 60 ? "🟡 Prioridad" : "🟢 Normal"
      : "—";
    return [
      [c.company], [c.contact], [c.phone || ""], [c.email || ""],
      [c.country || ""], [c.product || ""], [c.value, "Number"],
      [c.advisor], [c.stage], [days, "Number"], [alert], [c.createdAt],
    ];
  });
  downloadXls(`ILFIS_CRM_${label}_${now}.xls`, [{
    name: label, headers, rows
  }]);
}

function exportPayments(payments, clients) {
  const now = new Date().toLocaleDateString("es-CL");
  const headers = [
    "Empresa Cliente", "Razón Social Facturación", "RUT / NIT / Doc.",
    "Correo Facturación", "Teléfono", "Responsable", "Curso / Producto",
    "Monto Total", "Monto Pagado", "Saldo Pendiente",
    "Fecha Pago", "Método", "Estado Pago", "Estado Factura", "Notas"
  ];
  const rows = payments.map(p => {
    const cl = clients.find(c => c.id === (p.clientFirestoreId || p.clientId));
    return [
      [cl?.company || "—"], [p.billName || ""], [p.billDoc || ""],
      [p.billEmail || ""], [p.billPhone || ""], [p.billResponsable || ""], [p.billCurso || ""],
      [p.total, "Number"], [p.paid, "Number"], [p.pending, "Number"],
      [p.date || ""], [p.method || ""], [p.status || ""], [p.billInvoiceStatus || ""], [p.billNotes || ""],
    ];
  });
  downloadXls(`ILFIS_CRM_Pagos_${now}.xls`, [{ name: "Pagos", headers, rows }]);
}

function exportFull(clients, payments, advisors) {
  const now = new Date().toLocaleDateString("es-CL");

  const clientHeaders = ["Empresa", "Contacto", "Teléfono", "Correo", "País", "Producto", "Valor (USD)", "Asesor", "Etapa", "Días Activo", "Alerta", "Fecha Creación"];
  const clientRows = clients.map(c => {
    const days = Math.floor((new Date() - new Date(c.createdAt)) / 86400000);
    const alert = (c.stage !== "Ganado" && c.stage !== "Perdido")
      ? days >= 90 ? "Expirado" : days >= 60 ? "Prioridad" : "Normal" : "—";
    return [
      [c.company], [c.contact], [c.phone || ""], [c.email || ""],
      [c.country || ""], [c.product || ""], [c.value, "Number"],
      [c.advisor], [c.stage], [days, "Number"], [alert], [c.createdAt],
    ];
  });

  const payHeaders = [
    "Empresa", "Razón Social", "RUT/NIT", "Correo", "Teléfono", "Responsable", "Curso",
    "Total", "Pagado", "Pendiente", "Fecha", "Método", "Estado Pago", "Estado Factura", "Notas"
  ];
  const payRows = payments.map(p => {
    const cl = clients.find(c => c.id === (p.clientFirestoreId || p.clientId));
    return [
      [cl?.company || "—"], [p.billName || ""], [p.billDoc || ""],
      [p.billEmail || ""], [p.billPhone || ""], [p.billResponsable || ""], [p.billCurso || ""],
      [p.total, "Number"], [p.paid, "Number"], [p.pending, "Number"],
      [p.date || ""], [p.method || ""], [p.status || ""], [p.billInvoiceStatus || ""], [p.billNotes || ""],
    ];
  });

  const advHeaders = ["Nombre", "Correo", "Teléfono", "Rol", "Estado", "Clientes Asignados", "Ganados", "Revenue (USD)"];
  const advRows = advisors.map(a => {
    const myClients = clients.filter(c => c.advisor === a.name);
    const won = myClients.filter(c => c.stage === "Ganado");
    return [
      [a.name], [a.email || ""], [a.phone || ""], [a.role || ""],
      [a.active ? "Activo" : "Inactivo"],
      [myClients.length, "Number"], [won.length, "Number"],
      [won.reduce((s, c) => s + c.value, 0), "Number"],
    ];
  });

  downloadXls(`ILFIS_CRM_Reporte_Completo_${now}.xls`, [
    { name: "Clientes y Pipeline", headers: clientHeaders, rows: clientRows },
    { name: "Pagos y Facturación", headers: payHeaders, rows: payRows },
    { name: "Asesores", headers: advHeaders, rows: advRows },
  ]);
}


const ILFIS_LOGO = "data:image/png;base64,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";

const INITIAL_ADVISORS = [
  { id: 1, name: "Laura Gómez", email: "l.gomez@crmpro.com", phone: "+56 9 1111 2222", role: "Senior", active: true },
  { id: 2, name: "Carlos Mendoza", email: "c.mendoza@crmpro.com", phone: "+52 55 3333 4444", role: "Senior", active: true },
  { id: 3, name: "Ana Ríos", email: "a.rios@crmpro.com", phone: "+51 1 5555 6666", role: "Junior", active: true },
  { id: 4, name: "Pedro Salinas", email: "p.salinas@crmpro.com", phone: "+57 1 7777 8888", role: "Junior", active: true },
  { id: 5, name: "María Torres", email: "m.torres@crmpro.com", phone: "+54 11 9999 0000", role: "Senior", active: true },
];

const STAGES = ["Contacto Inicial", "En Negociación", "Propuesta Enviada", "Ganado", "Perdido"];

const STAGE_COLORS = {
  "Contacto Inicial": { bg: "bg-slate-100", text: "text-slate-700", border: "border-slate-300", dot: "bg-slate-400" },
  "En Negociación": { bg: "bg-red-50", text: "text-red-700", border: "border-red-300", dot: "bg-red-500" },
  "Propuesta Enviada": { bg: "bg-amber-50", text: "text-amber-700", border: "border-amber-300", dot: "bg-amber-500" },
  "Ganado": { bg: "bg-emerald-50", text: "text-emerald-700", border: "border-emerald-300", dot: "bg-emerald-500" },
  "Perdido": { bg: "bg-red-50", text: "text-red-700", border: "border-red-300", dot: "bg-red-400" },
};

const now = new Date();
const daysAgo = (d) => new Date(now - d * 86400000).toISOString().split("T")[0];

const INITIAL_CLIENTS = [
  { id: 1, company: "TechNova S.A.", contact: "Rodrigo Fuentes", phone: "+56 9 1234 5678", email: "r.fuentes@technova.cl", country: "Chile", product: "CRM Enterprise", value: 8500, advisor: "Laura Gómez", stage: "En Negociación", createdAt: daysAgo(75) },
  { id: 2, company: "Grupo Alianza", contact: "Patricia Mora", phone: "+52 55 4321 9876", email: "pmora@alianza.mx", country: "México", product: "ERP Básico", value: 4200, advisor: "Carlos Mendoza", stage: "Propuesta Enviada", createdAt: daysAgo(95) },
  { id: 3, company: "DataSoft Perú", contact: "Luis Vargas", phone: "+51 1 9876 5432", email: "lvargas@datasoft.pe", country: "Perú", product: "Analytics Pro", value: 12000, advisor: "Ana Ríos", stage: "Ganado", createdAt: daysAgo(40) },
  { id: 4, company: "InnovaTech", contact: "Sandra Pérez", phone: "+57 1 1234 5678", email: "sperez@innovatech.co", country: "Colombia", product: "CRM Starter", value: 2800, advisor: "Pedro Salinas", stage: "Contacto Inicial", createdAt: daysAgo(15) },
  { id: 5, company: "Retail Express", contact: "Miguel Torres", phone: "+54 11 2345 6789", email: "mtorres@retailexpress.ar", country: "Argentina", product: "POS Cloud", value: 5600, advisor: "María Torres", stage: "En Negociación", createdAt: daysAgo(110) },
  { id: 6, company: "Consultora BCD", contact: "Andrea Luna", phone: "+56 2 3456 7890", email: "aluna@bcd.cl", country: "Chile", product: "ERP Avanzado", value: 18000, advisor: "Laura Gómez", stage: "Ganado", createdAt: daysAgo(60) },
  { id: 7, company: "StartUp Hub", contact: "Diego Ramírez", phone: "+52 33 8765 4321", email: "dramirez@startuphub.mx", country: "México", product: "CRM Starter", value: 1500, advisor: "Carlos Mendoza", stage: "Perdido", createdAt: daysAgo(120) },
  { id: 8, company: "Logística Sur", contact: "Carmen Vidal", phone: "+56 9 9876 5432", email: "cvidal@logisur.cl", country: "Chile", product: "Analytics Pro", value: 7300, advisor: "Ana Ríos", stage: "Propuesta Enviada", createdAt: daysAgo(85) },
  { id: 9, company: "EduMax Global", contact: "Jorge Blanco", phone: "+51 1 2345 6789", email: "jblanco@edumax.pe", country: "Perú", product: "LMS Empresarial", value: 9900, advisor: "Pedro Salinas", stage: "Contacto Inicial", createdAt: daysAgo(10) },
  { id: 10, company: "Finanzas Plus", contact: "Isabel Castro", phone: "+57 1 9876 1234", email: "icastro@finanzasplus.co", country: "Colombia", product: "ERP Básico", value: 6100, advisor: "María Torres", stage: "En Negociación", createdAt: daysAgo(70) },
];

const INITIAL_PAYMENTS = [
  {
    id: 1, clientId: 3, total: 12000, paid: 12000, pending: 0,
    date: daysAgo(35), method: "Transferencia", status: "Total",
    billName: "DataSoft Perú S.A.C.", billDoc: "20601234567",
    billEmail: "facturacion@datasoft.pe", billPhone: "+51 1 9876 5432",
    billResponsable: "Luis Vargas", billCurso: "Analytics Pro",
    billInvoiceStatus: "Factura Enviada", billNotes: "Transferencia BCP ref #TRF-2025-441"
  },
  {
    id: 2, clientId: 6, total: 18000, paid: 9000, pending: 9000,
    date: daysAgo(55), method: "PayPal", status: "Parcial",
    billName: "Consultora BCD SpA", billDoc: "76.543.210-9",
    billEmail: "admin@bcd.cl", billPhone: "+56 2 3456 7890",
    billResponsable: "Andrea Luna", billCurso: "ERP Avanzado",
    billInvoiceStatus: "Pendiente de Envío", billNotes: "Primer cuota PayPal #PAY-88321"
  },
];

const daysBetween = (d) => Math.floor((now - new Date(d)) / 86400000);

const formatCurrency = (v) => `$${Number(v).toLocaleString("es-CL")}`;

const AlertIcon = () => (
  <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
  </svg>
);

const LockIcon = () => (
  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" />
  </svg>
);

const CloseIcon = ({ onClick }) => (
  <button onClick={onClick} className="text-slate-400 hover:text-slate-600 transition-colors">
    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
    </svg>
  </button>
);

function Modal({ children, onClose, title }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(15,23,42,0.7)" }}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100">
          <h2 className="text-lg font-semibold text-slate-800">{title}</h2>
          <CloseIcon onClick={onClose} />
        </div>
        <div className="px-6 py-5">{children}</div>
      </div>
    </div>
  );
}

function BlockAlert({ advisor, onClose }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(15,23,42,0.7)" }}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden">
        <div className="bg-red-50 px-6 py-5 border-b border-red-100">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-full bg-red-100 flex items-center justify-center text-red-600">
              <LockIcon />
            </div>
            <div>
              <h3 className="font-bold text-red-800 text-lg">Acceso Bloqueado</h3>
              <p className="text-red-600 text-sm">Cliente asignado a otro asesor</p>
            </div>
          </div>
        </div>
        <div className="px-6 py-5">
          <p className="text-slate-700 text-base leading-relaxed">
            Este cliente está siendo atendido por{" "}
            <span className="font-bold text-slate-900">{advisor}</span>.
            No puedes modificar su información ni registrar interacciones.
          </p>
          <p className="text-slate-500 text-sm mt-2">
            Si necesitas acceso, comunícate con el asesor asignado o con el supervisor comercial.
          </p>
        </div>
        <div className="px-6 pb-5">
          <button onClick={onClose} className="w-full bg-red-600 hover:bg-red-700 text-white font-semibold py-2.5 rounded-xl transition-colors">
            Entendido
          </button>
        </div>
      </div>
    </div>
  );
}

function Badge({ stage }) {
  const c = STAGE_COLORS[stage] || STAGE_COLORS["Contacto Inicial"];
  return (
    <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold ${c.bg} ${c.text} border ${c.border}`}>
      <span className={`w-1.5 h-1.5 rounded-full ${c.dot}`}></span>
      {stage}
    </span>
  );
}

function TimeBadge({ createdAt, stage }) {
  const days = daysBetween(createdAt);
  if (stage === "Ganado" || stage === "Perdido") return null;
  if (days >= 90) {
    return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-bold bg-red-100 text-red-700 border border-red-300">
        <AlertIcon />Expirado ({days}d)
      </span>
    );
  }
  if (days >= 60) {
    return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-bold bg-amber-100 text-amber-700 border border-amber-300">
        <AlertIcon />Prioridad ({days}d)
      </span>
    );
  }
  return (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-500 border border-slate-200">
      {days}d activo
    </span>
  );
}

const emptyForm = { company: "", contact: "", phone: "", email: "", country: "", product: "", value: "", advisor: "", stage: "Contacto Inicial" };

function ClientForm({ initial, advisors = [], onSave, onClose }) {
  const [form, setForm] = useState(initial || { ...emptyForm, advisor: advisors[0] || "" });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!form.company || !form.contact || !form.advisor) return;
    onSave({ ...form, value: parseFloat(form.value) || 0 });
  };

  const inputClass = "w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-red-500 focus:border-transparent bg-slate-50 focus:bg-white transition-all";
  const labelClass = "block text-xs font-semibold text-slate-500 mb-1.5 uppercase tracking-wide";

  return (
    <form onSubmit={handleSubmit} className="space-y-4">
      <div className="grid grid-cols-2 gap-4">
        <div className="col-span-2">
          <label className={labelClass}>Empresa / Nombre *</label>
          <input className={inputClass} value={form.company} onChange={e => set("company", e.target.value)} placeholder="Ej: TechNova S.A." required />
        </div>
        <div>
          <label className={labelClass}>Contacto Principal *</label>
          <input className={inputClass} value={form.contact} onChange={e => set("contact", e.target.value)} placeholder="Nombre del contacto" required />
        </div>
        <div>
          <label className={labelClass}>Teléfono</label>
          <input className={inputClass} value={form.phone} onChange={e => set("phone", e.target.value)} placeholder="+56 9 XXXX XXXX" />
        </div>
        <div>
          <label className={labelClass}>Correo Electrónico</label>
          <input className={inputClass} type="email" value={form.email} onChange={e => set("email", e.target.value)} placeholder="email@empresa.com" />
        </div>
        <div>
          <label className={labelClass}>País</label>
          <input className={inputClass} value={form.country} onChange={e => set("country", e.target.value)} placeholder="Chile, México..." />
        </div>
        <div>
          <label className={labelClass}>Producto / Curso de Interés</label>
          <input className={inputClass} value={form.product} onChange={e => set("product", e.target.value)} placeholder="Ej: CRM Enterprise" />
        </div>
        <div>
          <label className={labelClass}>Valor de Oportunidad (USD)</label>
          <input className={inputClass} type="number" value={form.value} onChange={e => set("value", e.target.value)} placeholder="0.00" min="0" />
        </div>
        <div>
          <label className={labelClass}>Asesor Asignado *</label>
          <select className={inputClass} value={form.advisor} onChange={e => set("advisor", e.target.value)} required>
            <option value="">Seleccionar asesor</option>
            {advisors.map(a => <option key={a}>{a}</option>)}
          </select>
        </div>
        <div className="col-span-2">
          <label className={labelClass}>Etapa del Pipeline</label>
          <select className={inputClass} value={form.stage} onChange={e => set("stage", e.target.value)}>
            {STAGES.map(s => <option key={s}>{s}</option>)}
          </select>
        </div>
      </div>
      <div className="flex gap-3 pt-2">
        <button type="submit" className="flex-1 bg-red-600 hover:bg-red-700 text-white font-semibold py-2.5 rounded-xl transition-colors">
          {initial ? "Actualizar Cliente" : "Registrar Cliente"}
        </button>
        <button type="button" onClick={onClose} className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2.5 rounded-xl transition-colors">
          Cancelar
        </button>
      </div>
    </form>
  );
}

function KanbanCard({ client, onEdit, onBlock, currentUser }) {
  const days = daysBetween(client.createdAt);
  const isBlocked = client.advisor !== currentUser;
  const isExpired = days >= 90 && client.stage !== "Ganado" && client.stage !== "Perdido";
  const isPriority = days >= 60 && days < 90 && client.stage !== "Ganado" && client.stage !== "Perdido";

  const handleClick = () => {
    if (isBlocked) onBlock(client.advisor);
    else onEdit(client);
  };

  return (
    <div
      onClick={handleClick}
      className={`bg-white rounded-xl border p-3.5 cursor-pointer transition-all hover:shadow-md group ${isExpired ? "border-red-200 hover:border-red-400" : isPriority ? "border-amber-200 hover:border-amber-400" : "border-slate-200 hover:border-red-300"}`}
    >
      <div className="flex items-start justify-between gap-2 mb-2">
        <span className="font-semibold text-slate-800 text-sm leading-tight group-hover:text-red-700 transition-colors">{client.company}</span>
        {isBlocked && <div className="text-slate-400 shrink-0"><LockIcon /></div>}
      </div>
      <p className="text-xs text-slate-500 mb-1">{client.contact}</p>
      <p className="text-xs text-slate-400 mb-2.5">{client.product}</p>
      <div className="flex items-center justify-between">
        <span className="text-xs font-bold text-red-700">{formatCurrency(client.value)}</span>
        <TimeBadge createdAt={client.createdAt} stage={client.stage} />
      </div>
      {isBlocked && (
        <div className="mt-2 flex items-center gap-1 text-xs text-slate-400">
          <LockIcon /><span>{client.advisor}</span>
        </div>
      )}
    </div>
  );
}

function KanbanBoard({ clients, onEdit, onBlock, currentUser }) {
  return (
    <div className="grid grid-cols-5 gap-3 min-w-max">
      {STAGES.map(stage => {
        const cols = clients.filter(c => c.stage === stage);
        const total = cols.reduce((s, c) => s + c.value, 0);
        const sc = STAGE_COLORS[stage];
        return (
          <div key={stage} className="w-56">
            <div className={`rounded-xl px-3 py-2.5 mb-3 ${sc.bg} border ${sc.border}`}>
              <div className="flex items-center justify-between">
                <span className={`text-xs font-bold uppercase tracking-wider ${sc.text}`}>{stage}</span>
                <span className={`text-xs font-bold ${sc.text} bg-white/70 px-2 py-0.5 rounded-full`}>{cols.length}</span>
              </div>
              <p className={`text-xs mt-0.5 font-semibold ${sc.text} opacity-80`}>{formatCurrency(total)}</p>
            </div>
            <div className="space-y-2.5">
              {cols.map(c => (
                <KanbanCard key={c.id} client={c} onEdit={onEdit} onBlock={onBlock} currentUser={currentUser} />
              ))}
              {cols.length === 0 && (
                <div className="border-2 border-dashed border-slate-200 rounded-xl p-4 text-center text-xs text-slate-400">Sin registros</div>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function ClientTable({ clients, onEdit, onBlock, currentUser }) {
  return (
    <div className="overflow-x-auto rounded-2xl border border-slate-200">
      <table className="w-full text-sm">
        <thead>
          <tr className="bg-slate-50 border-b border-slate-200">
            {["Empresa", "Contacto", "País", "Producto", "Valor", "Asesor", "Estado", "Tiempo", "Acción"].map(h => (
              <th key={h} className="px-4 py-3 text-left text-xs font-bold text-slate-500 uppercase tracking-wide">{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {clients.map((c, i) => {
            const isBlocked = c.advisor !== currentUser;
            return (
              <tr key={c.id} className={`border-b border-slate-100 transition-colors ${i % 2 === 0 ? "bg-white" : "bg-slate-50/50"} hover:bg-blue-50/40`}>
                <td className="px-4 py-3 font-semibold text-slate-800">{c.company}</td>
                <td className="px-4 py-3 text-slate-600">{c.contact}</td>
                <td className="px-4 py-3 text-slate-500">{c.country}</td>
                <td className="px-4 py-3 text-slate-600 max-w-32 truncate">{c.product}</td>
                <td className="px-4 py-3 font-bold text-red-700">{formatCurrency(c.value)}</td>
                <td className="px-4 py-3">
                  <div className={`flex items-center gap-1.5 text-xs font-medium ${isBlocked ? "text-slate-400" : "text-emerald-700"}`}>
                    {isBlocked && <LockIcon />}
                    <span>{c.advisor}</span>
                  </div>
                </td>
                <td className="px-4 py-3"><Badge stage={c.stage} /></td>
                <td className="px-4 py-3"><TimeBadge createdAt={c.createdAt} stage={c.stage} /></td>
                <td className="px-4 py-3">
                  <button
                    onClick={() => isBlocked ? onBlock(c.advisor) : onEdit(c)}
                    className={`text-xs font-semibold px-3 py-1.5 rounded-lg transition-colors ${isBlocked ? "bg-slate-100 text-slate-400 cursor-not-allowed" : "bg-red-50 text-red-700 hover:bg-red-100"}`}
                  >
                    {isBlocked ? "🔒 Bloqueado" : "✏️ Editar"}
                  </button>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}


// ── Utilidad: convertir File a base64 ────────────────────────────────────────
const fileToBase64 = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(reader.result);
  reader.onerror = reject;
  reader.readAsDataURL(file);
});

const ALLOWED_TYPES = ["image/jpeg", "image/png", "image/webp", "image/gif", "application/pdf"];
const MAX_SIZE_MB = 5;

// ── Componente: visor de comprobante adjunto ──────────────────────────────────
function ComprobanteBadge({ files, onRemove }) {
  const [preview, setPreview] = useState(null);
  if (!files || files.length === 0) return null;
  return (
    <div className="flex flex-wrap gap-2 mt-2">
      {files.map((f, i) => {
        const isPdf = f.type === "application/pdf" || f.name?.endsWith(".pdf");
        const isImg = !isPdf;
        return (
          <div key={i} className="relative group">
            {isImg && (
              <button type="button" onClick={() => setPreview(f)}
                className="block w-16 h-16 rounded-xl overflow-hidden border-2 border-emerald-300 hover:border-red-400 transition-all shadow">
                <img src={f.data} alt={f.name} className="w-full h-full object-cover" />
              </button>
            )}
            {isPdf && (
              <a href={f.data} target="_blank" rel="noopener noreferrer"
                className="flex flex-col items-center justify-center w-16 h-16 rounded-xl border-2 border-red-200 bg-red-50 hover:bg-red-100 transition-all gap-1 shadow">
                <span className="text-2xl">📄</span>
                <span className="text-xs font-bold text-red-700 truncate w-14 text-center">{f.name?.split(".")[0]?.slice(0,8)}</span>
              </a>
            )}
            {onRemove && (
              <button type="button" onClick={() => onRemove(i)}
                className="absolute -top-1.5 -right-1.5 w-5 h-5 bg-red-600 text-white rounded-full text-xs font-bold flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity shadow">
                ×
              </button>
            )}
          </div>
        );
      })}
      {preview && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{background:"rgba(0,0,0,0.85)"}}
          onClick={() => setPreview(null)}>
          <div className="relative max-w-3xl max-h-[90vh]" onClick={e => e.stopPropagation()}>
            <img src={preview.data} alt={preview.name} className="max-h-[85vh] max-w-full rounded-2xl shadow-2xl object-contain" />
            <button onClick={() => setPreview(null)}
              className="absolute -top-3 -right-3 w-8 h-8 bg-red-600 text-white rounded-full font-black text-lg flex items-center justify-center shadow-lg hover:bg-red-700">
              ×
            </button>
            <a href={preview.data} download={preview.name}
              className="absolute bottom-3 left-1/2 -translate-x-1/2 bg-black/70 text-white text-xs font-semibold px-4 py-2 rounded-full hover:bg-black transition-colors">
              ⬇ Descargar {preview.name}
            </a>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Componente: zona de carga de archivos ────────────────────────────────────
function FileUploadZone({ files, onChange }) {
  const [dragging, setDragging] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadError, setUploadError] = useState("");
  const inputRef = useRef();

  const processFiles = async (rawFiles) => {
    setUploadError("");
    setUploading(true);
    const results = [];
    for (const file of Array.from(rawFiles)) {
      if (!ALLOWED_TYPES.includes(file.type)) {
        setUploadError(`Tipo no permitido: ${file.name}. Solo JPG, PNG, PDF.`);
        continue;
      }
      if (file.size > MAX_SIZE_MB * 1024 * 1024) {
        setUploadError(`${file.name} supera el límite de ${MAX_SIZE_MB}MB.`);
        continue;
      }
      const data = await fileToBase64(file);
      results.push({ name: file.name, type: file.type, size: file.size, data });
    }
    onChange([...(files || []), ...results]);
    setUploading(false);
  };

  const handleDrop = (e) => {
    e.preventDefault(); setDragging(false);
    processFiles(e.dataTransfer.files);
  };

  return (
    <div className="space-y-2">
      <div
        onDragOver={e => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={handleDrop}
        onClick={() => inputRef.current.click()}
        className={`border-2 border-dashed rounded-xl p-5 text-center cursor-pointer transition-all ${dragging ? "border-red-500 bg-red-50" : "border-slate-300 hover:border-red-400 hover:bg-red-50/30"}`}
      >
        <input ref={inputRef} type="file" accept=".jpg,.jpeg,.png,.webp,.gif,.pdf" multiple className="hidden"
          onChange={e => processFiles(e.target.files)} />
        {uploading ? (
          <div className="flex items-center justify-center gap-2 text-slate-500">
            <div className="w-4 h-4 border-2 border-red-500 border-t-transparent rounded-full animate-spin" />
            <span className="text-sm">Procesando archivos...</span>
          </div>
        ) : (
          <>
            <div className="text-3xl mb-1">📎</div>
            <p className="text-sm font-semibold text-slate-600">Arrastra aquí o haz clic para subir</p>
            <p className="text-xs text-slate-400 mt-0.5">JPG, PNG, PDF · Máx. {MAX_SIZE_MB}MB por archivo · Múltiples archivos permitidos</p>
          </>
        )}
      </div>
      {uploadError && (
        <p className="text-xs font-semibold text-red-600 flex items-center gap-1">
          <AlertIcon />  {uploadError}
        </p>
      )}
      <ComprobanteBadge files={files} onRemove={i => onChange(files.filter((_, idx) => idx !== i))} />
    </div>
  );
}

function PaymentRow({ payment, client, onEdit }) {
  const [expanded, setExpanded] = useState(false);
  const statusColors = { "Total": "bg-emerald-100 text-emerald-700 border-emerald-300", "Parcial": "bg-amber-100 text-amber-700 border-amber-300", "Pendiente": "bg-red-100 text-red-700 border-red-300" };
  const invoiceColors = { "Factura Enviada": "bg-emerald-100 text-emerald-700 border-emerald-300", "Pendiente de Envío": "bg-amber-100 text-amber-700 border-amber-300", "No Aplica": "bg-slate-100 text-slate-500 border-slate-200", "En Revisión": "bg-blue-100 text-blue-700 border-blue-200" };
  const pct = payment.total > 0 ? Math.round((payment.paid / payment.total) * 100) : 0;
  return (
    <>
      <tr className={`border-b border-slate-100 hover:bg-red-50/20 transition-colors cursor-pointer ${expanded ? "bg-red-50/30" : ""}`} onClick={() => setExpanded(e => !e)}>
        <td className="px-4 py-3.5">
          <div className="flex items-center gap-2">
            <span className={`text-xs transition-transform ${expanded ? "rotate-90" : ""} text-slate-400`}>▶</span>
            <div>
              <p className="font-semibold text-slate-800 text-sm">{client?.company || "—"}</p>
              {payment.billName && payment.billName !== client?.company && (
                <p className="text-xs text-slate-400">{payment.billName}</p>
              )}
              {payment.comprobantes?.length > 0 && (
                <span className="inline-flex items-center gap-1 text-xs font-semibold text-emerald-700 mt-0.5">
                  📎 {payment.comprobantes.length} comprobante{payment.comprobantes.length > 1 ? "s" : ""}
                </span>
              )}
            </div>
          </div>
        </td>
        <td className="px-4 py-3.5 font-bold text-slate-900">{formatCurrency(payment.total)}</td>
        <td className="px-4 py-3.5">
          <div className="flex items-center gap-2">
            <div className="flex-1 h-2 bg-slate-200 rounded-full overflow-hidden min-w-16">
              <div className="h-full bg-emerald-500 rounded-full transition-all" style={{ width: `${pct}%` }} />
            </div>
            <span className="text-xs font-bold text-emerald-700 min-w-8">{pct}%</span>
          </div>
          <span className="text-xs text-emerald-700 font-semibold">{formatCurrency(payment.paid)}</span>
        </td>
        <td className="px-4 py-3.5 font-bold text-red-600">{formatCurrency(payment.pending)}</td>
        <td className="px-4 py-3.5 text-slate-500 text-sm">{payment.date}</td>
        <td className="px-4 py-3.5 text-slate-600 text-sm">{payment.method}</td>
        <td className="px-4 py-3.5">
          <span className={`px-2.5 py-1 rounded-full text-xs font-bold border ${statusColors[payment.status] || statusColors["Pendiente"]}`}>
            {payment.status}
          </span>
        </td>
        <td className="px-4 py-3.5">
          <span className={`px-2.5 py-1 rounded-full text-xs font-bold border ${invoiceColors[payment.billInvoiceStatus] || invoiceColors["Pendiente de Envío"]}`}>
            {payment.billInvoiceStatus || "Sin estado"}
          </span>
        </td>
        <td className="px-4 py-3.5" onClick={e => e.stopPropagation()}>
          <button onClick={() => onEdit(payment)} className="text-xs font-semibold px-3 py-1.5 rounded-lg bg-red-50 text-red-700 hover:bg-red-100 transition-colors">
            Editar
          </button>
        </td>
      </tr>
      {expanded && (
        <tr className="border-b border-slate-100 bg-slate-50/70">
          <td colSpan={9} className="px-6 py-4">
            <div className="grid grid-cols-4 gap-4">
              <div className="col-span-4 flex items-center gap-2 mb-1">
                <div className="h-px flex-1 bg-slate-200" />
                <span className="text-xs font-bold text-slate-400 uppercase tracking-widest px-2">📄 Datos de Facturación</span>
                <div className="h-px flex-1 bg-slate-200" />
              </div>
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-1">Razón Social / Nombre</p>
                <p className="text-sm font-semibold text-slate-800">{payment.billName || "—"}</p>
              </div>
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-1">RUT / NIT / Doc. Comprobante</p>
                <p className="text-sm font-semibold text-slate-800 font-mono">{payment.billDoc || "—"}</p>
              </div>
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-1">Correo Facturación</p>
                <p className="text-sm font-semibold text-slate-800">{payment.billEmail || "—"}</p>
              </div>
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-1">Teléfono</p>
                <p className="text-sm font-semibold text-slate-800">{payment.billPhone || "—"}</p>
              </div>
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-1">Responsable</p>
                <p className="text-sm font-semibold text-slate-800">{payment.billResponsable || "—"}</p>
              </div>
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-1">Curso / Producto</p>
                <p className="text-sm font-semibold text-slate-800">{payment.billCurso || "—"}</p>
              </div>
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-1">Estado Factura</p>
                <span className={`px-2.5 py-1 rounded-full text-xs font-bold border ${invoiceColors[payment.billInvoiceStatus] || invoiceColors["Pendiente de Envío"]}`}>
                  {payment.billInvoiceStatus || "—"}
                </span>
              </div>
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <p className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-1">Notas / Comprobante</p>
                <p className="text-sm text-slate-600">{payment.billNotes || "—"}</p>
              </div>
              {payment.comprobantes && payment.comprobantes.length > 0 && (
                <div className="col-span-4 bg-white rounded-xl border border-emerald-200 p-4">
                  <p className="text-xs font-bold text-slate-400 uppercase tracking-wide mb-2">
                    📎 Comprobantes Adjuntos ({payment.comprobantes.length})
                  </p>
                  <div className="flex flex-wrap gap-3">
                    {payment.comprobantes.map((f, i) => {
                      const isPdf = f.type === "application/pdf" || f.name?.endsWith(".pdf");
                      return isPdf ? (
                        <a key={i} href={f.data} target="_blank" rel="noopener noreferrer"
                          className="flex items-center gap-2 px-3 py-2 bg-red-50 border border-red-200 rounded-xl hover:bg-red-100 transition-colors">
                          <span className="text-lg">📄</span>
                          <div>
                            <p className="text-xs font-bold text-red-700 max-w-32 truncate">{f.name}</p>
                            <p className="text-xs text-slate-400">{f.size ? (f.size/1024).toFixed(0)+"KB" : "PDF"}</p>
                          </div>
                          <span className="text-xs text-red-600 font-bold">↗ Abrir</span>
                        </a>
                      ) : (
                        <div key={i} className="flex items-center gap-2 px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl">
                          <img src={f.data} alt={f.name} className="w-10 h-10 object-cover rounded-lg border border-slate-200 cursor-pointer"
                            onClick={() => window.open(f.data, "_blank")} />
                          <div>
                            <p className="text-xs font-bold text-slate-700 max-w-28 truncate">{f.name}</p>
                            <p className="text-xs text-slate-400">{f.size ? (f.size/1024).toFixed(0)+"KB" : "Imagen"}</p>
                          </div>
                          <a href={f.data} download={f.name} className="text-xs text-blue-600 font-bold hover:underline">⬇</a>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          </td>
        </tr>
      )}
    </>
  );
}

function PaymentForm({ clients, initial, onSave, onClose }) {
  const wonClients = clients.filter(c => c.stage === "Ganado");
  const defaultForm = {
    clientId: wonClients[0]?.id || "", total: "", paid: "",
    date: new Date().toISOString().split("T")[0], method: "Transferencia", status: "Parcial",
    billName: "", billDoc: "", billEmail: "", billPhone: "",
    billResponsable: "", billCurso: "", billInvoiceStatus: "Pendiente de Envío", billNotes: "",
    comprobantes: []
  };
  const [form, setForm] = useState(initial || defaultForm);
  const [section, setSection] = useState("pago");
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const pending = Math.max(0, parseFloat(form.total || 0) - parseFloat(form.paid || 0));

  const autoFillFromClient = (clientId) => {
    const cl = clients.find(c => c.id === parseInt(clientId));
    if (!cl) return;
    setForm(f => ({
      ...f, clientId,
      billName: f.billName || cl.company,
      billEmail: f.billEmail || cl.email,
      billPhone: f.billPhone || cl.phone,
      billResponsable: f.billResponsable || cl.contact,
      billCurso: f.billCurso || cl.product,
    }));
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    const total = parseFloat(form.total) || 0;
    const paid = parseFloat(form.paid) || 0;
    const autoStatus = paid >= total ? "Total" : paid > 0 ? "Parcial" : "Pendiente";
    onSave({ ...form, clientId: parseInt(form.clientId), total, paid, pending, status: autoStatus });
  };

  const ic = "w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-red-500 bg-slate-50 focus:bg-white transition-all";
  const lc = "block text-xs font-semibold text-slate-500 mb-1.5 uppercase tracking-wide";

  return (
    <form onSubmit={handleSubmit} className="space-y-4">
      <div className="flex gap-1 bg-slate-100 rounded-xl p-1 mb-2">
        {[["pago","💳 Datos del Pago"],["factura","📄 Datos de Facturación"],["comprobante","📎 Comprobantes"]].map(([s, label]) => (
          <button key={s} type="button" onClick={() => setSection(s)}
            className={`flex-1 py-2 rounded-lg text-xs font-bold transition-colors ${section === s ? "bg-white text-slate-800 shadow" : "text-slate-500 hover:text-slate-700"}`}>
            {label}
            {s === "comprobante" && form.comprobantes?.length > 0 && (
              <span className="ml-1 bg-emerald-500 text-white text-xs font-black px-1.5 py-0.5 rounded-full">{form.comprobantes.length}</span>
            )}
          </button>
        ))}
      </div>

      {section === "pago" && (
        <div className="grid grid-cols-2 gap-4">
          <div className="col-span-2">
            <label className={lc}>Cliente (Oportunidad Ganada) *</label>
            <select className={ic} value={form.clientId} onChange={e => autoFillFromClient(e.target.value)} required>
              <option value="">Seleccionar cliente</option>
              {wonClients.map(c => <option key={c.id} value={c.id}>{c.company} — {c.contact}</option>)}
            </select>
            <p className="text-xs text-slate-400 mt-1">Al seleccionar, los datos de facturación se pre-rellenan automáticamente.</p>
          </div>
          <div>
            <label className={lc}>Monto Total (USD)</label>
            <input className={ic} type="number" value={form.total} onChange={e => set("total", e.target.value)} min="0" placeholder="0.00" />
          </div>
          <div>
            <label className={lc}>Monto Pagado (USD)</label>
            <input className={ic} type="number" value={form.paid} onChange={e => set("paid", e.target.value)} min="0" placeholder="0.00" />
          </div>
          <div>
            <label className={lc}>Saldo Pendiente</label>
            <div className={`${ic} bg-red-50 text-red-700 font-bold`}>{formatCurrency(pending)}</div>
          </div>
          <div>
            <label className={lc}>Fecha de Pago</label>
            <input className={ic} type="date" value={form.date} onChange={e => set("date", e.target.value)} />
          </div>
          <div className="col-span-2">
            <label className={lc}>Método de Pago</label>
            <select className={ic} value={form.method} onChange={e => set("method", e.target.value)}>
              {["Transferencia", "PayPal", "Stripe", "Tarjeta de Crédito", "Efectivo", "Cheque"].map(m => <option key={m}>{m}</option>)}
            </select>
          </div>
        </div>
      )}

      {section === "factura" && (
        <div className="grid grid-cols-2 gap-4">
          <div className="col-span-2 bg-amber-50 border border-amber-200 rounded-xl px-4 py-2.5 text-xs text-amber-700 font-medium">
            💡 Estos datos aparecerán en el registro expandible de la tabla de pagos.
          </div>
          <div className="col-span-2">
            <label className={lc}>Razón Social / Nombre Facturación</label>
            <input className={ic} value={form.billName} onChange={e => set("billName", e.target.value)} placeholder="Ej: DataSoft Perú S.A.C." />
          </div>
          <div>
            <label className={lc}>RUT / NIT / N° Comprobante</label>
            <input className={ic} value={form.billDoc} onChange={e => set("billDoc", e.target.value)} placeholder="Ej: 20601234567" />
          </div>
          <div>
            <label className={lc}>Correo Facturación</label>
            <input className={ic} type="email" value={form.billEmail} onChange={e => set("billEmail", e.target.value)} placeholder="facturacion@empresa.com" />
          </div>
          <div>
            <label className={lc}>Teléfono</label>
            <input className={ic} value={form.billPhone} onChange={e => set("billPhone", e.target.value)} placeholder="+56 9 XXXX XXXX" />
          </div>
          <div>
            <label className={lc}>Responsable</label>
            <input className={ic} value={form.billResponsable} onChange={e => set("billResponsable", e.target.value)} placeholder="Nombre del responsable" />
          </div>
          <div>
            <label className={lc}>Curso / Producto</label>
            <input className={ic} value={form.billCurso} onChange={e => set("billCurso", e.target.value)} placeholder="Ej: CRM Enterprise" />
          </div>
          <div className="col-span-2">
            <label className={lc}>Estado de Factura</label>
            <select className={ic} value={form.billInvoiceStatus} onChange={e => set("billInvoiceStatus", e.target.value)}>
              {["Pendiente de Envío", "Factura Enviada", "En Revisión", "No Aplica"].map(s => <option key={s}>{s}</option>)}
            </select>
          </div>
          <div className="col-span-2">
            <label className={lc}>Notas / Referencia del Comprobante</label>
            <textarea className={ic} rows={2} value={form.billNotes} onChange={e => set("billNotes", e.target.value)} placeholder="Ej: Transferencia BCP ref #TRF-2025-441" />
          </div>
        </div>
      )}

      {section === "comprobante" && (
        <div className="space-y-4">
          <div className="bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-xs text-slate-600">
            <p className="font-bold text-slate-700 mb-1">📎 Adjunta los comprobantes de pago</p>
            <p>Puedes subir capturas de transferencia, vouchers bancarios, recibos o facturas en formato <strong>JPG, PNG o PDF</strong>.</p>
            <p className="mt-1 text-slate-400">Los archivos se guardan junto al registro del pago y son visibles para todos los asesores.</p>
          </div>
          <FileUploadZone
            files={form.comprobantes || []}
            onChange={files => set("comprobantes", files)}
          />
          {(!form.comprobantes || form.comprobantes.length === 0) && (
            <div className="text-center py-4 text-slate-400 text-sm">
              Aún no hay comprobantes adjuntos para este pago.
            </div>
          )}
        </div>
      )}

      <div className="flex gap-3 pt-2">
        <button type="submit" className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-2.5 rounded-xl transition-colors">
          {initial ? "Actualizar Pago" : "Registrar Pago"}
        </button>
        <button type="button" onClick={onClose} className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2.5 rounded-xl transition-colors">
          Cancelar
        </button>
      </div>
    </form>
  );
}

function MetricCard({ label, value, sub, color = "blue" }) {
  const colors = { blue: "from-red-600 to-red-700", emerald: "from-emerald-600 to-emerald-700", amber: "from-amber-500 to-amber-600", red: "from-red-700 to-red-800", slate: "from-zinc-700 to-zinc-800" };
  return (
    <div className={`rounded-2xl bg-gradient-to-br ${colors[color]} p-5 text-white`}>
      <p className="text-xs font-semibold uppercase tracking-widest opacity-70 mb-1">{label}</p>
      <p className="text-3xl font-bold mb-1">{value}</p>
      {sub && <p className="text-xs opacity-70">{sub}</p>}
    </div>
  );
}

function DashboardTab({ clients, payments, advisors = [], currentUser }) {
  const won = clients.filter(c => c.stage === "Ganado");
  const lost = clients.filter(c => c.stage === "Perdido");
  const active = clients.filter(c => c.stage !== "Ganado" && c.stage !== "Perdido");
  const thisMonth = won.filter(c => daysBetween(c.createdAt) <= 30);
  const convRate = clients.length > 0 ? Math.round((won.length / clients.length) * 100) : 0;
  const totalRevenue = won.reduce((s, c) => s + c.value, 0);
  const totalPaid = payments.reduce((s, p) => s + p.paid, 0);
  const expired = clients.filter(c => daysBetween(c.createdAt) >= 90 && c.stage !== "Ganado" && c.stage !== "Perdido");
  const priority = clients.filter(c => { const d = daysBetween(c.createdAt); return d >= 60 && d < 90 && c.stage !== "Ganado" && c.stage !== "Perdido"; });

  const advisorNames = advisors.length > 0 ? advisors.map(a => a.name) : [...new Set(clients.map(c => c.advisor).filter(Boolean))];
  const byAdvisor = advisorNames.map(a => ({
    name: a,
    total: clients.filter(c => c.advisor === a).length,
    won: clients.filter(c => c.advisor === a && c.stage === "Ganado").length,
    value: clients.filter(c => c.advisor === a && c.stage === "Ganado").reduce((s, c) => s + c.value, 0),
  })).sort((a, b) => b.value - a.value);

  const maxValue = Math.max(...byAdvisor.map(a => a.value), 1);

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
        <MetricCard label="Ingresos Totales" value={formatCurrency(totalRevenue)} sub={`${won.length} oportunidades cerradas`} color="blue" />
        <MetricCard label="Cobrado" value={formatCurrency(totalPaid)} sub="Pagos registrados" color="emerald" />
        <MetricCard label="Tasa de Conversión" value={`${convRate}%`} sub={`${won.length} ganados / ${clients.length} totales`} color="slate" />
        <MetricCard label="Cierres del Mes" value={thisMonth.length} sub={formatCurrency(thisMonth.reduce((s, c) => s + c.value, 0))} color="amber" />
        <MetricCard label="En Negociación" value={active.length} sub="Oportunidades activas" color="slate" />
      </div>

      {(expired.length > 0 || priority.length > 0) && (
        <div className="grid grid-cols-2 gap-4">
          {expired.length > 0 && (
            <div className="rounded-2xl border-2 border-red-200 bg-red-50 p-5">
              <div className="flex items-center gap-2 mb-3">
                <div className="text-red-600"><AlertIcon /></div>
                <h3 className="font-bold text-red-800">Oportunidades Expiradas</h3>
                <span className="ml-auto bg-red-200 text-red-800 text-xs font-bold px-2.5 py-0.5 rounded-full">{expired.length}</span>
              </div>
              <div className="space-y-2">
                {expired.map(c => (
                  <div key={c.id} className="bg-white rounded-xl p-3 border border-red-200 flex items-center justify-between">
                    <div>
                      <p className="font-semibold text-slate-800 text-sm">{c.company}</p>
                      <p className="text-xs text-slate-500">{c.advisor} · {c.stage}</p>
                    </div>
                    <div className="text-right">
                      <p className="text-xs font-bold text-red-700">{daysBetween(c.createdAt)}d</p>
                      <p className="text-xs font-bold text-red-700">{formatCurrency(c.value)}</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
          {priority.length > 0 && (
            <div className="rounded-2xl border-2 border-amber-200 bg-amber-50 p-5">
              <div className="flex items-center gap-2 mb-3">
                <div className="text-amber-600"><AlertIcon /></div>
                <h3 className="font-bold text-amber-800">Prioridad de Cierre (60–90d)</h3>
                <span className="ml-auto bg-amber-200 text-amber-800 text-xs font-bold px-2.5 py-0.5 rounded-full">{priority.length}</span>
              </div>
              <div className="space-y-2">
                {priority.map(c => (
                  <div key={c.id} className="bg-white rounded-xl p-3 border border-amber-200 flex items-center justify-between">
                    <div>
                      <p className="font-semibold text-slate-800 text-sm">{c.company}</p>
                      <p className="text-xs text-slate-500">{c.advisor} · {c.stage}</p>
                    </div>
                    <div className="text-right">
                      <p className="text-xs font-bold text-amber-700">{daysBetween(c.createdAt)}d</p>
                      <p className="text-xs font-bold text-red-700">{formatCurrency(c.value)}</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      <div className="grid grid-cols-2 gap-4">
        <div className="rounded-2xl border border-slate-200 bg-white p-5">
          <h3 className="font-bold text-slate-800 mb-4">Rendimiento por Asesor</h3>
          <div className="space-y-3">
            {byAdvisor.map(a => (
              <div key={a.name}>
                <div className="flex items-center justify-between mb-1">
                  <span className={`text-sm font-medium ${a.name === currentUser ? "text-red-700 font-bold" : "text-slate-700"}`}>
                    {a.name === currentUser ? "★ " : ""}{a.name}
                  </span>
                  <div className="flex gap-3 text-xs text-slate-500">
                    <span className="font-semibold text-emerald-700">{formatCurrency(a.value)}</span>
                    <span>{a.won}/{a.total}</span>
                  </div>
                </div>
                <div className="h-2 bg-slate-100 rounded-full overflow-hidden">
                  <div className="h-full bg-red-600 rounded-full transition-all" style={{ width: `${(a.value / maxValue) * 100}%` }} />
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="rounded-2xl border border-slate-200 bg-white p-5">
          <h3 className="font-bold text-slate-800 mb-4">Distribución del Pipeline</h3>
          <div className="space-y-2.5">
            {STAGES.map(stage => {
              const stageClients = clients.filter(c => c.stage === stage);
              const stageValue = stageClients.reduce((s, c) => s + c.value, 0);
              const pct = clients.length > 0 ? Math.round((stageClients.length / clients.length) * 100) : 0;
              const sc = STAGE_COLORS[stage];
              return (
                <div key={stage} className="flex items-center gap-3">
                  <span className={`w-2.5 h-2.5 rounded-full ${sc.dot} shrink-0`}></span>
                  <span className="text-sm text-slate-700 w-36 truncate">{stage}</span>
                  <div className="flex-1 h-2 bg-slate-100 rounded-full overflow-hidden">
                    <div className={`h-full ${sc.dot} rounded-full`} style={{ width: `${pct}%` }} />
                  </div>
                  <span className="text-xs font-bold text-slate-600 w-8 text-right">{stageClients.length}</span>
                  <span className="text-xs font-semibold text-blue-700 w-20 text-right">{formatCurrency(stageValue)}</span>
                </div>
              );
            })}
          </div>
          <div className="mt-4 pt-4 border-t border-slate-100 flex justify-between">
            <span className="text-sm font-semibold text-slate-600">Pipeline Total</span>
            <span className="text-sm font-bold text-red-700">{formatCurrency(clients.reduce((s, c) => s + c.value, 0))}</span>
          </div>
        </div>
      </div>
    </div>
  );
}

function AdvisorsTab({ advisors, setAdvisors, clients, currentUser, setCurrentUser }) {
  const [showForm, setShowForm] = useState(false);
  const [editAdvisor, setEditAdvisor] = useState(null);
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [form, setForm] = useState({ name: "", email: "", phone: "", role: "Junior", active: true });

  const openNew = () => { setForm({ name: "", email: "", phone: "", role: "Junior", active: true }); setEditAdvisor(null); setShowForm(true); };
  const openEdit = (a) => { setForm({ ...a }); setEditAdvisor(a); setShowForm(true); };

  const handleSave = (e) => {
    e.preventDefault();
    if (!form.name.trim()) return;
    if (editAdvisor) {
      setAdvisors(prev => prev.map(a => a.id === editAdvisor.id ? { ...a, ...form } : a));
    } else {
      setAdvisors(prev => [...prev, { ...form, id: Date.now() }]);
    }
    setShowForm(false);
  };

  const handleDelete = (advisor) => {
    const assigned = clients.filter(c => c.advisor === advisor.name).length;
    if (assigned > 0) { setDeleteConfirm({ advisor, assigned }); return; }
    setAdvisors(prev => prev.filter(a => a.id !== advisor.id));
  };

  const confirmDelete = () => {
    setAdvisors(prev => prev.filter(a => a.id !== deleteConfirm.advisor.id));
    setDeleteConfirm(null);
  };

  const inputClass = "w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-red-500 bg-slate-50 focus:bg-white transition-all";
  const labelClass = "block text-xs font-semibold text-slate-500 mb-1.5 uppercase tracking-wide";
  const roleColors = { Senior: "bg-blue-100 text-blue-700 border-blue-200", Junior: "bg-slate-100 text-slate-600 border-slate-200" };

  return (
    <div>
      <div className="flex items-center justify-between mb-5">
        <div>
          <h1 className="text-2xl font-black text-slate-900">Gestión de Asesores</h1>
          <p className="text-slate-500 text-sm mt-0.5">{advisors.length} asesores registrados · {advisors.filter(a => a.active).length} activos</p>
        </div>
        <button onClick={openNew} className="flex items-center gap-2 bg-red-600 hover:bg-red-700 text-white font-bold px-5 py-2.5 rounded-xl transition-colors shadow-md shadow-red-200">
          <span className="text-lg leading-none">+</span>
          <span>Nuevo Asesor</span>
        </button>
      </div>

      <div className="mb-5 bg-red-50 border border-red-200 rounded-2xl px-5 py-3.5 flex items-center gap-3">
        <span className="text-red-500 text-xl">👤</span>
        <div className="flex-1">
          <p className="text-sm font-semibold text-red-800">Sesión activa como: <span className="font-black">{currentUser}</span></p>
          <p className="text-xs text-red-600">Cambiar de sesión simula el acceso de otro asesor y activa la lógica de bloqueo</p>
        </div>
        <select
          className="border border-red-300 bg-white text-red-800 font-semibold rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-red-500"
          value={currentUser}
          onChange={e => setCurrentUser(e.target.value)}
        >
          {advisors.filter(a => a.active).map(a => <option key={a.id}>{a.name}</option>)}
        </select>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        {advisors.map(advisor => {
          const myClients = clients.filter(c => c.advisor === advisor.name);
          const won = myClients.filter(c => c.stage === "Ganado").length;
          const revenue = myClients.filter(c => c.stage === "Ganado").reduce((s, c) => s + c.value, 0);
          const isMe = advisor.name === currentUser;
          return (
            <div key={advisor.id} className={`bg-white rounded-2xl border-2 p-5 transition-all ${isMe ? "border-red-500 shadow-lg shadow-red-100" : "border-slate-200 hover:border-slate-300"} ${!advisor.active ? "opacity-60" : ""}`}>
              <div className="flex items-start justify-between mb-3">
                <div className="flex items-center gap-3">
                  <div className={`w-11 h-11 rounded-xl flex items-center justify-center text-sm font-black ${isMe ? "bg-red-600 text-white" : "bg-slate-100 text-slate-700"}`}>
                    {advisor.name.split(" ").map(n => n[0]).join("").slice(0, 2)}
                  </div>
                  <div>
                    <p className={`font-bold text-slate-900 ${isMe ? "text-red-700" : ""}`}>
                      {isMe && <span className="text-xs">⭐ </span>}{advisor.name}
                    </p>
                    <div className="flex items-center gap-1.5 mt-0.5">
                      <span className={`text-xs font-semibold px-2 py-0.5 rounded-full border ${roleColors[advisor.role] || roleColors.Junior}`}>{advisor.role}</span>
                      <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${advisor.active ? "bg-emerald-100 text-emerald-700" : "bg-slate-100 text-slate-500"}`}>
                        {advisor.active ? "Activo" : "Inactivo"}
                      </span>
                    </div>
                  </div>
                </div>
                {isMe && <span className="text-xs font-bold bg-red-100 text-red-700 px-2 py-1 rounded-lg">Tú</span>}
              </div>

              <div className="space-y-1 mb-3 text-xs text-slate-500">
                {advisor.email && <p className="flex items-center gap-1.5">✉️ <span className="truncate">{advisor.email}</span></p>}
                {advisor.phone && <p className="flex items-center gap-1.5">📞 {advisor.phone}</p>}
              </div>

              <div className="grid grid-cols-3 gap-2 mb-4 pt-3 border-t border-slate-100">
                <div className="text-center">
                  <p className="text-xl font-black text-slate-800">{myClients.length}</p>
                  <p className="text-xs text-slate-400">Clientes</p>
                </div>
                <div className="text-center">
                  <p className="text-xl font-black text-emerald-600">{won}</p>
                  <p className="text-xs text-slate-400">Ganados</p>
                </div>
                <div className="text-center">
                  <p className="text-sm font-black text-red-700">{formatCurrency(revenue)}</p>
                  <p className="text-xs text-slate-400">Revenue</p>
                </div>
              </div>

              <div className="flex gap-2">
                <button onClick={() => openEdit(advisor)} className="flex-1 text-xs font-semibold py-2 rounded-lg bg-zinc-100 text-zinc-700 hover:bg-zinc-200 transition-colors">
                  ✏️ Editar
                </button>
                <button
                  onClick={() => setAdvisors(prev => prev.map(a => a.id === advisor.id ? { ...a, active: !a.active } : a))}
                  className={`flex-1 text-xs font-semibold py-2 rounded-lg transition-colors ${advisor.active ? "bg-amber-50 text-amber-700 hover:bg-amber-100" : "bg-emerald-50 text-emerald-700 hover:bg-emerald-100"}`}
                >
                  {advisor.active ? "⏸ Desactivar" : "▶ Activar"}
                </button>
                <button onClick={() => handleDelete(advisor)} className="px-3 py-2 rounded-lg bg-red-50 text-red-500 hover:bg-red-100 transition-colors text-xs font-semibold">
                  🗑
                </button>
              </div>
            </div>
          );
        })}
      </div>

      {showForm && (
        <Modal title={editAdvisor ? `Editar: ${editAdvisor.name}` : "Nuevo Asesor"} onClose={() => setShowForm(false)}>
          <form onSubmit={handleSave} className="space-y-4">
            <div className="grid grid-cols-2 gap-4">
              <div className="col-span-2">
                <label className={labelClass}>Nombre Completo *</label>
                <input className={inputClass} value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))} placeholder="Ej: Juan Pérez" required />
              </div>
              <div>
                <label className={labelClass}>Correo Electrónico</label>
                <input className={inputClass} type="email" value={form.email} onChange={e => setForm(f => ({ ...f, email: e.target.value }))} placeholder="juan@empresa.com" />
              </div>
              <div>
                <label className={labelClass}>Teléfono</label>
                <input className={inputClass} value={form.phone} onChange={e => setForm(f => ({ ...f, phone: e.target.value }))} placeholder="+56 9 XXXX XXXX" />
              </div>
              <div>
                <label className={labelClass}>Rol / Nivel</label>
                <select className={inputClass} value={form.role} onChange={e => setForm(f => ({ ...f, role: e.target.value }))}>
                  <option>Junior</option>
                  <option>Senior</option>
                  <option>Líder Comercial</option>
                  <option>Director</option>
                </select>
              </div>
              <div>
                <label className={labelClass}>Estado</label>
                <select className={inputClass} value={form.active ? "activo" : "inactivo"} onChange={e => setForm(f => ({ ...f, active: e.target.value === "activo" }))}>
                  <option value="activo">Activo</option>
                  <option value="inactivo">Inactivo</option>
                </select>
              </div>
            </div>
            <div className="flex gap-3 pt-2">
              <button type="submit" className="flex-1 bg-red-600 hover:bg-red-700 text-white font-semibold py-2.5 rounded-xl transition-colors">
                {editAdvisor ? "Actualizar Asesor" : "Agregar Asesor"}
              </button>
              <button type="button" onClick={() => setShowForm(false)} className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2.5 rounded-xl transition-colors">
                Cancelar
              </button>
            </div>
          </form>
        </Modal>
      )}

      {deleteConfirm && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(15,23,42,0.7)" }}>
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-6">
            <div className="text-center mb-4">
              <div className="text-5xl mb-3">⚠️</div>
              <h3 className="font-bold text-slate-900 text-lg">¿Eliminar asesor?</h3>
              <p className="text-slate-600 text-sm mt-2">
                <strong>{deleteConfirm.advisor.name}</strong> tiene <strong>{deleteConfirm.assigned} cliente(s)</strong> asignados.
                Si lo eliminas, esos clientes quedarán sin asesor asignado.
              </p>
            </div>
            <div className="flex gap-3">
              <button onClick={confirmDelete} className="flex-1 bg-red-600 hover:bg-red-700 text-white font-semibold py-2.5 rounded-xl transition-colors">
                Eliminar de todas formas
              </button>
              <button onClick={() => setDeleteConfirm(null)} className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2.5 rounded-xl transition-colors">
                Cancelar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default function CRM() {
  const [tab, setTab] = useState("dashboard");
  const [view, setView] = useState("kanban");
  const [clients, setClients] = useState(INITIAL_CLIENTS);
  const [payments, setPayments] = useState(INITIAL_PAYMENTS);
  const [advisors, setAdvisors] = useState(INITIAL_ADVISORS);
  const [currentUser, setCurrentUser] = useState("Laura Gómez");
  const [search, setSearch] = useState("");
  const [filterStage, setFilterStage] = useState("Todas");
  const [showClientForm, setShowClientForm] = useState(false);
  const [editClient, setEditClient] = useState(null);
  const [blockAlert, setBlockAlert] = useState(null);
  const [showPayForm, setShowPayForm] = useState(false);
  const [editPayment, setEditPayment] = useState(null);

  const activeAdvisorNames = advisors.filter(a => a.active).map(a => a.name);

  const filteredClients = useMemo(() => {
    let list = clients;
    if (search) list = list.filter(c => c.company.toLowerCase().includes(search.toLowerCase()) || c.contact.toLowerCase().includes(search.toLowerCase()));
    if (filterStage !== "Todas") list = list.filter(c => c.stage === filterStage);
    return list;
  }, [clients, search, filterStage]);

  const handleSaveClient = (data) => {
    if (editClient) {
      setClients(prev => prev.map(c => c.id === editClient.id ? { ...c, ...data } : c));
    } else {
      setClients(prev => [...prev, { ...data, id: Date.now(), createdAt: new Date().toISOString().split("T")[0] }]);
    }
    setShowClientForm(false);
    setEditClient(null);
  };

  const handleEditClient = (client) => {
    if (client.advisor !== currentUser) { setBlockAlert(client.advisor); return; }
    setEditClient(client);
    setShowClientForm(true);
  };

  const handleSavePayment = (data) => {
    if (editPayment) {
      setPayments(prev => prev.map(p => p.id === editPayment.id ? { ...p, ...data } : p));
    } else {
      setPayments(prev => [...prev, { ...data, id: Date.now() }]);
    }
    setShowPayForm(false);
    setEditPayment(null);
  };

  const TABS = [
    { id: "dashboard", label: "Dashboard & Reportes", icon: "📊" },
    { id: "pipeline", label: "Pipeline & Clientes", icon: "🎯" },
    { id: "payments", label: "Relación de Pagos", icon: "💳" },
    { id: "advisors", label: "Asesores", icon: "👥" },
  ];

  const wonClients = clients.filter(c => c.stage === "Ganado");
  const totalPending = payments.reduce((s, p) => s + p.pending, 0);

  return (
    <div className="min-h-screen bg-gray-50 font-sans">
      {/* Header */}
      <header className="bg-black text-white sticky top-0 z-40 shadow-xl border-b-2 border-red-600">
        <div className="max-w-screen-2xl mx-auto px-6 py-0">
          <div className="flex items-center gap-6 h-16">
            <div className="flex items-center gap-3">
              <img src={ILFIS_LOGO} alt="ILFIS" className="h-7 w-auto brightness-0 invert" />
              <div className="h-6 w-px bg-red-600 mx-1"></div>
              <span className="text-xs font-bold text-red-500 uppercase tracking-widest">CRM</span>
            </div>
            <nav className="flex gap-1 flex-1">
              {TABS.map(t => (
                <button
                  key={t.id}
                  onClick={() => setTab(t.id)}
                  className={`flex items-center gap-2 px-4 py-1.5 rounded-lg text-sm font-semibold transition-all ${tab === t.id ? "bg-red-600 text-white shadow" : "text-slate-400 hover:text-white hover:bg-zinc-800"}`}
                >
                  <span>{t.icon}</span>
                  <span>{t.label}</span>
                </button>
              ))}
            </nav>
            <div className="flex items-center gap-3 text-sm">
              <div className="flex items-center gap-2 bg-zinc-900 rounded-xl px-3 py-2 border border-zinc-700">
                <div className="w-6 h-6 rounded-full bg-red-600 flex items-center justify-center text-xs font-bold">
                  {currentUser.split(" ").map(n => n[0]).join("")}
                </div>
                <span className="text-slate-300 text-xs font-medium">{currentUser}</span>
                <span className="text-emerald-400 text-xs">● Online</span>
              </div>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-screen-2xl mx-auto px-6 py-6">
        {tab === "dashboard" && (
          <div>
            <div className="flex items-center justify-between mb-6">
              <div>
                <h1 className="text-2xl font-black text-slate-900">Dashboard Comercial</h1>
                <p className="text-slate-500 text-sm mt-0.5">Resumen ejecutivo y reportes de cierre</p>
              </div>
              <div className="flex items-center gap-3">
                <div className="flex items-center gap-2 text-xs text-slate-500 bg-white border border-slate-200 rounded-xl px-4 py-2">
                  <span>📅</span>
                  <span>{new Date().toLocaleDateString("es-CL", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}</span>
                </div>
                <button
                  onClick={() => exportFull(clients, payments, advisors)}
                  className="flex items-center gap-2 bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-bold px-4 py-2 rounded-xl transition-colors shadow"
                  title="Exportar reporte completo a Excel"
                >
                  <span>⬇</span>
                  <span>Exportar Todo</span>
                </button>
              </div>
            </div>
            <DashboardTab clients={clients} payments={payments} advisors={advisors} currentUser={currentUser} />
          </div>
        )}

        {tab === "pipeline" && (
          <div>
            <div className="flex items-center justify-between mb-5">
              <div>
                <h1 className="text-2xl font-black text-slate-900">Pipeline de Ventas</h1>
                <p className="text-slate-500 text-sm mt-0.5">{clients.length} clientes registrados · {filteredClients.length} visibles</p>
              </div>
              <div className="flex gap-2">
                <div className="flex bg-white border border-slate-200 rounded-xl overflow-hidden">
                  <button onClick={() => setView("kanban")} className={`px-4 py-2 text-sm font-semibold transition-colors ${view === "kanban" ? "bg-red-600 text-white" : "text-slate-600 hover:bg-gray-50"}`}>🗂 Kanban</button>
                  <button onClick={() => setView("table")} className={`px-4 py-2 text-sm font-semibold transition-colors ${view === "table" ? "bg-red-600 text-white" : "text-slate-600 hover:bg-gray-50"}`}>📋 Tabla</button>
                </div>
                <button
                  onClick={() => exportClients(filteredClients, filterStage === "Todas" ? "Clientes" : filterStage.replace(/ /g,"_"))}
                  className="flex items-center gap-2 bg-emerald-700 hover:bg-emerald-800 text-white text-sm font-bold px-4 py-2.5 rounded-xl transition-colors shadow"
                  title="Exportar clientes visibles a Excel"
                >
                  <span>⬇</span>
                  <span>Excel</span>
                </button>
                <button
                  onClick={() => { setEditClient(null); setShowClientForm(true); }}
                  className="flex items-center gap-2 bg-red-600 hover:bg-red-700 text-white font-bold px-5 py-2.5 rounded-xl transition-colors shadow-md shadow-red-200"
                >
                  <span className="text-lg leading-none">+</span>
                  <span>Nuevo Cliente</span>
                </button>
              </div>
            </div>

            <div className="flex gap-3 mb-5">
              <div className="flex-1 relative">
                <span className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400">🔍</span>
                <input
                  className="w-full pl-10 pr-4 py-2.5 border border-slate-200 rounded-xl text-sm bg-white focus:outline-none focus:ring-2 focus:ring-red-500"
                  placeholder="Buscar empresa o contacto..."
                  value={search}
                  onChange={e => setSearch(e.target.value)}
                />
              </div>
              <select
                className="border border-slate-200 rounded-xl px-4 py-2.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-red-500"
                value={filterStage}
                onChange={e => setFilterStage(e.target.value)}
              >
                <option>Todas</option>
                {STAGES.map(s => <option key={s}>{s}</option>)}
              </select>
            </div>

            {view === "kanban" ? (
              <div className="overflow-x-auto pb-4">
                <KanbanBoard clients={filteredClients} onEdit={handleEditClient} onBlock={setBlockAlert} currentUser={currentUser} />
              </div>
            ) : (
              <ClientTable clients={filteredClients} onEdit={handleEditClient} onBlock={setBlockAlert} currentUser={currentUser} />
            )}
          </div>
        )}

        {tab === "payments" && (
          <div>
            <div className="flex items-center justify-between mb-5">
              <div>
                <h1 className="text-2xl font-black text-slate-900">Relación de Pagos</h1>
                <p className="text-slate-500 text-sm mt-0.5">Gestión financiera de clientes con oportunidades ganadas</p>
              </div>
              <div className="flex gap-2">
                <button
                  onClick={() => exportPayments(payments, clients)}
                  className="flex items-center gap-2 bg-emerald-700 hover:bg-emerald-800 text-white text-sm font-bold px-4 py-2.5 rounded-xl transition-colors shadow"
                  title="Exportar pagos a Excel"
                >
                  <span>⬇</span>
                  <span>Excel</span>
                </button>
                <button
                  onClick={() => { setEditPayment(null); setShowPayForm(true); }}
                  className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-5 py-2.5 rounded-xl transition-colors shadow-md shadow-emerald-200"
                >
                  <span className="text-lg leading-none">+</span>
                  <span>Registrar Pago</span>
                </button>
              </div>
            </div>

            <div className="grid grid-cols-5 gap-4 mb-6">
              <div className="bg-white rounded-2xl border border-slate-200 p-4">
                <p className="text-xs font-bold text-slate-500 uppercase tracking-wide mb-1">Clientes Ganados</p>
                <p className="text-3xl font-black text-slate-900">{wonClients.length}</p>
              </div>
              <div className="bg-white rounded-2xl border border-slate-200 p-4">
                <p className="text-xs font-bold text-slate-500 uppercase tracking-wide mb-1">Valor Total Ganado</p>
                <p className="text-3xl font-black text-red-700">{formatCurrency(wonClients.reduce((s, c) => s + c.value, 0))}</p>
              </div>
              <div className="bg-white rounded-2xl border border-slate-200 p-4">
                <p className="text-xs font-bold text-slate-500 uppercase tracking-wide mb-1">Total Cobrado</p>
                <p className="text-3xl font-black text-emerald-700">{formatCurrency(payments.reduce((s, p) => s + p.paid, 0))}</p>
              </div>
              <div className="bg-white rounded-2xl border border-slate-200 p-4">
                <p className="text-xs font-bold text-slate-500 uppercase tracking-wide mb-1">Saldo Pendiente</p>
                <p className={`text-3xl font-black ${totalPending > 0 ? "text-red-600" : "text-slate-400"}`}>{formatCurrency(totalPending)}</p>
              </div>
              <div className="bg-white rounded-2xl border border-slate-200 p-4">
                <p className="text-xs font-bold text-slate-500 uppercase tracking-wide mb-1">Facturas Enviadas</p>
                <p className="text-3xl font-black text-emerald-600">{payments.filter(p => p.billInvoiceStatus === "Factura Enviada").length}</p>
              </div>
            </div>

            {wonClients.length === 0 && (
              <div className="text-center py-16 text-slate-400">
                <div className="text-5xl mb-3">🏆</div>
                <p className="font-semibold text-slate-600">No hay clientes con oportunidades ganadas aún</p>
                <p className="text-sm">Mueve clientes a estado "Ganado" en el Pipeline para registrar pagos</p>
              </div>
            )}

            {wonClients.length > 0 && (
              <div className="rounded-2xl border border-slate-200 overflow-hidden bg-white">
                <div className="px-5 py-3 bg-slate-50 border-b border-slate-200">
                  <h3 className="font-bold text-slate-700 text-sm">Registros de Pago</h3>
                </div>
                {payments.length === 0 ? (
                  <div className="text-center py-12 text-slate-400">
                    <div className="text-4xl mb-2">💳</div>
                    <p>No hay pagos registrados. Usa el botón "Registrar Pago" para comenzar.</p>
                  </div>
                ) : (
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b border-slate-100">
                        {["Cliente", "Monto Total", "Pagado / Progreso", "Saldo Pendiente", "Fecha", "Método", "Estado Pago", "Estado Factura", "Acción"].map(h => (
                          <th key={h} className="px-4 py-3 text-left text-xs font-bold text-slate-500 uppercase tracking-wide">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {payments.map(p => (
                        <PaymentRow key={p.id} payment={p} client={clients.find(c => c.id === p.clientId)} onEdit={(pay) => { setEditPayment(pay); setShowPayForm(true); }} />
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            )}

            {wonClients.filter(c => !payments.find(p => p.clientId === c.id)).length > 0 && (
              <div className="mt-4 rounded-2xl border-2 border-dashed border-amber-300 bg-amber-50 p-4">
                <div className="flex items-center gap-2 mb-3">
                  <div className="text-amber-600"><AlertIcon /></div>
                  <h3 className="font-bold text-amber-800 text-sm">Clientes Ganados sin Pago Registrado</h3>
                </div>
                <div className="flex flex-wrap gap-2">
                  {wonClients.filter(c => !payments.find(p => p.clientId === c.id)).map(c => (
                    <div key={c.id} className="bg-white rounded-xl border border-amber-200 px-3 py-2 flex items-center gap-2">
                      <span className="text-sm font-semibold text-slate-800">{c.company}</span>
                      <span className="text-xs font-bold text-red-700">{formatCurrency(c.value)}</span>
                      <button
                        onClick={() => { setEditPayment(null); setShowPayForm(true); }}
                        className="text-xs font-bold text-amber-700 hover:text-amber-900 underline transition-colors"
                      >Registrar</button>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {tab === "advisors" && (
          <AdvisorsTab
            advisors={advisors}
            setAdvisors={setAdvisors}
            clients={clients}
            currentUser={currentUser}
            setCurrentUser={setCurrentUser}
          />
        )}
      </main>

      {/* Modals */}
      {showClientForm && (
        <Modal title={editClient ? `Editar: ${editClient.company}` : "Registrar Nuevo Cliente"} onClose={() => { setShowClientForm(false); setEditClient(null); }}>
          <ClientForm initial={editClient} advisors={activeAdvisorNames} onSave={handleSaveClient} onClose={() => { setShowClientForm(false); setEditClient(null); }} />
        </Modal>
      )}
      {showPayForm && (
        <Modal title={editPayment ? "Editar Pago" : "Registrar Nuevo Pago"} onClose={() => { setShowPayForm(false); setEditPayment(null); }}>
          <PaymentForm clients={clients} initial={editPayment} onSave={handleSavePayment} onClose={() => { setShowPayForm(false); setEditPayment(null); }} />
        </Modal>
      )}
      {blockAlert && <BlockAlert advisor={blockAlert} onClose={() => setBlockAlert(null)} />}
    </div>
  );
}
