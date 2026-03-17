const STORAGE_KEY = "apeiron_ops_hub_en_v1";
const INITIAL_DATA = window.INITIAL_DATA || [];
const STATUS = ["New", "In Contact", "Meeting Scheduled", "Proposal Sent", "Client", "Discarded"];
const PRIORITY = ["High", "Medium", "Low"];
const CHANNEL_METHODS = ["Email", "Phone", "LinkedIn", "WhatsApp"];
const COLUMN_FILTER_IDS = [
  "cfId",
  "cfCompany",
  "cfName",
  "cfRole",
  "cfStatus",
  "cfPriority",
  "cfNextAction",
  "cfDate",
  "cfChannel",
  "cfLeadInsight",
  "cfSharePoint",
];

let onlyOverdue = false;
let operationalMode = "contacts";
let currentPage = 1;
let pageSize = 25;
let leads = normalizeLeadChannels(loadData());

function text(v) {
  return String(v ?? "");
}

function norm(v) {
  return text(v).trim().toLowerCase();
}

function escHtml(v) {
  return text(v).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function escAttr(v) {
  return escHtml(v).replaceAll('"', "&quot;");
}

function canonicalChannelToken(token) {
  const t = norm(token);
  if (!t) return "";
  if (t.includes("no contact") || t.includes("sem contato")) return "NO_CONTACT";
  if (t.includes("telefone") || t.includes("phone")) return "Phone";
  if (t.includes("email") || t.includes("e-mail")) return "Email";
  if (t.includes("linkedin")) return "LinkedIn";
  if (t.includes("whatsapp") || t.includes("whats")) return "WhatsApp";
  return "";
}

function parseChannelSet(rawValue) {
  const raw = text(rawValue).trim();
  if (!raw) return new Set();

  const normalized = raw
    .replace(/\s*(\+|,|;|\/|\|)\s*/g, ",")
    .replace(/\s+e\s+/gi, ",")
    .replace(/\s+and\s+/gi, ",");
  const tokens = normalized.split(",").map((x) => x.trim()).filter(Boolean);
  if (!tokens.length) return new Set();

  const out = new Set();
  for (const tk of tokens) {
    const canonical = canonicalChannelToken(tk);
    if (!canonical) continue;
    if (canonical === "NO_CONTACT") {
      out.clear();
      return out;
    }
    out.add(canonical);
  }
  return out;
}

function channelSetToStorage(setLike) {
  const set = new Set(setLike || []);
  const ordered = CHANNEL_METHODS.filter((m) => set.has(m));
  return ordered.length ? ordered.join(", ") : "No contact";
}

function formatChannelLabel(rawValue) {
  return channelSetToStorage(parseChannelSet(rawValue));
}

function leadMatchesChannel(lead, selectedChannel) {
  const desired = canonicalChannelToken(selectedChannel);
  const current = parseChannelSet(lead.channel);
  if (!desired || desired === "NO_CONTACT") return current.size === 0;
  return current.has(desired);
}

function normalizeLeadChannels(data) {
  return (data || []).map((lead) => ({ ...lead, channel: formatChannelLabel(lead.channel) }));
}

function safeUrl(v) {
  const value = text(v).trim();
  if (!value) return "";
  if (/^https?:\/\//i.test(value)) return value;
  return `https://${value}`;
}

function detailField(label, value, options = {}) {
  const raw = text(value).trim();
  const isFull = options.full ? " full" : "";
  let rendered = raw ? escHtml(raw) : "Not available";
  if (raw && options.type === "email") rendered = `<a href="mailto:${escAttr(raw)}">${escHtml(raw)}</a>`;
  if (raw && options.type === "phone") rendered = `<a href="tel:${escAttr(raw)}">${escHtml(raw)}</a>`;
  if (raw && options.type === "url") {
    const href = safeUrl(raw);
    rendered = `<a href="${escAttr(href)}" target="_blank" rel="noopener noreferrer">${escHtml(raw)}</a>`;
  }
  return `<div class="contact-field${isFull}"><div class="k">${escHtml(label)}</div><div class="v">${rendered}</div></div>`;
}

function getLeadById(id) {
  return leads.find((l) => Number(l.id) === Number(id)) || null;
}

function firstNonEmpty(leadsList, key) {
  for (const item of leadsList) {
    const value = text(item[key]).trim();
    if (value) return value;
  }
  return "";
}

function getCompanyLeads(companyName) {
  const wanted = norm(companyName);
  return leads.filter((l) => norm(l.company) === wanted);
}

function companyChannelSummary(companyLeads) {
  const counts = { Email: 0, Phone: 0, LinkedIn: 0, WhatsApp: 0 };
  companyLeads.forEach((lead) => {
    const channels = parseChannelSet(lead.channel);
    CHANNEL_METHODS.forEach((method) => {
      if (channels.has(method)) counts[method] += 1;
    });
  });
  return `Email: ${counts.Email} | Phone: ${counts.Phone} | LinkedIn: ${counts.LinkedIn} | WhatsApp: ${counts.WhatsApp}`;
}

function openCompanyModal(companyName) {
  const companyLeads = getCompanyLeads(companyName);
  if (!companyLeads.length) return;

  const modal = document.getElementById("contactModal");
  const title = document.getElementById("contactModalTitle");
  const body = document.getElementById("contactModalBody");
  if (!modal || !title || !body) return;

  const total = companyLeads.length;
  const inContact = companyLeads.filter((l) => l.status === "In Contact").length;
  const meetings = companyLeads.filter((l) => l.status === "Meeting Scheduled").length;
  const highPriority = companyLeads.filter((l) => l.priority === "High").length;
  const withDirectContact = companyLeads.filter((l) => parseChannelSet(l.channel).size > 0).length;
  const inSharePoint = companyLeads.filter((l) => text(l.sharePointId).trim() !== "").length;

  title.textContent = `Company Details - ${companyName}`;
  body.innerHTML = [
    detailField("Company", companyName),
    detailField("Total Leads", total),
    detailField("In Contact", inContact),
    detailField("Meetings Scheduled", meetings),
    detailField("High Priority Leads", highPriority),
    detailField("Leads With Direct Contact", withDirectContact),
    detailField("Leads In SharePoint", inSharePoint),
    detailField("Main Contact Phone", firstNonEmpty(companyLeads, "accountPhone") || firstNonEmpty(companyLeads, "phone"), { type: "phone" }),
    detailField("Website", firstNonEmpty(companyLeads, "accountWebsite"), { type: "url" }),
    detailField("Sector", firstNonEmpty(companyLeads, "accountSector")),
    detailField("City / State", firstNonEmpty(companyLeads, "accountCityState")),
    detailField("Company Size", firstNonEmpty(companyLeads, "accountSize")),
    detailField("Revenue", firstNonEmpty(companyLeads, "accountRevenue")),
    detailField("Channel Coverage", companyChannelSummary(companyLeads), { full: true }),
    detailField("Account Summary", firstNonEmpty(companyLeads, "accountSummary"), { full: true }),
  ].join("");

  modal.setAttribute("aria-hidden", "false");
}

function openContactModal(id) {
  const lead = getLeadById(id);
  if (!lead) return;
  const modal = document.getElementById("contactModal");
  const title = document.getElementById("contactModalTitle");
  const body = document.getElementById("contactModalBody");
  if (!modal || !title || !body) return;

  title.textContent = `Contact Details - ${lead.name || "Unnamed contact"}`;
  body.innerHTML = [
    detailField("Company", lead.company),
    detailField("Role", lead.role),
    detailField("Email", lead.email, { type: "email" }),
    detailField("Phone", lead.phone, { type: "phone" }),
    detailField("LinkedIn", lead.linkedin, { type: "url" }),
    detailField("Location", lead.location),
    detailField("Department", lead.department),
    detailField("Status", lead.status),
    detailField("Priority", lead.priority),
    detailField("Contact Channel", formatChannelLabel(lead.channel)),
    detailField("Source", lead.source),
    detailField("SharePoint ID", lead.sharePointId),
    detailField("Lead Insight", lead.owner, { full: true }),
    detailField("Notes", lead.notes, { full: true }),
  ].join("");

  modal.setAttribute("aria-hidden", "false");
}

function closeContactModal() {
  const modal = document.getElementById("contactModal");
  if (modal) modal.setAttribute("aria-hidden", "true");
}

function setMsg(message, color) {
  const el = document.getElementById("syncMsg");
  el.textContent = message || "";
  el.style.color = color || "#1f6d41";
}

function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (_) {}
  return INITIAL_DATA;
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(leads));
}

function normalizeDate(s) {
  if (!s) return null;
  const str = text(s).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return new Date(`${str}T00:00:00`);
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) {
    const [d, m, y] = str.split("/");
    return new Date(`${y}-${m}-${d}T00:00:00`);
  }
  const dt = new Date(str);
  return Number.isNaN(dt.valueOf()) ? null : new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
}

function isOverdue(item) {
  const d = normalizeDate(item.nextActionDate);
  if (!d) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return d < today && item.status !== "Client";
}

function fillSelect(selectEl, options, firstLabel) {
  if (!selectEl) return;
  selectEl.innerHTML = "";
  const first = document.createElement("option");
  first.value = "";
  first.textContent = firstLabel;
  selectEl.appendChild(first);
  options.forEach((value) => {
    const op = document.createElement("option");
    op.value = value;
    op.textContent = value;
    selectEl.appendChild(op);
  });
}

function fillFilters() {
  fillSelect(document.getElementById("fStatus"), STATUS, "Status: all");
  fillSelect(document.getElementById("fPrio"), PRIORITY, "Priority: all");
  const channels = getChannelChoices();
  fillSelect(document.getElementById("fChannel"), channels, "Channel: all");

  fillSelect(document.getElementById("cfStatus"), STATUS, "All");
  fillSelect(document.getElementById("cfPriority"), PRIORITY, "All");
  fillSelect(document.getElementById("cfChannel"), channels, "All");
}

function getChannelChoices() {
  return ["No contact", ...CHANNEL_METHODS];
}

function duplicateKeyForLead(lead) {
  const nameKey = norm(lead?.name);
  const companyKey = norm(lead?.company);
  if (!nameKey || !companyKey) return "";
  return `${nameKey}|${companyKey}`;
}

function buildDuplicateKeySet(data) {
  const counts = new Map();
  (data || []).forEach((lead) => {
    const key = duplicateKeyForLead(lead);
    if (!key) return;
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  return new Set([...counts.entries()].filter(([, count]) => count > 1).map(([key]) => key));
}

function paginateRows(rows) {
  const total = rows.length;
  const size = Number(pageSize) || 0;
  if (size <= 0) {
    currentPage = 1;
    return { rows, total, totalPages: 1, page: 1, start: total ? 1 : 0, end: total };
  }
  const totalPages = Math.max(1, Math.ceil(total / size));
  currentPage = Math.min(Math.max(currentPage, 1), totalPages);
  const startIndex = (currentPage - 1) * size;
  const pageRows = rows.slice(startIndex, startIndex + size);
  return {
    rows: pageRows,
    total,
    totalPages,
    page: currentPage,
    start: total ? startIndex + 1 : 0,
    end: startIndex + pageRows.length,
  };
}

function getColumnFilters() {
  return {
    id: norm(document.getElementById("cfId")?.value),
    company: norm(document.getElementById("cfCompany")?.value),
    name: norm(document.getElementById("cfName")?.value),
    role: norm(document.getElementById("cfRole")?.value),
    status: norm(document.getElementById("cfStatus")?.value),
    priority: norm(document.getElementById("cfPriority")?.value),
    nextAction: norm(document.getElementById("cfNextAction")?.value),
    date: norm(document.getElementById("cfDate")?.value),
    channel: norm(document.getElementById("cfChannel")?.value),
    leadInsight: norm(document.getElementById("cfLeadInsight")?.value),
    sharePoint: norm(document.getElementById("cfSharePoint")?.value),
  };
}

function matchesColumnFilters(lead, f) {
  if (f.id && !norm(lead.id).includes(f.id)) return false;
  if (f.company && !norm(lead.company).includes(f.company)) return false;
  if (f.name && !norm(lead.name).includes(f.name)) return false;
  if (f.role && !norm(lead.role).includes(f.role)) return false;
  if (f.status && norm(lead.status) !== f.status) return false;
  if (f.priority && norm(lead.priority) !== f.priority) return false;
  if (f.nextAction && !norm(lead.nextAction).includes(f.nextAction)) return false;
  if (f.date && !norm(lead.nextActionDate).includes(f.date)) return false;
  if (f.channel && !leadMatchesChannel(lead, f.channel)) return false;
  if (f.leadInsight && !norm(lead.owner).includes(f.leadInsight)) return false;

  const inSharePoint = text(lead.sharePointId).trim() !== "";
  if (f.sharePoint === "in" && !inSharePoint) return false;
  if (f.sharePoint === "pending" && inSharePoint) return false;

  return true;
}

function filtered() {
  const q = norm(document.getElementById("q").value);
  const fs = document.getElementById("fStatus").value;
  const fp = document.getElementById("fPrio").value;
  const fc = document.getElementById("fChannel").value;
  const colFilters = getColumnFilters();

  return leads.filter((lead) => {
    const txt = [lead.name, lead.company, lead.role, lead.owner, lead.nextAction].map((v) => norm(v)).join(" ");
    if (q && !txt.includes(q)) return false;
    if (fs && lead.status !== fs) return false;
    if (fp && lead.priority !== fp) return false;
    if (fc && !leadMatchesChannel(lead, fc)) return false;
    if (onlyOverdue && !isOverdue(lead)) return false;
    if (!matchesColumnFilters(lead, colFilters)) return false;
    return true;
  });
}

function renderKPIs(data) {
  const uniqueCompanies = new Set(data.map((l) => norm(l.company)).filter(Boolean)).size;
  const directContact = data.filter((l) => parseChannelSet(l.channel).size > 0).length;
  const contactCoverage = data.length ? Math.round((directContact / data.length) * 100) : 0;
  const duplicateNameCompanyKeys = buildDuplicateKeySet(data).size;

  const kpis = [
    ["Total Leads", data.length],
    ["Unique Companies", uniqueCompanies],
    ["High Priority", data.filter((l) => l.priority === "High").length],
    ["In Contact", data.filter((l) => l.status === "In Contact").length],
    ["Meeting Scheduled", data.filter((l) => l.status === "Meeting Scheduled").length],
    ["In SharePoint", data.filter((l) => text(l.sharePointId).trim() !== "").length],
    ["No Contact", data.filter((l) => parseChannelSet(l.channel).size === 0).length],
    ["Direct Contact Coverage", `${contactCoverage}%`],
    ["Duplicate Name+Company", duplicateNameCompanyKeys],
    ["Overdue Actions", data.filter(isOverdue).length],
  ];

  document.getElementById("kpis").innerHTML = kpis
    .map(([title, value]) => `<div class="kpi"><div class="t">${title}</div><div class="v">${value}</div></div>`)
    .join("");
}

function pClass(p) {
  return p === "High" ? "high" : p === "Medium" ? "medium" : "low";
}

function renderKanban(data) {
  const box = document.getElementById("kanban");
  box.innerHTML = "";

  STATUS.forEach((status) => {
    const arr = data.filter((l) => l.status === status);
    const lane = document.createElement("div");
    lane.className = "lane";
    lane.innerHTML = `<div class="head"><span>${status}</span><span>${arr.length}</span></div><div class="list"></div>`;
    const list = lane.querySelector(".list");

    arr.slice(0, 40).forEach((l) => {
      list.insertAdjacentHTML(
        "beforeend",
        `<div class="card">
          <div class="name"><button class="contact-link" type="button" data-open-contact="${escAttr(l.id)}">${escHtml(l.name || "(No name)")}</button> - <button class="company-link" type="button" data-open-company="${escAttr(l.company || "")}">${escHtml(l.company || "")}</button></div>
          <div class="meta">${escHtml(l.role || "")}</div>
          <span class="pill ${pClass(l.priority)}">${escHtml(l.priority || "No priority")}</span>
          <div class="meta" style="margin-top:4px">${l.nextActionDate ? `Next: ${escHtml(l.nextActionDate)}` : "No date"}${isOverdue(l) ? " - OVERDUE" : ""}</div>
        </div>`
      );
    });

    box.appendChild(lane);
  });
}

function renderAgenda(data) {
  const arr = [...data]
    .filter(isOverdue)
    .sort((a, b) => text(a.nextActionDate).localeCompare(text(b.nextActionDate)))
    .slice(0, 12);
  const el = document.getElementById("agenda");

  el.innerHTML = arr.length
    ? arr
        .map(
          (l) =>
            `<div class="agenda-item"><strong>${escHtml(l.name)}</strong> - ${escHtml(l.company)} - ${escHtml(
              l.nextActionDate || "no date"
            )} - ${escHtml(l.nextAction || "no action")}</div>`
        )
        .join("")
    : '<div class="agenda-item">No overdue actions in current filters.</div>';
}

function stOpts(selected) {
  return STATUS.map((s) => `<option ${s === selected ? "selected" : ""}>${s}</option>`).join("");
}

function prOpts(selected) {
  return PRIORITY.map((s) => `<option ${s === selected ? "selected" : ""}>${s}</option>`).join("");
}

function renderChannelEditor(rawValue) {
  const selected = parseChannelSet(rawValue);
  return `<div class="channel-multi">${CHANNEL_METHODS.map(
    (method) =>
      `<label class="channel-option"><input type="checkbox" data-k="channelMethod" data-method="${method}" ${
        selected.has(method) ? "checked" : ""
      } /><span>${method}</span></label>`
  ).join("")}</div>`;
}

function parseDateForSort(value) {
  const d = normalizeDate(value);
  return d ? d.getTime() : Number.POSITIVE_INFINITY;
}

function priorityRank(priority) {
  if (priority === "High") return 0;
  if (priority === "Medium") return 1;
  if (priority === "Low") return 2;
  return 3;
}

function pickPrimaryLead(groupLeads) {
  return [...groupLeads].sort((a, b) => {
    const overdueA = isOverdue(a) ? 0 : 1;
    const overdueB = isOverdue(b) ? 0 : 1;
    if (overdueA !== overdueB) return overdueA - overdueB;
    const prA = priorityRank(a.priority);
    const prB = priorityRank(b.priority);
    if (prA !== prB) return prA - prB;
    const dtA = parseDateForSort(a.nextActionDate);
    const dtB = parseDateForSort(b.nextActionDate);
    if (dtA !== dtB) return dtA - dtB;
    return Number(a.id || 0) - Number(b.id || 0);
  })[0];
}

function summarizeNames(groupLeads) {
  const names = [...new Set(groupLeads.map((l) => text(l.name).trim()).filter(Boolean))];
  if (!names.length) return "No named contacts";
  if (names.length <= 2) return names.join(" | ");
  return `${names.slice(0, 2).join(" | ")} +${names.length - 2} more`;
}

function summarizeRoles(groupLeads) {
  const roles = [...new Set(groupLeads.map((l) => text(l.role).trim()).filter(Boolean))];
  if (!roles.length) return "";
  if (roles.length <= 2) return roles.join(" | ");
  return `${roles.slice(0, 2).join(" | ")} +${roles.length - 2} more`;
}

function buildCompanyRows(data) {
  const groups = new Map();
  data.forEach((lead) => {
    const companyValue = text(lead.company).trim();
    const key = norm(companyValue) || `__no_company__${lead.id}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(lead);
  });

  return [...groups.values()].map((groupLeads) => {
    const primary = pickPrimaryLead(groupLeads);
    const idsLabel = groupLeads.map((l) => l.id).filter(Boolean).join(", ");

    const unionChannels = new Set();
    groupLeads.forEach((l) => {
      parseChannelSet(l.channel).forEach((method) => unionChannels.add(method));
    });

    const sharePointIds = groupLeads.map((l) => text(l.sharePointId).trim()).filter(Boolean);
    return {
      ...primary,
      id: idsLabel,
      name: summarizeNames(groupLeads),
      role: summarizeRoles(groupLeads),
      channel: channelSetToStorage(unionChannels),
      sharePointId: sharePointIds[0] || "",
      _groupLeads: groupLeads,
      _companyCount: groupLeads.length,
    };
  });
}

function updateOperationalHeaders() {
  const idHeader = document.getElementById("hId");
  const nameHeader = document.getElementById("hName");
  const roleHeader = document.getElementById("hRole");
  if (!idHeader || !nameHeader || !roleHeader) return;

  if (operationalMode === "companies") {
    idHeader.textContent = "ID(s)";
    nameHeader.textContent = "Contacts";
    roleHeader.textContent = "Roles";
  } else {
    idHeader.textContent = "ID";
    nameHeader.textContent = "Name";
    roleHeader.textContent = "Role";
  }
}

function renderTable(data) {
  const tb = document.getElementById("tbody");
  tb.innerHTML = "";
  const duplicateKeySet = operationalMode === "contacts" ? buildDuplicateKeySet(data) : new Set();

  data.forEach((l) => {
    const companyGroup = Array.isArray(l._groupLeads) ? l._groupLeads : null;
    const targetLeads = companyGroup || [l];
    const rowOverdue = companyGroup ? companyGroup.some(isOverdue) : isOverdue(l);
    const inSharePointCount = targetLeads.filter((x) => text(x.sharePointId).trim() !== "").length;

    const tr = document.createElement("tr");
    if (rowOverdue) tr.classList.add("overdue");

    const nameCell = companyGroup
      ? `<strong>${companyGroup.length} contact${companyGroup.length > 1 ? "s" : ""}</strong><span class="company-subline">${escHtml(
          l.name || "No named contacts"
        )}</span>`
      : `<button class="contact-link" type="button" data-open-contact="${escAttr(l.id)}">${escHtml(l.name || "")}</button>${
          duplicateKeySet.has(duplicateKeyForLead(l)) ? '<span class="dup-flag">Possible duplicate</span>' : ""
        }`;

    const sharePointTag =
      inSharePointCount === 0
        ? '<span class="tag wait">Pending</span>'
        : inSharePointCount === targetLeads.length
          ? '<span class="tag share">In SharePoint</span>'
          : '<span class="tag wait">Partial</span>';

    tr.innerHTML = `<td>${escHtml(l.id)}</td>
      <td><button class="company-link" type="button" data-open-company="${escAttr(l.company || "")}">${escHtml(l.company || "")}</button></td>
      <td>${nameCell}</td>
      <td>${escHtml(l.role || "")}</td>
      <td><select class="inline-select" data-k="status">${stOpts(l.status)}</select></td>
      <td><select class="inline-select" data-k="priority">${prOpts(l.priority)}</select></td>
      <td><textarea class="inline-textarea" data-k="nextAction">${escHtml(l.nextAction || "")}</textarea></td>
      <td><input class="inline-input" data-k="nextActionDate" value="${escAttr(l.nextActionDate || "")}" /></td>
      <td>${renderChannelEditor(l.channel)}</td>
      <td><textarea class="inline-textarea" data-k="owner">${escHtml(l.owner || "")}</textarea></td>
      <td>${sharePointTag}</td>`;

    tr.querySelectorAll('[data-k]:not([data-k="channelMethod"])').forEach((inp) =>
      inp.addEventListener("change", () => {
        const key = inp.getAttribute("data-k");
        targetLeads.forEach((lead) => {
          lead[key] = inp.value;
        });
        persist();
        refresh();
      })
    );

    tr.querySelectorAll('[data-k="channelMethod"]').forEach((inp) =>
      inp.addEventListener("change", () => {
        const selected = [...tr.querySelectorAll('[data-k="channelMethod"]:checked')].map((cb) => cb.getAttribute("data-method"));
        targetLeads.forEach((lead) => {
          lead.channel = channelSetToStorage(selected);
        });
        persist();
        refresh();
      })
    );

    tb.appendChild(tr);
  });
}

function renderPagination(meta) {
  const info = document.getElementById("pageInfo");
  const prev = document.getElementById("btnPrevPage");
  const next = document.getElementById("btnNextPage");
  if (!info || !prev || !next) return;

  const totalPages = meta.totalPages || 1;
  const page = meta.page || 1;
  info.textContent = meta.total
    ? `Showing ${meta.start}-${meta.end} of ${meta.total} | Page ${page}/${totalPages}`
    : "No records";

  prev.disabled = page <= 1 || !meta.total;
  next.disabled = page >= totalPages || !meta.total;
}

function refresh() {
  const data = filtered();
  renderKPIs(data);
  renderKanban(data);
  renderAgenda(data);
  updateOperationalHeaders();
  const tableData = operationalMode === "companies" ? buildCompanyRows(data) : data;
  const pageMeta = paginateRows(tableData);
  renderTable(pageMeta.rows);
  renderPagination(pageMeta);
}

function exportCsv() {
  const cols = ["id", "company", "name", "role", "status", "priority", "nextAction", "nextActionDate", "channel", "owner", "source", "sharePointId"];
  const lines = [cols.join(",")].concat(filtered().map((l) => cols.map((k) => `"${text(l[k]).replaceAll('"', '""')}"`).join(",")));
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "apeiron_operations_en.csv";
  a.click();
}

function headerMap(sheet, headerRow) {
  const map = {};
  for (let c = 1; c <= 120; c += 1) {
    const v = sheet.cell(headerRow, c).value();
    if (v !== undefined && v !== null && text(v).trim() !== "") map[text(v).trim()] = c;
  }
  return map;
}

function findCol(headers, names) {
  for (const name of names) if (headers[name]) return headers[name];
  return null;
}

function setv(sheet, row, col, value) {
  if (!col) return;
  sheet.cell(row, col).value(value === undefined || value === null ? "" : value);
}

function updateById(sheet, headerRow, leadMap) {
  const headers = headerMap(sheet, headerRow);
  const idCol = findCol(headers, ["ID"]);
  if (!idCol) return 0;

  const used = sheet.usedRange();
  const maxRow = used ? used.endCell().rowNumber() : headerRow + 1;
  let updated = 0;

  for (let r = headerRow + 1; r <= maxRow; r += 1) {
    const id = Number(sheet.cell(r, idCol).value());
    if (!Number.isFinite(id)) continue;
    const lead = leadMap.get(id);
    if (!lead) continue;
    const chSet = parseChannelSet(lead.channel);

    setv(sheet, r, findCol(headers, ["Status"]), lead.status);
    setv(sheet, r, findCol(headers, ["Priority"]), lead.priority);
    setv(sheet, r, findCol(headers, ["Channel", "Contact Channel"]), formatChannelLabel(lead.channel));
    setv(sheet, r, findCol(headers, ["Channel Email"]), chSet.has("Email") ? "Yes" : "");
    setv(sheet, r, findCol(headers, ["Channel Phone"]), chSet.has("Phone") ? "Yes" : "");
    setv(sheet, r, findCol(headers, ["Channel LinkedIn"]), chSet.has("LinkedIn") ? "Yes" : "");
    setv(sheet, r, findCol(headers, ["Channel WhatsApp"]), chSet.has("WhatsApp") ? "Yes" : "");
    setv(sheet, r, findCol(headers, ["Next Action"]), lead.nextAction);
    setv(sheet, r, findCol(headers, ["Next Action Date"]), lead.nextActionDate);
    setv(sheet, r, findCol(headers, ["Lead Insight", "Owner", "Vendedor"]), lead.owner);
    updated += 1;
  }
  return updated;
}

async function saveToExcel(file) {
  try {
    if (!window.XlsxPopulate) {
      setMsg("Spreadsheet library unavailable.", "#b71c1c");
      return;
    }
    setMsg("Processing file...", "#1f6d41");
    const arrayBuffer = await file.arrayBuffer();
    const wb = await XlsxPopulate.fromDataAsync(arrayBuffer);
    const leadMap = new Map(leads.map((l) => [Number(l.id), l]));

    let updated = 0;
    const operations = wb.sheet("Leads_Operations") || wb.sheet("Leads_Operacao");
    if (operations) updated += updateById(operations, 1, leadMap);

    wb.sheets().forEach((sheet) => {
      const name = sheet.name();
      if (name && (name.includes("Contacts") || name.includes("Contatos"))) updated += updateById(sheet, 3, leadMap);
    });

    const out = await wb.outputAsync();
    if (window.showSaveFilePicker) {
      const handle = await window.showSaveFilePicker({
        suggestedName: file.name,
        types: [
          {
            description: "Excel Workbook",
            accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] },
          },
        ],
      });
      const writable = await handle.createWritable();
      await writable.write(out);
      await writable.close();
    } else {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(out);
      a.download = file.name.replace(/\.xlsx$/i, "") + "_updated.xlsx";
      a.click();
    }
    setMsg(`Synced successfully (${updated} rows updated).`, "#1e8449");
  } catch (e) {
    console.error(e);
    setMsg("Failed to save to Excel. Please try again.", "#b71c1c");
  }
}

function wireUi() {
  const refreshFromFirstPage = () => {
    currentPage = 1;
    refresh();
  };

  document.getElementById("q").addEventListener("input", refreshFromFirstPage);
  document.getElementById("fStatus").addEventListener("change", refreshFromFirstPage);
  document.getElementById("fPrio").addEventListener("change", refreshFromFirstPage);
  document.getElementById("fChannel").addEventListener("change", refreshFromFirstPage);
  document.getElementById("operationalMode").addEventListener("change", (event) => {
    operationalMode = event.target.value === "companies" ? "companies" : "contacts";
    currentPage = 1;
    refresh();
  });
  document.getElementById("rowsPerPage").addEventListener("change", (event) => {
    const value = Number(event.target.value);
    pageSize = Number.isFinite(value) ? value : 25;
    currentPage = 1;
    refresh();
  });
  document.getElementById("btnPrevPage").addEventListener("click", () => {
    currentPage = Math.max(1, currentPage - 1);
    refresh();
  });
  document.getElementById("btnNextPage").addEventListener("click", () => {
    currentPage += 1;
    refresh();
  });

  COLUMN_FILTER_IDS.forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("input", refreshFromFirstPage);
    el.addEventListener("change", refreshFromFirstPage);
  });

  document.getElementById("btnOverdue").addEventListener("click", () => {
    onlyOverdue = !onlyOverdue;
    document.getElementById("btnOverdue").textContent = onlyOverdue ? "Show all" : "Only overdue";
    currentPage = 1;
    refresh();
  });
  document.getElementById("btnExport").addEventListener("click", exportCsv);
  document.getElementById("btnSaveExcel").addEventListener("click", () => document.getElementById("excelFile").click());
  document.getElementById("excelFile").addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    await saveToExcel(file);
    e.target.value = "";
  });

  document.addEventListener("click", (event) => {
    const contactTarget = event.target.closest("[data-open-contact]");
    if (contactTarget) {
      openContactModal(contactTarget.getAttribute("data-open-contact"));
      return;
    }

    const companyTarget = event.target.closest("[data-open-company]");
    if (companyTarget) {
      const company = companyTarget.getAttribute("data-open-company");
      if (text(company).trim()) openCompanyModal(company);
    }
  });

  document.getElementById("contactModalClose").addEventListener("click", closeContactModal);
  document.getElementById("contactModal").addEventListener("click", (event) => {
    if (event.target && event.target.getAttribute("data-close-modal") === "true") closeContactModal();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeContactModal();
  });
}

fillFilters();
wireUi();
refresh();
