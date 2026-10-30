const REQUIRED_SHEETS = [
  "ASIGNACION RANKING",
  "COMISIONES",
  "INFORME H A H",
  "DATOS ASESORES",
  "PUESTOS",
  "DOMINGO 1",
  "DOMINGO 2",
  "FESTIVO 1"
];

const dashboard = document.getElementById("main-menu");
const appContent = document.getElementById("app-content");
const uploadGate = document.getElementById("upload-gate");
const uploadMessage = document.getElementById("upload-message");
const fileInput = document.getElementById("excel-input");
const navigationOptions = document.querySelectorAll(".nav-option");

let currentData = null;
let selectedTeamM1 = "ALL";
let selectedAppM2 = "ALL";
let selectedTierM3 = "ALL";
let selectedTeamM4 = "ALL";

dashboard.addEventListener("transitionend", event => {
  if (event.propertyName === "opacity" && dashboard.classList.contains("is-leaving")) {
    dashboard.hidden = true;
  }
});

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, character => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[character]);
}

function textValue(value) {
  if (value === null || value === undefined || value === "") return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toLocaleDateString("es-CO");
  }
  return String(value).trim();
}

function numberValue(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string") return null;

  const cleaned = value.trim().replace(/[$\s]/g, "").replace(/,/g, "");
  if (!cleaned || /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)/i.test(cleaned)) return null;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function sumNumbers(values) {
  const numbers = values.filter(value => value !== null && value !== undefined);
  return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) : null;
}

function formatNumber(value) {
  return value === null || value === undefined
    ? "—"
    : new Intl.NumberFormat("es-CO", { maximumFractionDigits: 0 }).format(value);
}

function formatMoney(value) {
  return value === null || value === undefined
    ? "—"
    : new Intl.NumberFormat("es-CO", {
      style: "currency",
      currency: "COP",
      maximumFractionDigits: 0
    }).format(value);
}

function formatRate(value) {
  return value === null || value === undefined || !Number.isFinite(value)
    ? "—"
    : `${value.toFixed(2)}%`;
}

function renderDashboardPieChart(containerId, slices, total, centerValue, centerLabel, emptyMessage) {
  const container = document.getElementById(containerId);
  const visibleSlices = slices.filter(slice => Number.isFinite(slice.value) && slice.value > 0);

  if (!Number.isFinite(total) || total <= 0 || !visibleSlices.length) {
    container.innerHTML = `
      <div class="chart-content">
        <div class="pie-chart" role="img" aria-label="${escapeHtml(emptyMessage)}" style="background: conic-gradient(#39445a 0deg 360deg)">
          <div class="pie-chart-center"><strong>—</strong><span>sin datos</span></div>
        </div>
        <p class="chart-legend-label">${escapeHtml(emptyMessage)}</p>
      </div>
    `;
    return;
  }

  let accumulated = 0;
  const gradient = visibleSlices.map((slice, index) => {
    const start = accumulated;
    accumulated += (slice.value / total) * 360;
    const end = index === visibleSlices.length - 1 ? 360 : accumulated;
    return `${slice.color} ${start}deg ${end}deg`;
  }).join(", ");
  const description = visibleSlices
    .map(slice => `${slice.label}: ${formatRate((slice.value / total) * 100)}`)
    .join("; ");

  container.innerHTML = `
    <div class="chart-content">
      <div class="pie-chart" role="img" aria-label="${escapeHtml(description)}" style="background: conic-gradient(${gradient})">
        <div class="pie-chart-center">
          <strong>${escapeHtml(centerValue)}</strong>
          <span>${escapeHtml(centerLabel)}</span>
        </div>
      </div>
      <ul class="chart-legend">
        ${visibleSlices.map(slice => `
          <li>
            <span class="chart-legend-swatch" style="background: ${slice.color}; color: ${slice.color}" aria-hidden="true"></span>
            <span class="chart-legend-label">
              ${escapeHtml(slice.label)}
              <span class="chart-legend-value">${escapeHtml(slice.formattedValue)} · ${formatRate((slice.value / total) * 100)}</span>
            </span>
          </li>
        `).join("")}
      </ul>
    </div>
  `;
}

function getInitials(value) {
  return String(value ?? "")
    .replace(/\([^)]*\)/g, "")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map(part => part[0].toUpperCase())
    .join("") || "—";
}

function sheetMatrix(workbook, name) {
  const actualName = workbook.SheetNames.find(sheetName =>
    normalizeText(sheetName) === normalizeText(name)
  );
  if (!actualName) throw new Error(`No se encontró la hoja "${name}".`);
  return XLSX.utils.sheet_to_json(workbook.Sheets[actualName], {
    header: 1,
    raw: true,
    defval: null
  });
}

function excelErrorAt(workbook, sheetName, rowIndex, columnIndex) {
  const actualName = workbook.SheetNames.find(name =>
    normalizeText(name) === normalizeText(sheetName)
  );
  const sheet = actualName ? workbook.Sheets[actualName] : null;
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
  const cell = sheet ? sheet[address] : null;
  if (!cell || cell.t !== "e") return "";
  const knownErrors = {
    0: "#NULL!",
    7: "#DIV/0!",
    15: "#VALUE!",
    23: "#REF!",
    29: "#NAME?",
    36: "#NUM!",
    42: "#N/A",
    43: "#GETTING_DATA"
  };
  return cell.w || (typeof cell.v === "string" && cell.v.startsWith("#") ? cell.v : "") || knownErrors[cell.v] || "#ERROR";
}

function findColumn(headers, candidates, fallback, sheetName) {
  const normalizedHeaders = headers.map(normalizeText);
  for (const candidate of candidates) {
    const target = normalizeText(candidate);
    const index = normalizedHeaders.findIndex(header => header.includes(target));
    if (index >= 0) return index;
  }
  if (fallback !== undefined && fallback < headers.length) return fallback;
  throw new Error(`No se encontró la columna "${candidates[0]}" en "${sheetName}".`);
}

function valueAt(row, index) {
  return index < 0 || index >= row.length ? null : row[index];
}

function recordsFromSheet(workbook, sheetName, headerRow = 0) {
  const matrix = sheetMatrix(workbook, sheetName);
  const headers = matrix[headerRow] || [];
  return {
    headers,
    rows: matrix.slice(headerRow + 1).filter(row => row.some(value => value !== null && value !== ""))
  };
}

function parseAssignment(workbook) {
  const { headers, rows } = recordsFromSheet(workbook, "ASIGNACION RANKING");
  const columns = {
    date: findColumn(headers, ["fecha"], 0, "ASIGNACION RANKING"),
    company: findColumn(headers, ["empresa de cobranza"], 1, "ASIGNACION RANKING"),
    gestor: findColumn(headers, ["gestor de cobranza"], 3, "ASIGNACION RANKING"),
    team: findColumn(headers, ["equipo"], 4, "ASIGNACION RANKING"),
    accounts: findColumn(headers, ["pedidos en cartera"], 8, "ASIGNACION RANKING"),
    portfolio: findColumn(headers, ["importe de pedidos en cartera"], 9, "ASIGNACION RANKING"),
    received: findColumn(headers, ["importe efectivamente recibido"], 13, "ASIGNACION RANKING")
  };

  const advisors = rows.map(row => ({
    fecha: textValue(valueAt(row, columns.date)),
    empresa: textValue(valueAt(row, columns.company)),
    gestor: textValue(valueAt(row, columns.gestor)),
    equipo: textValue(valueAt(row, columns.team)),
    cuentas: numberValue(valueAt(row, columns.accounts)),
    cartera: numberValue(valueAt(row, columns.portfolio)),
    recibido: numberValue(valueAt(row, columns.received))
  })).filter(advisor => advisor.gestor);

  if (!advisors.length) {
    throw new Error('La hoja "ASIGNACION RANKING" no contiene registros de gestores.');
  }

  const teams = new Map();
  for (const advisor of advisors) {
    const key = advisor.equipo || "Sin equipo";
    if (!teams.has(key)) {
      teams.set(key, {
        equipo: key,
        empresa: advisor.empresa,
        asesores: new Set(),
        cuentas: 0,
        cartera: 0,
        recibido: 0
      });
    }
    const team = teams.get(key);
    team.asesores.add(advisor.gestor);
    team.cuentas += advisor.cuentas || 0;
    team.cartera += advisor.cartera || 0;
    team.recibido += advisor.recibido || 0;
  }

  const summary = [...teams.values()].map(team => ({
    ...team,
    asesores: team.asesores.size
  })).sort((a, b) => a.equipo.localeCompare(b.equipo, "es"));

  return { advisors, summary };
}

function parseCommissions(workbook) {
  const matrix = sheetMatrix(workbook, "COMISIONES");
  const headers = matrix[0] || [];
  const rows = matrix.slice(1).map((row, index) => ({ row, sourceRow: index + 1 }))
    .filter(({ row }) => row.some(value => value !== null && value !== ""));
  const columns = {
    segment: findColumn(headers, ["segmento"], 0, "COMISIONES"),
    app: findColumn(headers, ["aplicacion"], 1, "COMISIONES"),
    name: findColumn(headers, ["nombre sistema", "ejecutivo"], 2, "COMISIONES"),
    leader: findColumn(headers, ["tlj"], 5, "COMISIONES"),
    base: findColumn(headers, ["monto de pedidos en base"], 8, "COMISIONES"),
    received: findColumn(headers, ["monto recuperado"], 9, "COMISIONES"),
    sunday1Base: findColumn(headers, ["domingo 1"], 10, "COMISIONES"),
    sunday2Base: findColumn(headers, ["domingo 2"], 11, "COMISIONES"),
    sunday1Received: findColumn(headers, ["pagos domingo1"], 12, "COMISIONES"),
    sunday2Received: findColumn(headers, ["pagos domingo2"], 13, "COMISIONES")
  };

  return rows.map(({ row, sourceRow }) => {
    const sundayBaseValues = [
      valueAt(row, columns.sunday1Base),
      valueAt(row, columns.sunday2Base)
    ];
    const sundayReceivedValues = [
      valueAt(row, columns.sunday1Received),
      valueAt(row, columns.sunday2Received)
    ];
    const sundayBase = sumNumbers(sundayBaseValues.map(numberValue));
    const sundayReceived = sumNumbers(sundayReceivedValues.map(numberValue));
    return {
      seg: textValue(valueAt(row, columns.segment)),
      app: textValue(valueAt(row, columns.app)),
      nombre: textValue(valueAt(row, columns.name)),
      tlj: textValue(valueAt(row, columns.leader)),
      base: numberValue(valueAt(row, columns.base)),
      baseError: excelErrorAt(workbook, "COMISIONES", sourceRow, columns.base),
      rec: numberValue(valueAt(row, columns.received)),
      recError: excelErrorAt(workbook, "COMISIONES", sourceRow, columns.received),
      d_base: sundayBase,
      d_baseError: sundayBase === null
        ? [columns.sunday1Base, columns.sunday2Base].map(column => excelErrorAt(workbook, "COMISIONES", sourceRow, column)).find(Boolean) || ""
        : "",
      d_rec: sundayReceived,
      d_recError: sundayReceived === null
        ? [columns.sunday1Received, columns.sunday2Received].map(column => excelErrorAt(workbook, "COMISIONES", sourceRow, column)).find(Boolean) || ""
        : ""
    };
  }).filter(record => record.nombre);
}

function parseDailyReport(workbook) {
  const sheetName = "INFORME H A H";
  const matrix = sheetMatrix(workbook, sheetName);
  const headers = matrix[1] || [];
  const columns = {
    team: 8,
    name: findColumn(headers, ["nombre del ejecutivo"], 9, sheetName),
    goal: findColumn(headers, ["objetivo del dia"], 10, sheetName),
    accounts: findColumn(headers, ["cuentas asignadas"], 11, sheetName),
    amount: findColumn(headers, ["monto asignado"], 12, sheetName),
    received: findColumn(headers, ["recaudo"], 13, sheetName)
  };
  let currentTeam = "";
  const advisors = [];

  matrix.slice(2).forEach(row => {
    const teamValue = textValue(valueAt(row, columns.team));
    if (teamValue) currentTeam = teamValue;

    const name = textValue(valueAt(row, columns.name));
    if (!name || normalizeText(name) === "nombre del ejecutivo") return;
    advisors.push({
      ej: name,
      team: currentTeam || "Sin equipo",
      meta: numberValue(valueAt(row, columns.goal)),
      cuentas: numberValue(valueAt(row, columns.accounts)),
      monto: numberValue(valueAt(row, columns.amount)),
      pagado: numberValue(valueAt(row, columns.received))
    });
  });

  return advisors;
}

function parseStaff(workbook) {
  const { headers, rows } = recordsFromSheet(workbook, "DATOS ASESORES");
  const columns = {
    nombre: findColumn(headers, ["nombre"], 0, "DATOS ASESORES"),
    remoto: findColumn(headers, ["estan en teletrabajo"], 1, "DATOS ASESORES"),
    teletrabajo: findColumn(headers, ["puede teletrabajar"], 2, "DATOS ASESORES"),
    moto: findColumn(headers, ["tiene moto"], 3, "DATOS ASESORES"),
    sede2: findColumn(headers, ["registro sede 2"], 4, "DATOS ASESORES"),
    prueba: findColumn(headers, ["periodo de prueba"], 5, "DATOS ASESORES"),
    ingreso: findColumn(headers, ["fecha ingreso"], 7, "DATOS ASESORES"),
    disciplinario: findColumn(headers, ["procesos disciplinario"], 8, "DATOS ASESORES")
  };

  return rows.map(row => ({
    nombre: textValue(valueAt(row, columns.nombre)),
    remoto: textValue(valueAt(row, columns.remoto)),
    puede_tt: textValue(valueAt(row, columns.teletrabajo)),
    moto: textValue(valueAt(row, columns.moto)),
    sede2: textValue(valueAt(row, columns.sede2)),
    prueba: textValue(valueAt(row, columns.prueba)),
    ingreso: textValue(valueAt(row, columns.ingreso)),
    disc: textValue(valueAt(row, columns.disciplinario))
  })).filter(person => person.nombre);
}

function parseSeats(workbook) {
  const matrix = sheetMatrix(workbook, "PUESTOS");
  const zones = new Set(["SALA 1", "SALA 2", "SALA 3", "SEDE 2.0"]);
  const nonSeatLabels = /^(bano|entrada|cafeteria|parqueadero|horarios|break|almuerzo|pausas)\b/i;
  let currentZone = "SALA 1";
  const seats = [];
  const seen = new Set();

  matrix.slice(0, 60).forEach(row => row.forEach(value => {
    const name = textValue(value);
    if (!name) return;
    const zone = [...zones].find(candidate => normalizeText(name) === normalizeText(candidate));
    if (zone) {
      currentZone = zone;
      return;
    }
    if (nonSeatLabels.test(normalizeText(name))) return;
    const key = `${normalizeText(name)}|${currentZone}`;
    if (seen.has(key)) return;
    seen.add(key);
    seats.push({ nombre: name, zona: currentZone });
  }));
  return seats;
}

function parseReportSummary(workbook, sheetName) {
  const { headers, rows } = recordsFromSheet(workbook, sheetName);
  const columns = {
    name: findColumn(headers, ["gestor de cobranza"], 3, sheetName),
    portfolio: findColumn(headers, ["importe de pedidos en cartera"], 9, sheetName),
    received: findColumn(headers, ["importe efectivamente recibido"], 13, sheetName),
    date: findColumn(headers, ["fecha"], 0, sheetName)
  };
  const validRows = rows.filter(row => textValue(valueAt(row, columns.name)));
  const amounts = validRows.map(row => numberValue(valueAt(row, columns.received)));
  const portfolios = validRows.map(row => numberValue(valueAt(row, columns.portfolio)));
  const people = new Set(validRows.map(row => textValue(valueAt(row, columns.name))));
  const firstDate = validRows.map(row => textValue(valueAt(row, columns.date))).find(Boolean) || "";
  return {
    amount: sumNumbers(amounts),
    portfolio: sumNumbers(portfolios),
    people: people.size,
    date: firstDate
  };
}

function parseWorkbook(workbook) {
  const sheetNames = new Set(workbook.SheetNames.map(normalizeText));
  const missingSheets = REQUIRED_SHEETS.filter(name => !sheetNames.has(normalizeText(name)));
  if (missingSheets.length) {
    throw new Error(`Faltan hojas requeridas: ${missingSheets.join(", ")}.`);
  }

  const assignment = parseAssignment(workbook);
  const commissions = parseCommissions(workbook);
  const daily = parseDailyReport(workbook);
  const staff = parseStaff(workbook);
  const seats = parseSeats(workbook);
  const sundayReports = {
    sunday1: parseReportSummary(workbook, "DOMINGO 1"),
    sunday2: parseReportSummary(workbook, "DOMINGO 2"),
    holiday: parseReportSummary(workbook, "FESTIVO 1")
  };

  return {
    sheetCount: workbook.SheetNames.length,
    advisors: assignment.advisors,
    teams: assignment.summary,
    commissions,
    daily,
    staff,
    seats,
    sundayReports,
    assignmentDate: assignment.advisors.map(advisor => advisor.fecha).find(Boolean) || ""
  };
}

function setText(id, value) {
  const element = document.getElementById(id);
  if (element) element.textContent = value;
}

function setCurrentNavigation(moduleId) {
  navigationOptions.forEach(option => {
    if (option.dataset.module === moduleId) {
      option.setAttribute("aria-current", "page");
    } else {
      option.removeAttribute("aria-current");
    }
  });
}

function openModule(moduleId) {
  if (!currentData) return;
  const module = document.getElementById(moduleId);
  if (!module || !module.classList.contains("module-view")) {
    throw new Error(`No se encontró el módulo solicitado: ${moduleId}`);
  }

  dashboard.classList.add("is-leaving");
  document.querySelectorAll(".module-view").forEach(section => {
    section.classList.remove("active");
    section.hidden = true;
  });
  module.hidden = false;
  module.classList.add("active");
  setCurrentNavigation(moduleId);
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function backToMenu() {
  if (!currentData) return;
  document.querySelectorAll(".module-view").forEach(section => {
    section.classList.remove("active");
    section.hidden = true;
  });
  dashboard.hidden = false;
  dashboard.classList.add("is-leaving");
  requestAnimationFrame(() => dashboard.classList.remove("is-leaving"));
  setCurrentNavigation("dashboard");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function showEmptyRow(columns, message) {
  return `<tr><td colspan="${columns}">${escapeHtml(message)}</td></tr>`;
}

function updateGlobalMetrics() {
  const allAdvisors = currentData.advisors;
  const uniqueAdvisors = new Set(allAdvisors.map(advisor => advisor.gestor));
  const accounts = sumNumbers(allAdvisors.map(advisor => advisor.cuentas)) || 0;
  const portfolio = sumNumbers(allAdvisors.map(advisor => advisor.cartera)) || 0;
  const received = sumNumbers(allAdvisors.map(advisor => advisor.recibido)) || 0;
  const rate = portfolio > 0 ? (received / portfolio) * 100 : null;
  const teamCount = currentData.teams.length;

  setText("dashboard-active-count", formatNumber(uniqueAdvisors.size));
  setText("dashboard-team-count", `${formatNumber(teamCount)} equipos`);
  setText("dashboard-portfolio", formatMoney(portfolio));
  setText("dashboard-account-count", `${formatNumber(accounts)} obligaciones asignadas`);
  setText("dashboard-collected", formatMoney(received));
  setText("dashboard-rate", formatRate(rate));
  setText("m1-active-count", formatNumber(uniqueAdvisors.size));
  setText("m1-team-count", `Distribuidos en ${formatNumber(teamCount)} equipos`);
  setText("m1-account-count", formatNumber(accounts));
  setText("m1-portfolio", formatMoney(portfolio));
  setText("m1-collected", formatMoney(received));
  setText("m1-rate", `Tasa de recuperación: ${formatRate(rate)}`);
  setText("m1-source", `📂 Origen: Hoja ASIGNACION RANKING${currentData.assignmentDate ? ` · ${currentData.assignmentDate}` : ""}`);

  const collection = Math.max(0, Math.min(received, portfolio));
  const outstanding = Math.max(portfolio - Math.max(received, 0), 0);
  renderDashboardPieChart("dashboard-collection-chart", [
    { label: "Recaudado", value: collection, formattedValue: formatMoney(received), color: "#37d6a0" },
    { label: "Saldo pendiente", value: outstanding, formattedValue: formatMoney(outstanding), color: "#ffbd58" }
  ], portfolio, formatRate(rate), "recaudado", "No hay cartera asignada para calcular el recaudo.");

  const rankedTeams = [...currentData.teams]
    .filter(team => Number.isFinite(team.cartera) && team.cartera > 0)
    .sort((first, second) => second.cartera - first.cartera);
  const topTeams = rankedTeams.slice(0, 5);
  const teamSlices = topTeams.map((team, index) => ({
    label: team.equipo || "Sin equipo",
    value: team.cartera,
    formattedValue: formatMoney(team.cartera),
    color: ["#66a3ff", "#b18aff", "#ff719c", "#37d6a0", "#ffbd58"][index]
  }));
  const otherPortfolio = rankedTeams.slice(5).reduce((sum, team) => sum + team.cartera, 0);
  if (otherPortfolio > 0) {
    teamSlices.push({
      label: "Otros equipos",
      value: otherPortfolio,
      formattedValue: formatMoney(otherPortfolio),
      color: "#8c9bb5"
    });
  }
  const teamPortfolio = teamSlices.reduce((sum, slice) => sum + slice.value, 0);
  renderDashboardPieChart(
    "dashboard-team-chart",
    teamSlices,
    teamPortfolio,
    formatNumber(rankedTeams.length),
    "con cartera",
    "No hay cartera asignada a los equipos."
  );
}

function selectM1Team(team) {
  selectedTeamM1 = team;
  renderM1();
}

function renderM1() {
  if (!currentData) return;
  const query = normalizeText(document.getElementById("m1-search").value);
  const grid = document.getElementById("m1-teams-grid");
  const teams = currentData.teams;
  const totalCollected = currentData.advisors.reduce((sum, advisor) => sum + (advisor.recibido || 0), 0);
  const cards = [{
    equipo: "ALL",
    label: "🌐 Todos los Equipos",
    detail: `${formatNumber(new Set(currentData.advisors.map(advisor => advisor.gestor)).size)} gestores · ${formatMoney(totalCollected)}`
  }, ...teams.map(team => ({
    equipo: team.equipo,
    label: team.equipo,
    detail: `${formatNumber(team.asesores)} gestores · ${formatMoney(team.recibido)} (${formatRate(team.cartera > 0 ? team.recibido / team.cartera * 100 : null)})`
  }))];

  grid.innerHTML = cards.map(team => `
    <button type="button" class="subcasilla ${selectedTeamM1 === team.equipo ? "selected" : ""}" data-team="${escapeHtml(team.equipo)}">
      <span class="subcasilla-title">${escapeHtml(team.label)}</span>
      <span class="subcasilla-meta">${escapeHtml(team.detail)}</span>
    </button>
  `).join("");

  const filtered = currentData.advisors.filter(advisor => {
    const matchesTeam = selectedTeamM1 === "ALL" || advisor.equipo === selectedTeamM1;
    const matchesQuery = !query || normalizeText(`${advisor.gestor} ${advisor.equipo}`).includes(query);
    return matchesTeam && matchesQuery;
  });

  document.getElementById("m1-tbody").innerHTML = filtered.length ? filtered.map(advisor => {
    const rate = advisor.cartera > 0 ? advisor.recibido / advisor.cartera * 100 : null;
    return `
      <tr>
        <td><div class="user-cell"><div class="avatar-circle">${escapeHtml(getInitials(advisor.gestor))}</div><strong>${escapeHtml(advisor.gestor)}</strong></div></td>
        <td>${escapeHtml(advisor.equipo || "—")}</td>
        <td><span class="casilla-tag tag-blue">${escapeHtml(advisor.empresa || "—")}</span></td>
        <td>${escapeHtml(formatNumber(advisor.cuentas))}</td>
        <td>${escapeHtml(formatMoney(advisor.cartera))}</td>
        <td><strong>${escapeHtml(formatMoney(advisor.recibido))}</strong></td>
        <td><strong>${escapeHtml(formatRate(rate))}</strong><div class="progress-track"><div class="progress-fill" style="width:${rate === null ? 0 : Math.min(Math.max(rate, 0), 100)}%"></div></div></td>
      </tr>`;
  }).join("") : showEmptyRow(7, "No hay registros para los filtros seleccionados.");
}

function selectM2App(app) {
  selectedAppM2 = app;
  renderM2();
}

function segmentTarget(segment) {
  const normalized = normalizeText(segment);
  if (normalized === "s1") return 0.12;
  if (normalized === "m2") return 0.012;
  if (normalized === "s2") return 0.018;
  return null;
}

function renderM2() {
  if (!currentData) return;
  const query = normalizeText(document.getElementById("m2-search").value);
  const subtract = document.getElementById("toggle-domingos").checked;
  const apps = [...new Set(currentData.commissions.map(row => row.app).filter(Boolean))].sort();
  const appGrid = document.getElementById("m2-app-grid");
  const selectedAppExists = selectedAppM2 === "ALL" || apps.includes(selectedAppM2);
  if (!selectedAppExists) selectedAppM2 = "ALL";
  const appCards = [{ id: "ALL", label: "🌐 Todas las Apps", detail: `${currentData.commissions.length} ejecutivos` },
    ...apps.map(app => ({ id: app, label: app, detail: `${currentData.commissions.filter(row => row.app === app).length} ejecutivos` }))];

  appGrid.innerHTML = appCards.map(app => `
    <button type="button" class="subcasilla ${selectedAppM2 === app.id ? "selected" : ""}" data-app="${escapeHtml(app.id)}">
      <span class="subcasilla-title">${escapeHtml(app.label)}</span>
      <span class="subcasilla-meta">${escapeHtml(app.detail)}</span>
    </button>
  `).join("");

  const rows = currentData.commissions.filter(record => {
    const matchesApp = selectedAppM2 === "ALL" || record.app === selectedAppM2;
    const matchesQuery = !query || normalizeText(`${record.nombre} ${record.tlj}`).includes(query);
    return matchesApp && matchesQuery;
  });

  document.getElementById("m2-tbody").innerHTML = rows.length ? rows.map(record => {
    const baseFinal = subtract
      ? record.base === null || record.d_base === null ? null : Math.max(0, record.base - record.d_base)
      : record.base;
    const receivedFinal = subtract
      ? record.rec === null || record.d_rec === null ? null : Math.max(0, record.rec - record.d_rec)
      : record.rec;
    const returnRate = baseFinal > 0 && receivedFinal !== null ? receivedFinal / baseFinal : null;
    const target = segmentTarget(record.seg);
    const completion = returnRate !== null && target ? returnRate / target : null;
    const baseError = record.baseError.startsWith("#") ? record.baseError : record.d_baseError;
    const receivedError = record.recError.startsWith("#") ? record.recError : record.d_recError;
    const baseDisplay = baseFinal === null ? baseError || "—" : formatMoney(baseFinal);
    const receivedDisplay = receivedFinal === null ? receivedError || "—" : formatMoney(receivedFinal);
    let grade = "—", gradeClass = "tag-blue";
    if (completion !== null) {
      grade = "F";
      gradeClass = "tag-rose";
      if (completion >= 0.95) { grade = "S"; gradeClass = "tag-mint"; }
      else if (completion >= 0.90) { grade = "A"; gradeClass = "tag-mint"; }
      else if (completion >= 0.80) { grade = "B"; gradeClass = "tag-blue"; }
      else if (completion >= 0.70) { grade = "C"; gradeClass = "tag-amber"; }
      else if (completion >= 0.60) { grade = "D"; gradeClass = "tag-amber"; }
    }
    return `
      <tr>
        <td><div class="user-cell"><div class="avatar-circle">${escapeHtml(getInitials(record.nombre))}</div><strong>${escapeHtml(record.nombre)}</strong></div></td>
        <td><span class="casilla-tag tag-blue">${escapeHtml(record.app || "—")}</span></td>
        <td><strong>${escapeHtml(record.seg || "—")}</strong></td>
        <td>${escapeHtml(record.tlj || "—")}</td>
        <td>${escapeHtml(baseDisplay)}</td>
        <td><strong>${escapeHtml(receivedDisplay)}</strong></td>
        <td>${escapeHtml(formatRate(returnRate === null ? null : returnRate * 100))}</td>
        <td><span class="casilla-tag ${gradeClass}">${escapeHtml(grade)}</span></td>
      </tr>`;
  }).join("") : showEmptyRow(8, "No hay registros para los filtros seleccionados.");
}

function fillSundayMetrics(report, prefix, label) {
  const date = report.date ? ` (${report.date})` : "";
  setText(`${prefix}-label`, `${label}${date}`);
  setText(`${prefix}-value`, formatMoney(report.amount));
  setText(`${prefix}-sub`, `Cartera: ${formatMoney(report.portfolio)} · ${formatNumber(report.people)} gestores`);
}

function updateCommissionMetrics() {
  const reports = currentData.sundayReports;
  fillSundayMetrics(reports.sunday1, "m2-sunday1", "Domingo 1");
  fillSundayMetrics(reports.sunday2, "m2-sunday2", "Domingo 2");
  fillSundayMetrics(reports.holiday, "m2-holiday", "Festivo 1");
  setText("m2-executive-count", `${formatNumber(new Set(currentData.commissions.map(row => row.nombre)).size)} Ejecutivos`);
}

function selectM3Tier(tier) {
  selectedTierM3 = tier;
  renderM3();
}

function renderM3() {
  if (!currentData) return;
  const query = normalizeText(document.getElementById("m3-search").value);
  const ranked = [...currentData.advisors].sort((a, b) => (b.recibido || 0) - (a.recibido || 0));
  const tiers = [
    { id: "ALL", title: "🏆 Todos los Escalafones", bono: "Vista general" },
    { id: "A1", title: "🥇 Rango A1 (Top 1-2)", bono: "Bono: $1,650,000" },
    { id: "A2", title: "🥈 Rango A2 (Top 3-5)", bono: "Bono: $1,290,000" },
    { id: "A3", title: "🥉 Rango A3 (Top 6-10)", bono: "Bono: $1,020,000" },
    { id: "B1", title: "🎖️ Rango B1 (Top 11+)", bono: "Bono: $520,000" }
  ];
  document.getElementById("m3-tier-grid").innerHTML = tiers.map(tier => `
    <button type="button" class="subcasilla ${selectedTierM3 === tier.id ? "selected" : ""}" data-tier="${tier.id}">
      <span class="subcasilla-title">${escapeHtml(tier.title)}</span><span class="subcasilla-meta">${escapeHtml(tier.bono)}</span>
    </button>
  `).join("");

  const visible = ranked.map((advisor, index) => {
    const rank = index + 1;
    const tier = rank <= 2 ? "A1" : rank <= 5 ? "A2" : rank <= 10 ? "A3" : "B1";
    const bonus = { A1: "$1,650,000", A2: "$1,290,000", A3: "$1,020,000", B1: "$520,000" }[tier];
    return { advisor, rank, tier, bonus, previous: ranked[index - 1] };
  }).filter(item => {
    const matchesTier = selectedTierM3 === "ALL" || item.tier === selectedTierM3;
    const matchesQuery = !query || normalizeText(item.advisor.gestor).includes(query);
    return matchesTier && matchesQuery;
  });

  document.getElementById("m3-tbody").innerHTML = visible.length ? visible.map(({ advisor, rank, tier, bonus, previous }) => {
    const rate = advisor.cartera > 0 ? advisor.recibido / advisor.cartera * 100 : null;
    const nextAmount = previous ? previous.recibido : (advisor.recibido || 0) * 1.1;
    const amountToNext = Math.max(0, nextAmount - (advisor.recibido || 0));
    const rankClass = rank === 1 ? "rank-1" : rank === 2 ? "rank-2" : rank === 3 ? "rank-3" : "";
    return `
      <tr>
        <td><span class="rank-badge ${rankClass}">#${rank}</span></td>
        <td><div class="user-cell"><div class="avatar-circle">${escapeHtml(getInitials(advisor.gestor))}</div><strong>${escapeHtml(advisor.gestor)}</strong></div></td>
        <td>${escapeHtml(advisor.equipo || "—")}</td>
        <td><strong>${escapeHtml(formatMoney(advisor.recibido))}</strong></td>
        <td>${escapeHtml(formatRate(rate))}</td>
        <td>${rank === 1 ? "🌟 Líder #1 (+10% meta)" : `Faltan ${escapeHtml(formatMoney(amountToNext))}`}</td>
        <td><span class="casilla-tag tag-amber">${tier} · ${bonus}</span></td>
      </tr>`;
  }).join("") : showEmptyRow(7, "No hay registros para los filtros seleccionados.");
}

function selectM4Team(team) {
  selectedTeamM4 = team;
  renderM4();
}

function updateDailyMetrics() {
  const daily = currentData.daily;
  const people = new Set(daily.map(record => record.ej));
  const accounts = sumNumbers(daily.map(record => record.cuentas)) || 0;
  const amount = sumNumbers(daily.map(record => record.monto)) || 0;
  const received = sumNumbers(daily.map(record => record.pagado)) || 0;
  const rate = amount > 0 ? received / amount * 100 : null;
  setText("m4-executive-count", formatNumber(people.size));
  setText("m4-account-count", formatNumber(accounts));
  setText("m4-assigned-total", formatMoney(amount));
  setText("m4-collected-total", formatMoney(received));
  setText("m4-average-rate", `Avance promedio: ${formatRate(rate)}`);
}

function renderM4() {
  if (!currentData) return;
  const teams = [...new Set(currentData.daily.map(record => record.team).filter(Boolean))].sort();
  const grid = document.getElementById("m4-team-grid");
  const cards = ["ALL", ...teams];
  if (!cards.includes(selectedTeamM4)) selectedTeamM4 = "ALL";
  grid.innerHTML = cards.map(team => `
    <button type="button" class="subcasilla ${selectedTeamM4 === team ? "selected" : ""}" data-team="${escapeHtml(team)}">
      <span class="subcasilla-title">${team === "ALL" ? "🌐 Todos los equipos" : `⚡ ${escapeHtml(team)}`}</span>
      <span class="subcasilla-meta">${team === "ALL" ? `${currentData.daily.length} registros` : `${currentData.daily.filter(record => record.team === team).length} registros`}</span>
    </button>
  `).join("");

  const rows = currentData.daily.filter(record => selectedTeamM4 === "ALL" || record.team === selectedTeamM4);
  document.getElementById("m4-tbody").innerHTML = rows.length ? rows.map(record => {
    const rate = record.monto > 0 ? record.pagado / record.monto * 100 : null;
    return `
      <tr>
        <td><div class="user-cell"><div class="avatar-circle">${escapeHtml(getInitials(record.ej))}</div><strong>${escapeHtml(record.ej)}</strong></div></td>
        <td><span class="casilla-tag tag-violet">${escapeHtml(record.team)}</span></td>
        <td>${escapeHtml(formatNumber(record.cuentas))}</td>
        <td>${escapeHtml(formatMoney(record.monto))}</td>
        <td>${escapeHtml(formatMoney(record.meta))}</td>
        <td><strong>${escapeHtml(formatMoney(record.pagado))}</strong></td>
        <td><strong>${escapeHtml(formatRate(rate))}</strong><div class="progress-track"><div class="progress-fill" style="width:${rate === null ? 0 : Math.min(Math.max(rate, 0), 100)}%"></div></div></td>
      </tr>`;
  }).join("") : showEmptyRow(7, "La hoja INFORME H A H no contiene registros para esta selección.");
}

function isYes(value) {
  return ["si", "sí", "yes", "true", "1"].includes(normalizeText(value));
}

function updateStaffMetrics() {
  const people = currentData.staff;
  const telework = people.filter(person => isYes(person.puede_tt)).length;
  const remote = people.filter(person => isYes(person.remoto)).length;
  const probation = people.filter(person => isYes(person.prueba)).length;
  const motorcycle = people.filter(person => isYes(person.moto)).length;
  const disciplinary = people.filter(person => person.disc && normalizeText(person.disc) !== "ninguno").length;
  setText("m5-telework-count", `${formatNumber(telework)} / ${formatNumber(people.length)}`);
  setText("m5-remote-count", `${formatNumber(remote)} actualmente en remoto`);
  setText("m5-probation-count", formatNumber(probation));
  setText("m5-contracted-count", `${formatNumber(Math.max(0, people.length - probation))} fuera de periodo de prueba`);
  setText("m5-motorcycle-count", formatNumber(motorcycle));
  setText("m5-disciplinary-count", formatNumber(disciplinary));
  setText("m5-staff-count", `${formatNumber(people.length)} fichas del Excel`);
  setText("m5-seat-count", `${formatNumber(currentData.seats.length)} puestos de ${formatNumber(new Set(currentData.seats.map(seat => seat.zona)).size)} zonas`);
}

function renderM5List() {
  if (!currentData) return;
  const query = normalizeText(document.getElementById("m5-search").value);
  const rows = currentData.staff.filter(person =>
    !query || normalizeText(`${person.nombre} ${person.disc}`).includes(query)
  );
  document.getElementById("m5-tbody").innerHTML = rows.length ? rows.map(person => `
    <tr>
      <td><div class="user-cell"><div class="avatar-circle">${escapeHtml(getInitials(person.nombre))}</div><strong>${escapeHtml(person.nombre)}</strong></div></td>
      <td><span class="casilla-tag ${isYes(person.puede_tt) ? "tag-mint" : "tag-blue"}">${escapeHtml(person.puede_tt || "—")}</span></td>
      <td>${isYes(person.moto) ? "🏍️ Sí" : escapeHtml(person.moto || "—")}</td>
      <td>${escapeHtml(person.sede2 || "—")}</td>
      <td>${escapeHtml(person.prueba || "—")}</td>
      <td>${escapeHtml(person.ingreso || "—")}</td>
      <td><span class="casilla-tag ${!person.disc || normalizeText(person.disc) === "ninguno" ? "tag-mint" : "tag-rose"}">${escapeHtml(person.disc || "—")}</span></td>
    </tr>
  `).join("") : showEmptyRow(7, "No hay registros de personal para esta búsqueda.");
}

function renderM5Seats() {
  if (!currentData) return;
  const select = document.getElementById("m5-zone-filter");
  const selectedZone = select.value;
  const zones = [...new Set(currentData.seats.map(seat => seat.zona))].sort();
  select.replaceChildren(new Option("Todas las Salas y Sedes", "ALL"));
  zones.forEach(zone => select.add(new Option(zone, zone)));
  select.value = zones.includes(selectedZone) ? selectedZone : "ALL";

  const filtered = currentData.seats.filter(seat => select.value === "ALL" || seat.zona === select.value);
  document.getElementById("m5-seats-grid").innerHTML = filtered.length ? filtered.map(seat => `
    <div class="seat-box">
      <div class="seat-icon">🪑</div>
      <div class="seat-name">${escapeHtml(seat.nombre)}</div>
      <div class="seat-zone">${escapeHtml(seat.zona)}</div>
    </div>
  `).join("") : `<p>No hay puestos registrados en la hoja PUESTOS para esta zona.</p>`;
}

function switchRRHHSub(mode) {
  document.getElementById("sub-rrhh-list").classList.toggle("selected", mode === "list");
  document.getElementById("sub-rrhh-map").classList.toggle("selected", mode === "map");
  document.getElementById("rrhh-panel-list").style.display = mode === "list" ? "block" : "none";
  document.getElementById("rrhh-panel-map").style.display = mode === "map" ? "block" : "none";
  if (mode === "map") renderM5Seats();
}

function resetFilters() {
  ["m1-search", "m2-search", "m3-search", "m5-search"].forEach(id => {
    document.getElementById(id).value = "";
  });
  document.getElementById("toggle-domingos").checked = true;
  selectedTeamM1 = "ALL";
  selectedAppM2 = "ALL";
  selectedTierM3 = "ALL";
  selectedTeamM4 = "ALL";
  document.getElementById("m5-zone-filter").value = "ALL";
  switchRRHHSub("list");
}

function renderWorkbook() {
  updateGlobalMetrics();
  renderM1();
  updateCommissionMetrics();
  renderM2();
  renderM3();
  updateDailyMetrics();
  renderM4();
  updateStaffMetrics();
  renderM5List();
  renderM5Seats();
}

function bindCardSelection(containerId, dataAttribute, callback) {
  document.getElementById(containerId).addEventListener("click", event => {
    const card = event.target.closest(`[data-${dataAttribute}]`);
    if (card && event.currentTarget.contains(card)) callback(card.dataset[dataAttribute]);
  });
}

bindCardSelection("m1-teams-grid", "team", selectM1Team);
bindCardSelection("m2-app-grid", "app", selectM2App);
bindCardSelection("m3-tier-grid", "tier", selectM3Tier);
bindCardSelection("m4-team-grid", "team", selectM4Team);

fileInput.addEventListener("change", async event => {
  const file = event.target.files[0];
  if (!file) return;

  currentData = null;
  appContent.hidden = true;
  uploadGate.hidden = false;
  uploadMessage.textContent = "";
  setText("status-subtitle", "Cargando archivo…");

  try {
    if (typeof XLSX === "undefined") {
      throw new Error("No se pudo cargar la librería para leer Excel. Revisa tu conexión e inténtalo de nuevo.");
    }
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
    const parsedData = parseWorkbook(workbook);
    currentData = parsedData;
    resetFilters();
    renderWorkbook();
    setText("status-subtitle", `Archivo cargado: ${file.name} · ${parsedData.sheetCount} hojas`);
    uploadGate.hidden = true;
    appContent.hidden = false;
  } catch (error) {
    console.error("No se pudo cargar el archivo Excel:", error);
    const message = error instanceof Error ? error.message : "Ocurrió un error inesperado al leer el archivo.";
    uploadMessage.textContent = `No se pudo cargar "${file.name}": ${message}`;
    setText("status-subtitle", "No hay un Excel válido cargado");
  } finally {
    fileInput.value = "";
  }
});

document.getElementById("m1-search").addEventListener("input", renderM1);
document.getElementById("m2-search").addEventListener("input", renderM2);
document.getElementById("m3-search").addEventListener("input", renderM3);
document.getElementById("m5-search").addEventListener("input", renderM5List);
document.getElementById("toggle-domingos").addEventListener("change", renderM2);
document.getElementById("m5-zone-filter").addEventListener("change", renderM5Seats);
document.getElementById("sub-rrhh-list").addEventListener("click", () => switchRRHHSub("list"));
document.getElementById("sub-rrhh-map").addEventListener("click", () => switchRRHHSub("map"));
