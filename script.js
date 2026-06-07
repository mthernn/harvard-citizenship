/* Harvard Citizenship Tutor Ranker
   Static browser-only app for Google Forms Excel exports.
   Requires SheetJS loaded from index.html.
*/

const CONFIG = {
  outputFileName: "ranked_citizenship_tutors.xlsx",
  previewLimit: 50,
  languagePriority: [
    { key: "Spanish", aliases: ["spanish"] },
    { key: "Mandarin", aliases: ["mandarin", "chinese mandarin"] },
    { key: "Haitian Creole", aliases: ["haitian creole", "haitian"] },
    { key: "French", aliases: ["french"] },
    { key: "Cantonese", aliases: ["cantonese"] },
    { key: "Vietnamese", aliases: ["vietnamese"] },
    { key: "Arabic", aliases: ["arabic"] },
    { key: "Russian", aliases: ["russian"] },
    { key: "Italian", aliases: ["italian"] },
    { key: "Hindi", aliases: ["hindi"] },
    { key: "Cape Verdean Creole", aliases: ["cape verdean creole", "cape verdean", "kriolu", "crioulo"] },
    { key: "Korean", aliases: ["korean"] },
    { key: "German", aliases: ["german"] },
    { key: "Telugu", aliases: ["telugu"] },
    { key: "Amharic", aliases: ["amharic"] },
    { key: "Tagalog", aliases: ["tagalog", "filipino"] },
    { key: "Japanese", aliases: ["japanese"] },
    { key: "Urdu", aliases: ["urdu"] },
    { key: "Albanian", aliases: ["albanian"] },
    { key: "Persian", aliases: ["persian", "farsi"] },
    { key: "Thai", aliases: ["thai"] },
    { key: "Punjabi", aliases: ["punjabi", "panjabi"] },
    { key: "Yoruba", aliases: ["yoruba"] },
    { key: "Swahili", aliases: ["swahili", "kiswahili"] },
    { key: "Mongolian", aliases: ["mongolian"] },
    { key: "Latin", aliases: ["latin"] }
  ],
  fields: {
    forename: {
      label: "Forename",
      required: true,
      exact: ["Forename"],
      includesAll: [["forename"], ["first", "name"]]
    },
    surname: {
      label: "Surname",
      required: true,
      exact: ["Surname"],
      includesAll: [["surname"], ["last", "name"], ["family", "name"]]
    },
    email: {
      label: "Email",
      required: false,
      exact: ["Harvard Email", "Email Address", "Email"],
      includesAll: [["email"]]
    },
    phone: {
      label: "Phone Number",
      required: false,
      exact: ["Phone Number", "Phone Number (555/555-5555)"],
      includesAll: [["phone"]]
    },
    birthdate: {
      label: "Birthdate",
      required: false,
      exact: ["Birthdate"],
      includesAll: [["birthdate"], ["birthday"], ["date", "birth"]]
    },
    classYear: {
      label: "Class Year",
      required: true,
      exact: ["Class Year"],
      includesAll: [["class", "year"]]
    },
    languages: {
      label: "Language(s) Other than English",
      required: false,
      exact: ["Language(s) Other than English (if any)"],
      includesAll: [["language", "english"], ["languages", "below"], ["do", "you", "know", "languages"]]
    },
    currentOrFormerBoard: {
      label: "Current or Former Board Member",
      required: false,
      exact: ["Are you a current or former Board Member?"],
      includesAll: [["current", "former", "board"]]
    },
    currentBoard: {
      label: "Current Board Member",
      required: false,
      exact: ["Current Board Member?", "Are you a CURRENT Board Member?"],
      includesAll: [["current", "board", "member"]]
    },
    currentPosition: {
      label: "Current Board Position",
      required: false,
      exact: ["If current board member, position:"],
      includesAll: [["current", "board", "position"]]
    },
    formerBoard: {
      label: "Former Board Member",
      required: false,
      exact: ["Former Board Member?", "Are you a FORMER Board Member?"],
      includesAll: [["former", "board", "member"]]
    },
    formerHighestPosition: {
      label: "Former Highest Position",
      required: false,
      exact: ["If former board member, former highest position:", "If you are a Former Board Member, what was your highest position?"],
      includesAll: [["former", "highest", "position"], ["former", "board", "position"]]
    },
    yearsAgoSelected: {
      label: "Years Ago Selected",
      required: false,
      exact: ["How many years ago were you selected for the position selected above?", "If applicable, how many years ago were you selected for the position selected above?"],
      includesAll: [["years", "ago", "selected"]]
    },
    returningMember: {
      label: "Returning Member",
      required: false,
      exact: ["Returning Member?", "Are you a new or returning member of Citizenship Tutoring?"],
      includesAll: [["returning", "member"], ["new", "returning", "citizenship"]]
    },
    previousTutee: {
      label: "Previous Tutee Name",
      required: false,
      exact: ["If your tutee has not naturalized yet, what is their name? (so we can match them again)"],
      includesAll: [["tutee", "naturalized"], ["assigned", "tutee", "previous", "semester"]]
    },
    concentration: {
      label: "Concentration",
      required: false,
      exact: ["Concentration"],
      includesAll: [["concentration"]]
    },
    dorm: {
      label: "Dorm",
      required: false,
      exact: ["Dorm"],
      includesAll: [["dorm"]]
    },
    hometown: {
      label: "Hometown",
      required: false,
      exact: ["Hometown"],
      includesAll: [["hometown"]]
    },
    foodRestrictions: {
      label: "Food Restrictions",
      required: false,
      exact: ["Food Restrictions"],
      includesAll: [["food", "restriction"], ["dietary", "restriction"]]
    }
  }
};

let workbook = null;
let rawRows = [];
let headers = [];
let rankedRows = [];
let activeSheetName = "";

const els = {
  fileInput: document.getElementById("fileInput"),
  fileName: document.getElementById("fileName"),
  rankButton: document.getElementById("rankButton"),
  downloadButton: document.getElementById("downloadButton"),
  resetButton: document.getElementById("resetButton"),
  status: document.getElementById("status"),
  summaryCard: document.getElementById("summaryCard"),
  summaryGrid: document.getElementById("summaryGrid"),
  diagnosticsCard: document.getElementById("diagnosticsCard"),
  diagnostics: document.getElementById("diagnostics"),
  previewCard: document.getElementById("previewCard"),
  previewTable: document.getElementById("previewTable")
};

els.fileInput.addEventListener("change", handleFileUpload);
els.rankButton.addEventListener("click", rankTutors);
els.downloadButton.addEventListener("click", downloadRankedWorkbook);
els.resetButton.addEventListener("click", resetApp);

function handleFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  resetResultsOnly();
  els.fileName.textContent = file.name;
  setStatus("Reading spreadsheet...", "warn");

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: "array", cellDates: false, raw: true });
      activeSheetName = chooseSheetName(workbook);
      const sheet = workbook.Sheets[activeSheetName];
      const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });

      if (!matrix.length) throw new Error("The selected sheet is empty.");
      headers = matrix[0].map(value => String(value || "").trim());
      rawRows = matrix.slice(1)
        .filter(row => row.some(cell => String(cell ?? "").trim() !== ""))
        .map(row => rowToObject(headers, row));

      if (!rawRows.length) throw new Error("No response rows were found after the header row.");

      els.rankButton.disabled = false;
      els.resetButton.disabled = false;
      setStatus(`Loaded ${rawRows.length} response row${rawRows.length === 1 ? "" : "s"} from “${activeSheetName}”.`, "ok");
      showDiagnostics(buildFieldMap(headers));
    } catch (error) {
      console.error(error);
      setStatus(error.message || "Unable to read this spreadsheet.", "error");
      els.rankButton.disabled = true;
      els.downloadButton.disabled = true;
    }
  };
  reader.onerror = () => setStatus("Unable to read this file.", "error");
  reader.readAsArrayBuffer(file);
}

function chooseSheetName(wb) {
  const preferred = wb.SheetNames.find(name => normalizeText(name).includes("form responses"));
  return preferred || wb.SheetNames[0];
}

function rowToObject(headerRow, row) {
  const obj = {};
  headerRow.forEach((header, index) => {
    if (!header) return;
    const key = obj.hasOwnProperty(header) ? `${header}__${index + 1}` : header;
    obj[key] = row[index] ?? "";
  });
  return obj;
}

function rankTutors() {
  try {
    const fieldMap = buildFieldMap(headers);
    const enriched = rawRows.map((row, index) => enrichRow(row, fieldMap, index));
    rankedRows = enriched.sort(compareTutors).map((row, index) => ({ ...row, seniorityRank: index + 1 }));

    renderSummary(rankedRows);
    renderPreview(rankedRows);
    showDiagnostics(fieldMap);
    els.downloadButton.disabled = false;
    setStatus(`Ranked ${rankedRows.length} tutor${rankedRows.length === 1 ? "" : "s"}.`, "ok");
  } catch (error) {
    console.error(error);
    setStatus(error.message || "Unable to rank tutors.", "error");
  }
}

function buildFieldMap(headerRow) {
  const allHeaders = [...headerRow];
  const map = {};
  Object.entries(CONFIG.fields).forEach(([fieldKey, spec]) => {
    map[fieldKey] = findMatchingHeaders(allHeaders, spec);
  });
  return map;
}

function findMatchingHeaders(headerRow, spec) {
  const exactNorms = (spec.exact || []).map(normalizeHeader);
  const matches = [];

  headerRow.forEach(header => {
    const norm = normalizeHeader(header);
    if (!norm) return;
    const exactMatch = exactNorms.includes(norm);
    const includeMatch = (spec.includesAll || []).some(group => group.every(term => norm.includes(normalizeHeader(term))));
    if (exactMatch || includeMatch) matches.push(header);
  });

  return [...new Set(matches)];
}

function enrichRow(row, fieldMap, originalIndex) {
  const currentPosition = getValue(row, fieldMap.currentPosition);
  const formerPosition = getValue(row, fieldMap.formerHighestPosition);
  const currentBoardValue = getValue(row, fieldMap.currentBoard) || getValue(row, fieldMap.currentOrFormerBoard);
  const formerBoardValue = getValue(row, fieldMap.formerBoard) || getValue(row, fieldMap.currentOrFormerBoard);

  const currentBoard = isCurrentBoardMember(currentBoardValue, currentPosition);
  const formerBoard = isFormerBoardMember(formerBoardValue, formerPosition);
  const director = currentBoard && roleIncludes(currentPosition, "director");
  const secretary = currentBoard && roleIncludes(currentPosition, "secretary");
  const formerDirector = formerBoard && roleIncludes(formerPosition, "director");
  const formerSecretary = formerBoard && roleIncludes(formerPosition, "secretary");
  const yearsAgo = toNumber(getValue(row, fieldMap.yearsAgoSelected));
  const languageRaw = getValue(row, fieldMap.languages);
  const detectedLanguages = detectLanguages(languageRaw);
  const priorityLanguageFlags = Object.fromEntries(
    CONFIG.languagePriority.map(lang => [lang.key, hasLanguage(detectedLanguages, lang)])
  );
  const birthdateValue = getValue(row, fieldMap.birthdate);
  const birthdate = parseDateValue(birthdateValue);

  const calculated = {
    seniorityRank: null,
    director,
    secretary,
    boardMember: currentBoard,
    formerDirector,
    formerSecretary,
    formerBoardMember: formerBoard,
    yearsSinceDirector: formerDirector ? yearsAgo : 0,
    yearsSinceSecretary: formerSecretary ? yearsAgo : 0,
    yearsSinceBoardMember: formerBoard ? yearsAgo : 0,
    returningTutor: isReturningMember(getValue(row, fieldMap.returningMember)),
    classYear: toNumber(getValue(row, fieldMap.classYear), Number.POSITIVE_INFINITY),
    numberOfLotes: detectedLanguages.length,
    detectedLanguages,
    priorityLanguageFlags,
    birthdateSortValue: birthdate ? birthdate.getTime() : Number.POSITIVE_INFINITY,
    surnameSortValue: normalizeText(getValue(row, fieldMap.surname)),
    forenameSortValue: normalizeText(getValue(row, fieldMap.forename)),
    originalIndex
  };

  return { original: row, calculated };
}

function compareTutors(a, b) {
  const ac = a.calculated;
  const bc = b.calculated;

  const comparisons = [
    descBool(ac.director, bc.director),
    descBool(ac.secretary, bc.secretary),
    descBool(ac.boardMember, bc.boardMember),
    descBool(ac.formerDirector, bc.formerDirector),
    descBool(ac.formerSecretary, bc.formerSecretary),
    descBool(ac.formerBoardMember, bc.formerBoardMember),
    descNum(ac.yearsSinceDirector, bc.yearsSinceDirector),
    descNum(ac.yearsSinceSecretary, bc.yearsSinceSecretary),
    descNum(ac.yearsSinceBoardMember, bc.yearsSinceBoardMember),
    descBool(ac.returningTutor, bc.returningTutor),
    ascNum(ac.classYear, bc.classYear),
    descNum(ac.numberOfLotes, bc.numberOfLotes)
  ];

  for (const lang of CONFIG.languagePriority) {
    comparisons.push(descBool(ac.priorityLanguageFlags[lang.key], bc.priorityLanguageFlags[lang.key]));
  }

  comparisons.push(
    ascNum(ac.birthdateSortValue, bc.birthdateSortValue),
    ascText(ac.surnameSortValue, bc.surnameSortValue),
    ascText(ac.forenameSortValue, bc.forenameSortValue),
    ascNum(ac.originalIndex, bc.originalIndex)
  );

  return comparisons.find(value => value !== 0) || 0;
}

function descBool(a, b) { return Number(Boolean(b)) - Number(Boolean(a)); }
function descNum(a, b) { return safeNumber(b) - safeNumber(a); }
function ascNum(a, b) { return safeNumber(a) - safeNumber(b); }
function ascText(a, b) { return String(a || "").localeCompare(String(b || "")); }
function safeNumber(value) { return Number.isFinite(value) ? value : Number.POSITIVE_INFINITY; }

function getValue(row, possibleHeaders) {
  if (!possibleHeaders || !possibleHeaders.length) return "";
  for (const header of possibleHeaders) {
    const value = row[header];
    if (value !== undefined && value !== null && String(value).trim() !== "") return value;
  }
  return "";
}

function isCurrentBoardMember(value, position) {
  const combined = normalizeText(`${value} ${position}`);
  return isTruthy(value) || combined.includes("current board") || roleIncludes(position, "director") || roleIncludes(position, "secretary") || roleIncludes(position, "board");
}

function isFormerBoardMember(value, position) {
  const combined = normalizeText(`${value} ${position}`);
  return isTruthy(value) || combined.includes("former board") || roleIncludes(position, "director") || roleIncludes(position, "secretary") || roleIncludes(position, "board");
}

function roleIncludes(value, role) {
  const text = normalizeText(value);
  if (!text) return false;
  if (role === "director") return text.includes("director");
  if (role === "secretary") return text.includes("secretary");
  if (role === "board") return text.includes("board") || text.includes("director") || text.includes("secretary");
  return text.includes(role);
}

function isReturningMember(value) {
  const text = normalizeText(value);
  return isTruthy(value) || text.includes("returning");
}

function isTruthy(value) {
  const text = normalizeText(value);
  return ["true", "t", "yes", "y", "1", "selected"].includes(text) || text.startsWith("yes") || text.includes("i am a returning") || text.includes("current") || text.includes("former");
}

function toNumber(value, fallback = 0) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const match = String(value ?? "").match(/-?\d+(\.\d+)?/);
  return match ? Number(match[0]) : fallback;
}

function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && Number.isFinite(value)) return excelSerialToDate(value);
  const text = String(value ?? "").trim();
  if (!text) return null;
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  const fractionalDay = serial - Math.floor(serial) + 0.0000001;
  const totalSeconds = Math.floor(86400 * fractionalDay);
  const seconds = totalSeconds % 60;
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor(totalSeconds / 60) % 60;
  return new Date(dateInfo.getFullYear(), dateInfo.getMonth(), dateInfo.getDate(), hours, minutes, seconds);
}

function detectLanguages(value) {
  const text = String(value ?? "").trim();
  if (!text) return [];
  const cleaned = text
    .replace(/[\u{1F1E6}-\u{1F1FF}]/gu, " ")
    .replace(/[()\[\]]/g, " ");

  const pieces = cleaned
    .split(/,|;|\n|\r|\|/)
    .map(item => item.trim())
    .filter(Boolean);

  const detected = [];
  for (const lang of CONFIG.languagePriority) {
    if (hasLanguage(pieces, lang) || lang.aliases.some(alias => normalizeText(cleaned).includes(alias))) detected.push(lang.key);
  }

  const extraPieces = pieces
    .map(piece => piece.replace(/\s+/g, " ").trim())
    .filter(piece => piece && !detected.some(lang => normalizeText(piece).includes(normalizeText(lang))));

  return [...new Set([...detected, ...extraPieces])];
}

function hasLanguage(languageList, lang) {
  const normalizedItems = languageList.map(normalizeText).join(" | ");
  return lang.aliases.some(alias => normalizedItems.includes(normalizeText(alias)));
}

function normalizeHeader(value) {
  return normalizeText(value).replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
}

function normalizeText(value) {
  return String(value ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function renderSummary(rows) {
  const currentBoard = rows.filter(r => r.calculated.boardMember).length;
  const formerBoard = rows.filter(r => r.calculated.formerBoardMember).length;
  const returning = rows.filter(r => r.calculated.returningTutor).length;
  const withLanguages = rows.filter(r => r.calculated.numberOfLotes > 0).length;
  const directors = rows.filter(r => r.calculated.director).length;
  const secretaries = rows.filter(r => r.calculated.secretary).length;

  const items = [
    [rows.length, "Total tutors"],
    [currentBoard, "Current board"],
    [formerBoard, "Former board"],
    [directors, "Current directors"],
    [secretaries, "Current secretaries"],
    [returning, "Returning members"],
    [withLanguages, "With LOTEs"]
  ];

  els.summaryGrid.innerHTML = items.map(([value, label]) => `
    <div class="summary-item">
      <div class="summary-value">${escapeHtml(value)}</div>
      <div class="summary-label">${escapeHtml(label)}</div>
    </div>
  `).join("");
  els.summaryCard.classList.remove("hidden");
}

function renderPreview(rows) {
  const displayHeaders = ["Rank", "Forename", "Surname", "Class Year", "Current Role", "Former Role", "Years Ago", "Returning", "LOTEs", "Detected Languages", "Birthdate", "Email"];
  const body = rows.slice(0, CONFIG.previewLimit).map(row => {
    const c = row.calculated;
    const o = row.original;
    const fieldMap = buildFieldMap(headers);
    const values = [
      row.seniorityRank,
      getValue(o, fieldMap.forename),
      getValue(o, fieldMap.surname),
      c.classYear === Number.POSITIVE_INFINITY ? "" : c.classYear,
      getValue(o, fieldMap.currentPosition),
      getValue(o, fieldMap.formerHighestPosition),
      getValue(o, fieldMap.yearsAgoSelected),
      c.returningTutor ? "TRUE" : "FALSE",
      c.numberOfLotes,
      c.detectedLanguages.join(", "),
      formatDateForDisplay(getValue(o, fieldMap.birthdate)),
      getValue(o, fieldMap.email)
    ];
    return `<tr>${values.map(value => `<td>${escapeHtml(value)}</td>`).join("")}</tr>`;
  }).join("");

  els.previewTable.innerHTML = `
    <thead><tr>${displayHeaders.map(header => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>
    <tbody>${body}</tbody>
  `;
  els.previewCard.classList.remove("hidden");
}

function showDiagnostics(fieldMap) {
  const items = [];
  Object.entries(CONFIG.fields).forEach(([key, spec]) => {
    const matches = fieldMap[key] || [];
    if (matches.length) {
      items.push(`<li class="good"><strong>${escapeHtml(spec.label)}:</strong> matched ${matches.map(escapeHtml).join("; ")}</li>`);
    } else if (spec.required) {
      items.push(`<li class="bad"><strong>${escapeHtml(spec.label)}:</strong> required field not found</li>`);
    } else {
      items.push(`<li class="note"><strong>${escapeHtml(spec.label)}:</strong> optional field not found</li>`);
    }
  });
  els.diagnostics.innerHTML = `<ul>${items.join("")}</ul>`;
  els.diagnosticsCard.classList.remove("hidden");
}

function downloadRankedWorkbook() {
  if (!rankedRows.length) return;
  const fieldMap = buildFieldMap(headers);
  const outputRows = rankedRows.map(row => buildOutputRow(row, fieldMap));
  const worksheet = XLSX.utils.json_to_sheet(outputRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, worksheet, "Ranked Tutors");

  const rulesSheet = XLSX.utils.aoa_to_sheet([
    ["Harvard Citizenship Tutor Ranker"],
    ["Ranking rule", "Direction"],
    ["Director?", "TRUE first"],
    ["Secretary?", "TRUE first"],
    ["Current Board Member?", "TRUE first"],
    ["Former Director?", "TRUE first"],
    ["Former Secretary?", "TRUE first"],
    ["Former Board Member?", "TRUE first"],
    ["Years since Director", "largest to smallest"],
    ["Years since Secretary", "largest to smallest"],
    ["Years since Board Member", "largest to smallest"],
    ["Returning Tutor?", "TRUE first"],
    ["Class Year", "smallest to largest"],
    ["Number of LOTEs", "largest to smallest"],
    ...CONFIG.languagePriority.map(lang => [`LOTE includes ${lang.key}?`, "TRUE first"]),
    ["Birthdate", "oldest to youngest"],
    ["Surname", "A to Z"],
    ["Forename", "A to Z"]
  ]);
  XLSX.utils.book_append_sheet(wb, rulesSheet, "Ranking Rules");
  XLSX.writeFile(wb, CONFIG.outputFileName);
}

function buildOutputRow(row, fieldMap) {
  const c = row.calculated;
  const o = row.original;
  const calculatedColumns = {
    "Seniority Rank": row.seniorityRank,
    "Director?": boolText(c.director),
    "Secretary?": boolText(c.secretary),
    "Current Board Member?": boolText(c.boardMember),
    "Former Director?": boolText(c.formerDirector),
    "Former Secretary?": boolText(c.formerSecretary),
    "Former Board Member?": boolText(c.formerBoardMember),
    "Years Since Director": c.yearsSinceDirector,
    "Years Since Secretary": c.yearsSinceSecretary,
    "Years Since Board Member": c.yearsSinceBoardMember,
    "Returning Tutor?": boolText(c.returningTutor),
    "Calculated Class Year": c.classYear === Number.POSITIVE_INFINITY ? "" : c.classYear,
    "Number of LOTEs": c.numberOfLotes,
    "Detected LOTEs": c.detectedLanguages.join(", "),
    ...Object.fromEntries(CONFIG.languagePriority.map(lang => [`LOTE: ${lang.key}?`, boolText(c.priorityLanguageFlags[lang.key])])),
    "Birthdate Parsed": formatDateForDisplay(getValue(o, fieldMap.birthdate))
  };
  return { ...calculatedColumns, ...o };
}

function boolText(value) { return value ? "TRUE" : "FALSE"; }

function formatDateForDisplay(value) {
  const date = parseDateValue(value);
  if (!date) return String(value ?? "");
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function setStatus(message, kind = "") {
  els.status.textContent = message;
  els.status.className = `status ${kind}`.trim();
}

function resetResultsOnly() {
  rankedRows = [];
  els.downloadButton.disabled = true;
  els.summaryCard.classList.add("hidden");
  els.previewCard.classList.add("hidden");
  els.diagnosticsCard.classList.add("hidden");
  els.previewTable.innerHTML = "";
}

function resetApp() {
  workbook = null;
  rawRows = [];
  headers = [];
  rankedRows = [];
  activeSheetName = "";
  els.fileInput.value = "";
  els.fileName.textContent = "No file selected";
  els.rankButton.disabled = true;
  els.downloadButton.disabled = true;
  els.resetButton.disabled = true;
  resetResultsOnly();
  setStatus("Upload a spreadsheet to begin.", "");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
