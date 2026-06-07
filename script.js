/* Harvard Citizenship Tutor Ranker and Matcher
   Static browser-only application. Requires SheetJS loaded from CDN in index.html. */

const LANGUAGE_PRIORITY = [
  "Spanish", "Mandarin", "Haitian Creole", "French", "Cantonese", "Vietnamese",
  "Arabic", "Russian", "Italian", "Hindi", "Cape Verdean Creole", "Korean",
  "German", "Telugu", "Amharic", "Tagalog", "Japanese", "Urdu", "Albanian",
  "Persian", "Thai", "Punjabi", "Yoruba", "Swahili", "Mongolian", "Latin"
];

const LANGUAGE_ALIASES = new Map([
  ["farsi", "Persian"], ["persian", "Persian"],
  ["cape verde creole", "Cape Verdean Creole"], ["cape verdean creole", "Cape Verdean Creole"],
  ["haitian", "Haitian Creole"], ["haitian creole", "Haitian Creole"],
  ["mandarin chinese", "Mandarin"], ["chinese", "Mandarin"],
  ["vietnamese", "Vietnamese"], ["tagalog", "Tagalog"], ["filipino", "Tagalog"]
]);

const DAY_MAP = { mon: "MON", monday: "MON", tue: "TUE", tues: "TUE", tuesday: "TUE", wed: "WED", wednesday: "WED", thu: "THU", thur: "THU", thurs: "THU", thursday: "THU", fri: "FRI", friday: "FRI", sat: "SAT", saturday: "SAT", sun: "SUN", sunday: "SUN" };
const DAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"];

let tutorRowsRaw = [];
let rankedTutors = [];
let tuteeRowsRaw = [];
let preparedTutees = [];
let matchRows = [];
let manualReviewRows = [];
let finalWorkbookData = null;

const $ = id => document.getElementById(id);

$("rankTutorsButton").addEventListener("click", rankTutorsFromUpload);
$("downloadRankedTutorsButton").addEventListener("click", downloadRankedTutors);
$("matchButton").addEventListener("click", matchTuteesFromUpload);
$("downloadMatchesButton").addEventListener("click", downloadMatches);

async function readWorkbookFile(file) {
  if (!file) throw new Error("Please choose a spreadsheet file first.");
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellDates: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("The workbook does not contain a worksheet.");
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
  return { workbook, sheetName, rows };
}

async function rankTutorsFromUpload() {
  try {
    setStatus("tutorStatus", "Reading tutor spreadsheet...", "muted");
    const file = $("tutorFile").files[0];
    const { rows } = await readWorkbookFile(file);
    tutorRowsRaw = rows;
    rankedTutors = rows.map((row, index) => prepareTutor(row, index)).sort(compareTutors);
    rankedTutors.forEach((t, i) => t.rank = i + 1);
    renderTutorResults();
    $("downloadRankedTutorsButton").disabled = rankedTutors.length === 0;
    $("matchButton").disabled = rankedTutors.length === 0;
    setStatus("tutorStatus", `Ranked ${rankedTutors.length} tutors successfully.`, "good");
  } catch (err) {
    console.error(err);
    setStatus("tutorStatus", err.message, "bad");
  }
}

async function matchTuteesFromUpload() {
  try {
    if (!rankedTutors.length) throw new Error("Rank tutors before generating matches.");
    setStatus("matchStatus", "Reading tutee spreadsheet and generating matches...", "muted");
    const file = $("tuteeFile").files[0];
    const { rows } = await readWorkbookFile(file);
    tuteeRowsRaw = rows;
    preparedTutees = rows.map((row, index) => prepareTutee(row, index)).sort(compareTutees);
    preparedTutees.forEach((t, i) => t.priorityRank = i + 1);
    runMatching();
    renderMatchResults();
    $("downloadMatchesButton").disabled = false;
    const complete = matchRows.filter(r => r.Match_Status && r.Match_Status.startsWith("Complete")).length;
    const flagged = manualReviewRows.length + matchRows.filter(r => !String(r.Match_Status || "").startsWith("Complete")).length;
    setStatus("matchStatus", `Generated ${matchRows.length} tutee-session rows. ${complete} complete rows, ${flagged} review notes.`, flagged ? "warn" : "good");
  } catch (err) {
    console.error(err);
    setStatus("matchStatus", err.message, "bad");
  }
}

function prepareTutor(row, index) {
  const forename = getFirst(row, ["forename", "first name", "given name"]);
  const surname = getFirst(row, ["surname", "last name", "family name"]);
  const email = getFirst(row, ["harvard email", "email address", "email"]);
  const currentBoardText = getFirst(row, ["current board member?", "are you a current board member"]);
  const currentOrFormerBoardText = getFirst(row, ["current or former board member"]);
  const currentPosition = getFirst(row, ["if current board member, position", "current position", "select your position"]);
  const formerBoardText = getFirst(row, ["former board member?", "are you a former board member"]);
  const formerHighestPosition = getFirst(row, ["former highest position", "highest position", "former position"]);
  const yearsAgo = toNumber(getFirst(row, ["how many years ago", "years ago"]));
  const returningText = getFirst(row, ["returning member", "new or returning member"]);
  const classYear = toNumber(getFirst(row, ["class year"]), 9999);
  const birthdateRaw = getFirst(row, ["birthdate", "birth date", "date of birth"]);
  const languageText = getFirst(row, ["language(s) other than english", "languages below", "do you know any", "lote", "language"]);
  const tuteeContinuationName = getFirst(row, ["tutee has not naturalized", "assigned a tutee", "match them again", "previous semester"]);
  const modePreference = normalizeMode(getFirst(row, ["prefer", "zoom", "in-person", "in person"]));
  const languages = parseLanguages(languageText);
  const availability = parseAvailabilityFromRow(row, ["availability"]);

  const currentBoard = yesish(currentBoardText) || /current/i.test(String(currentOrFormerBoardText));
  const formerBoard = yesish(formerBoardText) || /former/i.test(String(currentOrFormerBoardText));
  const director = currentBoard && hasPosition(currentPosition, "director");
  const secretary = currentBoard && hasPosition(currentPosition, "secretary");
  const boardMember = currentBoard;
  const formerDirector = formerBoard && hasPosition(formerHighestPosition, "director");
  const formerSecretary = formerBoard && hasPosition(formerHighestPosition, "secretary");
  const returningTutor = yesish(returningText) || /returning/i.test(String(returningText));

  return {
    originalIndex: index,
    rank: 0,
    forename, surname, email,
    fullName: clean(`${forename} ${surname}`) || `Tutor ${index + 1}`,
    currentPosition,
    formerHighestPosition,
    yearsAgo: yearsAgo || 0,
    director, secretary, boardMember, formerDirector, formerSecretary, formerBoard,
    yearsSinceDirector: formerDirector ? yearsAgo || 0 : 0,
    yearsSinceSecretary: formerSecretary ? yearsAgo || 0 : 0,
    yearsSinceBoard: formerBoard ? yearsAgo || 0 : 0,
    returningTutor,
    classYear,
    birthdate: parseDateValue(birthdateRaw),
    languages,
    languageText,
    languageCount: languages.length,
    tuteeContinuationName: clean(tuteeContinuationName),
    availability,
    modePreference,
    assignmentCount: 0,
    assignedTuteeKeys: new Set(),
    raw: row,
    sortKey: []
  };
}

function compareTutors(a, b) {
  const checks = [
    boolDesc(a.director, b.director), boolDesc(a.secretary, b.secretary), boolDesc(a.boardMember, b.boardMember),
    boolDesc(a.formerDirector, b.formerDirector), boolDesc(a.formerSecretary, b.formerSecretary), boolDesc(a.formerBoard, b.formerBoard),
    numDesc(a.yearsSinceDirector, b.yearsSinceDirector), numDesc(a.yearsSinceSecretary, b.yearsSinceSecretary), numDesc(a.yearsSinceBoard, b.yearsSinceBoard),
    boolDesc(a.returningTutor, b.returningTutor), numAsc(a.classYear, b.classYear), numDesc(a.languageCount, b.languageCount),
    ...LANGUAGE_PRIORITY.map(lang => boolDesc(hasLang(a, lang), hasLang(b, lang))),
    dateAsc(a.birthdate, b.birthdate), strAsc(a.surname, b.surname), strAsc(a.forename, b.forename), numAsc(a.originalIndex, b.originalIndex)
  ];
  return checks.find(x => x !== 0) || 0;
}

function prepareTutee(row, index) {
  const forename = getFirst(row, ["forename", "first name", "given name"]);
  const surname = getFirst(row, ["surname", "surename", "last name", "family name"]);
  const email = getFirst(row, ["email address", "email"]);
  const phone = getFirst(row, ["phone"]);
  const communityPartner = getFirst(row, ["associated community partner", "community partner"]);
  const statusText = getFirst(row, ["new or returning", "is the tutee"]);
  const languageText = getFirst(row, ["native language"]);
  const weeklyMeetings = Math.max(1, Math.round(toNumber(getFirst(row, ["how many times per week", "meetings per week", "tutored"]), 1)));
  const stage = getFirst(row, ["naturalization process", "n-400", "interview"]);
  const modePreference = normalizeMode(getFirst(row, ["prefer", "in-person", "in person", "zoom"]));
  const availability = parseAvailabilityFromRow(row, ["tutee's availability", "tutees availability", "availability"]);
  const nativeLanguages = parseLanguages(languageText);
  const returning = /return/i.test(String(statusText));
  const newTutee = !returning;
  const baseTutors = baseTutorsForStage(stage);
  const tutorsPerSession = returning ? Math.max(1, Math.ceil(baseTutors / 2)) : baseTutors;
  const priorityScore = (newTutee ? 10000 : 0) + stageNeedWeight(stage) * 100 + weeklyMeetings * 10 + nativeLanguages.length;
  return {
    originalIndex: index,
    priorityRank: 0,
    forename, surname, email, phone, communityPartner,
    fullName: clean(`${forename} ${surname}`) || `Tutee ${index + 1}`,
    key: nameKey(`${forename} ${surname}`),
    statusText, returning, newTutee,
    nativeLanguages, languageText,
    weeklyMeetings, stage,
    baseTutors, tutorsPerSession,
    totalTutorSlots: weeklyMeetings * tutorsPerSession,
    modePreference, availability,
    priorityScore,
    raw: row
  };
}

function compareTutees(a, b) {
  return numDesc(a.priorityScore, b.priorityScore) || boolDesc(a.newTutee, b.newTutee) || numDesc(stageNeedWeight(a.stage), stageNeedWeight(b.stage)) || numDesc(a.weeklyMeetings, b.weeklyMeetings) || strAsc(a.surname, b.surname) || strAsc(a.forename, b.forename);
}

function runMatching() {
  rankedTutors.forEach(t => { t.assignmentCount = 0; t.assignedTuteeKeys = new Set(); });
  matchRows = [];
  manualReviewRows = [];
  const maxSessions = Math.max(1, toNumber($("maxSessions").value, 2));
  const sessionAssignments = new Map();

  // Continuity pass: a returning tutor who named a returning tutee is assigned first when availability overlaps.
  for (const tutor of rankedTutors) {
    if (!tutor.tuteeContinuationName) continue;
    const tutee = findTuteeByTutorContinuation(tutor, preparedTutees);
    if (!tutee) {
      manualReviewRows.push(reviewRow("Continuity pair not found", tutor.fullName, tutor.tuteeContinuationName, "Tutor listed a returning tutee, but no matching tutee name was found."));
      continue;
    }
    if (!tutee.returning) {
      manualReviewRows.push(reviewRow("Continuity target is not marked returning", tutor.fullName, tutee.fullName, "Tutor listed this tutee, but the tutee row is not marked RETURNING."));
      continue;
    }
    const overlap = intersectSlots(tutor.availability, tutee.availability);
    if (!overlap.length) {
      manualReviewRows.push(reviewRow("Continuity availability conflict", tutor.fullName, tutee.fullName, "Tutor requested this returning tutee, but no availability slot overlaps."));
      continue;
    }
    const key = sessionKey(tutee, 1);
    if (!sessionAssignments.has(key)) sessionAssignments.set(key, []);
    if (!sessionAssignments.get(key).includes(tutor)) {
      sessionAssignments.get(key).push(tutor);
      tutor.assignmentCount += 1;
      tutor.assignedTuteeKeys.add(tutee.key);
    }
  }

  for (const tutee of preparedTutees) {
    for (let sessionNumber = 1; sessionNumber <= tutee.weeklyMeetings; sessionNumber++) {
      const key = sessionKey(tutee, sessionNumber);
      const preAssigned = sessionAssignments.get(key) || [];
      const assigned = [...preAssigned];
      let notes = preAssigned.length ? [`Continuity tutor assigned: ${preAssigned.map(t => t.fullName).join(", ")}.`] : [];
      let statusParts = [];

      if (!tutee.availability.length) statusParts.push("Needs tutee availability");

      const languageAlreadySatisfied = assigned.some(t => tutorMatchesTuteeLanguage(t, tutee));
      if (!languageAlreadySatisfied) {
        const languageCandidate = bestCandidates(tutee, rankedTutors, assigned, maxSessions, true)[0];
        if (languageCandidate) {
          assigned.push(languageCandidate);
          languageCandidate.assignmentCount += 1;
          languageCandidate.assignedTuteeKeys.add(tutee.key);
        } else {
          statusParts.push("Needs language-compatible tutor");
        }
      }

      while (assigned.length < tutee.tutorsPerSession) {
        const candidate = bestCandidates(tutee, rankedTutors, assigned, maxSessions, false)[0];
        if (!candidate) break;
        assigned.push(candidate);
        candidate.assignmentCount += 1;
        candidate.assignedTuteeKeys.add(tutee.key);
      }

      if (!assigned.length && !statusParts.length) statusParts.push("Needs availability-compatible tutor");
      if (assigned.length < tutee.tutorsPerSession) statusParts.push(`Underfilled: ${assigned.length}/${tutee.tutorsPerSession} tutors assigned`);
      if (!assigned.some(t => tutorMatchesTuteeLanguage(t, tutee))) statusParts.push("Native language rule unmet");

      const prefAligned = assigned.filter(t => modeCompatible(tutee.modePreference, t.modePreference)).length;
      const shared = assigned.length ? unionAssignedOverlaps(assigned, tutee.availability) : [];
      const complete = statusParts.length === 0;
      const preferenceOverride = complete && assigned.some(t => !modeCompatible(tutee.modePreference, t.modePreference) && t.modePreference !== "no preference" && tutee.modePreference !== "no preference");
      const matchStatus = complete ? (preferenceOverride ? "Complete, preference overridden" : "Complete") : statusParts.join("; ");
      const row = buildMatchRow(tutee, sessionNumber, assigned, shared, prefAligned, matchStatus, notes.join(" "));
      matchRows.push(row);
      if (!complete) manualReviewRows.push(reviewRow("Session requires review", assigned.map(t => t.fullName).join(", ") || "No tutor assigned", tutee.fullName, matchStatus));
    }
  }

  finalWorkbookData = buildFinalWorkbookData();
}

function bestCandidates(tutee, tutors, assigned, maxSessions, requireLanguage) {
  return tutors
    .filter(t => !assigned.includes(t))
    .filter(t => t.assignmentCount < maxSessions)
    .filter(t => intersectSlots(t.availability, tutee.availability).length > 0)
    .filter(t => !t.assignedTuteeKeys.has(tutee.key) || assigned.length === 0)
    .filter(t => !requireLanguage || tutorMatchesTuteeLanguage(t, tutee))
    .sort((a, b) => scoreCandidate(b, tutee, requireLanguage) - scoreCandidate(a, tutee, requireLanguage));
}

function scoreCandidate(tutor, tutee, requireLanguage) {
  const overlapCount = intersectSlots(tutor.availability, tutee.availability).length;
  const languageScore = tutorMatchesTuteeLanguage(tutor, tutee) ? 500000 : 0;
  const tuteePrefScore = modeCompatible(tutee.modePreference, tutor.modePreference) ? 50000 : 0;
  const noPrefScore = tutor.modePreference === "no preference" ? 5000 : 0;
  const seniorityScore = Math.max(0, 10000 - tutor.rank);
  const workloadPenalty = tutor.assignmentCount * 20000;
  const repeatPenalty = tutor.assignedTuteeKeys.has(tutee.key) ? 15000 : 0;
  const hardLanguageBoost = requireLanguage ? languageScore * 2 : languageScore;
  return hardLanguageBoost + tuteePrefScore + noPrefScore + seniorityScore + overlapCount * 100 - workloadPenalty - repeatPenalty;
}

function buildMatchRow(tutee, sessionNumber, assigned, sharedSlots, prefAligned, status, notes) {
  const out = {
    Tutee_Priority_Rank: tutee.priorityRank,
    Tutee_Name: tutee.fullName,
    Tutee_Email: tutee.email,
    Tutee_Phone: tutee.phone,
    Community_Partner: tutee.communityPartner,
    New_or_Returning: tutee.returning ? "RETURNING" : "NEW",
    Native_Language: tutee.nativeLanguages.join(", ") || tutee.languageText,
    Naturalization_Stage: tutee.stage,
    Weekly_Session_Number: sessionNumber,
    Requested_Sessions_Per_Week: tutee.weeklyMeetings,
    Required_Tutors_This_Session: tutee.tutorsPerSession,
    Shared_Availability_Options: sharedSlots.join(", "),
    Tutee_Mode_Preference: tutee.modePreference,
    Preference_Aligned_Tutors: `${prefAligned}/${assigned.length}`,
    Match_Status: status,
    Notes: notes
  };
  for (let i = 0; i < Math.max(3, assigned.length); i++) {
    const t = assigned[i];
    out[`Tutor_${i + 1}_Rank`] = t ? t.rank : "";
    out[`Tutor_${i + 1}_Name`] = t ? t.fullName : "";
    out[`Tutor_${i + 1}_Email`] = t ? t.email : "";
    out[`Tutor_${i + 1}_Language_Match`] = t ? (tutorMatchesTuteeLanguage(t, tutee) ? "Yes" : "No") : "";
    out[`Tutor_${i + 1}_Mode_Preference`] = t ? t.modePreference : "";
  }
  return out;
}

function buildFinalWorkbookData() {
  return {
    "Generated Matches": matchRows,
    "Manual Review": manualReviewRows,
    "Tutee Priority Rankings": preparedTutees.map(t => ({
      Tutee_Priority_Rank: t.priorityRank,
      Tutee_Name: t.fullName,
      Email: t.email,
      Phone: t.phone,
      New_or_Returning: t.returning ? "RETURNING" : "NEW",
      Native_Language: t.nativeLanguages.join(", ") || t.languageText,
      Naturalization_Stage: t.stage,
      Requested_Sessions_Per_Week: t.weeklyMeetings,
      Tutors_Needed_Per_Session: t.tutorsPerSession,
      Total_Tutor_Slots_Needed: t.totalTutorSlots,
      Mode_Preference: t.modePreference,
      Availability_Slot_Count: t.availability.length,
      Availability: t.availability.join(", ")
    })),
    "Tutor Seniority Rankings": rankedTutorExportRows(),
    "Tutor Workload": rankedTutors.map(t => ({ Rank: t.rank, Tutor_Name: t.fullName, Email: t.email, Assigned_Sessions: t.assignmentCount, Languages: t.languages.join(", "), Availability_Slot_Count: t.availability.length, Mode_Preference: t.modePreference }))
  };
}

function rankedTutorExportRows() {
  return rankedTutors.map(t => ({
    Seniority_Rank: t.rank,
    Forename: t.forename,
    Surname: t.surname,
    Email: t.email,
    Director: yesNo(t.director),
    Secretary: yesNo(t.secretary),
    Current_Board_Member: yesNo(t.boardMember),
    Former_Director: yesNo(t.formerDirector),
    Former_Secretary: yesNo(t.formerSecretary),
    Former_Board_Member: yesNo(t.formerBoard),
    Years_Since_Director: t.yearsSinceDirector,
    Years_Since_Secretary: t.yearsSinceSecretary,
    Years_Since_Board: t.yearsSinceBoard,
    Returning_Tutor: yesNo(t.returningTutor),
    Class_Year: t.classYear === 9999 ? "" : t.classYear,
    Number_of_LOTEs: t.languageCount,
    Detected_LOTEs: t.languages.join(", "),
    Birthdate: t.birthdate ? dateToYMD(t.birthdate) : "",
    Availability_Slot_Count: t.availability.length,
    Availability: t.availability.join(", "),
    Teaching_Mode_Preference: t.modePreference,
    Returning_Tutee_Requested: t.tuteeContinuationName,
    ...t.raw
  }));
}

function renderTutorResults() {
  $("tutorSummary").innerHTML = cards([
    [rankedTutors.length, "Tutors ranked"],
    [rankedTutors.filter(t => t.boardMember).length, "Current board members"],
    [rankedTutors.filter(t => t.returningTutor).length, "Returning tutors"],
    [rankedTutors.filter(t => t.availability.length).length, "Tutors with availability"]
  ]);
  renderTable("tutorTable", rankedTutors.slice(0, 100).map(t => ({ Rank: t.rank, Name: t.fullName, Email: t.email, Class_Year: t.classYear === 9999 ? "" : t.classYear, Languages: t.languages.join(", "), Current_Position: t.currentPosition, Returning_Tutee: t.tuteeContinuationName, Availability_Slots: t.availability.length, Preference: t.modePreference })));
}

function renderMatchResults() {
  const complete = matchRows.filter(r => String(r.Match_Status).startsWith("Complete")).length;
  const underfilled = matchRows.filter(r => String(r.Match_Status).includes("Underfilled")).length;
  $("matchSummary").innerHTML = cards([
    [preparedTutees.length, "Tutees processed"],
    [matchRows.length, "Session rows generated"],
    [complete, "Complete sessions"],
    [manualReviewRows.length + underfilled, "Review flags"]
  ]);
  renderTable("matchTable", matchRows.slice(0, 150));
}

function renderTable(tableId, rows) {
  const table = $(tableId);
  table.innerHTML = "";
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  headers.forEach(h => { const th = document.createElement("th"); th.textContent = h.replaceAll("_", " "); hr.appendChild(th); });
  thead.appendChild(hr);
  const tbody = document.createElement("tbody");
  rows.forEach(row => {
    const tr = document.createElement("tr");
    headers.forEach(h => { const td = document.createElement("td"); td.textContent = row[h] == null ? "" : String(row[h]); tr.appendChild(td); });
    tbody.appendChild(tr);
  });
  table.append(thead, tbody);
}

function cards(items) {
  return items.map(([num, label]) => `<div class="card"><strong>${escapeHtml(String(num))}</strong><span>${escapeHtml(label)}</span></div>`).join("");
}

function downloadRankedTutors() {
  const wb = XLSX.utils.book_new();
  appendSheet(wb, "Tutor Seniority Rankings", rankedTutorExportRows());
  XLSX.writeFile(wb, `ranked_tutors_${todayStamp()}.xlsx`);
}

function downloadMatches() {
  const wb = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(finalWorkbookData)) appendSheet(wb, name, rows);
  XLSX.writeFile(wb, `citizenship_final_matches_${todayStamp()}.xlsx`);
}

function appendSheet(wb, name, rows) {
  const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ Note: "No rows generated." }]);
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
  ws["!cols"] = Array.from({ length: range.e.c + 1 }, (_, c) => {
    let max = 10;
    for (let r = range.s.r; r <= Math.min(range.e.r, 200); r++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v != null) max = Math.max(max, String(cell.v).length);
    }
    return { wch: Math.min(42, Math.max(10, max + 2)) };
  });
  XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
}

function getFirst(row, patterns) {
  const keys = Object.keys(row);
  const matches = [];
  for (const key of keys) {
    const nk = normalize(key);
    if (patterns.some(p => nk.includes(normalize(p)))) matches.push(key);
  }
  // Prefer non-empty values, then shorter headers to avoid grabbing long unrelated prompts too early.
  matches.sort((a, b) => {
    const av = clean(row[a]) ? 0 : 1;
    const bv = clean(row[b]) ? 0 : 1;
    return av - bv || a.length - b.length;
  });
  for (const key of matches) {
    const value = clean(row[key]);
    if (value) return value;
  }
  return "";
}

function parseAvailabilityFromRow(row, headerPatterns) {
  const slots = new Set();
  for (const [header, value] of Object.entries(row)) {
    if (!headerPatterns.some(p => normalize(header).includes(normalize(p)))) continue;
    const dayFromHeader = dayFromText(header);
    const text = clean(value);
    if (!text) continue;
    if (dayFromHeader) {
      splitMulti(text).forEach(token => { if (token) slots.add(`${dayFromHeader}|${normalizeTime(token)}`); });
    } else {
      const pieces = splitMulti(text);
      for (const piece of pieces) {
        const day = dayFromText(piece);
        if (day) {
          const time = piece.replace(new RegExp(Object.keys(DAY_MAP).join("|"), "ig"), "").replace(/[:\-–—]/g, " ");
          if (clean(time)) slots.add(`${day}|${normalizeTime(time)}`);
        }
      }
    }
  }
  return [...slots].sort(slotSort);
}

function parseLanguages(text) {
  const out = new Set();
  for (const raw of splitMulti(text)) {
    const normalized = normalizeLanguage(raw);
    if (normalized) out.add(normalized);
  }
  // Catch languages embedded in long Google Forms option labels.
  const nt = normalize(text);
  for (const lang of LANGUAGE_PRIORITY) {
    if (nt.includes(normalize(lang))) out.add(lang);
  }
  if (nt.includes("farsi")) out.add("Persian");
  if (nt.includes("persian")) out.add("Persian");
  if (nt.includes("cape verde creole")) out.add("Cape Verdean Creole");
  return [...out].sort((a, b) => LANGUAGE_PRIORITY.indexOf(a) - LANGUAGE_PRIORITY.indexOf(b));
}

function normalizeLanguage(raw) {
  const s = clean(String(raw).replace(/[🇦-🇿]/gu, ""));
  if (!s) return "";
  const n = normalize(s);
  if (LANGUAGE_ALIASES.has(n)) return LANGUAGE_ALIASES.get(n);
  const exact = LANGUAGE_PRIORITY.find(l => normalize(l) === n);
  if (exact) return exact;
  const contained = LANGUAGE_PRIORITY.find(l => n.includes(normalize(l)) || normalize(l).includes(n));
  return contained || "";
}

function baseTutorsForStage(stageText) {
  const s = normalize(stageText);
  const n400 = Math.max(1, toNumber($("n400Tutors")?.value, 1));
  const interview = Math.max(1, toNumber($("interviewTutors")?.value, 2));
  const intensive = Math.max(1, toNumber($("intensiveTutors")?.value, 3));
  if (s.includes("interview") && (s.includes("soon") || s.includes("scheduled") || s.includes("intensive") || s.includes("close"))) return intensive;
  if (s.includes("interview") || s.includes("civic") || s.includes("english") || s.includes("test")) return interview;
  if (s.includes("n 400") || s.includes("n-400") || s.includes("n400") || s.includes("fill")) return n400;
  return interview;
}

function stageNeedWeight(stageText) {
  const s = normalize(stageText);
  if (s.includes("interview") && (s.includes("soon") || s.includes("scheduled") || s.includes("intensive") || s.includes("close"))) return 3;
  if (s.includes("interview") || s.includes("civic") || s.includes("english") || s.includes("test")) return 2;
  if (s.includes("n 400") || s.includes("n-400") || s.includes("n400") || s.includes("fill")) return 1;
  return 2;
}

function findTuteeByTutorContinuation(tutor, tutees) {
  const target = nameKey(tutor.tuteeContinuationName);
  if (!target) return null;
  let exact = tutees.find(t => t.key && (t.key === target || target.includes(t.key) || t.key.includes(target)));
  if (exact) return exact;
  const targetTokens = new Set(target.split(" ").filter(Boolean));
  return tutees.map(t => ({ t, score: [...targetTokens].filter(x => t.key.includes(x)).length })).sort((a, b) => b.score - a.score)[0]?.score ? tutees.map(t => ({ t, score: [...targetTokens].filter(x => t.key.includes(x)).length })).sort((a, b) => b.score - a.score)[0].t : null;
}

function tutorMatchesTuteeLanguage(tutor, tutee) {
  if (!tutee.nativeLanguages.length) return false;
  return tutee.nativeLanguages.some(lang => hasLang(tutor, lang));
}
function hasLang(tutor, lang) { return tutor.languages.some(l => normalize(l) === normalize(lang)); }
function intersectSlots(a, b) { const bs = new Set(b); return a.filter(x => bs.has(x)); }
function unionAssignedOverlaps(tutors, tuteeSlots) { return [...new Set(tutors.flatMap(t => intersectSlots(t.availability, tuteeSlots)))].sort(slotSort); }
function sessionKey(tutee, sessionNumber) { return `${tutee.key || tutee.originalIndex}::${sessionNumber}`; }
function reviewRow(issue, tutor, tutee, note) { return { Issue: issue, Tutor: tutor, Tutee: tutee, Note: note }; }

function normalizeMode(text) {
  const s = normalize(text);
  if (!s || s.includes("no preference") || s.includes("both")) return "no preference";
  if (s.includes("person")) return "in-person";
  if (s.includes("zoom") || s.includes("virtual") || s.includes("online")) return "zoom";
  return "no preference";
}
function modeCompatible(tuteeMode, tutorMode) { return tuteeMode === "no preference" || tutorMode === "no preference" || tuteeMode === tutorMode; }

function dayFromText(text) {
  const n = normalize(text);
  for (const [k, v] of Object.entries(DAY_MAP)) if (new RegExp(`\\b${k}\\b`).test(n)) return v;
  const bracket = String(text).match(/\[(MON|TUE|WED|THU|FRI|SAT|SUN)\]/i);
  return bracket ? bracket[1].toUpperCase() : "";
}
function normalizeTime(text) { return clean(text).toUpperCase().replace(/\s+/g, "").replace(/[–—]/g, "-"); }
function slotSort(a, b) { const [da, ta] = a.split("|"); const [db, tb] = b.split("|"); return DAYS.indexOf(da) - DAYS.indexOf(db) || ta.localeCompare(tb); }

function splitMulti(value) {
  return String(value ?? "")
    .split(/,|;|\n|\||\/| and /i)
    .map(x => clean(x))
    .filter(Boolean);
}
function clean(value) { return String(value ?? "").replace(/\s+/g, " ").trim(); }
function normalize(value) { return clean(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, " ").trim(); }
function nameKey(value) { return normalize(value).replace(/\b(mr|mrs|ms|miss|dr)\b/g, "").replace(/\s+/g, " ").trim(); }
function yesish(value) { const s = normalize(value); return ["true", "yes", "y", "1", "current", "returning"].some(x => s === x || s.includes(x)); }
function hasPosition(value, position) { return normalize(value).includes(normalize(position)); }
function toNumber(value, fallback = 0) { const n = Number(String(value ?? "").replace(/[^0-9.-]/g, "")); return Number.isFinite(n) ? n : fallback; }
function boolDesc(a, b) { return Number(Boolean(b)) - Number(Boolean(a)); }
function numDesc(a, b) { return (Number(b) || 0) - (Number(a) || 0); }
function numAsc(a, b) { return (Number(a) || 0) - (Number(b) || 0); }
function strAsc(a, b) { return clean(a).localeCompare(clean(b), undefined, { sensitivity: "base" }); }
function dateAsc(a, b) { if (!a && !b) return 0; if (!a) return 1; if (!b) return -1; return a - b; }
function yesNo(x) { return x ? "Yes" : "No"; }
function parseDateValue(value) {
  if (value == null || value === "") return null;
  if (typeof value === "number") return new Date(Math.round((value - 25569) * 86400 * 1000));
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? null : d;
}
function dateToYMD(date) { return date.toISOString().slice(0, 10); }
function todayStamp() { return new Date().toISOString().slice(0, 10); }
function setStatus(id, msg, cls) { const el = $(id); el.className = `status ${cls || "muted"}`; el.textContent = msg; }
function escapeHtml(s) { return s.replace(/[&<>'"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[c])); }
