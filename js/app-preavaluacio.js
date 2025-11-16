/* ========= CONFIGURACIÓ BÀSICA ========= */
// Calculem el redirectUri per defecte a partir de la URL actual
const path = window.location.pathname.endsWith(".html")
  ? window.location.pathname
  : (window.location.pathname.endsWith("/") ? window.location.pathname : window.location.pathname + "/");
const redirectGuess = window.location.origin + path;

// Agafem la config del fitxer extern i, si no porta redirectUri, fem servir el calculat
const baseConfig = (window.PREAVAL_CONFIG || {});
const config = Object.assign({}, baseConfig, {
  redirectUri: baseConfig.redirectUri && baseConfig.redirectUri.trim() !== "" ? baseConfig.redirectUri : redirectGuess
});

/* ========= MSAL + GRAPH ========= */
const msalConfig = {
  auth: {
    clientId: config.clientId,
    authority: `https://login.microsoftonline.com/${config.tenantId}`,
    redirectUri: config.redirectUri
  },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
};

const graphScopes = ["User.Read","Files.ReadWrite", ...(config.useSharePointSite ? ["Sites.ReadWrite.All"] : [])];
const msalApp = new msal.PublicClientApplication(msalConfig);

const $ = (id) => document.getElementById(id);

let currentUserEmail = null;
let alumnes = [];
let assignacions = [];
let professors = [];
let preevaluacions = [];
let preCols = [];
let preFileUrlCache = null;

let currentProfessor = null;
let currentModule = null;
let currentStudentsInModule = [];
let currentStudent = null;
let currentPreeval = null;

/* Wizard */
let formSteps = [];
let currentStep = 1;
let totalSteps = 1;

/* ========= UTILITATS ========= */
function log(msg, cls="muted"){
  const el = $("status");
  el.className = cls;
  el.textContent = msg || "";
}

function showView(view){
  const v1 = $("view-modules");
  const v2 = $("view-students");
  const v3 = $("view-form");
  v1.classList.add("hidden");
  v2.classList.add("hidden");
  v3.classList.add("hidden");

  if(view === "modules"){ v1.classList.remove("hidden"); }
  else if(view === "students"){ v2.classList.remove("hidden"); }
  else { v3.classList.remove("hidden"); }
}

function resetUI(){
  $("app").classList.add("hidden");
  $("appbar").classList.add("hidden");
  $("hero").classList.remove("hidden");
  $("whoami").textContent = "—";
  currentUserEmail = null;
  alumnes = [];
  assignacions = [];
  professors = [];
  preevaluacions = [];
  preCols = [];
  preFileUrlCache = null;
  currentProfessor = null;
  currentModule = null;
  currentStudentsInModule = [];
  currentStudent = null;
  currentPreeval = null;
  log("");
  $("modulesTableBody").innerHTML = "";
  $("studentsTableBody").innerHTML = "";
  $("preComment").value = "";
  $("preRecommendation").value = "";
  $("preGrade").value = "";
  $("preTrafficLight").value = "";
  $("preAcademic").value = "1";
  $("preBehavior").value = "1";
  $("preInteractions").value = "1";
  $("preMotivation").value = "1";
  $("preNeeds").value = "";
  const cExtra = $("preCommentExtra");
  if(cExtra) cExtra.value = "";
  $("saveStatus").textContent = "";
}

/* ========= GRAPH HELPERS ========= */
async function getToken(){
  const account = msalApp.getAllAccounts()[0];
  const req = { account, scopes: graphScopes };
  try{
    return (await msalApp.acquireTokenSilent(req)).accessToken;
  }catch(e){
    return (await msalApp.acquireTokenPopup({ scopes: graphScopes })).accessToken;
  }
}

async function gfetch(url, opts={}){
  const token = await getToken();
  const r = await fetch(url, {
    ...opts,
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(opts.headers || {})
    }
  });
  if(!r.ok){
    let msg = "";
    try{ msg = await r.text(); }catch{}
    throw new Error(msg || (r.status + " error"));
  }
  return r.json();
}

function graphBase(){
  return config.useSharePointSite
    ? `https://graph.microsoft.com/v1.0/sites/${config.sharePoint.hostname}:${config.sharePoint.sitePath}:/drive`
    : "https://graph.microsoft.com/v1.0/me/drive";
}

function graphPathToFile(base, path){
  return `${base}/root:${encodeURI(path)}:`;
}

async function loadTableRows(filePath, tableName){
  const base = graphBase();
  const fileUrl = graphPathToFile(base, filePath);
  const cols = await gfetch(`${fileUrl}/workbook/tables/${encodeURIComponent(tableName)}/columns`);
  const colNames = cols.value.map(c => c.name);
  const rows = await gfetch(`${fileUrl}/workbook/tables/${encodeURIComponent(tableName)}/rows`);
  return { colNames, rows: rows.value.map(r => r.values[0]) };
}

function requireColumns(actual, required, tableName){
  for(const c of required){
    if(!actual.includes(c)){
      throw new Error(`Falta la columna "${c}" a la taula ${tableName}.`);
    }
  }
}

/* ========= CÀRREGA DE DADES ========= */
async function initData(){
  log("Carregant dades d'Excel…");

  // 1) Alumnes
  const A = await loadTableRows(config.refsFilePath, config.refsTables.alumnes);
  requireColumns(A.colNames, ["course","group","student_id","student_name","student_email"], "Alumnes");
  const ia = Object.fromEntries(A.colNames.map((n,i) => [n,i]));
  alumnes = A.rows.map(v => ({
    course: v[ia.course],
    group: v[ia.group],
    student_id: v[ia.student_id],
    student_name: v[ia.student_name],
    student_email: v[ia.student_email]
  }));

  // 2) Assignacions
  const S = await loadTableRows(config.refsFilePath, config.refsTables.assign);
  requireColumns(S.colNames, ["course","group","module_code","module_name","prof_email"], "Assignacions");
  const is = Object.fromEntries(S.colNames.map((n,i) => [n,i]));
  assignacions = S.rows.map(v => ({
    course: v[is.course],
    group: v[is.group],
    module_code: v[is.module_code],
    module_name: v[is.module_name],
    prof_email: (v[is.prof_email] || "").toLowerCase()
  }));

  // 3) Professors
  try{
    const P = await loadTableRows(config.refsFilePath, config.refsTables.professors);
    requireColumns(P.colNames, ["prof_email","profe_name","profe_lastname"], "Professors");
    const ip = Object.fromEntries(P.colNames.map((n,i) => [n,i]));
    professors = P.rows.map(v => ({
      prof_email: (v[ip.prof_email] || "").toLowerCase(),
      profe_name: v[ip.profe_name] || "",
      profe_lastname: v[ip.profe_lastname] || ""
    }));
  }catch(e){
    console.warn("No s'ha pogut llegir la taula Professors (es mostrarà només el correu).", e);
    professors = [];
  }

  // 4) Preevaluacions
  try{
    const base = graphBase();
    const fileUrl = graphPathToFile(base, config.preFilePath);
    preFileUrlCache = fileUrl;

    const cols = await gfetch(`${fileUrl}/workbook/tables/${encodeURIComponent(config.preTableName)}/columns`);
    preCols = cols.value.map(c => c.name);

    requireColumns(
      preCols,
      ["course","group","student_id","student_name","module_code","module_name",
       "nota_0a10","academic_1a4","comportament_1a4","interaccions_1a4","motivacio_1a4",
       "necessitats","semafor","comentari","prof_email"],
      config.preTableName
    );

    const rowsRes = await gfetch(`${fileUrl}/workbook/tables/${encodeURIComponent(config.preTableName)}/rows`);
    const rows = rowsRes.value || [];
    const ipre = Object.fromEntries(preCols.map((n,i) => [n,i]));

    const getVal = (v, name) =>
      (ipre[name] !== undefined && v[ipre[name]] !== undefined) ? v[ipre[name]] : "";

    preevaluacions = rows.map((r, idx) => {
      const v = r.values[0];
      const idxRow = (typeof r.index === "number") ? r.index : idx;
      return {
        _rowIndex: idxRow,
        rawValues: v.slice(),
        course: getVal(v,"course") || "",
        group: getVal(v,"group") || "",
        student_id: getVal(v,"student_id") || "",
        student_name: getVal(v,"student_name") || "",
        module_code: getVal(v,"module_code") || "",
        module_name: getVal(v,"module_name") || "",
        nota_0a10: getVal(v,"nota_0a10") || "",
        academic_1a4: getVal(v,"academic_1a4") || "",
        comportament_1a4: getVal(v,"comportament_1a4") || "",
        interaccions_1a4: getVal(v,"interaccions_1a4") || "",
        motivacio_1a4: getVal(v,"motivacio_1a4") || "",
        necessitats: getVal(v,"necessitats") || "",
        semafor: getVal(v,"semafor") || "",
        comentari: getVal(v,"comentari") || "",
        prof_email: (getVal(v,"prof_email") || "").toLowerCase(),
        prof_name: getVal(v,"prof_name") || "",
        timestamp: getVal(v,"timestamp") || "",
        recomanacio: getVal(v,"recomanacio") || ""
      };
    });
  }catch(e){
    console.warn("No s'ha pogut llegir la taula de Preavaluacions (es crearà a mesura que desis).", e);
    preevaluacions = [];
    preCols = ["course","group","student_id","student_name","module_code","module_name",
               "nota_0a10","academic_1a4","comportament_1a4","interaccions_1a4","motivacio_1a4",
               "necessitats","semafor","comentari","prof_email","prof_name","timestamp","recomanacio"];
    preFileUrlCache = null;
  }

  updateProfessorHeader();
  renderModules();
  log("Dades carregades correctament.","ok");
}

/* ========= PROFESSOR ACTUAL ========= */
function updateProfessorHeader(){
  currentProfessor = null;
  if(currentUserEmail){
    const lower = currentUserEmail.toLowerCase();
    currentProfessor = professors.find(p => p.prof_email === lower) || null;
  }
  if(currentProfessor){
    $("whoami").textContent = `${currentProfessor.profe_name} ${currentProfessor.profe_lastname} (${currentUserEmail})`;
  }else{
    $("whoami").textContent = currentUserEmail || "—";
  }
}

/* ========= HELPERS DE NEGOCI ========= */
function hasPreevaluation(profEmail, moduleCode, studentId){
  const email = (profEmail || "").toLowerCase();
  const sid = String(studentId || "");
  return preevaluacions.some(p =>
    p.prof_email === email &&
    p.module_code === moduleCode &&
    String(p.student_id) === sid
  );
}

function findPreevaluation(profEmail, moduleCode, studentId){
  const email = (profEmail || "").toLowerCase();
  const sid = String(studentId || "");
  return preevaluacions.find(p =>
    p.prof_email === email &&
    p.module_code === moduleCode &&
    String(p.student_id) === sid
  ) || null;
}

function getModulesForCurrentProfessor(){
  if(!currentUserEmail) return [];
  const email = currentUserEmail.toLowerCase();
  const mine = assignacions.filter(a => a.prof_email === email);
  const seen = new Set();
  const res = [];
  for(const a of mine){
    const key = `${a.module_code}||${a.module_name}`;
    if(!seen.has(key)){
      seen.add(key);
      res.push({ module_code: a.module_code, module_name: a.module_name });
    }
  }
  return res;
}

function getStudentsForModule(moduleCode){
  if(!currentUserEmail) return [];
  const email = currentUserEmail.toLowerCase();
  const assignmentsForModule = assignacions.filter(a => a.prof_email === email && a.module_code === moduleCode);
  if(!assignmentsForModule.length) return [];
  const allowedPairs = new Set(assignmentsForModule.map(a => `${a.course}||${a.group}`));
  return alumnes.filter(stu => allowedPairs.has(`${stu.course}||${stu.group}`));
}

/* ========= RENDERITZAT DE VISTES ========= */
function renderModules(){
  const tbody = $("modulesTableBody");
  tbody.innerHTML = "";
  const modules = getModulesForCurrentProfessor();
  if(!modules.length){
    $("noModules").style.display = "block";
    showView("modules");
    return;
  }
  $("noModules").style.display = "none";

  modules.forEach(mod => {
    const students = getStudentsForModule(mod.module_code);
    const total = students.length;
    let done = 0;
    students.forEach(stu => {
      if(hasPreevaluation(currentUserEmail, mod.module_code, stu.student_id)) done++;
    });

    const tr = document.createElement("tr");
    tr.className = "clickable-row";
    tr.addEventListener("click", () => openModule(mod));

    const tdName = document.createElement("td");
    tdName.textContent = mod.module_name || "—";

    const tdCode = document.createElement("td");
    tdCode.textContent = mod.module_code || "—";

    const tdCount = document.createElement("td");
    tdCount.textContent = total ? `${done} / ${total}` : "0 / 0";

    const tdAction = document.createElement("td");
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn btn-small";
    btn.textContent = "Veure alumnes";
    btn.addEventListener("click", (ev) => {
      ev.stopPropagation();
      openModule(mod);
    });
    tdAction.appendChild(btn);

    tr.appendChild(tdName);
    tr.appendChild(tdCode);
    tr.appendChild(tdCount);
    tr.appendChild(tdAction);
    tbody.appendChild(tr);
  });

  showView("modules");
}

function openModule(mod){
  currentModule = mod;
  currentStudentsInModule = getStudentsForModule(mod.module_code);
  renderStudents();
  showView("students");
}

function renderStudents(){
  const tbody = $("studentsTableBody");
  tbody.innerHTML = "";
  const filter = $("filterStatus").value || "all";
  const students = currentStudentsInModule || [];
  const email = currentUserEmail || "";
  const modCode = currentModule ? currentModule.module_code : "";

  const total = students.length;
  let doneCount = 0;
  students.forEach(stu => {
    if(hasPreevaluation(email, modCode, stu.student_id)) doneCount++;
  });

  $("studentsModuleTitle").textContent = currentModule
    ? `Alumnes del mòdul ${currentModule.module_name || ""} (${currentModule.module_code || ""})`
    : "Alumnes del mòdul";
  $("studentsModuleSummary").textContent = total
    ? `${doneCount} de ${total} alumnes tenen preavaluació.`
    : "Aquest mòdul no té cap alumne assignat.";

  if(!total){
    $("noStudents").style.display = "block";
    return;
  }
  $("noStudents").style.display = "none";

  students.forEach(stu => {
    const isDone = hasPreevaluation(email, modCode, stu.student_id);
    if(filter === "pending" && isDone) return;
    if(filter === "done" && !isDone) return;

    const tr = document.createElement("tr");
    tr.className = "clickable-row";
    tr.addEventListener("click", () => openFormForStudent(stu));

    const tdName = document.createElement("td");
    tdName.textContent = `${stu.student_name || "—"} (${stu.student_id || "?"})`;

    const tdGroup = document.createElement("td");
    tdGroup.textContent = `${stu.course || ""} / ${stu.group || ""}`;

    const tdStatus = document.createElement("td");
    const span = document.createElement("span");
    span.className = "pill " + (isDone ? "done" : "pending");
    span.textContent = isDone ? "Fet" : "Pendent";
    tdStatus.appendChild(span);

    const tdAction = document.createElement("td");
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn btn-small";
    btn.textContent = isDone ? "Veure / editar" : "Fer preavaluació";
    btn.addEventListener("click", (ev) => {
      ev.stopPropagation();
      openFormForStudent(stu);
    });
    tdAction.appendChild(btn);

    tr.appendChild(tdName);
    tr.appendChild(tdGroup);
    tr.appendChild(tdStatus);
    tr.appendChild(tdAction);
    tbody.appendChild(tr);
  });
}

/* ========= FORMULARI / WIZARD ========= */
function openFormForStudent(stu){
  currentStudent = stu;
  const email = currentUserEmail || "";
  const modCode = currentModule ? currentModule.module_code : "";

  const pe = findPreevaluation(email, modCode, stu.student_id);
  currentPreeval = pe;

  $("formTitle").textContent = `Preavaluació de ${stu.student_name || "—"} (${stu.student_id || "?"})`;
  const profName = currentProfessor
    ? `${currentProfessor.profe_name} ${currentProfessor.profe_lastname}`
    : email;
  const modName = currentModule ? `${currentModule.module_name || ""} (${currentModule.module_code || ""})` : "";
  $("formSubtitle").textContent = `${profName} — ${modName}`;

  if(pe){
    $("preComment").value      = pe.comentari || "";
    const cExtra = $("preCommentExtra");
    if(cExtra) cExtra.value = "";
    $("preRecommendation").value = pe.recomanacio || "";
    $("preGrade").value        = pe.nota_0a10 || "";
    $("preTrafficLight").value = pe.semafor || "";
    $("preAcademic").value     = pe.academic_1a4 || "1";
    $("preBehavior").value     = pe.comportament_1a4 || "1";
    $("preInteractions").value = pe.interaccions_1a4 || "1";
    $("preMotivation").value   = pe.motivacio_1a4 || "1";
    $("preNeeds").value        = pe.necessitats || "";
    $("saveStatus").textContent = "Preavaluació carregada des d'Excel.";
  }else{
    $("preComment").value = "";
    const cExtra = $("preCommentExtra");
    if(cExtra) cExtra.value = "";
    $("preRecommendation").value = "";
    $("preGrade").value = "";
    $("preTrafficLight").value = "";
    $("preAcademic").value = "1";
    $("preBehavior").value = "1";
    $("preInteractions").value = "1";
    $("preMotivation").value = "1";
    $("preNeeds").value = "";
    $("saveStatus").textContent = "Fent una nova preavaluació.";
  }

  syncAllControlsFromModel();
  currentStep = 1;
  updateStepUI();
  showView("form");
}

async function ensurePreFileUrl(){
  if(preFileUrlCache) return preFileUrlCache;
  const base = graphBase();
  preFileUrlCache = graphPathToFile(base, config.preFilePath);
  return preFileUrlCache;
}

function buildValuesFromPreeval(pe){
  const values = new Array(preCols.length).fill("");
  const map = {
    course: pe.course,
    group: pe.group,
    student_id: pe.student_id,
    student_name: pe.student_name || "",
    module_code: pe.module_code,
    module_name: pe.module_name || "",
    nota_0a10: pe.nota_0a10 || "",
    academic_1a4: pe.academic_1a4 || "",
    comportament_1a4: pe.comportament_1a4 || "",
    interaccions_1a4: pe.interaccions_1a4 || "",
    motivacio_1a4: pe.motivacio_1a4 || "",
    necessitats: pe.necessitats || "",
    semafor: pe.semafor || "",
    comentari: pe.comentari || "",
    prof_email: pe.prof_email || "",
    prof_name: pe.prof_name || "",
    timestamp: pe.timestamp || "",
    recomanacio: pe.recomanacio || ""
  };
  preCols.forEach((name, idx) => {
    if(Object.prototype.hasOwnProperty.call(map, name)){
      values[idx] = map[name];
    }else if(pe.rawValues && pe.rawValues.length > idx){
      values[idx] = pe.rawValues[idx];
    }else{
      values[idx] = "";
    }
  });
  return values;
}

async function savePreevaluation(){
  if(!currentStudent || !currentModule || !currentUserEmail){
    alert("Falta informació per desar la preavaluació.");
    return;
  }

  const commentTop       = $("preComment").value.trim();
  const commentExtraEl   = $("preCommentExtra");
  const commentExtra     = commentExtraEl ? commentExtraEl.value.trim() : "";
  const recommendation   = $("preRecommendation").value;
  const nota_0a10        = $("preGrade").value;
  const semafor          = $("preTrafficLight").value;
  const academic_1a4     = $("preAcademic").value;
  const comportament_1a4 = $("preBehavior").value;
  const interaccions_1a4 = $("preInteractions").value;
  const motivacio_1a4    = $("preMotivation").value;
  const necessitats      = $("preNeeds").value;

  const comentari        = commentTop || commentExtra;

  if(!semafor && !nota_0a10 && !comentari){
    if(!confirm("No hi ha cap nota, semàfor ni comentari. Vols desar igualment?")){
      return;
    }
  }

  $("btnSaveForm").disabled = true;
  $("saveStatus").textContent = "Desant la preavaluació…";

  const email = (currentUserEmail || "").toLowerCase();
  const profName = currentProfessor
    ? `${currentProfessor.profe_name} ${currentProfessor.profe_lastname}`
    : email;
  const timestamp = new Date().toISOString();

  const peBase = {
    course: currentStudent.course || "",
    group: currentStudent.group || "",
    student_id: currentStudent.student_id || "",
    student_name: currentStudent.student_name || "",
    module_code: currentModule.module_code || "",
    module_name: currentModule.module_name || "",
    nota_0a10,
    academic_1a4,
    comportament_1a4,
    interaccions_1a4,
    motivacio_1a4,
    necessitats,
    semafor,
    comentari,
    prof_email: email,
    prof_name: profName,
    timestamp,
    recomanacio: recommendation
  };

  try{
    const fileUrl = await ensurePreFileUrl();
    const tableNameEnc = encodeURIComponent(config.preTableName);

    if(currentPreeval && typeof currentPreeval._rowIndex === "number"){
      const merged = { ...currentPreeval, ...peBase };
      const values = buildValuesFromPreeval(merged);

      const body = JSON.stringify({
        index: currentPreeval._rowIndex,
        values: [values]
      });

      await gfetch(`${fileUrl}/workbook/tables/${tableNameEnc}/rows/${currentPreeval._rowIndex}`, {
        method: "PATCH",
        body
      });

      Object.assign(currentPreeval, peBase, { rawValues: values.slice() });
      $("saveStatus").textContent = "Preavaluació actualitzada correctament.";
    }else{
      const tempPe = { ...peBase, rawValues: [] };
      const values = buildValuesFromPreeval(tempPe);

      const body = JSON.stringify({
        index: null,
        values: [values]
      });

      const res = await gfetch(`${fileUrl}/workbook/tables/${tableNameEnc}/rows/add`, {
        method: "POST",
        body
      });

      const newIndex = (typeof res.index === "number") ? res.index : preevaluacions.length;
      const newPe = { _rowIndex: newIndex, rawValues: values.slice(), ...peBase };
      preevaluacions.push(newPe);
      currentPreeval = newPe;
      $("saveStatus").textContent = "Preavaluació desada correctament.";
    }

    renderStudents();
    renderModules();
  }catch(e){
    console.error(e);
    $("saveStatus").textContent = "Error en desar la preavaluació.";
    alert("No s'ha pogut desar la preavaluació: " + e.message);
  }finally{
    $("btnSaveForm").disabled = false;
  }
}

/* ========= WIZARD: CONTROL DE PASSOS ========= */
function updateStepUI(){
  if(!formSteps.length) return;
  if(currentStep < 1) currentStep = 1;
  if(currentStep > totalSteps) currentStep = totalSteps;

  formSteps.forEach(step => {
    const n = parseInt(step.dataset.step || "1", 10);
    if(n === currentStep) step.classList.remove("hidden");
    else step.classList.add("hidden");
  });

  const lbl = $("stepLabel");
  const bar = $("stepProgress");
  if(lbl){
    lbl.textContent = `Pàgina ${currentStep} de ${totalSteps}`;
  }
  if(bar){
    const pct = totalSteps > 1 ? ((currentStep - 1) / (totalSteps - 1)) * 100 : 100;
    bar.style.width = pct + "%";
  }

  const btnPrev = $("btnPrevStep");
  const btnNext = $("btnNextStep");
  if(btnPrev) btnPrev.disabled = currentStep === 1;
  if(btnNext) btnNext.textContent = currentStep === totalSteps ? "Finalitzar" : "Següent";
}

function nextStep(){
  if(currentStep < totalSteps){
    currentStep++;
    updateStepUI();
  }
}

function prevStep(){
  if(currentStep > 1){
    currentStep--;
    updateStepUI();
  }
}

/* ========= SYNC DE CONTROLS (botons, ràdios, checkboxes) ========= */
function setupGradeButtons(){
  const container = $("gradeButtons");
  if(!container) return;
  const buttons = Array.from(container.querySelectorAll(".rating-pill"));
  buttons.forEach(btn => {
    btn.addEventListener("click", () => {
      const v = btn.dataset.value;
      $("preGrade").value = v;
      buttons.forEach(b => b.classList.toggle("selected", b === btn));
    });
  });
}

function syncGradeButtonsFromField(){
  const container = $("gradeButtons");
  if(!container) return;
  const buttons = Array.from(container.querySelectorAll(".rating-pill"));
  const val = $("preGrade").value;
  buttons.forEach(btn => {
    btn.classList.toggle("selected", String(btn.dataset.value) === String(val));
  });
}

function setupScaleRadio(name, selectId){
  const radios = Array.from(document.querySelectorAll(`input[name="${name}"]`));
  const selectEl = $(selectId);
  if(!radios.length || !selectEl) return;
  radios.forEach(radio => {
    radio.addEventListener("change", () => {
      if(radio.checked){
        selectEl.value = radio.value;
      }
    });
  });
}

function syncScaleRadioFromSelect(name, selectId){
  const radios = Array.from(document.querySelectorAll(`input[name="${name}"]`));
  const selectEl = $(selectId);
  if(!radios.length || !selectEl) return;
  const val = selectEl.value;
  radios.forEach(radio => {
    radio.checked = (radio.value === val);
  });
}

function setupTrafficLightRadios(){
  const radios = Array.from(document.querySelectorAll('input[name="preTrafficLightRadio"]'));
  const selectEl = $("preTrafficLight");
  if(!radios.length || !selectEl) return;
  radios.forEach(radio => {
    radio.addEventListener("change", () => {
      if(radio.checked){
        selectEl.value = radio.value;
      }
    });
  });
}

function syncTrafficLightRadios(){
  const radios = Array.from(document.querySelectorAll('input[name="preTrafficLightRadio"]'));
  const selectEl = $("preTrafficLight");
  if(!radios.length || !selectEl) return;
  const val = selectEl.value || "";
  radios.forEach(radio => {
    radio.checked = (radio.value === val);
  });
}

function setupNeedsCheckboxes(){
  const boxes = Array.from(document.querySelectorAll('input[data-need]'));
  const hidden = $("preNeeds");
  if(!boxes.length || !hidden) return;

  const updateHidden = () => {
    const selected = boxes.filter(b => b.checked).map(b => b.value.trim());
    hidden.value = selected.join("; ");
  };

  boxes.forEach(box => {
    box.addEventListener("change", updateHidden);
  });
}

function syncNeedsCheckboxesFromField(){
  const boxes = Array.from(document.querySelectorAll('input[data-need]'));
  const hidden = $("preNeeds");
  if(!boxes.length || !hidden) return;
  const text = hidden.value || "";
  const tokens = text.split(";").map(t => t.trim()).filter(Boolean);
  boxes.forEach(box => {
    const value = box.value.trim().toLowerCase();
    const matched = tokens.some(t => t.toLowerCase() === value);
    box.checked = matched;
  });
}

function syncAllControlsFromModel(){
  syncGradeButtonsFromField();
  syncScaleRadioFromSelect("preAcademicRadio","preAcademic");
  syncScaleRadioFromSelect("preBehaviorRadio","preBehavior");
  syncScaleRadioFromSelect("preInteractionsRadio","preInteractions");
  syncScaleRadioFromSelect("preMotivationRadio","preMotivation");
  syncNeedsCheckboxesFromField();
  syncTrafficLightRadios();
}

/* ========= LOGIN / LOGOUT ========= */
async function start(){
  try{
    $("btnStart").disabled = true;
    const res = await msalApp.loginPopup({ scopes: graphScopes, prompt: "select_account" });
    currentUserEmail = (res.account.username || "").toLowerCase();
    $("hero").classList.add("hidden");
    $("appbar").classList.remove("hidden");
    $("app").classList.remove("hidden");
    await initData();
    showView("modules");
  }catch(e){
    console.error(e);
    alert("No s'ha pogut iniciar sessió: " + e.message);
    $("btnStart").disabled = false;
  }
}

async function logout(){
  try{
    const acc = msalApp.getAllAccounts()[0];
    await msalApp.logoutPopup({ account: acc, postLogoutRedirectUri: config.redirectUri });
  }catch(e){
    console.warn("logoutPopup va fallar, es neteja la UI igualment", e);
  }finally{
    resetUI();
  }
}

/* ========= ESDEVENIMENTS ========= */
$("btnStart").addEventListener("click", start);
$("btnLogout").addEventListener("click", logout);

$("btnBackToModules").addEventListener("click", () => {
  renderModules();
  showView("modules");
});

$("btnBackToStudents").addEventListener("click", () => {
  renderStudents();
  showView("students");
});

$("filterStatus").addEventListener("change", () => {
  renderStudents();
});

$("btnSaveForm").addEventListener("click", () => {
  savePreevaluation();
});

/* Wizard buttons */
const prevBtn = $("btnPrevStep");
const nextBtn = $("btnNextStep");
if(prevBtn) prevBtn.addEventListener("click", prevStep);
if(nextBtn) nextBtn.addEventListener("click", () => {
  if(currentStep < totalSteps) nextStep();
  else {
    // última pàgina: no fem res especial (pots prémer "Desa")
  }
});

/* Inicialització del wizard i controls UI */
formSteps = Array.from(document.querySelectorAll("#formSteps .form-step"));
totalSteps = formSteps.length || 1;
currentStep = 1;
updateStepUI();

setupGradeButtons();
setupScaleRadio("preAcademicRadio","preAcademic");
setupScaleRadio("preBehaviorRadio","preBehavior");
setupScaleRadio("preInteractionsRadio","preInteractions");
setupScaleRadio("preMotivationRadio","preMotivation");
setupNeedsCheckboxes();
setupTrafficLightRadios();

/* Sessió ja oberta? */
msalApp.handleRedirectPromise().then(async () => {
  const acc = msalApp.getAllAccounts()[0];
  if(acc){
    currentUserEmail = (acc.username || "").toLowerCase();
    $("hero").classList.add("hidden");
    $("appbar").classList.remove("hidden");
    $("app").classList.remove("hidden");
    await initData();
    showView("modules");
  }
}).catch(e => console.error(e));
