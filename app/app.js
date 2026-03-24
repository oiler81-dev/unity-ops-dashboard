const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

async function parseApiResponse(res) {
  const text = await res.text();
  let data = null;

  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    throw new Error(text || "Invalid response");
  }

  if (!res.ok) {
    throw new Error(data?.details || data?.error || "Request failed");
  }

  return data;
}

async function apiGet(url) {
  const res = await fetch(url);
  return parseApiResponse(res);
}

async function apiPost(url, payload) {
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  return parseApiResponse(res);
}

function setStatus(msg, isError = false) {
  const el = document.getElementById("statusMessage");
  el.textContent = msg;
  el.style.color = isError ? "#ff8a8a" : "#7CFC98";
}

function setDebug(data) {
  document.getElementById("debugOutput").textContent =
    typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function renderUser(userData) {
  document.getElementById("userInfo").innerText =
    `${userData.user.userDetails} (${userData.access.role})`;
}

function setupEntityDropdown(userData) {
  const select = document.getElementById("entitySelect");

  select.innerHTML = "";

  if (userData.access.isAdmin) {
    ENTITIES.forEach((e) => {
      const opt = document.createElement("option");
      opt.value = e;
      opt.textContent = e;
      select.appendChild(opt);
    });
  } else {
    const opt = document.createElement("option");
    opt.value = userData.access.entity;
    opt.textContent = userData.access.entity;
    select.appendChild(opt);

    select.disabled = true;
  }
}

function getSelectedEntity() {
  return document.getElementById("entitySelect").value;
}

function renderForm() {
  const fields = [
    "visitVolume",
    "callVolume",
    "newPatients",
    "noShowRate",
    "cancellationRate",
    "abandonedCallRate"
  ];

  const container = document.getElementById("kpiForm");
  container.innerHTML = "";

  fields.forEach((key) => {
    const div = document.createElement("div");
    div.innerHTML = `
      <label>${key}</label>
      <input type="number" id="${key}" />
    `;
    container.appendChild(div);
  });
}

function setFormValues(data) {
  Object.keys(data).forEach((key) => {
    const el = document.getElementById(key);
    if (el) el.value = data[key] ?? "";
  });
}

function getFormValues() {
  return {
    visitVolume: document.getElementById("visitVolume").value,
    callVolume: document.getElementById("callVolume").value,
    newPatients: document.getElementById("newPatients").value,
    noShowRate: document.getElementById("noShowRate").value,
    cancellationRate: document.getElementById("cancellationRate").value,
    abandonedCallRate: document.getElementById("abandonedCallRate").value
  };
}

function getDefaultWeekEnding() {
  const d = new Date();
  const diff = (5 - d.getDay() + 7) % 7;
  d.setDate(d.getDate() + diff);
  return d.toISOString().slice(0, 10);
}

async function loadWeek() {
  const weekEnding = document.getElementById("weekEnding").value;
  const entity = getSelectedEntity();

  setStatus("Loading...");
  const res = await apiGet(`/api/weekly?weekEnding=${weekEnding}&entity=${entity}`);

  setFormValues(res.data);
  setStatus(`Loaded ${entity}`);
  setDebug(res);
}

async function saveWeek() {
  const payload = {
    weekEnding: document.getElementById("weekEnding").value,
    entity: getSelectedEntity(),
    data: getFormValues()
  };

  setStatus("Saving...");
  setDebug(payload);

  const res = await apiPost("/api/weekly-save", payload);

  setStatus("Saved");
  setDebug(res);

  await loadWeek();
}

(async function init() {
  try {
    const user = await apiGet("/api/me");

    renderUser(user);
    setupEntityDropdown(user);
    renderForm();

    const weekInput = document.getElementById("weekEnding");
    weekInput.value = getDefaultWeekEnding();

    document.getElementById("entitySelect").addEventListener("change", loadWeek);
    weekInput.addEventListener("change", loadWeek);
    document.getElementById("saveBtn").addEventListener("click", saveWeek);

    await loadWeek();
  } catch (e) {
    setStatus(e.message, true);
    setDebug(e);
  }
})();
