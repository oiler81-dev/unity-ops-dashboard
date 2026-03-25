const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

let currentUser = null;
let currentWeekData = null;

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
  const res = await fetch(url, {
    method: "GET",
    headers: {
      "Accept": "application/json"
    }
  });

  return parseApiResponse(res);
}

async function apiPost(url, payload) {
  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    body: JSON.stringify(payload)
  });

  return parseApiResponse(res);
}

function setStatus(message, isError = false) {
  const el = document.getElementById("statusMessage");
  if (!el) return;
  el.textContent = message;
  el.style.color = isError ? "#ff8a8a" : "#7CFC98";
}

function setDebug(data) {
  const el = document.getElementById("debugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function getDefaultWeekEnding() {
  const d = new Date();
  const diff = (5 - d.getDay() + 7) % 7;
  d.setDate(d.getDate() + diff);
  return d.toISOString().slice(0, 10);
}

function renderUser(userData) {
  document.getElementById("userInfo").innerText =
    `${userData.user.userDetails} (${userData.access.role})`;
}

function setupEntityDropdown(userData) {
  const select = document.getElementById("entitySelect");
  select.innerHTML = "";

  if (userData.access.isAdmin) {
    ENTITIES.forEach((entity) => {
      const option = document.createElement("option");
      option.value = entity;
      option.textContent = entity;
      select.appendChild(option);
    });
  } else {
    const option = document.createElement("option");
    option.value = userData.access.entity;
    option.textContent = userData.access.entity;
    select.appendChild(option);
    select.disabled = true;
  }
}

function getSelectedEntity() {
  return document.getElementById("entitySelect").value;
}

function renderForm() {
  const fields = [
    { key: "visitVolume", label: "Visit Volume" },
    { key: "callVolume", label: "Call Volume" },
    { key: "newPatients", label: "New Patients" },
    { key: "noShowRate", label: "No Show Rate" },
    { key: "cancellationRate", label: "Cancellation Rate" },
    { key: "abandonedCallRate", label: "Abandoned Call Rate" }
  ];

  const container = document.getElementById("kpiForm");
  container.innerHTML = "";

  fields.forEach((field) => {
    const div = document.createElement("div");
    div.style.marginBottom = "12px";
    div.innerHTML = `
      <label for="${field.key}" style="display:block;margin-bottom:4px;">${field.label}</label>
      <input type="number" id="${field.key}" step="any" />
    `;
    container.appendChild(div);
  });
}

function setFormValues(data) {
  const keys = [
    "visitVolume",
    "callVolume",
    "newPatients",
    "noShowRate",
    "cancellationRate",
    "abandonedCallRate"
  ];

  keys.forEach((key) => {
    const input = document.getElementById(key);
    if (!input) return;
    input.value = data && data[key] != null ? data[key] : "";
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

function updateButtonState() {
  const saveBtn = document.getElementById("saveBtn");
  const submitBtn = document.getElementById("submitBtn");
  const approveBtn = document.getElementById("approveBtn");

  const status = currentWeekData?.status || "draft";
  const isAdmin = !!currentUser?.access?.isAdmin;

  saveBtn.disabled = status === "approved";
  submitBtn.disabled = status === "submitted" || status === "approved";
  approveBtn.disabled = !isAdmin || status !== "submitted";
}

async function loadWeek() {
  const weekEnding = document.getElementById("weekEnding").value;
  const entity = getSelectedEntity();

  setStatus("Loading...");
  const result = await apiGet(
    `/api/weekly?weekEnding=${encodeURIComponent(weekEnding)}&entity=${encodeURIComponent(entity)}`
  );

  currentWeekData = result;

  setFormValues(result.data || {});
  updateButtonState();

  setStatus(`Loaded ${entity} for ${weekEnding} (${result.status || "draft"})`);
  setDebug(result);
}

async function saveWeek() {
  const payload = {
    weekEnding: document.getElementById("weekEnding").value,
    entity: getSelectedEntity(),
    data: getFormValues()
  };

  setStatus("Saving...");
  setDebug(payload);

  const result = await apiPost("/api/weekly-save", payload);

  setStatus(result.message || "Saved successfully");
  setDebug(result);

  await loadWeek();
}

async function submitWeek() {
  const payload = {
    weekEnding: document.getElementById("weekEnding").value,
    entity: getSelectedEntity()
  };

  setStatus("Submitting...");
  setDebug(payload);

  const result = await apiPost("/api/submit-week", payload);

  setStatus(result.message || "Submitted successfully");
  setDebug(result);

  await loadWeek();
}

async function approveWeek() {
  const payload = {
    weekEnding: document.getElementById("weekEnding").value,
    entity: getSelectedEntity()
  };

  setStatus("Approving...");
  setDebug(payload);

  const result = await apiPost("/api/approve-week", payload);

  setStatus(result.message || "Approved successfully");
  setDebug(result);

  await loadWeek();
}

(async function init() {
  try {
    currentUser = await apiGet("/api/me");

    renderUser(currentUser);
    setupEntityDropdown(currentUser);
    renderForm();

    const weekInput = document.getElementById("weekEnding");
    weekInput.value = getDefaultWeekEnding();

    document.getElementById("entitySelect").addEventListener("change", async () => {
      try {
        await loadWeek();
      } catch (error) {
        setStatus(error.message || "Failed to load week", true);
        setDebug(String(error));
        console.error(error);
      }
    });

    weekInput.addEventListener("change", async () => {
      try {
        await loadWeek();
      } catch (error) {
        setStatus(error.message || "Failed to load week", true);
        setDebug(String(error));
        console.error(error);
      }
    });

    document.getElementById("saveBtn").addEventListener("click", async () => {
      try {
        await saveWeek();
      } catch (error) {
        setStatus(error.message || "Failed to save", true);
        setDebug(String(error));
        console.error(error);
      }
    });

    document.getElementById("submitBtn").addEventListener("click", async () => {
      try {
        await submitWeek();
      } catch (error) {
        setStatus(error.message || "Failed to submit", true);
        setDebug(String(error));
        console.error(error);
      }
    });

    document.getElementById("approveBtn").addEventListener("click", async () => {
      try {
        await approveWeek();
      } catch (error) {
        setStatus(error.message || "Failed to approve", true);
        setDebug(String(error));
        console.error(error);
      }
    });

    await loadWeek();
  } catch (error) {
    setStatus(error.message || "Failed to load app", true);
    setDebug(String(error));
    console.error(error);
  }
})();
