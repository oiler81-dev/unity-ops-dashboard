async function parseApiResponse(res) {
  const text = await res.text();
  let data = null;

  try {
    data = text ? JSON.parse(text) : null;
  } catch (error) {
    throw new Error(`Non-JSON response from ${res.url}: ${text || "[empty response]"}`);
  }

  if (!res.ok) {
    throw new Error(data?.error || data?.details || `Request failed with status ${res.status}`);
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

function getDefaultWeekEnding() {
  const today = new Date();
  const date = new Date(today);
  const day = date.getDay();
  const diffToFriday = (5 - day + 7) % 7;
  date.setDate(date.getDate() + diffToFriday);
  return date.toISOString().slice(0, 10);
}

function setStatus(message, isError = false) {
  const el = document.getElementById("statusMessage");
  el.textContent = message;
  el.style.color = isError ? "#ff8a8a" : "#7CFC98";
}

function setDebug(data) {
  const el = document.getElementById("debugOutput");
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function renderUser(userData) {
  document.getElementById("userInfo").innerText =
    `${userData.user.userDetails} (${userData.access.role})`;
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
    visitVolume: Number(document.getElementById("visitVolume").value || 0),
    callVolume: Number(document.getElementById("callVolume").value || 0),
    newPatients: Number(document.getElementById("newPatients").value || 0),
    noShowRate: Number(document.getElementById("noShowRate").value || 0),
    cancellationRate: Number(document.getElementById("cancellationRate").value || 0),
    abandonedCallRate: Number(document.getElementById("abandonedCallRate").value || 0)
  };
}

function resolveEntity(userData) {
  return userData.access.entity === "admin" ? "LAOSS" : userData.access.entity;
}

async function loadWeek(userData) {
  const weekEnding = document.getElementById("weekEnding").value;
  const entity = resolveEntity(userData);

  setStatus("Loading week...");
  const result = await apiGet(
    `/api/weekly?weekEnding=${encodeURIComponent(weekEnding)}&entity=${encodeURIComponent(entity)}`
  );

  setFormValues(result.data || {});
  setStatus(`Loaded ${entity} for ${weekEnding}`);
  setDebug(result);
}

async function saveWeek(userData) {
  const weekEnding = document.getElementById("weekEnding").value;
  const entity = resolveEntity(userData);

  const payload = {
    weekEnding,
    entity,
    data: getFormValues()
  };

  setStatus("Saving...");
  setDebug(payload);

  const result = await apiPost("/api/weekly-save", payload);

  setStatus(result.message || "Saved successfully");
  setDebug(result);

  await loadWeek(userData);
}

(async function init() {
  try {
    const userData = await apiGet("/api/me");

    renderUser(userData);
    renderForm();

    const weekInput = document.getElementById("weekEnding");
    weekInput.value = getDefaultWeekEnding();

    await loadWeek(userData);

    weekInput.addEventListener("change", async () => {
      try {
        await loadWeek(userData);
      } catch (error) {
        setStatus(error.message || "Failed to load week", true);
        setDebug(String(error));
        console.error(error);
      }
    });

    document.getElementById("saveBtn").addEventListener("click", async () => {
      try {
        await saveWeek(userData);
      } catch (error) {
        setStatus(error.message || "Failed to save", true);
        setDebug(String(error));
        console.error(error);
      }
    });
  } catch (error) {
    setStatus(error.message || "Failed to load app", true);
    setDebug(String(error));
    console.error(error);
  }
})();
