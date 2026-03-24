async function apiGet(url) {
  const res = await fetch(url);
  const data = await res.json();
  if (!res.ok) {
    throw new Error(data?.error || "Request failed");
  }
  return data;
}

async function apiPost(url, payload) {
  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  const data = await res.json();
  if (!res.ok) {
    throw new Error(data?.error || "Request failed");
  }
  return data;
}

function getDefaultWeekEnding() {
  const today = new Date();
  const date = new Date(today);
  const day = date.getDay();
  const diffToFriday = (5 - day + 7) % 7;
  date.setDate(date.getDate() + diffToFriday);
  return date.toISOString().slice(0, 10);
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

async function loadWeek(userData) {
  const weekEnding = document.getElementById("weekEnding").value;
  const entity = userData.access.entity === "admin" ? "LAOSS" : userData.access.entity;

  const result = await apiGet(
    `/api/weekly?weekEnding=${encodeURIComponent(weekEnding)}&entity=${encodeURIComponent(entity)}`
  );

  setFormValues(result.data);
}

async function saveWeek(userData) {
  const weekEnding = document.getElementById("weekEnding").value;
  const entity = userData.access.entity === "admin" ? "LAOSS" : userData.access.entity;

  const payload = {
    weekEnding,
    entity,
    data: getFormValues()
  };

  const result = await apiPost("/api/weekly-save", payload);
  alert(result.message);

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
      await loadWeek(userData);
    });

    document.getElementById("saveBtn").addEventListener("click", async () => {
      await saveWeek(userData);
    });
  } catch (error) {
    alert(error.message || "Failed to load app");
    console.error(error);
  }
})();
