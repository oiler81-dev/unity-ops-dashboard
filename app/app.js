async function loadUser() {
  const res = await fetch("/api/me");
  const data = await res.json();

  document.getElementById("userInfo").innerText =
    data.user.userDetails + " (" + data.access.role + ")";

  return data;
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

async function saveData(user) {
  const weekEnding = document.getElementById("weekEnding").value;

  const payload = {
    weekEnding,
    entity: user.access.entity,
    data: {
      visitVolume: Number(document.getElementById("visitVolume").value),
      callVolume: Number(document.getElementById("callVolume").value),
      newPatients: Number(document.getElementById("newPatients").value),
      noShowRate: Number(document.getElementById("noShowRate").value),
      cancellationRate: Number(document.getElementById("cancellationRate").value),
      abandonedCallRate: Number(document.getElementById("abandonedCallRate").value)
    }
  };

  const res = await fetch("/api/weekly-save", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });

  const result = await res.json();
  alert(JSON.stringify(result));
}

(async function init() {
  const user = await loadUser();
  renderForm();

  document.getElementById("saveBtn").onclick = () => saveData(user);
})();
