let storeData = [];
let currentDate = new Date(2026, 1, 1); // Feb 2026

const calendarEl = document.getElementById("calendar");
const monthLabel = document.getElementById("monthLabel");
const filterInput = document.getElementById("storeFilter");

document.getElementById("excelUpload").addEventListener("change", handleUpload);
filterInput.addEventListener("input", renderCalendar);

function handleUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = evt => {
    const workbook = XLSX.read(evt.target.result, { type: "binary" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

    storeData = rows
      .filter(r => r["Store"] && r["Off Date"] && r["Activation Date"])
      .map(r => ({
        store: String(r["Store"]).trim(),
        off: excelDateToJSDate(r["Off Date"]),
        act: excelDateToJSDate(r["Activation Date"])
      }))
      .filter(r => r.off && r.act);

    alert(`Loaded ${storeData.length} store records`);
    renderCalendar();
  };

  reader.readAsBinaryString(file);
}

function excelDateToJSDate(value) {
  if (value instanceof Date) return normalizeDate(value);

  if (typeof value === "number") {
    return normalizeDate(new Date((value - 25569) * 86400 * 1000));
  }

  const parsed = new Date(value);
  if (!isNaN(parsed)) return normalizeDate(parsed);

  return null;
}

function normalizeDate(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function renderCalendar() {
  calendarEl.innerHTML = "";

  const year = currentDate.getFullYear();
  const month = currentDate.getMonth();

  monthLabel.textContent = currentDate.toLocaleString("default", {
    month: "long",
    year: "numeric"
  });

  ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"].forEach(d => {
    const h = document.createElement("div");
    h.className = "day-header";
    h.textContent = d;
    calendarEl.appendChild(h);
  });

  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  for (let i = 0; i < firstDay; i++) {
    calendarEl.appendChild(document.createElement("div"));
  }

  for (let day = 1; day <= daysInMonth; day++) {
    const cellDate = normalizeDate(new Date(year, month, day));
    const cell = document.createElement("div");
    cell.className = "day-cell";

    const dateNum = document.createElement("div");
    dateNum.className = "date-number";
    dateNum.textContent = day;
    cell.appendChild(dateNum);

    let count = 0;

    storeData.forEach(s => {
      if (
        cellDate >= s.off &&
        cellDate <= s.act &&
        (!filterInput.value ||
          s.store.toLowerCase().includes(filterInput.value.toLowerCase()))
      ) {
        const div = document.createElement("div");
        div.className = "store";
        div.textContent = s.store;

        if (+cellDate === +s.off) div.classList.add("off");
        else if (+cellDate === +s.act) div.classList.add("active");
        else div.classList.add("between");

        div.title = `Off: ${s.off.toDateString()} | Activation: ${s.act.toDateString()}`;
        cell.appendChild(div);
        count++;
      }
    });

    if (count > 0) {
      const c = document.createElement("div");
      c.className = "count";
      c.textContent = `Stores: ${count}`;
      cell.appendChild(c);
    }

    calendarEl.appendChild(cell);
  }
}

function prevMonth() {
  currentDate.setMonth(currentDate.getMonth() - 1);
  renderCalendar();
}

function nextMonth() {
  currentDate.setMonth(currentDate.getMonth() + 1);
  renderCalendar();
}

renderCalendar();
