import React, { useEffect, useMemo, useState } from "react";

const EXCEL_WEBHOOK_URL = "https://defaulta1dce605051e42ce9ba7342cabd36c.67.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/883f9b78868849b2ab5ebb5a7727a4ac/triggers/manual/paths/invoke?api-version=1";

const EXCEL_WORKBOOK_NAME = "QR_Machine_Hours_BETA.xlsx";
const EXCEL_TABLE_NAME = "EquipmentHours";
const LOCAL_QUEUE_KEY = "equipment-hours-offline-queue";

const EQUIPMENT_CODES = [
  { code: "200-1", name: "Laser" },
  { code: "200-2", name: "Water Jet" },
  { code: "200-3", name: "Burn Table" },
  { code: "200-4", name: "Break" },
  { code: "200-5", name: "Saw" },
  { code: "200-7", name: "Rolls" },
  { code: "300-1", name: "Innocenti" },
  { code: "300-5", name: "OKK" },
  { code: "300-7", name: "Knee Mill" },
  { code: "300-10", name: "Monarch" },
  { code: "300-11", name: "G&L" },
  { code: "300-12", name: "Mazak" },
  { code: "300-13", name: "Hyundai" },
];

const DEFAULT_OPERATORS = [
  "AnDrew Toole",
  "Cameron Ott",
  "Leonel Maldonado",
  "Johnny Padilla",
  "Enedino Pecina",
  "Jose Guerrero Sr",
  "Timothy Howard",
  "Allen Green",
  "Richard Thomas",
  "Claro Lopez",
  "Norman Alderman",
  "Steve Hrynkiw",
  "Samantha Morgan",
  "Adam Levesque",
  "Nichole Wilson",
  "Gilbert Rice Sr",
  "Gilbert Rice Jr",
  "Agustin Huichapa",
  "Christopher Sherry",
  "Kelly McCain",
  "Alexander Medina",
  "Miguel Lara",
  "Chestan Baker",
  "Matthew Crisp",
  "Kenneth Krontz",
  "Bryan Schnurr",
  "Corbett Combs",
  "Jeffrey Schmit",
  "Todd Allen",
  "William Rodgers",
  "Timothy Slaten",
  "Duhamel Torres",
  "Dominique Luna",
  "Angel Salmeron",
  "Kale McCloughen",
  "Daniel Bright",
  "Lucinda Watts",
  "Jeffery Ashley",
  "Brandon Pressnell",
  "Thomas Rodman"
];

const BETA_MODE = false;
const COMPANY_NAME = "Specialty Fabrication LLC";
const SAVED_OPERATORS_KEY = "equipment-hours-saved-operators";
const COMPANY_LOGO_URL = "";

function safeNow() {
  return new Date();
}

function todayIso() {
  return safeNow().toISOString().slice(0, 10);
}

function getAppBaseUrl() {
  if (typeof window === "undefined") return "https://equipment-hours-beta.local/entry";
  return `${window.location.origin}${window.location.pathname}`;
}

function isExcelWebhookConfigured() {
  return Boolean(EXCEL_WEBHOOK_URL && !EXCEL_WEBHOOK_URL.includes("PASTE_POWER_AUTOMATE"));
}

function loadSavedOperators() {
  if (typeof window === "undefined") return [];
  try {
    return JSON.parse(localStorage.getItem(SAVED_OPERATORS_KEY) || "[]");
  } catch {
    return [];
  }
}

function saveOperators(operators) {
  if (typeof window === "undefined") return;
  localStorage.setItem(SAVED_OPERATORS_KEY, JSON.stringify(operators));
}

function formatStopwatch(startedAt, stoppedAt = safeNow()) {
  if (!startedAt) return "00:00:00";

  const start = new Date(startedAt).getTime();
  const stop = new Date(stoppedAt).getTime();
  const diffSeconds = Math.max(0, Math.floor((stop - start) / 1000));

  const hours = String(Math.floor(diffSeconds / 3600)).padStart(2, "0");
  const minutes = String(Math.floor((diffSeconds % 3600) / 60)).padStart(2, "0");
  const seconds = String(diffSeconds % 60).padStart(2, "0");

  return `${hours}:${minutes}:${seconds}`;
}

function formatDurationHours(startedAt, stoppedAt = safeNow()) {
  if (!startedAt) return "0.00";

  const start = new Date(startedAt).getTime();
  const stop = new Date(stoppedAt).getTime();
  const diffSeconds = Math.max(0, (stop - start) / 1000);

  if (diffSeconds <= 10) return "0.00";

  const rawHours = diffSeconds / 60 / 60;
  const roundedUpHalfHour = Math.ceil(rawHours * 2) / 2;
  return roundedUpHalfHour.toFixed(2);
}

function loadOfflineQueue() {
  return [];
}

function saveOfflineQueue() {
  return;
}

async function sendRowToExcel(row) {
  if (!isExcelWebhookConfigured()) {
    throw new Error("Excel webhook URL is not configured yet.");
  }

  const response = await fetch("/api/submit", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(row),
  });

  if (!response.ok) {
    let message = `Excel webhook failed with status ${response.status}`;
    try {
      const details = await response.json();
      if (details?.error) message = details.error;
    } catch {
      // Keep default message.
    }
    throw new Error(message);
  }

  return true;
}

function getUniqueEquipmentCodes(items) {
  const seen = new Set();
  return items.filter((item) => {
    if (seen.has(item.code)) return false;
    seen.add(item.code);
    return true;
  });
}

function getEquipmentName(code) {
  return EQUIPMENT_CODES.find((item) => item.code === code)?.name || "";
}

function getEquipmentByCode(code) {
  return EQUIPMENT_CODES.find((item) => item.code === code) || null;
}

function hasDuplicateCodes(items) {
  return new Set(items.map((item) => item.code)).size !== items.length;
}

function buildMachineUrl(item) {
  const params = new URLSearchParams({
    equipmentCode: item.code,
    equipmentName: item.name,
    source: "machine-qr",
  });
  return `${getAppBaseUrl()}?${params.toString()}`;
}

function buildQrImageUrl(item) {
  const data = encodeURIComponent(buildMachineUrl(item));
  return `https://api.qrserver.com/v1/create-qr-code/?size=220x220&margin=12&data=${data}`;
}

function formatHours(value) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return "0.00";
  return (Math.round(number * 2) / 2).toFixed(2);
}

function isValidHalfHour(value) {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 && Math.round(number * 2) === number * 2;
}

function makeEntryId() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function buildCsv(rows) {
  const header = [
    "Employee Name",
    "Job Number",
    "DATE",
    "Equipment Code",
    "Equipment Name",
    "Equipment Hours",
    "Notes",
    "Submitted At",
    "Excel Status",
  ];

  const body = rows.map((r) => [
    r.employeeName,
    r.jobNumber,
    r.date,
    r.equipmentCode,
    r.equipmentName,
    r.equipmentHours,
    r.notes,
    r.submittedAt,
    r.excelStatus,
  ]);

  return [header, ...body]
    .map((row) => row.map((cell) => `"${String(cell ?? "").replaceAll('"', '""')}"`).join(","))
    .join("\n");
}

function runSelfTests() {
  const removedNames = EQUIPMENT_CODES.map((item) => item.name);
  console.assert(!removedNames.includes("TOS"), "TOS should not be in EQUIPMENT_CODES");
  console.assert(!removedNames.includes("Tarnow"), "Tarnow should not be in EQUIPMENT_CODES");
  console.assert(!removedNames.includes("TB B&B"), "TB B&B should not be in EQUIPMENT_CODES");
  console.assert(getEquipmentName("300-7") === "Knee Mill", "300-7 should be Knee Mill");
  console.assert(getEquipmentName("300-8") === "", "300-8 should be merged into 300-7 and removed");
  console.assert(formatHours("1.5") === "1.50", "Hours should format to two decimals");
  console.assert(isValidHalfHour("1.5"), "1.5 should be a valid half-hour entry");
  console.assert(!isValidHalfHour("1.25"), "1.25 should not be valid for half-hour increments");
  console.assert(!hasDuplicateCodes(EQUIPMENT_CODES), "Each machine QR must have a unique equipment code");
  console.assert(getUniqueEquipmentCodes(EQUIPMENT_CODES).length === EQUIPMENT_CODES.length, "QR list should not duplicate equipment codes");
  console.assert(buildMachineUrl({ code: "300-7", name: "Knee Mill" }).includes("equipmentCode=300-7"), "QR URL should include equipment code");
  console.assert(getEquipmentByCode("200-1")?.name === "Laser", "Equipment lookup should return machine details");
  console.assert(isExcelWebhookConfigured(), "Webhook should be configured with the Power Automate URL before deployment");
  console.assert(EXCEL_WORKBOOK_NAME.endsWith(".xlsx"), "Excel workbook should be an xlsx file");
  console.assert(EXCEL_TABLE_NAME === "EquipmentHours", "Power Automate should target the EquipmentHours table");
  console.assert(Array.isArray(loadOfflineQueue()), "Offline queue loader should return an array");
}

runSelfTests();

function IconBadge({ children, dark = false }) {
  return (
    <span
      className={`inline-flex h-11 w-11 items-center justify-center rounded-2xl text-xl ${
        dark ? "bg-slate-900 text-white" : "bg-white text-slate-900 shadow-sm"
      }`}
      aria-hidden="true"
    >
      {children}
    </span>
  );
}

function Button({ children, className = "", variant = "solid", ...props }) {
  const base =
    "inline-flex items-center justify-center rounded-xl px-5 py-3 text-sm font-semibold transition active:scale-[0.99] disabled:cursor-not-allowed disabled:opacity-50";
  const styles =
    variant === "outline"
      ? "border border-slate-300 bg-white text-slate-900 hover:bg-slate-100"
      : "bg-slate-900 text-white hover:bg-slate-700";

  return (
    <button className={`${base} ${styles} ${className}`} {...props}>
      {children}
    </button>
  );
}

function Card({ children, className = "" }) {
  return <section className={`rounded-2xl border-0 bg-white shadow-sm ${className}`}>{children}</section>;
}

function CardContent({ children, className = "" }) {
  return <div className={className}>{children}</div>;
}

export default function EquipmentHoursQRApp() {
  const [jobNumber, setJobNumber] = useState("");
  const [date, setDate] = useState(todayIso());
  const [equipmentCode, setEquipmentCode] = useState("");
  const [equipmentHours, setEquipmentHours] = useState("");
  const [notes, setNotes] = useState("");
  const [submitted, setSubmitted] = useState(false);
  const [employeeName, setEmployeeName] = useState("");

  const [savedOperators, setSavedOperators] = useState(() => {
    const saved = loadSavedOperators();
    return saved.length ? saved : DEFAULT_OPERATORS;
  });

  const [timerStartedAt, setTimerStartedAt] = useState(null);
  const [timerDisplay, setTimerDisplay] = useState("00:00:00");
  const [offlineQueue, setOfflineQueue] = useState([]);
  const [rows, setRows] = useState([]);
  const [qrImageFailures, setQrImageFailures] = useState({});
  const [formError, setFormError] = useState("");
  const [scannedMachine, setScannedMachine] = useState(null);
  const [excelMessage, setExcelMessage] = useState("");
  const [isSending, setIsSending] = useState(false);

  const selectedEquipment = useMemo(
    () => EQUIPMENT_CODES.find((item) => item.code === equipmentCode),
    [equipmentCode]
  );

  const machineQrCodes = useMemo(() => getUniqueEquipmentCodes(EQUIPMENT_CODES), []);

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const scannedEquipmentCode = params.get("equipmentCode");
    const machine = getEquipmentByCode(scannedEquipmentCode);

    if (machine) {
      setEquipmentCode(machine.code);
      setScannedMachine(machine);
      setExcelMessage(`Machine loaded from QR: ${machine.code} - ${machine.name}`);
    }
  }, []);

  useEffect(() => {
    saveOfflineQueue(offlineQueue);
  }, [offlineQueue]);

  useEffect(() => {
    saveOperators(savedOperators);
  }, [savedOperators]);

  useEffect(() => {
    if (!timerStartedAt) {
      setTimerDisplay("00:00:00");
      return;
    }

    const timer = window.setInterval(() => {
      const calculatedHours = formatDurationHours(timerStartedAt);
      setTimerDisplay(formatStopwatch(timerStartedAt));
      setEquipmentHours(calculatedHours);
    }, 1000);

    return () => window.clearInterval(timer);
  }, [timerStartedAt]);

  const totalHours = useMemo(
    () => rows.reduce((sum, row) => sum + Number(row.equipmentHours || 0), 0),
    [rows]
  );

  function resetForm({ keepEquipment = false } = {}) {
    setEmployeeName("");
    setJobNumber("");
    setDate(todayIso());

    if (!keepEquipment) {
      setEquipmentCode("");
      setScannedMachine(null);
    }

    setEquipmentHours("");
    setNotes("");
    setTimerStartedAt(null);
    setTimerDisplay("00:00:00");
    setFormError("");
  }

  function startMachineTimer() {
    setEquipmentHours("0.00");
    setTimerDisplay("00:00:00");
    setTimerStartedAt(new Date().toISOString());
    setExcelMessage("Machine timer started. Equipment Hours will become 0.50 after 10 seconds.");
  }

  function stopMachineTimer() {
    if (!timerStartedAt) return;
    const calculatedHours = formatDurationHours(timerStartedAt);
    setEquipmentHours(calculatedHours);
    setTimerStartedAt(null);
    setTimerDisplay(formatStopwatch(timerStartedAt));
    setExcelMessage(`Machine timer stopped. Equipment Hours set to ${calculatedHours}.`);
  }

  function validateForm() {
    if (!employeeName.trim()) return "Employee Name is required.";
    if (!jobNumber.trim()) return "Job Number is required.";
    if (!selectedEquipment) return "Scan a machine QR code or select a valid Equipment Code.";
    if (!isValidHalfHour(equipmentHours)) return "Start and stop the machine timer. Equipment Hours must be at least 0.50 before submitting.";
    return "";
  }

  function updateRowStatus(rowId, excelStatus) {
    setRows((currentRows) => currentRows.map((row) => (row.id === rowId ? { ...row, excelStatus } : row)));
  }

  function queueRow(row) {
    const queuedRow = { ...row, excelStatus: "queued" };
    setOfflineQueue((current) => [queuedRow, ...current]);
    setRows((currentRows) => currentRows.map((item) => (item.id === row.id ? queuedRow : item)));
  }

  async function handleSubmit(e) {
    e.preventDefault();

    const validationError = validateForm();
    if (validationError) {
      setFormError(validationError);
      return;
    }

    const newRow = {
      id: makeEntryId(),
      employeeName: employeeName.trim(),
      jobNumber: jobNumber.trim(),
      date,
      equipmentCode,
      equipmentName: selectedEquipment.name,
      equipmentHours: formatHours(equipmentHours),
      notes: notes.trim(),
      submittedAt: safeNow().toLocaleString(),
      excelStatus: "sending",
    };

    setRows((currentRows) => [newRow, ...currentRows]);
    setSubmitted(true);
    setIsSending(true);
    setExcelMessage("Sending to Excel...");

    try {
      await sendRowToExcel(newRow);
      updateRowStatus(newRow.id, "sent");
      setExcelMessage("Saved directly to Excel.");
    } catch (error) {
      queueRow(newRow);
      setExcelMessage(`${error.message} Entry saved in the local queue.`);
    } finally {
      setIsSending(false);
      resetForm({ keepEquipment: Boolean(scannedMachine) });
      window.setTimeout(() => setSubmitted(false), 3000);
    }
  }

  async function syncQueuedRows() {
    if (!offlineQueue.length) {
      setExcelMessage("No queued entries to sync.");
      return;
    }

    setIsSending(true);
    setExcelMessage("Syncing queued entries to Excel...");

    const stillQueued = [];

    for (const row of offlineQueue) {
      try {
        await sendRowToExcel({ ...row, excelStatus: "syncing" });
        updateRowStatus(row.id, "sent");
      } catch {
        stillQueued.push(row);
        updateRowStatus(row.id, "queued");
      }
    }

    setOfflineQueue(stillQueued);
    setExcelMessage(stillQueued.length ? `${stillQueued.length} entries still queued.` : "All queued entries synced to Excel.");
    setIsSending(false);
  }

  function downloadCsv() {
    const csv = buildCsv(rows);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "equipment_hours_charged.csv";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function printQrCodes() {
    window.print();
  }

  return (
    <div className="min-h-screen bg-slate-50 p-4 text-slate-900">
      <div className="mx-auto max-w-6xl space-y-6">
        {BETA_MODE && (
          <div className="rounded-2xl border border-amber-300 bg-amber-100 p-4 text-amber-900 shadow-sm print:hidden">
            <div className="flex flex-wrap items-center justify-between gap-4">
              <div>
                <div className="text-lg font-bold">🧪 Shop Floor Beta Test Mode</div>
                <div className="text-sm">
                  Sends entries to Excel when configured, keeps failed submissions queued locally, and supports QR machine auto-fill.
                </div>
              </div>
              <div className="rounded-xl bg-white px-4 py-2 text-sm font-semibold shadow-sm">
                Offline Queue: {offlineQueue.length}
              </div>
            </div>
          </div>
        )}

        <div className="grid gap-4 md:grid-cols-[1.2fr_.8fr] print:block">
          <Card className="print:hidden">
            <CardContent className="p-6">
              <div className="mb-6 flex items-center gap-4">
                {COMPANY_LOGO_URL ? (
                  <div className="flex items-center justify-center rounded-2xl bg-white p-3 shadow-sm border border-slate-200">
                    <img
                      src={COMPANY_LOGO_URL}
                      alt={`${COMPANY_NAME} logo`}
                      className="h-16 w-16 object-contain"
                      onError={(e) => {
                        e.currentTarget.style.display = "none";
                      }}
                    />
                  </div>
                ) : null}

                <IconBadge dark>⏱</IconBadge>
                <div>
                  <div className="text-xs font-semibold uppercase tracking-[0.2em] text-slate-500">
                    {COMPANY_NAME}
                  </div>
                  <h1 className="text-3xl font-bold tracking-tight">Equipment Hours Charged</h1>
                  <p className="text-slate-600">QR-launched entry form based on your Equipment Hours Sheet.</p>
                </div>
              </div>

              {submitted && (
                <div className="mb-4 flex items-center gap-2 rounded-xl bg-green-100 p-3 text-green-800">
                  <span aria-hidden="true">✓</span>
                  Entry submitted.
                </div>
              )}

              {scannedMachine && (
                <div className="mb-4 rounded-xl border border-blue-200 bg-blue-50 p-4 text-blue-900">
                  <div className="text-sm font-semibold">Scanned Machine Auto-Fill</div>
                  <div className="mt-1 text-lg font-bold">{scannedMachine.code} - {scannedMachine.name}</div>
                  <div className="mt-1 text-xs text-blue-700">
                    This machine was filled from the QR scan and is locked so the operator cannot accidentally pick the wrong machine.
                  </div>
                </div>
              )}

              {formError && (
                <div className="mb-4 rounded-xl bg-red-100 p-3 text-sm font-medium text-red-800">
                  {formError}
                </div>
              )}

              <form onSubmit={handleSubmit} className="grid gap-4 md:grid-cols-2">
                <label className="space-y-1">
                  <span className="text-sm font-medium">Employee Name</span>
                  <select
                    required
                    value={employeeName}
                    onChange={(e) => setEmployeeName(e.target.value)}
                    className="w-full rounded-xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-900"
                  >
                    <option value="">Select operator</option>
                    {savedOperators.map((name) => (
                      <option key={name} value={name}>
                        {name}
                      </option>
                    ))}
                  </select>
                </label>

                <label className="space-y-1">
                  <span className="text-sm font-medium">Job Number</span>
                  <input
                    required
                    value={jobNumber}
                    onChange={(e) => setJobNumber(e.target.value)}
                    className="w-full rounded-xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-900"
                    placeholder="Enter job number"
                  />
                </label>

                <label className="space-y-1">
                  <span className="text-sm font-medium">DATE</span>
                  <input
                    required
                    type="date"
                    value={date}
                    onChange={(e) => setDate(e.target.value)}
                    className="w-full rounded-xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-900"
                  />
                </label>

                <label className="space-y-1">
                  <span className="text-sm font-medium">Equipment Code</span>
                  <select
                    required
                    value={equipmentCode}
                    disabled={Boolean(scannedMachine)}
                    onChange={(e) => setEquipmentCode(e.target.value)}
                    className="w-full rounded-xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-900 disabled:bg-slate-100 disabled:text-slate-700"
                  >
                    <option value="">Select code</option>
                    {EQUIPMENT_CODES.map((item) => (
                      <option key={item.code} value={item.code}>
                        {item.code} - {item.name}
                      </option>
                    ))}
                  </select>
                </label>

                <label className="space-y-1">
                  <span className="text-sm font-medium">Equipment Hours</span>
                  <div className="rounded-xl border border-slate-200 bg-white p-3">
                    <div className="mb-2 flex items-center justify-between gap-2">
                      <span className="text-xs font-medium text-slate-500">Machine Timer</span>
                      <span className="rounded-lg bg-slate-900 px-3 py-2 font-mono text-lg font-bold tracking-wider text-green-400 shadow-inner">
                        {timerDisplay}
                      </span>
                    </div>
                    <div className="flex gap-2">
                      <Button type="button" variant="outline" onClick={startMachineTimer} disabled={Boolean(timerStartedAt)}>
                        Start
                      </Button>
                      <Button type="button" variant="outline" onClick={stopMachineTimer} disabled={!timerStartedAt}>
                        Stop
                      </Button>
                    </div>
                    <div className="mt-2 rounded-lg bg-slate-50 p-2 text-xs text-slate-600">
                      <div><strong>Calculated Equipment Hours:</strong> {equipmentHours || "0.00"} hr</div>
                      <div className="mt-1">0-10 sec = 0.00, after 10 sec = 0.50, then rounds up in 1/2-hour increments.</div>
                    </div>
                  </div>
                </label>

                <label className="space-y-1 md:col-span-2">
                  <span className="text-sm font-medium">Notes</span>
                  <textarea
                    value={notes}
                    onChange={(e) => setNotes(e.target.value)}
                    className="min-h-24 w-full rounded-xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-900"
                    placeholder="Optional notes"
                  />
                </label>

                <div className="flex flex-wrap items-center gap-3 md:col-span-2">
                  <Button type="submit" className="py-4 text-base" disabled={isSending}>
                    <span className="mr-2" aria-hidden="true">➤</span> {isSending ? "Sending..." : "Submit equipment hours"}
                  </Button>

                  <Button type="button" variant="outline" className="py-4 text-base" onClick={() => resetForm({ keepEquipment: Boolean(scannedMachine) })}>
                    Clear entry
                  </Button>

                  {scannedMachine && (
                    <Button type="button" variant="outline" className="py-4 text-base" onClick={() => resetForm()}>
                      Change machine
                    </Button>
                  )}

                  {selectedEquipment && (
                    <span className="rounded-xl bg-white px-4 py-3 text-sm shadow-sm">
                      Selected: {selectedEquipment.code} - {selectedEquipment.name}
                    </span>
                  )}
                </div>
              </form>
            </CardContent>
          </Card>

          <Card>
            <CardContent className="p-6">
              <div className="mb-4 flex items-center justify-between gap-3 print:hidden">
                <div className="flex items-center gap-3">
                  <IconBadge>▦</IconBadge>
                  <div>
                    <h2 className="text-xl font-bold">{COMPANY_NAME} Machine QR Codes</h2>
                    <p className="text-sm text-slate-600">Each QR opens this app with that machine already selected and locked.</p>
                  </div>
                </div>
                <Button type="button" variant="outline" onClick={printQrCodes}>
                  Print QR Codes
                </Button>
              </div>

              <div className="rounded-2xl bg-slate-100 p-4 shadow-inner print:bg-white print:shadow-none">
                <div className="grid max-h-[420px] gap-3 overflow-auto sm:grid-cols-2 print:max-h-none print:grid-cols-3 print:overflow-visible">
                  {machineQrCodes.map((item) => (
                    <div key={item.code} className="break-inside-avoid rounded-xl bg-white p-3 text-center shadow-sm print:border print:shadow-none">
                      {qrImageFailures[item.code] ? (
                        <div className="mx-auto flex h-36 w-36 items-center justify-center rounded-lg border border-dashed border-slate-300 p-3 text-xs text-slate-600">
                          QR image unavailable. Use link below.
                        </div>
                      ) : (
                        <img
                          src={buildQrImageUrl(item)}
                          alt={`QR code for ${item.code} ${item.name}`}
                          className="mx-auto h-36 w-36 rounded-lg"
                          loading="lazy"
                          onError={() => setQrImageFailures((current) => ({ ...current, [item.code]: true }))}
                        />
                      )}
                      <div className="mt-2 text-sm font-bold">{item.code}</div>
                      <div className="text-xs text-slate-600">{item.name}</div>
                      <a className="mt-2 block break-all text-[10px] text-slate-500" href={buildMachineUrl(item)}>
                        {buildMachineUrl(item)}
                      </a>
                    </div>
                  ))}
                </div>
              </div>

              <div className="mt-4 space-y-3 text-sm text-slate-700 print:hidden">
                <p><strong>Single QR per machine:</strong> duplicate Knee Mill QR codes were merged into one machine entry.</p>
                <p><strong>Spreadsheet match:</strong> Job Number, DATE, Equipment Code, Equipment Hours, Notes.</p>
                <p><strong>Machine prefill:</strong> every QR includes the machine code in the URL, so scanning it auto-selects and locks that machine in the form.</p>
                <p><strong>Excel delivery:</strong> when the webhook is configured, submissions append directly to your Excel table.</p>
              </div>

              <div className="mt-5 grid grid-cols-2 gap-3 print:hidden">
                <div className="rounded-xl bg-white p-4 shadow-sm">
                  <div className="text-2xl font-bold">{rows.length}</div>
                  <div className="text-xs text-slate-500">Entries</div>
                </div>
                <div className="rounded-xl bg-white p-4 shadow-sm">
                  <div className="text-2xl font-bold">{totalHours.toFixed(2)}</div>
                  <div className="text-xs text-slate-500">Total Hours</div>
                </div>
              </div>

              <Button onClick={downloadCsv} variant="outline" className="mt-5 w-full py-4 print:hidden">
                <span className="mr-2" aria-hidden="true">⬇</span> Export CSV
              </Button>
            </CardContent>
          </Card>
        </div>

        <div className="grid gap-4 lg:grid-cols-[1fr_.7fr] print:hidden">
          <Card>
            <CardContent className="p-6">
              <div className="mb-4 flex items-center justify-between gap-3">
                <div className="flex items-center gap-2">
                  <span className="text-xl" aria-hidden="true">▤</span>
                  <h2 className="text-xl font-bold">{COMPANY_NAME} Digital Excel Preview</h2>
                </div>
                <div className="flex items-center gap-2 text-sm text-slate-600">
                  <span aria-hidden="true">✉</span> Sends to Excel after submit
                </div>
              </div>

              <div className="overflow-x-auto rounded-xl border border-slate-200 bg-white">
                <table className="w-full min-w-[950px] text-left text-sm">
                  <thead className="bg-slate-100 text-slate-700">
                    <tr>
                      <th className="p-3">Employee</th>
                      <th className="p-3">Job Number</th>
                      <th className="p-3">DATE</th>
                      <th className="p-3">Equipment Code</th>
                      <th className="p-3">Equipment</th>
                      <th className="p-3">Equipment Hours</th>
                      <th className="p-3">Notes</th>
                      <th className="p-3">Submitted</th>
                      <th className="p-3">Excel Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((row) => (
                      <tr key={row.id} className="border-t border-slate-100">
                        <td className="p-3">{row.employeeName}</td>
                        <td className="p-3 font-medium">{row.jobNumber}</td>
                        <td className="p-3">{row.date}</td>
                        <td className="p-3">{row.equipmentCode}</td>
                        <td className="p-3">{row.equipmentName}</td>
                        <td className="p-3">{row.equipmentHours}</td>
                        <td className="p-3">{row.notes}</td>
                        <td className="p-3 text-slate-500">{row.submittedAt}</td>
                        <td className="p-3 font-medium">{row.excelStatus}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardContent className="p-6">
              <div className="mb-4 flex items-center gap-2">
                <span className="text-xl" aria-hidden="true">☑</span>
                <h2 className="text-xl font-bold">Equipment Codes</h2>
              </div>
              <div className="max-h-[360px] overflow-auto rounded-xl border border-slate-200 bg-white">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-slate-100 text-left text-slate-700">
                    <tr>
                      <th className="p-3">Code</th>
                      <th className="p-3">Equipment</th>
                    </tr>
                  </thead>
                  <tbody>
                    {EQUIPMENT_CODES.map((item) => (
                      <tr key={item.code} className="border-t border-slate-100">
                        <td className="p-3 font-medium">{item.code}</td>
                        <td className="p-3">{item.name}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
}
