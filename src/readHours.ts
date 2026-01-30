import * as XLSX from "xlsx";

type WorklogRow = {
    workItemId: number;
    date: string; // dd/mm/yyyy
    hours: number;
    minutes: number;
};

type PunchLine = {
    date: string; // dd/mm/yyyy
    time: string; // HH:mm
    type: "ENTRY" | "EXIT";
};

export function buildPunchLinesFromXlsx(
    filePath: string,
    defaultEntryHHmm: string = "09:00",
    startDate?: string // dd/mm/yyyy (inclusive)
): PunchLine[] {
    const rows = readWorklogFromExcel(filePath);

    // 1) agrega minutos por dia
    const totalMinutesByDay = aggregateMinutesByDay(rows);

    // 2) ordena datas (dd/mm/yyyy)
    let daysSorted = Object.keys(totalMinutesByDay).sort(
        (a, b) => toDateKey(a) - toDateKey(b)
    );

    if (startDate) {
        const startKey = toDateKey(startDate);
        daysSorted = daysSorted.filter((d) => toDateKey(d) >= startKey);
    }

    // 3) gera lista sequencial: entrada + saida
    const punches: PunchLine[] = [];
    for (const day of daysSorted) {
        const totalMinutes = totalMinutesByDay[day];
        const entry = defaultEntryHHmm;
        const exit = addMinutesToHHmm(entry, totalMinutes);

        punches.push({ date: day, time: entry, type: "ENTRY" });
        punches.push({ date: day, time: exit, type: "EXIT" });
    }

    return punches;
}

/** ---------------- Excel read ---------------- */
function readWorklogFromExcel(filePath: string): WorklogRow[] {
    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const raw = XLSX.utils.sheet_to_json<any>(sheet, { defval: "" });

    return raw
        .map((r) => {
            const row = normalizeRowKeys(r);
            const date = excelSerialToDDMMYYYY(row["Data da Tarefa"]); // <- aqui resolve seu caso

            return {
                workItemId: Number(row["WorkItemId"] ?? 0),
                date,
                hours: Number(row["CampoHora"] ?? 0),
                minutes: Number(row["campoMinuto"] ?? 0),
            };
        })
        .filter((r) => r.date !== "");
}

/** Converte serial Excel (ex.: 46052) -> "dd/mm/yyyy" */
function excelSerialToDDMMYYYY(serial: any): string {
    if (serial instanceof Date && !isNaN(serial.getTime())) {
        const dd = String(serial.getDate()).padStart(2, "0");
        const mm = String(serial.getMonth() + 1).padStart(2, "0");
        const yyyy = String(serial.getFullYear());
        return `${dd}/${mm}/${yyyy}`;
    }

    if (typeof serial === "string") {
        const trimmed = serial.trim();
        if (trimmed === "") return "";

        // dd/mm/yyyy
        const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(trimmed);
        if (m) {
            const dd = m[1].padStart(2, "0");
            const mm = m[2].padStart(2, "0");
            const yyyy = m[3];
            return `${dd}/${mm}/${yyyy}`;
        }

        // numeric in string
        const asNumber = Number(trimmed);
        if (!isNaN(asNumber)) serial = asNumber;
    }

    if (typeof serial !== "number" || !isFinite(serial)) return "";

    const parsed = XLSX.SSF.parse_date_code(serial);
    if (!parsed || !parsed.y || !parsed.m || !parsed.d) return "";

    const dd = String(parsed.d).padStart(2, "0");
    const mm = String(parsed.m).padStart(2, "0");
    const yyyy = String(parsed.y);
    return `${dd}/${mm}/${yyyy}`;
}

function normalizeRowKeys(row: Record<string, any>): Record<string, any> {
    const normalized: Record<string, any> = {};
    for (const key of Object.keys(row)) {
        const trimmedKey = key.trim();
        normalized[trimmedKey] = row[key];
    }
    return normalized;
}

/** ---------------- Aggregate ---------------- */
function aggregateMinutesByDay(rows: WorklogRow[]): Record<string, number> {
    const totals: Record<string, number> = {};
    for (const r of rows) {
        const mins = r.hours * 60 + r.minutes;
        totals[r.date] = (totals[r.date] ?? 0) + mins;
    }
    return totals;
}

/** ---------------- Sort helper (dd/mm/yyyy) ---------------- */
function toDateKey(ddmmyyyy: string): number {
    const [dd, mm, yyyy] = ddmmyyyy.split("/");
    return Number(`${yyyy}${mm}${dd}`); // yyyymmdd
}

/** ---------------- Time helpers ---------------- */
function hhmmToMinutes(hhmm: string): number {
    const m = /^(\d{1,2}):(\d{2})$/.exec(hhmm.trim());
    if (!m) throw new Error(`Invalid time HH:mm: "${hhmm}"`);
    const h = Number(m[1]);
    const min = Number(m[2]);
    return h * 60 + min;
}

function minutesToHHmm(totalMinutes: number): string {
    const h = Math.floor(totalMinutes / 60);
    const m = totalMinutes % 60;
    return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

function addMinutesToHHmm(startHHmm: string, addMinutes: number): string {
    const end = hhmmToMinutes(startHHmm) + addMinutes;
    if (end >= 24 * 60) {
        throw new Error(`Exit time overflow: start=${startHHmm}, add=${addMinutes}min`);
    }
    return minutesToHHmm(end);
}

/** ---------------- Example usage ---------------- */
// const punches = buildPunchLinesFromXlsx(
//     "D:\\git\\laranjaAutomacao\\src\\relatorioApontamento2026-01-30T12_17_55.3025122Z.xlsx",
//     "09:00", "05/01/2026");
//
// console.table(punches);
// punches.forEach(p => console.log(`${p.date} ${p.time} ${p.type}`));
