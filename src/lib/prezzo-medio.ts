import * as XLSX from "xlsx";

export type Riepilogo = {
  puntoVendita: string;
  movimenti: number;
  incassato: number;
  venduto: number;
  erogazioni: number;
  prezzoMedioIncassato: number | null;
  prezzoMedioVenduto: number | null;
  dataDa: Date | null;
  dataA: Date | null;
};

const COL = { data: 1, pv: 3, incassato: 12, venduto: 17, erogazioni: 26 };

function toNumber(value: unknown): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value !== "string") return 0;
  const cleaned = value
    .replace(/[^\d,.-]/g, "")
    .replace(/\.(?=\d{3}\b)/g, "")
    .replace(",", ".");
  const n = parseFloat(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function toDate(value: unknown): Date | null {
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  if (typeof value === "string") {
    const m = value.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})/);
    if (m) {
      const year = Number(m[3]!.length === 2 ? `20${m[3]}` : m[3]);
      return new Date(year, Number(m[2]) - 1, Number(m[1]));
    }
    const d = new Date(value);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  return null;
}

export function elaboraWorkbook(data: ArrayBuffer): Riepilogo[] {
  const wb = XLSX.read(data, { cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]!]!;
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    blankrows: false,
    defval: null,
  });

  const map = new Map<string, Riepilogo>();

  for (const row of rows) {
    const pv = row[COL.pv];
    if (typeof pv !== "string" || !pv.trim()) continue;
    const erogazioni = toNumber(row[COL.erogazioni]);
    const incassato = toNumber(row[COL.incassato]);
    const venduto = toNumber(row[COL.venduto]);
    if (!erogazioni && !incassato && !venduto) continue;

    const key = pv.trim();
    const entry =
      map.get(key) ??
      ({
        puntoVendita: key,
        movimenti: 0,
        incassato: 0,
        venduto: 0,
        erogazioni: 0,
        prezzoMedioIncassato: null,
        prezzoMedioVenduto: null,
        dataDa: null,
        dataA: null,
      } satisfies Riepilogo);

    entry.movimenti += 1;
    entry.incassato += incassato;
    entry.venduto += venduto;
    entry.erogazioni += erogazioni;

    const d = toDate(row[COL.data]);
    if (d) {
      if (!entry.dataDa || d < entry.dataDa) entry.dataDa = d;
      if (!entry.dataA || d > entry.dataA) entry.dataA = d;
    }

    map.set(key, entry);
  }

  const result = [...map.values()].map((e) => ({
    ...e,
    prezzoMedioIncassato: e.erogazioni ? e.incassato / e.erogazioni : null,
    prezzoMedioVenduto: e.erogazioni ? e.venduto / e.erogazioni : null,
  }));

  return result.sort((a, b) => a.puntoVendita.localeCompare(b.puntoVendita, "it"));
}

export function esportaExcel(righe: Riepilogo[], nomeFile = "prezzo-medio.xlsx") {
  const fmtDate = (d: Date | null) => (d ? d.toLocaleDateString("it-IT") : "");
  const aoa = [
    [
      "Punto vendita",
      "Movimenti",
      "Dal",
      "Al",
      "Incassato",
      "Venduto",
      "Erogazioni",
      "Prezzo medio incassato",
      "Prezzo medio venduto",
      "Scostamento",
    ],
    ...righe.map((r) => [
      r.puntoVendita,
      r.movimenti,
      fmtDate(r.dataDa),
      fmtDate(r.dataA),
      r.incassato,
      r.venduto,
      r.erogazioni,
      r.prezzoMedioIncassato,
      r.prezzoMedioVenduto,
      r.prezzoMedioIncassato !== null && r.prezzoMedioVenduto !== null
        ? r.prezzoMedioIncassato - r.prezzoMedioVenduto
        : null,
    ]),
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [
    { wch: 34 },
    { wch: 11 },
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 22 },
    { wch: 22 },
    { wch: 14 },
  ];
  for (let r = 1; r <= righe.length; r++) {
    for (const c of [4, 5, 6]) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell) cell.z = "#,##0.00";
    }
    for (const c of [7, 8, 9]) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell) cell.z = "0.0000";
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Prezzo medio");
  XLSX.writeFile(wb, nomeFile);
}
