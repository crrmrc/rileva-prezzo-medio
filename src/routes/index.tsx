import { createFileRoute } from "@tanstack/react-router";
import { useMemo, useRef, useState } from "react";
import { Download, FileSpreadsheet, Fuel, Loader2, Upload } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { elaboraWorkbook, esportaExcel, type Riepilogo } from "@/lib/prezzo-medio";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Prezzo medio per punto vendita | Analisi erogazioni" },
      {
        name: "description",
        content:
          "Carica il file Excel dei movimenti e ottieni per ogni punto vendita il prezzo medio da incassato/erogazioni e venduto/erogazioni, con export in Excel.",
      },
      { property: "og:title", content: "Prezzo medio per punto vendita" },
      {
        property: "og:description",
        content:
          "Calcolo automatico del prezzo medio per punto vendita da incassato, venduto ed erogazioni, con export Excel.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

const num = (v: number, dec = 2) =>
  v.toLocaleString("it-IT", { minimumFractionDigits: dec, maximumFractionDigits: dec });
const prezzo = (v: number | null) => (v === null ? "—" : `€ ${num(v, 4)}`);
const data = (d: Date | null) => (d ? d.toLocaleDateString("it-IT") : "—");

function Index() {
  const [righe, setRighe] = useState<Riepilogo[] | null>(null);
  const [nomeFile, setNomeFile] = useState("");
  const [errore, setErrore] = useState<string | null>(null);
  const [caricamento, setCaricamento] = useState(false);
  const [filtro, setFiltro] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  const totali = useMemo(() => {
    if (!righe) return null;
    const t = righe.reduce(
      (acc, r) => ({
        incassato: acc.incassato + r.incassato,
        venduto: acc.venduto + r.venduto,
        erogazioni: acc.erogazioni + r.erogazioni,
        movimenti: acc.movimenti + r.movimenti,
      }),
      { incassato: 0, venduto: 0, erogazioni: 0, movimenti: 0 },
    );
    return {
      ...t,
      medioIncassato: t.erogazioni ? t.incassato / t.erogazioni : null,
      medioVenduto: t.erogazioni ? t.venduto / t.erogazioni : null,
    };
  }, [righe]);

  const visibili = useMemo(() => {
    if (!righe) return [];
    const q = filtro.trim().toLowerCase();
    return q ? righe.filter((r) => r.puntoVendita.toLowerCase().includes(q)) : righe;
  }, [righe, filtro]);

  const parziali = useMemo(() => {
    const t = visibili.reduce(
      (acc, r) => ({
        incassato: acc.incassato + r.incassato,
        venduto: acc.venduto + r.venduto,
        erogazioni: acc.erogazioni + r.erogazioni,
        movimenti: acc.movimenti + r.movimenti,
      }),
      { incassato: 0, venduto: 0, erogazioni: 0, movimenti: 0 },
    );
    return {
      ...t,
      medioIncassato: t.erogazioni ? t.incassato / t.erogazioni : null,
      medioVenduto: t.erogazioni ? t.venduto / t.erogazioni : null,
    };
  }, [visibili]);

  const filtroAttivo = filtro.trim().length > 0;

  async function onFile(file: File) {
    setCaricamento(true);
    setErrore(null);
    try {
      const buffer = await file.arrayBuffer();
      const result = elaboraWorkbook(buffer);
      if (!result.length) {
        setRighe(null);
        setErrore(
          "Nessun movimento riconosciuto. Verifica che i dati siano nelle colonne B, D, M, R e AA del primo foglio.",
        );
      } else {
        setRighe(result);
        setNomeFile(file.name);
      }
    } catch {
      setRighe(null);
      setErrore("Non è stato possibile leggere il file. Formati supportati: .xlsx, .xls, .csv");
    } finally {
      setCaricamento(false);
    }
  }

  return (
    <main className="min-h-screen bg-background">
      <header
        className="border-b border-border/40 text-primary-foreground"
        style={{ background: "var(--surface-gradient)" }}
      >
        <div className="mx-auto flex max-w-6xl flex-col gap-3 px-6 py-10">
          <div className="flex items-center gap-3">
            <span className="flex size-11 items-center justify-center rounded-lg bg-white/15">
              <Fuel className="size-6" />
            </span>
            <div>
              <h1 className="text-2xl font-semibold tracking-tight">
                Prezzo medio per punto vendita
              </h1>
              <p className="text-sm opacity-80">
                Incassato / erogazioni e venduto / erogazioni, aggregati sul mese
              </p>
            </div>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-6xl space-y-8 px-6 py-10">
        <Card>
          <CardHeader>
            <CardTitle className="text-base">1. Carica il file dei movimenti</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            <div
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                const f = e.dataTransfer.files?.[0];
                if (f) void onFile(f);
              }}
              className="flex flex-col items-center gap-3 rounded-lg border border-dashed border-border bg-secondary/40 px-6 py-10 text-center"
            >
              <FileSpreadsheet className="size-8 text-muted-foreground" />
              <p className="text-sm text-muted-foreground">
                Trascina qui il file Excel oppure selezionalo dal computer
              </p>
              <Button onClick={() => inputRef.current?.click()} disabled={caricamento}>
                {caricamento ? (
                  <Loader2 className="animate-spin" />
                ) : (
                  <Upload />
                )}
                Seleziona file
              </Button>
              <input
                ref={inputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) void onFile(f);
                  e.target.value = "";
                }}
              />
              {nomeFile && !errore ? (
                <p className="text-xs text-muted-foreground">File caricato: {nomeFile}</p>
              ) : null}
            </div>
            {errore ? <p className="text-sm text-destructive">{errore}</p> : null}
            <p className="text-xs text-muted-foreground">
              Colonne lette: B data · D punto vendita · M incassato · R venduto · AA erogazioni.
            </p>
          </CardContent>
        </Card>

        {righe && totali ? (
          <>
            <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
              {[
                { label: "Punti vendita", value: String(righe.length) },
                { label: "Movimenti", value: num(totali.movimenti, 0) },
                {
                  label: "Prezzo medio incassato",
                  value: prezzo(totali.medioIncassato),
                  evidenza: true,
                },
                {
                  label: "Prezzo medio venduto",
                  value: prezzo(totali.medioVenduto),
                  evidenza: true,
                },
              ].map((k) => (
                <Card key={k.label}>
                  <CardContent className="pt-6">
                    <p className="text-xs uppercase tracking-wide text-muted-foreground">
                      {k.label}
                    </p>
                    <p
                      className={`mt-2 text-2xl font-semibold ${
                        k.evidenza ? "text-primary" : "text-foreground"
                      }`}
                    >
                      {k.value}
                    </p>
                  </CardContent>
                </Card>
              ))}
            </div>

            <Card>
              <CardHeader className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                <CardTitle className="text-base">2. Riepilogo per punto vendita</CardTitle>
                <div className="flex flex-col gap-2 sm:flex-row">
                  <Input
                    value={filtro}
                    onChange={(e) => setFiltro(e.target.value)}
                    placeholder="Cerca punto vendita"
                    className="sm:w-56"
                  />
                  <Button variant="secondary" onClick={() => esportaExcel(righe)}>
                    <Download />
                    Esporta in Excel
                  </Button>
                </div>
              </CardHeader>
              <CardContent>
                <div className="overflow-x-auto">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>Punto vendita</TableHead>
                        <TableHead className="text-right">Mov.</TableHead>
                        <TableHead>Periodo</TableHead>
                        <TableHead className="text-right">Incassato</TableHead>
                        <TableHead className="text-right">Venduto</TableHead>
                        <TableHead className="text-right">Erogazioni</TableHead>
                        <TableHead className="text-right">P. medio incassato</TableHead>
                        <TableHead className="text-right">P. medio venduto</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {visibili.map((r) => (
                        <TableRow key={r.puntoVendita}>
                          <TableCell className="font-medium">{r.puntoVendita}</TableCell>
                          <TableCell className="text-right">{r.movimenti}</TableCell>
                          <TableCell className="whitespace-nowrap text-muted-foreground">
                            {data(r.dataDa)} – {data(r.dataA)}
                          </TableCell>
                          <TableCell className="text-right">€ {num(r.incassato)}</TableCell>
                          <TableCell className="text-right">€ {num(r.venduto)}</TableCell>
                          <TableCell className="text-right">{num(r.erogazioni, 3)}</TableCell>
                          <TableCell className="text-right font-semibold text-primary">
                            {prezzo(r.prezzoMedioIncassato)}
                          </TableCell>
                          <TableCell className="text-right font-semibold">
                            {prezzo(r.prezzoMedioVenduto)}
                          </TableCell>
                        </TableRow>
                      ))}
                      {filtroAttivo && visibili.length < (righe?.length ?? 0) ? (
                        <TableRow className="bg-primary/5">
                          <TableCell className="font-semibold text-primary">
                            Parziale ({visibili.length} di {righe?.length})
                          </TableCell>
                          <TableCell className="text-right font-semibold">
                            {parziali.movimenti}
                          </TableCell>
                          <TableCell />
                          <TableCell className="text-right font-semibold">
                            € {num(parziali.incassato)}
                          </TableCell>
                          <TableCell className="text-right font-semibold">
                            € {num(parziali.venduto)}
                          </TableCell>
                          <TableCell className="text-right font-semibold">
                            {num(parziali.erogazioni, 3)}
                          </TableCell>
                          <TableCell className="text-right font-semibold">
                            {prezzo(parziali.medioIncassato)}
                          </TableCell>
                          <TableCell className="text-right font-semibold">
                            {prezzo(parziali.medioVenduto)}
                          </TableCell>
                        </TableRow>
                      ) : null}
                      <TableRow className="bg-secondary/60">
                        <TableCell className="font-semibold">Totale</TableCell>
                        <TableCell className="text-right font-semibold">
                          {totali.movimenti}
                        </TableCell>
                        <TableCell />
                        <TableCell className="text-right font-semibold">
                          € {num(totali.incassato)}
                        </TableCell>
                        <TableCell className="text-right font-semibold">
                          € {num(totali.venduto)}
                        </TableCell>
                        <TableCell className="text-right font-semibold">
                          {num(totali.erogazioni, 3)}
                        </TableCell>
                        <TableCell className="text-right font-semibold">
                          {prezzo(totali.medioIncassato)}
                        </TableCell>
                        <TableCell className="text-right font-semibold">
                          {prezzo(totali.medioVenduto)}
                        </TableCell>
                      </TableRow>
                    </TableBody>
                  </Table>
                </div>
              </CardContent>
            </Card>
          </>
        ) : null}
      </div>
    </main>
  );
}
