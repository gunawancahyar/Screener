"use client";

import React, { useState, useMemo, useEffect } from "react";

// ==========================================
// 1. HELPER FUNCTIONS & CONSTANTS (Di luar komponen)
// Taruh zeroNeraca, zeroLabaRugi, computeRatios, dll. di area ini
// ==========================================


const PIE_COLORS = ["#34d399", "#10b981", "#059669", "#047857", "#065f46", "#a3e635"];


export default function ScreenerPage() {
  // 1. State Hydration (isMounted)
  const [isMounted, setIsMounted] = useState(false);

  // 2. State Aplikasi & Input
  const [tab, setTab] = useState("dashboard");
  const [companyName, setCompanyName] = useState("");
  const [neraca, setNeraca] = useState(zeroNeraca);
  const [lr, setLr] = useState(zeroLabaRugi);
  const [pasar, setPasar] = useState({ hargaSaham: 0 });
  const [prev, setPrev] = useState(zeroPrev);
  const [params, setParams] = useState({
    perTarget: 15,
    pbvTarget: 2,
    growth: 0.1,
    discountRate: 0.12,
    terminalGrowth: 0.03,
  });
  const [history, setHistory] = useState([]);

  // 3. Helper Setter
  const setN = (k) => (v) => setNeraca((s) => ({ ...s, [k]: v }));
  const setL = (k) => (v) => setLr((s) => ({ ...s, [k]: v }));
  const setP = (k) => (v) => setPrev((s) => ({ ...s, [k]: v }));

  // 4. Calculations (useMemo)
  const ratios = useMemo(() => computeRatios(neraca, lr, pasar), [neraca, lr, pasar]);
  const scoring = useMemo(() => computeScoring(ratios), [ratios]);
  const valuation = useMemo(() => computeValuation(ratios, lr, neraca, pasar, params), [ratios, lr, neraca, pasar, params]);
  const altman = useMemo(() => computeAltmanZ(neraca, lr, pasar, ratios), [neraca, lr, pasar, ratios]);
  const piotroski = useMemo(() => computePiotroski(neraca, lr, prev, ratios), [neraca, lr, prev, ratios]);
  const cfq = useMemo(() => computeCashFlowQuality(prev, lr), [prev, lr]);
  const redFlags = useMemo(() => detectRedFlags(neraca, lr, prev, ratios), [neraca, lr, prev, ratios]);
  const invScore = useMemo(() => computeInvestmentScore(scoring, valuation, prev, lr, ratios), [scoring, valuation, prev, lr, ratios]);
  const analysisText = useMemo(() => generateAnalysis(scoring, ratios, redFlags), [scoring, ratios, redFlags]);

  const hasData = neraca.totalAset > 0 && lr.penjualan > 0;

  // 5. Hydration Effect
  useEffect(() => {
    setIsMounted(true);
  }, []);

  // 6. Functions
  function saveToHistory() {
    setHistory((h) => [
      {
        id: Date.now(),
        name: companyName || `Analisa #${h.length + 1}`,
        date: new Date().toLocaleString("id-ID"),
        grade: scoring.grade,
        pct: scoring.pct,
        avgFair: valuation.avgFair,
        harga: pasar.hargaSaham,
        verdict: valuation.verdict,
        recommendation: invScore.recommendation,
      },
      ...h,
    ]);
  }

  // D. HYDRATION GUARD
  if (!isMounted) {
    return (
      <div className="min-h-screen bg-zinc-950 text-emerald-400 flex items-center justify-center font-mono">
        Loading Fundamental Screener...
      </div>
    );
  }

  // E. JSX TAMPILAN UTAMA
  return (
    <div className="min-h-screen bg-zinc-950 text-white p-6">
      {/* MASUKKAN SELURUH TAMPILAN JSX DASHBOARD KAMU DI SINI */}
    </div>
  );

  function exportExcel() {
    const wb = XLSX.utils.book_new();
    const inputRows = [
      ["NERACA", ""], ...Object.entries(neraca).map(([k, v]) => [k, v]),
      ["", ""], ["LABA RUGI", ""], ...Object.entries(lr).map(([k, v]) => [k, v]),
      ["", ""], ["DATA PASAR", ""], ["hargaSaham", pasar.hargaSaham],
    ];
    wb.SheetNames.push("Data Input");
    wb.Sheets["Data Input"] = XLSX.utils.aoa_to_sheet(inputRows);

    const ratioRows = [["Kategori", "Rasio", "Nilai", "Penilaian"]];
    scoring.rows.forEach(r => ratioRows.push([r.cat, r.label, fmtVal(r.value, r.unit), r.scored ? r.scored.label : "-"]));
    wb.SheetNames.push("Rasio & Skor");
    wb.Sheets["Rasio & Skor"] = XLSX.utils.aoa_to_sheet(ratioRows);

    const valRows = [["Metode", "Estimasi Harga Wajar"]];
    valuation.methods.forEach(m => valRows.push([m.label, m.value ? Math.round(m.value) : "-"]));
    valRows.push(["Rata-rata Harga Wajar", valuation.avgFair ? Math.round(valuation.avgFair) : "-"]);
    valRows.push(["Harga Saat Ini", pasar.hargaSaham]);
    valRows.push(["Margin of Safety (%)", valuation.mos ? valuation.mos.toFixed(2) : "-"]);
    valRows.push(["Kesimpulan", valuation.verdict || "-"]);
    valRows.push(["Skor Investasi (0-100)", invScore.total.toFixed(1)]);
    valRows.push(["Rekomendasi", invScore.recommendation]);
    wb.SheetNames.push("Valuasi");
    wb.Sheets["Valuasi"] = XLSX.utils.aoa_to_sheet(valRows);

    XLSX.writeFile(wb, `analisa-fundamental-${(companyName || "saham").replace(/\s+/g, "_")}.xlsx`);
  }

  function exportPDF() {
    window.print();
  }

  const radarData = CATEGORIES.map(c => ({
    category: c,
    score: scoring.byCategory[c].max ? (scoring.byCategory[c].sum / scoring.byCategory[c].max) * 100 : 0,
  }));

  const assetPieData = [
    { name: "Kas", value: neraca.kas },
    { name: "Piutang", value: neraca.piutang },
    { name: "Persediaan", value: neraca.persediaan },
    { name: "Aset Tetap", value: neraca.asetTetap },
    { name: "Aset Lainnya", value: Math.max(0, neraca.totalAset - neraca.kas - neraca.piutang - neraca.persediaan - neraca.asetTetap) },
  ].filter(d => d.value > 0);

  const liabEquityPieData = [
    { name: "Liabilitas Lancar", value: neraca.liabilitasLancar },
    { name: "Liabilitas Jk. Panjang", value: neraca.liabilitasJangkaPanjang },
    { name: "Ekuitas", value: neraca.totalEkuitas },
  ].filter(d => d.value > 0);

  const marginBarData = [
    { name: "Gross", value: ratios.grossMargin ? ratios.grossMargin * 100 : 0 },
    { name: "Operating", value: ratios.operatingMargin ? ratios.operatingMargin * 100 : 0 },
    { name: "Net", value: ratios.netMargin ? ratios.netMargin * 100 : 0 },
    { name: "ROA", value: ratios.roa ? ratios.roa * 100 : 0 },
    { name: "ROE", value: ratios.roe ? ratios.roe * 100 : 0 },
  ];

  const trendData = prev.penjualanTahunLalu || prev.labaBersihTahunLalu ? [
    { name: "Tahun Lalu", Penjualan: prev.penjualanTahunLalu, LabaBersih: prev.labaBersihTahunLalu },
    { name: "Tahun Ini", Penjualan: lr.penjualan, LabaBersih: lr.labaBersih },
  ] : [];

  const scoreDonutData = CATEGORIES.map(c => ({ name: c, value: scoring.byCategory[c].sum })).filter(d => d.value > 0);
  
  const gaugeColor = scoring.grade === "A+" || scoring.grade === "A" ? "#34d399" : scoring.grade === "B" ? "#a3e635" : scoring.grade === "C" ? "#facc15" : "#f87171";

  return (
    <div className="min-h-screen bg-black text-gray-100 font-sans">
      <style>{`
        @media print {
          nav, .no-print { display: none !important; }
          body { background: white !important; }
        }
        input[type=number]::-webkit-inner-spin-button { opacity: 1; }
      `}</style>

      {/* NAVBAR */}
      <nav className="sticky top-0 z-20 border-b border-emerald-500/20 bg-black/90 backdrop-blur">
        <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between gap-4 flex-wrap">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-emerald-500/10 border border-emerald-400/40 flex items-center justify-center">
              <span className="text-emerald-400 font-mono font-bold text-sm">FS</span>
            </div>
            <div>
              <div className="text-sm font-bold tracking-widest text-emerald-400 uppercase leading-none">Fundamental Screener</div>
              <div className="text-[10px] text-gray-500 leading-none mt-0.5">Analisa Saham Indonesia</div>
            </div>
          </div>
          <div className="flex items-center gap-1 overflow-x-auto">
            {NAV.map(n => (
              <button
                key={n.id}
                onClick={() => setTab(n.id)}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium whitespace-nowrap transition-colors ${
                  tab === n.id ? "bg-emerald-500/15 text-emerald-300 border border-emerald-400/40" : "text-gray-400 hover:text-emerald-300 border border-transparent"
                }`}
              >
                <n.icon className="w-3.5 h-3.5" /> {n.label}
              </button>
            ))}
          </div>
          <div className="flex items-center gap-2 no-print">
            <button onClick={exportPDF} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium border border-emerald-500/30 text-gray-300 hover:bg-emerald-500/10">
              <FileDown className="w-3.5 h-3.5" /> PDF
            </button>
            <button onClick={exportExcel} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium border border-emerald-500/30 text-gray-300 hover:bg-emerald-500/10">
              <FileSpreadsheet className="w-3.5 h-3.5" /> Excel
            </button>
          </div>
        </div>
      </nav>
    

      <main className="max-w-7xl mx-auto px-4 py-6" id="print-report">
        {/* DASHBOARD */}
        {tab === "dashboard" && (
          <div className="space-y-5">
            <div className="flex items-center justify-between flex-wrap gap-3">
              <div>
                <h1 className="text-xl font-bold text-white">Ringkasan Analisa</h1>
                <p className="text-sm text-gray-500">{companyName || "Belum ada nama emiten — isi di tab Input Data"}</p>
              </div>
              {hasData && (
                <button onClick={saveToHistory} className="no-print flex items-center gap-1.5 px-3 py-2 rounded-md text-xs font-medium bg-emerald-500 text-black hover:bg-emerald-400">
                  <Save className="w-3.5 h-3.5" /> Simpan ke Riwayat
                </button>
              )}
            </div>

            {!hasData ? (
              <Card className="p-10 text-center">
                <Info className="w-8 h-8 text-emerald-400 mx-auto mb-3" />
                <p className="text-gray-300 font-medium">Belum ada data untuk dianalisa.</p>
                <p className="text-sm text-gray-500 mt-1">Isi Neraca dan Laba Rugi pada tab "Input Data" untuk mulai.</p>
                <button onClick={() => setTab("input")} className="mt-4 inline-flex items-center gap-1 px-4 py-2 rounded-md bg-emerald-500 text-black text-sm font-semibold hover:bg-emerald-400">
                  Mulai Input Data <ChevronRight className="w-4 h-4" />
                </button>
              </Card>
            ) : (
              <>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                  <StatCard label="Skor Kesehatan" value={scoring.grade} sub={`${scoring.pct.toFixed(1)} / 100`} color={gaugeColor} />
                  <StatCard label="Harga Wajar (rata-rata)" value={fmtIDR(valuation.avgFair)} sub={valuation.verdict || "-"} color="#34d399" />
                  <StatCard label="Margin of Safety" value={valuation.mos !== null ? `${valuation.mos.toFixed(1)}%` : "-"} sub={valuation.mosLabel || "-"} color="#34d399" />
                  <StatCard label="Skor Investasi" value={invScore.total.toFixed(0)} sub={invScore.recommendation} color="#34d399" />
                </div>

                <div className="grid md:grid-cols-2 gap-4">
                  <Card>
                    <CardHeader title="Radar Rasio per Kategori" icon={BarChart3} />
                    <div className="p-4 h-72">
                      <ResponsiveContainer width="100%" height="100%">
                        <RadarChart data={radarData}>
                          <PolarGrid stroke="#065f46" />
                          <PolarAngleAxis dataKey="category" tick={{ fill: "#9ca3af", fontSize: 11 }} />
                          <PolarRadiusAxis domain={[0, 100]} tick={{ fill: "#4b5563", fontSize: 9 }} />
                          <Radar dataKey="score" stroke="#34d399" fill="#34d399" fillOpacity={0.35} />
                        </RadarChart>
                      </ResponsiveContainer>
                    </div>
                  </Card>

                  <Card>
                    <CardHeader title="Gauge Health Score" icon={Calculator} />
                    <div className="p-4 h-72 flex flex-col items-center justify-center">
                      <ResponsiveContainer width="100%" height="80%">
                        <RadialBarChart innerRadius="60%" outerRadius="100%" data={[{ name: "score", value: scoring.pct, fill: gaugeColor }]} startAngle={180} endAngle={0}>
                          <RadialBar dataKey="value" background={{ fill: "#1f2937" }} cornerRadius={8} />
                        </RadialBarChart>
                      </ResponsiveContainer>
                      <div className="text-center -mt-16">
                        <div className="text-3xl font-bold font-mono" style={{ color: gaugeColor }}>{scoring.grade}</div>
                        <div className="text-xs text-gray-500">{scoring.pct.toFixed(1)} / 100</div>
                      </div>
                    </div>
                  </Card>
                </div>

                <Card>
                  <CardHeader title="Analisa Otomatis" icon={Info} />
                  <div className="p-5 space-y-2">
                    {analysisText.map((t, i) => (
                      <p key={i} className="text-sm text-gray-300 leading-relaxed flex gap-2">
                        <span className="text-emerald-400">•</span> {t}
                      </p>
                    ))}
                  </div>
                </Card>
              </>
            )}
          </div>
        )}

        {/* INPUT DATA */}
        {tab === "input" && (
          <div className="space-y-5">
            <h1 className="text-xl font-bold text-white">Input Data Laporan Keuangan</h1>

            <Card>
              <CardHeader title="Identitas Emiten" icon={ClipboardEdit} />
              <div className="p-5">
                <NumberFieldText label="Nama Emiten / Kode Saham" value={companyName} onChange={setCompanyName} />
              </div>
            </Card>

            <Card>
              <CardHeader title="Neraca" icon={ClipboardEdit} />
              <div className="p-5 grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
                <NumberField label="Kas dan Setara Kas" value={neraca.kas} onChange={setN("kas")} suffix="IDR" />
                <NumberField label="Piutang" value={neraca.piutang} onChange={setN("piutang")} suffix="IDR" />
                <NumberField label="Persediaan" value={neraca.persediaan} onChange={setN("persediaan")} suffix="IDR" />
                <NumberField label="Aset Lancar" value={neraca.asetLancar} onChange={setN("asetLancar")} suffix="IDR" />
                <NumberField label="Aset Tetap" value={neraca.asetTetap} onChange={setN("asetTetap")} suffix="IDR" />
                <NumberField label="Total Aset" value={neraca.totalAset} onChange={setN("totalAset")} suffix="IDR" />
                <NumberField label="Liabilitas Lancar" value={neraca.liabilitasLancar} onChange={setN("liabilitasLancar")} suffix="IDR" />
                <NumberField label="Liabilitas Jangka Panjang" value={neraca.liabilitasJangkaPanjang} onChange={setN("liabilitasJangkaPanjang")} suffix="IDR" />
                <NumberField label="Total Liabilitas" value={neraca.totalLiabilitas} onChange={setN("totalLiabilitas")} suffix="IDR" />
                <NumberField label="Total Ekuitas" value={neraca.totalEkuitas} onChange={setN("totalEkuitas")} suffix="IDR" />
                <NumberField label="Jumlah Saham Beredar" value={neraca.jumlahSaham} onChange={setN("jumlahSaham")} suffix="lembar" />
              </div>
            </Card>

            <Card>
              <CardHeader title="Laporan Laba Rugi" icon={ClipboardEdit} />
              <div className="p-5 grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
                <NumberField label="Penjualan" value={lr.penjualan} onChange={setL("penjualan")} suffix="IDR" />
                <NumberField label="Harga Pokok Penjualan" value={lr.hpp} onChange={setL("hpp")} suffix="IDR" />
                <NumberField label="Laba Kotor" value={lr.labaKotor} onChange={setL("labaKotor")} suffix="IDR" />
                <NumberField label="EBIT" value={lr.ebit} onChange={setL("ebit")} suffix="IDR" />
                <NumberField label="EBITDA" value={lr.ebitda} onChange={setL("ebitda")} suffix="IDR" />
                <NumberField label="Beban Bunga" value={lr.bebanBunga} onChange={setL("bebanBunga")} suffix="IDR" />
                <NumberField label="Laba Bersih" value={lr.labaBersih} onChange={setL("labaBersih")} suffix="IDR" />
                <NumberField label="Dividen" value={lr.dividen} onChange={setL("dividen")} suffix="IDR" />
              </div>
            </Card>

            <Card>
              <CardHeader title="Data Pasar" icon={ClipboardEdit} />
              <div className="p-5 grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
                <NumberField label="Harga Saham Saat Ini" value={pasar.hargaSaham} onChange={v => setPasar({ hargaSaham: v })} suffix="IDR" />
              </div>
            </Card>

            <Card>
              <CardHeader title="Data Tambahan (Opsional)" icon={ClipboardEdit}
                right={<span className="text-[10px] text-gray-500">Untuk tren, red flags, Altman Z-Score & Piotroski F-Score</span>} />
              <div className="p-5 grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
                <NumberField label="Laba Bersih Tahun Lalu" value={prev.labaBersihTahunLalu} onChange={setP("labaBersihTahunLalu")} suffix="IDR" />
                <NumberField label="Penjualan Tahun Lalu" value={prev.penjualanTahunLalu} onChange={setP("penjualanTahunLalu")} suffix="IDR" />
                <NumberField label="Laba Kotor Tahun Lalu" value={prev.labaKotorTahunLalu} onChange={setP("labaKotorTahunLalu")} suffix="IDR" />
                <NumberField label="Piutang Tahun Lalu" value={prev.piutangTahunLalu} onChange={setP("piutangTahunLalu")} suffix="IDR" />
                <NumberField label="Persediaan Tahun Lalu" value={prev.persediaanTahunLalu} onChange={setP("persediaanTahunLalu")} suffix="IDR" />
                <NumberField label="Total Aset Tahun Lalu" value={prev.totalAsetTahunLalu} onChange={setP("totalAsetTahunLalu")} suffix="IDR" />
                <NumberField label="Total Liabilitas Tahun Lalu" value={prev.totalLiabilitasTahunLalu} onChange={setP("totalLiabilitasTahunLalu")} suffix="IDR" />
                <NumberField label="Aset Lancar Tahun Lalu" value={prev.asetLancarTahunLalu} onChange={setP("asetLancarTahunLalu")} suffix="IDR" />
                <NumberField label="Liabilitas Lancar Tahun Lalu" value={prev.liabilitasLancarTahunLalu} onChange={setP("liabilitasLancarTahunLalu")} suffix="IDR" />
                <NumberField label="Jumlah Saham Tahun Lalu" value={prev.jumlahSahamTahunLalu} onChange={setP("jumlahSahamTahunLalu")} suffix="lembar" />
                <NumberField label="Arus Kas Operasi (CFO)" value={prev.arusKasOperasi} onChange={setP("arusKasOperasi")} suffix="IDR" hint="isi periode berjalan" />
              </div>
            </Card>
          </div>
        )}

        {/* HASIL ANALISA */}
        {tab === "hasil" && (
          <div className="space-y-5">
            <h1 className="text-xl font-bold text-white">Hasil Analisa Rasio</h1>

            <div className="grid md:grid-cols-2 gap-4">
              <Card>
                <CardHeader title="Komposisi Aset" icon={BarChart3} />
                <div className="p-4 h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={assetPieData} dataKey="value" nameKey="name" outerRadius={80} label={{ fill: "#9ca3af", fontSize: 10 }}>
                        {assetPieData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                      </Pie>
                      <Tooltip formatter={v => fmtIDR(v)} contentStyle={{ background: "#111827", border: "1px solid #10b981" }} />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </Card>
              <Card>
                <CardHeader title="Komposisi Liabilitas & Ekuitas" icon={BarChart3} />
                <div className="p-4 h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={liabEquityPieData} dataKey="value" nameKey="name" outerRadius={80} label={{ fill: "#9ca3af", fontSize: 10 }}>
                        {liabEquityPieData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                      </Pie>
                      <Tooltip formatter={v => fmtIDR(v)} contentStyle={{ background: "#111827", border: "1px solid #10b981" }} />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </Card>
            </div>

            <div className="grid md:grid-cols-2 gap-4">
              <Card>
                <CardHeader title="Bar Chart Profitabilitas (%)" icon={BarChart3} />
                <div className="p-4 h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={marginBarData}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" />
                      <XAxis dataKey="name" tick={{ fill: "#9ca3af", fontSize: 11 }} />
                      <YAxis tick={{ fill: "#9ca3af", fontSize: 11 }} />
                      <Tooltip contentStyle={{ background: "#111827", border: "1px solid #10b981" }} formatter={v => v.toFixed(2) + "%"} />
                      <Bar dataKey="value" fill="#34d399" radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </Card>
              <Card>
                <CardHeader title="Tren Penjualan & Laba" icon={BarChart3} />
                <div className="p-4 h-64">
                  {trendData.length ? (
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={trendData}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" />
                        <XAxis dataKey="name" tick={{ fill: "#9ca3af", fontSize: 11 }} />
                        <YAxis tick={{ fill: "#9ca3af", fontSize: 11 }} />
                        <Tooltip contentStyle={{ background: "#111827", border: "1px solid #10b981" }} formatter={v => fmtIDR(v)} />
                        <Legend wrapperStyle={{ fontSize: 11 }} />
                        <Line type="monotone" dataKey="Penjualan" stroke="#34d399" strokeWidth={2} />
                        <Line type="monotone" dataKey="LabaBersih" stroke="#a3e635" strokeWidth={2} />
                      </LineChart>
                    </ResponsiveContainer>
                  ) : (
                    <div className="h-full flex items-center justify-center text-sm text-gray-500 text-center px-6">
                      Isi "Data Tambahan Opsional" (penjualan & laba tahun lalu) untuk melihat tren.
                    </div>
                  )}
                </div>
              </Card>
            </div>

            <Card>
              <CardHeader title="Donut Chart Skor per Kategori" icon={BarChart3} />
              <div className="p-4 h-64">
                {scoreDonutData.length ? (
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={scoreDonutData} dataKey="value" nameKey="name" innerRadius={50} outerRadius={80} label={{ fill: "#9ca3af", fontSize: 10 }}>
                        {scoreDonutData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                      </Pie>
                      <Tooltip contentStyle={{ background: "#111827", border: "1px solid #10b981" }} />
                    </PieChart>
                  </ResponsiveContainer>
                ) : (
                  <div className="h-full flex items-center justify-center text-sm text-gray-500">Belum ada data.</div>
                )}
              </div>
            </Card>

            {CATEGORIES.map(cat => (
              <Card key={cat}>
                <CardHeader
                  title={`Rasio ${cat}`}
                  icon={Calculator}
                  right={scoring.byCategory[cat].max > 0 && (
                    <span className="text-xs text-gray-400 font-mono">{scoring.byCategory[cat].sum} / {scoring.byCategory[cat].max}</span>
                  )}
                />
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-left text-gray-500 text-xs border-b border-emerald-500/10">
                        <th className="py-2 px-5 font-medium">Rasio</th>
                        <th className="py-2 px-5 font-medium">Nilai</th>
                        <th className="py-2 px-5 font-medium">Penilaian</th>
                      </tr>
                    </thead>
                    <tbody>
                      {scoring.rows.filter(r => r.cat === cat).map(r => (
                        <tr key={r.key} className="border-b border-gray-800/50 hover:bg-emerald-500/5">
                          <td className="py-2 px-5 text-gray-300">{r.label}</td>
                          <td className="py-2 px-5 font-mono text-emerald-100">{fmtVal(r.value, r.unit)}</td>
                          <td className="py-2 px-5">{r.scored ? <Badge label={r.scored.label} /> : <span className="text-gray-600 text-xs">—</span>}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Card>
            ))}

            <Card>
              <CardHeader title="Red Flags" icon={AlertTriangle} />
              <div className="p-5">
                {redFlags.length === 0 ? (
                  <p className="text-sm text-gray-400 flex items-center gap-2"><CheckCircle2 className="w-4 h-4 text-emerald-400" /> Tidak terdeteksi red flag signifikan dari data yang diisi.</p>
                ) : (
                  <ul className="space-y-2">
                    {redFlags.map((f, i) => (
                      <li key={i} className="text-sm text-red-300 flex items-start gap-2">
                        <AlertTriangle className="w-4 h-4 mt-0.5 shrink-0" /> {f}
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            </Card>

            <div className="grid md:grid-cols-3 gap-4">
              <Card>
                <CardHeader title="Altman Z-Score" icon={Calculator} />
                <div className="p-5">
                  {altman ? (
                    <>
                      <div className="text-2xl font-bold font-mono text-emerald-300">{altman.z.toFixed(2)}</div>
                      <div className="text-xs text-gray-400 mt-1">{altman.zone}</div>
                      <p className="text-[11px] text-gray-600 mt-2">*Proksi laba ditahan menggunakan Total Ekuitas (simplifikasi).</p>
                    </>
                  ) : <p className="text-sm text-gray-500">Isi data neraca & laba rugi lengkap.</p>}
                </div>
              </Card>
              <Card>
                <CardHeader title="Piotroski F-Score" icon={Calculator} />
                <div className="p-5">
                  {piotroski ? (
                    <>
                      <div className="text-2xl font-bold font-mono text-emerald-300">{piotroski.score} / 9</div>
                      <ul className="mt-2 space-y-1">
                        {piotroski.criteria.map((c, i) => (
                          <li key={i} className="text-[11px] text-gray-400 flex items-center gap-1.5">
                            {c.pass ? <CheckCircle2 className="w-3 h-3 text-emerald-400 shrink-0" /> : <XCircle className="w-3 h-3 text-red-400 shrink-0" />} {c.text}
                          </li>
                        ))}
                      </ul>
                    </>
                  ) : <p className="text-sm text-gray-500">Isi "Data Tambahan Opsional" untuk menghitung.</p>}
                </div>
              </Card>
              <Card>
                <CardHeader title="Kualitas Arus Kas" icon={Calculator} />
                <div className="p-5">
                  {cfq ? (
                    <>
                      <div className="text-2xl font-bold font-mono text-emerald-300">{cfq.cfq.toFixed(2)}x</div>
                      <div className="mt-1"><Badge label={cfq.label} /></div>
                      <p className="text-[11px] text-gray-600 mt-2">CFO ÷ Laba Bersih. Idealnya &gt; 1x.</p>
                    </>
                  ) : <p className="text-sm text-gray-500">Isi Arus Kas Operasi (CFO) untuk menghitung.</p>}
                </div>
              </Card>
            </div>
          </div>
        )}

        {/* VALUASI */}
        {tab === "valuasi" && (
          <div className="space-y-5">
            <h1 className="text-xl font-bold text-white">Valuasi & Estimasi Harga Wajar</h1>

            <Card>
              <CardHeader title="Parameter Valuasi" icon={Calculator} />
              <div className="p-5 grid sm:grid-cols-2 lg:grid-cols-5 gap-4">
                <NumberField label="PER Target" value={params.perTarget} onChange={v => setParams(s => ({ ...s, perTarget: v }))} suffix="x" />
                <NumberField label="PBV Target" value={params.pbvTarget} onChange={v => setParams(s => ({ ...s, pbvTarget: v }))} suffix="x" />
                <NumberField label="Growth (DCF)" value={params.growth * 100} onChange={v => setParams(s => ({ ...s, growth: v / 100 }))} suffix="%" />
                <NumberField label="Discount Rate (DCF)" value={params.discountRate * 100} onChange={v => setParams(s => ({ ...s, discountRate: v / 100 }))} suffix="%" />
                <NumberField label="Terminal Growth (DCF)" value={params.terminalGrowth * 100} onChange={v => setParams(s => ({ ...s, terminalGrowth: v / 100 }))} suffix="%" />
              </div>
            </Card>

            <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-4">
              {valuation.methods.map(m => {
                const diff = m.value ? ((pasar.hargaSaham - m.value) / m.value) * 100 : null;
                const v = diff === null ? null : diff < -5 ? "UNDERVALUED" : diff > 5 ? "OVERVALUED" : "FAIR VALUE";
                const vColor = v === "UNDERVALUED" ? "text-emerald-400" : v === "OVERVALUED" ? "text-red-400" : "text-yellow-400";
                return (
                  <Card key={m.key} className="p-5">
                    <div className="text-xs text-gray-400 uppercase tracking-wide">{m.label}</div>
                    <div className="text-xl font-bold font-mono text-white mt-1">{fmtIDR(m.value)}</div>
                    {v && <div className={`text-xs font-semibold mt-2 ${vColor}`}>{v === "UNDERVALUED" ? "🟢" : v === "OVERVALUED" ? "🔴" : "🟡"} {v}</div>}
                  </Card>
                );
              })}
            </div>

            <div className="grid md:grid-cols-2 gap-4">
              <Card className="p-6">
                <div className="text-xs text-gray-400 uppercase tracking-wide mb-1">Harga Wajar Rata-Rata Gabungan</div>
                <div className="text-3xl font-bold font-mono text-emerald-300">{fmtIDR(valuation.avgFair)}</div>
                <div className="text-sm text-gray-500 mt-1">Harga saat ini: <span className="font-mono text-gray-300">{fmtIDR(pasar.hargaSaham)}</span></div>
                {valuation.verdict && (
                  <div className={`mt-4 inline-flex items-center gap-2 px-3 py-1.5 rounded-lg text-sm font-bold ${
                    valuation.verdict === "UNDERVALUED" ? "bg-emerald-500/10 text-emerald-300 border border-emerald-500/40" :
                    valuation.verdict === "OVERVALUED" ? "bg-red-500/10 text-red-300 border border-red-500/40" :
                    "bg-yellow-500/10 text-yellow-300 border border-yellow-500/40"
                  }`}>
                    {valuation.verdict === "UNDERVALUED" ? "🟢" : valuation.verdict === "OVERVALUED" ? "🔴" : "🟡"} {valuation.verdict}
                  </div>
                )}
              </Card>

              <Card className="p-6">
                <div className="text-xs text-gray-400 uppercase tracking-wide mb-1">Margin of Safety</div>
                <div className="text-3xl font-bold font-mono text-emerald-300">{valuation.mos !== null ? `${valuation.mos.toFixed(1)}%` : "-"}</div>
                {valuation.mosLabel && <div className="mt-2"><Badge label={valuation.mosLabel === "Overvalued" ? "Buruk" : valuation.mosLabel === "Tipis" ? "Cukup" : valuation.mosLabel === "Aman" ? "Baik" : "Sangat Baik"} /> <span className="text-xs text-gray-400 ml-1">{valuation.mosLabel}</span></div>}
              </Card>
            </div>

            <Card className="p-6">
              <div className="flex items-center justify-between flex-wrap gap-4">
                <div>
                  <div className="text-xs text-gray-400 uppercase tracking-wide mb-1">Skor Investasi Keseluruhan</div>
                  <div className="text-4xl font-bold font-mono text-emerald-300">{invScore.total.toFixed(0)} <span className="text-lg text-gray-500">/ 100</span></div>
                </div>
                <div className={`px-4 py-2 rounded-lg text-lg font-bold border ${
                  invScore.recommendation === "Strong Buy" ? "bg-emerald-500/10 text-emerald-300 border-emerald-500/40" :
                  invScore.recommendation === "Buy" ? "bg-lime-500/10 text-lime-300 border-lime-500/40" :
                  invScore.recommendation === "Hold" ? "bg-yellow-500/10 text-yellow-300 border-yellow-500/40" :
                  invScore.recommendation === "Sell" ? "bg-orange-500/10 text-orange-300 border-orange-500/40" :
                  "bg-red-500/10 text-red-300 border-red-500/40"
                }`}>{invScore.recommendation}</div>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-5">
                <MiniBar label="Fundamental (40%)" value={invScore.breakdown.fundamentalScore} />
                <MiniBar label="Valuasi (30%)" value={invScore.breakdown.valuationScore} />
                <MiniBar label="Growth (20%)" value={invScore.breakdown.growthScore} />
                <MiniBar label="Risiko (10%)" value={invScore.breakdown.riskScore} />
              </div>
              <p className="text-[11px] text-gray-600 mt-4">
                Aturan rekomendasi: ≥80 Strong Buy · 65–79 Buy · 50–64 Hold · 35–49 Sell · &lt;35 Avoid.
              </p>
            </Card>
          </div>
        )}

        {/* RIWAYAT */}
        {tab === "riwayat" && (
          <div className="space-y-5">
            <h1 className="text-xl font-bold text-white">Riwayat Analisa</h1>
            <Card>
              {history.length === 0 ? (
                <div className="p-10 text-center text-sm text-gray-500">
                  Belum ada riwayat. Simpan analisa dari tab Dashboard setelah data lengkap.
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-left text-gray-500 text-xs border-b border-emerald-500/10">
                        <th className="py-2 px-5 font-medium">Emiten</th>
                        <th className="py-2 px-5 font-medium">Tanggal</th>
                        <th className="py-2 px-5 font-medium">Grade</th>
                        <th className="py-2 px-5 font-medium">Harga Wajar</th>
                        <th className="py-2 px-5 font-medium">Harga Saat Ini</th>
                        <th className="py-2 px-5 font-medium">Kesimpulan</th>
                        <th className="py-2 px-5 font-medium">Rekomendasi</th>
                      </tr>
                    </thead>
                    <tbody>
                      {history.map(h => (
                        <tr key={h.id} className="border-b border-gray-800/50 hover:bg-emerald-500/5">
                          <td className="py-2 px-5 text-gray-200">{h.name}</td>
                          <td className="py-2 px-5 text-gray-400 text-xs">{h.date}</td>
                          <td className="py-2 px-5 font-mono text-emerald-300">{h.grade}</td>
                          <td className="py-2 px-5 font-mono">{fmtIDR(h.avgFair)}</td>
                          <td className="py-2 px-5 font-mono">{fmtIDR(h.harga)}</td>
                          <td className="py-2 px-5">{h.verdict || "-"}</td>
                          <td className="py-2 px-5">{h.recommendation}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </Card>
          </div>
        )}
      </main>
    </div>
  );
}
  

function StatCard({ label, value, sub, color }) {
  return (
    <div className="rounded-xl border border-emerald-500/20 bg-zinc-900/60 backdrop-blur-sm p-4">
      <div className="text-[11px] text-gray-500 uppercase tracking-wide">{label}</div>
      <div className="text-2xl font-bold font-mono mt-1" style={{ color }}>{value}</div>
      <div className="text-xs text-gray-500 mt-0.5">{sub}</div>
    </div>
  );
}
function MiniBar({ label, value }) {
  return (
    <div>
      <div className="flex justify-between text-[10px] text-gray-500 mb-1">
        <span>{label}</span><span className="font-mono">{value.toFixed(0)}</span>
      </div>
      <div className="h-1.5 w-full bg-gray-800 rounded-full overflow-hidden">
        <div className="h-full bg-emerald-400" style={{ width: `${Math.min(100, Math.max(0, value))}%` }} />
      </div>
    </div>
  );
}
function NumberFieldText({ label, value, onChange }) {
  return (
    <label className="flex flex-col gap-1 max-w-sm">
      <span className="text-xs text-gray-400">{label}</span>
      <input
        type="text"
        value={value}
        onChange={e => onChange(e.target.value)}
        placeholder="Contoh: PT Bank Central Asia Tbk (BBCA)"
        className="w-full bg-black/40 border border-emerald-500/20 rounded-lg px-3 py-2 text-sm text-emerald-100 focus:outline-none focus:ring-1 focus:ring-emerald-400 focus:border-emerald-400 placeholder-gray-600"
      />
    </label>
  );
}
