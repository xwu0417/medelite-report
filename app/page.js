/**
 * @author Angela Wu
 * @project Medelite Facility Assessment Report Generator
 * @date 2026
 * @description Fetches CMS nursing home data by CCN and generates
 *              PDF and Word Doc facility assessment snapshots.
 */
"use client";
import { useState } from "react";
import jsPDF from "jspdf";
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, WidthType, AlignmentType, ImageRun, HeadingLevel } from "docx";
import { saveAs } from "file-saver";

// ─── CMS API Dataset IDs ───────────────────────────────────────────────────
const DS_PROVIDER  = "4pq5-n9py";
const DS_STATE_AVG = "xcdc-v8bm";
const DS_CLAIMS    = "ijh5-nb2v";

const CODE_STR_HOSP = "521";
const CODE_STR_ED   = "522";
const CODE_LT_HOSP  = "551";
const CODE_LT_ED    = "552";

const COL_STR_HOSP_AVG = "percentage_of_short_stay_residents_who_were_rehospitalized__1d02";
const COL_STR_ED_AVG   = "percentage_of_short_stay_residents_who_had_an_outpatient_em_d911";
const COL_LT_HOSP_AVG  = "number_of_hospitalizations_per_1000_longstay_resident_days";
const COL_LT_ED_AVG    = "number_of_outpatient_emergency_department_visits_per_1000_l_de9d";

const CMS_BASE = "https://data.cms.gov/provider-data/api/1/datastore/query";
function cmsUrl(dataset, params) {
  const full = `${CMS_BASE}/${dataset}/0?${params}`;
  return `/api/cms?endpoint=${encodeURIComponent(full)}`;
}

// ─── Star renderer ────────────────────────────────────────────────────────
function Stars({ value }) {
  const n = parseInt(value) || 0;
  return (
    <span>
      {[1,2,3,4,5].map(i => (
        <span key={i} style={{ color: i <= n ? "#f59e0b" : "#d1d5db", fontSize: 18 }}>★</span>
      ))}
      <span style={{ marginLeft: 6, fontSize: 13, color: "#6b7280", fontWeight: 500 }}>{n}/5</span>
    </span>
  );
}

// ─── Metric badge ─────────────────────────────────────────────────────────
function MetricBadge({ value, facilityVal, avgVal, higherIsBad = true }) {
  if (!facilityVal || facilityVal === "N/A" || !avgVal || avgVal === "N/A") {
    return <span style={{ fontWeight: 600, color: "#111" }}>{value}</span>;
  }
  const fNum = parseFloat(facilityVal);
  const aNum = parseFloat(avgVal);
  const isBad = higherIsBad ? fNum > aNum : fNum < aNum;
  const color = isBad ? "#dc2626" : "#16a34a";
  const bg    = isBad ? "#fef2f2" : "#f0fdf4";
  return (
    <span style={{ fontWeight: 700, color, background: bg, borderRadius: 6, padding: "2px 8px", fontSize: 13 }}>
      {value}
    </span>
  );
}

export default function Home() {
  const [ccn, setCcn] = useState("");
  const [facilityData, setFacilityData] = useState(null);
  const [claimsData,   setClaimsData]   = useState([]);
  const [stateAvg,     setStateAvg]     = useState({});
  const [nationAvg,    setNationAvg]    = useState({});
  const [loading,      setLoading]      = useState(false);
  const [error,        setError]        = useState(null);
  const [dataReady,    setDataReady]    = useState(false);

  const [nameOverride,        setNameOverride]        = useState("");
  const [emr,                 setEmr]                 = useState("");
  const [currentCensus,       setCurrentCensus]       = useState("");
  const [patientType,         setPatientType]         = useState("");
  const [previousCoverage,    setPreviousCoverage]    = useState("");
  const [previousPerformance, setPreviousPerformance] = useState("");
  const [medicalCoverage,     setMedicalCoverage]     = useState("");

  async function fetchFacility() {
    if (!ccn.trim()) return;
    setLoading(true);
    setDataReady(false);
    setError(null);
    try {
      const provRes  = await fetch(cmsUrl(DS_PROVIDER,
        `conditions[0][property]=cms_certification_number_ccn&conditions[0][value]=${ccn.trim()}&conditions[0][operator]==`
      ));
      const provJson = await provRes.json();
      if (!provJson.results?.length) {
        setError("No facility found for that CCN. Double-check and try again.");
        setLoading(false);
        return;
      }
      const facility = provJson.results[0];
      const state    = facility.state;

      const [claimsRes, stateRes, nationRes] = await Promise.all([
        fetch(cmsUrl(DS_CLAIMS,
          `conditions[0][property]=cms_certification_number_ccn&conditions[0][value]=${ccn.trim()}&conditions[0][operator]==&limit=100`
        )),
        fetch(cmsUrl(DS_STATE_AVG,
          `conditions[0][property]=state_or_nation&conditions[0][value]=${state}&conditions[0][operator]==`
        )),
        fetch(cmsUrl(DS_STATE_AVG,
          `conditions[0][property]=state_or_nation&conditions[0][value]=NATION&conditions[0][operator]==`
        )),
      ]);
      const [claimsJson, stateJson, nationJson] = await Promise.all([
        claimsRes.json(), stateRes.json(), nationRes.json(),
      ]);
      setFacilityData(facility);
      setClaimsData(claimsJson.results || []);
      setStateAvg(stateJson.results?.[0]  || {});
      setNationAvg(nationJson.results?.[0] || {});
      setNameOverride("");
      setDataReady(true);
    } catch (err) {
      console.error(err);
      setError("Something went wrong. Check your network and try again.");
    }
    setLoading(false);
  }

  const displayName = nameOverride.trim() || facilityData?.provider_name || "";

  function fmt(val, isPercent = false) {
    if (val === null || val === undefined || val === "") return "N/A";
    const num = parseFloat(val);
    if (isNaN(num)) return "N/A";
    return isPercent ? `${num.toFixed(1)}%` : num.toFixed(2);
  }

  function claimScore(code) {
    const row = claimsData.find(r => String(r.measure_code) === String(code));
    return row?.adjusted_score ?? row?.observed_score ?? null;
  }

  function getAllRows() {
    if (!facilityData) return [];
    return [
      ["Name of Facility",                            displayName],
      ["Location",                                    `${facilityData.provider_address}, ${facilityData.citytown}, ${facilityData.state} ${facilityData.zip_code}`],
      ["EMR",                                         emr || ""],
      ["Census Capacity",                             facilityData.number_of_certified_beds || ""],
      ["Current Census",                              currentCensus || ""],
      ["Type of Patient",                             patientType || ""],
      ["Previous Coverage from Medelite",             previousCoverage || ""],
      ["Previous Provider Performance from Medelite", previousPerformance ? `About ${previousPerformance} patients/day` : ""],
      ["Medical Coverage",                            medicalCoverage || ""],
      ["Overall Star Rating",                         facilityData.overall_rating || ""],
      ["Health Inspection",                           facilityData.health_inspection_rating || ""],
      ["Staffing",                                    facilityData.staffing_rating || ""],
      ["Quality of Resident Care",                    facilityData.qm_rating || ""],
      ["Short Term Hospitalization",                  fmt(claimScore(CODE_STR_HOSP), true)],
      ["STR National Avg. for Hospitalization",       fmt(nationAvg[COL_STR_HOSP_AVG], true)],
      ["STR State Avg. for Hospitalization",          fmt(stateAvg[COL_STR_HOSP_AVG],  true)],
      ["STR ED Visit",                                fmt(claimScore(CODE_STR_ED), true)],
      ["STR ED Visits National Avg.",                 fmt(nationAvg[COL_STR_ED_AVG], true)],
      ["STR ED Visits State Avg.",                    fmt(stateAvg[COL_STR_ED_AVG],  true)],
      ["LT Hospitalization",                          fmt(claimScore(CODE_LT_HOSP))],
      ["LT National Avg. for Hospitalization",        fmt(nationAvg[COL_LT_HOSP_AVG])],
      ["LT State Avg. for Hospitalization",           fmt(stateAvg[COL_LT_HOSP_AVG])],
      ["ED Visit",                                    fmt(claimScore(CODE_LT_ED))],
      ["LT ED Visits National Avg.",                  fmt(nationAvg[COL_LT_ED_AVG])],
      ["LT ED Visits State Avg.",                     fmt(stateAvg[COL_LT_ED_AVG])],
    ];
  }

  // ─── PDF ──────────────────────────────────────────────────────────────────
  async function downloadPDF() {
    if (!facilityData || !dataReady) return;
    const doc = new jsPDF();
    try {
      const blob   = await (await fetch(window.location.origin + "/infinite_logo.png")).blob();
      const base64 = await new Promise(res => { const r = new FileReader(); r.onloadend = () => res(r.result); r.readAsDataURL(blob); });
      doc.addImage(base64, "PNG", 65, 8, 80, 22);
    } catch { console.warn("Logo skipped"); }

    doc.setFontSize(13); doc.setFont("helvetica", "bold"); doc.setTextColor(0,0,0);
    doc.text("FACILITY ASSESSMENT SNAPSHOT", 105, 38, { align: "center" });
    doc.setFontSize(11);
    doc.text(facilityData.state || "", 105, 45, { align: "center" });

    const startY = 52, leftX = 20, colSplit = 100, rightX = 190, rowH = 8.5, pageH = 280;
    let y = startY;
    getAllRows().forEach(([label, value], i) => {
      if (y + rowH > pageH) { doc.addPage(); y = 20; }
      if (i % 2 === 0) { doc.setFillColor(245,245,245); doc.rect(leftX, y, rightX-leftX, rowH, "F"); }
      doc.setDrawColor(180,180,180);
      doc.rect(leftX, y, rightX-leftX, rowH, "S");
      doc.line(colSplit, y, colSplit, y+rowH);
      doc.setFontSize(8.5); doc.setFont("helvetica","bold"); doc.setTextColor(0,0,0);
      doc.text(doc.splitTextToSize(label, colSplit-leftX-3)[0], leftX+2, y+5.5);
      doc.setFont("helvetica","normal");
      doc.text(String(value), colSplit+3, y+5.5);
      y += rowH;
    });
    doc.setFontSize(9); doc.setTextColor(6,69,173);
    doc.textWithLink("View official Medicare Care Compare profile", leftX, y+10,
      { url: `https://www.medicare.gov/care-compare/details/nursing-home/${ccn.trim()}` });
    doc.save(`Facility_Assessment_${displayName.replace(/\s+/g,"_")}.pdf`);
  }

  // ─── DOCX ─────────────────────────────────────────────────────────────────
  async function downloadDocx() {
    if (!facilityData || !dataReady) return;
    let logoBuffer = null;
    try { logoBuffer = await (await fetch(window.location.origin + "/infinite_logo.png")).arrayBuffer(); } catch {}
    const allRows = getAllRows();
    const tableRows = allRows.map(([label, value], i) => new TableRow({ children: [
      new TableCell({ width:{size:50,type:WidthType.PERCENTAGE}, shading: i%2===0?{fill:"F5F5F5"}:undefined,
        children:[new Paragraph({children:[new TextRun({text:label,bold:true,size:20})]})] }),
      new TableCell({ width:{size:50,type:WidthType.PERCENTAGE}, shading: i%2===0?{fill:"F5F5F5"}:undefined,
        children:[new Paragraph({children:[new TextRun({text:String(value),size:20})]})] }),
    ]}));
    const children = [];
    if (logoBuffer) children.push(new Paragraph({ alignment:AlignmentType.CENTER, children:[new ImageRun({data:logoBuffer,transformation:{width:200,height:55},type:"png"})] }));
    children.push(
      new Paragraph({text:"FACILITY ASSESSMENT SNAPSHOT",heading:HeadingLevel.HEADING_1,alignment:AlignmentType.CENTER}),
      new Paragraph({text:facilityData.state||"",alignment:AlignmentType.CENTER,spacing:{after:200}}),
      new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:tableRows}),
      new Paragraph({text:""}),
      new Paragraph({children:[new TextRun({text:"View official Medicare Care Compare profile",color:"0645AD",underline:{},size:18})]})
    );
    const blob = await Packer.toBlob(new Document({sections:[{children}]}));
    saveAs(blob, `Facility_Assessment_${displayName.replace(/\s+/g,"_")}.docx`);
  }

  // ─── Styles ───────────────────────────────────────────────────────────────
  const inputStyle = {
    padding: "10px 14px", fontSize: 15, width: "100%",
    borderRadius: 10, border: "1.5px solid #e5e7eb",
    boxSizing: "border-box", outline: "none",
    background: "#fafafa", color: "#111",
    transition: "border-color 0.2s",
  };
  const labelStyle = {
    fontWeight: 600, fontSize: 12, color: "#6b7280",
    letterSpacing: "0.05em", textTransform: "uppercase",
    display: "block", marginBottom: 6,
  };
  const card = {
    background: "#fff", borderRadius: 16,
    border: "1px solid #f0f0f0",
    boxShadow: "0 1px 4px rgba(0,0,0,0.06)",
    padding: 28, marginBottom: 24,
  };
  const sectionTitle = {
    fontSize: 11, fontWeight: 700, letterSpacing: "0.1em",
    textTransform: "uppercase", color: "#9ca3af",
    margin: "0 0 16px 0",
  };

  const PINK = "#c0218e";
  const NAVY = "#1a56a0";

  return (
    <main style={{ fontFamily: "'Inter', 'Segoe UI', sans-serif", background: "#f8f9fb", minHeight: "100vh", padding: "0 0 60px" }}>

      {/* ── Top bar ── */}
      <div style={{ background: "#fff", borderBottom: "1px solid #f0f0f0", padding: "18px 0", marginBottom: 36 }}>
        <div style={{ maxWidth: 780, margin: "0 auto", padding: "0 24px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img src="/infinite_logo.png" alt="INFINITE Managed by MEDELITE" style={{ height: 36, width: "auto" }} />
          <span style={{ fontSize: 12, color: "#9ca3af", letterSpacing: "0.08em", textTransform: "uppercase" }}>
            Facility Assessment Tool
          </span>
        </div>
      </div>

      <div style={{ maxWidth: 780, margin: "0 auto", padding: "0 24px" }}>

        {/* ── Hero ── */}
        <div style={{ marginBottom: 32 }}>
          <h1 style={{ fontSize: 28, fontWeight: 800, color: "#111", margin: "0 0 6px", letterSpacing: "-0.02em" }}>
            Facility Assessment
          </h1>
          <p style={{ fontSize: 15, color: "#6b7280", margin: 0 }}>
            Enter a CMS Certification Number to generate a snapshot report.
          </p>
        </div>

        {/* ── CCN Lookup ── */}
        <div style={card}>
          <p style={sectionTitle}>CMS Certification Number</p>
          <div style={{ display: "flex", gap: 12 }}>
            <input
              type="text"
              placeholder="e.g. 686123"
              value={ccn}
              onChange={(e) => setCcn(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && fetchFacility()}
              style={{ ...inputStyle, flex: 1 }}
            />
            <button
              onClick={fetchFacility}
              disabled={loading}
              style={{
                padding: "10px 24px", background: PINK, color: "#fff",
                border: "none", borderRadius: 10, fontSize: 14,
                fontWeight: 700, cursor: loading ? "not-allowed" : "pointer",
                whiteSpace: "nowrap", opacity: loading ? 0.7 : 1,
                transition: "opacity 0.2s",
              }}
            >
              {loading ? "Loading…" : "Fetch Data"}
            </button>
          </div>
          {error && (
            <div style={{ marginTop: 12, padding: "10px 14px", background: "#fef2f2", borderRadius: 8, color: "#dc2626", fontSize: 13, fontWeight: 500 }}>
              ⚠ {error}
            </div>
          )}
        </div>

        {/* ── Loading ── */}
        {loading && (
          <div style={{ ...card, textAlign: "center", padding: 48 }}>
            <div style={{ fontSize: 36, marginBottom: 12 }}>🏥</div>
            <p style={{ color: "#6b7280", margin: 0, fontSize: 15 }}>Pulling facility data from CMS…</p>
          </div>
        )}

        {dataReady && facilityData && (
          <>
            {/* ── Facility header ── */}
            <div style={{
              ...card,
              background: `linear-gradient(135deg, ${PINK} 0%, #7b1fa2 100%)`,
              border: "none", color: "#fff", marginBottom: 24,
            }}>
              <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", opacity: 0.75, marginBottom: 8 }}>
                {facilityData.state} · CCN {ccn.trim()}
              </div>
              <div style={{ fontSize: 22, fontWeight: 800, letterSpacing: "-0.01em", marginBottom: 4 }}>
                {displayName}
              </div>
              <div style={{ fontSize: 13, opacity: 0.8 }}>
                {facilityData.provider_address}, {facilityData.citytown}, {facilityData.state} {facilityData.zip_code}
              </div>
            </div>

            {/* ── Star ratings ── */}
            <div style={{ ...card }}>
              <p style={sectionTitle}>CMS Star Ratings</p>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                {[
                  ["Overall Rating",          facilityData.overall_rating],
                  ["Health Inspection",       facilityData.health_inspection_rating],
                  ["Staffing",               facilityData.staffing_rating],
                  ["Quality of Resident Care",facilityData.qm_rating],
                ].map(([label, val]) => (
                  <div key={label} style={{ background: "#f8f9fb", borderRadius: 12, padding: "14px 16px" }}>
                    <div style={{ fontSize: 11, color: "#9ca3af", fontWeight: 600, letterSpacing: "0.05em", textTransform: "uppercase", marginBottom: 6 }}>{label}</div>
                    <Stars value={val} />
                  </div>
                ))}
              </div>
            </div>

            {/* ── Manual inputs ── */}
            <div style={card}>
              <p style={sectionTitle}>Facility Details</p>

              <label style={labelStyle}>Facility Name Override</label>
              <input type="text" placeholder={facilityData.provider_name} value={nameOverride}
                onChange={(e) => setNameOverride(e.target.value)}
                style={{ ...inputStyle, marginBottom: 4 }} />
              <small style={{ color: "#9ca3af", display: "block", marginBottom: 16, fontSize: 12 }}>
                Leave blank to use: &quot;{facilityData.provider_name}&quot;
              </small>

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                <div>
                  <label style={labelStyle}>EMR System</label>
                  <input type="text" placeholder="e.g. PCC" value={emr} onChange={(e) => setEmr(e.target.value)} style={inputStyle} />
                </div>
                <div>
                  <label style={labelStyle}>Current Census</label>
                  <input type="number" placeholder="e.g. 112" value={currentCensus} onChange={(e) => setCurrentCensus(e.target.value)} style={inputStyle} />
                </div>
                <div>
                  <label style={labelStyle}>Type of Patient</label>
                  <select value={patientType} onChange={(e) => setPatientType(e.target.value)} style={inputStyle}>
                    <option value="">Select…</option>
                    <option>Long-term &amp; Short-term</option>
                    <option>Long-term</option>
                    <option>Short-term</option>
                  </select>
                </div>
                <div>
                  <label style={labelStyle}>Previous Medelite Coverage</label>
                  <select value={previousCoverage} onChange={(e) => setPreviousCoverage(e.target.value)} style={inputStyle}>
                    <option value="">Select…</option>
                    <option>Yes</option>
                    <option>No</option>
                  </select>
                </div>
                <div>
                  <label style={labelStyle}>Prev. Provider Performance (pts/day)</label>
                  <input type="text" placeholder="e.g. 30" value={previousPerformance} onChange={(e) => setPreviousPerformance(e.target.value)} style={inputStyle} />
                </div>
                <div>
                  <label style={labelStyle}>Medical Coverage</label>
                  <input type="text" placeholder="e.g. Optometry, PCP" value={medicalCoverage} onChange={(e) => setMedicalCoverage(e.target.value)} style={inputStyle} />
                </div>
              </div>
            </div>

            {/* ── Metrics ── */}
            <div style={card}>
              <p style={sectionTitle}>Hospitalization & ED Metrics</p>

              {/* STR Hospitalization */}
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10, borderBottom: "1px solid #f0f0f0", paddingBottom: 6 }}>
                  Short-Term Hospitalization
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
                  {[
                    ["Facility",  fmt(claimScore(CODE_STR_HOSP), true)],
                    ["National",  fmt(nationAvg[COL_STR_HOSP_AVG], true)],
                    ["State",     fmt(stateAvg[COL_STR_HOSP_AVG],  true)],
                  ].map(([lbl, val]) => (
                    <div key={lbl} style={{ background: "#f8f9fb", borderRadius: 10, padding: "12px 14px" }}>
                      <div style={{ fontSize: 10, color: "#9ca3af", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>{lbl}</div>
                      <MetricBadge value={val} facilityVal={fmt(claimScore(CODE_STR_HOSP), true)} avgVal={fmt(nationAvg[COL_STR_HOSP_AVG], true)} />
                    </div>
                  ))}
                </div>
              </div>

              {/* STR ED Visit */}
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10, borderBottom: "1px solid #f0f0f0", paddingBottom: 6 }}>
                  Short-Term ED Visits
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
                  {[
                    ["Facility", fmt(claimScore(CODE_STR_ED), true)],
                    ["National", fmt(nationAvg[COL_STR_ED_AVG], true)],
                    ["State",    fmt(stateAvg[COL_STR_ED_AVG],  true)],
                  ].map(([lbl, val]) => (
                    <div key={lbl} style={{ background: "#f8f9fb", borderRadius: 10, padding: "12px 14px" }}>
                      <div style={{ fontSize: 10, color: "#9ca3af", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>{lbl}</div>
                      <MetricBadge value={val} facilityVal={fmt(claimScore(CODE_STR_ED), true)} avgVal={fmt(nationAvg[COL_STR_ED_AVG], true)} />
                    </div>
                  ))}
                </div>
              </div>

              {/* LT Hospitalization */}
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10, borderBottom: "1px solid #f0f0f0", paddingBottom: 6 }}>
                  Long-Term Hospitalization (per 1,000 days)
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
                  {[
                    ["Facility", fmt(claimScore(CODE_LT_HOSP))],
                    ["National", fmt(nationAvg[COL_LT_HOSP_AVG])],
                    ["State",    fmt(stateAvg[COL_LT_HOSP_AVG])],
                  ].map(([lbl, val]) => (
                    <div key={lbl} style={{ background: "#f8f9fb", borderRadius: 10, padding: "12px 14px" }}>
                      <div style={{ fontSize: 10, color: "#9ca3af", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>{lbl}</div>
                      <MetricBadge value={val} facilityVal={fmt(claimScore(CODE_LT_HOSP))} avgVal={fmt(nationAvg[COL_LT_HOSP_AVG])} />
                    </div>
                  ))}
                </div>
              </div>

              {/* LT ED Visit */}
              <div>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10, borderBottom: "1px solid #f0f0f0", paddingBottom: 6 }}>
                  Long-Term ED Visits (per 1,000 days)
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
                  {[
                    ["Facility", fmt(claimScore(CODE_LT_ED))],
                    ["National", fmt(nationAvg[COL_LT_ED_AVG])],
                    ["State",    fmt(stateAvg[COL_LT_ED_AVG])],
                  ].map(([lbl, val]) => (
                    <div key={lbl} style={{ background: "#f8f9fb", borderRadius: 10, padding: "12px 14px" }}>
                      <div style={{ fontSize: 10, color: "#9ca3af", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>{lbl}</div>
                      <MetricBadge value={val} facilityVal={fmt(claimScore(CODE_LT_ED))} avgVal={fmt(nationAvg[COL_LT_ED_AVG])} />
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* ── Full data table (for reference / export) ── */}
            <div style={card}>
              <p style={sectionTitle}>Full Report Preview</p>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <tbody>
                  {getAllRows().map(([label, value], i) => (
                    <tr key={i} style={{ background: i % 2 === 0 ? "#f8f9fb" : "#fff" }}>
                      <td style={{ padding: "9px 12px", border: "1px solid #f0f0f0", fontWeight: 600, width: "55%", color: "#374151" }}>{label}</td>
                      <td style={{ padding: "9px 12px", border: "1px solid #f0f0f0", color: "#111" }}>{value || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <p style={{ marginTop: 12, fontSize: 13 }}>
                <a href={`https://www.medicare.gov/care-compare/details/nursing-home/${ccn.trim()}`}
                  target="_blank" rel="noreferrer" style={{ color: NAVY, fontWeight: 600 }}>
                  ↗ View official Medicare Care Compare profile
                </a>
              </p>
            </div>

            {/* ── Download buttons ── */}
            <div style={{ display: "flex", gap: 14, justifyContent: "center" }}>
              <button onClick={downloadPDF} style={{
                padding: "14px 32px", background: PINK, color: "#fff",
                border: "none", borderRadius: 12, fontSize: 15, fontWeight: 700, cursor: "pointer",
                boxShadow: `0 4px 14px ${PINK}55`,
              }}>
                ⬇ Download PDF
              </button>
              <button onClick={downloadDocx} style={{
                padding: "14px 32px", background: NAVY, color: "#fff",
                border: "none", borderRadius: 12, fontSize: 15, fontWeight: 700, cursor: "pointer",
                boxShadow: `0 4px 14px ${NAVY}55`,
              }}>
                ⬇ Download Word Doc
              </button>
            </div>
          </>
        )}

        {/* ── Footer ── */}
        <div style={{ textAlign: "center", color: "#d1d5db", fontSize: 12, marginTop: 48 }}>
          Built by <strong style={{ color: "#9ca3af" }}>Angela Wu</strong> · Medelite Technical Assessment · {new Date().getFullYear()}
        </div>

      </div>
    </main>
  );
}