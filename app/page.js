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

const PINK = "#c0218e";

// ─── Inline styles as constants ───────────────────────────────────────────
const S = {
  page: {
    fontFamily: "'Inter', 'Helvetica Neue', Arial, sans-serif",
    background: "#fff",
    minHeight: "100vh",
    color: "#111",
  },
  topbar: {
    borderBottom: "2px solid #111",
    padding: "14px 40px",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    background: "#fff",
  },
  topbarLabel: {
    fontSize: 11,
    letterSpacing: "0.12em",
    textTransform: "uppercase",
    color: "#888",
    fontWeight: 500,
  },
  wrap: {
    maxWidth: 860,
    margin: "0 auto",
    padding: "48px 40px 80px",
  },
  rule: { borderTop: "1px solid #e5e7eb", margin: "32px 0" },
  sectionEyebrow: {
    fontSize: 10,
    fontWeight: 700,
    letterSpacing: "0.14em",
    textTransform: "uppercase",
    color: PINK,
    marginBottom: 16,
  },
  h1: {
    fontSize: 26,
    fontWeight: 700,
    letterSpacing: "-0.02em",
    margin: "0 0 6px",
    color: "#111",
  },
  sub: { fontSize: 14, color: "#666", margin: 0 },
  label: {
    display: "block",
    fontSize: 11,
    fontWeight: 600,
    letterSpacing: "0.08em",
    textTransform: "uppercase",
    color: "#888",
    marginBottom: 6,
  },
  input: {
    width: "100%",
    padding: "9px 12px",
    fontSize: 14,
    border: "1px solid #d1d5db",
    borderRadius: 0,
    outline: "none",
    background: "#fff",
    color: "#111",
    boxSizing: "border-box",
    fontFamily: "inherit",
  },
  select: {
    width: "100%",
    padding: "9px 12px",
    fontSize: 14,
    border: "1px solid #d1d5db",
    borderRadius: 0,
    outline: "none",
    background: "#fff",
    color: "#111",
    boxSizing: "border-box",
    fontFamily: "inherit",
    appearance: "none",
  },
  btnPrimary: {
    padding: "10px 28px",
    background: PINK,
    color: "#fff",
    border: "none",
    borderRadius: 0,
    fontSize: 13,
    fontWeight: 700,
    letterSpacing: "0.04em",
    cursor: "pointer",
    fontFamily: "inherit",
  },
  btnSecondary: {
    padding: "10px 28px",
    background: "#fff",
    color: "#111",
    border: "1px solid #111",
    borderRadius: 0,
    fontSize: 13,
    fontWeight: 700,
    letterSpacing: "0.04em",
    cursor: "pointer",
    fontFamily: "inherit",
  },
  mono: {
    fontFamily: "'JetBrains Mono', 'Courier New', monospace",
    fontSize: 13,
  },
  facilityName: {
    fontSize: 20,
    fontWeight: 700,
    letterSpacing: "-0.01em",
    color: "#111",
    margin: "0 0 4px",
  },
  facilityMeta: { fontSize: 13, color: "#666", margin: 0 },
};

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
    if (val === null || val === undefined || val === "") return "—";
    const num = parseFloat(val);
    if (isNaN(num)) return "—";
    return isPercent ? `${num.toFixed(1)}%` : num.toFixed(2);
  }

  function claimScore(code) {
    const row = claimsData.find(r => String(r.measure_code) === String(code));
    return row?.adjusted_score ?? row?.observed_score ?? null;
  }

  // Compare facility to national: red if worse, green if better
  function deltaColor(facilityRaw, nationalRaw, higherIsBad = true) {
    const f = parseFloat(facilityRaw);
    const n = parseFloat(nationalRaw);
    if (isNaN(f) || isNaN(n)) return "#111";
    if (higherIsBad) return f > n ? "#b91c1c" : "#15803d";
    return f < n ? "#b91c1c" : "#15803d";
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
      if (i % 2 === 0) { doc.setFillColor(248,249,251); doc.rect(leftX, y, rightX-leftX, rowH, "F"); }
      doc.setDrawColor(220,220,220);
      doc.rect(leftX, y, rightX-leftX, rowH, "S");
      doc.line(colSplit, y, colSplit, y+rowH);
      doc.setFontSize(8.5); doc.setFont("helvetica","bold"); doc.setTextColor(50,50,50);
      doc.text(doc.splitTextToSize(label, colSplit-leftX-3)[0], leftX+2, y+5.5);
      doc.setFont("helvetica","normal"); doc.setTextColor(0,0,0);
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
      new TableCell({ width:{size:50,type:WidthType.PERCENTAGE}, shading: i%2===0?{fill:"F8F9FB"}:undefined,
        children:[new Paragraph({children:[new TextRun({text:label,bold:true,size:20})]})] }),
      new TableCell({ width:{size:50,type:WidthType.PERCENTAGE}, shading: i%2===0?{fill:"F8F9FB"}:undefined,
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

  // ─── Render ───────────────────────────────────────────────────────────────
  return (
    <main style={S.page}>

      {/* Top bar */}
      <div style={S.topbar}>
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src="/infinite_logo.png" alt="INFINITE Managed by MEDELITE" style={{ height: 32, width: "auto" }} />
        <span style={S.topbarLabel}>Facility Assessment Tool</span>
      </div>

      <div style={S.wrap}>

        {/* Page title */}
        <div style={{ marginBottom: 40 }}>
          <h1 style={S.h1}>Facility Assessment Snapshot</h1>
          <p style={S.sub}>Enter a CMS Certification Number to pull live facility data and generate a report.</p>
        </div>

        {/* CCN input */}
        <div style={{ marginBottom: 40 }}>
          <p style={S.sectionEyebrow}>Lookup</p>
          <div style={{ display: "flex", gap: 0, maxWidth: 480 }}>
            <input
              type="text"
              placeholder="CMS Certification Number (e.g. 686123)"
              value={ccn}
              onChange={(e) => setCcn(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && fetchFacility()}
              style={{ ...S.input, flex: 1, borderRight: "none" }}
            />
            <button
              onClick={fetchFacility}
              disabled={loading}
              style={{ ...S.btnPrimary, opacity: loading ? 0.6 : 1 }}
            >
              {loading ? "Loading…" : "Fetch"}
            </button>
          </div>
          {error && (
            <p style={{ marginTop: 10, fontSize: 13, color: "#b91c1c", fontWeight: 500 }}>
              {error}
            </p>
          )}
        </div>

        {/* Loading state */}
        {loading && (
          <div style={{ borderTop: "1px solid #e5e7eb", paddingTop: 32, color: "#888", fontSize: 14 }}>
            Pulling data from CMS Medicare…
          </div>
        )}

        {dataReady && facilityData && (
          <>
            <hr style={S.rule} />

            {/* Facility identity */}
            <div style={{ marginBottom: 40 }}>
              <p style={S.sectionEyebrow}>Facility</p>
              <p style={S.facilityName}>{displayName}</p>
              <p style={S.facilityMeta}>
                {facilityData.provider_address}, {facilityData.citytown}, {facilityData.state} {facilityData.zip_code}
                &nbsp;·&nbsp;CCN {ccn.trim()}
              </p>
            </div>

            {/* Star ratings — plain text, monospace values */}
            <div style={{ marginBottom: 40 }}>
              <p style={S.sectionEyebrow}>CMS Star Ratings</p>
              <table style={{ borderCollapse: "collapse", width: "100%" }}>
                <tbody>
                  {[
                    ["Overall Rating",           facilityData.overall_rating],
                    ["Health Inspection",         facilityData.health_inspection_rating],
                    ["Staffing",                  facilityData.staffing_rating],
                    ["Quality of Resident Care",  facilityData.qm_rating],
                  ].map(([label, val], i) => (
                    <tr key={label} style={{ borderBottom: "1px solid #f0f0f0" }}>
                      <td style={{ padding: "11px 0", fontSize: 14, color: "#444", width: "60%" }}>{label}</td>
                      <td style={{ padding: "11px 0", ...S.mono, fontWeight: 700 }}>
                        {val || "—"} / 5
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <hr style={S.rule} />

            {/* Manual details — 2 col grid */}
            <div style={{ marginBottom: 40 }}>
              <p style={S.sectionEyebrow}>Facility Details</p>
              <div style={{ marginBottom: 16 }}>
                <label style={S.label}>Facility Name Override</label>
                <input type="text" placeholder={facilityData.provider_name} value={nameOverride}
                  onChange={(e) => setNameOverride(e.target.value)} style={{ ...S.input, maxWidth: 480 }} />
                <span style={{ fontSize: 11, color: "#aaa", display: "block", marginTop: 4 }}>
                  Leave blank to use: &quot;{facilityData.provider_name}&quot;
                </span>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "16px 32px", maxWidth: 700 }}>
                <div>
                  <label style={S.label}>EMR System</label>
                  <input type="text" placeholder="e.g. PCC" value={emr}
                    onChange={(e) => setEmr(e.target.value)} style={S.input} />
                </div>
                <div>
                  <label style={S.label}>Current Census</label>
                  <input type="number" placeholder="e.g. 112" value={currentCensus}
                    onChange={(e) => setCurrentCensus(e.target.value)} style={S.input} />
                </div>
                <div>
                  <label style={S.label}>Type of Patient</label>
                  <select value={patientType} onChange={(e) => setPatientType(e.target.value)} style={S.select}>
                    <option value="">Select…</option>
                    <option>Long-term &amp; Short-term</option>
                    <option>Long-term</option>
                    <option>Short-term</option>
                  </select>
                </div>
                <div>
                  <label style={S.label}>Previous Medelite Coverage</label>
                  <select value={previousCoverage} onChange={(e) => setPreviousCoverage(e.target.value)} style={S.select}>
                    <option value="">Select…</option>
                    <option>Yes</option>
                    <option>No</option>
                  </select>
                </div>
                <div>
                  <label style={S.label}>Prev. Performance (patients/day)</label>
                  <input type="text" placeholder="e.g. 30" value={previousPerformance}
                    onChange={(e) => setPreviousPerformance(e.target.value)} style={S.input} />
                </div>
                <div>
                  <label style={S.label}>Medical Coverage</label>
                  <input type="text" placeholder="e.g. Optometry, PCP, Podiatry" value={medicalCoverage}
                    onChange={(e) => setMedicalCoverage(e.target.value)} style={S.input} />
                </div>
              </div>
            </div>

            <hr style={S.rule} />

            {/* Metrics — inline compact rows */}
            <div style={{ marginBottom: 40 }}>
              <p style={S.sectionEyebrow}>Hospitalization &amp; ED Metrics</p>

              {[
                {
                  label: "Short-Term Hospitalization",
                  facility: fmt(claimScore(CODE_STR_HOSP), true),
                  national: fmt(nationAvg[COL_STR_HOSP_AVG], true),
                  state:    fmt(stateAvg[COL_STR_HOSP_AVG],  true),
                  raw: { f: claimScore(CODE_STR_HOSP), n: nationAvg[COL_STR_HOSP_AVG] },
                },
                {
                  label: "Short-Term ED Visits",
                  facility: fmt(claimScore(CODE_STR_ED), true),
                  national: fmt(nationAvg[COL_STR_ED_AVG], true),
                  state:    fmt(stateAvg[COL_STR_ED_AVG],  true),
                  raw: { f: claimScore(CODE_STR_ED), n: nationAvg[COL_STR_ED_AVG] },
                },
                {
                  label: "Long-Term Hospitalization (per 1,000 days)",
                  facility: fmt(claimScore(CODE_LT_HOSP)),
                  national: fmt(nationAvg[COL_LT_HOSP_AVG]),
                  state:    fmt(stateAvg[COL_LT_HOSP_AVG]),
                  raw: { f: claimScore(CODE_LT_HOSP), n: nationAvg[COL_LT_HOSP_AVG] },
                },
                {
                  label: "Long-Term ED Visits (per 1,000 days)",
                  facility: fmt(claimScore(CODE_LT_ED)),
                  national: fmt(nationAvg[COL_LT_ED_AVG]),
                  state:    fmt(stateAvg[COL_LT_ED_AVG]),
                  raw: { f: claimScore(CODE_LT_ED), n: nationAvg[COL_LT_ED_AVG] },
                },
              ].map(({ label, facility, national, state, raw }) => (
                <div key={label} style={{ borderBottom: "1px solid #f0f0f0", padding: "14px 0", display: "flex", alignItems: "baseline", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}>
                  <span style={{ fontSize: 14, color: "#333", minWidth: 260 }}>{label}</span>
                  <span style={{ display: "flex", gap: 28, alignItems: "baseline" }}>
                    <span>
                      <span style={{ fontSize: 10, color: "#aaa", textTransform: "uppercase", letterSpacing: "0.08em", marginRight: 6 }}>Facility</span>
                      <span style={{ ...S.mono, fontWeight: 700, fontSize: 14, color: deltaColor(raw.f, raw.n) }}>{facility}</span>
                    </span>
                    <span>
                      <span style={{ fontSize: 10, color: "#aaa", textTransform: "uppercase", letterSpacing: "0.08em", marginRight: 6 }}>Natl</span>
                      <span style={{ ...S.mono, fontSize: 14, color: "#555" }}>{national}</span>
                    </span>
                    <span>
                      <span style={{ fontSize: 10, color: "#aaa", textTransform: "uppercase", letterSpacing: "0.08em", marginRight: 6 }}>State</span>
                      <span style={{ ...S.mono, fontSize: 14, color: "#555" }}>{state}</span>
                    </span>
                  </span>
                </div>
              ))}
            </div>

            <hr style={S.rule} />

            {/* Full table */}
            <div style={{ marginBottom: 40 }}>
              <p style={S.sectionEyebrow}>Full Report Data</p>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <tbody>
                  {getAllRows().map(([label, value], i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #f0f0f0", background: i % 2 === 0 ? "#fafafa" : "#fff" }}>
                      <td style={{ padding: "9px 12px", color: "#555", fontWeight: 500, width: "55%" }}>{label}</td>
                      <td style={{ padding: "9px 12px", color: "#111", ...S.mono }}>{value || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={{ marginTop: 14 }}>
                <a href={`https://www.medicare.gov/care-compare/details/nursing-home/${ccn.trim()}`}
                  target="_blank" rel="noreferrer"
                  style={{ fontSize: 13, color: PINK, fontWeight: 600, textDecoration: "none" }}>
                  View Medicare Care Compare profile →
                </a>
              </div>
            </div>

            {/* Download buttons */}
            <div style={{ display: "flex", gap: 12 }}>
              <button onClick={downloadPDF} style={S.btnPrimary}>Download PDF</button>
              <button onClick={downloadDocx} style={S.btnSecondary}>Download Word Doc</button>
            </div>
          </>
        )}

        {/* Footer */}
        <div style={{ marginTop: 64, borderTop: "1px solid #f0f0f0", paddingTop: 20, fontSize: 11, color: "#bbb", letterSpacing: "0.04em" }}>
          Built by Angela Wu &nbsp;·&nbsp; Medelite Technical Assessment &nbsp;·&nbsp; {new Date().getFullYear()}
        </div>

      </div>
    </main>
  );
}