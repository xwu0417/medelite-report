/**
 * @author Angela Wu
 * @project Medelite Facility Assessment Report Generator
 * @date 06/16/2026
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

// ─── Claims measure codes ──────────────────────────────────────────────────
const CODE_STR_HOSP = "521";
const CODE_STR_ED   = "522";
const CODE_LT_HOSP  = "551";
const CODE_LT_ED    = "552";

// ─── State-avg column keys ─────────────────────────────────────────────────
const COL_STR_HOSP_AVG = "percentage_of_short_stay_residents_who_were_rehospitalized__1d02";
const COL_STR_ED_AVG   = "percentage_of_short_stay_residents_who_had_an_outpatient_em_d911";
const COL_LT_HOSP_AVG  = "number_of_hospitalizations_per_1000_longstay_resident_days";
const COL_LT_ED_AVG    = "number_of_outpatient_emergency_department_visits_per_1000_l_de9d";

// ─── Proxy helper — all CMS fetches go through /api/cms to avoid CORS ─────
const CMS_BASE = "https://data.cms.gov/provider-data/api/1/datastore/query";

function cmsUrl(dataset, params) {
  const full = `${CMS_BASE}/${dataset}/0?${params}`;
  return `/api/cms?endpoint=${encodeURIComponent(full)}`;
}

export default function Home() {
  const [ccn, setCcn] = useState("");

  const [facilityData, setFacilityData] = useState(null);
  const [claimsData,   setClaimsData]   = useState([]);
  const [stateAvg,     setStateAvg]     = useState({});
  const [nationAvg,    setNationAvg]    = useState({});

  const [loading,   setLoading]   = useState(false);
  const [error,     setError]     = useState(null);
  const [dataReady, setDataReady] = useState(false);

  // Manual inputs
  const [nameOverride,        setNameOverride]        = useState("");
  const [emr,                 setEmr]                 = useState("");
  const [currentCensus,       setCurrentCensus]       = useState("");
  const [patientType,         setPatientType]         = useState("");
  const [previousCoverage,    setPreviousCoverage]    = useState("");
  const [previousPerformance, setPreviousPerformance] = useState("");
  const [medicalCoverage,     setMedicalCoverage]     = useState("");

  // ─── Fetch ALL data ────────────────────────────────────────────────────────
  async function fetchFacility() {
    if (!ccn.trim()) return;
    setLoading(true);
    setDataReady(false);
    setError(null);

    try {
      // Step 1: provider info (need state before firing avg queries)
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

      // Step 2: claims + state avg + national avg in parallel
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
        claimsRes.json(),
        stateRes.json(),
        nationRes.json(),
      ]);

      const claims = claimsJson.results || [];
      const sAvg   = stateJson.results?.[0]  || {};
      const nAvg   = nationJson.results?.[0] || {};

      // Commit all at once — no partial renders
      setFacilityData(facility);
      setClaimsData(claims);
      setStateAvg(sAvg);
      setNationAvg(nAvg);
      setNameOverride("");
      setDataReady(true);

    } catch (err) {
      console.error(err);
      setError("Something went wrong fetching data. Check your network and try again.");
    }

    setLoading(false);
  }

  // ─── Helpers ──────────────────────────────────────────────────────────────
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

  // ─── 12 metric rows ───────────────────────────────────────────────────────
  function getMetricRows() {
    return [
      ["Short Term Hospitalization",           fmt(claimScore(CODE_STR_HOSP), true)],
      ["STR National Avg. for Hospitalization",fmt(nationAvg[COL_STR_HOSP_AVG], true)],
      ["STR State Avg. for Hospitalization",   fmt(stateAvg[COL_STR_HOSP_AVG],  true)],
      ["STR ED Visit",                         fmt(claimScore(CODE_STR_ED), true)],
      ["STR ED Visits National Avg.",          fmt(nationAvg[COL_STR_ED_AVG], true)],
      ["STR ED Visits State Avg.",             fmt(stateAvg[COL_STR_ED_AVG],  true)],
      ["LT Hospitalization",                   fmt(claimScore(CODE_LT_HOSP))],
      ["LT National Avg. for Hospitalization", fmt(nationAvg[COL_LT_HOSP_AVG])],
      ["LT State Avg. for Hospitalization",    fmt(stateAvg[COL_LT_HOSP_AVG])],
      ["ED Visit",                             fmt(claimScore(CODE_LT_ED))],
      ["LT ED Visits National Avg.",           fmt(nationAvg[COL_LT_ED_AVG])],
      ["LT ED Visits State Avg.",              fmt(stateAvg[COL_LT_ED_AVG])],
    ];
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
      ...getMetricRows(),
    ];
  }

  // ─── PDF Download ──────────────────────────────────────────────────────────
  async function downloadPDF() {
    if (!facilityData || !dataReady) return;
    const doc = new jsPDF();

    try {
      const logoUrl  = window.location.origin + "/infinite_logo.png";
      const response = await fetch(logoUrl);
      const blob     = await response.blob();
      const base64   = await new Promise((res) => {
        const reader = new FileReader();
        reader.onloadend = () => res(reader.result);
        reader.readAsDataURL(blob);
      });
      doc.addImage(base64, "PNG", 65, 8, 80, 22);
    } catch {
      console.warn("Logo not loaded; skipping.");
    }

    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(0, 0, 0);
    doc.text("FACILITY ASSESSMENT SNAPSHOT", 105, 38, { align: "center" });
    doc.setFontSize(11);
    doc.text(facilityData.state || "", 105, 45, { align: "center" });

    const startY   = 52;
    const leftX    = 20;
    const colSplit = 100;
    const rightX   = 190;
    const rowH     = 8.5;
    const pageH    = 280;
    const allRows  = getAllRows();
    let y = startY;

    allRows.forEach(([label, value], i) => {
      if (y + rowH > pageH) { doc.addPage(); y = 20; }

      if (i % 2 === 0) {
        doc.setFillColor(245, 245, 245);
        doc.rect(leftX, y, rightX - leftX, rowH, "F");
      }

      doc.setDrawColor(180, 180, 180);
      doc.rect(leftX, y, rightX - leftX, rowH, "S");
      doc.line(colSplit, y, colSplit, y + rowH);

      doc.setFontSize(8.5);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(0, 0, 0);
      doc.text(doc.splitTextToSize(label, colSplit - leftX - 3)[0], leftX + 2, y + 5.5);
      doc.setFont("helvetica", "normal");
      doc.text(String(value), colSplit + 3, y + 5.5);

      y += rowH;
    });

    doc.setFontSize(9);
    doc.setTextColor(6, 69, 173);
    doc.textWithLink(
      "View official Medicare Care Compare profile",
      leftX, y + 10,
      { url: `https://www.medicare.gov/care-compare/details/nursing-home/${ccn.trim()}` }
    );

    doc.save(`Facility_Assessment_${displayName.replace(/\s+/g, "_")}.pdf`);
  }

  // ─── DOCX Download ────────────────────────────────────────────────────────
  async function downloadDocx() {
    if (!facilityData || !dataReady) return;

    let logoBuffer = null;
    try {
      const res = await fetch(window.location.origin + "/infinite_logo.png");
      logoBuffer = await res.arrayBuffer();
    } catch {
      console.warn("Logo not loaded");
    }

    const allRows = getAllRows();

    const tableRows = allRows.map(([label, value], i) =>
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            shading: i % 2 === 0 ? { fill: "F5F5F5" } : undefined,
            children: [new Paragraph({
              children: [new TextRun({ text: label, bold: true, size: 20 })]
            })],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            shading: i % 2 === 0 ? { fill: "F5F5F5" } : undefined,
            children: [new Paragraph({
              children: [new TextRun({ text: String(value), size: 20 })]
            })],
          }),
        ],
      })
    );

    const children = [];

    if (logoBuffer) {
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({
          data: logoBuffer,
          transformation: { width: 200, height: 55 },
          type: "png",
        })],
      }));
    }

    children.push(
      new Paragraph({
        text: "FACILITY ASSESSMENT SNAPSHOT",
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        text: facilityData.state || "",
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
      }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: tableRows,
      }),
      new Paragraph({ text: "" }),
      new Paragraph({
        children: [new TextRun({
          text: "View official Medicare Care Compare profile",
          color: "0645AD",
          underline: {},
          size: 18,
        })],
      })
    );

    const doc = new Document({ sections: [{ children }] });
    const blob = await Packer.toBlob(doc);
    saveAs(blob, `Facility_Assessment_${displayName.replace(/\s+/g, "_")}.docx`);
  }

  // ─── Styles ───────────────────────────────────────────────────────────────
  const box        = { background: "#fff", border: "1px solid #ddd", borderRadius: 12, padding: 24, marginBottom: 20 };
  const labelStyle = { fontWeight: 600, fontSize: 13, color: "#555", display: "block", marginBottom: 4 };
  const inputStyle = { padding: "9px 12px", fontSize: 14, width: "100%", borderRadius: 8, border: "1px solid #ccc", boxSizing: "border-box" };
  const btnStyle   = { padding: "11px 24px", background: "#c0218e", color: "#fff", border: "none", borderRadius: 8, fontSize: 15, fontWeight: 700, cursor: "pointer" };

  // ─── Render ───────────────────────────────────────────────────────────────
  return (
    <main style={{ fontFamily: "sans-serif", maxWidth: 700, margin: "40px auto", padding: "0 20px 60px" }}>

      <div style={{ textAlign: "center", marginBottom: 32 }}>
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src="/infinite_logo.png" alt="INFINITE Managed by MEDELITE" style={{ width: 260, height: "auto" }} />
      </div>

      <h2 style={{ textAlign: "center", marginBottom: 24 }}>FACILITY ASSESSMENT SNAPSHOT</h2>

      {/* CCN Lookup */}
      <div style={box}>
        <label style={labelStyle}>Enter CCN (CMS Certification Number)</label>
        <input
          type="text"
          placeholder="e.g. 686123"
          value={ccn}
          onChange={(e) => setCcn(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && fetchFacility()}
          style={inputStyle}
        />
        <button onClick={fetchFacility} disabled={loading} style={{ ...btnStyle, marginTop: 12 }}>
          {loading ? "Loading..." : "Fetch Facility Data"}
        </button>
        {error && <p style={{ color: "red", marginTop: 12 }}>{error}</p>}
      </div>

      {loading && (
        <div style={{ ...box, textAlign: "center", color: "#888", padding: 40 }}>
          <div style={{ fontSize: 32, marginBottom: 12 }}>⏳</div>
          <p style={{ margin: 0 }}>Fetching facility data, ratings, and averages…</p>
        </div>
      )}

      {dataReady && facilityData && (
        <>
          <div style={box}>
            <h3 style={{ marginTop: 0 }}>Facility Details</h3>

            <label style={labelStyle}>Facility Name Override (optional)</label>
            <input
              type="text"
              placeholder={facilityData.provider_name}
              value={nameOverride}
              onChange={(e) => setNameOverride(e.target.value)}
              style={{ ...inputStyle, marginBottom: 4 }}
            />
            <small style={{ color: "#888", display: "block", marginBottom: 12 }}>
              Leave blank to use CMS name: &quot;{facilityData.provider_name}&quot;
            </small>

            <label style={labelStyle}>EMR System</label>
            <input type="text" placeholder="e.g. PCC" value={emr} onChange={(e) => setEmr(e.target.value)} style={{ ...inputStyle, marginBottom: 12 }} />

            <label style={labelStyle}>Current Census</label>
            <input type="number" placeholder="e.g. 112" value={currentCensus} onChange={(e) => setCurrentCensus(e.target.value)} style={{ ...inputStyle, marginBottom: 12 }} />

            <label style={labelStyle}>Type of Patient</label>
            <select value={patientType} onChange={(e) => setPatientType(e.target.value)} style={{ ...inputStyle, marginBottom: 12 }}>
              <option value="">Select...</option>
              <option>Long-term &amp; Short-term</option>
              <option>Long-term</option>
              <option>Short-term</option>
            </select>

            <label style={labelStyle}>Previous Coverage from Medelite</label>
            <select value={previousCoverage} onChange={(e) => setPreviousCoverage(e.target.value)} style={{ ...inputStyle, marginBottom: 12 }}>
              <option value="">Select...</option>
              <option>Yes</option>
              <option>No</option>
            </select>

            <label style={labelStyle}>Previous Provider Performance (patients/day)</label>
            <input type="text" placeholder="e.g. 30" value={previousPerformance} onChange={(e) => setPreviousPerformance(e.target.value)} style={{ ...inputStyle, marginBottom: 12 }} />

            <label style={labelStyle}>Medical Coverage</label>
            <input type="text" placeholder="e.g. Optometry, PCP, Podiatry" value={medicalCoverage} onChange={(e) => setMedicalCoverage(e.target.value)} style={inputStyle} />
          </div>

          {/* Report Preview */}
          <div style={box}>
            <h3 style={{ marginTop: 0 }}>Report Preview</h3>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 14 }}>
              <tbody>
                {getAllRows().map(([label, value], i) => (
                  <tr key={i} style={{ background: i % 2 === 0 ? "#f9f9f9" : "#fff" }}>
                    <td style={{ padding: "8px 10px", border: "1px solid #ddd", fontWeight: 600, width: "55%" }}>{label}</td>
                    <td style={{ padding: "8px 10px", border: "1px solid #ddd" }}>{value || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            <p style={{ marginTop: 12, fontSize: 13 }}>
              <a
                href={`https://www.medicare.gov/care-compare/details/nursing-home/${ccn.trim()}`}
                target="_blank"
                rel="noreferrer"
                style={{ color: "#0645ad" }}
              >
                View official Medicare Care Compare profile
              </a>
            </p>
          </div>

          {/* Download Buttons */}
          <div style={{ textAlign: "center", display: "flex", gap: 16, justifyContent: "center" }}>
            <button
              onClick={downloadPDF}
              style={{ padding: "14px 36px", background: "#c0218e", color: "#fff", border: "none", borderRadius: 10, fontSize: 16, fontWeight: 700, cursor: "pointer" }}
            >
              Download PDF
            </button>
            <button
              onClick={downloadDocx}
              style={{ padding: "14px 36px", background: "#1a56a0", color: "#fff", border: "none", borderRadius: 10, fontSize: 16, fontWeight: 700, cursor: "pointer" }}
            >
              Download Word Doc
            </button>
          </div>
        </>
      )}
    </main>
  );
}