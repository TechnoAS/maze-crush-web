/**
 * WhatsApp Daily Production Report Parser
 * TypeScript port of parse_whatsapp_report.py
 */

const FLAGS = "i"; // case-insensitive (dotAll handled via [\s\S] where needed)

function extractVal(pattern: string, text: string, defaultVal = 0.0, flags = "is"): number {
  try {
    const re = new RegExp(pattern, flags);
    const m = text.match(re);
    if (m && m[1] !== undefined) {
      const raw = m[1].trim().replace(/^-+/, "").replace(/,/g, "").trim();
      const v = parseFloat(raw);
      if (!isNaN(v)) return v;
    }
  } catch (_) {}
  return defaultVal;
}

function extractStr(pattern: string, text: string, defaultVal = "", flags = "is"): string {
  try {
    const re = new RegExp(pattern, flags);
    const m = text.match(re);
    if (m && m[1] !== undefined) return m[1].trim();
  } catch (_) {}
  return defaultVal;
}

export interface ParsedReport {
  [key: string]: number | string;
}

export function parseReport(text: string): ParsedReport {
  const data: ParsedReport = {};

  // ── DATE ──────────────────────────────────────────────────────────────────
  data["Date"] = extractStr(String.raw`(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})`, text, "", "i");

  // ══ A. CRUSHING & PRODUCTION ══════════════════════════════════════════════
  data["Maize Crushing (MT)"] = extractVal(String.raw`Maize\s+crushing\s*=\s*([\d,]+\.?\d*)\s*mt`, text);
  data["SO2 Water Usage (m3)"] = extractVal(String.raw`[Ss][Oo]2?\s+water\s+usage\s*[:\-=]\s*([\d,]+\.?\d*)\s*m`, text);
  // Accept GRIND =1.00 or GRIND =1.00% — tolerate missing percent sign
  data["SO2 Water Usage Ratio (m3/MT)"] = extractVal(String.raw`(?:GRIND|grind)\s*[=:\-]\s*([\d,]+\.?\d*)\s*%?`, text);
  data["SO2 PPM Steeped Maize Avg"] = extractVal(String.raw`[Ss][Oo]2?\s+PPM\s+of\s+Steeped\s+Maize\s*[-–:=\s]+Avg\s+([\d,]+\.?\d*)\s*PPM`, text);
  data["Total CSL Generated (m3)"] = extractVal(String.raw`Total\s+CSL\s+Generated\s*[:\-=]\s*([\d,]+\.?\d*)\s*m`, text);
  data["Total Effluent (m3)"] = extractVal(String.raw`Total\s+Effluent\s*[=:\-]\s*([\d,]+\.?\d*)\s*m3`, text);
  // Prefer explicit "per MT" values (e.g. "per MT of Grind 3.06m3")
  data["Effluent Per MT (m3/MT)"] = extractVal(String.raw`per\s+MT(?:\s+of\s+Grind)?[^\d]*([\d,]+\.?\d*)\s*m3`, text, 0.0, "is");
  // Fallback to older pattern (Total Effluent = 920 m3 /300 m3) — compute ratio when denominator present
  if (!data["Effluent Per MT (m3/MT)"] || data["Effluent Per MT (m3/MT)"] === 0) {
    const effMatch = text.match(/Total\s+Effluent\s*[=:\-]\s*([\d,]+\.?\d*)\s*m3\s*\/?\s*([\d,]+\.?\d*)?\s*m3?/i);
    if (effMatch) {
      const totalEff = parseFloat(effMatch[1].replace(/,/g, ""));
      const denom = effMatch[2] ? parseFloat(effMatch[2].replace(/,/g, "")) : NaN;
      if (!isNaN(totalEff) && !isNaN(denom) && denom !== 0) {
        data["Effluent Per MT (m3/MT)"] = +(totalEff / denom).toFixed(3);
      }
    }
  }
  data["CSL to ETP (m3)"] = extractVal(String.raw`CSL\s+to\s+ETP\s*[:\-=]\s*([\d,]+\.?\d*)\s*m`, text);
  data["CSL Plant (m3)"] = extractVal(String.raw`CSL\s+PLANT\s*[:\-=]\s*([\d,]+\.?\d*)\s*m`, text);
  data["HCSL to Gluten (m3)"] = extractVal(String.raw`HCSL\s+to\s+Gluten\s*[:\-=]\s*([\d,]+\.?\d*)\s*m`, text);
  data["HCSL to Fiber (m3)"] = extractVal(String.raw`HCSL\s+to\s+Fiber\s*[:\-=]\s*([\d,]+\.?\d*)\s*m`, text);

  // ── DRYER PRODUCTION ────────────────────────────────────────────────────
  for (let d = 1; d <= 10; d++) {
    data[`Dryer ${d} (MT)`] = extractVal(`Dryer\\s+${d}\\s*[:\\-=]\\s*([\\d,]+\\.?\\d*)\\s*MT`, text);
  }
  data["Total Dryer Production (MT)"] = extractVal(String.raw`Total\s+Dryer\s+Production\s*[:\-=]\s*([\d,]+\.?\d*)\s*MT`, text);

  // ── LGP / OUTPUT ────────────────────────────────────────────────────────
  data["Slurry to LGP (MT)"] = extractVal(String.raw`Slurry\s+to\s+LGP\s*[=:\-]\s*([\d,]+\.?\d*)\s*MT`, text);
  for (let lg = 1; lg <= 10; lg++) {
    data[`LG-${lg} (MT)`] = extractVal(`LG[-\\s]*${lg}\\s*[=:\\-]\\s*([\\d,]+\\.?\\d*)\\s*MT`, text);
  }
  data["HMCS (MT)"] = extractVal(String.raw`HMCS\s*[=:\-]\s*([\d,]+\.?\d*)\s*MT`, text);
  data["MD (MT)"] = extractVal(String.raw`\bMD\s*[=:\-]\s*([\d,]+\.?\d*)\s*MT`, text);
  data["Total Output (MT)"] = extractVal(String.raw`Total\s+Output\s*[:\-=]\s*([\d,]+\.?\d*)\s*MT`, text);
  data["Recovery (%)"] = extractVal(String.raw`Recovery\s*%?\s*[:\-=]\s*([\d,]+\.?\d*)\s*%`, text);
  data["Average Recovery (%)"] = extractVal(String.raw`Average\s*[:\-=]\s*([\d,]+\.?\d*)\s*%`, text);

  // ── BYPRODUCTS ─────────────────────────────────────────────────────────
  data["Germ (MT)"] = extractVal(String.raw`Germ\s*[:\-=\s]+\s*([\d,]+\.?\d*)\s*MT`, text);
  data["Germ (%)"] = extractVal(String.raw`Germ\s*[:\-=\s]+\s*[\d,]+\.?\d*\s*MT[-\s]+\(?([\d,]+\.?\d*)\s*%`, text);
  data["Gluten (MT)"] = extractVal(String.raw`Gluten\s*[:\-=]\s*([\d,]+\.?\d*)\s*MT`, text);
  data["Gluten (%)"] = extractVal(String.raw`Gluten\s*[:\-=]\s*[\d,]+\.?\d*\s*MT\s*\(?\s*([\d,]+\.?\d*)\s*%`, text);
  data["Gluten HCSL (m3)"] = extractVal(String.raw`Gluten\s*[:\-=][^\n]*Hcsl\s+([\d,]+\.?\d*)\s*m`, text);
  data["Dry Husk (MT)"] = extractVal(String.raw`Dry\s+Husk\s*[:\-=]\s*([\d,]+\.?\d*)\s*MT`, text);
  data["Dry Husk (%)"] = extractVal(String.raw`Dry\s+Husk\s*[:\-=]\s*[\d,]+\.?\d*\s*MT\s*\(?\s*([\d,]+\.?\d*)\s*%`, text);
  data["Dry Husk HCSL (m3)"] = extractVal(String.raw`Dry\s+Husk[^\n]*Hcsl\s+([\d,]+\.?\d*)\s*m`, text);
  data["Wet Husk (MT)"] = extractVal(String.raw`Wet\s+Husk\s*[:\-=]\s*([\d,]+\.?\d*)\s*MT`, text);

  // ══ B. BOILER & TURBINE ════════════════════════════════════════════════
  data["Boiler Load (MT)"] = extractVal(String.raw`Boiler\s+Load\s*[:\-=\s]+\s*([\d,]+\.?\d*)\s*MT`, text);
  data["Steam Generated (MT)"] = extractVal(String.raw`Steam\s+Generated\s*[:\-=]\s*([\d,\s]+\.?\s*\d*)\s*MT`, text.replace(/ \./g, "."));
  data["Turbine Load (KWH)"] = extractVal(String.raw`Turbine\s+Load\s*[:\-=\s]+\s*([\d,]+\.?\d*)\s*KWH`, text);
  data["Inlet Steam (TPH)"] = extractVal(String.raw`Inlet\s+Steam\s*[:\-=\s]+\s*([\d,]+\.?\d*)\s*TPH`, text);
  data["Outlet Steam (TPH)"] = extractVal(String.raw`Outlet\s+Steam\s*[:\-=\s]+\s*([\d,]+\.?\d*)\s*TPH`, text);
  data["1st Exhaust Pressure (kg/cm2)"] = extractVal(String.raw`1st\s+Exhaust\s+Pressure\s*[:\-=\s]+\s*([\d,]+\.?\d*)\s*kg`, text);
  data["2nd Exhaust Pressure (kg/cm2)"] = extractVal(String.raw`2nd\s+Exhaust\s+Pressure\s*[:\-=\s]+\s*([\d,]+\.?\d*)\s*kg`, text);
  data["Bearing Temp (C)"] = extractVal(String.raw`Bearing\s+Temp\s*[:\-=\s]+\s*([\d,]+\.?\d*)\s*[°]?C`, text);

  // ── POWER USED ─────────────────────────────────────────────────────────
  data["MSEB Power (KWH)"] = extractVal(String.raw`MSEB\s*[:\-=]\s*([\d,]+\.?\d*)\s*KWH`, text);
  data["Gas Engine Power (KWH)"] = extractVal(String.raw`Gas\s+Engine\s*[:\-=]\s*([\d,]+\.?\d*)\s*KWH`, text);
  data["Steam Turbine Power (KWH)"] = extractVal(String.raw`Steam\s+Turbine\s*[:\-=]\s*([\d,]+\.?\d*)\s*KWH`, text);
  data["Total Power (KWH)"] = extractVal(String.raw`Total\s+Power\s*[:\-=]\s*([\d,]+\.?\d*)\s*KWH`, text);
  data["Power Per Ton (Units/Ton)"] = extractVal(String.raw`Total\s+Power[^\n]*\((\d+)\s+units\s*\/\s*ton`, text);

  // ══ C. FUEL CONSUMPTION ════════════════════════════════════════════════
  data["Coal (MT)"] = extractVal(String.raw`Coal\s*[:\-=]\s*([\d,]+\.?\d*)\s*MT`, text);
  data["Coal Ratio"] = extractVal(String.raw`Coal[^\n]*Ratio\s*[:\-=\s]+\s*([\d,]+\.?\d*)`, text);
  data["Loose Bagasse (MT)"] = extractVal(String.raw`Loose\s+Bagasse\s*[:\-=]\s*([\d,]+\.?\d*)\s*MT`, text);
  data["Loose Bagasse Ratio"] = extractVal(String.raw`Loose\s+Bagasse[^\n]*Ratio\s*[:\-=\s]+\s*([\d,]+\.?\d*)`, text);
  data["Firewood (MT)"] = extractVal(String.raw`Firewood\s*[:\-=]\s*([\d,]+\.?\d*)\s*MT`, text);
  data["Wood Dust (MT)"] = extractVal(String.raw`Wood\s+Dust\s*[:\-=]\s*([\d,]+\.?\d*)\s*MT`, text);
  data["Gas Fuel (MT)"] = extractVal(String.raw`(?:^|\n)\s*Gas\s*[:\-=]\s*([\d,]+\.?\d*)\s*MT`, text, 0.0, "im");

  // ══ D. UTILITIES ════════════════════════════════════════════════════════
  // Tank formats:
  //   Numeric:      "Level Tank: 234"            → number (234)
  //   Full-string:  "Big Tank: 4.5ft +3.00 ft"  → string ("4.5ft +3.00 ft")
  //   Text status:  "Hermes Tank: full"           → string ("full")
  //                 "Process Water Tank: Full"    → string ("Full")
  data["Water Used (m3)"] = extractVal(String.raw`Water\s+Used\s*[:\-=]\s*([\d,]+\.?\d*)\s*m`, text);
  data["Waste Water (ft)"] = extractVal(String.raw`Waste\s+Water\s*[:\-=\s]+\s*([\d,]+\.?\d*)\s*ft`, text);
  // Level Tank — plain numeric
  data["Level Tank"] = extractVal(String.raw`Level\s+Tank\s*[:\-=]\s*([\d]+\.?\d*)`, text);
  // Big Tank — full string until end of line (e.g. "4.5ft +3.00 ft")
  data["Big Tank"] = extractStr(String.raw`Big\s+Tank\s*[:\-=]\s*([^\n]+)`, text);
  // Text-status tanks
  data["Hermes Tank"] = extractStr(String.raw`Hermes\s+Tank\s*[:\-=]\s*([^\n]+)`, text);
  data["Process Water Tank"] = extractStr(String.raw`Process\s+Water\s+Tank\s*[:\-=]\s*([^\n]+)`, text);

  // ── WATER QUALITY ──────────────────────────────────────────────────────
  data["B/F TDS (ppm)"] = extractVal(String.raw`B\/F\s+TDS\s*[:\-=\s]+-?\s*([\d,]+\.?\d*)\s*(?:ppm)?`, text);
  data["M/B SiO2 (ppm)"] = extractVal(String.raw`M\/B\s+Si[Oo][2₂]\s*[:\-=\s]+-?\s*([\d,]+\.?\d*)\s*(?:ppm)?`, text);
  data["RO TDS (ppm)"] = extractVal(String.raw`RO\s+TDS\s*[:\-=\s]+-+\s*([\d,]+\.?\d*)`, text);
  data["B/D TDS (ppm)"] = extractVal(String.raw`B\/D\s+TDS\s*[:\-=\s]+-?\s*([\d,]+\.?\d*)\s*(?:ppm)?`, text);
  data["MIDC TDS (ppm)"] = extractVal(String.raw`MIDC\s+TDS\s*[:\-=\s]+-?\s*([\d,]+\.?\d*)\s*(?:ppm)?`, text);

  // ══ E. MAIZE FILLING ════════════════════════════════════════════════════
  data["Maize From Silo 1 (Bags)"] = extractVal(String.raw`From\s+Silo\s+1\s*[:\-=]\s*([\d,]+\.?\d*)\s*bags`, text);
  data["Maize From Silo 2 (Bags)"] = extractVal(String.raw`From\s+Silo\s+2\s*[:\-=]\s*([\d,]+\.?\d*)\s*bags`, text);
  data["Maize Filled in Vats (Bags)"] = extractVal(String.raw`Maize\s+Filled\s+in\s+Vats\s*[:\-=]\s*([\d,]+\.?\d*)\s*bags`, text);
  data["OB Maize Silo 1 (Bags)"] = extractVal(String.raw`OB\s+Maize\s+in\s+Silo\s+1\s*[:\-=]\s*([\d,]+\.?\d*)\s*bags`, text);
  data["Closing Silo 1 (Bags)"] = extractVal(String.raw`Closing\s+in\s+Silo\s+1[^\n]*=\s*([\d,]+\.?\d*)\s*bags`, text);
  data["OB Maize Silo 2 (Bags)"] = extractVal(String.raw`OB\s+Maize\s+in\s+Si[ol][oa]?\s+no\.?\s*[=:\-]?\s*2\s*[:\-=]\s*([\d,]+\.?\d*)\s*bags`, text);
  data["Closing Silo 2 (Bags)"] = extractVal(String.raw`Closing\s+in\s+Silo\s*2[^\n]*=\s*([\d,]+\.?\d*)\s*bags`, text);
  data["Total Maize Filled (Bags)"] = extractVal(String.raw`Total\s+Maize\s+Filled\s*[:\-=]\s*([\d,]+\.?\d*)\s*bags`, text);
  data["Maize Filling Station (Bags)"] = extractVal(String.raw`Maize\s+Filling\s+Station\s*[:\-=]\s*([\d,]+\.?\d*)\s*bags`, text);

  // ══ F. VAT TABLE ════════════════════════════════════════════════════════
  const vatSectionMatch = text.match(/(?:Water\s+Damage|Vat\s+No|WD\s*%)[\s\S]*?(?=\n[-─=*\s]{3,}|$)/i);
  const vatSection = vatSectionMatch ? vatSectionMatch[0] : text;
  // More tolerant VAT line regex: accept separators like space, dot, '=' and extra punctuation
  const vatLineRe = /^\s*(\d{1,3})\D+?([\d.]+)\s*%[^\d]*([\d.]+)\s*%/gm;

  const parsedVats: Record<number, [number, number]> = {};
  let vm: RegExpExecArray | null;
  while ((vm = vatLineRe.exec(vatSection)) !== null) {
    const vatNo = parseInt(vm[1], 10);
    if (vatNo >= 1 && vatNo <= 999 && !(vatNo in parsedVats)) {
      const wd = parseFloat(vm[2]);
      const ss = parseFloat(vm[3]);
      if (!isNaN(wd) && !isNaN(ss)) parsedVats[vatNo] = [wd, ss];
    }
  }
  for (const vatNo of Object.keys(parsedVats).map(Number).sort((a, b) => a - b)) {
    const [wd, ss] = parsedVats[vatNo];
    data[`Vat ${vatNo} WD (%)`] = wd;
    data[`Vat ${vatNo} SS (%)`] = ss;
  }

  return data;
}

export function toCSVRow(data: ParsedReport): string {
  const headers = Object.keys(data);
  const values = Object.values(data);
  const escape = (v: string | number) => {
    const s = String(v);
    if (s.includes(",") || s.includes('"') || s.includes("\n")) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  };
  return `${headers.map(escape).join(",")}\n${values.map(escape).join(",")}`;
}
