const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");
const readlineSync = require("readline-sync"); // Diperlukan untuk input sheet

// --- Fungsi Ekstraksi GSM (dari skrip kedua, lebih komprehensif) ---
function extractGsmValue(desc) {
  if (typeof desc !== "string") return "N/A";

  let match;
  // Prioritized search for GSM patterns
  // 1. G/M2 (most specific)
  match = desc.match(/(\d[\d.,]*)\s*G\/M2/i);
  if (match) return match[1].replace(",", ".");

  // 2. GSM
  match = desc.match(/(\d[\d.,]*)\s*GSM/i);
  if (match) return match[1].replace(",", ".");

  // 3. GR/M2
  match = desc.match(/(\d[\d.,]*)\s*GR\/M2/i);
  if (match) return match[1].replace(",", ".");

  // 4. XXG (e.g., 30G KIMLON, WEIGHT 15G)
  match = desc.match(/\b(WEIGHT|AVERAGE WEIGHT|BASIS WEIGHT)\s*:?\s*(\d[\d.,]*)\s*G\b/i);
  if (match) return match[2].replace(",", "."); // Adjusted index for capturing group
  match = desc.match(/(\d[\d.,]*)\s*G\s+(KIMLON|TYPE)/i);
  if (match) return match[1].replace(",", ".");
  // Regex to avoid GSM, G/M2, GR/M2 but capture things like '15G' or '20 G'
  match = desc.match(/\b(\d[\d.,]*)\s*G\b(?!\s*SM)(?!\s*\/M2)(?!\s*R\/M2)(?!\s*[A-DF-LN-QS-Z])/i);
  if (match) return match[1].replace(",", ".");

  // 5. GR/YD
  match = desc.match(/(\d[\d.,]*)\s*GR\/YD/i);
  if (match) return match[1].replace(",", ".");

  return "N/A";
}

// --- Fungsi Ekstraksi Width (dari skrip pertama) ---
function extractWidthValue(desc) {
  if (typeof desc !== "string") return "N/A";

  let match;
  match = desc.match(/(\d+[\d.,]*)\s*(inch|inches|\")/i);
  if (match) {
    const inches = parseFloat(match[1].replace(",", "."));
    return (inches * 2.54).toFixed(2); // Konversi ke cm
  }

  match = desc.match(/(\d+[\d.,]*)\s*cm/i);
  if (match) return parseFloat(match[1].replace(",", ".")).toFixed(2); // Pastikan format desimal konsisten

  match = desc.match(/(\d+[\d.,]*)\s*mm/i);
  if (match) {
    const mm = parseFloat(match[1].replace(",", "."));
    return (mm / 10).toFixed(2); // Konversi ke cm
  }

  return "N/A";
}

// --- Fungsi Ekstraksi ITEM (dari skrip kedua) ---
function extractItemTypeAndPattern(desc) {
  if (typeof desc !== "string") return { item: "N/A", pattern: "" };
  const descUpper = desc.toUpperCase();

  // Priority 1: AT (Air Thru)
  const atPattern = /\bAIR\s*THRU(?:\s*NONWOVEN)?\b|\bAIR\s*THROUGH\b|\bAIRTHRU\b/i;
  if (atPattern.test(descUpper)) {
    const matchedPattern = descUpper.match(atPattern);
    return { item: "AT", pattern: matchedPattern ? matchedPattern[0] : "" };
  }

  // Priority 2: SMS (Spunbond-Meltblown-Spunbond) and its variants (e.g., SMMS, SSMMMS)
  const smsVariants = /\bS[SM]*M[SM]*S*\b/i; // SMM, SMS, SSMMS, SMMS, SMMMS, SSMMMS etc.
  if (smsVariants.test(descUpper)) {
    const match = descUpper.match(smsVariants);
    if (match && match[0].includes("M") && match[0].includes("S")) {
      if (/\bSMS\sNON-WOVEN\sFABRIC\b/i.test(descUpper) || /\bSMS\sNONWOVEN\b/i.test(descUpper) || /\bSPUNMELT\s\(SMS\)\b/i.test(descUpper)) {
        return { item: "SMS", pattern: "SMS" };
      }
      return { item: "SMS", pattern: match[0] };
    }
  }
  if (/\bSMS\b/i.test(descUpper) && !/\bS[SM]*M[SM]*S*\b/.test(descUpper.replace(/\bSMS\b/i, ""))) {
    return { item: "SMS", pattern: "SMS" };
  }

  // Priority 3: SB (Spunbond - only S, no M)
  const sbPatternSpecific = /\bS(SSSBS|S{2,})\b/i;
  const sbPatternGeneral = /\b(SPUNBOND|3S|HO\sSSS)\b/i;

  if (sbPatternSpecific.test(descUpper)) {
    const matchedPattern = descUpper.match(sbPatternSpecific);
    if (matchedPattern && !matchedPattern[0].includes("M")) {
      return { item: "SB", pattern: matchedPattern[0] };
    }
  }
  if (sbPatternGeneral.test(descUpper)) {
    if (!/\bS[SM]*M[SM]*S*\b/.test(descUpper) && !/\bSMS\b/.test(descUpper)) {
      const matchedPattern = descUpper.match(sbPatternGeneral);
      if (matchedPattern && !matchedPattern[0].includes("M")) {
        return { item: "SB", pattern: matchedPattern[0] };
      }
    }
  }
  if (/\bSPUNBOND\b/i.test(descUpper) && !(/\bS[SM]*M[SM]*S*\b/i.test(descUpper) || /\bSMS\b/i.test(descUpper))) {
    return { item: "SB", pattern: "SPUNBOND" };
  }

  // Priority 4: MB (Meltblown)
  const mbPattern = /\bMELTBLOWN\b|\bMELT\sBLOWN\b/i;
  if (mbPattern.test(descUpper)) {
    const matchedPattern = descUpper.match(mbPattern);
    return { item: "MB", pattern: matchedPattern ? matchedPattern[0] : "" };
  }

  // Priority 5: Nonwoven (generic)
  const nonwovenPattern = /NON-WOVEN(?:\sFABRIC)?|NON\sWOVEN(?:\sFABRIC)?|NONWOVEN(?:\sFABRIC)?/i;
  if (nonwovenPattern.test(descUpper)) {
    const matchedPattern = descUpper.match(nonwovenPattern);
    return { item: "Nonwoven", pattern: matchedPattern ? matchedPattern[0] : "" };
  }

  return { item: "N/A", pattern: "" };
}

// --- Fungsi Ekstraksi ADD ON (dari skrip kedua) ---
function extractAddOnFromList(desc) {
  if (typeof desc !== "string") return "-";
  const descUpper = desc.toUpperCase();
  let foundAddOns = new Set();

  const specificCoatings = ["PU Coated", "PVC Coated", "PE Coated", "HDPE Coated", "Thermoplastic Coated"];

  const addOnDefinitions = [
    { keywords: [/\bPU\sCOATED\b/i, /\bPOLYURETHANE\sCOATED\b/i], output: "PU Coated" },
    { keywords: [/\bPVC\sCOATED\b/i, /\bVINYL\sCHLORIDE\s\(PVC\sPLASTIC\)\b/i], output: "PVC Coated" },
    { keywords: [/\bPE\sCOATED\b/i], output: "PE Coated" },
    { keywords: [/\bHDPE\s(?:GLUE\sON\sTHE\sSURFACE|COATED)\b/i], output: "HDPE Coated" },
    { keywords: [/\bTHERMOPLASTIC\s(?:NYLON\sPA|ADHESIVE|COATED)\b/i], output: "Thermoplastic Coated" },
    { keywords: [/\bLAMINATED\b/i, /\bMULTI-LAYER\sLAMINATED\b/i], output: "Laminated", negations: [/\bNOT\sLAMINATED\b/i, /\bUNLAMINATED\b/i] },
    { keywords: [/\bCOATED\b/i, /\bSURFACE\sCOATED\b/i], output: "Coated", negations: [/\bNOT\sCOATED\b/i, /\bUNCOATED\b/i, /\bNOT\sCOATED\sWITH\sGLUE\b/i] },
    { keywords: [/\bIMPREGNATED\b/i, /\bCHEMICALLY\sIMPREGNATED\b/i, /\bSOAKED\b/i], output: "Impregnated", negations: [/\bNOT\sIMPREGNATED\b/i, /\bUNIMPREGNATED\b/i] },
    { keywords: [/\bPERFORATED\b/i, /\bPUNCHED\b/i, /\bNEEDLE-PUNCHED\b/i, /\bNEEDLE\sPUNCHED\b/i], output: "Perforated" },
    { keywords: [/\bEMBOSSED\b/i, /\bTEXTURED\b/i, /\b3D\sSINGLE\sPEARL\sEMBOSSING\b/i, /\bSINGLE\sPEARL\sEMBOSSING\b/i, /\bPEARL\sEMBOSSING\b/i], output: "Embossed" },
    { keywords: [/\bULTRASONIC\sSEALED\b/i], output: "Ultrasonic Sealed" },
    { keywords: [/\bHEAT\sSEALED\b/i, /\bHOT\sMELT\b/i], output: "Heat Sealed" },
    { keywords: [/\bREINFORCED\b/i], output: "Reinforced" },
    { keywords: [/\bPRESSED\b/i, /\bCOMPRESSED\b/i], output: "Pressed" },
    { keywords: [/\bNON-SLIP\b/i, /\bANTI-SLIP\b/i, /\bNON-STICK\b/i], output: "Non-Slip" },
    { keywords: [/\bANTISTATIC\b/i, /\bANTI-STATIC\b/i, /\bESD\b/i, /\bELECTROSTATIC\sFILTER\b/i], output: "Antistatic" },
    { keywords: [/\bHILOFT\b/i, /\bHIGH\sLOFT\b/i], output: "High Loft" },
    { keywords: [/\bLOW\sLOFT\b/i], output: "Low Loft" },
    { keywords: [/\bBREATHABLE\b/i, /\bAIR\sPERMEABLE\b/i, /\bFULL\sBREATHABLE\b/i], output: "Breathable" },
    { keywords: [/\bNON-BREATHABLE\b/i], output: "Non-Breathable" },
    { keywords: [/\bHYDROPHILIC\b/i, /\bLEG\sHI\b/i, /\bTOP\sHI\b/i, /\bCARRIER\sHI\b/i], output: "Hydrophilic" },
    { keywords: [/\bHYDROPHOBIC\b/i, /\bLEG\sHO\b/i, /\bEAR\sHO\b/i, /\bST\sHO\b/i], output: "Hydrophobic" },
    { keywords: [/\bNON-ABSORBENT\b/i], output: "Non-Absorbent" },
    { keywords: [/\bANTIMICROBIAL\b/i, /\bANTI-MICROBIAL\b/i], output: "Antimicrobial" },
    { keywords: [/\bANTIBACTERIAL\b/i, /\bANTI-BACTERIAL\b/i], output: "Antibacterial" },
    { keywords: [/\bANTIVIRAL\b/i, /\bANTI-VIRAL\b/i], output: "Antiviral" },
    { keywords: [/\bFLAME\sRETARDANT\b/i, /\bFIRE\sRESISTANT\b/i], output: "Flame Retardant" },
    { keywords: [/\bUV\sSTABILIZED\b/i, /\bUV\sRESISTANT\b/i], output: "UV Stabilized" },
    { keywords: [/\bOIL\sABSORBENT\b/i], output: "Oil Absorbent" },
    { keywords: [/\bOIL\sREPELLENT\b/i], output: "Oil Repellent" },
    { keywords: [/\bCHEMICAL\sRESISTANT\b/i, /\bALCOHOL-RESISTANT\b/i], output: "Chemical Resistant" },
    { keywords: [/\bANTIFUNGAL\b/i, /\bMOLD\sRESISTANT\b/i], output: "Antifungal" },
    { keywords: [/\bODOR\sCONTROL\b/i, /\bDEODORIZING\b/i], output: "Odor Control" },
    { keywords: [/\bANTI-MILDEW\b/i], output: "Anti-Mildew" },
    { keywords: [/\bCONDUCTIVE\sFABRIC\b/i], output: "Conductive" },
    { keywords: [/\bEXTRA\sSOFT\b/i, /\bSUPER\sSOFT\b/i, /\bULTRA\sSOFT\b/i], output: "Extra Soft" },
    { keywords: [/\bCOTTON\sSOFT\b/i], output: "Cotton Soft" },
    { keywords: [/\bSOFT\b/i], output: "Soft" },
    { keywords: [/\bSMOOTH\b/i], output: "Smooth" },
    { keywords: [/\bSILKY\sFEEL\b/i], output: "Silky Feel" },
    { keywords: [/\bMATTE\b/i], output: "Matte" },
    { keywords: [/\bGLOSSY\b/i, /\bSHINY\b/i], output: "Glossy" },
    { keywords: [/\bSTIFF\b/i, /\bFIRM\b/i, /\bHARD\b/i], output: "Stiff" },
    { keywords: [/\bANTI-WRINKLE\b/i, /\bWRINKLE\sRESISTANT\b/i], output: "Anti-Wrinkle" },
    { keywords: [/\bFLEECE-LIKE\b/i], output: "Fleece-Like" },
    { keywords: [/\bVELVETY\b/i, /\bSUPERFINE\sVELVET\b/i], output: "Velvety" },
    { keywords: [/\bPLUSH\b/i], output: "Plush" },
    { keywords: [/\bDUST-FREE\b/i, /\bANTI\sDUST\b/i], output: "Dust-Free" },
    { keywords: [/\bLOW\sLINT\b/i], output: "Low Lint" },
    { keywords: [/\bANTI-PILLING\b/i], output: "Anti-Pilling" },
    { keywords: [/\bANTI-STRETCH\b/i], output: "Anti-Stretch" },
    { keywords: [/\bPRINTED\b/i, /\bPATTERNED\b/i], output: "Printed", negations: [/\bUNPRINTED\b/i] },
    { keywords: [/\bTWO-TONE\b/i, /\bBICOLOR\b/i], output: "Two-Tone" },
    { keywords: [/\bREFLECTIVE\b/i], output: "Reflective" },
    { keywords: [/\bFLUORESCENT\b/i], output: "Fluorescent" },
    { keywords: [/\bDYED\b/i], output: "Dyed", negations: [/\bUNDYED\b/i, /\bUNBLEACHED,\sUNDYED\b/i] },
    { keywords: [/\bCOLORED\b/i, /\bCOLOUR\b/i], output: "Colored", negations: [/\bUNCOLORED\b/i] },
    { keywords: [/\bMEDICAL\sGRADE\b/i, /\bMEDICAL\sUSE\b/i, /\bAMMI\sLEVEL\s\d+\b/i], output: "Medical Grade" },
    { keywords: [/\bFOOD\sGRADE\b/i], output: "Food Grade" },
    { keywords: [/\bECO-FRIENDLY\b/i, /\bRECYCLED\b/i, /\bREC\sPOLYESTER\b/i], output: "Eco-Friendly" },
    { keywords: [/\bBIODEGRADABLE\b/i, /\bCOMPOSTABLE\b/i], output: "Biodegradable" },
    { keywords: [/\bWATERPROOF\b/i, /\bWPN\sINSOLE\b/i], output: "Waterproof" },
    { keywords: [/\bWATER\sRESISTANT\b/i, /\bWATER\sREPELLENT\b/i], output: "Water Resistant" },
    { keywords: [/\bCHEMICAL\sFREE\b/i], output: "Chemical Free" },
    { keywords: [/\bDUSTPROOF\b/i], output: "Dustproof" },
    { keywords: [/\bHIGH\sTENSILE\sSTRENGTH\b/i], output: "High Tensile Strength" },
    { keywords: [/\bHIGH\sELONGATION\b/i], output: "High Elongation" },
    { keywords: [/\bELASTICITY\b/i, /\bSTRETCHY\b/i], output: "Elasticity" },
    { keywords: [/\bULTRA\sLIGHTWEIGHT\b/i], output: "Ultra Lightweight" },
    { keywords: [/\bLIGHTWEIGHT\b/i], output: "Lightweight" },
    { keywords: [/\bSOUND\sABSORBING\b/i, /\bSOUND\sINSULATING\b/i, /\bNOISE-PROOF\b/i], output: "Sound Absorbing" },
    { keywords: [/\bHEAT\sINSULATING\b/i, /\bTHERMAL\sINSULATING\b/i], output: "Heat Insulating" },
    { keywords: [/\bFILM\b/i], output: "Film", negations: [/NON-WOVEN\sFILM/i, /LAMINATED\sPE\sFILM/i] },
    {
      keywords: [/\bADHESIVE\b/i, /\bGLUE\b/i, /\bSELF-ADHESIVE\b/i, /\bCONSTRUCTION\sGLUE\b/i, /\bFABRIC\sGLUE\b/i, /\bWITH\sGLUE\b/i, /\bGLUED\b/i, /\bADHESIVE\sLAYER\b/i, /\bDOUBLE-SIDED\sTAPE\b/i],
      output: "Adhesive",
      negations: [/\bNOT\sCOATED\sWITH\sGLUE\b/i],
    },
    { keywords: [/\bFIBERFILL\b/i], output: "Fiberfill" },
    { keywords: [/\bMESH\b/i, /\bWEB\b/i], output: "Mesh" },
    { keywords: [/\bFAUX\sLEATHER\b/i, /\bSYNTHETIC\sLEATHER\b/i, /\bIMITATION\sLEATHER\b/i, /\bLEATHERETTE\b/i], output: "Faux Leather" },
  ];

  const globalNegationComplexRegex = /\bNOT\sIMPREGNATED,\sCOATED\sOR\sLAMINATED\b/i;
  const globalNegationUnUnUnRegex = /\bUNIMPREGNATED,\sUNCOATED,\sUNLAMINATED\b/i;
  const globalNegationNotImpOrCoatRegex = /\bNOT\sIMPREGNATED\sOR\sCOATED\b/i;
  const globalNegationUnImpAndUnCoatRegex = /\bUNIMPREGNATED\sAND\sUNCOATED\b/i;

  function isGloballyNegated(featureName) {
    if (globalNegationComplexRegex.test(descUpper) && (featureName === "Impregnated" || featureName === "Coated" || featureName === "Laminated")) return true;
    if (globalNegationUnUnUnRegex.test(descUpper) && (featureName === "Impregnated" || featureName === "Coated" || featureName === "Laminated")) return true;
    if (globalNegationNotImpOrCoatRegex.test(descUpper) && (featureName === "Impregnated" || featureName === "Coated")) return true;
    if (globalNegationUnImpAndUnCoatRegex.test(descUpper) && (featureName === "Impregnated" || featureName === "Coated")) return true;
    return false;
  }

  function isAffirmedOutsideGlobalNegation(featureName, keywordRegex) {
    const cleanedDesc = descUpper
      .replace(globalNegationComplexRegex, " G_NEG_ICL ")
      .replace(globalNegationUnUnUnRegex, " G_NEG_UUU ")
      .replace(globalNegationNotImpOrCoatRegex, " G_NEG_IC ")
      .replace(globalNegationUnImpAndUnCoatRegex, " G_NEG_UU ");
    return keywordRegex.test(cleanedDesc);
  }

  for (const def of addOnDefinitions) {
    let applyAddon = false;
    let keywordMatched = null;

    for (const keyword of def.keywords) {
      if (keyword.test(descUpper)) {
        keywordMatched = keyword;
        applyAddon = true;
        break;
      }
    }

    if (applyAddon) {
      let negatedByKeywordSpecific = false;
      if (def.negations) {
        for (const negation of def.negations) {
          if (negation.test(descUpper)) {
            if (def.output === "Adhesive" && negation.source.includes("NOT\\sCOATED\\sWITH\\sGLUE")) {
              const tempDescNoNegation = descUpper.replace(negation, "");
              if (def.keywords.some((k) => k.test(tempDescNoNegation))) {
                continue;
              }
            }
            negatedByKeywordSpecific = true;
            break;
          }
        }
      }
      if (negatedByKeywordSpecific) {
        applyAddon = false;
      }

      if (applyAddon && (def.output === "Impregnated" || def.output === "Coated" || def.output === "Laminated")) {
        if (isGloballyNegated(def.output)) {
          let isSpecificType = false;
          if (def.output === "Coated") {
            isSpecificType = specificCoatings.includes(def.output);
            if (!isSpecificType) {
              isSpecificType = specificCoatings.some((sc) => addOnDefinitions.find((d) => d.output === sc)?.keywords.some((k) => k.test(descUpper)));
            }
          }
          if (!isSpecificType && keywordMatched && !isAffirmedOutsideGlobalNegation(def.output, keywordMatched)) {
            applyAddon = false;
          }
        }
      }

      if (applyAddon && def.output === "Film" && /LAMINATED\s(?:PE\s)?FILM/i.test(descUpper)) {
        if (foundAddOns.has("Laminated") || foundAddOns.has("PE Coated")) {
          let filmIsStandalone = false;
          const filmMatches = [...descUpper.matchAll(/\bFILM\b/gi)];
          for (const filmMatch of filmMatches) {
            const surroundingText = descUpper.substring(Math.max(0, filmMatch.index - 20), Math.min(descUpper.length, filmMatch.index + filmMatch[0].length + 20));
            if (!/LAMINATED\s(?:PE\s)?FILM/i.test(surroundingText) && !/NON-WOVEN\sFILM/i.test(surroundingText)) {
              filmIsStandalone = true;
              break;
            }
          }
          if (!filmIsStandalone) applyAddon = false;
        }
      }

      if (applyAddon) {
        foundAddOns.add(def.output);
      }
    }
  }

  const colors = [
    { name: "Light Beige", pattern: /\bLIGHT\sBEIGE\b/i },
    { name: "Silver Gray", pattern: /\bSILVER\sGRAY\b/i },
    { name: "Sky Blue", pattern: /\bSKY\sBLUE\b/i },
    { name: "Pale Mauve", pattern: /\bPALE\sMAUVE\b/i },
    { name: "Monk's Robe", pattern: /\bMONK'S\sROBE\b/i },
    { name: "Dress Blue", pattern: /\bDRESS\sBLUE\b/i },
    { name: "China Blue", pattern: /\bCHINA\sBLUE\b/i },
    { name: "Blue Nights", pattern: /\bBLUE\sNIGHTS\b/i },
    { name: "Chateau Rose", pattern: /\bCHATEAU\sROSE\b/i },
    { name: "Cloud Dancer", pattern: /\bCLOUD\sDANCER\b/i },
    { name: "Moonlite Mauve", pattern: /\bMOONLITE\sMAUVE\b/i },
    { name: "Purple Haze", pattern: /\bPURPLE\sHAZE\b/i },
    { name: "Love Potion", pattern: /\bLOVE\sPOTION\b/i },
    { name: "Baltic Sea", pattern: /\bBALTIC\sSEA\b/i },
    { name: "Cloudburst", pattern: /\bCLOUDBURST\b/i },
    { name: "Orange Popsicle", pattern: /\bORANGE\sPOPSICLE\b/i },
    { name: "Purple Rose", pattern: /\bPURPLE\sROSE\b/i },
    { name: "Bright White", pattern: /\bBRIGHT\sWHITE\b/i },
    { name: "Cool White", pattern: /\bCOOL\sWHITE\b/i },
    { name: "Classic White", pattern: /\bCLASSIC\sWHITE\b/i },
    { name: "Black Beauty", pattern: /\bBLACK\sBEAUTY\b/i },
    { name: "White", pattern: /\bWHITE\b/i, negations: [/\bBRIGHT\sWHITE\b/i, /\bCOOL\sWHITE\b/i, /\bCLASSIC\sWHITE\b/i] },
    { name: "Black", pattern: /\bBLACK\b/i, negations: [/\bBLACK\sBEAUTY\b/i] },
    { name: "Pink", pattern: /\bPINK\b/i },
    { name: "Green", pattern: /\bGREEN\b/i, negations: [/\bASPG\sGRN\b/i] },
    { name: "Blue", pattern: /\bBLUE\b/i, negations: [/\bSKY\sBLUE\b/i, /\bDRESS\sBLUE\b/i, /\bCHINA\sBLUE\b/i, /\bBLUE\sNIGHTS\b/i, /\bBALTIC\sSEA\b/i] },
    { name: "Gray", pattern: /\bGRAY\b/i, negations: [/\bSILVER\sGRAY\b/i] },
    { name: "Grey", pattern: /\bGREY\b/i, negations: [/\bSILVER\sGRAY\b/i] }, // Sama dengan Gray
    { name: "Beige", pattern: /\bBEIGE\b/i, negations: [/\bLIGHT\sBEIGE\b/i] },
    { name: "Turquoise", pattern: /\bTURQUOISE\b/i },
    { name: "Charcoal", pattern: /\bCHARCOAL\b/i },
    { name: "Cream", pattern: /\bCREAM\b/i },
    { name: "Salsa", pattern: /\bSALSA\b/i },
    { name: "Fedora", pattern: /\bFEDORA\b/i },
    { name: "Caviar", pattern: /\bCAVIAR\b/i },
    { name: "Tomato", pattern: /\bTOMATO\b/i },
    { name: "Humus", pattern: /\bHUMUS\b/i },
    { name: "Cork", pattern: /\bCORK\b/i },
    { name: "Periscope", pattern: /\bPERISCOPE\b/i },
    { name: "Mediterranea", pattern: /\bMEDITERRANEA\b/i },
    { name: "Aspg Grn", pattern: /\bASPG\sGRN\b/i },
  ];

  for (const color of colors) {
    let negated = false;
    if (color.negations) {
      for (const negation of color.negations) {
        if (negation.test(descUpper) && !foundAddOns.has(color.name)) {
          negated = true;
          break;
        }
      }
    }
    if (negated) continue;

    if (color.pattern.test(descUpper)) {
      if (color.name === "Green" && foundAddOns.has("Aspg Grn")) continue;
      let specificVariantExists = false;
      if (color.name === "White" && (foundAddOns.has("Bright White") || foundAddOns.has("Cool White") || foundAddOns.has("Classic White"))) specificVariantExists = true;
      if (color.name === "Black" && foundAddOns.has("Black Beauty")) specificVariantExists = true;
      if (color.name === "Blue" && (foundAddOns.has("Sky Blue") || foundAddOns.has("Dress Blue") || foundAddOns.has("China Blue") || foundAddOns.has("Blue Nights") || foundAddOns.has("Baltic Sea"))) specificVariantExists = true;
      if (color.name === "Gray" && foundAddOns.has("Silver Gray")) specificVariantExists = true;
      if (color.name === "Grey" && foundAddOns.has("Silver Gray")) specificVariantExists = true;
      if (color.name === "Beige" && foundAddOns.has("Light Beige")) specificVariantExists = true;

      if (!specificVariantExists) {
        foundAddOns.add(color.name === "Aspg Grn" ? "Green" : color.name);
      }
    }
  }
  if (foundAddOns.has("Gray") && foundAddOns.has("Grey")) {
    foundAddOns.delete("Grey");
  }

  let finalAddOns = Array.from(foundAddOns);
  let hasSpecificCoating = specificCoatings.some((sc) => finalAddOns.includes(sc));

  if (hasSpecificCoating && finalAddOns.includes("Coated")) {
    const tempDescNoSpecificCoating = descUpper
      .replace(/\bPU\sCOATED\b/gi, "")
      .replace(/\bPVC\sCOATED\b/gi, "")
      .replace(/\bPE\sCOATED\b/gi, "")
      .replace(/\bHDPE\sCOATED\b/gi, "")
      .replace(/\bHDPE\sGLUE\sON\sTHE\sSURFACE\b/gi, "")
      .replace(/\bTHERMOPLASTIC\sCOATED\b/gi, "");
    if (!/\bCOATED\b/i.test(tempDescNoSpecificCoating) || /\bNOT\sCOATED\b/i.test(descUpper) || /\bUNCOATED\b/i.test(descUpper)) {
      if (!((/\bCOATED\sWITH\sGLUE\b/i.test(descUpper) || /\bGLUE-COATED\b/i.test(descUpper)) && finalAddOns.includes("Adhesive"))) {
        finalAddOns = finalAddOns.filter((a) => a !== "Coated");
      }
    }
  }

  if (finalAddOns.includes("Adhesive") && finalAddOns.includes("Coated")) {
    if (/\bCOATED\sWITH\sGLUE\b/i.test(descUpper) || /\bGLUE-COATED\b/i.test(descUpper)) {
      const tempDescNoGlueCoating = descUpper.replace(/\bCOATED\sWITH\sGLUE\b/gi, "").replace(/\bGLUE-COATED\b/gi, "");
      let otherCoatingExists = specificCoatings.some((sc) => finalAddOns.includes(sc));
      if (!otherCoatingExists) {
        otherCoatingExists = /\bCOATED\b/i.test(tempDescNoGlueCoating) && !(/\bNOT\sCOATED\b/i.test(tempDescNoGlueCoating) || /\bUNCOATED\b/i.test(tempDescNoGlueCoating));
      }
      if (!otherCoatingExists) {
        finalAddOns = finalAddOns.filter((a) => a !== "Coated");
      }
    }
  }

  const globalNegationIsPresent = isGloballyNegated("Impregnated") || isGloballyNegated("Coated") || isGloballyNegated("Laminated");
  if (globalNegationIsPresent) {
    const checkAndRemoveGeneral = (generalAddon, generalKeywordRegex, specificVersionsArray) => {
      if (finalAddOns.includes(generalAddon)) {
        let hasSpecificTypeInFinal = specificVersionsArray ? specificVersionsArray.some((sv) => finalAddOns.includes(sv)) : false;
        let affirmedOutside = isAffirmedOutsideGlobalNegation(generalAddon, generalKeywordRegex);
        if (hasSpecificTypeInFinal) {
          if (!affirmedOutside) {
            finalAddOns = finalAddOns.filter((a) => a !== generalAddon);
          }
        } else {
          if (!affirmedOutside) {
            finalAddOns = finalAddOns.filter((a) => a !== generalAddon);
          }
        }
      }
    };
    checkAndRemoveGeneral("Coated", /\bCOATED\b/i, specificCoatings);
    checkAndRemoveGeneral("Laminated", /\bLAMINATED\b/i, null);
    checkAndRemoveGeneral("Impregnated", /\bIMPREGNATED\b/i, null);
  }

  if (finalAddOns.includes("Extra Soft") && finalAddOns.includes("Soft")) finalAddOns = finalAddOns.filter((a) => a !== "Soft");
  if (finalAddOns.includes("Cotton Soft") && finalAddOns.includes("Soft")) finalAddOns = finalAddOns.filter((a) => a !== "Soft");
  if (finalAddOns.includes("Ultra Lightweight") && finalAddOns.includes("Lightweight")) finalAddOns = finalAddOns.filter((a) => a !== "Lightweight");
  if (finalAddOns.includes("Waterproof") && finalAddOns.includes("Water Resistant")) finalAddOns = finalAddOns.filter((a) => a !== "Water Resistant");

  if (finalAddOns.length === 0) return "-";
  return finalAddOns.sort().join("; ");
}

// --- Fungsi Utama (Gabungan dan Modifikasi) ---
async function processExcelFile(inputFilePath, outputFilePath) {
  try {
    if (!fs.existsSync(inputFilePath)) {
      console.error(`File input tidak ditemukan: ${inputFilePath}`);
      return;
    }

    const workbook = xlsx.readFile(inputFilePath);
    const sheetNames = workbook.SheetNames;

    console.log("\nSheet yang tersedia:");
    sheetNames.forEach((name, index) => {
      console.log(`${index + 1}. ${name}`);
    });

    const selectedSheetIndexesInput = readlineSync.question("\nMasukkan nomor sheet yang ingin diproses (pisahkan dengan koma jika lebih dari satu, misal: 1,3,5, atau biarkan kosong untuk memproses semua): ");

    let selectedSheetIndexes;
    if (selectedSheetIndexesInput.trim() === "") {
      selectedSheetIndexes = sheetNames.map((_, index) => index); // Proses semua sheet
      console.log("Memproses semua sheet...");
    } else {
      selectedSheetIndexes = selectedSheetIndexesInput
        .split(",")
        .map((idx) => parseInt(idx.trim()) - 1)
        .filter((idx) => idx >= 0 && idx < sheetNames.length);
    }

    if (selectedSheetIndexes.length === 0 && selectedSheetIndexesInput.trim() !== "") {
      console.error("Tidak ada sheet valid yang dipilih.");
      return;
    }
    if (selectedSheetIndexes.length === 0) {
      console.error("Tidak ada sheet yang dipilih atau ditemukan untuk diproses.");
      return;
    }

    const newWorkbook = xlsx.utils.book_new(); // Buat workbook baru di awal

    selectedSheetIndexes.forEach((sheetIndex) => {
      const sheetName = sheetNames[sheetIndex];
      const worksheet = workbook.Sheets[sheetName];
      const headerRowJson = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

      if (!headerRowJson || headerRowJson.length === 0) {
        console.warn(`Sheet '${sheetName}' kosong atau tidak memiliki header. Dilewati.`);
        return;
      }
      const headerRow = headerRowJson[0] || [];

      const descColumnName = headerRow.find((h) => h && (h.toUpperCase() === "ITEM DESC" || h.toUpperCase() === "PRODUCT DESCRIPTION(EN)"));

      if (!descColumnName) {
        console.error(`Kolom 'ITEM DESC' atau 'PRODUCT DESCRIPTION(EN)' tidak ditemukan di sheet '${sheetName}'. Sheet ini akan dilewati.`);
        return;
      }

      const jsonData = xlsx.utils.sheet_to_json(worksheet);
      if (jsonData.length === 0) {
        console.warn(`Sheet '${sheetName}' tidak memiliki data (selain header). Hasil sheet ini mungkin kosong.`);
      }

      const processedData = jsonData.map((row) => {
        const itemDesc = row[descColumnName] || "";

        const gsm = extractGsmValue(itemDesc);
        const width = extractWidthValue(itemDesc);
        const { item: itemType } = extractItemTypeAndPattern(itemDesc);
        const addOn = extractAddOnFromList(itemDesc);

        // Salin semua kolom asli dan tambahkan/timpa kolom baru
        let newRow = { ...row };
        newRow["GSM"] = gsm;
        newRow["Width (cm)"] = width;
        newRow["ITEM"] = itemType;
        newRow["ADD ON"] = addOn;
        return newRow;
      });

      // Menentukan header untuk output
      let outputHeaders;
      const newColumns = ["GSM", "Width (cm)", "ITEM", "ADD ON"];

      if (jsonData.length > 0) {
        const originalHeaders = Object.keys(jsonData[0]);
        // Filter original headers to remove any that will be re-added, to control order
        const uniqueOriginalHeaders = originalHeaders.filter((h) => !newColumns.some((nc) => nc.toUpperCase() === h.toUpperCase()));
        outputHeaders = [...uniqueOriginalHeaders, ...newColumns];
      } else {
        // Jika jsonData kosong tapi headerRow ada
        const uniqueOriginalHeaders = headerRow.filter((h) => !newColumns.some((nc) => nc.toUpperCase() === (h ? h.toUpperCase() : "")));
        outputHeaders = [...uniqueOriginalHeaders, ...newColumns];
      }
      // Pastikan kolom deskripsi ada di awal jika belum
      if (descColumnName && outputHeaders.includes(descColumnName) && outputHeaders[0].toUpperCase() !== descColumnName.toUpperCase()) {
        outputHeaders = outputHeaders.filter((h) => h.toUpperCase() !== descColumnName.toUpperCase());
        outputHeaders.unshift(descColumnName);
      } else if (descColumnName && !outputHeaders.includes(descColumnName)) {
        outputHeaders.unshift(descColumnName);
      }
      outputHeaders = [...new Set(outputHeaders)]; // Hapus duplikat jika ada

      const newWorksheet = xlsx.utils.json_to_sheet(processedData, { header: outputHeaders });
      xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName); // Tambahkan sheet ke workbook yang sama
      console.log(`Sheet '${sheetName}' berhasil diproses.`);
    });

    if (newWorkbook.SheetNames.length > 0) {
      xlsx.writeFile(newWorkbook, outputFilePath);
      console.log(`\nSemua sheet yang dipilih berhasil diproses dan disimpan sebagai ${outputFilePath}`);
    } else {
      console.log("\nTidak ada sheet yang diproses atau data untuk disimpan.");
    }
  } catch (error) {
    console.error("Terjadi kesalahan saat memproses file:", error);
  }
}

// --- Konfigurasi dan Eksekusi ---
const inputFileName = "input.xlsx"; // Ganti dengan nama file input Anda
const outputFileName = "output.xlsx"; // Nama file output yang diinginkan

const inputFilePath = path.join(__dirname, inputFileName);
const outputFilePath = path.join(__dirname, outputFileName);

processExcelFile(inputFilePath, outputFilePath);
