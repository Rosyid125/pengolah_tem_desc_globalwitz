const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");
const readlineSync = require("readline-sync"); // Diperlukan untuk input sheet

// --- Fungsi Ekstraksi GSM (dari skrip kedua, lebih komprehensif) ---
function extractGsmValue(desc) {
  if (typeof desc !== "string") return "-";

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

  return "-";
}

// --- Fungsi Ekstraksi Width (dari skrip pertama) ---
function extractWidthValue(desc) {
  if (typeof desc !== "string") return "-";

  let match;
    // 1. Priority patterns for specific data examples - highest priority
  // 60'' WIDE, 54'' WIDE, 58'' WIDE, 44'' WIDE, 36'' WIDE, etc.
  match = desc.match(/(\d+[\d.,]*)\s*(?:''|"|inch|inches)\s+wide/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // 150CM WIDE, 152CM WIDE, 1M 32GSM WIDE, 1M 36GSM WIDE, 1.5M 60GSM WIDE, etc.
  match = desc.match(/(\d+[\d.,]*)\s*(?:cm|m)\s+(?:\d+gsm\s+)?wide/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // 22IN WIDTH, 18IN WIDTH, etc.
  match = desc.match(/(\d+[\d.,]*)\s*(?:in|inch|inches)\s+width/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // CUT WIDTH and EDGE WIDTH patterns
  // 150CM CUT WIDTH, 60 INCH EDGE WIDTH, etc.
  match = desc.match(/(\d+[\d.,]*)\s*(?:cm|inch|inches|''|")\s+(?:cut\s+width|edge\s+width)/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // WIDTH FROM patterns
  // WIDTH FROM 40", etc.
  match = desc.match(/width\s+from\s+(\d+[\d.,]*)\s*(?:''|"|inch|inches)/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // NON-UNIFORM WIDTH patterns with ranges
  // NON-UNIFORM WIDTH (48-60 INCHES), etc.
  match = desc.match(/non-uniform\s+width\s*\(\s*(\d+[\d.,]*)\s*-\s*\d+[\d.,]*\s*(?:inches|inch|''|")\s*\)/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // Special dimensional patterns
  // 30GSM*10MM(WIDTH), etc.
  match = desc.match(/\*(\d+[\d.,]*)\s*mm\s*\(\s*width\s*\)/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // 2. Width context with inch patterns - all variations including [angka][unit] WIDTH
  // WIDTH = 65", WIDTH: 50", WIDTH 70", WIDTH65", 50 inch WIDE, WIDE 60", 65" WIDTH, etc.
  match = desc.match(/(?:width\s*[=:]\s*(\d+[\d.,]*)\s*(?:inch|inches|\")|width\s+(\d+[\d.,]*)\s*(?:inch|inches|\")|width(\d+[\d.,]*)\s*(?:inch|inches|\")|(\d+[\d.,]*)\s*(?:inch|inches|\'\')\s*wide|wide\s+(\d+[\d.,]*)\s*(?:inch|inches|\")|(\d+[\d.,]*)\s*(?:inch|inches|\")\s*width)/i);
  if (match) {
    const value = match[1] || match[2] || match[3] || match[4] || match[5] || match[6];
    if (value) {
      return parseFloat(value.replace(",", ".")).toFixed(2);
    }
  }

  // 3. Width context with meter patterns - all variations including [angka][unit] WIDTH
  // WIDTH = 1.5M, WIDTH: 2M, WIDTH 1.2M, WIDTH1.5M, 2 meter WIDE, WIDE 1M, 1.5M WIDTH, etc.
  match = desc.match(/(?:width\s*[=:]\s*(\d+[\d.,]*)\s*(?:m|meter|metres?)|width\s+(\d+[\d.,]*)\s*(?:m|meter|metres?)|width(\d+[\d.,]*)\s*(?:m|meter|metres?)|(\d+[\d.,]*)\s*(?:m|meter|metres?)\s*wide|wide\s+(\d+[\d.,]*)\s*(?:m|meter|metres?)|(\d+[\d.,]*)\s*(?:m|meter|metres?)\s*width)/i);
  if (match) {
    const value = match[1] || match[2] || match[3] || match[4] || match[5] || match[6];
    if (value) {
      return parseFloat(value.replace(",", ".")).toFixed(2);
    }
  }

  // 4. Width context with foot/feet patterns - all variations including [angka][unit] WIDTH
  // WIDTH = 5ft, WIDTH: 3ft, WIDTH 4ft, WIDTH5ft, 3 feet WIDE, WIDE 2ft, 5ft WIDTH, etc.
  match = desc.match(/(?:width\s*[=:]\s*(\d+[\d.,]*)\s*(?:ft|foot|feet)|width\s+(\d+[\d.,]*)\s*(?:ft|foot|feet)|width(\d+[\d.,]*)\s*(?:ft|foot|feet)|(\d+[\d.,]*)\s*(?:ft|foot|feet)\s*wide|wide\s+(\d+[\d.,]*)\s*(?:ft|foot|feet)|(\d+[\d.,]*)\s*(?:ft|foot|feet)\s*width)/i);
  if (match) {
    const value = match[1] || match[2] || match[3] || match[4] || match[5] || match[6];
    if (value) {
      return parseFloat(value.replace(",", ".")).toFixed(2);
    }
  }

  // 5. Width context with yard patterns - all variations including [angka][unit] WIDTH
  // WIDTH = 2yd, WIDTH: 1yd, WIDTH 3yd, WIDTH2yd, 1 yard WIDE, WIDE 2yd, 2yd WIDTH, etc.
  match = desc.match(/(?:width\s*[=:]\s*(\d+[\d.,]*)\s*(?:yd|yards?)|width\s+(\d+[\d.,]*)\s*(?:yd|yards?)|width(\d+[\d.,]*)\s*(?:yd|yards?)|(\d+[\d.,]*)\s*(?:yd|yards?)\s*wide|wide\s+(\d+[\d.,]*)\s*(?:yd|yards?)|(\d+[\d.,]*)\s*(?:yd|yards?)\s*width)/i);
  if (match) {
    const value = match[1] || match[2] || match[3] || match[4] || match[5] || match[6];
    if (value) {
      return parseFloat(value.replace(",", ".")).toFixed(2);
    }
  }

  // 6. Width context with centimeter patterns - all variations including [angka][unit] WIDTH
  // WIDTH = 150cm, WIDTH: 100cm, WIDTH 120cm, WIDTH150cm, 100 cm WIDE, WIDE 80cm, 150cm WIDTH, etc.
  match = desc.match(/(?:width\s*[=:]\s*(\d+[\d.,]*)\s*cm|width\s+(\d+[\d.,]*)\s*cm|width(\d+[\d.,]*)\s*cm|(\d+[\d.,]*)\s*cm\s*wide|wide\s+(\d+[\d.,]*)\s*cm|(\d+[\d.,]*)\s*cm\s*width)/i);
  if (match) {
    const value = match[1] || match[2] || match[3] || match[4] || match[5] || match[6];
    if (value) {
      return parseFloat(value.replace(",", ".")).toFixed(2);
    }
  }

  // 7. Width context with millimeter patterns - all variations including [angka][unit] WIDTH
  // WIDTH = 1500mm, WIDTH: 1000mm, WIDTH 1200mm, WIDTH1500mm, 1000 mm WIDE, WIDE 800mm, 1500mm WIDTH, etc.
  match = desc.match(/(?:width\s*[=:]\s*(\d+[\d.,]*)\s*mm|width\s+(\d+[\d.,]*)\s*mm|width(\d+[\d.,]*)\s*mm|(\d+[\d.,]*)\s*mm\s*wide|wide\s+(\d+[\d.,]*)\s*mm|(\d+[\d.,]*)\s*mm\s*width)/i);
  if (match) {
    const value = match[1] || match[2] || match[3] || match[4] || match[5] || match[6];
    if (value) {
      return parseFloat(value.replace(",", ".")).toFixed(2);
    }
  }

  // 8. Width context with inch patterns (handles '' apostrophe variations)
  // 60'', 54'', 58'', etc.
  match = desc.match(/(\d+[\d.,]*)\s*\'\'\s*wide/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // 9. Width context with IN patterns (22IN, 18IN, etc.)
  match = desc.match(/(?:width\s*[=:]\s*(\d+[\d.,]*)\s*in|width\s+(\d+[\d.,]*)\s*in|width(\d+[\d.,]*)\s*in|(\d+[\d.,]*)\s*in\s*width)/i);
  if (match) {
    const value = match[1] || match[2] || match[3] || match[4];
    if (value) {
      return parseFloat(value.replace(",", ".")).toFixed(2);
    }
  }

  // 10. Width context with GSM patterns (like "1M 32GSM WIDE")
  match = desc.match(/(\d+[\d.,]*)\s*m\s+\d+gsm\s*wide/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // 11. Width context WITHOUT units - all variations 
  // WIDTH = 150, WIDTH: 100, WIDTH 120, WIDTH150, 150 WIDE, WIDE 100, 150WIDTH, etc.
  // Also handle cases like "WIDTH 150//40G" where GSM follows, and "60 WIDTH"
  match = desc.match(/(?:width\s*[=:]\s*(\d+[\d.,]*)|width\s+(\d+[\d.,]*)|width(\d+[\d.,]*)|(\d+[\d.,]*)\s*wide|wide\s+(\d+[\d.,]*)|(\d+[\d.,]*)width|(\d+[\d.,]*)\s+width)(?!\s*(?:inch|inches|\"|\'\'|in|m|meter|metres?|ft|foot|feet|yd|yards?|cm|mm|μm|microns?|micrometers?))/i);
  if (match) {
    const value = match[1] || match[2] || match[3] || match[4] || match[5] || match[6] || match[7];
    if (value) {
      return parseFloat(value.replace(",", ".")).toFixed(2);
    }
  }
  // 12. Additional WIDTH FROM patterns (for cases not caught by priority patterns)
  // "WIDTH FROM 40"", "WIDTH FROM 48"
  match = desc.match(/width\s+from\s+(\d+[\d.,]*)\s*(?:inch|inches|\"|\'\')?\b/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // 13. Range patterns in parentheses
  // WIDTH (45"-72"), WIDTH (12CM~26CM), etc.
  match = desc.match(/width\s*\(\s*(\d+[\d.,]*)\s*(?:inch|inches|\")\s*-\s*(\d+[\d.,]*)\s*(?:inch|inches|\")\s*\)/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  match = desc.match(/width\s*\(\s*(\d+[\d.,]*)\s*cm\s*~\s*(\d+[\d.,]*)\s*cm\s*\)/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // 14. CUT WIDTH patterns
  // "150CM CUT WIDTH", "152CM CUT WIDTH"
  match = desc.match(/(\d+[\d.,]*)\s*cm\s+cut\s+width/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // 15. EDGE WIDTH patterns  
  // "60 INCH EDGE WIDTH"
  match = desc.match(/(\d+[\d.,]*)\s*(?:inch|inches)\s+edge\s+width/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // 16. Complex patterns with parentheses and multiple values
  // 152.4 CM (60 IN) WIDE, 150CM (+/-5CM) WIDE, etc.
  match = desc.match(/(\d+[\d.,]*)\s*cm\s*\([^)]*\)\s*wide/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  match = desc.match(/(\d+[\d.,]*)\s*(?:inch|inches|\"|\'\')?\s*\([^)]*\)\s*wide/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // 17. Patterns with slashes
  // 35/36'' WIDE, patterns with slashes
  match = desc.match(/(\d+[\d.,]*)\s*\/\s*\d+[\d.,]*\s*(?:inch|inches|\'\')\s*wide/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  // 18. Small width patterns
  // "WIDTH LESS THAN 10CM", "10MM SMALL WIDTH"
  match = desc.match(/width\s+less\s+than\s+(\d+[\d.,]*)\s*cm/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }

  match = desc.match(/(\d+[\d.,]*)\s*mm\s+small\s+width/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }
  // 19. Dimensional patterns (LENGTH*WIDTH)
  // "320*300MM (LENGTH*WIDTH)", "80*260MM (LENGTH * WIDTH)"
  match = desc.match(/\d+[\d.,]*\*(\d+[\d.,]*)\s*mm\s*\(\s*length\s*\*\s*width\s*\)/i);
  if (match) {
    return parseFloat(match[1].replace(",", ".")).toFixed(2);
  }  // 20. Four-step fallback extraction when no WIDTH/WIDE keywords present
  // Step 1: Check for WIDTH/WIDE keywords first (already handled above)
  // Step 2-4: Only if no WIDTH or WIDE found in the description
  if (!/\b(?:width|wide)\b/i.test(desc)) {
    
    // Step 2: Standard [angka][unit] with spaces
    // Inches (most common for width)
    match = desc.match(/\b(\d+[\d.,]*)\s+(?:''|"|inch|inches|in)\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Centimeters
    match = desc.match(/\b(\d+[\d.,]*)\s+cm\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Meters
    match = desc.match(/\b(\d+[\d.,]*)\s+(?:m|meter|metres?)\b/i);
    if (match) {
      const value = parseFloat(match[1].replace(",", "."));
      return value.toFixed(2);
    }
    
    // Millimeters
    match = desc.match(/\b(\d+[\d.,]*)\s+mm\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Feet
    match = desc.match(/\b(\d+[\d.,]*)\s+(?:ft|foot|feet)\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Yards
    match = desc.match(/\b(\d+[\d.,]*)\s+(?:yd|yards?)\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Step 3: [angka][unit] without space in front (concatenated at start)
    // Examples: 40IN, 150CM, 80MM (but not 80MMGSM), 1.5M, 5FT, 2YD
    
    // Inches without space in front
    match = desc.match(/\b(\d+[\d.,]*)(?:''|"|inch|inches|in)\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Centimeters without space in front
    match = desc.match(/\b(\d+[\d.,]*)cm\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Meters without space in front
    match = desc.match(/\b(\d+[\d.,]*)(?:m|meter|metres?)\b/i);
    if (match) {
      const value = parseFloat(match[1].replace(",", "."));
      return value.toFixed(2);
    }
    
    // Millimeters without space in front (avoid GSM conflicts)
    match = desc.match(/\b(\d+[\d.,]*)mm\b(?!.*gsm)/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Feet without space in front
    match = desc.match(/\b(\d+[\d.,]*)(?:ft|foot|feet)\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Yards without space in front
    match = desc.match(/\b(\d+[\d.,]*)(?:yd|yards?)\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Step 4: [angka][unit] without space behind (concatenated patterns)
    // Examples: 35G40IN, PPU.70", 80MM40GSM - extract the second number+unit
    
    // Extract inch values from concatenated strings like 35G40IN, PPU.70"
    match = desc.match(/(?:\d+[A-Z]*\.?)(\d+[\d.,]*)(?:''|"|IN|INCH|INCHES)\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Extract CM values from concatenated strings
    match = desc.match(/(?:\d+[A-Z]*\.?)(\d+[\d.,]*)CM\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Extract M values from concatenated strings
    match = desc.match(/(?:\d+[A-Z]*\.?)(\d+[\d.,]*)M\b/i);
    if (match) {
      const value = parseFloat(match[1].replace(",", "."));
      return value.toFixed(2);
    }
    
    // Extract MM values from concatenated strings (but avoid GSM conflicts)
    match = desc.match(/(?:^|[^G])(\d+[\d.,]*)MM\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
    
    // Extract FT values from concatenated strings
    match = desc.match(/(?:\d+[A-Z]*\.?)(\d+[\d.,]*)FT\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
      // Extract YD values from concatenated strings
    match = desc.match(/(?:\d+[A-Z]*\.?)(\d+[\d.,]*)YD\b/i);
    if (match) {
      return parseFloat(match[1].replace(",", ".")).toFixed(2);
    }
  }

  return "-";
}

// --- Fungsi Ekstraksi ITEM (dari skrip kedua) ---
function extractItemTypeAndPattern(desc) {
  if (typeof desc !== "string") return { item: "-", pattern: "" };
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
    const matchedPattern = descUpper.match(nonwovenPattern);    return { item: "Nonwoven", pattern: matchedPattern ? matchedPattern[0] : "" };
  }

  return { item: "-", pattern: "" };
}

// --- Fungsi Ekstraksi ADD ON (dari skrip kedua) ---

function extractAddOnFromList(desc) {
  if (typeof desc !== "string") return "-";
  const descUpper = desc.toUpperCase();
  const colors = [
    { name: "BLACK", pattern: /\bBLACK\b/i },
    { name: "BLUE", pattern: /\bBLUE\b(?!\s+SKY)/i }, // Avoid matching "SKY BLUE"
    { name: "BROWN", pattern: /\bBROWN\b/i },
    { name: "CHARCOAL", pattern: /\bCHARCOAL\b/i },
    { name: "CREAM", pattern: /\bCREAM\b/i },
    { name: "GRAY", pattern: /\bGRA[Y|E]\b/i }, // Matches both GRAY and GREY
    { name: "GREEN", pattern: /\bGREEN\b/i },
    { name: "LIGHT BEIGE", pattern: /\bLIGHT\s+BEIGE\b/i },
    { name: "NAVY", pattern: /\bNAVY\b/i },
    { name: "NEUTRAL", pattern: /\bNEUTRAL\b/i },
    { name: "OFF WHITE", pattern: /\bOFF\s+WHITE\b/i },
    { name: "PINK", pattern: /\bPINK\b/i },
    { name: "RED", pattern: /\bRED\b/i },
    { name: "SILVER", pattern: /\bSILVER\b/i },
    { name: "SKY BLUE", pattern: /\bSKY\s+BLUE\b/i },
    { name: "TURQUOISE", pattern: /\bTURQUOISE\b/i },
    { name: "WHITE", pattern: /\b(?:MILKY\s+WHITE|SNOW\s+WHITE|WHITE)\b/i }, // Includes variations
    { name: "YELLOW", pattern: /\bYELLOW\b/i },
  ];  const sifat = [
    { name: "HO", pattern: /\b(?:HO|PHOBIC|HYDROPHOBIC)\b/i },
    { name: "HI", pattern: /\b(?:HI|PHILIC|HYDROPHILIC)\b/i },
  ];
  
  const softness = [
    { name: "SUPER SOFT", pattern: /\b(?:SUPER\s+SOFT|EXTRA\s+SOFT)\b/i },
    { name: "SOFT", pattern: /\bSOFT\b(?!\s+(?:SUPER|EXTRA))/i }, // SOFT but not SUPER SOFT or EXTRA SOFT
  ];  let result = [];
  
  // Find detected colors
  let detectedColors = colors.filter(color => color.pattern.test(descUpper));
  
  // Find detected softness
  let detectedSoftness = softness.filter(soft => soft.pattern.test(descUpper));
  
  // Find detected properties (HO/HI)
  let detectedSifat = sifat.filter(s => s.pattern.test(descUpper));
  
  // Build combinations: COLOR + SOFTNESS + HI/HO
  if (detectedColors.length > 0) {
    for (const color of detectedColors) {
      let combination = color.name;
      
      // Add softness if detected
      if (detectedSoftness.length > 0) {
        for (const soft of detectedSoftness) {
          combination += ` ${soft.name}`;
        }
      }
      
      // Add properties if detected
      if (detectedSifat.length > 0) {
        for (const sifat of detectedSifat) {
          combination += ` ${sifat.name}`;
        }
      }
      
      result.push(combination);
    }
  }
  
  // If no colors but softness/properties detected, add them independently
  if (detectedColors.length === 0) {
    if (detectedSoftness.length > 0) {
      for (const soft of detectedSoftness) {
        let combination = soft.name;
        if (detectedSifat.length > 0) {
          for (const sifat of detectedSifat) {
            combination += ` ${sifat.name}`;
          }
        }
        result.push(combination);
      }
    } else if (detectedSifat.length > 0) {
      for (const sifat of detectedSifat) {
        result.push(sifat.name);
      }
    }
  }
  
  if (result.length === 0) return "-";
  return result.join(" ");
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
        newRow["WIDTH"] = width;
        newRow["ITEM"] = itemType;
        newRow["ADD ON"] = addOn;
        return newRow;
      });

      // Menentukan header untuk output
      let outputHeaders;
      const newColumns = ["GSM", "WIDTH", "ITEM", "ADD ON"];

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
