/**
 * test_normalizeType_ - Runs assertions for normalizeType_ function.
 *
 * Covers:
 * - Null/undefined/empty inputs
 * - Case insensitivity
 * - Whitespace variations (spaces, nbsp)
 * - Dash variants (hyphens, en/em dashes)
 * - Unicode normalization (e.g., accented chars if any)
 * - Plurals and singulars
 * - Invalid inputs
 *
 * Run this in Google Apps Script editor or console.
 * Logs "✅ All tests passed" if successful.
 */
function test_normalizeType() {
  // Helper assert
  function assertEqual(actual, expected, msg) {
    if (actual !== expected) {
      throw new Error(`ASSERT FAILED: ${msg}\nExpected: ${expected}\nActual: ${actual}`);
    }
  }

  // ---- Null/Empty ----
  assertEqual(normalizeType_(null), null, "null input");
  assertEqual(normalizeType_(undefined), null, "undefined input");
  assertEqual(normalizeType_(""), null, "empty string");
  assertEqual(normalizeType_("   "), null, "whitespace only");

  // ---- STOCK ----
  // Basics
  assertEqual(normalizeType_("stock"), "stock", "stock basic");
  assertEqual(normalizeType_("STOCK"), "stock", "stock uppercase");
  assertEqual(normalizeType_("Stock"), "stock", "stock mixed case");
  // Plurals
  assertEqual(normalizeType_("stocks"), "stock", "stocks plural");
  assertEqual(normalizeType_("SHARES"), "stock", "shares uppercase");
  assertEqual(normalizeType_("share"), "stock", "share singular");
  // Whitespace/dashes
  assertEqual(normalizeType_(" stock "), "stock", "stock with spaces");
  assertEqual(normalizeType_("st ock"), null, "invalid with space inside");
  // Unicode/dashes (though not directly in "stock")
  assertEqual(normalizeType_("stock\u00a0"), "stock", "stock with nbsp"); // nbsp -> space -> trim

  // ---- BULL CALL SPREAD ----
  // Basics
  assertEqual(normalizeType_("bcs"), "bull-call-spread", "bcs abbr");
  assertEqual(normalizeType_("BCS"), "bull-call-spread", "BCS upper");
  assertEqual(normalizeType_("bull call spread"), "bull-call-spread", "bull call spread spaces");
  assertEqual(normalizeType_("Bull-Call-Spread"), "bull-call-spread", "bull-call-spread dashes");
  assertEqual(normalizeType_("bull.call.spread"), "bull-call-spread", "bull.call.spread dots");
  // Plurals
  assertEqual(normalizeType_("bull call spreads"), "bull-call-spread", "bull call spreads plural");
  // Variations
  assertEqual(normalizeType_("bull\u2013call\u2014spread"), "bull-call-spread", "bull en/em dash spread"); // dashes → "-"
  assertEqual(normalizeType_(" bull  call   spread "), "bull-call-spread", "extra spaces");
  assertEqual(normalizeType_("BuLl CaLl SpReAd"), "bull-call-spread", "mixed case");
  // Invalid for BCS
  assertEqual(normalizeType_("bull calls spread"), null, "invalid plural mismatch");
  assertEqual(normalizeType_("bcs extra"), null, "bcs with extra");

  // ---- BULL PUT SPREAD ----
  // Basics
  assertEqual(normalizeType_("bps"), "bull-put-spread", "bps abbr");
  assertEqual(normalizeType_("BPS"), "bull-put-spread", "BPS upper");
  assertEqual(normalizeType_("bull put spread"), "bull-put-spread", "bull put spread spaces");
  assertEqual(normalizeType_("Bull-Put-Spread"), "bull-put-spread", "bull-put-spread dashes");
  assertEqual(normalizeType_("bull.put.spread"), "bull-put-spread", "bull.put.spread dots");
  // Plurals
  assertEqual(normalizeType_("bull put spreads"), "bull-put-spread", "bull put spreads plural");
  // Variations
  assertEqual(normalizeType_("bull\u2013put\u2014spread"), "bull-put-spread", "bull en/em dash spread");
  assertEqual(normalizeType_(" bull  put   spread "), "bull-put-spread", "extra spaces");
  assertEqual(normalizeType_("BuLl PuT SpReAd"), "bull-put-spread", "mixed case");
  // Invalid for BPS
  assertEqual(normalizeType_("bull puts spread"), null, "invalid plural mismatch");
  assertEqual(normalizeType_("bps extra"), null, "bps with extra");

  // ---- Invalid/General ----
  assertEqual(normalizeType_("call"), null, "partial match");
  assertEqual(normalizeType_("bull spread"), null, "missing type");
  assertEqual(normalizeType_("bear call spread"), null, "wrong direction");
  assertEqual(normalizeType_("123"), null, "numbers");
  assertEqual(normalizeType_("stock spread"), null, "mixed invalid");

  // ---- Unicode edge ----
  // Assuming no accents in types, but normalization
  assertEqual(normalizeType_("sto\u0301ck"), null, "accented stock -> 'stóck' -> 'stock' after NFKD? Wait, 'ó' -> 'o' + combining acute");
  Logger.log("✅ All normalizeType_ tests passed");
}

function normalizeType_(raw) {
  if (raw == null) return null;

  const t = String(raw)
    .normalize("NFKD")                     // Unicode canonicalization
    .toLowerCase()
    .replace(/[\u2010-\u2015\u2212]/g, "-") // dash variants → "-"
    .replace(/\u00a0/g, " ")               // nbsp → space
    .replace(/\s+/g, " ")                  // collapse whitespace
    .trim();

  // ---- STOCK ----
  if (/^(stock|stocks|share|shares)$/.test(t)) {
    return "stock";
  }

  // ---- CALL ----
  if (/^(call|calls)$/.test(t)) {
    return "Call";
  }

  // ---- PUT ----
  if (/^(put|puts)$/.test(t)) {
    return "Put";
  }

  // ---- BULL CALL SPREAD ----
  // Matches:
  //   bcs
  //   bull-call-spread
  //   bull call spreads
  //   bull.call.spread
  if (/^(bcs|bull[\s.\-]?call[\s.\-]?spread(s)?)$/.test(t)) {
    return "bull-call-spread";
  }

  // ---- BULL PUT SPREAD ----
  // Matches:
  //   bps
  //   bull-put-spread
  //   bull put spreads
  //   bull.put.spread
  if (/^(bps|bull[\s.\-]?put[\s.\-]?spread(s)?)$/.test(t)) {
    return "bull-put-spread";
  }

  return null;
}
