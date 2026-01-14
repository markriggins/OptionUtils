function assertEqual(actual, expected, message = "", eps = 1e-9) {
    // Same reference or primitive equality (covers identical strings, numbers, booleans)
    if (actual === expected) return;

    const aType = typeof actual;
    const eType = typeof expected;

    // Both numbers → floating-point tolerant compare
    if (aType === "number" && eType === "number") {
      if (!Number.isFinite(actual) || !Number.isFinite(expected)) {
        throw new Error(
          `${message}\nNon-finite number comparison\nActual: ${actual}\nExpected: ${expected}`
        );
      }
      if (Math.abs(actual - expected) <= eps) return;

      throw new Error(
        `${message}\nNumber mismatch\nExpected: ${expected}\nActual:   ${actual}\nΔ:        ${actual - expected}`
      );
    }

    // Both strings → exact match
    if (aType === "string" && eType === "string") {
      throw new Error(
        `${message}\nString mismatch\nExpected: "${expected}"\nActual:   "${actual}"`
      );
    }

    // Mismatched types → hard failure
    throw new Error(
      `${message}\nType mismatch\nExpected (${eType}): ${expected}\nActual   (${aType}): ${actual}`
    );
  }



/**
 * dateToMidnightTimestamp - Returns the Unix timestamp (ms) for 00:00:00 on the same calendar day
 * @param {Date} d - Input date (any time of day)
 * @returns {number|null} Milliseconds since 1970-01-01 UTC, or null if invalid
 */
function dateToMidnightTimestamp(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) {
    return null;
  }
  
  // Create new Date with only year/month/date (midnight local time)
  const midnight = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  
  // Return timestamp in milliseconds (UTC)
  return midnight.getTime();
}


function runAllTests() {
  const global = this;               // Apps Script global scope
  const testFns = [];

  // Discover all functions named test_*
  for (const name in global) {
    if (typeof global[name] === "function" && name.startsWith("test_")) {
      testFns.push(name);
    }
  }

  if (testFns.length === 0) {
    throw new Error("No test_* functions found");
  }

  // Run tests
  for (const name of testFns) {
    try {
      global[name]();
      console.log(`✅ ${name} PASSED`);
    } catch (e) {
      console.error(`❌ ${name} FAILED`);
      throw e;   // fail fast → red execution
    }
  }

  console.log(`🎉 ALL ${testFns.length} TESTS PASSED`);
}

