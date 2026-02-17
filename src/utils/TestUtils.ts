// @ts-check
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



// dateToMidnightTimestamp is in CommonUtils.js

/**
 * Assert two 2D arrays are deeply equal (supports Date comparison).
 */
function assertArrayDeepEqual(actual, expected, msg = "") {
  if (actual.length !== expected.length) {
    throw new Error(`ASSERT FAILED${msg ? " – " + msg : ""}\nLength mismatch: ${actual.length} != ${expected.length}`);
  }
  actual.forEach((row, i) => {
    const expRow = expected[i];
    if (row.length !== expRow.length) {
      throw new Error(`ASSERT FAILED${msg ? " – " + msg : ""}\nRow ${i} length mismatch`);
    }
    row.forEach((v, j) => {
      if (v instanceof Date && expRow[j] instanceof Date) {
        if (v.getTime() !== expRow[j].getTime()) {
          throw new Error(`ASSERT FAILED${msg ? " – " + msg : ""}\nRow ${i}, Col ${j}: Dates differ`);
        }
      } else if (v !== expRow[j]) {
        throw new Error(`ASSERT FAILED${msg ? " – " + msg : ""}\nRow ${i}, Col ${j}: ${v} != ${expRow[j]}`);
      }
    });
  });
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

