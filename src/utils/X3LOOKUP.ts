// @ts-check
/**
 * X3LOOKUP - Generic 3-key exact lookup with caching
 * Does NOT assume anything about dates, symbols, strikes, or sheet structure.  But note that
 * for indexing efficiency, date inputs will be converted to timestamps.
 * 
 * @param {*} key1         First key (e.g. symbol)
 * @param {*} key2         Second key (e.g. expiration date)
 * @param {*} key3         Third key (e.g. strike)
 * @param {Range} key1Col  Column range for key1 values
 * @param {Range} key2Col  Column range for key2 values
 * @param {Range} key3Col  Column range for key3 values
 * @param {Range} returnCol Column range to return values from
 * @returns {*} matched value or error string
 * @customfunction
 */
function X3LOOKUP(key1, key2, key3, key1Col, key2Col, key3Col, returnCol) {
  if (arguments.length !== 7) return "#VALUE!";
  if ([key1, key2, key3, key1Col, key2Col, key3Col, returnCol].some(v => v == null)) return "#REF!";

  const cache = CacheService.getScriptCache();
  const cacheKey = "X3LOOKUP_Cache";
  let indexJson = cache.get(cacheKey);

  let indexMap;
  if (indexJson) {
    try {
      const obj = JSON.parse(indexJson);
      indexMap = new Map(Object.entries(obj));
    } catch (e) {
      // corrupt → rebuild below
    }
  }

  if (!indexMap) {
    const sheet = key1Col.getSheet();
    if (!sheet) return "#REF! Invalid key column range";

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return "#N/A No data";

    indexMap = new Map();

    const k1Idx = key1Col.getColumn() - 1;
    const k2Idx = key2Col.getColumn() - 1;
    const k3Idx = key3Col.getColumn() - 1;
    const retIdx = returnCol.getColumn() - 1;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row.length <= Math.max(k1Idx, k2Idx, k3Idx, retIdx)) continue;

      let v1 = row[k1Idx];
      let v2 = row[k2Idx];
      let v3 = row[k3Idx];
      const retVal = row[retIdx];

      // Normalize
      if (v1 instanceof Date) v1 = dateToMidnightTimestamp(v1);
      else if (typeof v1 === "string") v1 = v1.trim().toUpperCase();
      else v1 = Number(v1);

      if (v2 instanceof Date) v2 = dateToMidnightTimestamp(v2);
      else if (typeof v2 === "string") v2 = v2.trim().toUpperCase();
      else v2 = Number(v2);

      if (v3 instanceof Date) v3 = dateToMidnightTimestamp(v3);
      else if (typeof v3 === "string") v3 = v3.trim().toUpperCase();
      else v3 = Number(v3);

      const cacheKey = `${v1}|${v2}|${v3}`;
      indexMap.set(cacheKey, retVal);
    }

    const serializable = Object.fromEntries(indexMap);
    cache.put(cacheKey, JSON.stringify(serializable), 21600);
  }

  let lk1 = key1;
  let lk2 = key2;
  let lk3 = key3;

  if (lk1 instanceof Date) lk1 = dateToMidnightTimestamp(lk1);
  else if (typeof lk1 === "string") lk1 = lk1.trim().toUpperCase();
  else lk1 = Number(lk1);

  if (lk2 instanceof Date) lk2 = dateToMidnightTimestamp(lk2);
  else if (typeof lk2 === "string") lk2 = lk2.trim().toUpperCase();
  else lk2 = Number(lk2);

  if (lk3 instanceof Date) lk3 = dateToMidnightTimestamp(lk3);
  else if (typeof lk3 === "string") lk3 = lk3.trim().toUpperCase();
  else lk3 = Number(lk3);

  const lookupKey = `${lk1}|${lk2}|${lk3}`;
  const value = indexMap.get(lookupKey);

  return value !== undefined ? value : "#N/A";
}

/**
 * TEST FUNCTION - run from editor to verify X3LOOKUP
 */
function test_X3LOOKUP() {
  // Mock table
  const mockTable = [
    ["Ticker", "ExpDate",      "Strike", "Price", "Volume"],
    ["TSLA",  new Date(2028,11,15), 350,     205.00,  120],
    ["TSLA",  new Date(2028,11,15), 400,     180.50,  85],
    ["AMZN",  new Date(2028,11,15), 180,     22.30,   45],
    ["TSLA",  new Date(2028, 5,16), 580,     147.75,  30],
    ["TSLA",  new Date(2028,11,15), 580,     142.90,  65],
    ["NVDA",  new Date(2028,11,15), 900,     88.40,   200]
  ];

  const mockSheet = {
    getDataRange: () => ({ getValues: () => mockTable }),
    getSheet: () => mockSheet
  };

  const mockRange = (colLetter) => ({
    getColumnLetter: () => colLetter,
    getSheet: () => mockSheet,
    getColumn: () => "ABCDE".indexOf(colLetter) + 1,
    getNumColumns: () => 1
  });

  const mockCache = {
    data: {},
    get: function(k) { return this.data[k]; },
    put: function(k, v) { this.data[k] = v; }
  };
  const originalCache = CacheService.getScriptCache;
  CacheService.getScriptCache = () => mockCache;

  // Build mock cache (using dateToMidnightTimestamp as timestamp)
  const tempMap = new Map();
  for (let i = 1; i < mockTable.length; i++) {
    const row = mockTable[i];
    let v1 = row[0]; if (typeof v1 === "string") v1 = v1.trim().toUpperCase();
    let v2 = row[1] instanceof Date ? dateToMidnightTimestamp(row[1]) : row[1];
    let v3 = Number(row[2]);
    const retVal = row[3];
    if (v2 !== null && !isNaN(v3)) {
      tempMap.set(`${v1}|${v2}|${v3}`, retVal);
    }
  }
  mockCache.put("X3LOOKUP_Cache", JSON.stringify(Object.fromEntries(tempMap)));

  // Tests
  const dec2028 = new Date(2028,11,15);
  const jun2028 = new Date(2028, 5,16);

  assertEqual(X3LOOKUP("TSLA", dec2028, 350, mockRange("A"), mockRange("B"), mockRange("C"), mockRange("D")), 205.00);
  assertEqual(X3LOOKUP("tsla", dec2028, 400, mockRange("A"), mockRange("B"), mockRange("C"), mockRange("D")), 180.50);
  assertEqual(X3LOOKUP("AMZN", dec2028, 180, mockRange("A"), mockRange("B"), mockRange("C"), mockRange("D")), 22.30);
  assertEqual(X3LOOKUP("TSLA", jun2028, 580, mockRange("A"), mockRange("B"), mockRange("C"), mockRange("D")), 147.75);

  assertEqual(X3LOOKUP("TSLA", dec2028, 999, mockRange("A"), mockRange("B"), mockRange("C"), mockRange("D")), "#N/A");
  console.log("✅ X3LOOKUP test passed");
}

