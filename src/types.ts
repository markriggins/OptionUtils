/**
 * Type definitions for OptionUtils - Phase 4
 * @file src/types.ts
 */

interface SpreadFinderConfig {
  minROI?: number;
  patience?: number;
  minLiquidity?: number;
  maxSpreadWidth?: number;
  minExpectedGain?: number;
  [key: string]: any;  // allow your existing fields
}

interface SpreadResult {
  symbol: string;
  expiration: string;
  longStrike: number;
  shortStrike: number;
  expectedGain?: number;
  roi?: number;
  // add any other fields your outputSpreadResults_ uses
}

