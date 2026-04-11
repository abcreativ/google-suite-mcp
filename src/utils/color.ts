/**
 * Color conversion utilities for Google Sheets.
 *
 * Google Sheets API represents colors as { red, green, blue, alpha } where
 * each component is a float in the range [0, 1].
 */

export interface SheetsColor {
  red: number;
  green: number;
  blue: number;
  alpha?: number;
}

/** Named CSS colors mapped to their hex equivalents. */
const NAMED_COLORS: Record<string, string> = {
  black: "#000000",
  white: "#ffffff",
  red: "#ff0000",
  green: "#008000",
  blue: "#0000ff",
  yellow: "#ffff00",
  orange: "#ffa500",
  purple: "#800080",
  pink: "#ffc0cb",
  cyan: "#00ffff",
  magenta: "#ff00ff",
  lime: "#00ff00",
  navy: "#000080",
  teal: "#008080",
  maroon: "#800000",
  olive: "#808000",
  silver: "#c0c0c0",
  gray: "#808080",
  grey: "#808080",
  aqua: "#00ffff",
  fuchsia: "#ff00ff",
  transparent: "#00000000",
};

/**
 * Parses a hex or named color string into a Google Sheets Color object.
 *
 * Supported formats:
 *   - "#RGB"       - 3-digit hex shorthand (e.g. "#F00")
 *   - "#RRGGBB"    - 6-digit hex (e.g. "#FF0000")
 *   - "#RRGGBBAA"  - 8-digit hex with alpha (e.g. "#FF0000FF")
 *   - Named colors (e.g. "red", "blue", "transparent")
 *
 * All returned values are in the [0, 1] range as required by the Sheets API.
 */
export function parseColor(color: string): SheetsColor {
  const input = color.trim().toLowerCase();

  // Resolve named colors to hex
  const resolved = NAMED_COLORS[input] ?? input;

  if (!resolved.startsWith("#")) {
    throw new Error(
      `Unsupported color format: "${color}". Use hex (#RGB, #RRGGBB, #RRGGBBAA) or a named color.`
    );
  }

  const hex = resolved.slice(1); // strip leading #

  let r: number, g: number, b: number, a: number | undefined;

  if (hex.length === 3) {
    // #RGB shorthand → #RRGGBB
    r = parseInt(hex[0] + hex[0], 16);
    g = parseInt(hex[1] + hex[1], 16);
    b = parseInt(hex[2] + hex[2], 16);
  } else if (hex.length === 6) {
    r = parseInt(hex.slice(0, 2), 16);
    g = parseInt(hex.slice(2, 4), 16);
    b = parseInt(hex.slice(4, 6), 16);
  } else if (hex.length === 8) {
    r = parseInt(hex.slice(0, 2), 16);
    g = parseInt(hex.slice(2, 4), 16);
    b = parseInt(hex.slice(4, 6), 16);
    a = parseInt(hex.slice(6, 8), 16) / 255;
  } else {
    throw new Error(
      `Invalid hex color: "${color}". Expected #RGB, #RRGGBB, or #RRGGBBAA.`
    );
  }

  const result: SheetsColor = {
    red: r / 255,
    green: g / 255,
    blue: b / 255,
  };

  if (a !== undefined) {
    result.alpha = a;
  }

  return result;
}
