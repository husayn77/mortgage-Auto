// Converts a number into Nigerian Naira / Kobo words using Title Case,
// hyphenated tens (e.g. "Thirty-Eight") and comma-separated scale groups,
// matching the exact style used in the Excel template, e.g.
// 1238883.4676 -> "One Million, Two Hundred Thirty-Eight Thousand, Eight Hundred Eighty-Three Naira and Forty-Seven Kobo"

const ONES = [
  "",
  "One",
  "Two",
  "Three",
  "Four",
  "Five",
  "Six",
  "Seven",
  "Eight",
  "Nine",
  "Ten",
  "Eleven",
  "Twelve",
  "Thirteen",
  "Fourteen",
  "Fifteen",
  "Sixteen",
  "Seventeen",
  "Eighteen",
  "Nineteen",
];

const TENS = [
  "",
  "",
  "Twenty",
  "Thirty",
  "Forty",
  "Fifty",
  "Sixty",
  "Seventy",
  "Eighty",
  "Ninety",
];

const SCALES = ["", "Thousand", "Million", "Billion", "Trillion", "Quadrillion"];

function twoDigits(n: number): string {
  if (n < 20) return ONES[n];
  const t = Math.floor(n / 10);
  const o = n % 10;
  return o > 0 ? `${TENS[t]}-${ONES[o]}` : TENS[t];
}

function threeDigits(n: number): string {
  const h = Math.floor(n / 100);
  const rem = n % 100;
  const parts: string[] = [];
  if (h > 0) parts.push(`${ONES[h]} Hundred`);
  if (rem > 0) parts.push(twoDigits(rem));
  return parts.join(" ");
}

function integerToWords(value: number): string {
  if (value === 0) return "Zero";

  const groups: number[] = [];
  let n = value;
  while (n > 0) {
    groups.push(n % 1000);
    n = Math.floor(n / 1000);
  }

  const parts: string[] = [];
  for (let i = groups.length - 1; i >= 0; i--) {
    const g = groups[i];
    if (g === 0) continue;
    const scale = SCALES[i];
    parts.push(scale ? `${threeDigits(g)} ${scale}` : threeDigits(g));
  }

  return parts.join(", ");
}

/**
 * Convert a monetary amount to words in Naira and Kobo.
 * Returns e.g. "One Million ... Naira and Forty-Seven Kobo".
 */
export function amountToWords(amount: number | null | undefined): string {
  if (amount === null || amount === undefined || isNaN(Number(amount))) {
    return "";
  }

  const value = Math.abs(Number(amount));
  const naira = Math.floor(value);
  let kobo = Math.round((value - naira) * 100);

  let nairaForWords = naira;
  // Handle rounding overflow (e.g. 0.999 -> 100 kobo)
  if (kobo === 100) {
    nairaForWords += 1;
    kobo = 0;
  }

  const nairaWords = integerToWords(nairaForWords);

  if (kobo > 0) {
    return `${nairaWords} Naira and ${twoDigits(kobo)} Kobo`;
  }
  return `${nairaWords} Naira Only`;
}
