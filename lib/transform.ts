/**
 * Maps a raw SharePoint list item (Graph `fields` object) into the typed shape
 * we persist, including the Price-string → integer conversion.
 */

/**
 * SharePoint column *internal* names.
 *
 * IMPORTANT: Graph returns columns under `fields` keyed by their INTERNAL name,
 * which is NOT always the display name you see in the SharePoint UI. Spaces are
 * usually encoded (e.g. "Full Name" -> "FullName" or "Full_x0020_Name"), and
 * columns created long ago can have opaque names like "field_1".
 *
 * Run `npm run inspect:columns` once against your list to print the real
 * internal names, then adjust the right-hand side below to match.
 */
export const FIELD_MAP = {
  fullName: "FullName",
  phone: "Phone",
  age: "Age",
  gender: "Gender",
  email: "Email",
  idNumber: "IDNumber",
  address: "Address",
  shopName: "ShopName",
  price: "Price",
  notes: "Notes",
  numberOfCoupons: "NumberOfCoupons",
  receiptNumber: "ReceiptNumber",
  couponCode: "CouponCode",
  submissionDate: "SubmissionDate",
} as const;

export interface SyncEntry {
  sharepointId: string;
  fullName: string | null;
  phone: string | null;
  age: number | null;
  gender: string | null;
  email: string | null;
  idNumber: string | null;
  address: string | null;
  shopName: string | null;
  price: number | null;
  notes: string | null;
  numberOfCoupons: number | null;
  receiptNumber: string | null;
  couponCode: string | null;
  submissionDate: string | null; // ISO 8601 or null
}

/**
 * Parses a formatted currency string into a pure integer.
 *
 * Examples:
 *   "200,750 IQD" -> 200750
 *   "1,000"       -> 1000
 *   "IQD 50000"   -> 50000
 *   "75.000 IQD"  -> 75000   (treats '.' as a thousands separator too)
 *   ""            -> null
 *
 * Strategy: discard every character that is not a digit. For IQD (a
 * zero-decimal currency) thousands separators — whether ',' or '.' — and the
 * currency label are all noise, so stripping non-digits yields the integer
 * amount directly.
 */
export function parsePriceToInteger(raw: unknown): number | null {
  if (raw === null || raw === undefined) return null;

  // Already numeric (Graph sometimes returns Number for numeric columns).
  if (typeof raw === "number") {
    return Number.isFinite(raw) ? Math.round(raw) : null;
  }

  const text = String(raw).trim();
  if (text === "") return null;

  const digits = text.replace(/[^\d]/g, "");
  if (digits === "") return null;

  const value = Number.parseInt(digits, 10);
  return Number.isNaN(value) ? null : value;
}

function asString(value: unknown): string | null {
  if (value === null || value === undefined) return null;
  const s = String(value).trim();
  return s === "" ? null : s;
}

function asInteger(value: unknown): number | null {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? Math.round(n) : null;
}

function asIsoDate(value: unknown): string | null {
  const s = asString(value);
  if (s === null) return null;
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d.toISOString();
}

/** A SharePoint item as returned by Graph (`expand=fields`). */
export interface GraphListItem {
  id: string;
  fields?: Record<string, unknown>;
}

export function mapItemToEntry(item: GraphListItem): SyncEntry {
  const f = item.fields ?? {};
  return {
    sharepointId: String(item.id),
    fullName: asString(f[FIELD_MAP.fullName]),
    phone: asString(f[FIELD_MAP.phone]),
    age: asInteger(f[FIELD_MAP.age]),
    gender: asString(f[FIELD_MAP.gender]),
    email: asString(f[FIELD_MAP.email]),
    idNumber: asString(f[FIELD_MAP.idNumber]),
    address: asString(f[FIELD_MAP.address]),
    shopName: asString(f[FIELD_MAP.shopName]),
    price: parsePriceToInteger(f[FIELD_MAP.price]),
    notes: asString(f[FIELD_MAP.notes]),
    numberOfCoupons: asInteger(f[FIELD_MAP.numberOfCoupons]),
    receiptNumber: asString(f[FIELD_MAP.receiptNumber]),
    couponCode: asString(f[FIELD_MAP.couponCode]),
    submissionDate: asIsoDate(f[FIELD_MAP.submissionDate]),
  };
}
