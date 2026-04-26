/**
 * Rental income importer.
 *
 * Drop CSV exports from Airbnb / Vrbo / Guesty into ONEDRIVE_RENTALS_INBOX.
 * The handler detects the format and writes normalized rows to rentals_YYYY.xlsx.
 *
 * Format detection by header signature:
 *   - Airbnb transaction history: contains "Confirmation Code", "Listing", "Type"
 *   - Vrbo earnings: contains "Reservation ID", "Property", "Net rental income"
 *   - Guesty Lite payout: contains "Reservation", "Check-in", "Net amount"
 */

export interface NormalizedRental {
  booking_id: string;
  platform: 'Airbnb' | 'Vrbo' | 'Direct' | 'Guesty';
  property: string;
  guest_name: string;
  check_in: string; // YYYY-MM-DD
  check_out: string;
  nights: number;
  gross_revenue: number;
  platform_fees: number; // positive number
  cleaning_fees_collected: number;
  taxes_collected: number;
  net_payout: number;
  payout_date: string; // YYYY-MM-DD
  entity: string;
  notes: string;
}

export function detectFormat(headers: string[]): 'airbnb' | 'vrbo' | 'guesty' | 'unknown' {
  const lc = headers.map((h) => h.toLowerCase().trim());
  if (lc.includes('confirmation code') && lc.includes('listing')) return 'airbnb';
  if (lc.some((h) => h.includes('reservation id')) && lc.some((h) => h.includes('property')))
    return 'vrbo';
  if (lc.includes('reservation') && lc.some((h) => h.includes('net amount'))) return 'guesty';
  return 'unknown';
}

export function parseCsv(content: string): { headers: string[]; rows: string[][] } {
  const lines = content.split(/\r?\n/).filter((l) => l.trim().length > 0);
  if (lines.length === 0) return { headers: [], rows: [] };
  const headers = parseCsvLine(lines[0]);
  const rows = lines.slice(1).map(parseCsvLine);
  return { headers, rows };
}

function parseCsvLine(line: string): string[] {
  const out: string[] = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else inQuotes = !inQuotes;
    } else if (ch === ',' && !inQuotes) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

function getCol(
  headers: string[],
  row: string[],
  ...names: string[]
): string {
  for (const name of names) {
    const idx = headers.findIndex(
      (h) => h.toLowerCase().trim() === name.toLowerCase(),
    );
    if (idx >= 0) return (row[idx] || '').trim();
  }
  return '';
}

function num(s: string): number {
  if (!s) return 0;
  const cleaned = s.replace(/[$,]/g, '').replace(/\((.*)\)/, '-$1');
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function isoDate(s: string): string {
  if (!s) return '';
  const d = new Date(s);
  if (isNaN(d.getTime())) return s;
  return d.toISOString().slice(0, 10);
}

export function normalizeAirbnb(headers: string[], row: string[]): NormalizedRental | null {
  const type = getCol(headers, row, 'Type').toLowerCase();
  if (type !== 'reservation' && type !== 'payout') return null;
  if (type === 'payout') return null; // skip aggregate payouts; we want per-reservation

  const checkIn = isoDate(getCol(headers, row, 'Start Date', 'Check-in'));
  const checkOut = isoDate(getCol(headers, row, 'End Date', 'Check-out'));
  const nights = parseInt(getCol(headers, row, 'Nights') || '0', 10);
  const gross = num(getCol(headers, row, 'Gross Earnings', 'Amount'));
  const fees = num(getCol(headers, row, 'Service Fee', 'Host Fee'));
  const cleaning = num(getCol(headers, row, 'Cleaning Fee'));
  const taxes = num(getCol(headers, row, 'Occupancy Taxes', 'Taxes'));
  const net = num(getCol(headers, row, 'Earnings', 'Paid Out'));
  const payoutDate = isoDate(getCol(headers, row, 'Date', 'Payout Date'));

  return {
    booking_id: getCol(headers, row, 'Confirmation Code'),
    platform: 'Airbnb',
    property: getCol(headers, row, 'Listing'),
    guest_name: getCol(headers, row, 'Guest'),
    check_in: checkIn,
    check_out: checkOut,
    nights,
    gross_revenue: gross,
    platform_fees: Math.abs(fees),
    cleaning_fees_collected: cleaning,
    taxes_collected: taxes,
    net_payout: net || gross - Math.abs(fees),
    payout_date: payoutDate,
    entity: 'Sol Haus',
    notes: '',
  };
}

export function normalizeVrbo(headers: string[], row: string[]): NormalizedRental | null {
  const id = getCol(headers, row, 'Reservation ID', 'Reservation');
  if (!id) return null;
  return {
    booking_id: id,
    platform: 'Vrbo',
    property: getCol(headers, row, 'Property'),
    guest_name: getCol(headers, row, 'Guest', 'Traveler Name'),
    check_in: isoDate(getCol(headers, row, 'Check-in', 'Arrival')),
    check_out: isoDate(getCol(headers, row, 'Check-out', 'Departure')),
    nights: parseInt(getCol(headers, row, 'Nights') || '0', 10),
    gross_revenue: num(getCol(headers, row, 'Gross Booking Amount', 'Total')),
    platform_fees: Math.abs(num(getCol(headers, row, 'Vrbo Fees', 'Service Fee'))),
    cleaning_fees_collected: num(getCol(headers, row, 'Cleaning Fee')),
    taxes_collected: num(getCol(headers, row, 'Taxes')),
    net_payout: num(getCol(headers, row, 'Net rental income', 'Payout')),
    payout_date: isoDate(getCol(headers, row, 'Payment Date', 'Payout Date')),
    entity: 'Sol Haus',
    notes: '',
  };
}

export function normalizeGuesty(headers: string[], row: string[]): NormalizedRental | null {
  const id = getCol(headers, row, 'Reservation', 'Reservation ID');
  if (!id) return null;
  const channel = getCol(headers, row, 'Channel', 'Source').toLowerCase();
  const platform: NormalizedRental['platform'] = channel.includes('airbnb')
    ? 'Airbnb'
    : channel.includes('vrbo') || channel.includes('homeaway')
      ? 'Vrbo'
      : channel.includes('direct')
        ? 'Direct'
        : 'Guesty';

  return {
    booking_id: id,
    platform,
    property: getCol(headers, row, 'Property', 'Listing'),
    guest_name: getCol(headers, row, 'Guest'),
    check_in: isoDate(getCol(headers, row, 'Check-in')),
    check_out: isoDate(getCol(headers, row, 'Check-out')),
    nights: parseInt(getCol(headers, row, 'Nights') || '0', 10),
    gross_revenue: num(getCol(headers, row, 'Total', 'Gross')),
    platform_fees: Math.abs(num(getCol(headers, row, 'Channel Fee', 'Commission'))),
    cleaning_fees_collected: num(getCol(headers, row, 'Cleaning Fee')),
    taxes_collected: num(getCol(headers, row, 'Tax', 'Taxes')),
    net_payout: num(getCol(headers, row, 'Net amount', 'Payout')),
    payout_date: isoDate(getCol(headers, row, 'Payout Date', 'Date')),
    entity: 'Sol Haus',
    notes: `via Guesty (${channel || 'unknown channel'})`,
  };
}

export function normalize(
  format: 'airbnb' | 'vrbo' | 'guesty',
  headers: string[],
  row: string[],
): NormalizedRental | null {
  if (format === 'airbnb') return normalizeAirbnb(headers, row);
  if (format === 'vrbo') return normalizeVrbo(headers, row);
  return normalizeGuesty(headers, row);
}
