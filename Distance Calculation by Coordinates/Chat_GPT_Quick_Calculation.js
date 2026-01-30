/**
 * Bulk calculate truckable distance (NO API)
 * Uses Haversine × Road Factor
 * Optimized for 12,000+ rows
 */

function CALCULATE_TRUCKABLE_DISTANCE() {

  const ROAD_FACTOR = 1.25; // adjust if needed
  const EARTH_RADIUS = 6371; // km

  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length < 2) return;

  // Read headers
  const headers = values[0];

  const fromLatCol = headers.indexOf("From_Lat");
  const fromLngCol = headers.indexOf("From_Lng");
  const toLatCol   = headers.indexOf("To_Lat");
  const toLngCol   = headers.indexOf("To_Lng");
  const distCol    = headers.indexOf("Distance in KM");

  if (
    fromLatCol === -1 || fromLngCol === -1 ||
    toLatCol === -1   || toLngCol === -1 ||
    distCol === -1
  ) {
    throw new Error("Required columns not found. Check header names.");
  }

  // Process rows in memory
  for (let i = 1; i < values.length; i++) {

    const fromLat = values[i][fromLatCol];
    const fromLng = values[i][fromLngCol];
    const toLat   = values[i][toLatCol];
    const toLng   = values[i][toLngCol];

    if (!fromLat || !fromLng || !toLat || !toLng) {
      values[i][distCol] = "";
      continue;
    }

    const dLat = toRad(toLat - fromLat);
    const dLng = toRad(toLng - fromLng);

    const a =
      Math.sin(dLat / 2) ** 2 +
      Math.cos(toRad(fromLat)) *
      Math.cos(toRad(toLat)) *
      Math.sin(dLng / 2) ** 2;

    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

    const straightKm = EARTH_RADIUS * c;
    const truckableKm = straightKm * ROAD_FACTOR;

    values[i][distCol] = Math.round(truckableKm * 100) / 100;
  }

  // Write back once (VERY FAST)
  dataRange.setValues(values);

  SpreadsheetApp.getUi().alert(
    "Done",
    "Truckable distance calculated for " + (values.length - 1) + " rows.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Degree → Radian
function toRad(deg) {
  return deg * Math.PI / 180;
}
