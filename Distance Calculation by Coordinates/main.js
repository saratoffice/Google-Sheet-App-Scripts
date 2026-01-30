/**
 * Custom function for Sheets:
 * Returns Google Maps DRIVING DISTANCE in KM between two coordinates.
 *
 * Usage in Sheets:
 * =GET_DRIVING_DISTANCE(D2, E2, H2, I2)
 */
function GET_DRIVING_DISTANCE(fromLat, fromLng, toLat, toLng) {
  try {
    var origin = fromLat + "," + fromLng;
    var dest = toLat + "," + toLng;

    var directions = Maps.newDirectionFinder()
      .setOrigin(origin)
      .setDestination(dest)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    var meters = directions.routes[0].legs[0].distance.value;
    return meters / 1000; // convert meters → km
  }
  catch (e) {
    return "ERR"; // if invalid or API quota exceeded
  }
}

/**
 * Bulk fill distances for 1000+ rows automatically.
 * Assumes your From_Lat is column D (4), From_Lng E (5),
 * To_Lat H (8), To_Lng I (9), output column J (10)
 */
function FILL_ALL_DISTANCES() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();

  for (var row = 2; row <= lastRow; row++) {
    var fromLat = sheet.getRange(row, 4).getValue();
    var fromLng = sheet.getRange(row, 5).getValue();
    var toLat   = sheet.getRange(row, 8).getValue();
    var toLng   = sheet.getRange(row, 9).getValue();

    if (fromLat && fromLng && toLat && toLng) {
      var km = GET_DRIVING_DISTANCE(fromLat, fromLng, toLat, toLng);
      sheet.getRange(row, 10).setValue(km);
    }
  }
}
