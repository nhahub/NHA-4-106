/**
 * RAW2INSIGHT - Supply Chain Logistics Intelligence Dashboard
 * Apps Script Backend Ecosystem
 *
 * ============================================================================
 * FIELD-NAME ASSUMPTIONS (adjust the candidate arrays below if your sheet
 * headers differ - the pickField() helper tries each candidate in order):
 *   loads:                 revenue, fuel_surcharge, accessorial_charges,
 *                           weight_lbs, customer_id, booking_type, load_type,
 *                           destination_state, pickup_date/load_date/date
 *   trips:                 truck_id, load_id, actual_distance_miles,
 *                           fuel_gallons_used, trip_date/date
 *   trucks:                truck_id, make
 *   customers:              customer_id, customer_name, customer_type
 *   driverMonthlyMetrics:  first_name, trips_completed, average_idle_hours,
 *                           on_time_delivery_rate, total_miles,
 *                           month/YearMonth/date
 *   truckUtilizationMetrics: truck_id, total_miles, maintenance_cost,
 *                           month/YearMonth/date
 *   maintenanceRecords:    truck_id, total_cost, maintenance_type,
 *                           location_state, maintenance_date/date
 *   deliveryEvents:        load_id, on_time_flag, location_city,
 *                           detention_minutes
 *   fuelPurchases:         truck_id, total_cost, location_state,
 *                           purchase_date/date
 * ============================================================================
 */

// CONSTANTS: Replace with your actual Google Spreadsheet IDs
const SPREADSHEET_A_ID = "1xyUXii3g2u1p55VUUA8BjAP5uhJzwg2Jf7WdKSpJmyI";    // Contains: loads, trips, trucks
const SPREADSHEET_B_ID = "1mwOyWryvlPSdeEi9_ckzo7iuoQ41FqoTC8H4ybBMyKA";   // Contains: customers, deliveryEvents, driverMonthlyMetrics, truckUtilizationMetrics, maintenanceRecords, fuelPurchases

// Google Drive file ID of the RAW2INSIGHT logo image
const LOGO_FILE_ID = "1li1_RTr8NBVmYUwp7bdYP-BPJX4bIOOs";

/**
 * Serves the core HTML layout framework
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('dashboard');
  template.logoDataUri = getLogoDataUri();

  return template.evaluate()
      .setTitle('RAW2INSIGHT – Supply Chain Logistics Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Reads the logo directly from Drive and returns it as a base64 data URI.
 *
 * Why this instead of the old <img src="https://drive.google.com/uc?export=view&id=...">
 * link: that URL format is a raw hotlink, and Drive frequently refuses to serve it
 * inline - it can return an HTML "can't scan this file for viruses" interstitial page,
 * a 403, or nothing at all, especially inside the sandboxed iframe that Apps Script
 * HtmlService renders in. None of that shows up as a normal broken-image icon; the
 * <img> tag just silently fails. Fetching the file server-side with DriveApp and
 * inlining it as base64 sidesteps all of that - the browser never talks to Drive at
 * all, so there's nothing for Drive to block.
 *
 * Requires: the account running this script must have at least "Viewer" access to
 * LOGO_FILE_ID (owns it, or it's shared with them / with "Anyone with the link").
 */
function getLogoDataUri() {
  try {
    const file = DriveApp.getFileById(LOGO_FILE_ID);
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    return "data:" + mimeType + ";base64," + base64;
  } catch (err) {
    Logger.log("Logo load failed: " + err.toString());
    return ""; // dashboard.html falls back to the fa-chart-line icon when this is empty
  }
}

// ==========================================================================
// DATA LOADING
// ==========================================================================

function loadAllData() {
  const ssA = SpreadsheetApp.openById(SPREADSHEET_A_ID);
  const ssB = SpreadsheetApp.openById(SPREADSHEET_B_ID);

  return {
    loads:          getSheetData(ssA, "loads"),
    trips:          getSheetData(ssA, "trips"),
    trucks:         getSheetData(ssA, "trucks"),
    customers:      getSheetData(ssB, "customers"),
    driverMonthly:  getSheetData(ssB, "driverMonthlyMetrics"),
    truckUtil:      getSheetData(ssB, "truckUtilizationMetrics"),
    maintenance:    getSheetData(ssB, "maintenanceRecords"),
    deliveryEvents: getSheetData(ssB, "deliveryEvents"),
    fuelPurchases:  getSheetData(ssB, "fuelPurchases")
  };
}

/**
 * Helper: Extracts rows as JSON key-value objects matching header values
 */
function getSheetData(ssConnection, sheetName) {
  const sheet = ssConnection.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => h.toString().trim());
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    let rowObj = {};
    for (let j = 0; j < headers.length; j++) {
      rowObj[headers[j]] = data[i][j];
    }
    rows.push(rowObj);
  }
  return rows;
}

// ==========================================================================
// GENERIC UTILITY HELPERS
// ==========================================================================

function pickField(row, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const v = row[candidates[i]];
    if (v !== undefined && v !== null && v !== "") return v;
  }
  return null;
}

function num(row, candidates) {
  const v = pickField(row, candidates);
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function toDate(val) {
  if (val instanceof Date) return val;
  if (typeof val === "string" || typeof val === "number") {
    const d = new Date(val);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function quarterLabel(dateVal) {
  const d = toDate(dateVal);
  if (!d) return null;
  return "Q" + (Math.floor(d.getMonth() / 3) + 1) + " " + d.getFullYear();
}

function yearMonthLabel(dateVal) {
  const d = toDate(dateVal);
  if (!d) return null;
  return MONTH_NAMES[d.getMonth()] + " " + d.getFullYear();
}

// Sortable key so "Jan 2024" < "Feb 2024" < "Jan 2025" etc.
function yearMonthSortKey(label) {
  if (!label) return "";
  const parts = label.split(" ");
  const mIdx = MONTH_NAMES.indexOf(parts[0]);
  return parts[1] + "-" + (mIdx < 10 ? "0" + mIdx : mIdx);
}

// Passes if the filter is empty/"All", otherwise requires exact match
function passFilter(value, filterValue) {
  if (filterValue === undefined || filterValue === null || filterValue === "" || filterValue === "All") return true;
  return String(value) === String(filterValue);
}

function uniqueSorted(arr) {
  return Array.from(new Set(arr.filter(v => v !== null && v !== undefined && v !== ""))).sort();
}

function groupSum(rows, keyFn, valFn) {
  const map = {};
  rows.forEach(r => {
    const k = keyFn(r);
    if (k === null || k === undefined || k === "") return;
    map[k] = (map[k] || 0) + valFn(r);
  });
  return map;
}

function sortMapDesc(map, limit) {
  const arr = Object.keys(map).map(k => ({ label: k, value: map[k] })).sort((a, b) => b.value - a.value);
  return limit ? arr.slice(0, limit) : arr;
}

// ==========================================================================
// FILTER OPTIONS (populates all dropdowns on initial dashboard load)
// ==========================================================================

function getFilterOptions() {
  const d = loadAllData();

  const truckMakeMap = {};
  d.trucks.forEach(t => { truckMakeMap[pickField(t, ["truck_id"])] = pickField(t, ["make"]); });

  const quarters = uniqueSorted(d.loads.map(r => quarterLabel(pickField(r, ["pickup_date", "load_date", "date"]))));
  const customerTypes = uniqueSorted(d.customers.map(r => pickField(r, ["customer_type"])));
  const makes = uniqueSorted(d.trucks.map(r => pickField(r, ["make"])));
  const loadTypes = uniqueSorted(d.loads.map(r => pickField(r, ["load_type"])));

  const yearMonths = uniqueSorted(
    d.truckUtil.map(r => yearMonthLabel(pickField(r, ["month", "YearMonth", "date"])))
      .concat(d.maintenance.map(r => yearMonthLabel(pickField(r, ["maintenance_date", "date"]))))
  ).sort((a, b) => yearMonthSortKey(a).localeCompare(yearMonthSortKey(b)));

  const maintenanceTypes = uniqueSorted(d.maintenance.map(r => pickField(r, ["maintenance_type"])));
  const locationStates = uniqueSorted(
    d.maintenance.map(r => pickField(r, ["location_state"]))
      .concat(d.fuelPurchases.map(r => pickField(r, ["location_state"])))
  );

  return { quarters, customerTypes, makes, loadTypes, yearMonths, maintenanceTypes, locationStates };
}

// ==========================================================================
// STRATEGIC DASHBOARD  (filters: quarter, customerType)
// ==========================================================================

function getStrategicData(filters) {
  filters = filters || {};
  const d = loadAllData();

  const customerTypeMap = {};
  const customerNameMap = {};
  d.customers.forEach(c => {
    const id = pickField(c, ["customer_id"]);
    customerTypeMap[id] = pickField(c, ["customer_type"]);
    customerNameMap[id] = pickField(c, ["customer_name"]);
  });

  // ---- Filter loads by quarter + customer_type ----
  const filteredLoads = d.loads.filter(row => {
    const custId = pickField(row, ["customer_id"]);
    const q = quarterLabel(pickField(row, ["pickup_date", "load_date", "date"]));
    return passFilter(q, filters.quarter) && passFilter(customerTypeMap[custId], filters.customerType);
  });

  let totalRevenue = 0;
  const revenueByBookingType = {};
  const revenueByLoadType = {};
  const revenueByCustomerName = {};
  const revenueByDestState = {};
  const revenueByYearMonth = {};

  filteredLoads.forEach(row => {
    const rev = num(row, ["revenue"]) + num(row, ["fuel_surcharge"]) + num(row, ["accessorial_charges"]);
    totalRevenue += rev;

    const booking = pickField(row, ["booking_type"]) || "Unknown";
    revenueByBookingType[booking] = (revenueByBookingType[booking] || 0) + rev;

    const lt = pickField(row, ["load_type"]) || "Unknown";
    revenueByLoadType[lt] = (revenueByLoadType[lt] || 0) + rev;

    const custId = pickField(row, ["customer_id"]);
    const custName = customerNameMap[custId] || custId || "Unknown";
    revenueByCustomerName[custName] = (revenueByCustomerName[custName] || 0) + rev;

    const state = pickField(row, ["destination_state"]) || "Unknown";
    revenueByDestState[state] = (revenueByDestState[state] || 0) + rev;

    const ym = yearMonthLabel(pickField(row, ["pickup_date", "load_date", "date"]));
    if (ym) revenueByYearMonth[ym] = (revenueByYearMonth[ym] || 0) + rev;
  });

  const loadsCount = filteredLoads.length;

  // ---- Costs (Total Costs = fuel_purchases.total_cost + truck_utilization_metrics.maintenance_cost) ----
  // Quarter filter applies to cost tables via their own date fields; customer_type has no
  // relationship to cost tables so it is intentionally NOT applied here (mirrors Power BI
  // behaviour when there is no relationship path between the slicer table and the fact table).
  let totalCosts = 0;
  const costByYearMonth = {};

  d.fuelPurchases.forEach(row => {
    const q = quarterLabel(pickField(row, ["purchase_date", "date"]));
    if (!passFilter(q, filters.quarter)) return;
    const cost = num(row, ["total_cost"]);
    totalCosts += cost;
    const ym = yearMonthLabel(pickField(row, ["purchase_date", "date"]));
    if (ym) costByYearMonth[ym] = (costByYearMonth[ym] || 0) + cost;
  });

  d.truckUtil.forEach(row => {
    const q = quarterLabel(pickField(row, ["month", "YearMonth", "date"]));
    if (!passFilter(q, filters.quarter)) return;
    const cost = num(row, ["maintenance_cost"]);
    totalCosts += cost;
    const ym = yearMonthLabel(pickField(row, ["month", "YearMonth", "date"]));
    if (ym) costByYearMonth[ym] = (costByYearMonth[ym] || 0) + cost;
  });

  const grossProfit = totalRevenue - totalCosts;
  const grossProfitMargin = totalRevenue > 0 ? (grossProfit / totalRevenue) * 100 : 0;
  const revenuePerLoad = loadsCount > 0 ? (totalRevenue / loadsCount) : 0;
  const profitPerLoad = loadsCount > 0 ? (grossProfit / loadsCount) : 0;

  // Revenue per Mile = Total Revenue / SUM(driver_monthly_metrics[total_miles]), quarter-filtered
  let sumDriverMiles = 0;
  d.driverMonthly.forEach(row => {
    const q = quarterLabel(pickField(row, ["month", "YearMonth", "date"]));
    if (!passFilter(q, filters.quarter)) return;
    sumDriverMiles += num(row, ["total_miles"]);
  });
  const revenuePerMile = sumDriverMiles > 0 ? (totalRevenue / sumDriverMiles) : 0;

  // Top Customer Revenue %
  let topCustomerRevenue = 0;
  Object.keys(revenueByCustomerName).forEach(name => {
    if (revenueByCustomerName[name] > topCustomerRevenue) topCustomerRevenue = revenueByCustomerName[name];
  });
  const topCustomerRevenuePct = totalRevenue > 0 ? (topCustomerRevenue / totalRevenue) * 100 : 0;

  // Gross Profit Margin by destination_state
  // (TotalCosts has no relationship to destination_state, so the same TotalCosts figure
  // is subtracted from every state's revenue slice - matching how the DAX measure would
  // evaluate without a state-aware cost relationship.)
  const gpmByDestState = {};
  Object.keys(revenueByDestState).forEach(state => {
    const rev = revenueByDestState[state];
    gpmByDestState[state] = rev > 0 ? ((rev - totalCosts) / rev) * 100 : 0;
  });

  // Gross Profit Margin and Total Revenue by YearMonth
  const monthKeys = uniqueSorted(Object.keys(revenueByYearMonth).concat(Object.keys(costByYearMonth)))
    .sort((a, b) => yearMonthSortKey(a).localeCompare(yearMonthSortKey(b)));
  const gpmSeries = monthKeys.map(m => {
    const rev = revenueByYearMonth[m] || 0;
    const cost = costByYearMonth[m] || 0;
    return rev > 0 ? ((rev - cost) / rev) * 100 : 0;
  });
  const revSeries = monthKeys.map(m => revenueByYearMonth[m] || 0);

  return {
    kpis: {
      totalRevenue: totalRevenue,
      grossProfitMargin: grossProfitMargin,
      revenuePerLoad: revenuePerLoad,
      profitPerLoad: profitPerLoad,
      revenuePerMile: revenuePerMile,
      topCustomerRevenuePct: topCustomerRevenuePct
    },
    charts: {
      revenueByBookingType: sortMapDesc(revenueByBookingType),
      topCustomerRevenuePct: topCustomerRevenuePct,
      gpmByDestState: sortMapDesc(gpmByDestState),
      yearMonthLabels: monthKeys,
      gpmSeries: gpmSeries,
      revSeries: revSeries,
      revenueByLoadType: sortMapDesc(revenueByLoadType),
      revenueByCustomerName: sortMapDesc(revenueByCustomerName, 8)
    }
  };
}

// ==========================================================================
// OPERATIONAL DASHBOARD  (filters: quarter, make, loadType)
// ==========================================================================

function getOperationalData(filters) {
  filters = filters || {};
  const d = loadAllData();

  const truckMakeMap = {};
  d.trucks.forEach(t => { truckMakeMap[pickField(t, ["truck_id"])] = pickField(t, ["make"]); });

  const loadTypeMap = {};
  const loadQuarterMap = {};
  d.loads.forEach(l => {
    const id = pickField(l, ["load_id"]);
    loadTypeMap[id] = pickField(l, ["load_type"]);
    loadQuarterMap[id] = quarterLabel(pickField(l, ["pickup_date", "load_date", "date"]));
  });

  // ---- Count of load_id by YearMonth (quarter + loadType filters) ----
  const loadCountByYearMonth = {};
  d.loads.forEach(row => {
    const q = quarterLabel(pickField(row, ["pickup_date", "load_date", "date"]));
    const lt = pickField(row, ["load_type"]);
    if (!passFilter(q, filters.quarter) || !passFilter(lt, filters.loadType)) return;
    const ym = yearMonthLabel(pickField(row, ["pickup_date", "load_date", "date"]));
    if (ym) loadCountByYearMonth[ym] = (loadCountByYearMonth[ym] || 0) + 1;
  });

  // ---- Average Load Weight (quarter + loadType filters) ----
  let weightSum = 0, weightCount = 0;
  d.loads.forEach(row => {
    const q = quarterLabel(pickField(row, ["pickup_date", "load_date", "date"]));
    const lt = pickField(row, ["load_type"]);
    if (!passFilter(q, filters.quarter) || !passFilter(lt, filters.loadType)) return;
    weightSum += num(row, ["weight_lbs"]);
    weightCount++;
  });
  const avgLoadWeight = weightCount > 0 ? (weightSum / weightCount) : 0;

  // ---- driverMonthlyMetrics: OTD rate, Average Idle Hours, trips completed, table by first_name ----
  let idleWeighted = 0, otdWeighted = 0, tripsCompletedTotal = 0;
  const idleByFirstName = {}; // { name: { weighted, trips } }
  const tripsByYearMonth = {};
  const otdByYearMonth = {}; // weighted sums, divided after

  d.driverMonthly.forEach(row => {
    const q = quarterLabel(pickField(row, ["month", "YearMonth", "date"]));
    if (!passFilter(q, filters.quarter)) return;

    const trips = num(row, ["trips_completed"]);
    const idle = num(row, ["average_idle_hours"]);
    const otd = num(row, ["on_time_delivery_rate"]);
    const name = pickField(row, ["first_name"]) || "Unknown";
    const ym = yearMonthLabel(pickField(row, ["month", "YearMonth", "date"]));

    idleWeighted += idle * trips;
    otdWeighted += otd * trips;
    tripsCompletedTotal += trips;

    if (!idleByFirstName[name]) idleByFirstName[name] = { weighted: 0, trips: 0 };
    idleByFirstName[name].weighted += idle * trips;
    idleByFirstName[name].trips += trips;

    if (ym) {
      if (!tripsByYearMonth[ym]) tripsByYearMonth[ym] = 0;
      tripsByYearMonth[ym] += trips;
      if (!otdByYearMonth[ym]) otdByYearMonth[ym] = { weighted: 0, trips: 0 };
      otdByYearMonth[ym].weighted += otd * trips;
      otdByYearMonth[ym].trips += trips;
    }
  });

  const avgIdleTime = tripsCompletedTotal > 0 ? (idleWeighted / tripsCompletedTotal) : 0;
  const onTimeDeliveryRate = tripsCompletedTotal > 0 ? (otdWeighted / tripsCompletedTotal) * 100 : 0;

  const idleTable = Object.keys(idleByFirstName).map(name => ({
    name: name,
    avgIdleHours: idleByFirstName[name].trips > 0 ? (idleByFirstName[name].weighted / idleByFirstName[name].trips) : 0
  })).sort((a, b) => b.avgIdleHours - a.avgIdleHours);

  // ---- truckUtilizationMetrics: Miles per Truck, maintenance cost & miles by month (make filter) ----
  let sumTruckMiles = 0, sumMaintCost = 0;
  const distinctTrucks = new Set();
  const maintCostByYearMonth = {};
  const milesByYearMonth = {};
  const truckCountByYearMonth = {}; // for miles-per-truck-by-month distinct count

  d.truckUtil.forEach(row => {
    const truckId = pickField(row, ["truck_id"]);
    const make = truckMakeMap[truckId];
    if (!passFilter(make, filters.make)) return;

    const miles = num(row, ["total_miles"]);
    const cost = num(row, ["maintenance_cost"]);
    sumTruckMiles += miles;
    sumMaintCost += cost;
    if (truckId) distinctTrucks.add(truckId);

    const ym = yearMonthLabel(pickField(row, ["month", "YearMonth", "date"]));
    if (ym) {
      maintCostByYearMonth[ym] = (maintCostByYearMonth[ym] || 0) + cost;
      milesByYearMonth[ym] = (milesByYearMonth[ym] || 0) + miles;
      if (!truckCountByYearMonth[ym]) truckCountByYearMonth[ym] = new Set();
      if (truckId) truckCountByYearMonth[ym].add(truckId);
    }
  });

  const milesPerTruck = distinctTrucks.size > 0 ? (sumTruckMiles / distinctTrucks.size) : 0;

  // ---- deliveryEvents: Sum of detention_minutes by location_city (quarter + loadType via loads join) ----
  const detentionByCity = {};
  d.deliveryEvents.forEach(row => {
    const loadId = pickField(row, ["load_id"]);
    const q = loadQuarterMap[loadId];
    const lt = loadTypeMap[loadId];
    if (loadId && (!passFilter(q, filters.quarter) || !passFilter(lt, filters.loadType))) return;

    const city = pickField(row, ["location_city"]) || "Unknown";
    detentionByCity[city] = (detentionByCity[city] || 0) + num(row, ["detention_minutes"]);
  });

  // ---- Build month-aligned series for the two combo charts ----
  const monthKeysA = uniqueSorted(Object.keys(maintCostByYearMonth).concat(Object.keys(milesByYearMonth)))
    .sort((a, b) => yearMonthSortKey(a).localeCompare(yearMonthSortKey(b)));
  const maintCostSeries = monthKeysA.map(m => maintCostByYearMonth[m] || 0);
  const milesPerTruckSeries = monthKeysA.map(m => {
    const truckSet = truckCountByYearMonth[m];
    const cnt = truckSet ? truckSet.size : 0;
    return cnt > 0 ? (milesByYearMonth[m] / cnt) : 0;
  });

  const monthKeysB = uniqueSorted(Object.keys(tripsByYearMonth).concat(Object.keys(otdByYearMonth)))
    .sort((a, b) => yearMonthSortKey(a).localeCompare(yearMonthSortKey(b)));
  const tripsSeries = monthKeysB.map(m => tripsByYearMonth[m] || 0);
  const otdSeries = monthKeysB.map(m => {
    const o = otdByYearMonth[m];
    return o && o.trips > 0 ? (o.weighted / o.trips) * 100 : 0;
  });

  return {
    kpis: {
      onTimeDeliveryRate: onTimeDeliveryRate,
      avgIdleTime: avgIdleTime,
      avgLoadWeight: avgLoadWeight,
      milesPerTruck: milesPerTruck
    },
    charts: {
      loadCountByYearMonth: Object.keys(loadCountByYearMonth)
        .sort((a, b) => yearMonthSortKey(a).localeCompare(yearMonthSortKey(b)))
        .map(k => ({ label: k, value: loadCountByYearMonth[k] })),
      idleTable: idleTable,
      detentionByCity: sortMapDesc(detentionByCity, 8),
      monthLabelsA: monthKeysA,
      maintCostSeries: maintCostSeries,
      milesPerTruckSeries: milesPerTruckSeries,
      monthLabelsB: monthKeysB,
      tripsSeries: tripsSeries,
      otdSeries: otdSeries
    }
  };
}

// ==========================================================================
// ANALYTICAL DASHBOARD  (filters: yearMonth, make, maintenanceType, locationState)
// ==========================================================================

function getAnalyticalData(filters) {
  filters = filters || {};
  const d = loadAllData();

  const truckMakeMap = {};
  d.trucks.forEach(t => { truckMakeMap[pickField(t, ["truck_id"])] = pickField(t, ["make"]); });

  // ---- maintenanceRecords: Total Maintenance Cost, table by truck_id, by make, by state ----
  let totalMaintenanceCost = 0;
  const maintCostByTruck = {};
  const maintCostByMake = {};
  const maintCostByState = {};

  d.maintenance.forEach(row => {
    const truckId = pickField(row, ["truck_id"]);
    const make = truckMakeMap[truckId];
    const mType = pickField(row, ["maintenance_type"]);
    const state = pickField(row, ["location_state"]);
    const ym = yearMonthLabel(pickField(row, ["maintenance_date", "date"]));

    if (!passFilter(ym, filters.yearMonth)) return;
    if (!passFilter(make, filters.make)) return;
    if (!passFilter(mType, filters.maintenanceType)) return;
    if (!passFilter(state, filters.locationState)) return;

    const cost = num(row, ["total_cost"]);
    totalMaintenanceCost += cost;
    if (truckId) maintCostByTruck[truckId] = (maintCostByTruck[truckId] || 0) + cost;
    if (make) maintCostByMake[make] = (maintCostByMake[make] || 0) + cost;
    if (state) maintCostByState[state] = (maintCostByState[state] || 0) + cost;
  });

  // ---- fuelPurchases: Total Fuel Cost, by state, by state+make ----
  let totalFuelCost = 0;
  const fuelCostByState = {};
  const fuelCostByStateMake = {}; // { state: { make: value } }

  d.fuelPurchases.forEach(row => {
    const truckId = pickField(row, ["truck_id"]);
    const make = truckMakeMap[truckId];
    const state = pickField(row, ["location_state"]);
    const ym = yearMonthLabel(pickField(row, ["purchase_date", "date"]));

    if (!passFilter(ym, filters.yearMonth)) return;
    if (!passFilter(make, filters.make)) return;
    if (!passFilter(state, filters.locationState)) return;

    const cost = num(row, ["total_cost"]);
    totalFuelCost += cost;
    if (state) {
      fuelCostByState[state] = (fuelCostByState[state] || 0) + cost;
      if (!fuelCostByStateMake[state]) fuelCostByStateMake[state] = {};
      const mk = make || "Unknown";
      fuelCostByStateMake[state][mk] = (fuelCostByStateMake[state][mk] || 0) + cost;
    }
  });

  // ---- trips: CPM / Fuel Cost per Mile / MPG denominator (yearMonth + make filters) ----
  let totalTripMiles = 0, totalFuelGallons = 0;
  d.trips.forEach(row => {
    const truckId = pickField(row, ["truck_id"]);
    const make = truckMakeMap[truckId];
    const ym = yearMonthLabel(pickField(row, ["trip_date", "date"]));

    if (!passFilter(ym, filters.yearMonth)) return;
    if (!passFilter(make, filters.make)) return;

    totalTripMiles += num(row, ["actual_distance_miles"]);
    totalFuelGallons += num(row, ["fuel_gallons_used"]);
  });

  const cpm = totalTripMiles > 0 ? ((totalMaintenanceCost + totalFuelCost) / totalTripMiles) : 0;
  const fuelCpm = totalTripMiles > 0 ? (totalFuelCost / totalTripMiles) : 0;
  const mpg = totalFuelGallons > 0 ? (totalTripMiles / totalFuelGallons) : 0;

  // total_Cost by destination_state (fuel + maintenance combined, per state)
  const totalCostByState = {};
  Object.keys(maintCostByState).forEach(s => { totalCostByState[s] = (totalCostByState[s] || 0) + maintCostByState[s]; });
  Object.keys(fuelCostByState).forEach(s => { totalCostByState[s] = (totalCostByState[s] || 0) + fuelCostByState[s]; });

  // Total Fuel Cost by location_state and make -> stacked series
  const stateLabels = Object.keys(fuelCostByStateMake).sort();
  const allMakes = uniqueSorted(Object.values(fuelCostByStateMake).flatMap(m => Object.keys(m)));
  const stackedSeries = allMakes.map(mk => ({
    make: mk,
    data: stateLabels.map(s => (fuelCostByStateMake[s] && fuelCostByStateMake[s][mk]) || 0)
  }));

  return {
    kpis: {
      totalMaintenanceCost: totalMaintenanceCost,
      totalFuelCost: totalFuelCost,
      cpm: cpm,
      fuelCpm: fuelCpm,
      mpg: mpg
    },
    charts: {
      maintCostByTruckTable: sortMapDesc(maintCostByTruck),
      totalCostByState: sortMapDesc(totalCostByState),
      maintCostByMake: sortMapDesc(maintCostByMake),
      stateLabels: stateLabels,
      stackedFuelByStateMake: stackedSeries
    }
  };
}
