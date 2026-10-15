const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { test } = require("node:test");
const XLSX = require("../PDIExport/libs/xlsx.full.min.js");

const background = vm.createContext({
  XLSX,
  importScripts() {},
  chrome: { runtime: { onMessage: { addListener() {} } } }
});
vm.runInContext(fs.readFileSync(path.join(__dirname, "../PDIExport/background.js"), "utf8"), background);

// Keep PDI's scheduleEventId/scheduleEventID casing and timestamps. These
// examples reproduce the reported AM delivery after noon and PM after midnight.
function makeTrip(id, scheduleEventId, start, deliveryStart, label, flags = {}) {
  return {
    id,
    scheduleEventId,
    unitCombinationId: 83861,
    start,
    end: deliveryStart,
    events: [{
      id: id * 10,
      type: 9,
      sequenceNumber: 1,
      label,
      start: deliveryStart,
      end: deliveryStart,
      customerId: id,
      referenceNumber: String(id)
    }],
    orders: [{ orderId: id, customerId: id }],
    ...flags
  };
}

function fixture() {
  return [{
    unitCombination: { id: 83861, label: "83861" },
    drivers: [
      { scheduleEventID: 274913, shiftNumber: 2, start: "2026-10-14T16:00:00", end: "2026-10-15T03:59:00" },
      { scheduleEventID: 247839, shiftNumber: 1, start: "2026-10-15T04:00:00", end: "2026-10-15T15:59:00" },
      { scheduleEventID: 274914, shiftNumber: 2, start: "2026-10-15T16:00:00", end: "2026-10-16T03:59:00" },
      { scheduleEventID: 247840, shiftNumber: 1, start: "2026-10-16T04:00:00", end: "2026-10-16T15:59:00" }
    ],
    trips: [
      makeTrip(179851, 274913, "2026-10-14T23:00:00", "2026-10-15T01:00:00", "Previous PM", { isLastTrip: true }),
      makeTrip(179904, 247839, "2026-10-15T04:00:00", "2026-10-15T07:53:00", "BP Keith - Diesel", { isFirstTrip: true }),
      makeTrip(179985, 247839, "2026-10-15T11:43:00", "2026-10-15T12:54:00", "X-Con Briens Road", { isLastTrip: true }),
      makeTrip(179977, 274914, "2026-10-15T16:00:00", "2026-10-15T17:36:00", "X-Con Unley", { isFirstTrip: true }),
      makeTrip(179993, 274914, "2026-10-15T19:14:00", "2026-10-15T21:08:00", "X-Con Seaford"),
      makeTrip(179994, 274914, "2026-10-15T23:09:00", "2026-10-16T00:29:00", "X-Con Anzac Highway", { isLastTrip: true }),
      makeTrip(180000, 247840, "2026-10-16T04:00:00", "2026-10-16T06:00:00", "Next AM", { isFirstTrip: true, isLastTrip: true })
    ]
  }];
}

function detailsFor(groups) {
  const details = new Map();
  for (const trip of groups.flatMap(g => g.trips)) {
    details.set(String(trip.id), {
      tripId: trip.id,
      orders: [{
        id: trip.id,
        referenceNumber: String(trip.id),
        customerId: trip.id,
        name: trip.events[0].label,
        orderPositions: [{
          productId: 1,
          articleName: "Diesel",
          quantity: 1000,
          compartments: [{ unitCompartmentId: 1, compartmentNumber: 1, quantity: 1000 }]
        }]
      }],
      tripLoad: [{
        trailer: { compartments: [{ id: 1, compartmentNumber: 1 }] },
        compartments: [{ unitCompartmentId: 1, productId: 1, quantity: 1000 }]
      }]
    });
  }
  return details;
}

function exportRows(groups, date = "2026-10-15", eventDetails = new Map()) {
  const { workbook } = background.buildWorkbookSheets(groups, detailsFor(groups), eventDetails, date);
  return Object.fromEntries(workbook.SheetNames.map(name => [
    name, XLSX.utils.sheet_to_json(workbook.Sheets[name])
  ]));
}

test("Briens Road stays AM after noon and is the second AM load", () => {
  const rows = exportRows(fixture());
  const briens = rows.Deliveries.find(r => r.tripId === 179985);
  assert.equal(briens.amPm, "AM");
  assert.equal(briens.deliveryDate, "2026-10-15");
  assert.equal(briens.deliveryOrder, 2);
  assert.equal(rows.Trips.find(r => r.tripId === 179985).scheduleDate, "2026-10-15");
  assert.equal(rows.Trips.find(r => r.tripId === 179985).amPm, "AM");
});

test("after-midnight deliveries remain on the owning PM shift and load sequence", () => {
  const rows = exportRows(fixture());
  const anzac = rows.Deliveries.find(r => r.tripId === 179994);
  assert.equal(anzac.deliveryDate, "2026-10-15");
  assert.equal(anzac.amPm, "PM");
  assert.equal(anzac.deliveryOrder, 3);
  assert.equal(rows["Trip Events"].find(r => r.tripId === 179994).start, "2026-10-16T00:29:00");
  assert.ok(rows["Order Position Compartments"].some(r => r.tripId === 179994));
});

test("a PM trip starting after midnight belongs to the previous schedule day", () => {
  const groups = fixture();
  groups[0].trips.find(t => t.id === 179994).isLastTrip = false;
  groups[0].trips.push(makeTrip(180001, 274914, "2026-10-16T01:00:00", "2026-10-16T02:00:00", "Late PM", { isLastTrip: true }));
  const rows = exportRows(groups);
  const late = rows.Deliveries.find(r => r.tripId === 180001);
  assert.equal(late.deliveryDate, "2026-10-15");
  assert.equal(late.amPm, "PM");
  assert.equal(late.deliveryOrder, 4);
});

test("all trip sheets contain the same selected shifts and exclude adjacent days", () => {
  const rows = exportRows(fixture());
  const expected = [179904, 179977, 179985, 179993, 179994];
  for (const name of ["Deliveries", "Trips", "Trip Events", "Trip Orders", "Trip Compartments", "Order Position Compartments"]) {
    assert.deepEqual([...new Set(rows[name].map(r => r.tripId))].sort((a, b) => a - b), expected, name);
  }
  assert.deepEqual(rows.Drivers.map(d => d.scheduleEventId), [247839, 274914]);
});

test("the following day does not re-export the previous PM tail", () => {
  const rows = exportRows(fixture(), "2026-10-16");
  assert.deepEqual(rows.Deliveries.map(r => r.tripId), [180000]);
  assert.equal(rows.Deliveries[0].amPm, "AM");
  assert.equal(rows.Deliveries[0].deliveryOrder, 1);
});

test("actual and planned delivery dates cannot change the assigned shift", () => {
  const eventDetails = new Map([[
    "179985:1799850",
    { actualStart: "2026-10-16T00:10:00", actualEnd: "2026-10-16T00:30:00", plannedStart: "2026-10-15T12:54:00" }
  ]]);
  const row = exportRows(fixture(), "2026-10-15", eventDetails).Deliveries.find(r => r.tripId === 179985);
  assert.equal(row.deliveryDate, "2026-10-15");
  assert.equal(row.amPm, "AM");
});

test("load numbering resets for each schedule date and shift", () => {
  const rows = exportRows(fixture(), "").Deliveries;
  assert.equal(rows.find(r => r.tripId === 179904).deliveryOrder, 1);
  assert.equal(rows.find(r => r.tripId === 179977).deliveryOrder, 1);
  assert.equal(rows.find(r => r.tripId === 180000).deliveryOrder, 1);
});

test("distinct shift assignments on the same vehicle and day have independent numbering", () => {
  const groups = fixture();
  groups[0].drivers.push({ scheduleEventID: 900000, shiftNumber: 1, start: "2026-10-15T10:00:00" });
  groups[0].trips.push(makeTrip(180002, 900000, "2026-10-15T10:00:00", "2026-10-15T11:00:00", "Other AM", { isFirstTrip: true, isLastTrip: true }));
  const rows = exportRows(groups).Deliveries;
  assert.equal(rows.find(r => r.tripId === 179985).deliveryOrder, 2);
  assert.equal(rows.find(r => r.tripId === 180002).deliveryOrder, 1);
});

test("shift IDs match across number/string values without requiring a driver ID", () => {
  const groups = fixture();
  groups[0].drivers[1].scheduleEventID = "247839";
  const row = exportRows(groups).Deliveries.find(r => r.tripId === 179985);
  assert.equal(row.amPm, "AM");
});

test("missing shift metadata falls back to trip start rather than delivery time", () => {
  const groups = fixture();
  groups[0].drivers = [];
  const rows = exportRows(groups).Deliveries;
  assert.equal(rows.find(r => r.tripId === 179985).amPm, "AM");
  assert.equal(rows.find(r => r.tripId === 179994).amPm, "PM");
  assert.equal(rows.find(r => r.tripId === 179994).deliveryDate, "2026-10-15");
  assert.ok(!rows.some(r => r.tripId === 179851));
});

test("filtering preserves its input and is safe to repeat before workbook creation", () => {
  const groups = fixture();
  const original = JSON.stringify(groups);
  const once = background.filterTripGroupsBySelectedDate(groups, "2026-10-15");
  const twice = background.filterTripGroupsBySelectedDate(once, "2026-10-15");
  assert.equal(JSON.stringify(groups), original);
  assert.equal(JSON.stringify(once), JSON.stringify(twice));
});

test("empty schedules create a valid workbook", () => {
  const rows = exportRows([]);
  assert.equal(Object.keys(rows).length, 8);
  assert.ok(Object.values(rows).every(sheet => sheet.length === 0));
});
