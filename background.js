// background.js
importScripts("libs/xlsx.full.min.js");

const API_ORIGIN = "https://svc-scheduling.logistics.pdisoftware.com";
const UI_ORIGINS = ["https://plc-mt.logistics.pdisoftware.com"];
const AUTH_TTL_MS = 10 * 60 * 1000;

function isAllowedSender(sender) {
  try {
    const senderUrl = sender?.url || sender?.tab?.url || "";
    return UI_ORIGINS.some(origin => senderUrl.startsWith(origin));
  } catch (_) {
    return false;
  }
}

function isAuthFresh(pdiAuth, ttlMs = AUTH_TTL_MS) {
  const capturedAt = pdiAuth?.capturedAt ? Date.parse(pdiAuth.capturedAt) : NaN;
  if (!Number.isFinite(capturedAt)) return false;
  return Date.now() - capturedAt <= ttlMs;
}

async function waitForAuthCapture({ timeoutMs = 15000, pollMs = 500 } = {}) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const auth = await chrome.storage.local.get(["pdiAuth"]);
    const pdiAuth = auth?.pdiAuth;
    if (pdiAuth?.authorization && pdiAuth?.tenantId && isAuthFresh(pdiAuth)) return pdiAuth;
    await new Promise(r => setTimeout(r, pollMs));
  }
  return null;
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  (async () => {
    try {
      if (msg?.type === "AUTH_CAPTURED") {
        if (!isAllowedSender(_sender)) {
          throw new Error("Auth capture from unauthorized origin");
        }
        const { authorization, tenantId } = msg.payload || {};
        if (authorization || tenantId) {
          const existing = await chrome.storage.local.get(["pdiAuth"]);
          const pdiAuth = {
            authorization: authorization ?? existing?.pdiAuth?.authorization ?? null,
            tenantId: tenantId ?? existing?.pdiAuth?.tenantId ?? null,
            capturedAt: new Date().toISOString()
          };
          await chrome.storage.local.set({ pdiAuth });
        }
        sendResponse({ ok: true });
        return;
      }

      if (msg?.type === "EXPORT_TRIPS") {
        const { profileId, start, end, language } = msg.payload;
        try {
          let pdiAuth = await waitForAuthCapture({ timeoutMs: 15000, pollMs: 500 });
          if (!pdiAuth) {
            const auth = await chrome.storage.local.get(["pdiAuth"]);
            pdiAuth = auth?.pdiAuth;
          }

          if (pdiAuth && !isAuthFresh(pdiAuth)) {
            await chrome.storage.local.remove(["pdiAuth"]);
            pdiAuth = null;
          }

          if (!pdiAuth?.authorization) {
            throw new Error("No Authorization token captured yet. While the export is waiting, trigger any action in the PDI app (e.g. search or navigation) to create an API call.");
          }
          if (!pdiAuth?.tenantId) {
            throw new Error("No tenant-id captured yet. While the export is waiting, trigger any action in the PDI app (e.g. search or navigation) to create an API call.");
          }

          // 1) Base list
          const tripsPayload = await apiFetchTrips({ profileId, start, end, language, pdiAuth });
          const tripGroups = Array.isArray(tripsPayload) ? tripsPayload : [tripsPayload];
          const baseTrips = tripGroups.flatMap(g => (Array.isArray(g?.trips) ? g.trips : []));

          // 2) Drilldown per tripId -> /api/Trips/{id}
          // Concurrency limit avoids smashing the API.
          const drilldownByTripId = new Map();

          await parallelLimit(
            baseTrips,
            6,
            async (trip) => {
              const tripId = trip?.id ?? trip?.tripId;
              if (tripId == null) return;
              const detail = await apiFetchTripDetail({ tripId, pdiAuth });
              drilldownByTripId.set(String(tripId), detail);
            }
          );

          // 2b) Delivery event details for loadIDs
          const deliveryEventDetails = new Map();
          await parallelLimit(
            baseTrips,
            4,
            async (trip) => {
              const tripId = trip?.id ?? trip?.tripId;
              if (tripId == null) return;
              const events = Array.isArray(trip?.events) ? trip.events : [];
              const deliveries = events.filter(e => e?.type === 9 && e?.id != null);
              await parallelLimit(
                deliveries,
                2,
                async (ev) => {
                  const eventId = ev?.id;
                  if (eventId == null) return;
                  const detail = await apiFetchTripEventDetail({ tripId, eventId, pdiAuth });
                  deliveryEventDetails.set(`${tripId}:${eventId}`, detail);
                }
              );
            }
          );

          // 3) Build workbook sheets (base + drilldown sheets)
          const sheets = buildWorkbookSheets(tripGroups, drilldownByTripId, deliveryEventDetails);

          const filename = `pdi_trips_${start}_to_${end}.xlsx`;
          await downloadWorkbook(sheets.workbook, filename);

          sendResponse({
            ok: true,
            message: [
              `Exported ${sheets.summary.tripCount} trips`,
              `Trip Compartments: ${sheets.summary.tripCompartmentRows}`,
              `Order Position Compartments: ${sheets.summary.orderPositionCompartmentRows}`
            ].join("\n")
          });
          return;
        } finally {
          await chrome.storage.local.remove(["pdiAuth"]);
        }
      }

      sendResponse({ ok: false, error: "Unknown message type" });
    } catch (err) {
      sendResponse({ ok: false, error: err?.message || String(err) });
    }
  })();

  return true;
});

async function apiFetchTrips({ profileId, start, end, language, pdiAuth }) {
  const url = new URL(`${API_ORIGIN}/api/Trips`);
  url.searchParams.set("profileId", profileId);
  url.searchParams.set("start", start);
  url.searchParams.set("end", end);
  url.searchParams.set("language", language || "en");

  const res = await fetch(url.toString(), {
    method: "GET",
    headers: {
      "accept": "application/json, text/plain, */*",
      "authorization": pdiAuth.authorization,
      "tenant-id": pdiAuth.tenantId
    },
    credentials: "include"
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Trips request failed: ${res.status} ${res.statusText}\n${text.slice(0, 300)}`);
  }

  return res.json();
}

async function apiFetchTripDetail({ tripId, pdiAuth }) {
  const res = await fetch(`${API_ORIGIN}/api/Trips/${encodeURIComponent(tripId)}`, {
    method: "GET",
    headers: {
      "accept": "application/json, text/plain, */*",
      "authorization": pdiAuth.authorization,
      "tenant-id": pdiAuth.tenantId
    },
    credentials: "include"
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Trip detail failed for ${tripId}: ${res.status} ${res.statusText}\n${text.slice(0, 200)}`);
  }

  return res.json();
}

async function apiFetchTripEventDetail({ tripId, eventId, pdiAuth }) {
  const res = await fetch(`${API_ORIGIN}/api/Trips/${encodeURIComponent(tripId)}/tripEvents/${encodeURIComponent(eventId)}/details`, {
    method: "GET",
    headers: {
      "accept": "application/json, text/plain, */*",
      "authorization": pdiAuth.authorization,
      "tenant-id": pdiAuth.tenantId
    },
    credentials: "include"
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Trip event detail failed for ${tripId}/${eventId}: ${res.status} ${res.statusText}\n${text.slice(0, 200)}`);
  }

  return res.json();
}

// Generic concurrency limiter
async function parallelLimit(items, concurrency, worker) {
  const list = Array.isArray(items) ? items : [];
  let idx = 0;

  const runners = Array.from({ length: Math.min(concurrency, list.length) }, async () => {
    while (true) {
      const i = idx++;
      if (i >= list.length) break;
      await worker(list[i], i);
    }
  });

  await Promise.all(runners);
}

function buildWorkbookSheets(groups, drilldownByTripId, deliveryEventDetails) {
  const payloads = Array.isArray(groups) ? groups : [groups];

  // Base sheets
  const unitRows = [];
  const driverRows = [];
  const tripRows = [];
  const eventRows = [];
  const orderRows = [];
  const allTrips = [];

  for (const payload of payloads) {
    const unit = payload?.unitCombination || {};
    const drivers = Array.isArray(payload?.drivers) ? payload.drivers : [];
    const trips = Array.isArray(payload?.trips) ? payload.trips : [];

    if (Object.keys(unit).length > 0) unitRows.push(flattenUnitCombination(unit));
    for (const d of drivers) driverRows.push(flattenDriver(d));
    for (const t of trips) {
      tripRows.push(flattenTrip(t, unit));
      allTrips.push(t);

      const evs = Array.isArray(t?.events) ? t.events : [];
      for (const ev of evs) eventRows.push(flattenTripEvent(t, ev));

      const ords = Array.isArray(t?.orders) ? t.orders : [];
      for (const o of ords) orderRows.push(flattenTripOrder(t, o));
    }
  }

  const deliveryRows = [];

  const tripOrderById = computeTripOrderById(allTrips);

  // Drilldown sheets
  const tripCompartmentRows = [];
  const orderPositionCompartmentRows = [];

  for (const t of allTrips) {
    const tripKey = String(t?.id ?? t?.tripId ?? "");
    const detail = tripKey ? drilldownByTripId.get(tripKey) : undefined;
    if (!detail) continue;

    const deliveryEvents = Array.isArray(t?.events) ? t.events.filter(e => e?.type === 9) : [];
    const deliveryByRef = new Map();
    const deliveryByCustomer = new Map();
    const deliveryDetailByEventId = new Map();
    for (const ev of deliveryEvents) {
      if (ev?.referenceNumber) deliveryByRef.set(String(ev.referenceNumber), ev);
      if (ev?.customerId != null) deliveryByCustomer.set(String(ev.customerId), ev);
      const eventId = ev?.id;
      if (eventId != null && deliveryEventDetails) {
        const key = `${t?.id ?? t?.tripId}:${eventId}`;
        const evDetail = deliveryEventDetails.get(key);
        if (evDetail) deliveryDetailByEventId.set(String(eventId), evDetail);
      }
    }

    const productsById = new Map();
    for (const a of (Array.isArray(detail?.articles) ? detail.articles : [])) {
      const key = a?.productId ?? null;
      if (key != null) productsById.set(String(key), a?.productName || a?.name || "");
    }

    const tripLoads = Array.isArray(detail?.tripLoad) ? detail.tripLoad : [];
    const tripLoadByPoint = new Map();
    for (const tl of tripLoads) {
      if (tl?.loadingPointId != null) tripLoadByPoint.set(String(tl.loadingPointId), tl);
    }

    const unitItems = Array.isArray(detail?.unit?.units) ? detail.unit.units : [];
    const tractor = unitItems.find(u => String(u?.unitType || "").toLowerCase().includes("tractor")) || {};
    
    // Collect ALL trailers (there can be multiple: A-Trailer, Semi-trailer, etc.)
    const trailerUnits = unitItems.filter(u => String(u?.unitType || "").toLowerCase().includes("trailer"));
    const trailerUnit = trailerUnits[0] || {};
    
    // Build a map of trailers by ID for compartment-level lookup
    const trailerById = new Map();
    for (const tu of trailerUnits) {
      if (tu?.id != null) trailerById.set(String(tu.id), tu);
    }
    
    // Build maps for trailer lookup by compartment or unit
    const trailerByNumber = new Map();
    for (const tu of trailerUnits) {
      if (tu?.number != null) trailerByNumber.set(String(tu.number), tu);
    }

    // Map unitCompartmentId -> trailer number (from physical compartments)
    const compartmentIdToTrailerNumber = new Map();
    for (const tl of tripLoads) {
      const trailer = tl?.trailer || {};
      for (const pc of (Array.isArray(trailer?.compartments) ? trailer.compartments : [])) {
        if (pc?.id != null && pc?.unitNumber != null) {
          compartmentIdToTrailerNumber.set(String(pc.id), String(pc.unitNumber));
        }
      }
    }

    const orders = Array.isArray(detail?.orders) ? detail.orders : [];
    for (const o of orders) {
      const ev =
        deliveryByRef.get(String(o?.referenceNumber ?? "")) ||
        deliveryByRef.get(String(o?.id ?? "")) ||
        (o?.customerId != null ? deliveryByCustomer.get(String(o.customerId)) : undefined);

      const evDetail = ev?.id != null ? deliveryDetailByEventId.get(String(ev.id)) : undefined;
      const tripStart = t?.start || "";
      const tripEnd = t?.end || "";
      const tripEventStart = Array.isArray(t?.events) ? t.events?.[0]?.start : "";
      const tripEventEnd = Array.isArray(t?.events) ? t.events?.[0]?.end : "";
      const eventStart =
        tripEventStart ||
        evDetail?.actualStart ||
        evDetail?.plannedStart ||
        ev?.start ||
        detail?.start ||
        tripStart ||
        "";
      const eventEnd =
        tripEventEnd ||
        evDetail?.actualEnd ||
        evDetail?.plannedEnd ||
        ev?.end ||
        detail?.end ||
        tripEnd ||
        "";
      const deliveryDate = formatDateOnly(tripStart || eventStart);
      const amPm = formatAmPm(tripStart || eventStart);
      const deliveryOrder = tripOrderById.get(String(t?.id ?? t?.tripId ?? "")) ?? "";

      const positions = Array.isArray(o?.orderPositions) ? o.orderPositions : [];
      for (const p of positions) {
        const comps = Array.isArray(p?.compartments) ? p.compartments : [];

        const loadPointId = p?.loadingPointID ?? o?.loadingPointID ?? o?.loadingPointId ?? "";
        const tripLoad = (loadPointId != null && loadPointId !== "")
          ? tripLoadByPoint.get(String(loadPointId))
          : tripLoads[0];

        const loadingTerminal = tripLoad
          ? `${tripLoad?.name ?? ""}${tripLoad?.loadingPointId != null ? ` (${tripLoad.loadingPointId})` : ""}`.trim()
          : "";

        const loadId = evDetail?.loadIDs ?? "";

        for (const c of comps) {
          const productName =
            p?.articleName ||
            productsById.get(String(p?.productId ?? "")) ||
            "";

          // Determine the specific trailer for this compartment
          let compTrailer = null;
          if (c?.unitId != null) {
            compTrailer = trailerById.get(String(c.unitId)) || null;
          }

          if (!compTrailer && c?.unitCompartmentId != null) {
            const trailerNumber = compartmentIdToTrailerNumber.get(String(c.unitCompartmentId));
            if (trailerNumber) compTrailer = trailerByNumber.get(String(trailerNumber)) || null;
          }

          const trailerId =
            compTrailer?.number ??
            compTrailer?.id ??
            tripLoad?.trailer?.number ??
            tripLoad?.trailer?.id ??
            trailerUnit?.number ??
            trailerUnit?.id ??
            "";

          deliveryRows.push({
            tripId: t?.id ?? t?.tripId ?? "",
            deliveryDate,
            customerName: o?.name ?? "",
            amPm,
            loadingTerminal,
            product: productName,
            quantity: c?.quantity ?? p?.quantity ?? "",
            shipmentId: evDetail?.orderReference ?? o?.referenceNumber ?? o?.id ?? "",
            totalTimeForLoadIncludingDeliveryMin: t?.tripLength ?? calcDurationMinutes(eventStart, eventEnd) ?? "",
            loadId,
            deliveryOrder,
            dropSequence: ev?.sequenceNumber ?? "",
            customerNumber: o?.customerNumber ?? "",
            customerAddress: formatAddress(evDetail) || o?.address || o?.street || o?.zip || "",
            truckId: tractor?.number ?? tractor?.id ?? "",
            trailerId,
            compartmentNumber: c?.compartmentNumber ?? "",
            preloadForNextShift: t?.preloadForNextShift ?? false,
            preloadInPreviousShift: t?.preloadInPreviousShift ?? false
          });
        }
      }
    }

    // TripLoad compartments (actual load plan compartments)
    for (const tl of tripLoads) {
      const trailer = tl?.trailer || {};

      // Map physical compartment info by unitCompartmentId
      const physById = new Map();
      for (const pc of (Array.isArray(trailer?.compartments) ? trailer.compartments : [])) {
        physById.set(pc?.id, pc);
      }

      for (const c of (Array.isArray(tl?.compartments) ? tl.compartments : [])) {
        const phys = physById.get(c?.unitCompartmentId) || {};
        const productId = c?.productId ?? "";
        const productName = productsById.get(productId) || "";

        // Determine the correct trailer for this compartment
        // Use unitId from compartment, or unitNumber from physical compartment to find the right trailer
        const compartmentUnitId = c?.unitId ?? "";
        const compartmentUnitNumber = phys?.unitNumber ?? "";
        
        // Try to find the specific trailer for this compartment
        let compTrailer = compartmentUnitId ? trailerById.get(String(compartmentUnitId)) : null;
        if (!compTrailer && compartmentUnitNumber) {
          // Find trailer by number if unitId lookup failed
          compTrailer = trailerUnits.find(tu => String(tu?.number) === String(compartmentUnitNumber));
        }
        // Fall back to the tripLoad trailer or first trailer unit
        compTrailer = compTrailer || trailer || trailerUnit || {};

        const trailerNumber = compTrailer?.number ?? trailer?.number ?? "";
        const trailerPlate = compTrailer?.licencePlate ?? trailer?.licencePlate ?? "";
        const trailerId = compTrailer?.id ?? trailer?.id ?? "";

        tripCompartmentRows.push({
          tripId: detail?.tripId ?? t?.id ?? "",
          tripReferenceNumber: detail?.referenceNumber ?? t?.reference ?? "",
          loadingPointId: tl?.loadingPointId ?? "",
          loadingPointName: tl?.name ?? "",

          trailerId,
          trailerNumber,
          trailerPlate,

          unitCompartmentId: c?.unitCompartmentId ?? "",
          compartmentNumber: phys?.compartmentNumber ?? "",
          compartmentCapacityL: phys?.capacity ?? "",
          compartmentMinLoadL: phys?.minLoadCapacity ?? "",

          productId,
          productName,
          quantityL: c?.quantity ?? ""
        });
      }
    }

    // Orders -> orderPositions -> compartments mapping (delivery allocation)
    for (const o of orders) {
      const positions = Array.isArray(o?.orderPositions) ? o.orderPositions : [];
      for (const p of positions) {
        const comps = Array.isArray(p?.compartments) ? p.compartments : [];
        for (const oc of comps) {
          orderPositionCompartmentRows.push({
            tripId: detail?.tripId ?? t?.id ?? "",
            tripReferenceNumber: detail?.referenceNumber ?? t?.reference ?? "",

            orderId: o?.id ?? "",
            orderReferenceNumber: o?.referenceNumber ?? "",
            customerId: o?.customerId ?? "",
            customerNumber: o?.customerNumber ?? "",
            customerName: o?.name ?? "",

            orderPositionId: p?.id ?? "",
            articleId: p?.articleId ?? "",
            articleName: p?.articleName ?? "",
            productId: p?.productId ?? "",
            orderedQuantityL: p?.quantity ?? "",

            unitCompartmentId: oc?.unitCompartmentId ?? "",
            compartmentNumber: oc?.compartmentNumber ?? "",
            compartmentQuantityL: oc?.quantity ?? ""
          });
        }
      }
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(deliveryRows), "Deliveries");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(tripRows), "Trips");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(eventRows), "Trip Events");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(orderRows), "Trip Orders");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(tripCompartmentRows), "Trip Compartments");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(orderPositionCompartmentRows), "Order Position Compartments");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(driverRows), "Drivers");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(unitRows), "Unit Combination");

  return {
    workbook: wb,
    summary: {
      tripCount: tripRows.length,
      tripCompartmentRows: tripCompartmentRows.length,
      orderPositionCompartmentRows: orderPositionCompartmentRows.length,
      deliveryRows: deliveryRows.length
    }
  };
}

function formatDateOnly(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function shiftNumberToAmPm(value) {
  const n = Number(value);
  if (n === 1) return "AM";
  if (n === 2) return "PM";
  return value ?? "";
}

function formatAmPm(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  return d.getHours() >= 12 ? "PM" : "AM";
}

function calcDurationMinutes(startIso, endIso) {
  if (!startIso || !endIso) return null;
  const s = new Date(startIso);
  const e = new Date(endIso);
  if (Number.isNaN(s.getTime()) || Number.isNaN(e.getTime())) return null;
  const minutes = Math.round((e - s) / 60000);
  return minutes >= 0 ? minutes : null;
}

function formatAddress(evDetail) {
  if (!evDetail) return "";
  const parts = [];
  const street = [evDetail?.streetNumber, evDetail?.street].filter(Boolean).join(" ").trim();
  if (street) parts.push(street);
  if (evDetail?.city) parts.push(evDetail.city);
  if (evDetail?.zip) parts.push(evDetail.zip);
  return parts.join(", ");
}

function computeTripOrderById(trips) {
  const orderById = new Map();
  const groups = new Map();

  trips.forEach((t, idx) => {
    const unitId = String(t?.unitCombinationId ?? "");
    const amPm = tripAmPm(t);
    const groupKey = `${unitId}:${amPm}`;
    if (!groups.has(groupKey)) groups.set(groupKey, []);
    groups.get(groupKey).push({ trip: t, idx });
  });

  for (const list of groups.values()) {
    const total = list.length;
    if (total === 0) continue;

    const firstEntry = earliestFlagged(list, "isFirstTrip");
    const lastEntry = latestFlagged(list, "isLastTrip");

    const middle = list
      .filter(x => x !== firstEntry && x !== lastEntry)
      .sort((a, b) => {
        const aTime = firstDeliveryTime(a.trip) ?? tripStartTime(a.trip);
        const bTime = firstDeliveryTime(b.trip) ?? tripStartTime(b.trip);
        if (aTime !== bTime) return aTime - bTime;
        return a.idx - b.idx; // response order fallback
      });

    const ordered = [firstEntry, ...middle, lastEntry].filter(Boolean);
    ordered.forEach((entry, i) => {
      const t = entry.trip;
      const tripKey = String(t?.id ?? t?.tripId ?? "");
      if (!tripKey) return;
      orderById.set(tripKey, i + 1);
    });
  }

  return orderById;
}

function firstDeliverySequence(trip) {
  const events = Array.isArray(trip?.events) ? trip.events : [];
  const deliveries = events.filter(e => e?.type === 9 && Number.isFinite(e?.sequenceNumber));
  if (deliveries.length === 0) return Number.POSITIVE_INFINITY;
  return Math.min(...deliveries.map(e => e.sequenceNumber));
}

function firstDeliveryTime(trip) {
  const events = Array.isArray(trip?.events) ? trip.events : [];
  const deliveries = events.filter(e => e?.type === 9 && e?.start);
  if (deliveries.length === 0) return null;
  return Math.min(...deliveries.map(e => new Date(e.start).getTime()).filter(n => !Number.isNaN(n)));
}

function tripStartTime(trip) {
  const t = new Date(trip?.start || "").getTime();
  return Number.isNaN(t) ? Number.POSITIVE_INFINITY : t;
}

function tripAmPm(trip) {
  const time = firstDeliveryTime(trip) ?? tripStartTime(trip);
  if (!Number.isFinite(time)) return "";
  const d = new Date(time);
  return d.getHours() >= 12 ? "PM" : "AM";
}

function earliestFlagged(list, flag) {
  const flagged = list.filter(x => x.trip?.[flag] === true);
  if (flagged.length === 0) return null;
  return flagged.sort((a, b) => (tripStartTime(a.trip) - tripStartTime(b.trip)) || (a.idx - b.idx))[0];
}

function latestFlagged(list, flag) {
  const flagged = list.filter(x => x.trip?.[flag] === true);
  if (flagged.length === 0) return null;
  return flagged.sort((a, b) => (tripStartTime(b.trip) - tripStartTime(a.trip)) || (b.idx - a.idx))[0];
}

function flattenUnitCombination(u) {
  const units = Array.isArray(u?.units) ? u.units : [];
  const prime = units.find(x => x.order === 1) || {};
  const rear = units.find(x => x.order === 2) || {};

  return {
    unitCombinationId: u?.id ?? "",
    label: u?.label ?? "",
    transportCompany: u?.transportCompany ?? "",
    vehicleBase: u?.vehicleBase ?? "",
    capacityL: u?.capacity ?? "",
    payload: u?.payload ?? "",
    gpsNumber: u?.gpsNumber ?? "",

    primeUnitId: prime?.id ?? "",
    primeNumber: prime?.number ?? u?.vehicleNumberPrimeMover ?? "",
    primePlate: prime?.licencePlate ?? u?.vehicleLicensePlatePrimeMover ?? "",

    rearUnitId: rear?.id ?? "",
    rearNumber: rear?.number ?? u?.vehicleNumberRearUnit ?? "",
    rearPlate: rear?.licencePlate ?? u?.vehicleLicensePlateRearUnit ?? "",

    driverNameLoggedIn: u?.ovcDriverInformation?.driverName ?? "",
    loggedInDateTime: u?.ovcDriverInformation?.loggedInDateTime ?? ""
  };
}

function flattenDriver(d) {
  const startDate = formatDateOnly(d?.start) || d?.start || "";
  const endDate = formatDateOnly(d?.end) || d?.end || "";
  return {
    driverId: d?.id ?? "",
    name: d?.name ?? "",
    fullName: d?.fullName ?? "",
    scheduleEventId: d?.scheduleEventID ?? "",
    shiftNumber: shiftNumberToAmPm(d?.shiftNumber),
    start: startDate,
    end: endDate,
    shiftLockedByDA: d?.shiftLockedByDA ?? false
  };
}

function flattenTrip(t, unit) {
  return {
    tripId: t?.id ?? "",
    reference: t?.reference ?? "",
    dispoStatus: t?.dispoStatus ?? "",
    dispoStatusName: t?.dispoStatusName ?? "",
    locked: t?.locked ?? false,

    driverId: t?.driverId ?? "",
    driverName: t?.driverName ?? "",

    start: t?.start ?? "",
    end: t?.end ?? "",

    plannedVolumeL: t?.plannedVolume ?? "",
    maxVolumeL: t?.maxVolume ?? "",
    plannedPayload: t?.plannedPayload ?? "",
    maxPayload: t?.maxPayload ?? "",

    tripLengthMin: t?.tripLength ?? "",
    scheduleEventId: t?.scheduleEventId ?? "",
    unitCombinationId: t?.unitCombinationId ?? "",

    vehicleBase: unit?.vehicleBase ?? "",
    transportCompany: unit?.transportCompany ?? ""
  };
}

function flattenTripEvent(t, ev) {
  const type = ev?.type ?? "";
  const typeGuess =
    type === 3 ? "Depot" :
    type === 8 ? "Loading" :
    type === 9 ? "Delivery" :
    "";

  return {
    tripId: t?.id ?? "",
    tripReference: ev?.tripReference ?? t?.reference ?? "",
    sequenceNumber: ev?.sequenceNumber ?? "",
    eventId: ev?.id ?? "",
    eventType: type,
    eventTypeGuess: typeGuess,
    label: ev?.label ?? "",
    start: ev?.start ?? "",
    end: ev?.end ?? "",
    distanceKm: ev?.distance ?? "",
    customerId: ev?.customerId ?? "",
    referenceNumber: ev?.referenceNumber ?? "",
    loadingPointId: ev?.loadingPointId ?? ""
  };
}

function flattenTripOrder(t, o) {
  return {
    tripId: t?.id ?? "",
    tripReference: t?.reference ?? "",
    orderId: o?.orderId ?? "",
    customerId: o?.customerId ?? "",
    dispoStatus: o?.dispoStatus ?? "",
    dispoStatusName: o?.dispoStatusName ?? "",
    creditStatus: o?.creditStatus ?? "",
    creditStatusName: o?.creditStatusName ?? "",
    vmi: o?.vmi ?? false,
    unAttendedDelivery: o?.unAttendedDelivery ?? false
  };
}

async function downloadWorkbook(workbook, filename) {
  // Service workers don't support URL.createObjectURL; use a data URL instead.
  const base64 = XLSX.write(workbook, { bookType: "xlsx", type: "base64" });
  const url = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}`;

  await chrome.downloads.download({
    url,
    filename,
    saveAs: true
  });
}


