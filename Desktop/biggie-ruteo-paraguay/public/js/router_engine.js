/**
 * Motor de Ruteo Inteligente y Optimización de Rutas para Biggie Paraguay
 * Soporta asignación directa por Nombres de Sucursales Biggie y por Zonas,
 * optimización TSP con 2-Opt y cálculo de tiempos y costos en Guaraníes (₲).
 */

const HUB_BIGGIE_DEFAULT = {
  id: 'HUB-01',
  name: 'Biggie Sacramento (Matriz)',
  address: 'Avda. Santísimo Sacramento c/ Roque Centurión Miranda',
  zone: 'Manorá',
  city: 'Asunción',
  lat: -25.2818,
  lng: -57.5870,
  type: 'hub'
};

const URBAN_WINDING_FACTOR = 1.32;
const AVG_SPEED_KMH = 26;
const SERVICE_TIME_MINS = 10;
const FUEL_PRICE_PYG_LITER = 7800;
const FUEL_LITERS_PER_100KM = 11.5;

function haversineDistance(lat1, lon1, lat2, lon2) {
  if (lat1 === lat2 && lon1 === lon2) return 0;
  const R = 6371;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = 
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * 
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function routeDistance(p1, p2) {
  const direct = haversineDistance(Number(p1.lat), Number(p1.lng), Number(p2.lat), Number(p2.lng));
  return direct * URBAN_WINDING_FACTOR;
}

function nearestNeighborTour(startPoint, points) {
  if (!points || points.length === 0) return [];
  const unvisited = [...points];
  const tour = [];
  let current = startPoint;

  while (unvisited.length > 0) {
    let nearestIdx = 0;
    let minDistance = Infinity;

    for (let i = 0; i < unvisited.length; i++) {
      const dist = routeDistance(current, unvisited[i]);
      if (dist < minDistance) {
        minDistance = dist;
        nearestIdx = i;
      }
    }

    const nextPoint = unvisited.splice(nearestIdx, 1)[0];
    tour.push(nextPoint);
    current = nextPoint;
  }

  return tour;
}

function optimize2Opt(startPoint, initialTour) {
  if (initialTour.length <= 2) return initialTour;

  let tour = [startPoint, ...initialTour];
  let improved = true;
  let iterations = 0;
  const MAX_ITERATIONS = 50;

  function calculateTotalDist(t) {
    let dist = 0;
    for (let i = 0; i < t.length - 1; i++) {
      dist += routeDistance(t[i], t[i + 1]);
    }
    return dist;
  }

  let bestDistance = calculateTotalDist(tour);

  while (improved && iterations < MAX_ITERATIONS) {
    improved = false;
    iterations++;

    for (let i = 1; i < tour.length - 1; i++) {
      for (let k = i + 1; k < tour.length; k++) {
        const newTour = tour.slice(0, i)
          .concat(tour.slice(i, k + 1).reverse())
          .concat(tour.slice(k + 1));
        
        const newDistance = calculateTotalDist(newTour);
        if (newDistance < bestDistance - 0.001) {
          tour = newTour;
          bestDistance = newDistance;
          improved = true;
          break;
        }
      }
      if (improved) break;
    }
  }

  return tour.slice(1);
}

function buildEmployeeRoute(employee, stops, hub = HUB_BIGGIE_DEFAULT, returnToHub = true) {
  if (!stops || stops.length === 0) {
    return {
      employee,
      stops: [],
      metrics: {
        totalStops: 0,
        totalDistanceKm: 0,
        drivingTimeMins: 0,
        serviceTimeMins: 0,
        totalTimeMins: 0,
        totalTimeFormatted: '0 min',
        estimatedFuelLiters: 0,
        estimatedFuelCostPyg: 0,
        estimatedFuelCostFormatted: '₲ 0'
      }
    };
  }

  const initialTour = nearestNeighborTour(hub, stops);
  const optimizedTour = optimize2Opt(hub, initialTour);

  let currentLoc = hub;
  let totalDist = 0;
  const detailedStops = [];

  for (let i = 0; i < optimizedTour.length; i++) {
    const stop = optimizedTour[i];
    const legDist = routeDistance(currentLoc, stop);
    totalDist += legDist;
    const legDriveTimeMins = Math.round((legDist / AVG_SPEED_KMH) * 60);

    detailedStops.push({
      stopNumber: i + 1,
      id: stop.id || ('STOP-' + (i + 1)),
      name: stop.name,
      address: stop.address || '',
      zone: stop.zone || '',
      city: stop.city || 'Gran Asunción',
      lat: Number(stop.lat),
      lng: Number(stop.lng),
      phone: stop.phone || '',
      orderNumber: stop.orderNumber || stop.invoice || stop.id || '',
      packages: stop.packages || stop.cajas || 1,
      priority: stop.priority || 'Normal',
      notes: stop.notes || '',
      distanceFromPrevKm: Number(legDist.toFixed(2)),
      cumulativeDistanceKm: Number(totalDist.toFixed(2)),
      estimatedDriveTimeMins: legDriveTimeMins,
      estimatedArrivalMin: Math.round(((totalDist / AVG_SPEED_KMH) * 60) + (i * SERVICE_TIME_MINS)),
      status: 'pending'
    });

    currentLoc = stop;
  }

  let returnDist = 0;
  if (returnToHub && optimizedTour.length > 0) {
    returnDist = routeDistance(currentLoc, hub);
    totalDist += returnDist;
  }

  const drivingTimeMins = Math.round((totalDist / AVG_SPEED_KMH) * 60);
  const serviceTimeMins = detailedStops.length * SERVICE_TIME_MINS;
  const totalTimeMins = drivingTimeMins + serviceTimeMins;
  
  const hours = Math.floor(totalTimeMins / 60);
  const mins = totalTimeMins % 60;
  const totalTimeFormatted = hours > 0 ? (hours + 'h ' + mins + 'm') : (mins + ' min');

  const fuelLiters = (totalDist * FUEL_LITERS_PER_100KM) / 100;
  const fuelCostPyg = Math.round(fuelLiters * FUEL_PRICE_PYG_LITER);
  const fuelCostFormatted = '₲ ' + fuelCostPyg.toLocaleString('es-PY');

  return {
    employee,
    hub,
    returnToHub,
    returnDistanceKm: Number(returnDist.toFixed(2)),
    stops: detailedStops,
    metrics: {
      totalStops: detailedStops.length,
      totalDistanceKm: Number(totalDist.toFixed(2)),
      drivingTimeMins,
      serviceTimeMins,
      totalTimeMins,
      totalTimeFormatted,
      estimatedFuelLiters: Number(fuelLiters.toFixed(2)),
      estimatedFuelCostPyg: fuelCostPyg,
      estimatedFuelCostFormatted: fuelCostFormatted
    }
  };
}

function normalizeStr(s) {
  if (!s) return '';
  return s.toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

/**
 * Distribuye sucursales Biggie entre empleados:
 * 1. Primero por coincidencia directa de Nombre de Sucursal Biggie asignada (emp.assignedStores).
 * 2. Segundo por Zona asignada (emp.assignedZones).
 * 3. Fallback por cercanía geográfica.
 */
function planMultiEmployeeRoutes(stops, employees, hub = HUB_BIGGIE_DEFAULT, returnToHub = true) {
  const activeEmployees = employees.filter(e => e.active !== false);
  const employeeBuckets = new Map();
  
  activeEmployees.forEach(emp => {
    employeeBuckets.set(emp.id, []);
  });

  const unassignedStops = [];

  stops.forEach(stop => {
    const stopNameNorm = normalizeStr(stop.name);
    const stopZoneNorm = normalizeStr(stop.zone);
    let assigned = false;

    // Prioridad 1: Coincidencia directa por Nombre de Biggie
    for (const emp of activeEmployees) {
      const assignedStores = emp.assignedStores || [];
      const hasDirectStore = assignedStores.some(storeName => {
        const storeNorm = normalizeStr(storeName);
        return storeNorm.includes(stopNameNorm) || stopNameNorm.includes(storeNorm.replace('biggie', ''));
      });

      if (hasDirectStore) {
        employeeBuckets.get(emp.id).push(stop);
        assigned = true;
        break;
      }
    }

    // Prioridad 2: Coincidencia por Zona
    if (!assigned) {
      for (const emp of activeEmployees) {
        const assignedZones = emp.assignedZones || [];
        const hasZone = assignedZones.some(z => {
          const empZoneNorm = normalizeStr(z);
          return empZoneNorm.includes(stopZoneNorm) || stopZoneNorm.includes(empZoneNorm);
        });

        if (hasZone) {
          employeeBuckets.get(emp.id).push(stop);
          assigned = true;
          break;
        }
      }
    }

    if (!assigned) {
      unassignedStops.push(stop);
    }
  });

  // Asignar paradas restantes por proximidad
  if (unassignedStops.length > 0 && activeEmployees.length > 0) {
    unassignedStops.forEach(stop => {
      let bestEmp = null;
      let minDistance = Infinity;

      for (const emp of activeEmployees) {
        const empStops = employeeBuckets.get(emp.id);
        const refPoint = empStops.length > 0 ? empStops[empStops.length - 1] : hub;
        const d = routeDistance(refPoint, stop);
        if (d < minDistance) {
          minDistance = d;
          bestEmp = emp;
        }
      }

      if (bestEmp) {
        employeeBuckets.get(bestEmp.id).push(Object.assign({}, stop, { autoAssigned: true }));
      }
    });
  }

  const routes = [];
  let grandTotalDistance = 0;
  let grandTotalTimeMins = 0;
  let grandTotalFuelPyg = 0;
  let grandTotalStops = 0;

  activeEmployees.forEach(emp => {
    const empStops = employeeBuckets.get(emp.id) || [];
    const empRoute = buildEmployeeRoute(emp, empStops, hub, returnToHub);
    routes.push(empRoute);

    grandTotalDistance += empRoute.metrics.totalDistanceKm;
    grandTotalTimeMins += empRoute.metrics.totalTimeMins;
    grandTotalFuelPyg += empRoute.metrics.estimatedFuelCostPyg;
    grandTotalStops += empRoute.metrics.totalStops;
  });

  const grandHours = Math.floor(grandTotalTimeMins / 60);
  const grandMins = grandTotalTimeMins % 60;

  return {
    hub,
    routes,
    summary: {
      totalEmployees: activeEmployees.length,
      activeRoutes: routes.filter(r => r.stops.length > 0).length,
      totalStops: grandTotalStops,
      grandTotalDistanceKm: Number(grandTotalDistance.toFixed(2)),
      grandTotalTimeFormatted: grandHours > 0 ? (grandHours + 'h ' + grandMins + 'm') : (grandMins + ' min'),
      grandTotalFuelCostFormatted: '₲ ' + grandTotalFuelPyg.toLocaleString('es-PY'),
      grandTotalFuelCostPyg: grandTotalFuelPyg,
      generatedAt: new Date().toISOString()
    }
  };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    haversineDistance,
    routeDistance,
    nearestNeighborTour,
    optimize2Opt,
    buildEmployeeRoute,
    planMultiEmployeeRoutes,
    HUB_BIGGIE_DEFAULT
  };
}