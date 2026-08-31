/**
 * BIGGIE PARAGUAY - SISTEMA DE RUTEO INTELIGENTE LOCAL
 * Controlador Principal con Asignación por Nombres de Sucursales Biggie y Zonas
 */

const AppState = {
  employees: [],
  zones: [],
  biggieStores: [],
  uploadedStops: [],
  routingPlan: null,
  activeFilterEmpId: 'ALL',
  map: null,
  mapLayers: {
    markersGroup: null,
    routesGroup: null
  },
  editingEmployeeId: null,
  editingZoneId: null,
  editingStoreId: null
};

document.addEventListener('DOMContentLoaded', async () => {
  console.log('Iniciando Sistema de Ruteo Biggie Paraguay...');
  initNavigationTabs();
  initMap();
  await loadInitialData();
  initDropzone();
  initEventListeners();
  renderBiggieStoresTable();
  console.log('Sistema de Ruteo Listo.');
});

function initNavigationTabs() {
  const tabButtons = document.querySelectorAll('.tab-btn');
  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetTab = btn.getAttribute('data-tab');
      tabButtons.forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.style.display = 'none');
      btn.classList.add('active');
      const content = document.getElementById(targetTab);
      if (content) {
        content.style.display = 'block';
      }
      if (targetTab === 'tab-routing' && AppState.map) {
        setTimeout(() => AppState.map.invalidateSize(), 200);
      }
    });
  });
}

function initMap() {
  const mapContainer = document.getElementById('map-container');
  if (!mapContainer) return;
  const ASU_LAT = -25.2980;
  const ASU_LNG = -57.5750;

  AppState.map = L.map('map-container').setView([ASU_LAT, ASU_LNG], 12);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '© OpenStreetMap &bull; Biggie Paraguay'
  }).addTo(AppState.map);

  AppState.mapLayers.markersGroup = L.layerGroup().addTo(AppState.map);
  AppState.mapLayers.routesGroup = L.layerGroup().addTo(AppState.map);
  drawHubMarker();
}

function drawHubMarker() {
  const hub = HUB_BIGGIE_DEFAULT;
  const hubIcon = L.divIcon({
    className: 'hub-leaflet-pin',
    html: '<span style="padding: 4px 6px;">🏢 MATRIZ</span>',
    iconSize: [80, 26],
    iconAnchor: [40, 13]
  });

  L.marker([hub.lat, hub.lng], { icon: hubIcon })
    .addTo(AppState.mapLayers.markersGroup)
    .bindPopup(`
      <div style="font-family: sans-serif; padding: 4px;">
        <strong style="color: #D90429; font-size: 13px;">🏢 ${hub.name}</strong><br>
        <span style="font-size: 11px; color: #555;">${hub.address}</span><br>
        <span style="display:inline-block; margin-top:4px; font-size:10px; background:#FFB703; color:#000; padding:2px 6px; border-radius:3px; font-weight:bold;">
          PUNTO DE SALIDA & LLEGADA
        </span>
      </div>
    `);
}

async function loadInitialData() {
  try {
    const resEmp = await fetch('/api/employees');
    AppState.employees = await resEmp.json();
  } catch (e) {
    AppState.employees = (typeof DEFAULT_EMPLOYEES !== 'undefined') ? DEFAULT_EMPLOYEES : [];
  }

  try {
    const resZones = await fetch('/api/zones');
    AppState.zones = await resZones.json();
  } catch (e) {
    AppState.zones = (typeof STANDARD_ZONES !== 'undefined') ? STANDARD_ZONES : [];
  }

  try {
    const resStores = await fetch('/api/biggie-stores');
    AppState.biggieStores = await resStores.json();
  } catch (e) {
    AppState.biggieStores = (typeof BIGGIE_STORES !== 'undefined') ? BIGGIE_STORES : [];
  }

  renderEmployeesSummary();
  renderEmployeesAdmin();
  renderZonesAdmin();
}
function initDropzone() {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('excel-file-input');

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleExcelFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      handleExcelFile(e.target.files[0]);
    }
  });
}

function handleExcelFile(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

      if (!rawRows || rawRows.length === 0) {
        alert('El archivo Excel está vacío.');
        return;
      }

      processParsedRows(rawRows, file.name);
    } catch (err) {
      alert('Error al leer el archivo Excel: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function cleanStr(s) {
  if (!s) return '';
  return String(s).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function matchLocalBiggieStore(nameQuery, addressQuery, zoneQuery, cityQuery) {
  const qName = cleanStr(nameQuery);
  const qAddr = cleanStr(addressQuery);
  const qZone = cleanStr(zoneQuery);
  const qCity = cleanStr(cityQuery);

  // Prioridad 1: Búsqueda exacta o cercana por nombre
  for (const store of AppState.biggieStores) {
    const sName = cleanStr(store.name);
    const sAddr = cleanStr(store.address);
    const sZone = cleanStr(store.zone);
    const sCity = cleanStr(store.city);

    // Búsqueda por nombre (incluyendo palabras parciales)
    if (qName && qName.length > 2) {
      // Si el query contiene parte del nombre de la tienda
      if (sName.includes(qName) || qName.includes(sName.replace('biggie ', '').trim())) {
        return store;
      }
      // Búsqueda de palabras clave de tienda en el query
      const storeKeywords = sName.replace('biggie', '').trim().split(/\s+/).filter(w => w.length > 2);
      if (storeKeywords.some(kw => qName.includes(kw))) {
        return store;
      }
    }
    // Búsqueda por dirección
    if (qAddr && qAddr.length > 3 && sAddr.includes(qAddr)) {
      return store;
    }
  }

  // Prioridad 2: Búsqueda por zona
  if (qZone && qZone.length > 2) {
    for (const store of AppState.biggieStores) {
      const sZone = cleanStr(store.zone);
      if (sZone.includes(qZone) || qZone.includes(sZone)) {
        return store;
      }
    }
  }

  // Prioridad 3: Búsqueda por ciudad y zona
  if (qCity && qCity.length > 2 && qZone && qZone.length > 2) {
    for (const store of AppState.biggieStores) {
      const sCity = cleanStr(store.city);
      const sZone = cleanStr(store.zone);
      if (sCity.includes(qCity) && sZone.includes(qZone)) {
        return store;
      }
    }
  }

  return null;
}

function processParsedRows(rawRows, fileName = 'Archivo cargado') {
  const parsedStops = [];
  let counter = 1;

  rawRows.forEach(row => {
    const idVal = row['ID'] || row['ID Sucursal'] || row['Id'] || row['id'] || ('STOP-' + counter);
    let nameVal = row['Nombre'] || row['Nombre Sucursal'] || row['Sucursal'] || row['Destino'] || ('Sucursal Biggie ' + counter);
    const addrVal = row['Direccion'] || row['Dirección'] || row['Ubicacion'] || '';
    const cityVal = row['Ciudad'] || row['Localidad'] || 'Asunción';
    let zoneVal = row['Zona'] || row['Sector'] || '';
    let latVal = parseFloat(row['Latitud'] || row['Lat'] || row['lat'] || 0);
    let lngVal = parseFloat(row['Longitud'] || row['Long'] || row['Lng'] || row['lng'] || 0);
    const orderVal = row['Nro Pedido'] || row['Pedido'] || row['Factura'] || ('PED-' + counter);
    const packagesVal = parseInt(row['Cajas/Bultos'] || row['Cajas'] || row['Bultos'] || 1, 10);
    const priorityVal = row['Prioridad'] || 'Normal';
    const phoneVal = row['Telefono'] || row['Teléfono'] || row['Telefono Contacto'] || '';
    const notesVal = row['Notas'] || row['Observaciones'] || '';

    // Intentar matching progresivo
    let matched = matchLocalBiggieStore(nameVal, addrVal, zoneVal, cityVal);
    
    if (matched) {
      nameVal = matched.name;
      if (!latVal || !lngVal || isNaN(latVal) || isNaN(lngVal)) {
        latVal = matched.lat;
        lngVal = matched.lng;
      }
      if (!zoneVal) {
        zoneVal = matched.zone;
      }
    } else {
      // Si no hay match, buscar por zona
      if (zoneVal) {
        const byZone = AppState.biggieStores.filter(s => cleanStr(s.zone) === cleanStr(zoneVal));
        if (byZone.length > 0) {
          const usedNames = parsedStops.map(p => p.name);
          const available = byZone.filter(s => !usedNames.includes(s.name));
          if (available.length > 0) {
            const selectedStore = available[0];
            nameVal = selectedStore.name;
            if (!latVal || !lngVal || isNaN(latVal) || isNaN(lngVal)) {
              latVal = selectedStore.lat;
              lngVal = selectedStore.lng;
            }
            zoneVal = selectedStore.zone;
          }
        }
      } 
      
      // Si aún no hay nombre válido de una sucursal real, buscar en cualquier zona disponible
      if (nameVal.includes('Sucursal Biggie') && AppState.biggieStores.length > 0) {
        const usedNames = parsedStops.map(p => p.name);
        const available = AppState.biggieStores.filter(s => !usedNames.includes(s.name));
        if (available.length > 0) {
          const selectedStore = available[0];
          nameVal = selectedStore.name;
          if (!latVal || !lngVal || isNaN(latVal) || isNaN(lngVal)) {
            latVal = selectedStore.lat;
            lngVal = selectedStore.lng;
          }
          zoneVal = selectedStore.zone;
        }
      }
    }

    if (!latVal || !lngVal || isNaN(latVal) || isNaN(lngVal)) {
      latVal = -25.2950 + ((Math.random() - 0.5) * 0.08);
      lngVal = -57.5800 + ((Math.random() - 0.5) * 0.08);
    }

if (!zoneVal) {
      zoneVal = 'Asunción';
    }

    parsedStops.push({
      id: String(idVal),
      name: String(nameVal),
      address: String(addrVal),
      city: String(cityVal),
      zone: String(zoneVal),
      lat: latVal,
      lng: lngVal,
      orderNumber: String(orderVal),
      packages: isNaN(packagesVal) ? 1 : packagesVal,
      priority: String(priorityVal),
      phone: String(phoneVal),
      notes: String(notesVal)
    });

    counter++;
  });

  AppState.uploadedStops = parsedStops;
  document.getElementById('stops-counter-badge').textContent = `${parsedStops.length} Cargadas`;
  
  renderEmployeesSummary();
  previewStopsOnMap(parsedStops);
  alert(`✅ Se cargaron exitosamente ${parsedStops.length} sucursales desde "${fileName}".\nPresiona "Generar y Optimizar Ruteo Automático" para trazar la ruta.`);
}

function previewStopsOnMap(stops) {
  AppState.mapLayers.markersGroup.clearLayers();
  AppState.mapLayers.routesGroup.clearLayers();
  drawHubMarker();

  const bounds = L.latLngBounds([HUB_BIGGIE_DEFAULT.lat, HUB_BIGGIE_DEFAULT.lng]);

  stops.forEach((s, idx) => {
    bounds.extend([s.lat, s.lng]);
    const pinIcon = L.divIcon({
      className: 'custom-leaflet-pin',
      html: `<span>${idx + 1}</span>`,
      iconSize: [24, 24],
      iconAnchor: [12, 12]
    });

    L.marker([s.lat, s.lng], { icon: pinIcon })
      .addTo(AppState.mapLayers.markersGroup)
      .bindPopup(`
        <div style="font-family: sans-serif; padding: 4px;">
          <strong style="color: #D90429; font-size: 12px;">${s.name}</strong><br>
          <span style="font-size: 11px;">📍 ${s.address || s.city}</span><br>
          <span style="font-size: 11px;">🏷️ Zona: <strong>${s.zone}</strong></span><br>
          <span style="font-size: 11px;">📦 Pedido: ${s.orderNumber} (${s.packages} cajas)</span>
        </div>
      `);
  });

  if (stops.length > 0) {
    AppState.map.fitBounds(bounds, { padding: [40, 40] });
  }
}

function renderEmployeesSummary() {
  const container = document.getElementById('active-employees-summary');
  if (!container) return;

  container.innerHTML = '';

  AppState.employees.forEach(emp => {
    const assignedStores = (emp.assignedStores || []).map(s => normalizeStr(s));
    const assignedZones = (emp.assignedZones || []).map(z => normalizeStr(z));

    const matchingStops = AppState.uploadedStops.filter(s => {
      const sNameNorm = normalizeStr(s.name);
      const sZoneNorm = normalizeStr(s.zone);
      const storeMatch = assignedStores.some(as => as.includes(sNameNorm) || sNameNorm.includes(as.replace('biggie', '')));
      const zoneMatch = assignedZones.some(az => az.includes(sZoneNorm) || sZoneNorm.includes(az));
      return storeMatch || zoneMatch;
    }).length;

    // Biggies por asignación directa + los pertenecientes a las zonas asignadas
    const storesFromZones = AppState.zones
      .filter(z => assignedZones.some(az => normalizeStr(z.name) === az))
      .flatMap(z => z.stores || []);
    const totalStoresConfigured = new Set([
      ...(emp.assignedStores || []),
      ...storesFromZones
    ]).size;
    const totalZonesConfigured = (emp.assignedZones || []).length;

    const row = document.createElement('div');
    row.style.cssText = 'display:flex; justify-content:space-between; align-items:center; background:rgba(255,255,255,0.04); border-radius:6px; padding:0.4rem 0.6rem; font-size:0.75rem;';
    
    row.innerHTML = `
      <div style="display:flex; align-items:center; gap:0.4rem; overflow:hidden; flex:1;">
        <span class="emp-color-bullet" style="background:${emp.color || '#D90429'}; flex-shrink:0;"></span>
        <div style="overflow:hidden;">
          <strong style="white-space:nowrap; text-overflow:ellipsis; overflow:hidden; display:block;">${emp.name}</strong>
          <span style="color:var(--text-muted); font-size:0.68rem;">
            🏪 ${totalStoresConfigured} Biggies &bull; 📍 ${totalZonesConfigured} Zonas
          </span>
        </div>
      </div>
      <div style="display:flex; align-items:center; gap:0.4rem; flex-shrink:0;">
        <span class="zone-tag" style="font-size:0.68rem; padding:0.1rem 0.35rem;">${matchingStops} tiendas en ruta</span>
        <button class="btn btn-secondary" onclick="openEditEmployeeModal('${emp.id}')" style="padding:0.15rem 0.4rem; font-size:0.7rem;" title="Modificar Empleado y sus Biggies">
          ✏️
        </button>
      </div>
    `;
    container.appendChild(row);
  });
}

function executeRouteOptimization() {
  if (!AppState.uploadedStops || AppState.uploadedStops.length === 0) {
    alert('Primero debes cargar un archivo Excel o presionar "Cargar Demo Biggie".');
    return;
  }

  const returnToHub = document.getElementById('chk-return-hub').checked;
  const result = planMultiEmployeeRoutes(AppState.uploadedStops, AppState.employees, HUB_BIGGIE_DEFAULT, returnToHub);

  AppState.routingPlan = result;

  document.getElementById('kpi-stops').textContent = result.summary.totalStops;
  document.getElementById('kpi-distance').textContent = `${result.summary.grandTotalDistanceKm} km`;
  document.getElementById('kpi-time').textContent = result.summary.grandTotalTimeFormatted;
  document.getElementById('kpi-fuel').textContent = result.summary.grandTotalFuelCostFormatted;

  document.getElementById('btn-export-excel').disabled = false;

  renderRoutesOnMap(result);
  renderEmployeeFilterPills(result.routes);
  renderRouteTable('ALL');
}

function renderRoutesOnMap(plan) {
  AppState.mapLayers.markersGroup.clearLayers();
  AppState.mapLayers.routesGroup.clearLayers();
  drawHubMarker();

  const bounds = L.latLngBounds([HUB_BIGGIE_DEFAULT.lat, HUB_BIGGIE_DEFAULT.lng]);

  plan.routes.forEach(route => {
    if (!route.stops || route.stops.length === 0) return;

    const empColor = route.employee.color || '#3B82F6';
    const latLngs = [];

    latLngs.push([HUB_BIGGIE_DEFAULT.lat, HUB_BIGGIE_DEFAULT.lng]);

    route.stops.forEach(stop => {
      bounds.extend([stop.lat, stop.lng]);
      latLngs.push([stop.lat, stop.lng]);

      const pinIcon = L.divIcon({
        className: 'custom-leaflet-pin',
        html: `<span style="background:${empColor}; border-color:#FFF;">${stop.stopNumber}</span>`,
        iconSize: [26, 26],
        iconAnchor: [13, 13]
      });

      L.marker([stop.lat, stop.lng], { icon: pinIcon })
        .addTo(AppState.mapLayers.markersGroup)
        .bindPopup(`
          <div style="font-family: sans-serif; padding: 4px;">
            <span style="font-size:10px; background:${empColor}; color:#FFF; padding:2px 6px; border-radius:3px; font-weight:bold;">
              PARADA #${stop.stopNumber} &bull; ${route.employee.name}
            </span><br>
            <strong style="color: #111; font-size: 13px; margin-top:4px; display:inline-block;">${stop.name}</strong><br>
            <span style="font-size: 11px; color:#555;">📍 ${stop.address}</span><br>
            <span style="font-size: 11px;">🏷️ Zona: <strong>${stop.zone}</strong></span><br>
            <span style="font-size: 11px;">📦 Pedido: ${stop.orderNumber} (${stop.packages} cajas)</span><br>
            <span style="font-size: 11px;">⏱️ Llegada Estimada: <strong>+${stop.estimatedArrivalMin} min</strong> (a ${stop.distanceFromPrevKm} km)</span>
          </div>
        `);
    });

    if (route.returnToHub) {
      latLngs.push([HUB_BIGGIE_DEFAULT.lat, HUB_BIGGIE_DEFAULT.lng]);
    }

    L.polyline(latLngs, {
      color: empColor,
      weight: 4,
      opacity: 0.85,
      dashArray: '8, 6'
    }).addTo(AppState.mapLayers.routesGroup);
  });

  AppState.map.fitBounds(bounds, { padding: [50, 50] });
}

function renderEmployeeFilterPills(routes) {
  const container = document.getElementById('route-employee-pills');
  container.innerHTML = '';

  const allBtn = document.createElement('div');
  allBtn.className = `emp-chip ${AppState.activeFilterEmpId === 'ALL' ? 'active' : ''}`;
  allBtn.innerHTML = `<span>Todos los Choferes (${routes.reduce((acc, r) => acc + r.stops.length, 0)})</span>`;
  allBtn.onclick = () => {
    AppState.activeFilterEmpId = 'ALL';
    renderEmployeeFilterPills(routes);
    renderRouteTable('ALL');
  };
  container.appendChild(allBtn);

  routes.forEach(r => {
    if (r.stops.length === 0) return;
    const chip = document.createElement('div');
    chip.className = `emp-chip ${AppState.activeFilterEmpId === r.employee.id ? 'active' : ''}`;
    chip.innerHTML = `
      <span class="emp-color-bullet" style="background:${r.employee.color};"></span>
      <span>${r.employee.name} (${r.stops.length} paradas - ${r.metrics.totalDistanceKm} km)</span>
    `;
    chip.onclick = () => {
      AppState.activeFilterEmpId = r.employee.id;
      renderEmployeeFilterPills(routes);
      renderRouteTable(r.employee.id);
    };
    container.appendChild(chip);
  });
}

function renderRouteTable(empFilterId = 'ALL') {
  const tbody = document.getElementById('route-table-body');
  if (!AppState.routingPlan || !AppState.routingPlan.routes) return;

  tbody.innerHTML = '';

  let allStops = [];

  AppState.routingPlan.routes.forEach(route => {
    if (empFilterId === 'ALL' || empFilterId === route.employee.id) {
      route.stops.forEach(s => {
        allStops.push({
          ...s,
          employeeName: route.employee.name,
          employeeColor: route.employee.color
        });
      });
    }
  });

  if (allStops.length === 0) {
    tbody.innerHTML = `<tr><td colspan="11" style="text-align:center; color:var(--text-muted); padding:1.5rem;">No hay paradas para este chofer.</td></tr>`;
    return;
  }

  allStops.forEach(s => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="stop-badge" style="background:${s.employeeColor || '#D90429'};">#${s.stopNumber}</span></td>
      <td><strong>${s.name}</strong></td>
      <td><span class="zone-tag">${s.zone}</span></td>
      <td><span style="font-weight:600; color:${s.employeeColor};">${s.employeeName}</span></td>
      <td>${s.distanceFromPrevKm} km</td>
    `;
    tbody.appendChild(tr);
  });
}

function exportRoutesToExcel() {
  try {
    if (!AppState.routingPlan || !AppState.routingPlan.routes) {
      alert('❌ No hay rutas generadas para exportar. Primero genera las rutas con "Generar y Optimizar Ruteo Automático".');
      return;
    }

    if (typeof XLSX === 'undefined') {
      alert('❌ Error: Librería XLSX no cargada. Recarga la página.');
      return;
    }

    const workbook = XLSX.utils.book_new();

    // Recorrer todas las rutas y armar una fila por sucursal visitada
    const rows = [];
    AppState.routingPlan.routes.forEach(route => {
      if (route.stops && route.stops.length > 0) {
        const employee = route.employee.name || '';
        route.stops.forEach(stop => {
          rows.push({
            'Empleado': employee,
            'Sucursal': stop.name || '',
            'Notas': stop.notes || ''
          });
        });
      }
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [{ wch: 28 }, { wch: 30 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(workbook, ws, 'Empleado y Sucursal');

    // Exportar archivo
    const dateStr = new Date().toISOString().slice(0, 10);
    const fileName = `Rutas_Biggie_Paraguay_${dateStr}.xlsx`;
    XLSX.writeFile(workbook, fileName);
    
    alert('✅ ARCHIVO EXPORTADO EXITOSAMENTE\n📁 ' + fileName + '\n\nVerifica tu carpeta de descargas.');
  } catch (error) {
    console.error('Error en exportación:', error);
    alert('❌ Error al exportar: ' + (error.message || error));
  }
}

// ==========================================================================
// GESTIÓN DE EMPLEADOS CON SELECCIÓN POR NOMBRES DE BIGGIES Y ZONAS
// ==========================================================================
function renderEmployeesAdmin() {
  const container = document.getElementById('employees-list-container');
  if (!container) return;

  container.innerHTML = '';

  AppState.employees.forEach(emp => {
    const card = document.createElement('div');
    card.style.cssText = `background:var(--bg-subtle); border:1px solid ${emp.active !== false ? 'var(--border-color)' : '#991B1B'}; border-radius:8px; padding:0.85rem; display:flex; flex-direction:column; gap:0.5rem; opacity:${emp.active !== false ? '1' : '0.6'};`;
    
    const storeBadges = (emp.assignedStores || []).map(s => `<span class="zone-tag" style="background:rgba(217,4,41,0.15); color:#EF233C; border-color:rgba(217,4,41,0.3); font-size:0.7rem;">🏪 ${s}</span>`).join(' ');
    const zoneChips = (emp.assignedZones || []).map(z => `<span class="zone-tag" style="font-size:0.7rem;">📍 ${z}</span>`).join(' ');

    card.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:flex-start;">
        <div style="display:flex; align-items:center; gap:0.5rem;">
          <span class="emp-color-bullet" style="background:${emp.color || '#3B82F6'}; width:14px; height:14px;"></span>
          <div>
            <strong style="font-size:0.95rem;">${emp.name}</strong>
            ${emp.active === false ? '<span style="color:#EF4444; font-size:0.72rem; margin-left:0.3rem;">[Inactivo]</span>' : ''}
          </div>
        </div>
        <div style="display:flex; gap:0.35rem;">
          <button class="btn btn-secondary" onclick="openEditEmployeeModal('${emp.id}')" style="padding:0.25rem 0.55rem; font-size:0.75rem;">
            ✏️ Modificar
          </button>
          <button class="btn btn-secondary" onclick="deleteEmployee('${emp.id}')" style="padding:0.25rem 0.55rem; font-size:0.75rem; color:#EF4444;" title="Eliminar Empleado">
            🗑️
          </button>
        </div>
      </div>
      
      <div style="font-size:0.78rem; font-weight:700; color:var(--biggie-yellow); margin-top:0.15rem;">
        SUCURSALES BIGGIE ASIGNADAS (${(emp.assignedStores || []).length}):
      </div>
      <div style="display:flex; flex-wrap:wrap; gap:0.35rem; max-height:80px; overflow-y:auto;">
        ${storeBadges || '<span style="color:var(--text-muted); font-size:0.75rem;">Sin Biggies asignados directamente</span>'}
      </div>

      <div style="font-size:0.75rem; color:var(--text-muted); margin-top:0.15rem;">
        Zonas de cobertura: ${zoneChips || '<span style="color:var(--text-muted);">Sin zonas</span>'}
      </div>
    `;
    container.appendChild(card);
  });
}

function renderBiggieStoreCheckboxes(selectedStores = [], selectedZones = []) {
  const container = document.getElementById('emp-form-zones-container');
  if (!container) return;

  container.innerHTML = '';

  AppState.zones.forEach(zone => {
    // Obtener las tiendas de esta zona (de zone.stores o de AppState.biggieStores)
    let zoneStores = zone.stores || [];
    if (!zoneStores || zoneStores.length === 0) {
      zoneStores = AppState.biggieStores
        .filter(s => normalizeStr(s.zone) === normalizeStr(zone.name))
        .map(s => s.name);
    }

    const isZoneChecked = selectedZones.includes(zone.name);

    const zoneGroup = document.createElement('div');
    zoneGroup.style.cssText = 'background:rgba(255,255,255,0.03); border:1px solid var(--border-color); border-radius:8px; padding:0.6rem;';

    const zoneHeader = document.createElement('div');
    zoneHeader.style.cssText = 'display:flex; justify-content:space-between; align-items:center;';
    zoneHeader.innerHTML = `
      <label style="display:flex; align-items:center; gap:0.45rem; font-weight:700; font-size:0.85rem; cursor:pointer;">
        <input type="checkbox" class="zone-master-checkbox" data-zone="${zone.name}" ${isZoneChecked ? 'checked' : ''} style="width:16px; height:16px;">
        <span class="emp-color-bullet" style="background:${zone.color || '#FFB703'};"></span>
        <span>${zone.name} <span style="font-weight:400; color:var(--text-muted); font-size:0.72rem;">(${zoneStores.length} Biggies)</span></span>
      </label>
    `;
    zoneGroup.appendChild(zoneHeader);
    container.appendChild(zoneGroup);
  });

  // Nota informativa: los Biggies se asignan automáticamente según la zona elegida
  const note = document.createElement('div');
  note.style.cssText = 'grid-column: 1 / -1; color:var(--text-muted); font-size:0.72rem; padding:0.4rem; background:rgba(0,0,0,0.15); border-radius:6px;';
  note.textContent = '💡 Al elegir una zona, todas sus sucursales Biggie se asignan automáticamente al empleado.';
  container.appendChild(note);
}

function openCreateEmployeeModal() {
  document.getElementById('modal-employee-title').textContent = '➕ Nuevo Empleado / Chofer';
  document.getElementById('emp-form-id').value = '';
  document.getElementById('emp-form-name').value = '';
  
  const colors = ['#EF4444', '#10B981', '#3B82F6', '#8B5CF6', '#EC4899', '#F59E0B', '#14B8A6'];
  const randomColor = colors[Math.floor(Math.random() * colors.length)];
  document.getElementById('emp-form-color').value = randomColor;
  document.getElementById('emp-form-color-hex').textContent = randomColor;
  document.getElementById('emp-form-active').checked = true;

  renderBiggieStoreCheckboxes([], []);

  const modal = document.getElementById('modal-employee-form');
  modal.classList.add('open');
}

function openEditEmployeeModal(empId) {
  const emp = AppState.employees.find(e => e.id === empId);
  if (!emp) return;

  document.getElementById('modal-employee-title').textContent = `✏️ Modificar a ${emp.name}`;
  document.getElementById('emp-form-id').value = emp.id;
  document.getElementById('emp-form-name').value = emp.name;
  document.getElementById('emp-form-color').value = emp.color || '#3B82F6';
  document.getElementById('emp-form-color-hex').textContent = emp.color || '#3B82F6';
  document.getElementById('emp-form-active').checked = emp.active !== false;

  renderBiggieStoreCheckboxes(emp.assignedStores || [], emp.assignedZones || []);

  const modal = document.getElementById('modal-employee-form');
  modal.classList.add('open');
}

async function handleSaveEmployeeForm(e) {
  e.preventDefault();

  const id = document.getElementById('emp-form-id').value;
  const name = document.getElementById('emp-form-name').value.trim();
  const color = document.getElementById('emp-form-color').value;
  const active = document.getElementById('emp-form-active').checked;

  // Zonas seleccionadas
  const checkedZoneBoxes = document.querySelectorAll('.zone-master-checkbox:checked');
  const assignedZones = Array.from(checkedZoneBoxes).map(cb => cb.getAttribute('data-zone'));

  // Los Biggies se asignan automáticamente según las zonas elegidas
  const assignedStores = [];
  AppState.zones.forEach(zone => {
    if (!assignedZones.includes(zone.name)) return;
    let zoneStores = zone.stores || [];
    if (!zoneStores || zoneStores.length === 0) {
      zoneStores = AppState.biggieStores
        .filter(s => normalizeStr(s.zone) === normalizeStr(zone.name))
        .map(s => s.name);
    }
    zoneStores.forEach(storeName => {
      if (!assignedStores.includes(storeName)) assignedStores.push(storeName);
    });
  });

  const payload = {
    name,
    color,
    active,
    assignedStores,
    assignedZones
  };

  try {
    let res;
    if (id) {
      res = await fetch(`/api/employees/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
    } else {
      res = await fetch('/api/employees/new', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
    }

    const data = await res.json();
    if (data.success && data.employees) {
      AppState.employees = data.employees;
    }
  } catch (err) {
    console.warn('Error al guardar en backend:', err);
  }

  closeModal('modal-employee-form');
  renderEmployeesSummary();
  renderEmployeesAdmin();
  renderZonesAdmin();

  if (AppState.uploadedStops.length > 0) {
    executeRouteOptimization();
  }
}

async function deleteEmployee(empId) {
  const emp = AppState.employees.find(e => e.id === empId);
  if (!emp) return;

  const confirmDelete = confirm(`¿Estás seguro de que deseas eliminar a "${emp.name}"?`);
  if (!confirmDelete) return;

  try {
    const res = await fetch(`/api/employees/${empId}`, { method: 'DELETE' });
    const data = await res.json();
    if (data.success && data.employees) {
      AppState.employees = data.employees;
    } else {
      AppState.employees = AppState.employees.filter(e => e.id !== empId);
    }
  } catch (err) {
    AppState.employees = AppState.employees.filter(e => e.id !== empId);
  }

  renderEmployeesSummary();
  renderEmployeesAdmin();
  renderZonesAdmin();

  if (AppState.uploadedStops.length > 0) {
    executeRouteOptimization();
  }
}

// ==========================================================================
// GESTIÓN DE ZONAS (MOSTRANDO SUS BIGGIES)
// ==========================================================================
function renderZonesAdmin() {
  const container = document.getElementById('zones-list-container');
  if (!container) return;

  container.innerHTML = '';

  AppState.zones.forEach(zone => {
    const card = document.createElement('div');
    card.style.cssText = 'background:var(--bg-subtle); border:1px solid var(--border-color); border-radius:8px; padding:0.85rem; display:flex; flex-direction:column; gap:0.5rem;';
    
    // Tiendas Biggie de esta zona
    let zoneStores = zone.stores || [];
    if (!zoneStores || zoneStores.length === 0) {
      zoneStores = AppState.biggieStores
        .filter(s => normalizeStr(s.zone) === normalizeStr(zone.name))
        .map(s => s.name);
    }

    const storeBadges = zoneStores.map(s => `<span class="zone-tag" style="background:rgba(217,4,41,0.15); color:#EF233C; border-color:rgba(217,4,41,0.3); font-size:0.7rem;">🏪 ${s}</span>`).join(' ');
    const assignedEmps = AppState.employees.filter(e => {
      const hasZone = (e.assignedZones || []).includes(zone.name);
      const hasAnyStore = (e.assignedStores || []).some(s => zoneStores.includes(s));
      return hasZone || hasAnyStore;
    }).map(e => e.name).join(', ');

    card.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center;">
        <div style="display:flex; align-items:center; gap:0.4rem;">
          <span class="emp-color-bullet" style="background:${zone.color || '#FFB703'};"></span>
          <strong style="font-size:0.95rem;">${zone.name}</strong>
          <span class="zone-tag" style="font-size:0.68rem;">${zoneStores.length} Biggies</span>
        </div>
        <div style="display:flex; gap:0.35rem;">
          <button class="btn btn-secondary" onclick="openEditZoneModal('${zone.id}')" style="padding:0.25rem 0.55rem; font-size:0.75rem;">
            ✏️ Editar
          </button>
          <button class="btn btn-secondary" onclick="deleteZone('${zone.id}')" style="padding:0.25rem 0.55rem; font-size:0.75rem; color:#EF4444;" title="Eliminar Zona">
            🗑️
          </button>
        </div>
      </div>
      
      <div style="font-size:0.75rem; color:var(--text-muted);">
        ${zone.description}
      </div>

      <!-- Sucursales Biggie dentro de esta zona -->
      <div style="display:flex; flex-wrap:wrap; gap:0.35rem; margin-top:0.2rem;">
        ${storeBadges || '<span style="color:var(--text-muted); font-size:0.72rem;">Sin tiendas asignadas a esta zona</span>'}
      </div>

      <div style="font-size:0.72rem; color:var(--biggie-yellow); margin-top:0.2rem;">
        👤 Choferes que atienden esta zona o sus Biggies: <strong>${assignedEmps || 'Ninguno'}</strong>
      </div>
    `;
    container.appendChild(card);
  });
}

function openCreateZoneModal() {
  const title = document.getElementById('modal-zone-title');
  const inputId = document.getElementById('zone-form-id');
  const inputName = document.getElementById('zone-form-name');
  const inputDesc = document.getElementById('zone-form-desc');
  const inputColor = document.getElementById('zone-form-color');
  const colorHex = document.getElementById('zone-form-color-hex');
  const modal = document.getElementById('modal-zone-form');

  if (!modal) {
    console.warn('No existe el modal de zona');
    return;
  }

  if (title) title.textContent = '📍 Agregar Nueva Zona en Paraguay';
  if (inputId) inputId.value = '';
  if (inputName) inputName.value = '';
  if (inputDesc) inputDesc.value = '';

  const colors = ['#F59E0B', '#10B981', '#3B82F6', '#8B5CF6', '#EC4899', '#14B8A6', '#F43F5E', '#D97706'];
  const randomColor = colors[Math.floor(Math.random() * colors.length)];
  if (inputColor) inputColor.value = randomColor;
  if (colorHex) colorHex.textContent = randomColor;

  modal.classList.add('open');
}

function openEditZoneModal(zoneId) {
  const zone = AppState.zones.find(z => z.id === zoneId);
  if (!zone) return;

  const title = document.getElementById('modal-zone-title');
  const inputId = document.getElementById('zone-form-id');
  const inputName = document.getElementById('zone-form-name');
  const inputDesc = document.getElementById('zone-form-desc');
  const inputColor = document.getElementById('zone-form-color');
  const colorHex = document.getElementById('zone-form-color-hex');
  const modal = document.getElementById('modal-zone-form');

  if (title) title.textContent = `✏️ Editar Zona: ${zone.name}`;
  if (inputId) inputId.value = zone.id;
  if (inputName) inputName.value = zone.name;
  if (inputDesc) inputDesc.value = zone.description || '';
  if (inputColor) inputColor.value = zone.color || '#F59E0B';
  if (colorHex) colorHex.textContent = zone.color || '#F59E0B';

  if (modal) modal.classList.add('open');
}

async function handleSaveZoneForm(e) {
  e.preventDefault();

  const id = document.getElementById('zone-form-id').value;
  const name = document.getElementById('zone-form-name').value.trim();
  const descriptionInput = document.getElementById('zone-form-desc');
  const description = descriptionInput ? descriptionInput.value.trim() : '';
  const color = document.getElementById('zone-form-color').value;

  if (!name) return;

  const payload = { name, description, color };

  try {
    let res;
    if (id) {
      res = await fetch(`/api/zones/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
    } else {
      res = await fetch('/api/zones/new', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
    }

    const data = await res.json();
    if (data.success && data.zones) {
      AppState.zones = data.zones;
    }
  } catch (err) {
    console.warn('Error al guardar zona:', err);
  }

  closeModal('modal-zone-form');
  renderZonesAdmin();
  renderBiggieStoresTable(document.getElementById('input-search-stores')?.value || '');
  renderEmployeesAdmin();
  renderEmployeesSummary();
}

async function deleteZone(zoneId) {
  const zone = AppState.zones.find(z => z.id === zoneId);
  if (!zone) return;

  const confirmDelete = confirm(`¿Estás seguro de que deseas eliminar la zona "${zone.name}"?`);
  if (!confirmDelete) return;

  try {
    const res = await fetch(`/api/zones/${zoneId}`, { method: 'DELETE' });
    const data = await res.json();
    if (data.success && data.zones) {
      AppState.zones = data.zones;
    } else {
      AppState.zones = AppState.zones.filter(z => z.id !== zoneId);
    }
  } catch (err) {
    AppState.zones = AppState.zones.filter(z => z.id !== zoneId);
  }

  renderZonesAdmin();
  renderBiggieStoresTable(document.getElementById('input-search-stores')?.value || '');
  renderEmployeesAdmin();
  renderEmployeesSummary();
}

function closeModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.remove('open');
}

function renderBiggieStoresTable(searchQuery = '') {
  const container = document.getElementById('stores-by-zone-container');
  if (!container) return;

  container.innerHTML = '';
  const q = cleanStr(searchQuery);

  // Obtener y filtrar sucursales
  const filtered = (AppState.biggieStores || []).filter(s => {
    if (!q) return true;
    return cleanStr(s.name).includes(q) || cleanStr(s.address || '').includes(q) || cleanStr(s.city || '').includes(q) || cleanStr(s.zone || '').includes(q);
  });

  // Usar zonas definidas o estándar
  const zones = (AppState.zones && AppState.zones.length > 0) ? AppState.zones : (typeof STANDARD_ZONES !== 'undefined' ? STANDARD_ZONES : []);

  // Si hay búsqueda pero sin resultados, mostrar mensaje
  if (q && filtered.length === 0 && zones.length > 0) {
    container.innerHTML = '<div style="color:var(--text-muted); padding:2rem; text-align:center;">No se encontraron sucursales que coincidan con la búsqueda.</div>';
    return;
  }

  // Agrupar sucursales por zona
  zones.forEach(zone => {
    const zoneStores = filtered.filter(s => normalizeStr(s.zone || '').includes(normalizeStr(zone.name)));
    
    // Si hay búsqueda activa, solo mostrar zonas con resultados
    if (q && zoneStores.length === 0) return;

    const zoneCard = document.createElement('div');
    zoneCard.style.cssText = `background:var(--bg-subtle); border:1px solid var(--border-color); border-radius:8px; padding:1rem; display:flex; flex-direction:column; gap:0.75rem;`;

    const headerColor = zone.color || '#FFB703';
    const zoneHeaderHtml = `
      <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid rgba(255,255,255,0.1); padding-bottom:0.75rem; margin-bottom:0.25rem;">
        <div style="display:flex; align-items:center; gap:0.6rem; flex:1;">
          <span class="emp-color-bullet" style="background:${headerColor}; width:16px; height:16px;"></span>
          <div style="flex:1;">
            <strong style="font-size:1.05rem; color:#FFF;">${zone.name}</strong>
            <span class="zone-tag" style="font-size:0.7rem; margin-left:0.5rem; background:rgba(255,255,255,0.1); color:var(--text-secondary);">📦 ${zoneStores.length} Biggie${zoneStores.length !== 1 ? 's' : ''}</span>
          </div>
        </div>
        <div style="display:flex; gap:0.35rem; flex-shrink:0;">
          <button class="btn btn-primary" onclick="openCreateStoreModal('${zone.name}')" style="padding:0.3rem 0.6rem; font-size:0.75rem; white-space:nowrap;">
            ➕ Agregar Biggie
          </button>
          <button class="btn btn-secondary" onclick="deleteZone('${zone.id}')" title="Eliminar Zona" style="padding:0.3rem 0.6rem; font-size:0.75rem; color:#EF4444; white-space:nowrap;">
            🗑️
          </button>
        </div>
      </div>
    `;
    
    zoneCard.innerHTML = zoneHeaderHtml;

    const listContainer = document.createElement('div');
    listContainer.style.cssText = 'display:flex; flex-direction:column; gap:0.5rem;';
    listContainer.id = `zone-stores-list-${zone.id || zone.name}`;

    if (zoneStores.length === 0) {
      const emptyMsg = document.createElement('div');
      emptyMsg.style.cssText = 'color:var(--text-muted); font-size:0.8rem; padding:1rem; text-align:center; background:rgba(0,0,0,0.15); border-radius:6px;';
      emptyMsg.textContent = '📭 Sin sucursales en esta zona. Agrega la primera.';
      listContainer.appendChild(emptyMsg);
    } else {
      zoneStores.forEach(store => {
        const row = document.createElement('div');
        row.style.cssText = 'display:flex; justify-content:space-between; align-items:flex-start; background:rgba(0,0,0,0.2); border-radius:6px; padding:0.7rem; font-size:0.8rem; border-left:3px solid ' + headerColor + ';';
        
        const storeInfoHtml = `
          <div style="display:flex; flex-direction:column; gap:0.2rem; flex:1; overflow:hidden;">
            <div style="display:flex; align-items:center; gap:0.4rem; flex-wrap:wrap;">
              <strong style="color:#FFF; font-size:0.9rem;">${store.name || 'Sin nombre'}</strong>
              <code style="font-size:0.65rem; color:var(--text-muted); background:rgba(255,255,255,0.05); padding:0.1rem 0.3rem; border-radius:3px;">${store.id || 'N/A'}</code>
            </div>
            <span style="color:var(--text-secondary); font-size:0.75rem;">📍 ${store.address || 'Sin dirección'} — ${store.city || 'Paraguay'}</span>
            <span style="color:var(--text-muted); font-size:0.7rem;">
              🌍 ${store.lat ? store.lat.toFixed(4) : 'N/A'}, ${store.lng ? store.lng.toFixed(4) : 'N/A'}
              ${store.phone ? ' 📞 ' + store.phone : ''}
            </span>
          </div>
          <div style="display:flex; gap:0.3rem; flex-shrink:0; margin-left:0.5rem;">
            <button class="btn btn-secondary" onclick="openEditStoreModal('${store.id}')" style="padding:0.2rem 0.45rem; font-size:0.72rem;" title="Modificar Sucursal">
              ✏️
            </button>
            <button class="btn btn-secondary" onclick="deleteStore('${store.id}')" style="padding:0.2rem 0.45rem; font-size:0.72rem; color:#EF4444;" title="Eliminar Sucursal">
              🗑️
            </button>
          </div>
        `;
        
        row.innerHTML = storeInfoHtml;
        listContainer.appendChild(row);
      });
    }

    zoneCard.appendChild(listContainer);
    container.appendChild(zoneCard);
  });

  // Si no hay zonas, mostrar mensaje
  if (zones.length === 0) {
    container.innerHTML = '<div style="color:var(--text-muted); padding:2rem; text-align:center;"><p>No hay zonas definidas.</p><button class="btn btn-primary" onclick="openCreateZoneModal()" style="margin-top:1rem;">➕ Crear Primera Zona</button></div>';
  }
}

// ==========================================================================
// GESTIÓN DE SUCURSALES BIGGIE (CRUD POR ZONA)
// ==========================================================================
function populateStoreZoneDropdown(selectedZone = '') {
  const select = document.getElementById('store-form-zone');
  if (!select) return;
  select.innerHTML = '';
  AppState.zones.forEach(zone => {
    const opt = document.createElement('option');
    opt.value = zone.name;
    opt.textContent = zone.name;
    if (zone.name === selectedZone) opt.selected = true;
    select.appendChild(opt);
  });
}

function openCreateStoreModal(preselectedZone = '') {
  AppState.editingStoreId = null;
  document.getElementById('modal-store-title').textContent = '➕ Nueva Sucursal Biggie';
  document.getElementById('store-form-id').value = '';
  document.getElementById('store-form-name').value = '';

  populateStoreZoneDropdown(preselectedZone);

  const modal = document.getElementById('modal-store-form');
  modal.classList.add('open');
}

function openEditStoreModal(storeId) {
  const store = AppState.biggieStores.find(s => s.id === storeId);
  if (!store) return;

  AppState.editingStoreId = storeId;
  document.getElementById('modal-store-title').textContent = `✏️ Editar: ${store.name}`;
  document.getElementById('store-form-id').value = store.id;
  document.getElementById('store-form-name').value = store.name;

  populateStoreZoneDropdown(store.zone);

  const modal = document.getElementById('modal-store-form');
  modal.classList.add('open');
}

async function handleSaveStoreForm(e) {
  e.preventDefault();

  const id = document.getElementById('store-form-id').value;
  const name = document.getElementById('store-form-name').value.trim();
  const zone = document.getElementById('store-form-zone').value;

  if (!name || !zone) return;

  // Conservar lat/lng/dirección existentes si se está editando
  const existing = id ? AppState.biggieStores.find(s => s.id === id) : null;
  const payload = {
    name,
    zone,
    address: existing ? existing.address || '' : '',
    city: existing ? existing.city || '' : '',
    phone: existing ? existing.phone || '' : '',
    lat: existing ? existing.lat : '',
    lng: existing ? existing.lng : ''
  };

  try {
    let res;
    if (id) {
      res = await fetch(`/api/biggie-stores/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
    } else {
      res = await fetch('/api/biggie-stores', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
    }

    const data = await res.json();
    if (data.success && data.stores) {
      AppState.biggieStores = data.stores;
    }
  } catch (err) {
    console.warn('Error al guardar sucursal:', err);
  }

  closeModal('modal-store-form');
  renderBiggieStoresTable(document.getElementById('input-search-stores').value);
  renderZonesAdmin();
  renderEmployeesAdmin();
  renderEmployeesSummary();
}

async function deleteStore(storeId) {
  const store = AppState.biggieStores.find(s => s.id === storeId);
  if (!store) return;

  const confirmDelete = confirm(`¿Eliminar la sucursal "${store.name}"?`);
  if (!confirmDelete) return;

  try {
    const res = await fetch(`/api/biggie-stores/${storeId}`, { method: 'DELETE' });
    const data = await res.json();
    if (data.success && data.stores) {
      AppState.biggieStores = data.stores;
    } else {
      AppState.biggieStores = AppState.biggieStores.filter(s => s.id !== storeId);
    }
  } catch (err) {
    AppState.biggieStores = AppState.biggieStores.filter(s => s.id !== storeId);
  }

  renderBiggieStoresTable(document.getElementById('input-search-stores').value);
  renderZonesAdmin();
  renderEmployeesAdmin();
  renderEmployeesSummary();
}

function initEventListeners() {
  const safeAddListener = (id, eventName, handler) => {
    const element = document.getElementById(id);
    if (element) {
      element.addEventListener(eventName, handler);
    }
  };

  safeAddListener('btn-generate-routes', 'click', executeRouteOptimization);

  safeAddListener('btn-load-demo', 'click', () => {
    const demoSample = [
      { ...AppState.biggieStores[1], orderNumber: 'PED-DEMO-01', packages: 12 },
      { ...AppState.biggieStores[2], orderNumber: 'PED-DEMO-02', packages: 8 },
      { ...AppState.biggieStores[5], orderNumber: 'PED-DEMO-03', packages: 15 },
      { ...AppState.biggieStores[6], orderNumber: 'PED-DEMO-04', packages: 20 },
      { ...AppState.biggieStores[7], orderNumber: 'PED-DEMO-05', packages: 10 },
      { ...AppState.biggieStores[12], orderNumber: 'PED-DEMO-06', packages: 14 },
      { ...AppState.biggieStores[13], orderNumber: 'PED-DEMO-07', packages: 18 },
      { ...AppState.biggieStores[14], orderNumber: 'PED-DEMO-08', packages: 11 },
      { ...AppState.biggieStores[17], orderNumber: 'PED-DEMO-09', packages: 16 },
      { ...AppState.biggieStores[18], orderNumber: 'PED-DEMO-10', packages: 22 },
      { ...AppState.biggieStores[22], orderNumber: 'PED-DEMO-11', packages: 9 },
      { ...AppState.biggieStores[26], orderNumber: 'PED-DEMO-12', packages: 15 }
    ];
    processParsedRows(demoSample, 'Demo_Biggie_Paraguay.xlsx');
    setTimeout(() => executeRouteOptimization(), 300);
  });

  safeAddListener('btn-download-template', 'click', () => {
    window.location.href = '/api/download-template';
  });

  safeAddListener('btn-export-excel', 'click', exportRoutesToExcel);

  safeAddListener('btn-print-route', 'click', () => {
    window.print();
  });

  safeAddListener('btn-fit-map', 'click', () => {
    if (AppState.map) {
      AppState.map.setView([-25.2980, -57.5750], 12);
    }
  });

  // Empleados
  safeAddListener('btn-add-employee', 'click', openCreateEmployeeModal);
  safeAddListener('form-employee', 'submit', handleSaveEmployeeForm);
  const employeeColorInput = document.getElementById('emp-form-color');
  if (employeeColorInput) {
    employeeColorInput.addEventListener('input', (e) => {
      const hex = document.getElementById('emp-form-color-hex');
      if (hex) hex.textContent = e.target.value;
    });
  }

  // Zonas
  safeAddListener('form-zone', 'submit', handleSaveZoneForm);
  const zoneColorInput = document.getElementById('zone-form-color');
  if (zoneColorInput) {
    zoneColorInput.addEventListener('input', (e) => {
      const hex = document.getElementById('zone-form-color-hex');
      if (hex) hex.textContent = e.target.value;
    });
  }

  document.getElementById('btn-quick-manage-zones').addEventListener('click', () => {
    const tabEmp = document.querySelector('[data-tab="tab-employees"]');
    if (tabEmp) tabEmp.click();
  });

  document.getElementById('input-search-stores').addEventListener('input', (e) => {
    renderBiggieStoresTable(e.target.value);
  });

  // Zona desde Directorio
  document.getElementById('btn-add-zone-from-stores').addEventListener('click', openCreateZoneModal);

  // Sucursales
  document.getElementById('form-store').addEventListener('submit', handleSaveStoreForm);
}
