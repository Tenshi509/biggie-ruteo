const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const { BIGGIE_STORES, STANDARD_ZONES, DEFAULT_EMPLOYEES } = require('./biggie_db');
const { planMultiEmployeeRoutes, buildEmployeeRoute, HUB_BIGGIE_DEFAULT } = require('./router_engine');

const app = express();
const PORT = process.env.PORT || 3456;

const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

const EMPLOYEES_FILE = path.join(DATA_DIR, 'employees.json');
const ZONES_FILE = path.join(DATA_DIR, 'zones.json');
const STORES_FILE = path.join(DATA_DIR, 'stores.json');

if (!fs.existsSync(EMPLOYEES_FILE)) {
  fs.writeFileSync(EMPLOYEES_FILE, JSON.stringify(DEFAULT_EMPLOYEES, null, 2), 'utf8');
}
if (!fs.existsSync(ZONES_FILE)) {
  fs.writeFileSync(ZONES_FILE, JSON.stringify(STANDARD_ZONES, null, 2), 'utf8');
}
if (!fs.existsSync(STORES_FILE)) {
  fs.writeFileSync(STORES_FILE, JSON.stringify(BIGGIE_STORES, null, 2), 'utf8');
}

function getEmployees() {
  try {
    return JSON.parse(fs.readFileSync(EMPLOYEES_FILE, 'utf8'));
  } catch (e) {
    return DEFAULT_EMPLOYEES;
  }
}

function saveEmployees(list) {
  fs.writeFileSync(EMPLOYEES_FILE, JSON.stringify(list, null, 2), 'utf8');
}

function getZones() {
  try {
    return JSON.parse(fs.readFileSync(ZONES_FILE, 'utf8'));
  } catch (e) {
    return STANDARD_ZONES;
  }
}

function saveZones(list) {
  fs.writeFileSync(ZONES_FILE, JSON.stringify(list, null, 2), 'utf8');
}

function getStores() {
  try {
    return JSON.parse(fs.readFileSync(STORES_FILE, 'utf8'));
  } catch (e) {
    return BIGGIE_STORES;
  }
}

function saveStores(list) {
  fs.writeFileSync(STORES_FILE, JSON.stringify(list, null, 2), 'utf8');
}

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({ storage: multer.memoryStorage() });

// -------------------------------------------------------------
// RUTAS API
// -------------------------------------------------------------

app.get('/api/status', (req, res) => {
  res.json({
    status: 'online',
    app: 'Biggie Ruteo Paraguay',
    version: '1.3.0',
    timestamp: new Date().toISOString()
  });
});

// EMPLEADOS
app.get('/api/employees', (req, res) => {
  res.json(getEmployees());
});

app.post('/api/employees', (req, res) => {
  const employees = req.body;
  if (!Array.isArray(employees)) {
    return res.status(400).json({ error: 'Formato inválido.' });
  }
  saveEmployees(employees);
  res.json({ success: true, employees });
});

app.post('/api/employees/new', (req, res) => {
  const empData = req.body;
  if (!empData || !empData.name) {
    return res.status(400).json({ error: 'El nombre del empleado es obligatorio.' });
  }

  const list = getEmployees();
  const newId = 'EMP-' + String(Date.now()).slice(-4);
  const newEmp = {
    id: empData.id || newId,
    name: empData.name,
    phone: empData.phone || '',
    assignedZones: Array.isArray(empData.assignedZones) ? empData.assignedZones : [],
    color: empData.color || '#3B82F6',
    active: empData.active !== false
  };

  list.push(newEmp);
  saveEmployees(list);
  res.json({ success: true, employee: newEmp, employees: list });
});

app.put('/api/employees/:id', (req, res) => {
  const empId = req.params.id;
  const updateData = req.body;
  const list = getEmployees();
  const index = list.findIndex(e => e.id === empId);

  if (index === -1) {
    return res.status(404).json({ error: 'Empleado no encontrado.' });
  }

  list[index] = {
    ...list[index],
    name: updateData.name !== undefined ? updateData.name : list[index].name,
    phone: updateData.phone !== undefined ? updateData.phone : list[index].phone,
    color: updateData.color !== undefined ? updateData.color : list[index].color,
    active: updateData.active !== undefined ? updateData.active : list[index].active,
    assignedZones: Array.isArray(updateData.assignedZones) ? updateData.assignedZones : list[index].assignedZones
  };

  saveEmployees(list);
  res.json({ success: true, employee: list[index], employees: list });
});

app.delete('/api/employees/:id', (req, res) => {
  const empId = req.params.id;
  let list = getEmployees();
  const initialLength = list.length;
  list = list.filter(e => e.id !== empId);

  if (list.length === initialLength) {
    return res.status(404).json({ error: 'Empleado no encontrado.' });
  }

  saveEmployees(list);
  res.json({ success: true, message: 'Empleado eliminado.', employees: list });
});

// ZONAS (CREAR, EDITAR, ELIMINAR)
app.get('/api/zones', (req, res) => {
  res.json(getZones());
});

app.post('/api/zones', (req, res) => {
  const zones = req.body;
  if (!Array.isArray(zones)) {
    return res.status(400).json({ error: 'Formato inválido.' });
  }
  saveZones(zones);
  res.json({ success: true, zones });
});

app.post('/api/zones/new', (req, res) => {
  const zoneData = req.body;
  if (!zoneData || !zoneData.name) {
    return res.status(400).json({ error: 'El nombre de la zona es obligatorio.' });
  }

  const list = getZones();
  const newId = 'ZONA-' + String(Date.now()).slice(-4);
  const newZone = {
    id: zoneData.id || newId,
    name: zoneData.name.trim(),
    description: zoneData.description ? zoneData.description.trim() : 'Zona personalizada de Paraguay',
    color: zoneData.color || '#F59E0B'
  };

  list.push(newZone);
  saveZones(list);
  res.json({ success: true, zone: newZone, zones: list });
});

app.put('/api/zones/:id', (req, res) => {
  const zoneId = req.params.id;
  const updateData = req.body;
  const list = getZones();
  const index = list.findIndex(z => z.id === zoneId);

  if (index === -1) {
    return res.status(404).json({ error: 'Zona no encontrada.' });
  }

  const oldName = list[index].name;
  const newName = updateData.name !== undefined ? updateData.name.trim() : oldName;

  list[index] = {
    ...list[index],
    name: newName,
    description: updateData.description !== undefined ? updateData.description.trim() : list[index].description,
    color: updateData.color !== undefined ? updateData.color : list[index].color
  };

  saveZones(list);

  // Si cambió de nombre, actualizar en los empleados asignados
  if (oldName !== newName) {
    const empList = getEmployees();
    let changed = false;
    empList.forEach(emp => {
      if (emp.assignedZones && emp.assignedZones.includes(oldName)) {
        emp.assignedZones = emp.assignedZones.map(z => z === oldName ? newName : z);
        changed = true;
      }
    });
    if (changed) saveEmployees(empList);
  }

  res.json({ success: true, zone: list[index], zones: list });
});

app.delete('/api/zones/:id', (req, res) => {
  const zoneId = req.params.id;
  let list = getZones();
  const target = list.find(z => z.id === zoneId);

  if (!target) {
    return res.status(404).json({ error: 'Zona no encontrada.' });
  }

  list = list.filter(z => z.id !== zoneId);
  saveZones(list);

  // Remover de empleados
  const empList = getEmployees();
  let changed = false;
  empList.forEach(emp => {
    if (emp.assignedZones && emp.assignedZones.includes(target.name)) {
      emp.assignedZones = emp.assignedZones.filter(z => z !== target.name);
      changed = true;
    }
  });
  if (changed) saveEmployees(empList);

  res.json({ success: true, message: 'Zona eliminada.', zones: list });
});

// SUCURSALES BIGGIE (CRUD)
app.get('/api/biggie-stores', (req, res) => {
  res.json(getStores());
});

app.post('/api/biggie-stores', (req, res) => {
  const storeData = req.body;
  if (!storeData || !storeData.name) {
    return res.status(400).json({ error: 'El nombre de la sucursal es obligatorio.' });
  }

  const list = getStores();
  const newId = 'BIG-' + String(Date.now()).slice(-6);
  const newStore = {
    id: storeData.id || newId,
    name: storeData.name.trim(),
    address: (storeData.address || '').trim(),
    zone: (storeData.zone || 'Sin zona').trim(),
    city: (storeData.city || '').trim(),
    lat: parseFloat(storeData.lat) || 0,
    lng: parseFloat(storeData.lng) || 0,
    phone: (storeData.phone || '').trim(),
    type: storeData.type || 'store'
  };

  list.push(newStore);
  saveStores(list);
  res.json({ success: true, store: newStore, stores: list });
});

app.put('/api/biggie-stores/:id', (req, res) => {
  const storeId = req.params.id;
  const updateData = req.body;
  const list = getStores();
  const index = list.findIndex(s => s.id === storeId);

  if (index === -1) {
    return res.status(404).json({ error: 'Sucursal no encontrada.' });
  }

  list[index] = {
    ...list[index],
    name: updateData.name !== undefined ? updateData.name.trim() : list[index].name,
    address: updateData.address !== undefined ? updateData.address.trim() : list[index].address,
    zone: updateData.zone !== undefined ? updateData.zone.trim() : list[index].zone,
    city: updateData.city !== undefined ? updateData.city.trim() : list[index].city,
    lat: updateData.lat !== undefined ? parseFloat(updateData.lat) : list[index].lat,
    lng: updateData.lng !== undefined ? parseFloat(updateData.lng) : list[index].lng,
    phone: updateData.phone !== undefined ? updateData.phone.trim() : list[index].phone
  };

  saveStores(list);
  res.json({ success: true, store: list[index], stores: list });
});

app.delete('/api/biggie-stores/:id', (req, res) => {
  const storeId = req.params.id;
  let list = getStores();
  const initialLength = list.length;
  list = list.filter(s => s.id !== storeId);

  if (list.length === initialLength) {
    return res.status(404).json({ error: 'Sucursal no encontrada.' });
  }

  saveStores(list);
  res.json({ success: true, message: 'Sucursal eliminada.', stores: list });
});

app.get('/api/download-template', (req, res) => {
  const templatePath = path.join(__dirname, 'sample_data', 'plantilla_rutas_biggie.xlsx');
  if (fs.existsSync(templatePath)) {
    res.download(templatePath, 'Plantilla_Rutas_Biggie_Paraguay.xlsx');
  } else {
    res.status(404).send('Plantilla no encontrada');
  }
});

function cleanStr(s) {
  if (!s) return '';
  return String(s).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function matchBiggieStore(nameQuery, addressQuery, cityQuery) {
  const qName = cleanStr(nameQuery);
  const qAddr = cleanStr(addressQuery);
  const qCity = cleanStr(cityQuery);

  const stores = getStores();

  for (const store of stores) {
    const sName = cleanStr(store.name);
    const sAddr = cleanStr(store.address);

    if (qName && (sName.includes(qName) || qName.includes(sName.replace('biggie ', '')))) {
      return store;
    }
    if (qAddr && sAddr.includes(qAddr)) {
      return store;
    }
  }

  for (const store of stores) {
    const sCity = cleanStr(store.city);
    if (qCity && sCity.includes(qCity)) {
      return store;
    }
  }

  return null;
}

app.post('/api/upload-excel', upload.single('excelFile'), (req, res) => {
  try {
    if (!req.file || !req.file.buffer) {
      return res.status(400).json({ error: 'No se recibió ningún archivo Excel.' });
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    if (!rawRows || rawRows.length === 0) {
      return res.status(400).json({ error: 'El archivo Excel está vacío.' });
    }

    const parsedStops = [];
    let idCounter = 1;

    rawRows.forEach(row => {
      const idVal = row['ID'] || row['ID Sucursal'] || row['Id'] || row['id'] || ('STOP-' + idCounter);
      const nameVal = row['Nombre'] || row['Nombre Sucursal'] || row['Sucursal'] || row['Destino'] || ('Sucursal Biggie ' + idCounter);
      const addrVal = row['Direccion'] || row['Dirección'] || row['Ubicacion'] || '';
      const cityVal = row['Ciudad'] || row['Localidad'] || 'Asunción';
      let zoneVal = row['Zona'] || row['Sector'] || '';
      let latVal = parseFloat(row['Latitud'] || row['Lat'] || row['lat'] || 0);
      let lngVal = parseFloat(row['Longitud'] || row['Long'] || row['Lng'] || row['lng'] || 0);
      const orderVal = row['Nro Pedido'] || row['Pedido'] || row['Factura'] || ('PED-' + idCounter);
      const packagesVal = parseInt(row['Cajas/Bultos'] || row['Cajas'] || row['Bultos'] || 1, 10);
      const priorityVal = row['Prioridad'] || 'Normal';
      const phoneVal = row['Telefono'] || row['Teléfono'] || row['Telefono Contacto'] || '';
      const notesVal = row['Notas'] || row['Observaciones'] || '';

      const matchedStore = matchBiggieStore(nameVal, addrVal, cityVal);
      if (matchedStore) {
        if (!latVal || !lngVal || isNaN(latVal) || isNaN(lngVal)) {
          latVal = matchedStore.lat;
          lngVal = matchedStore.lng;
        }
        if (!zoneVal) {
          zoneVal = matchedStore.zone;
        }
      }

      if (!latVal || !lngVal || isNaN(latVal) || isNaN(lngVal)) {
        latVal = -25.2950 + ((Math.random() - 0.5) * 0.08);
        lngVal = -57.5800 + ((Math.random() - 0.5) * 0.08);
      }

      if (!zoneVal) {
        zoneVal = 'Asunción Centro';
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
        notes: String(notesVal),
        matchedWithDb: !!matchedStore
      });

      idCounter++;
    });

    res.json({
      success: true,
      totalRows: parsedStops.length,
      stops: parsedStops
    });
  } catch (error) {
    console.error('Error al procesar Excel:', error);
    res.status(500).json({ error: 'Error al procesar el archivo Excel: ' + error.message });
  }
});

app.post('/api/optimize-routes', (req, res) => {
  try {
    const { stops, employees, hub, returnToHub } = req.body;

    if (!stops || !Array.isArray(stops) || stops.length === 0) {
      return res.status(400).json({ error: 'No se enviaron paradas para rutear.' });
    }

    const empList = employees && employees.length > 0 ? employees : getEmployees();
    const hubPoint = hub || HUB_BIGGIE_DEFAULT;
    const shouldReturn = returnToHub !== undefined ? returnToHub : true;

    const result = planMultiEmployeeRoutes(stops, empList, hubPoint, shouldReturn);
    res.json(result);
  } catch (error) {
    console.error('Error al optimizar rutas:', error);
    res.status(500).json({ error: 'Error al calcular ruteo: ' + error.message });
  }
});

app.post('/api/export-routes-excel', (req, res) => {
  try {
    const { routes, summary } = req.body;
    if (!routes || !Array.isArray(routes)) {
      return res.status(400).json({ error: 'Datos de ruta inválidos.' });
    }

    const workbook = XLSX.utils.book_new();

    const summaryRows = [
      { 'Métrica': 'Total de Paradas / Sucursales', 'Valor': summary.totalStops },
      { 'Métrica': 'Choferes / Empleados Activos', 'Valor': summary.activeRoutes },
      { 'Métrica': 'Distancia Total del Operativo', 'Valor': `${summary.grandTotalDistanceKm} km` },
      { 'Métrica': 'Tiempo Total Estimado', 'Valor': summary.grandTotalTimeFormatted },
      { 'Métrica': 'Costo Estimado de Combustible', 'Valor': summary.grandTotalFuelCostFormatted },
      { 'Métrica': 'Fecha de Generación', 'Valor': new Date().toLocaleString('es-PY') }
    ];
    const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
    wsSummary['!cols'] = [{ wch: 30 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(workbook, wsSummary, 'Resumen General');

    routes.forEach(r => {
      if (r.stops && r.stops.length > 0) {
        const empNameClean = r.employee.name.replace(/[^a-zA-Z0-9 ]/g, '').substring(0, 25);
        const routeRows = r.stops.map(s => ({
          'Parada #': s.stopNumber,
          'Sucursal': s.name,
          'Dirección': s.address,
          'Zona': s.zone,
          'Ciudad': s.city,
          'Nro Pedido': s.orderNumber,
          'Cajas/Bultos': s.packages,
          'Prioridad': s.priority,
          'Distancia Tramo (km)': s.distanceFromPrevKm,
          'Distancia Acumulada (km)': s.cumulativeDistanceKm,
          'Llegada Estimada (+min)': `${s.estimatedArrivalMin} min`,
          'Contacto': s.phone,
          'Observaciones': s.notes
        }));

        const wsEmp = XLSX.utils.json_to_sheet(routeRows);
        wsEmp['!cols'] = [
          { wch: 10 }, { wch: 30 }, { wch: 35 }, { wch: 22 }, { wch: 18 },
          { wch: 16 }, { wch: 14 }, { wch: 12 }, { wch: 20 }, { wch: 22 },
          { wch: 22 }, { wch: 16 }, { wch: 30 }
        ];
        XLSX.utils.book_append_sheet(workbook, wsEmp, empNameClean);
      }
    });

    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', 'attachment; filename="Rutas_Optimizadas_Biggie.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (error) {
    console.error('Error al exportar Excel:', error);
    res.status(500).json({ error: 'Error al generar Excel: ' + error.message });
  }
});

app.listen(PORT, () => {
  console.log(`====================================================`);
  console.log(` BIGGIE PARAGUAY - SISTEMA DE RUTEO INTELIGENTE LOCAL`);
  console.log(` Servidor activo en: http://localhost:${PORT}`);
  console.log(`====================================================`);
});
