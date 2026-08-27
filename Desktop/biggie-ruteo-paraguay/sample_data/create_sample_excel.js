const fs = require('fs');
const XLSX = require('xlsx');
const path = require('path');

const sampleRows = [
  {
'ID Sucursal': 'BIG-VM-01',
    'Nombre Sucursal': 'Biggie Villa Morra',
    'Direccion': 'Avda. Mariscal López c/ Senador Long',
    'Ciudad': 'Asunción',
    'Zona': 'Villa Morra',
    'Latitud': -25.2971,
    'Longitud': -57.5866,
    'Nro Pedido': 'PED-2026-081',
    'Cajas/Bultos': 15,
    'Prioridad': 'Alta',
    'Telefono Contacto': '0981 200001',
    'Notas': 'Entregar por portón lateral'
  },
  {
'ID Sucursal': 'BIG-ASU-01',
    'Nombre Sucursal': 'Biggie Palma',
    'Direccion': 'Calle Palma e/ 14 de Mayo y 15 de Agosto',
    'Ciudad': 'Asunción',
    'Zona': 'Palma Centro',
    'Latitud': -25.2833,
    'Longitud': -57.6331,
    'Nro Pedido': 'PED-2026-082',
    'Cajas/Bultos': 12,
    'Prioridad': 'Normal',
    'Telefono Contacto': '0981 100001',
    'Notas': 'Recepción abierta 24hs'
  },
  {
'ID Sucursal': 'BIG-LUQ-01',
    'Nombre Sucursal': 'Biggie Luque Centro',
    'Direccion': 'Gral. Aquino c/ Rosario',
    'Ciudad': 'Luque',
    'Zona': 'Luque Centro',
    'Latitud': -25.2668,
    'Longitud': -57.4962,
    'Nro Pedido': 'PED-2026-086',
    'Cajas/Bultos': 14,
    'Prioridad': 'Normal',
    'Telefono Contacto': '0981 300001',
    'Notas': 'Frente a plaza principal'
  },
  {
'ID Sucursal': 'BIG-MRA-01',
    'Nombre Sucursal': 'Biggie Mariano Roque Alonso',
    'Direccion': 'Ruta Transchaco c/ Bernardino Caballero',
    'Ciudad': 'Mariano Roque Alonso',
    'Zona': 'Mariano Roque Alonso Transchaco',
    'Latitud': -25.2289,
    'Longitud': -57.5342,
    'Nro Pedido': 'PED-2026-088',
    'Cajas/Bultos': 22,
    'Prioridad': 'Alta',
    'Telefono Contacto': '0981 300003',
    'Notas': 'Descarga en bahía de camiones'
  },
  {
'ID Sucursal': 'BIG-SL-01',
    'Nombre Sucursal': 'Biggie San Lorenzo Centro',
    'Direccion': 'Julia Miranda Cueto c/ Sgto. Silva',
    'Ciudad': 'San Lorenzo',
    'Zona': 'San Lorenzo Centro',
    'Latitud': -25.3212,
    'Longitud': -57.4974,
    'Nro Pedido': 'PED-2026-089',
    'Cajas/Bultos': 16,
    'Prioridad': 'Normal',
    'Telefono Contacto': '0981 400001',
    'Notas': 'Zona comercial centro'
  },
  {
'ID Sucursal': 'BIG-CAP-01',
    'Nombre Sucursal': 'Biggie Capiatá Km 16',
    'Direccion': 'Ruta PY02 Km 16.5 c/ Aratiri',
    'Ciudad': 'Capiatá',
    'Zona': 'Capiatá Km 16 Aratiri',
    'Latitud': -25.3585,
    'Longitud': -57.4372,
    'Nro Pedido': 'PED-2026-094',
    'Cajas/Bultos': 18,
    'Prioridad': 'Normal',
    'Telefono Contacto': '0981 400004',
    'Notas': 'Entrada por colectora'
  },
  {
'ID Sucursal': 'BIG-FDM-01',
    'Nombre Sucursal': 'Biggie Fernando Norte',
    'Direccion': 'Avda. Mcal. López c/ Insaurralde',
    'Ciudad': 'Fernando de la Mora',
    'Zona': 'Fernando de la Mora Norte',
    'Latitud': -25.3074,
    'Longitud': -57.5332,
    'Nro Pedido': 'PED-2026-091',
    'Cajas/Bultos': 11,
    'Prioridad': 'Normal',
    'Telefono Contacto': '0981 500001',
    'Notas': 'Lado norte'
  },
  {
'ID Sucursal': 'BIG-LAM-01',
    'Nombre Sucursal': 'Biggie Lambaré',
    'Direccion': 'Avda. Cacique Lambaré c/ Río Apa',
    'Ciudad': 'Lambaré',
    'Zona': 'Lambaré Cacique',
    'Latitud': -25.3389,
    'Longitud': -57.6184,
    'Nro Pedido': 'PED-2026-092',
    'Cajas/Bultos': 15,
    'Prioridad': 'Normal',
    'Telefono Contacto': '0981 600001',
    'Notas': 'Estacionamiento amplio'
  },
  {
'ID Sucursal': 'BIG-VEL-01',
    'Nombre Sucursal': 'Biggie Villa Elisa',
    'Direccion': 'Avda. Von Poleski c/ Colombia',
    'Ciudad': 'Villa Elisa',
    'Zona': 'Villa Elisa Centro',
    'Latitud': -25.3697,
    'Longitud': -57.5839,
    'Nro Pedido': 'PED-2026-095',
    'Cajas/Bultos': 13,
    'Prioridad': 'Normal',
    'Telefono Contacto': '0981 600003',
    'Notas': 'Sobre avenida principal'
  },
  {
'ID Sucursal': 'BIG-NEM-01',
    'Nombre Sucursal': 'Biggie Ñemby',
    'Direccion': 'Avda. Acceso Sur c/ Pratt Gill',
    'Ciudad': 'Ñemby',
    'Zona': 'Ñemby Acceso Sur',
    'Latitud': -25.4014,
    'Longitud': -57.5753,
    'Nro Pedido': 'PED-2026-096',
    'Cajas/Bultos': 20,
    'Prioridad': 'Alta',
    'Telefono Contacto': '0981 600004',
    'Notas': 'Descarga en patio trasero'
  }
];

const worksheet = XLSX.utils.json_to_sheet(sampleRows);
worksheet['!cols'] = [
  { wch: 14 }, { wch: 32 }, { wch: 42 }, { wch: 22 }, { wch: 24 },
  { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 14 }, { wch: 12 },
  { wch: 18 }, { wch: 35 }
];

const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Rutas Biggie Paraguay');

const targetPath = path.join(__dirname, 'plantilla_rutas_biggie.xlsx');
XLSX.writeFile(workbook, targetPath);
console.log('Sample Excel re-generated at:', targetPath);
