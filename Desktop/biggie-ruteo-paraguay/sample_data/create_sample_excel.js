const fs = require('fs');
const XLSX = require('xlsx');
const path = require('path');

const sampleRows = [
  { 'Empleado': 'Carlos Gómez', 'Sucursal': 'Biggie Villa Morra', 'Notas': 'Entregar por portón lateral' },
  { 'Empleado': 'Carlos Gómez', 'Sucursal': 'Biggie Palma', 'Notas': 'Recepción abierta 24hs' },
  { 'Empleado': 'María López', 'Sucursal': 'Biggie Luque Centro', 'Notas': 'Frente a plaza principal' },
  { 'Empleado': 'María López', 'Sucursal': 'Biggie Mariano Roque Alonso', 'Notas': 'Descarga en bahía de camiones' },
  { 'Empleado': 'Juan Pérez', 'Sucursal': 'Biggie San Lorenzo Centro', 'Notas': 'Zona comercial centro' },
  { 'Empleado': 'Juan Pérez', 'Sucursal': 'Biggie Capiatá Km 16', 'Notas': 'Entrada por colectora' },
  { 'Empleado': 'Ana Martínez', 'Sucursal': 'Biggie Fernando Norte', 'Notas': 'Lado norte' },
  { 'Empleado': 'Ana Martínez', 'Sucursal': 'Biggie Lambaré', 'Notas': 'Estacionamiento amplio' },
  { 'Empleado': 'Pedro Ramírez', 'Sucursal': 'Biggie Villa Elisa', 'Notas': 'Sobre avenida principal' },
  { 'Empleado': 'Pedro Ramírez', 'Sucursal': 'Biggie Ñemby', 'Notas': 'Descarga en patio trasero' }
];

const worksheet = XLSX.utils.json_to_sheet(sampleRows);
worksheet['!cols'] = [
  { wch: 22 }, { wch: 34 }, { wch: 40 }
];

const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Rutas Biggie Paraguay');

const targetPath = path.join(__dirname, 'plantilla_rutas_biggie.xlsx');
XLSX.writeFile(workbook, targetPath);
console.log('Sample Excel re-generated at:', targetPath);
