const express = require('express');
const fs = require('fs');
const ExcelJS = require('exceljs');
const cors = require('cors');
const app = express();
app.use(express.json());
app.use(cors());

app.use(express.static('public'));

function toGprmcCoords(decimal, isLat) {
  const absolute = Math.abs(decimal);
  const degrees = Math.floor(absolute);
  const minutes = (absolute - degrees) * 60;
  const formatted = isLat
    ? `${degrees.toString().padStart(2, '0')}${minutes.toFixed(2).padStart(5, '0')}`
    : `${degrees.toString().padStart(3, '0')}${minutes.toFixed(2).padStart(5, '0')}`;
  const direction = isLat ? (decimal >= 0 ? 'N' : 'S') : (decimal >= 0 ? 'E' : 'W');
  return { value: formatted, direction };
}

function calculateChecksum(sentence) {
  let checksum = 0;
  for (let i = 0; i < sentence.length; i++) {
    checksum ^= sentence.charCodeAt(i);
  }
  return checksum.toString(16).toUpperCase().padStart(2, '0');
}

function interpolateCoords(start, end, steps) {
  const points = [];
  for (let i = 0; i < steps; i++) {
    const lat = start[0] + ((end[0] - start[0]) * i) / (steps - 1);
    const lon = start[1] + ((end[1] - start[1]) * i) / (steps - 1);
    points.push([lat, lon]);
  }
  return points;
}

function generateGprmc(device, baseTime, index, lat, lon, speed = 60.0, heading = 0.0) {
  const timestamp = new Date(baseTime.getTime() + index * 10000); // 10 sec interval
  
  const hh = String(timestamp.getHours()).padStart(2, '0');
  const mm = String(timestamp.getMinutes()).padStart(2, '0');
  const ss = String(timestamp.getSeconds()).padStart(2, '0');
  const hhmmss = `${hh}${mm}${ss}`;

  const dd = String(timestamp.getDate()).padStart(2, '0');
  const MM = String(timestamp.getMonth() + 1).padStart(2, '0');
  const yy = String(timestamp.getFullYear()).slice(-2);
  const ddmmyy = `${dd}${MM}${yy}`;

  const { value: latVal, direction: latDir } = toGprmcCoords(lat, true);
  const { value: lonVal, direction: lonDir } = toGprmcCoords(lon, false);

  const sentence = `GPRMC,${hhmmss},A,${latVal},${latDir},${lonVal},${lonDir},${speed.toFixed(1)},${heading.toFixed(1)},${ddmmyy},,`;
  const checksum = calculateChecksum(sentence);
  return `${device},$${sentence}*${checksum}`;
}

app.post('/generate', async (req, res) => {
  const { from, to, count, device } = req.body;
  const baseTime = new Date();
  const coords = interpolateCoords(from, to, count);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('GPRMC Data');
  sheet.columns = [{key: 'gprmc' }];

  coords.forEach((point, index) => {
    const heading = 10 + Math.sin(index / 150) * 5;
    const gprmc = generateGprmc(device, baseTime, index, point[0], point[1], 60.0, heading);
    sheet.addRow({ gprmc });
  });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=gprmc_data_${device}.xlsx`);
  await workbook.xlsx.write(res);
  res.end();
});

app.listen(3000, () => console.log('🚀 Server running on http://localhost:3000'));
