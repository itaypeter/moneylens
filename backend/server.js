require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const { initDB } = require('./db');
const dataRoutes = require('./routes/data');
const ocrRoutes  = require('./routes/ocr');

const app  = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

// ── API routes ────────────────────────────────────────────────────────────────
app.get('/health', (req, res) => res.json({ ok: true, ts: new Date().toISOString() }));
app.use('/api', dataRoutes);
app.use('/api/ocr', ocrRoutes);

// ── Serve React frontend (built into ../frontend/dist) ────────────────────────
const DIST = path.join(__dirname, '..', 'frontend', 'dist');
app.use(express.static(DIST));
app.get('*', (req, res) => res.sendFile(path.join(DIST, 'index.html')));

async function start() {
  try {
    await initDB();
    app.listen(PORT, () => console.log(`✅ MoneyLens running on port ${PORT}`));
  } catch (e) {
    console.error('Failed to start:', e.message);
    process.exit(1);
  }
}

start();
