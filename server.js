const express = require('express');
const cors    = require('cors');
const path    = require('path');
const app     = express();

app.use(cors());
app.use(express.json());

// DART 프록시
app.use('/dart', require('./routes/dart'));

// ✅ 정적 파일 — assets 폴더 명시적으로 추가
const DIST = path.join(__dirname, 'client', 'dist');
app.use(express.static(DIST));

// ✅ SPA fallback — /dart 제외하고만 적용
app.get(/^(?!\/dart).*$/, (req, res) => {
  res.sendFile(path.join(DIST, 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Listening on ${PORT}`));
