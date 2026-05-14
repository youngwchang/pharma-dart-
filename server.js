const express = require('express');
const cors    = require('cors');
const path    = require('path');
const app     = express();

app.use(cors());
app.use(express.json());

// DART 프록시 라우터
app.use('/dart', require('./routes/dart'));

// React 빌드 결과물 서빙
app.use(express.static(path.join(__dirname, 'client/dist')));
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'client/dist/index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));