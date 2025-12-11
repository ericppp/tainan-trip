const XLSX = require('xlsx');
const fs = require('fs');

// 讀取 data/itinerary.json
const itineraryData = JSON.parse(fs.readFileSync('./data/itinerary.json', 'utf8'));

// 準備 Excel 資料
const rows = [
  ['日期', '時間', '項目', 'Emoji', '說明', '類型', '緯度', '經度', '連結', '標籤']
];

itineraryData.itinerary.forEach(day => {
  day.items.forEach(item => {
    rows.push([
      day.date,
      item.time || '',
      item.title,
      item.emoji || '',
      item.description || '',
      item.type || '',
      item.location ? item.location.lat : '',
      item.location ? item.location.lng : '',
      item.links ? item.links.map(l => l.url).join(', ') : '',
      item.tags ? item.tags.join(', ') : ''
    ]);
  });
});

// 建立 workbook 和 worksheet
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet(rows);

// 設定欄寬
ws['!cols'] = [
  { wch: 12 }, // 日期
  { wch: 8 },  // 時間
  { wch: 25 }, // 項目
  { wch: 6 },  // Emoji
  { wch: 40 }, // 說明
  { wch: 10 }, // 類型
  { wch: 12 }, // 緯度
  { wch: 12 }, // 經度
  { wch: 50 }, // 連結
  { wch: 20 }  // 標籤
];

XLSX.utils.book_append_sheet(wb, ws, '台南旅遊行程');

// 輸出檔案
XLSX.writeFile(wb, '台南旅遊行程.xlsx');

console.log('✅ Excel 檔案已建立: 台南旅遊行程.xlsx');
console.log(`   包含 ${rows.length - 1} 筆行程資料`);
