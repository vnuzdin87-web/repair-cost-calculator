const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

function createWindow() {
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1100,
    minHeight: 700,
    backgroundColor: '#0f1026',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    },
    title: 'Калькулятор стоимости ремонта'
  });

  win.loadFile('src/index.html');
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

ipcMain.handle('load-price', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [
      { name: 'Таблицы', extensions: ['xlsx', 'xls', 'csv'] },
      { name: 'Excel', extensions: ['xlsx', 'xls'] },
      { name: 'CSV', extensions: ['csv'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  });
  if (result.canceled || result.filePaths.length === 0) return { success: false };
  try {
    const filePath = result.filePaths[0];
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    return { success: true, data: jsonData, fileName: path.basename(filePath) };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('save-estimate', async (event, data) => {
  const result = await dialog.showSaveDialog({
    filters: [
      { name: 'Excel', extensions: ['xlsx'] },
      { name: 'JSON', extensions: ['json'] }
    ],
    defaultPath: 'Смета_расчет.xlsx'
  });
  if (result.canceled) return { success: false };
  try {
    if (result.filePath.endsWith('.json')) {
      fs.writeFileSync(result.filePath, JSON.stringify(data, null, 2), 'utf-8');
      return { success: true };
    }

    const items = Array.isArray(data.items) ? data.items : [];
    const works = items.filter(i => i.type === 'work');
    const mats = items.filter(i => i.type === 'material');
    const bath = items.filter(i => i.type === 'bath');
    const sum = (arr) => arr.reduce((s, i) => s + (Number(i.sum) || 0), 0);
    const sw = sum(works), sm = sum(mats), sb = sum(bath);
    const total = sw + sm + sb;

    const aoa = [];
    aoa.push(['СМЕТА НА КАПИТАЛЬНЫЙ РЕМОНТ КВАРТИРЫ']);
    aoa.push(['Объект', '5-комнатная квартира, 200 м²']);
    aoa.push(['Дата', new Date().toLocaleString('ru-RU')]);
    aoa.push([]);
    aoa.push(['Заказчик', (data.client && data.client.name) || '']);
    aoa.push(['Телефон', (data.client && data.client.phone) || '']);
    aoa.push(['Адрес', (data.client && data.client.address) || '']);
    aoa.push(['Сегмент отделки', data.segment || '']);
    aoa.push(['Комментарий', (data.client && data.client.comment) || '']);
    aoa.push([]);

    const addSection = (title, headers, rows) => {
      aoa.push([title]);
      aoa.push(headers);
      rows.forEach(r => aoa.push(r));
      aoa.push([]);
    };

    addSection('РАБОТЫ',
      ['№', 'Группа', 'Наименование', 'Ед. изм.', 'Объём', 'Цена за ед., руб.', 'Сумма, руб.', 'Требования к приёмке'],
      works.map((w, idx) => [idx + 1, w.group || '', w.name || '', w.unit || '', Number(w.vol) || 0, Number(w.price) || 0, Number(w.sum) || 0, w.req || ''])
    );
    addSection('МАТЕРИАЛЫ',
      ['Категория', 'Наименование', 'Сегмент', 'Бренд', 'Ед. изм.', 'Кол-во', 'Цена за ед., руб.', 'Сумма, руб.', 'Характеристика'],
      mats.map(m => [m.cat || '', m.name || '', m.seg || '', m.brand || '', m.unit || '', Number(m.vol) || 0, Number(m.price) || 0, Number(m.sum) || 0, m.desc || ''])
    );
    addSection('САНТЕХНИКА И САНУЗЛЫ',
      ['Помещение', 'Наименование', 'Сегмент', 'Бренд', 'Ед. изм.', 'Кол-во', 'Цена за ед., руб.', 'Сумма, руб.'],
      bath.map(b => [b.room || '', b.name || '', b.seg || '', b.brand || '', b.unit || '', Number(b.qty) || 0, Number(b.price) || 0, Number(b.sum) || 0])
    );

    aoa.push(['ИТОГО']);
    aoa.push(['Работы', sw]);
    aoa.push(['Материалы', sm]);
    aoa.push(['Сантехника и санузлы', sb]);
    aoa.push(['ВСЕГО К ОПЛАТЕ', total]);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [
      { wch: 22 }, { wch: 34 }, { wch: 14 }, { wch: 28 }, { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 18 }, { wch: 36 }
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Смета');
    XLSX.writeFile(wb, result.filePath);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});
