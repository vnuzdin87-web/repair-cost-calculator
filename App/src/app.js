const state = {
  priceData: [],
  works: [],
  materials: [],
  bath: [],
  selectedSegment: 'Эконом',
  estimate: {
    works: new Set(),
    materials: new Set(),
    bath: new Set(),
    workQty: {},
    matQty: {}
  }
};

const fmt = (n) => new Intl.NumberFormat('ru-RU').format(Math.round(n || 0)) + ' ₽';

function initTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
      if (btn.dataset.tab === 'estimate') buildEstimate();
    });
  });
}

function initSegments() {
  document.querySelectorAll('.seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.selectedSegment = btn.dataset.segment;
      document.getElementById('matSegmentHint').textContent = 'Показаны материалы и сантехника для сегмента: ' + state.selectedSegment;
      document.getElementById('estSeg').textContent = state.selectedSegment;
      state.estimate.materials.clear();
      state.materials.filter(m => m.seg === state.selectedSegment).forEach(m => state.estimate.materials.add(m.id));
      state.estimate.bath.clear();
      state.bath.filter(b => b.seg === state.selectedSegment).forEach(b => state.estimate.bath.add(b.id));
      filterMaterials();
      renderBath();
      document.querySelectorAll('#materialsTable input[type=checkbox][data-type="m"]').forEach(cb => { cb.checked = state.estimate.materials.has(cb.dataset.id); });
      document.querySelectorAll('#bathTable input[type=checkbox][data-type="b"]').forEach(cb => { cb.checked = state.estimate.bath.has(cb.dataset.id); });
      recalc();
    });
  });
}

function initFilters() {
  document.querySelectorAll('.filter-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      document.querySelectorAll('.filter-chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      const f = chip.dataset.filter;
      document.querySelectorAll('#worksTable tbody tr').forEach(tr => {
        if (f === 'all') { tr.style.display = ''; return; }
        tr.style.display = tr.dataset.group === f ? '' : 'none';
      });
    });
  });
}

const YANDEX_PUBLIC_URL = '';

function applyLoadedData(data, sourceName) {
  state.priceData = data;
  parseData(data);
  state.estimate.materials.clear();
  state.materials.filter(m => m.seg === state.selectedSegment).forEach(m => state.estimate.materials.add(m.id));
  state.estimate.bath.clear();
  state.bath.filter(b => b.seg === state.selectedSegment).forEach(b => state.estimate.bath.add(b.id));
  document.getElementById('fileStatus').textContent = 'Загружено: ' + sourceName + ' (работ: ' + state.works.length + ', мат+санузлы: ' + (state.materials.length + state.bath.length) + ')';
  renderAll();
  document.querySelectorAll('#materialsTable input[type=checkbox][data-type="m"]').forEach(cb => { cb.checked = state.estimate.materials.has(cb.dataset.id); });
  document.querySelectorAll('#bathTable input[type=checkbox][data-type="b"]').forEach(cb => { cb.checked = state.estimate.bath.has(cb.dataset.id); });
  recalc();
}

async function loadXlsxFromArrayBuffer(ab, sourceName) {
  if (typeof XLSX === 'undefined') throw new Error('XLSX не загружен');
  const wb = XLSX.read(ab);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  applyLoadedData(json, sourceName);
}

async function fetchArrayBuffer(url) {
  const res = await fetch(url);
  if (!res.ok) { const e = new Error('HTTP ' + res.status); e.status = res.status; throw e; }
  return await res.arrayBuffer();
}

async function loadFromYandexDisk() {
  const status = document.getElementById('fileStatus');
  if (!YANDEX_PUBLIC_URL) return;
  status.textContent = 'Загрузка данных с Яндекс.Диска...';
  const apiUrl = 'https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key=' + encodeURIComponent(YANDEX_PUBLIC_URL);

  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      const metaRes = await fetch(apiUrl);
      if (metaRes.status === 429 && attempt < 2) {
        status.textContent = 'Яндекс.Диск ограничил запросы, повтор через 2 сек...';
        await new Promise(r => setTimeout(r, 2000));
        continue;
      }
      if (!metaRes.ok) throw new Error('API Яндекс.Диска: ' + metaRes.status);
      const meta = await metaRes.json();
      if (!meta.href) throw new Error('Не получена ссылка для скачивания');
      const data = await fetchArrayBuffer(meta.href);
      await loadXlsxFromArrayBuffer(data, 'Прайс_шаблон.xlsx (Яндекс.Диск)');
      return;
    } catch (e) {
      console.error('Yandex Disk load error (attempt ' + attempt + '):', e);
      if (attempt >= 2) {
        try {
          status.textContent = 'Яндекс.Диск недоступен, загружаю локальный прайс...';
          const localData = await fetchArrayBuffer('./price_template.xlsx');
          await loadXlsxFromArrayBuffer(localData, 'price_template.xlsx (локальный)');
          return;
        } catch (e2) {
          console.error('Local template load error:', e2);
          status.textContent = 'Не удалось загрузить прайс. Используются демо-данные. Можно выбрать файл вручную.';
          loadDefaults();
          renderAll();
          state.estimate.materials.clear();
          state.materials.filter(m => m.seg === state.selectedSegment).forEach(m => state.estimate.materials.add(m.id));
          state.estimate.bath.clear();
          state.bath.filter(b => b.seg === state.selectedSegment).forEach(b => state.estimate.bath.add(b.id));
          renderAll();
          document.querySelectorAll('#materialsTable input[type=checkbox][data-type="m"]').forEach(cb => { cb.checked = state.estimate.materials.has(cb.dataset.id); });
          document.querySelectorAll('#bathTable input[type=checkbox][data-type="b"]').forEach(cb => { cb.checked = state.estimate.bath.has(cb.dataset.id); });
          recalc();
        }
      }
    }
  }
}

document.getElementById('btnLoad').addEventListener('click', async () => {
  const res = await window.api.loadPrice();
  if (!res || !res.success) return;
  applyLoadedData(res.data, res.fileName);
});

const btnReloadYandex = document.getElementById('btnReloadYandex');
if (btnReloadYandex) {
  btnReloadYandex.hidden = !YANDEX_PUBLIC_URL;
  btnReloadYandex.addEventListener('click', loadFromYandexDisk);
}

document.getElementById('btnSave').addEventListener('click', async () => {
  const data = {
    client: {
      name: document.getElementById('clientName').value,
      phone: document.getElementById('clientPhone').value,
      address: document.getElementById('clientAddress').value,
      comment: document.getElementById('clientComment').value
    },
    segment: state.selectedSegment,
    items: buildEstimateItems()
  };
  const res = await window.api.saveEstimate(data);
  const status = document.getElementById('fileStatus');
  if (res && res.success) {
    status.textContent = 'Смета сохранена в Excel';
  } else if (res && res.error) {
    status.textContent = 'Ошибка сохранения: ' + res.error;
  }
});

function newOrder() {
  document.getElementById('clientName').value = '';
  document.getElementById('clientPhone').value = '';
  document.getElementById('clientAddress').value = '';
  document.getElementById('clientComment').value = '';
  document.getElementById('estClient').textContent = '-';
  document.getElementById('estPhone').textContent = '-';
  document.getElementById('estAddr').textContent = '-';

  state.estimate.works.clear();
  state.estimate.materials.clear();
  state.estimate.bath.clear();
  state.estimate.workQty = {};
  state.estimate.matQty = {};

  state.selectedSegment = 'Эконом';
  document.querySelectorAll('.seg-btn').forEach(b => b.classList.toggle('active', b.dataset.segment === 'Эконом'));
  document.getElementById('matSegmentHint').textContent = 'Показаны материалы и сантехника для сегмента: Эконом';
  document.getElementById('estSeg').textContent = 'Эконом';

  // Auto-select default segment materials/bath
  state.materials.filter(m => m.seg === state.selectedSegment).forEach(m => state.estimate.materials.add(m.id));
  state.bath.filter(b => b.seg === state.selectedSegment).forEach(b => state.estimate.bath.add(b.id));

  renderAll();
  document.querySelectorAll('#materialsTable input[type=checkbox][data-type="m"]').forEach(cb => { cb.checked = state.estimate.materials.has(cb.dataset.id); });
  document.querySelectorAll('#bathTable input[type=checkbox][data-type="b"]').forEach(cb => { cb.checked = state.estimate.bath.has(cb.dataset.id); });
  recalc();
  document.getElementById('fileStatus').textContent = 'Новый заказ: данные клиента очищены';
}

document.getElementById('btnRecalc').addEventListener('click', recalc);
const btnNew = document.getElementById('btnNewOrder');
if (btnNew) btnNew.addEventListener('click', newOrder);
document.getElementById('clientName').addEventListener('input', () => document.getElementById('estClient').textContent = document.getElementById('clientName').value || '-');
document.getElementById('clientPhone').addEventListener('input', () => document.getElementById('estPhone').textContent = document.getElementById('clientPhone').value || '-');
document.getElementById('clientAddress').addEventListener('input', () => document.getElementById('estAddr').textContent = document.getElementById('clientAddress').value || '-');

function parseData(data) {
  state.works = [];
  state.materials = [];
  state.bath = [];
  let inMat = false, inBath = false, inWorks = false;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length === 0) continue;
    const s = (row[0]||'').toString();
    if (s.includes('РАЗДЕЛ 1')) { inWorks = true; inMat = false; inBath = false; continue; }
    if (s.includes('РАЗДЕЛ 2')) { inMat = true; inBath = false; inWorks = false; continue; }
    if (s.includes('РАЗДЕЛ 3')) { inMat = false; inBath = true; inWorks = false; continue; }
    if (s.includes('РАЗДЕЛ 4')) { inMat = false; inBath = false; inWorks = false; continue; }
    if (s === '№' && inWorks) continue;
    if (inWorks && row[0] && (typeof row[0] === 'number' || !isNaN(Number(row[0]))) && row[1] && row[2]) {
      state.works.push({ id: 'w'+row[0], num: row[0], group: row[1], name: row[2], unit: row[3], vol: Number(row[4])||0, price: Number(row[5])||0, req: row[7]||'' });
      continue;
    }
    if (inMat && row[0] && s !== 'Категория' && row[1]) {
      state.materials.push({ id: 'm'+i, cat: row[0], name: row[1], seg: row[2], brand: row[3], unit: row[4], vol: Number(row[5])||1, price: Number(row[6])||0, desc: row[8]||'' });
      continue;
    }
    if (inBath && row[0] && s !== 'Помещение' && row[1]) {
      state.bath.push({ id: 'b'+i, room: row[0], name: row[1], seg: row[2], brand: row[3], unit: row[4], qty: Number(row[5])||1, price: Number(row[6])||0 });
      continue;
    }
  }
  if (state.works.length === 0) loadDefaults();
}

function loadDefaults() {
  state.works = [
    {id:'w1',num:1,group:'Демонтаж',name:'Демонтаж напольных покрытий',unit:'м²',vol:200,price:110,req:'Полностью без остатков'},
    {id:'w2',num:2,group:'Демонтаж',name:'Демонтаж покрытий стен',unit:'м²',vol:380,price:95,req:'Без повреждения несущих'},
    {id:'w3',num:3,group:'Черновые',name:'Штукатурка стен под маяки',unit:'м²',vol:380,price:850,req:'Откл. не более 1мм/м'},
    {id:'w4',num:4,group:'Черновые',name:'Стяжка пола',unit:'м²',vol:200,price:850,req:'По 2м правилу'},
    {id:'w5',num:5,group:'Чистовые',name:'Шпаклёвка под покраску Q4',unit:'м²',vol:290,price:650,req:'Световая проверка'},
    {id:'w6',num:6,group:'Чистовые',name:'Керамогранит/плитка',unit:'м²',vol:65,price:1650,req:'Ровные швы'},
    {id:'w7',num:7,group:'Инженерия',name:'Электрика по проекту',unit:'групп',vol:42,price:1850,req:'Однолинейка'},
    {id:'w8',num:8,group:'Инженерия',name:'Сантехника с коллектором',unit:'точек',vol:18,price:2850,req:'Опрессовка'},
  ];
  state.materials = [
    {id:'m1',cat:'Стены',name:'Краска матовая',seg:'Эконом',brand:'Tikkurila',unit:'ведро 5л',vol:16,price:3250,desc:'Моющаяся'},
    {id:'m2',cat:'Стены',name:'Латекс премиум',seg:'Средний',brand:'Dulux Diamond',unit:'ведро 5л',vol:12,price:6850,desc:'Износостойкая'},
    {id:'m3',cat:'Пол',name:'Ламинат 33 класс',seg:'Эконом',brand:'Tarkett',unit:'м²',vol:162,price:1650,desc:'Практичный'},
    {id:'m4',cat:'Пол',name:'Инженерная доска',seg:'Средний',brand:'Boen',unit:'м²',vol:135,price:4850,desc:'Шпон'},
    {id:'m5',cat:'Пол',name:'Массив дуба',seg:'Люкс',brand:'Listone Giordano',unit:'м²',vol:162,price:16500,desc:'Премиум'},
    {id:'m6',cat:'Сантехника',name:'Премиум комплект',seg:'Средний',brand:'Villeroy&Boch/Hansgrohe',unit:'компл.',vol:2,price:165000,desc:'Качественный'},
    {id:'m7',cat:'Сантехника',name:'Люкс комплект',seg:'Люкс',brand:'Gessi/Dornbracht',unit:'компл.',vol:2,price:485000,desc:'Дизайнерский'},
  ];
  state.bath = [
    {id:'b1',room:'Главный санузел',name:'Инсталляция+унитаз',seg:'Средний',brand:'Geberit/Villeroy&Boch',unit:'компл.',qty:1,price:95000},
    {id:'b2',room:'Главный санузел',name:'Душевая система',seg:'Средний',brand:'Hansgrohe',unit:'компл.',qty:1,price:85000},
    {id:'b3',room:'Главный санузел',name:'Душевая панель люкс',seg:'Люкс',brand:'Dornbracht',unit:'компл.',qty:1,price:485000},
    {id:'b4',room:'Гостевой санузел',name:'Базовый комплект',seg:'Эконом',brand:'Santek',unit:'компл.',qty:1,price:24500},
  ];
}

function renderAll() {
  renderWorks();
  renderMaterials();
  renderBath();
}

function renderWorks() {
  const tbody = document.querySelector('#worksTable tbody');
  tbody.innerHTML = '';
  state.works.forEach(w => {
    const tr = document.createElement('tr');
    tr.dataset.group = w.group;
    tr.innerHTML = `<td><input type="checkbox" data-type="w" data-id="${w.id}" ${state.estimate.works.has(w.id)?'checked':''}></td>
    <td>${w.num}</td><td>${w.group}</td><td>${w.name}</td><td>${w.unit}</td>
    <td><input type="number" value="${w.vol}" min="0" step="0.1" data-type="wvol" data-id="${w.id}"></td>
    <td><input type="number" value="${w.price}" min="0" step="1" data-type="wprice" data-id="${w.id}"></td>
    <td class="num">${fmt(w.vol*w.price)}</td><td style="color:var(--text-sub)">${w.req}</td>`;
    tbody.appendChild(tr);
  });
  bindTable();
}

function renderMaterials() {
  const tbody = document.querySelector('#materialsTable tbody');
  tbody.innerHTML = '';
  state.materials.filter(m => m.seg === state.selectedSegment).forEach(m => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td><input type="checkbox" data-type="m" data-id="${m.id}" ${state.estimate.materials.has(m.id)?'checked':''}></td>
    <td>${m.cat}</td><td>${m.name}</td><td>${m.seg}</td><td>${m.brand}</td><td>${m.unit}</td>
    <td><input type="number" value="${m.vol}" min="0" step="0.1" data-type="mvol" data-id="${m.id}"></td>
    <td><input type="number" value="${m.price}" min="0" step="1" data-type="mprice" data-id="${m.id}"></td>
    <td class="num">${fmt(m.vol*m.price)}</td><td style="color:var(--text-sub)">${m.desc}</td>`;
    tbody.appendChild(tr);
  });
  bindTable();
}

function filterMaterials() { renderMaterials(); }

function renderBath() {
  const tbody = document.querySelector('#bathTable tbody');
  tbody.innerHTML = '';
  state.bath.filter(b => b.seg === state.selectedSegment).forEach(b => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td><input type="checkbox" data-type="b" data-id="${b.id}" ${state.estimate.bath.has(b.id)?'checked':''}></td>
    <td>${b.room}</td><td>${b.name}</td><td>${b.seg}</td><td>${b.brand}</td><td>${b.unit}</td>
    <td><input type="number" value="${b.qty}" min="0" step="1" data-type="bqty" data-id="${b.id}"></td>
    <td><input type="number" value="${b.price}" min="0" step="1" data-type="bprice" data-id="${b.id}"></td>
    <td class="num">${fmt(b.qty*b.price)}</td>`;
    tbody.appendChild(tr);
  });
  bindTable();
}

function bindTable() {
  document.querySelectorAll('input[type=checkbox][data-type]').forEach(cb => {
    cb.onchange = () => {
      const type = cb.dataset.type, id = cb.dataset.id;
      if (type === 'w') { cb.checked ? state.estimate.works.add(id) : state.estimate.works.delete(id); }
      if (type === 'm') { cb.checked ? state.estimate.materials.add(id) : state.estimate.materials.delete(id); }
      if (type === 'b') { cb.checked ? state.estimate.bath.add(id) : state.estimate.bath.delete(id); }
      recalc();
    };
  });
  document.querySelectorAll('input[type=number][data-type]').forEach(inp => {
    inp.oninput = () => {
      const type = inp.dataset.type, id = inp.dataset.id, val = Number(inp.value)||0;
      if (type === 'wvol') { const w = state.works.find(x=>x.id===id); if(w){w.vol=val; renderWorks(); recalc();}}
      if (type === 'wprice') { const w = state.works.find(x=>x.id===id); if(w){w.price=val; renderWorks(); recalc();}}
      if (type === 'mvol') { const m = state.materials.find(x=>x.id===id); if(m){m.vol=val; renderMaterials(); recalc();}}
      if (type === 'mprice') { const m = state.materials.find(x=>x.id===id); if(m){m.price=val; renderMaterials(); recalc();}}
      if (type === 'bqty') { const b = state.bath.find(x=>x.id===id); if(b){b.qty=val; renderBath(); recalc();}}
      if (type === 'bprice') { const b = state.bath.find(x=>x.id===id); if(b){b.price=val; renderBath(); recalc();}}
    };
  });
}

function recalc(showStatus = false) {
  let sw=0, sm=0, sb=0;
  state.works.forEach(w=>{ if(state.estimate.works.has(w.id)) sw += w.vol*w.price; });
  state.materials.forEach(m=>{ if(state.estimate.materials.has(m.id)) sm += m.vol*m.price; });
  state.bath.forEach(b=>{ if(state.estimate.bath.has(b.id)) sb += b.qty*b.price; });
  document.getElementById('sumWorks').textContent = fmt(sw);
  document.getElementById('sumMaterials').textContent = fmt(sm);
  document.getElementById('sumBath').textContent = fmt(sb);
  document.getElementById('sumTotal').textContent = fmt(sw+sm+sb);
  const estTab = document.getElementById('tab-estimate');
  if (estTab && estTab.classList.contains('active')) buildEstimate();
  if (showStatus) {
    const status = document.getElementById('fileStatus');
    if (status) status.textContent = 'Расчёт обновлён • ' + new Date().toLocaleTimeString('ru-RU');
  }
}

function buildEstimateItems() {
  const items = [];
  state.works.forEach(w=>{ if(state.estimate.works.has(w.id)) items.push({type:'work',...w,sum:w.vol*w.price}); });
  state.materials.forEach(m=>{ if(state.estimate.materials.has(m.id)) items.push({type:'material',...m,sum:m.vol*m.price}); });
  state.bath.forEach(b=>{ if(state.estimate.bath.has(b.id)) items.push({type:'bath',...b,sum:b.qty*b.price}); });
  return items;
}

function buildEstimate() {
  document.getElementById('estClient').textContent = document.getElementById('clientName').value || '-';
  document.getElementById('estPhone').textContent = document.getElementById('clientPhone').value || '-';
  document.getElementById('estAddr').textContent = document.getElementById('clientAddress').value || '-';
  document.getElementById('estSeg').textContent = state.selectedSegment;
  let sw=0, sm=0, sb=0;
  const ew = document.getElementById('estWorks'); ew.innerHTML='<tr><th>№</th><th>Наименование</th><th>Ед.</th><th>Объём</th><th>Цена</th><th>Сумма</th></tr>';
  state.works.forEach(w=>{ if(state.estimate.works.has(w.id)){ sw+=w.vol*w.price; ew.innerHTML+=`<tr><td>${w.num}</td><td>${w.name}</td><td>${w.unit}</td><td>${w.vol}</td><td>${fmt(w.price)}</td><td>${fmt(w.vol*w.price)}</td></tr>`;}});
  const em = document.getElementById('estMat'); em.innerHTML='<tr><th>Категория</th><th>Наименование</th><th>Бренд</th><th>Ед.</th><th>Кол-во</th><th>Цена</th><th>Сумма</th></tr>';
  state.materials.forEach(m=>{ if(state.estimate.materials.has(m.id)){ sm+=m.vol*m.price; em.innerHTML+=`<tr><td>${m.cat}</td><td>${m.name}</td><td>${m.brand}</td><td>${m.unit}</td><td>${m.vol}</td><td>${fmt(m.price)}</td><td>${fmt(m.vol*m.price)}</td></tr>`;}});
  const eb = document.getElementById('estBath'); eb.innerHTML='<tr><th>Помещение</th><th>Наименование</th><th>Бренд</th><th>Ед.</th><th>Кол-во</th><th>Цена</th><th>Сумма</th></tr>';
  state.bath.forEach(b=>{ if(state.estimate.bath.has(b.id)){ sb+=b.qty*b.price; eb.innerHTML+=`<tr><td>${b.room}</td><td>${b.name}</td><td>${b.brand}</td><td>${b.unit}</td><td>${b.qty}</td><td>${fmt(b.price)}</td><td>${fmt(b.qty*b.price)}</td></tr>`;}});
  document.getElementById('estSumW').textContent = fmt(sw).replace(' ₽','');
  document.getElementById('estSumM').textContent = fmt(sm).replace(' ₽','');
  document.getElementById('estSumB').textContent = fmt(sb).replace(' ₽','');
  document.getElementById('estSumG').textContent = fmt(sw+sm+sb).replace(' ₽','');
}

loadDefaults();
renderAll();
initTabs();
initSegments();
initFilters();
document.getElementById('matSegmentHint').textContent = 'Показаны материалы и сантехника для сегмента: Эконом';
document.getElementById('estClient').textContent='-'; document.getElementById('estPhone').textContent='-'; document.getElementById('estAddr').textContent='-';
document.getElementById('estSeg').textContent = 'Эконом';
loadFromYandexDisk();
