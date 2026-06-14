// === CONFIG ===
const CLIENT_ID = '113470943982-eujk6po7pepq3kbhvl8obfs4754vkjts.apps.googleusercontent.com';
const SPREADSHEET_ID = '1A2xbLM38T74E6o_XjgSn4MiWp3ITIeB99hHVYVhuu18';
const SHEET_RANGES = ['RCP_cash!A:T', 'RCP_in_kind!A:V', 'RCP_infrastructures!A:K'];

// Oblast centroids (lat, lng) for map
const OBLAST_COORDS = {
   'Запорізька': [47.8388, 35.1396],
   'Харківська': [49.9935, 36.2304],
   'Харків': [49.9935, 36.2304],
   'Херсонська': [46.6354, 32.6169],
   'Миколаївська': [46.9750, 31.9946],
   'Дніпропетровська': [48.4647, 35.0462],
   'Донецька': [48.0159, 37.8028],
   'Одеська': [46.4774, 30.7326],
   'Полтавська': [49.5883, 34.5514],
   'Чернігівська': [51.4982, 31.2893],
   'Хмельницька': [49.4229, 26.9871],
   'Житомирська': [50.2547, 28.6587],
   'Житомирска': [50.2547, 28.6587],
   'Волинська': [50.7472, 25.3254],
};

let tokenClient, accessToken = null;
let masterData = [], filteredData = [];
let charts = { oblasts: null, demographics: null, timeline: null, fuel: null, assistanceType: null, partners: null, popGroup: null, age: null, monthly: null };
let leafletMap = null, mapMarkers = [];

const colors = {
   red: '#711324', lightRed: '#a83246', gray: '#6c757d', dark: '#343a40',
   wood: '#8B4513', briq: '#CD853F', coal: '#2F4F4F',
   palette: ['#711324', '#a83246', '#c45c73', '#d4839a', '#e8b4bf', '#6c757d', '#495057', '#343a40', '#8B4513', '#CD853F', '#2F4F4F', '#5c7a2e', '#2874a6', '#1a5276']
};

// === GOOGLE AUTH ===
window.onload = function () {
   Chart.defaults.font.family = "'Segoe UI', sans-serif";
   Chart.defaults.color = '#555';
   try {
      tokenClient = google.accounts.oauth2.initTokenClient({
         client_id: CLIENT_ID,
         scope: 'https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/userinfo.email',
         callback: handleAuthResponse,
      });
   } catch (e) { console.error("GSI failed:", e); }
   document.getElementById('auth-btn').onclick = () => tokenClient.requestAccessToken({ prompt: 'consent' });
};

async function handleAuthResponse(response) {
   if (response.error) { showError("Помилка авторизації. Спробуйте ще раз."); return; }
   accessToken = response.access_token;
   try {
      const userInfoRes = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', { headers: { Authorization: `Bearer ${accessToken}` } });
      const userInfo = await userInfoRes.json();
      const email = userInfo.email.toLowerCase();
      const allowedDomains = ['@caritas.ua', '@cu.caritas.ua', '@db.caritas.ua'];
      if (!allowedDomains.some(d => email.endsWith(d))) { showError(`Пошта ${email} не має доступу.`); return; }

      document.getElementById('user-email').innerText = email;
      document.getElementById('user-initials').innerText = email.substring(0, 2).toUpperCase();
      document.getElementById('auth-screen').style.display = 'none';
      document.getElementById('dashboard-screen').style.display = 'flex';
      document.getElementById('loader').style.display = 'flex';

      await fetchSheetsData();
   } catch (err) { showError("Помилка зв'язку з Google."); console.error(err); }
}

function showError(msg) {
   document.getElementById('error-text').innerText = msg;
   document.getElementById('auth-error').style.display = 'block';
}

// === FETCH DATA ===
async function fetchSheetsData() {
   try {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values:batchGet?ranges=${SHEET_RANGES.join('&ranges=')}`;
      const response = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
      if (!response.ok) throw new Error("Немає доступу до таблиці.");
      const data = await response.json();
      processRawData(data.valueRanges);
   } catch (err) {
      console.error(err);
      alert("Помилка завантаження даних. Перевірте консоль.");
      document.getElementById('loader').style.display = 'none';
   }
}

// === PROCESS DATA ===
function processRawData(valueRanges) {
   masterData = [];

   // RCP_cash
   if (valueRanges[0]?.values) {
      valueRanges[0].values.slice(1).forEach(r => {
         if (!r[0]) return;
         masterData.push({
            typeGroup: 'cash',
            assistanceType: r[18] || 'Unknown',
            project: r[1] || 'Unknown',
            partner: r[2] || 'Unknown',
            oblast: normalizeOblast(r[3] || 'Unknown'),
            raion: r[4] || '',
            popGroup: normalizePopGroup(r[9] || 'Unknown'),
            hhs: parseInt(r[7]) || 0,
            people: parseInt(r[8]) || 0,
            w017: parseInt(r[10]) || 0, m017: parseInt(r[11]) || 0,
            w1859: parseInt(r[12]) || 0, m1859: parseInt(r[13]) || 0,
            w60: parseInt(r[14]) || 0, m60: parseInt(r[15]) || 0,
            disabled: parseInt(r[16]) || 0,
            date: parseDate(r[17]),
            amount: parseNum(r[19]),
            fuelTotal: 0, wood: 0, briq: 0, coal: 0
         });
         console.log("CASH:", r);
      });
   }

   // RCP_in_kind
   if (valueRanges[1]?.values) {
      valueRanges[1].values.slice(1).forEach(r => {
         if (!r[0]) return;
         const w = parseNum(r[19]), b = parseNum(r[20]), c = parseNum(r[21]);
         masterData.push({
            typeGroup: 'inkind',
            assistanceType: r[18] || 'Паливо для обігріву',
            project: r[1] || 'Unknown',
            partner: r[2] || 'Unknown',
            oblast: normalizeOblast(r[3] || 'Unknown'),
            raion: r[4] || '',
            popGroup: normalizePopGroup(r[9] || 'Unknown'),
            hhs: parseInt(r[7]) || 0,
            people: parseInt(r[8]) || 0,
            w017: parseInt(r[10]) || 0, m017: parseInt(r[11]) || 0,
            w1859: parseInt(r[12]) || 0, m1859: parseInt(r[13]) || 0,
            w60: parseInt(r[14]) || 0, m60: parseInt(r[15]) || 0,
            disabled: parseInt(r[16]) || 0,
            date: parseDate(r[17]),
            amount: 0,
            fuelTotal: w + b + c, wood: w, briq: b, coal: c
         });
         console.log("IN KIND:", r);
      });
   }

   // RCP_infrastructures
   if (valueRanges[2]?.values) {
      valueRanges[2].values.slice(1).forEach(r => {
         if (!r[0]) return;
         masterData.push({
            typeGroup: 'infra',
            assistanceType: r[9] || 'Інфраструктура',
            project: r[1] || 'Unknown',
            partner: r[2] || 'Unknown',
            oblast: normalizeOblast(r[3] || 'Unknown'),
            raion: r[4] || '',
            popGroup: 'Unknown',
            hhs: 0,
            people: parseInt(r[7]) || 0,
            w017: 0, m017: 0, w1859: 0, m1859: 0, w60: 0, m60: 0,
            disabled: 0,
            date: parseDate(r[8]),
            amount: parseNum(r[10]),
            fuelTotal: 0, wood: 0, briq: 0, coal: 0
         });
         console.log("INFRA:", r);
      });
   }

   filteredData = [...masterData];
   populateFilters();
   initMap();
   updateDashboard();

   setTimeout(() => {
      if (leafletMap) leafletMap.invalidateSize();
   }, 300);
   document.getElementById('loader').style.display = 'none';
}

function normalizeOblast(s) {
   s = s.trim();
   if (s === 'Харків') return 'Харківська';
   if (s === 'Житомирска') return 'Житомирська';
   return s;
}

function normalizePopGroup(s) {
   s = s.trim();
   if (s.includes('Внутрішньо переміщені') || s === 'ВПО' || s.startsWith('ВПО')) return 'ВПО';
   if (s.includes('фронт')) return 'Прифронтові мешканці';
   if (s.includes('Місцев')) return 'Місцеве населення';
   if (s.includes('повернул')) return 'Репатріанти';
   return s;
}

function parseNum(v) {
   if (v === undefined || v === null || v === '') return 0;

   return parseFloat(
      String(v)
         .replace(/\s/g, '')
         .replace(',', '.')
         .replace(/[^\d.-]/g, '')
   ) || 0;
}

function parseDate(s) {
   if (!s) return null;
   if (s instanceof Date) return s;
   const parts = String(s).split('.');
   if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
   // ISO format
   const d = new Date(s);
   return isNaN(d) ? null : d;
}

// === FILTERS ===
function populateFilters() {
   const projects = new Set(), oblasts = new Set(), years = new Set(), partners = new Set(), popGroups = new Set(), assistanceTypes = new Set();
   masterData.forEach(i => {
      projects.add(i.project); oblasts.add(i.oblast);
      partners.add(i.partner); popGroups.add(i.popGroup);
      assistanceTypes.add(i.assistanceType);
      if (i.date) years.add(i.date.getFullYear());
   });
   const fill = (id, set, sort = true) => {
      const sel = document.getElementById(id);
      const first = sel.options[0].outerHTML;
      sel.innerHTML = first;
      const arr = sort ? Array.from(set).sort() : Array.from(set);
      arr.forEach(v => { if (v) sel.innerHTML += `<option value="${v}">${v}</option>`; });
   };
   fill('flt-project', projects);
   fill('flt-oblast', oblasts);
   fill('flt-year', years);
   fill('flt-partner', partners);
   fill('flt-popgroup', popGroups);
   fill('flt-assistance-type', assistanceTypes);
}

function applyFilters() {
   const p = document.getElementById('flt-project').value;
   const t = document.getElementById('flt-type').value;
   const at = document.getElementById('flt-assistance-type').value;
   const o = document.getElementById('flt-oblast').value;
   const pa = document.getElementById('flt-partner').value;
   const pg = document.getElementById('flt-popgroup').value;
   const y = document.getElementById('flt-year').value;

   filteredData = masterData.filter(i => {
      if (p !== 'all' && i.project !== p) return false;
      if (t !== 'all' && i.typeGroup !== t) return false;
      if (at !== 'all' && i.assistanceType !== at) return false;
      if (o !== 'all' && i.oblast !== o) return false;
      if (pa !== 'all' && i.partner !== pa) return false;
      if (pg !== 'all' && i.popGroup !== pg) return false;
      if (y !== 'all' && (!i.date || i.date.getFullYear().toString() !== y)) return false;
      return true;
   });
   updateDashboard();
}

function resetFilters() {
   ['flt-project', 'flt-type', 'flt-assistance-type', 'flt-oblast', 'flt-partner', 'flt-popgroup', 'flt-year']
      .forEach(id => document.getElementById(id).value = 'all');
   filteredData = [...masterData];
   updateDashboard();
}

// === DASHBOARD UPDATE ===
function updateDashboard() {
   let tHhs = 0, tPeople = 0, tAmount = 0, tFuel = 0, tDisabled = 0;

   let obMap = {};

   let demoMap = {
      women: 0,
      men: 0
   };

   let ageMap = {
      w017: 0,
      m017: 0,
      w1859: 0,
      m1859: 0,
      w60: 0,
      m60: 0
   };

   let timeMap = {};

   let fuelMap = {
      wood: 0,
      briq: 0,
      coal: 0
   };

   let assistTypeMap = {};
   let partnerMap = {};
   let popGroupMap = {};
   let monthMap = {};

   filteredData.forEach(i => {
      tHhs += i.hhs; tPeople += i.people; tAmount += i.amount; tFuel += i.fuelTotal; tDisabled += i.disabled;
      obMap[i.oblast] = (obMap[i.oblast] || 0) + i.people;
      demoMap.women += i.w017 + i.w1859 + i.w60;
      demoMap.men += i.m017 + i.m1859 + i.m60;
      ageMap.w017 += i.w017; ageMap.m017 += i.m017;
      ageMap.w1859 += i.w1859; ageMap.m1859 += i.m1859;
      ageMap.w60 += i.w60; ageMap.m60 += i.m60;

      if (i.date) {
         const yr = i.date.getFullYear();
         timeMap[yr] = (timeMap[yr] || 0) + i.people;
         const mo = `${yr}-${String(i.date.getMonth() + 1).padStart(2, '0')}`;
         monthMap[mo] = (monthMap[mo] || 0) + i.people;
      }

      fuelMap.wood += i.wood; fuelMap.briq += i.briq; fuelMap.coal += i.coal;

      // Shorten assistance type labels
      const atLabel = shortenAssistType(i.assistanceType);
      assistTypeMap[atLabel] = (assistTypeMap[atLabel] || 0) + i.people;
      if (i.popGroup !== 'Unknown') popGroupMap[i.popGroup] = (popGroupMap[i.popGroup] || 0) + i.people;
   });

   // KPIs
   const kpiCards = document.querySelectorAll('.kpi-card');

   if (kpiCards[0]) kpiCards[0].style.display = tHhs ? 'flex' : 'none';
   if (kpiCards[1]) kpiCards[1].style.display = tPeople ? 'flex' : 'none';
   if (kpiCards[2]) kpiCards[2].style.display = tAmount ? 'flex' : 'none';
   if (kpiCards[3]) kpiCards[3].style.display = tFuel ? 'flex' : 'none';
   if (kpiCards[4]) kpiCards[4].style.display = tDisabled ? 'flex' : 'none';

   if (tHhs) animateValue('kpi-hh', tHhs);
   if (tPeople) animateValue('kpi-people', tPeople);
   if (tAmount) animateValue('kpi-amount', tAmount, true);
   if (tFuel) animateValue('kpi-fuel', Math.round(tFuel));
   if (tDisabled) animateValue('kpi-disabled', tDisabled);

   // Charts
   updateChartOblasts(obMap);
   updateChartDemographics(demoMap);
   updateChartTimeline(timeMap);
   updateChartFuel(fuelMap);
   updateChartAssistanceType(assistTypeMap);
   updateChartPartners(partnerMap);
   updateChartPopGroup(popGroupMap);
   updateChartAge(ageMap);
   updateChartMonthly(monthMap);
   updateMapMarkers(obMap);
   updateOblastTable(obMap);
}

function shortenAssistType(s) {
   if (!s) return 'Інше';
   if (s.includes('складних ремонтів')) return 'Складний ремонт';
   if (s.includes('середніх ремонтів')) return 'Середній ремонт';
   if (s.includes('легких ремонтів')) return 'Легкий ремонт';
   if (s.includes('комунальн')) return 'Комунальні послуги';
   if (s.includes('паливо') && s.includes('натуральн')) return 'Паливо (In-Kind)';
   if (s.includes('паливо') && s.includes('готівк')) return 'Паливо (Готівка)';
   if (s.includes('оренд')) return 'Оренда житла';
   if (s.includes('ВПО') || s.includes('вПО')) return 'Ремонт ВПО';
   if (s.includes('інфраструктур') || s.includes('ремонтів у житлових')) return 'Інфраструктура';
   return s.length > 30 ? s.substring(0, 30) + '…' : s;
}

// === CHARTS ===
function destroyAndCreate(key, canvasId, config) {
   const canvas = document.getElementById(canvasId);
   if (!canvas) return;

   const hasData =
      config?.data?.datasets?.some(ds =>
         ds.data && ds.data.some(v => v > 0)
      );

   if (!hasData) {
      if (charts[key]) {
         charts[key].destroy();
         charts[key] = null;
      }
      return;
   }

   if (charts[key]) charts[key].destroy();

   const ctx = canvas.getContext('2d');
   charts[key] = new Chart(ctx, config);
}

function updateChartOblasts(dataMap) {
   const sorted = Object.entries(dataMap).sort((a, b) => b[1] - a[1]).slice(0, 12);
   destroyAndCreate('oblasts', 'chartOblasts', {
      type: 'bar',
      data: { labels: sorted.map(e => e[0]), datasets: [{ label: 'Осіб', data: sorted.map(e => e[1]), backgroundColor: colors.red, borderRadius: 5 }] },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { font: { size: 11 } } } } }
   });
}

function updateChartDemographics(demoMap) {
   destroyAndCreate('demographics', 'chartDemographics', {
      type: 'doughnut',
      data: {
         labels: ['Жінки', 'Чоловіки'],
         datasets: [{
            data: [
               demoMap.women,
               demoMap.men
            ],
            backgroundColor: [
               colors.red,
               colors.gray
            ],
            borderWidth: 2,
            borderColor: '#fff'
         }]
      },
      options: {
         responsive: true,
         maintainAspectRatio: false,
         cutout: '68%',
         layout: {
            padding: 10
         },
         plugins: {
            legend: {
               position: 'right', // ЗМІНЕНО: легенда праворуч
               labels: {
                  boxWidth: 12,
                  padding: 14,
                  font: {
                     size: 11
                  }
               }
            }
         }
      }
   });
}

function updateChartTimeline(timeMap) {
   const labels = Object.keys(timeMap).sort();
   destroyAndCreate('timeline', 'chartTimeline', {
      type: 'line',
      data: { labels, datasets: [{ label: 'Осіб', data: labels.map(l => timeMap[l]), borderColor: colors.red, backgroundColor: 'rgba(113,19,36,0.1)', fill: true, tension: 0.3, pointBackgroundColor: colors.red, pointRadius: 5 }] },
      options: { responsive: true, maintainAspectRatio: false }
   });
}

function updateChartFuel(fuelMap) {
   destroyAndCreate('fuel', 'chartFuel', {
      type: 'bar',
      data: { labels: ['Дрова', 'Брикети', 'Вугілля'], datasets: [{ label: 'Тонни', data: [fuelMap.wood, fuelMap.briq, fuelMap.coal], backgroundColor: [colors.wood, colors.briq, colors.coal], borderRadius: 5 }] },
      options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', plugins: { legend: { display: false } } }
   });
}

function updateChartAssistanceType(dataMap) {
   const sorted = Object.entries(dataMap)
      .sort((a, b) => b[1] - a[1]);

   destroyAndCreate('assistanceType', 'chartAssistanceType', {
      type: 'doughnut',
      data: {
         labels: sorted.map(e => e[0]),
         datasets: [{
            data: sorted.map(e => e[1]),
            backgroundColor: colors.palette.slice(0, sorted.length),
            borderWidth: 2,
            borderColor: '#fff'
         }]
      },
      options: {
         responsive: true,
         maintainAspectRatio: false,
         cutout: '68%',
         layout: {
            padding: 10
         },
         plugins: {
            legend: {
               position: 'right', // ЗМІНЕНО: легенда праворуч
               labels: {
                  boxWidth: 12,
                  padding: 12,
                  font: {
                     size: 11
                  }
               }
            }
         }
      }
   });
}

function updateChartPartners() {
   const oblastTypeMap = {};

   filteredData.forEach(i => {
      if (!oblastTypeMap[i.oblast]) {
         oblastTypeMap[i.oblast] = {};
      }

      const type = shortenAssistType(i.assistanceType);

      oblastTypeMap[i.oblast][type] =
         (oblastTypeMap[i.oblast][type] || 0) + i.hhs;
   });

   const oblasts = Object.keys(oblastTypeMap);

   const allTypes = [...new Set(
      filteredData.map(i => shortenAssistType(i.assistanceType))
   )];

   const datasets = allTypes.map((type, idx) => ({
      label: type,
      data: oblasts.map(ob => oblastTypeMap[ob][type] || 0),
      backgroundColor: colors.palette[idx % colors.palette.length],
      borderRadius: 4
   }));

   destroyAndCreate('partners', 'chartPartners', {
      type: 'bar',
      data: {
         labels: oblasts,
         datasets
      },
      options: {
         responsive: true,
         maintainAspectRatio: false,
         plugins: {
            legend: {
               position: 'bottom'
            }
         },
         scales: {
            x: {
               stacked: true
            },
            y: {
               stacked: true
            }
         }
      }
   });
}

function updateChartPopGroup(dataMap) {
   const sorted = Object.entries(dataMap)
      .sort((a, b) => b[1] - a[1]);

   destroyAndCreate('popGroup', 'chartPopGroup', {
      type: 'doughnut',
      data: {
         labels: sorted.map(e => e[0]),
         datasets: [{
            data: sorted.map(e => e[1]),
            backgroundColor: [
               '#711324',
               '#a83246',
               '#c45c73',
               '#6c757d',
               '#2874a6',
               '#2e86c1'
            ],
            borderWidth: 2,
            borderColor: '#fff'
         }]
      },
      options: {
         responsive: true,
         maintainAspectRatio: false,
         cutout: '68%',
         layout: {
            padding: 10
         },
         plugins: {
            legend: {
               position: 'right', // ЗМІНЕНО: легенда праворуч
               labels: {
                  boxWidth: 12,
                  padding: 12,
                  font: {
                     size: 11
                  }
               }
            }
         }
      }
   });
}

function updateChartAge(a) {
   destroyAndCreate('age', 'chartAge', {
      type: 'bar',
      data: {
         labels: ['0–17', '18–59', '60+'],
         datasets: [
            { label: 'Жінки', data: [a.w017, a.w1859, a.w60], backgroundColor: colors.red, borderRadius: 4 },
            { label: 'Чоловіки', data: [a.m017, a.m1859, a.m60], backgroundColor: colors.gray, borderRadius: 4 }
         ]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
   });
}

function updateChartMonthly(monthMap) {
   const labels = Object.keys(monthMap).sort();
   destroyAndCreate('monthly', 'chartMonthly', {
      type: 'line',
      data: { labels, datasets: [{ label: 'Осіб/місяць', data: labels.map(l => monthMap[l]), borderColor: colors.red, backgroundColor: 'rgba(113,19,36,0.08)', fill: true, tension: 0.4, pointRadius: 3, pointBackgroundColor: colors.red }] },
      options: { responsive: true, maintainAspectRatio: false, scales: { x: { ticks: { maxTicksLimit: 18, font: { size: 10 } } } } }
   });
}

// === MAP ===
function initMap() {
   if (leafletMap) return;

   const ukraineBounds = [
      [44.2, 22.0],
      [52.5, 40.5]
   ];

   leafletMap = L.map('map', {
      zoomControl: true,
      scrollWheelZoom: false,
      maxBounds: ukraineBounds,
      maxBoundsViscosity: 1.0,
      minZoom: 6,
      maxZoom: 9
   }).setView([48.8, 31.5], 6);

   L.tileLayer(
      'https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png',
      {
         attribution: '© OpenStreetMap, © CARTO'
      }
   ).addTo(leafletMap);
}

function updateMapMarkers(obMap) {
   mapMarkers.forEach(m => leafletMap.removeLayer(m));
   mapMarkers = [];

   if (!leafletMap) return;

   const vals = Object.values(obMap);
   if (!vals.length) return;

   const maxVal = Math.max(...vals);

   Object.entries(obMap).forEach(([oblast, count]) => {
      const coords = OBLAST_COORDS[oblast];
      if (!coords) return;

      const oblastData = filteredData.filter(i => i.oblast === oblast);

      let women = 0;
      let men = 0;
      let types = {};

      oblastData.forEach(i => {
         women += i.w017 + i.w1859 + i.w60;
         men += i.m017 + i.m1859 + i.m60;

         const type = shortenAssistType(i.assistanceType);
         types[type] = (types[type] || 0) + i.people;
      });

      const typeList = Object.entries(types)
         .map(([t, c]) => `• ${t}: ${numFormat.format(c)}`)
         .join('<br>');

      const r = Math.max(15, Math.min(55, (count / maxVal) * 55));

      const marker = L.circleMarker(coords, {
         radius: r,
         fillColor: colors.red,
         fillOpacity: 0.65,
         color: '#fff',
         weight: 2
      }).bindTooltip(`
         <div style="min-width:220px;">
            <strong>${oblast}</strong><br><br>

            <strong>Всього:</strong> ${numFormat.format(count)} осіб<br>
            <strong>Жінки:</strong> ${numFormat.format(women)}<br>
            <strong>Чоловіки:</strong> ${numFormat.format(men)}<br><br>

            <strong>Види допомоги:</strong><br>
            ${typeList}
         </div>
      `, {
         direction: 'top',
         sticky: true
      });

      marker.addTo(leafletMap);
      mapMarkers.push(marker);
   });
}

// === MINI TABLE ===
const numFormat = new Intl.NumberFormat('uk-UA');
function updateOblastTable(obMap) {
   const sorted = Object.entries(obMap).sort((a, b) => b[1] - a[1]); const maxPeople = sorted[0]?.[1] || 1;
   // Also get amount per oblast from filteredData
   const amountMap = {};
   filteredData.forEach(i => { amountMap[i.oblast] = (amountMap[i.oblast] || 0) + i.amount; });

   const tbody = document.getElementById('oblast-table-body');
   if (!sorted.length) {
      tbody.innerHTML = `<tr><td colspan="4">Немає даних</td></tr>`;
      return;
   }

   tbody.innerHTML = sorted.map(([ob, cnt], idx) => {
      const barW = Math.round((cnt / maxPeople) * 80);
      const amt = amountMap[ob] || 0;
      return `<tr>
         <td>${idx + 1}</td>
         <td>
            <div class="rank-bar">
               <span>${ob}</span>
            </div>
            <div class="rank-bar"><div class="rank-bar-fill" style="width:${barW}px"></div></div>
         </td>
         <td class="num">${numFormat.format(cnt)}</td>
         <td class="num">${amt > 0 ? (amt / 1000000).toFixed(1) + 'M' : '—'}</td>
      </tr>`;
   }).join('');
}

// === ANIMATE KPI ===
function animateValue(id, end, isCurrency = false) {
   const obj = document.getElementById(id);
   let current = 0;
   const steps = 25, stepVal = end / steps;
   clearInterval(obj._timer);
   obj._timer = setInterval(() => {
      current = Math.min(current + stepVal, end);
      obj.innerText = isCurrency
         ? current.toLocaleString('uk-UA', { maximumFractionDigits: 0 })
         : numFormat.format(Math.round(current));
      if (current >= end) clearInterval(obj._timer);
   }, 28);
}