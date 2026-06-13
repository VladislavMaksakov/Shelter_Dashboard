// === КОНФІГУРАЦІЯ ===
const CLIENT_ID = '113470943982-eujk6po7pepq3kbhvl8obfs4754vkjts.apps.googleusercontent.com';
const SPREADSHEET_ID = '1A2xbLM38T74E6o_XjgSn4MiWp3ITIeB99hHVYVhuu18';

// Аркуші, з яких будемо тягнути дані
const SHEET_RANGES = ['RCP_cash!A:T', 'RCP_in_kind!A:V', 'RCP_infrastructures!A:K'];

// Глобальні змінні
let tokenClient;
let accessToken = null;
let masterData = [];
let filteredData = [];

// Інстанси графіків
let charts = {
   oblasts: null,
   demographics: null,
   timeline: null,
   fuel: null
};

// Кольори для графіків
const colors = {
   red: '#711324',
   lightRed: '#a83246',
   gray: '#6c757d',
   dark: '#343a40',
   wood: '#8B4513',
   briq: '#CD853F',
   coal: '#2F4F4F'
};

// === ІНІЦІАЛІЗУЄМО GOOGLE AUTH ===
window.onload = function () {
   Chart.defaults.font.family = "'Segoe UI', sans-serif";
   Chart.defaults.color = '#555';

   try {
      tokenClient = google.accounts.oauth2.initTokenClient({
         client_id: CLIENT_ID,
         scope: 'https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/userinfo.email',
         callback: handleAuthResponse,
      });
   } catch (e) {
      console.error("GSI library failed to load. Are you offline?", e);
   }

   document.getElementById('auth-btn').onclick = () => {
      tokenClient.requestAccessToken({ prompt: 'consent' });
   };
};

// Обробка відповіді авторизації
async function handleAuthResponse(response) {
   if (response.error !== undefined) {
      showError("Помилка авторизації. Спробуйте ще раз.");
      return;
   }
   accessToken = response.access_token;

   try {
      const userInfoRes = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
         headers: { Authorization: `Bearer ${accessToken}` }
      });
      const userInfo = await userInfoRes.json();
      const email = userInfo.email.toLowerCase();

      // ПЕРЕВІРКА ДОМЕНУ (Оновлено)
      const allowedDomains = ['@caritas.ua', '@cu.caritas.ua', '@db.caritas.ua'];
      const isAllowed = allowedDomains.some(domain => email.endsWith(domain));

      if (!isAllowed) {
         showError(`Пошта ${email} не має доступу.`);
         return;
      }

      // Авторизація успішна
      document.getElementById('user-email').innerText = email;
      document.getElementById('user-initials').innerText = email.substring(0, 2).toUpperCase();

      document.getElementById('auth-screen').style.display = 'none';
      document.getElementById('dashboard-screen').style.display = 'flex';
      document.getElementById('loader').style.display = 'flex';

      // Завантажуємо дані
      await fetchSheetsData();

   } catch (err) {
      showError("Помилка зв'язку з сервером Google.");
      console.error(err);
   }
}

function showError(msg) {
   const errDiv = document.getElementById('auth-error');
   document.getElementById('error-text').innerText = msg;
   errDiv.style.display = 'block';
}

// === ЗАВАНТАЖЕННЯ ДАНИХ З GOOGLE SHEETS ===
async function fetchSheetsData() {
   try {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values:batchGet?ranges=${SHEET_RANGES.join('&ranges=')}`;

      const response = await fetch(url, {
         headers: { Authorization: `Bearer ${accessToken}` }
      });

      if (!response.ok) throw new Error("Немає доступу до таблиці. Перевірте SPREADSHEET_ID.");

      const data = await response.json();
      processRawData(data.valueRanges);

   } catch (err) {
      console.error(err);
      alert("Помилка завантаження даних. Перевірте консоль.");
      document.getElementById('loader').style.display = 'none';
   }
}

// === ОБРОБКА ДАНИХ ===
function processRawData(valueRanges) {
   masterData = [];

   // 1. RCP_cash
   if (valueRanges[0] && valueRanges[0].values) {
      const rows = valueRanges[0].values.slice(1);
      rows.forEach(r => {
         if (!r[0]) return;
         masterData.push({
            typeAssistanceGroup: 'cash',
            project: r[1] || 'Unknown',
            oblast: r[3] || 'Unknown',
            hhs: parseInt(r[7]) || 0,
            people: parseInt(r[8]) || 0,
            women: (parseInt(r[10]) || 0) + (parseInt(r[12]) || 0) + (parseInt(r[14]) || 0),
            men: (parseInt(r[11]) || 0) + (parseInt(r[13]) || 0) + (parseInt(r[15]) || 0),
            disabled: parseInt(r[16]) || 0,
            date: parseDate(r[17]),
            amount: parseFloat((r[19] || "0").replace(/\s/g, '').replace(',', '.')) || 0,
            fuelTotal: 0, wood: 0, briq: 0, coal: 0
         });
      });
   }

   // 2. RCP_in_kind
   if (valueRanges[1] && valueRanges[1].values) {
      const rows = valueRanges[1].values.slice(1);
      rows.forEach(r => {
         if (!r[0]) return;
         const w = parseFloat((r[19] || "0").replace(',', '.')) || 0;
         const b = parseFloat((r[20] || "0").replace(',', '.')) || 0;
         const c = parseFloat((r[21] || "0").replace(',', '.')) || 0;

         masterData.push({
            typeAssistanceGroup: 'inkind',
            project: r[1] || 'Unknown',
            oblast: r[3] || 'Unknown',
            hhs: parseInt(r[7]) || 0,
            people: parseInt(r[8]) || 0,
            women: (parseInt(r[10]) || 0) + (parseInt(r[12]) || 0) + (parseInt(r[14]) || 0),
            men: (parseInt(r[11]) || 0) + (parseInt(r[13]) || 0) + (parseInt(r[15]) || 0),
            disabled: parseInt(r[16]) || 0,
            date: parseDate(r[17]),
            amount: 0,
            fuelTotal: w + b + c, wood: w, briq: b, coal: c
         });
      });
   }

   // 3. RCP_infrastructures
   if (valueRanges[2] && valueRanges[2].values) {
      const rows = valueRanges[2].values.slice(1);
      rows.forEach(r => {
         if (!r[0]) return;
         masterData.push({
            typeAssistanceGroup: 'infra',
            project: r[1] || 'Unknown',
            oblast: r[3] || 'Unknown',
            hhs: 0,
            people: parseInt(r[7]) || 0,
            women: 0, men: 0, disabled: 0,
            date: parseDate(r[8]),
            amount: parseFloat((r[10] || "0").replace(/\s/g, '').replace(',', '.')) || 0,
            fuelTotal: 0, wood: 0, briq: 0, coal: 0
         });
      });
   }

   filteredData = [...masterData];
   populateFilters();
   updateDashboard();
   document.getElementById('loader').style.display = 'none';
}

function parseDate(dateStr) {
   if (!dateStr) return null;
   const parts = dateStr.split('.');
   if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
   return null;
}

// === ВІДМАЛЬОВКА ТА ОНОВЛЕННЯ ===
function populateFilters() {
   const projects = new Set();
   const oblasts = new Set();
   const years = new Set();

   masterData.forEach(item => {
      projects.add(item.project);
      oblasts.add(item.oblast);
      if (item.date) years.add(item.date.getFullYear());
   });

   const fillSelect = (id, set) => {
      const select = document.getElementById(id);
      select.innerHTML = select.options[0].outerHTML;
      Array.from(set).sort().forEach(val => {
         select.innerHTML += `<option value="${val}">${val}</option>`;
      });
   };

   fillSelect('flt-project', projects);
   fillSelect('flt-oblast', oblasts);
   fillSelect('flt-year', years);
}

function applyFilters() {
   const p = document.getElementById('flt-project').value;
   const t = document.getElementById('flt-type').value;
   const o = document.getElementById('flt-oblast').value;
   const y = document.getElementById('flt-year').value;

   filteredData = masterData.filter(item => {
      let match = true;
      if (p !== 'all' && item.project !== p) match = false;
      if (t !== 'all' && item.typeAssistanceGroup !== t) match = false;
      if (o !== 'all' && item.oblast !== o) match = false;
      if (y !== 'all' && (!item.date || item.date.getFullYear().toString() !== y)) match = false;
      return match;
   });

   updateDashboard();
}

function resetFilters() {
   document.getElementById('flt-project').value = 'all';
   document.getElementById('flt-type').value = 'all';
   document.getElementById('flt-oblast').value = 'all';
   document.getElementById('flt-year').value = 'all';
   filteredData = [...masterData];
   updateDashboard();
}

function updateDashboard() {
   let tHhs = 0, tPeople = 0, tAmount = 0, tFuel = 0;
   let obMap = {}, demoMap = { women: 0, men: 0, disabled: 0 };
   let timeMap = {};
   let fuelMap = { wood: 0, briq: 0, coal: 0 };

   filteredData.forEach(i => {
      tHhs += i.hhs;
      tPeople += i.people;
      tAmount += i.amount;
      tFuel += i.fuelTotal;

      obMap[i.oblast] = (obMap[i.oblast] || 0) + i.people;
      demoMap.women += i.women;
      demoMap.men += i.men;
      demoMap.disabled += i.disabled;

      if (i.date) {
         const yr = i.date.getFullYear();
         timeMap[yr] = (timeMap[yr] || 0) + i.people;
      }

      fuelMap.wood += i.wood;
      fuelMap.briq += i.briq;
      fuelMap.coal += i.coal;
   });

   animateValue("kpi-hh", tHhs);
   animateValue("kpi-people", tPeople);
   animateValue("kpi-amount", tAmount, true);
   animateValue("kpi-fuel", Math.round(tFuel));

   updateChartOblasts(obMap);
   updateChartDemographics(demoMap);
   updateChartTimeline(timeMap);
   updateChartFuel(fuelMap);
}

const numFormat = new Intl.NumberFormat('uk-UA');
function animateValue(id, end, isCurrency = false) {
   const obj = document.getElementById(id);
   let current = 0;
   const steps = 20;
   const stepVal = end / steps;

   clearInterval(obj.timer);
   obj.timer = setInterval(() => {
      current += stepVal;
      if (current >= end) {
         current = end;
         clearInterval(obj.timer);
      }
      obj.innerText = isCurrency ? current.toLocaleString('uk-UA', { maximumFractionDigits: 0 }) : numFormat.format(Math.round(current));
   }, 30);
}

// === ЛОГІКА ГРАФІКІВ (CHART.JS) ===
function updateChartOblasts(dataMap) {
   const labels = Object.keys(dataMap).sort((a, b) => dataMap[b] - dataMap[a]).slice(0, 10);
   const values = labels.map(l => dataMap[l]);

   if (charts.oblasts) charts.oblasts.destroy();
   const ctx = document.getElementById('chartOblasts').getContext('2d');
   charts.oblasts = new Chart(ctx, {
      type: 'bar',
      data: {
         labels: labels,
         datasets: [{
            label: 'Осіб охоплено',
            data: values,
            backgroundColor: colors.red,
            borderRadius: 4
         }]
      },
      options: { responsive: true, maintainAspectRatio: false }
   });
}

function updateChartDemographics(demoMap) {
   if (charts.demographics) charts.demographics.destroy();
   const ctx = document.getElementById('chartDemographics').getContext('2d');
   charts.demographics = new Chart(ctx, {
      type: 'doughnut',
      data: {
         labels: ['Жінки/Дівчата', 'Чоловіки/Хлопці', 'Особи з інвалідністю (в т.ч.)'],
         datasets: [{
            data: [demoMap.women, demoMap.men, demoMap.disabled],
            backgroundColor: [colors.red, colors.gray, colors.lightRed],
            borderWidth: 0
         }]
      },
      options: { responsive: true, maintainAspectRatio: false, cutout: '65%' }
   });
}

function updateChartTimeline(timeMap) {
   const labels = Object.keys(timeMap).sort();
   const values = labels.map(l => timeMap[l]);

   if (charts.timeline) charts.timeline.destroy();
   const ctx = document.getElementById('chartTimeline').getContext('2d');
   charts.timeline = new Chart(ctx, {
      type: 'line',
      data: {
         labels: labels,
         datasets: [{
            label: 'Осіб',
            data: values,
            borderColor: colors.red,
            backgroundColor: 'rgba(113, 19, 36, 0.1)',
            fill: true,
            tension: 0.3,
            pointBackgroundColor: colors.red
         }]
      },
      options: { responsive: true, maintainAspectRatio: false }
   });
}

function updateChartFuel(fuelMap) {
   if (charts.fuel) charts.fuel.destroy();
   const ctx = document.getElementById('chartFuel').getContext('2d');
   charts.fuel = new Chart(ctx, {
      type: 'bar',
      data: {
         labels: ['Дрова', 'Брикети', 'Вугілля'],
         datasets: [{
            label: 'Тонни',
            data: [fuelMap.wood, fuelMap.briq, fuelMap.coal],
            backgroundColor: [colors.wood, colors.briq, colors.coal],
            borderRadius: 4
         }]
      },
      options: {
         responsive: true,
         maintainAspectRatio: false,
         indexAxis: 'y',
         plugins: { legend: { display: false } }
      }
   });
}