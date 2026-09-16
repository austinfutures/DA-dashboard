/* =========================================================
   CUSTOM ANALYSIS DASHBOARD
   Cleaned-up, defensive replacement
========================================================= */

import { cleanCaseRow, cleanDefRow } from '../cleanData.js';

/* =========================================================
   CONSTANTS
========================================================= */

const COLORS = [
  '#000000',
  '#e91e63',
  '#ff9800',
  '#ffe600',
  '#4caf50',
  '#00bcd4',
  '#9c27b0',
  '#f44336',
  '#3f51b5',
  '#2196f3',
  '#795548'
];

const STATUS_TYPES = [
  'Filed',
  'Dismissed',
  'Rejected',
  'Open',
  'Sentenced',
  'accepted',
  'rejected'
];

const COMPLETED_METRICS = new Set([
  'Sentenced',
  'Dismissed'
]);

const MONTH_NAMES = [
  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
];

/*
  Files are expected at:

  customAnalysis.html
  data/
    cases_2025.xlsx
    cases_2024.xlsx
    defendants_2025.xlsx

  If your data folder is somewhere else, change this.
*/
const FOLDER = './data/';

/* =========================================================
   STATE
========================================================= */

const caseRows = [];
const defRows = [];

let rows = caseRows;

let charts = [];
let pieChart = null;
let largeChart = null;

let MIN_CASE_YEAR = null;
let LATEST_CASE_YEAR = null;
let LATEST_DEF_YEAR = null;

const loaded = {
  cases: false,
  defendants: false
};

/* =========================================================
   DOM HELPERS
========================================================= */

function $(id) {
  return document.getElementById(id);
}

function getDataset() {
  return $('dataset')?.value || 'cases';
}

function getMetric() {
  return $('metric')?.value || 'all_cases';
}

function getRange() {
  return $('range')?.value || 'last12';
}

function getDimension() {
  return $('dimension')?.value || '';
}

/* =========================================================
   TEXT HELPERS
========================================================= */

function prettyName(key) {
  const map = {
    all_cases: 'All Cases Received',
    accepted: 'Accepted Cases',
    rejected: 'Rejected Cases',

    Filed: 'Cases Filed by Prosecutor',
    Dismissed: 'Dismissed by Court',
    Rejected: 'Declined to Prosecute',
    Open: 'Open Case',
    Sentenced: 'Sentenced'
  };

  return map[key] ||
    String(key || '')
      .replace(/_/g, ' ')
      .replace(/\b\w/g, c => c.toUpperCase());
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function isValidDate(date) {
  return date instanceof Date &&
    !Number.isNaN(date.getTime());
}

function fmt(value, isCount = true) {
  if (value == null || Number.isNaN(value)) {
    return 'N/A';
  }

  if (!isCount) {
    return `${value}%`;
  }

  return getDataset() === 'cases'
    ? `${value} cases`
    : `${value} defendants`;
}

function fadeColor(hex, alpha = 0.18) {
  if (!hex || typeof hex !== 'string') {
    return `rgba(0,0,0,${alpha})`;
  }

  const clean = hex.replace('#', '');
  const number = parseInt(clean, 16);

  if (Number.isNaN(number)) {
    return `rgba(0,0,0,${alpha})`;
  }

  const r = (number >> 16) & 255;
  const g = (number >> 8) & 255;
  const b = number & 255;

  return `rgba(${r},${g},${b},${alpha})`;
}

/* =========================================================
   DATE / BUCKET HELPERS
========================================================= */

function keyFromRecord(record, range, mode) {
  if (!record) return null;

  const year = mode === 'status'
    ? record.status_year
    : record.year;

  const month = mode === 'status'
    ? record.status_month
    : record.month;

  const quarter = mode === 'status'
    ? record.status_quarter
    : record.quarter;

  if (!year || !month) {
    return null;
  }

  if (range === 'monthly' || range === 'last12') {
    return `${year}-${month}`;
  }

  if (range === 'quarterly') {
    return `${year}-Q${quarter}`;
  }

  if (range === 'annual') {
    return String(year);
  }

  return `${year}-${month}`;
}

function buildBuckets(
  data,
  range,
  mode,
  metric,
  floorYear = null,
  fixedYear = null
) {
  const safeRows = Array.isArray(data) ? data : [];

  /*
    Last 12:
    If a latest file year exists, show Jan-Dec
    of that latest available year.
  */
  if (range === 'last12') {
    if (Number.isInteger(fixedYear)) {
      const buckets = [];

      for (let i = 0; i < 12; i++) {
        const month = i + 1;

        buckets.push({
          y: fixedYear,
          m: month,
          label: `${MONTH_NAMES[i]} '${String(fixedYear).slice(-2)}`,
          key: `${fixedYear}-${month}`
        });
      }

      return floorYear
        ? buckets.filter(b => b.y >= floorYear)
        : buckets;
    }

    /*
      Fallback rolling 12-month window.
    */
    const validRows = safeRows.filter(record => {
      if (mode !== 'status') return true;

      if (COMPLETED_METRICS.has(metric)) {
        return record.status === metric;
      }

      return true;
    });

    const timestamps = validRows
      .map(record =>
        mode === 'status'
          ? record.status_ts
          : record.ts
      )
      .filter(Number.isFinite);

    if (!timestamps.length) {
      return [];
    }

    const maxTimestamp = Math.max(...timestamps);
    const maxDate = new Date(maxTimestamp);

    const startYear = maxDate.getFullYear();
    const startMonth = maxDate.getMonth();

    const buckets = [];

    for (let i = 11; i >= 0; i--) {
      const offset = startMonth - i;

      const year =
        startYear + Math.floor(offset / 12);

      const monthIndex =
        ((offset % 12) + 12) % 12;

      buckets.push({
        y: year,
        m: monthIndex + 1,
        label: `${MONTH_NAMES[monthIndex]} '${String(year).slice(-2)}`,
        key: `${year}-${monthIndex + 1}`
      });
    }

    return floorYear
      ? buckets.filter(b => b.y >= floorYear)
      : buckets;
  }

  /*
    Monthly
  */
  if (range === 'monthly') {
    const years = [
      ...new Set(
        safeRows.map(record =>
          mode === 'status'
            ? record.status_year
            : record.year
        )
      )
    ]
      .filter(year =>
        year && (!floorYear || year >= floorYear)
      )
      .sort((a, b) => a - b);

    const buckets = [];

    years.forEach(year => {
      MONTH_NAMES.forEach((_, index) => {
        buckets.push({
          y: year,
          m: index + 1,
          label: `${MONTH_NAMES[index]} '${String(year).slice(-2)}`,
          key: `${year}-${index + 1}`
        });
      });
    });

    return buckets;
  }

  /*
    Quarterly
  */
  if (range === 'quarterly') {
    const years = [
      ...new Set(
        safeRows.map(record =>
          mode === 'status'
            ? record.status_year
            : record.year
        )
      )
    ]
      .filter(year =>
        year && (!floorYear || year >= floorYear)
      )
      .sort((a, b) => a - b);

    const buckets = [];

    years.forEach(year => {
      [1, 2, 3, 4].forEach(quarter => {
        buckets.push({
          y: year,
          q: quarter,
          label: `Q${quarter} '${String(year).slice(-2)}`,
          key: `${year}-Q${quarter}`
        });
      });
    });

    return buckets;
  }

  /*
    Annual
  */
  const years = [
    ...new Set(
      safeRows.map(record =>
        mode === 'status'
          ? record.status_year
          : record.year
      )
    )
  ]
    .filter(year =>
      year && (!floorYear || year >= floorYear)
    )
    .sort((a, b) => a - b);

  return years.map(year => ({
    y: year,
    label: String(year),
    key: String(year)
  }));
}

/* =========================================================
   CHART.JS HOVER PLUGIN
========================================================= */

const hoverBar = {
  id: 'hoverBar',

  afterDraw(chart) {
    if (!chart) return;
    if (chart.config?.type !== 'line') return;

    const {
      ctx,
      tooltip,
      chartArea
    } = chart;

    /*
      Chart.js may call afterDraw before tooltip
      or active elements have been initialized.
    */
    if (!tooltip) return;
    if (!tooltip._active) return;
    if (!tooltip._active.length) return;
    if (!chartArea) return;

    const active = tooltip._active[0];

    if (!active || !active.element) {
      return;
    }

    const x = active.element.x;

    ctx.save();

    ctx.fillStyle = 'rgba(0,0,0,.07)';

    ctx.fillRect(
      x - 18,
      chartArea.top,
      36,
      chartArea.bottom - chartArea.top
    );

    ctx.restore();
  }
};

if (typeof Chart !== 'undefined') {
  Chart.register(hoverBar);
}

/* =========================================================
   FILE DISCOVERY
========================================================= */

async function discoverYears(type) {
  const base = type === 'defendants'
    ? 'defendants'
    : 'cases';

  const found = [];
  const currentYear = new Date().getFullYear();

  /*
    Do not stop at the first missing year.
    Your files may have gaps.
  */
  for (let year = currentYear; year >= 2015; year--) {
    try {
      const response = await fetch(
        `${FOLDER}${base}_${year}.xlsx`,
        { method: 'HEAD' }
      );

      if (response.ok) {
        found.push(year);
      }
    } catch (error) {
      console.warn(
        `Could not check ${base}_${year}.xlsx`,
        error
      );
    }
  }

  return found.sort((a, b) => a - b);
}

/* =========================================================
   DATA LOADING
========================================================= */

async function ensureLoaded(dataset) {
  if (loaded[dataset]) {
    return;
  }

  const years = await discoverYears(dataset);

  if (!years.length) {
    console.warn(
      `No ${dataset} Excel files found in ${FOLDER}`
    );

    loaded[dataset] = true;
    return;
  }

  if (dataset === 'cases') {
    MIN_CASE_YEAR = Math.min(...years);
    LATEST_CASE_YEAR = Math.max(...years);

    await loadCasesData(years);
  } else {
    LATEST_DEF_YEAR = Math.max(...years);

    await loadDefendantsData(years);
  }

  loaded[dataset] = true;
}

async function loadCasesData(years) {
  for (const year of years) {
    const filename = `${FOLDER}cases_${year}.xlsx`;

    try {
      const response = await fetch(filename);

      if (!response.ok) {
        console.warn(
          `Skipping missing file: ${filename}`
        );
        continue;
      }

      const buffer = await response.arrayBuffer();

      const workbook = XLSX.read(buffer, {
        type: 'array'
      });

      const firstSheet =
        workbook.Sheets[workbook.SheetNames[0]];

      if (!firstSheet) {
        console.warn(`No worksheet found in ${filename}`);
        continue;
      }

      const rawRows = XLSX.utils.sheet_to_json(
        firstSheet,
        { defval: '' }
      );

      rawRows.forEach(raw => {
        const cleaned = cleanCaseRow(raw);

        if (!cleaned) return;

        const receivedDate = new Date(cleaned.date_da);

        if (!isValidDate(receivedDate)) {
          return;
        }

        cleaned.ts = receivedDate.getTime();
        cleaned.year = receivedDate.getFullYear();
        cleaned.month = receivedDate.getMonth() + 1;
        cleaned.quarter =
          Math.floor(receivedDate.getMonth() / 3) + 1;

        const statusDate = new Date(
          cleaned.status_date || ''
        );

        if (isValidDate(statusDate)) {
          cleaned.status_ts = statusDate.getTime();
          cleaned.status_year = statusDate.getFullYear();
          cleaned.status_month = statusDate.getMonth() + 1;
          cleaned.status_quarter =
            Math.floor(statusDate.getMonth() / 3) + 1;
        } else {
          cleaned.status_ts = null;
          cleaned.status_year = null;
          cleaned.status_month = null;
          cleaned.status_quarter = null;
        }

        caseRows.push(cleaned);
      });
    } catch (error) {
      console.error(
        `Failed loading ${filename}`,
        error
      );
    }
  }
}

async function loadDefendantsData(years) {
  for (const year of years) {
    const filename = `${FOLDER}defendants_${year}.xlsx`;

    try {
      const response = await fetch(filename);

      if (!response.ok) {
        console.warn(
          `Skipping missing file: ${filename}`
        );
        continue;
      }

      const buffer = await response.arrayBuffer();

      const workbook = XLSX.read(buffer, {
        type: 'array'
      });

      const firstSheet =
        workbook.Sheets[workbook.SheetNames[0]];

      if (!firstSheet) {
        console.warn(`No worksheet found in ${filename}`);
        continue;
      }

      const rawRows = XLSX.utils.sheet_to_json(
        firstSheet,
        { defval: '' }
      );

      rawRows.forEach(raw => {
        const cleaned = cleanDefRow(raw);

        if (!cleaned) return;

        const date =
          raw['Case Received By DA'] ||
          raw['case received by da'] ||
          raw['Case Received'] ||
          raw['Case Received Case ID'] ||
          '';

        const receivedDate = new Date(date);

        if (!isValidDate(receivedDate)) {
          return;
        }

        const age = Number.isFinite(cleaned.age)
          ? cleaned.age
          : null;

        let ageGroup = 'Not reported';

        if (age !== null) {
          if (age < 18) ageGroup = '<18';
          else if (age <= 24) ageGroup = '18–24';
          else if (age <= 34) ageGroup = '25–34';
          else if (age <= 49) ageGroup = '35–49';
          else if (age <= 64) ageGroup = '50–64';
          else ageGroup = '65+';
        }

        defRows.push({
          ...cleaned,

          date_da: date,

          ts: receivedDate.getTime(),
          year: receivedDate.getFullYear(),
          month: receivedDate.getMonth() + 1,
          quarter:
            Math.floor(receivedDate.getMonth() / 3) + 1,

          age_group: ageGroup
        });
      });
    } catch (error) {
      console.error(
        `Failed loading ${filename}`,
        error
      );
    }
  }
}

/* =========================================================
   CONTROLS
========================================================= */

function setupControls() {
  const controlIds = [
    'dataset',
    'metric',
    'range',
    'dimension'
  ];

  controlIds.forEach(id => {
    const element = $(id);

    if (!element) {
      console.warn(`Missing control: #${id}`);
      return;
    }

    element.addEventListener('change', async () => {
      try {
        if (id === 'dataset') {
          const dataset = getDataset();

          await ensureLoaded(dataset);

          initDimension();
        }

        build();
      } catch (error) {
        console.error(
          `Error handling ${id} change`,
          error
        );
      }
    });
  });

  const pieToggle = $('pieToggle');

  if (pieToggle) {
    pieToggle.addEventListener('change', () => {
      build();
    });
  }
}

/* =========================================================
   DIMENSION DROPDOWN
========================================================= */

function initDimension() {
  const dataset = getDataset();
  const select = $('dimension');

  if (!select) return;

  const source = dataset === 'cases'
    ? caseRows[0]
    : defRows[0];

  /*
    Prevent Object.keys(undefined)
  */
  if (!source) {
    select.innerHTML = '';
    console.warn(
      `No ${dataset} data available for dimension dropdown.`
    );
    return;
  }

  const ignore = [
    'case_id',
    'date_da',
    'year',
    'month',
    'quarter',
    'ts',
    'days_to_file',
    'days_file_to_sent',
    'age',
    'status',
    'status_date',
    'status_year',
    'status_month',
    'status_quarter',
    'status_ts'
  ];

  let keys = Object.keys(source)
    .filter(key => !ignore.includes(key));

  if (dataset === 'cases') {
    keys = keys.filter(key =>
      ![
        'ethnicity',
        'gender',
        'county_res',
        'age_group'
      ].includes(key)
    );
  } else {
    keys = keys.filter(key =>
      [
        'ethnicity',
        'gender',
        'county_res',
        'age_group'
      ].includes(key)
    );
  }

  select.innerHTML = keys.map(key => `
    <option value="${escapeHtml(key)}">
      ${escapeHtml(
        key.replace(/_/g, ' ')
          .replace(/\b\w/g, c => c.toUpperCase())
      )}
    </option>
  `).join('');
}

/* =========================================================
   BUILD DASHBOARD
========================================================= */

export function build() {
  const dataset = getDataset();
  const isCaseMode = dataset === 'cases';

  rows = isCaseMode ? caseRows : defRows;

  /*
    Show/hide metric picker.
  */
  const metricElement = $('metric');

  if (metricElement?.parentElement) {
    metricElement.parentElement.style.display =
      isCaseMode ? '' : 'none';
  }

  /*
    Destroy small charts safely.
  */
  charts.forEach(chart => {
    try {
      chart.destroy();
    } catch (error) {
      console.warn('Could not destroy chart', error);
    }
  });

  charts = [];

  /*
    Clear comparison chart safely.
  */
  if (largeChart) {
    largeChart.data.datasets = [];
    largeChart.data.labels = [];
    largeChart.update('none');

    const compareSection = $('compareSection');

    if (compareSection) {
      compareSection.style.display = 'none';
    }
  }

  /*
    No data state.
  */
  if (!rows.length) {
    renderNoData(
      `No ${dataset} data is currently available.`
    );
    return;
  }

  /*
    Rebuild temporary AlaSQL table.
    Wrapped defensively because AlaSQL is not needed
    for the actual aggregation below.
  */
  try {
    alasql('DROP TABLE IF EXISTS cases');
    alasql('CREATE TABLE cases');
    alasql('INSERT INTO cases SELECT * FROM ?', [rows]);
  } catch (error) {
    console.warn('AlaSQL setup skipped:', error);
  }

  const metric = isCaseMode
    ? getMetric()
    : 'all_cases';

  const timeMode =
    isCaseMode &&
    COMPLETED_METRICS.has(metric)
      ? 'status'
      : 'received';

  /*
    Closed-case metrics use latest 12-month display.
  */
  const rangeElement = $('range');

  if (rangeElement) {
    if (
      isCaseMode &&
      COMPLETED_METRICS.has(metric)
    ) {
      rangeElement.value = 'last12';
      rangeElement.disabled = true;

      Array.from(rangeElement.options).forEach(option => {
        option.hidden = option.value !== 'last12';
      });
    } else {
      rangeElement.disabled = false;

      Array.from(rangeElement.options).forEach(option => {
        option.hidden = false;
      });
    }
  }

  const range = getRange();

  const floorYear = isCaseMode
    ? MIN_CASE_YEAR
    : null;

  const fixedYear = isCaseMode
    ? LATEST_CASE_YEAR
    : LATEST_DEF_YEAR;

  const pieToggle = $('pieToggle');

  const pieMode =
    Boolean(pieToggle?.checked) &&
    (
      isCaseMode
        ? (
            metric === 'all_cases' ||
            STATUS_TYPES.includes(metric)
          )
        : true
    );

  const dimension = getDimension();

  const buckets = buildBuckets(
    rows,
    range,
    timeMode,
    metric,
    floorYear,
    fixedYear
  );

  if (!buckets.length) {
    renderNoData(
      'No data found for the selected filters.'
    );
    return;
  }

  /*
    Aggregation stores.
  */
  const allReceived = {};
  const statusReceived = {};
  const groupAllReceived = {};
  const groupStatusReceived = {};

  const statusCompleted = {};
  const groupStatusCompleted = {};

  rows.forEach(record => {
    let group = record[dimension];

    if (
      group === undefined ||
      group === null ||
      group === ''
    ) {
      group = 'Unknown';
    }

    /*
      Received-date bucket.
    */
    const receivedKey = keyFromRecord(
      record,
      range,
      'received'
    );

    if (receivedKey) {
      allReceived[receivedKey] =
        (allReceived[receivedKey] || 0) + 1;

      if (!groupAllReceived[group]) {
        groupAllReceived[group] = {};
      }

      groupAllReceived[group][receivedKey] =
        (groupAllReceived[group][receivedKey] || 0) + 1;

      /*
        Status counts by received date.
      */
      if (isCaseMode && record.status) {
        if (!statusReceived[record.status]) {
          statusReceived[record.status] = {};
        }

        statusReceived[record.status][receivedKey] =
          (statusReceived[record.status][receivedKey] || 0) + 1;

        if (!groupStatusReceived[record.status]) {
          groupStatusReceived[record.status] = {};
        }

        if (!groupStatusReceived[record.status][group]) {
          groupStatusReceived[record.status][group] = {};
        }

        groupStatusReceived[record.status][group][receivedKey] =
          (groupStatusReceived[record.status][group][receivedKey] || 0) + 1;
      }
    }

    /*
      Completed/status-date bucket.
    */
    const statusKey = keyFromRecord(
      record,
      range,
      'status'
    );

    if (
      statusKey &&
      isCaseMode &&
      record.status
    ) {
      if (!statusCompleted[record.status]) {
        statusCompleted[record.status] = {};
      }

      statusCompleted[record.status][statusKey] =
        (statusCompleted[record.status][statusKey] || 0) + 1;

      if (!groupStatusCompleted[record.status]) {
        groupStatusCompleted[record.status] = {};
      }

      if (!groupStatusCompleted[record.status][group]) {
        groupStatusCompleted[record.status][group] = {};
      }

      groupStatusCompleted[record.status][group][statusKey] =
        (groupStatusCompleted[record.status][group][statusKey] || 0) + 1;
    }
  });

  /*
    Determine which aggregation applies to metric.
  */
  function metricBuckets(selectedMetric) {
    if (!isCaseMode) {
      return {
        bucket: allReceived,
        group: groupAllReceived
      };
    }

    if (selectedMetric === 'all_cases') {
      return {
        bucket: allReceived,
        group: groupAllReceived
      };
    }

    if (selectedMetric === 'rejected') {
      return {
        bucket: statusReceived.Rejected || {},
        group: groupStatusReceived.Rejected || {}
      };
    }

    if (selectedMetric === 'accepted') {
      const bucket = {};
      const group = {};

      Object.keys(allReceived).forEach(key => {
        bucket[key] =
          (allReceived[key] || 0) -
          (statusReceived.Rejected?.[key] || 0);
      });

      Object.keys(groupAllReceived).forEach(groupName => {
        group[groupName] = {};

        Object.keys(groupAllReceived[groupName]).forEach(key => {
          const rejected =
            groupStatusReceived.Rejected?.[groupName]?.[key] || 0;

          group[groupName][key] =
            (groupAllReceived[groupName][key] || 0) -
            rejected;
        });
      });

      return { bucket, group };
    }

    if (
      selectedMetric === 'Sentenced' ||
      selectedMetric === 'Dismissed'
    ) {
      return {
        bucket: statusCompleted[selectedMetric] || {},
        group: groupStatusCompleted[selectedMetric] || {}
      };
    }

    if (
      ['Filed', 'Open', 'Rejected'].includes(selectedMetric)
    ) {
      return {
        bucket: statusReceived[selectedMetric] || {},
        group: groupStatusReceived[selectedMetric] || {}
      };
    }

    return {
      bucket: {},
      group: {}
    };
  }

  const {
    bucket: bucketBase,
    group: groupBase
  } = metricBuckets(metric);

  const loading = $('pieLoading');

  if (loading) {
    loading.remove();
  }

  /*
    PIE / LINE MODE
  */
  if (pieMode) {
    const lineData = buckets.map(bucket =>
      bucketBase[bucket.key] || 0
    );

    renderLinePie(
      buckets,
      lineData,
      groupBase,
      metric
    );

    return;
  }

  /*
    STANDARD MULTI-LINE MODE
  */
  const allLabel = isCaseMode
    ? prettyName(metric)
    : 'All Defendants';

  const datasets = [
    {
      label: allLabel,
      color: '#000000',
      values: buckets.map(bucket =>
        bucketBase[bucket.key] || 0
      )
    },

    ...Object.keys(groupBase).map((group, index) => ({
      label: group,
      color: COLORS[(index + 1) % COLORS.length],
      values: buckets.map(bucket =>
        groupBase[group]?.[bucket.key] || 0
      )
    }))
  ];

  render(
    datasets,
    buckets.map(bucket => bucket.label),
    true
  );
}

/* =========================================================
   NO DATA DISPLAY
========================================================= */

function renderNoData(message) {
  const grid = $('chartGrid');

  if (!grid) return;

  grid.innerHTML = `
    <div class="chart-box">
      <div class="chart-head">
        <div class="chart-title">No data</div>
      </div>
      <div class="chart-number">No data</div>
      <div style="font-size:.9rem;color:#666">
        ${escapeHtml(message)}
      </div>
    </div>
  `;

  const compareSection = $('compareSection');

  if (compareSection) {
    compareSection.style.display = 'none';
  }

  const loading = $('pieLoading');

  if (loading) {
    loading.remove();
  }
}

/* =========================================================
   STANDARD LINE CHART RENDERER
========================================================= */

function render(datasets, labels, isCount = true) {
  const grid = $('chartGrid');

  if (!grid) return;

  grid.innerHTML = '';

  charts.forEach(chart => {
    try {
      chart.destroy();
    } catch (error) {
      console.warn('Chart destroy failed', error);
    }
  });

  charts = [];

  if (!datasets.length || !labels.length) {
    renderNoData('No data available.');
    return;
  }

  const firstLabel = labels[0];
  const lastLabel = labels.at(-1);

  datasets.forEach((dataset, index) => {
    const canvasId = `c${index}`;

    grid.insertAdjacentHTML('beforeend', `
      <div class="chart-box">
        <div class="chart-head">
          <div class="chart-title">
            ${escapeHtml(dataset.label)}
          </div>
          <div class="chart-month" id="m${index}"></div>
        </div>

        <div class="chart-number" id="v${index}">
          ${escapeHtml(
            fmt(dataset.values.at(-1), isCount)
          )}
        </div>

        <div class="chart-canvas">
          <canvas
            id="${canvasId}"
            width="280"
            height="100">
          </canvas>
        </div>

        <div class="range-labels">
          <span>${escapeHtml(firstLabel)}</span>
          <span>${escapeHtml(lastLabel)}</span>
        </div>

        <label style="margin-top:8px;display:block;">
          <input
            type="checkbox"
            onchange="toggleLargeChart(${index})">
          Compare
        </label>
      </div>
    `);

    const canvas = $(canvasId);

    if (!canvas) return;

    const context = canvas.getContext('2d');

    if (!context) return;

    const chart = new Chart(context, {
      type: 'line',

      data: {
        labels: [...labels],

        datasets: [
          {
            label: dataset.label,
            data: [...dataset.values],
            borderColor: dataset.color,
            backgroundColor: dataset.color,
            tension: 0.18,
            pointRadius: 0,
            pointHoverRadius: 5
          }
        ]
      },

      options: {
        responsive: false,
        animation: false,

        plugins: {
          legend: {
            display: false
          },

          tooltip: {
            enabled: false
          }
        },

        interaction: {
          mode: 'nearest',
          axis: 'x',
          intersect: false
        },

        scales: {
          x: {
            display: false
          },

          y: {
            beginAtZero: true,

            ticks: {
              callback(value) {
                return Number.isInteger(value)
                  ? value
                  : '';
              }
            }
          }
        },

        onHover(event, elements) {
          if (elements?.length) {
            hover(
              elements[0].index,
              labels,
              isCount
            );
          } else {
            clear(isCount);
          }
        }
      },

      plugins: [hoverBar]
    });

    charts.push(chart);
  });
}

/* =========================================================
   LINE + PIE MODE
========================================================= */

function renderLinePie(
  buckets,
  lineData,
  groupCounts,
  metricName
) {
  const grid = $('chartGrid');

  if (!grid) return;

  const isCaseMode = getDataset() === 'cases';

  const unitWord = isCaseMode
    ? 'cases'
    : 'defendants';

  const titleText = isCaseMode
    ? prettyName(metricName)
    : 'All Defendants';

  /*
    Destroy any previous pie chart before replacing it.
  */
  if (pieChart) {
    try {
      pieChart.destroy();
    } catch (error) {
      console.warn('Could not destroy pie chart', error);
    }

    pieChart = null;
  }

  grid.innerHTML = `
    <div class="chart-box" style="flex:1 1 100%;">
      <div class="chart-head">
        <div class="chart-title">
          ${escapeHtml(titleText)}
        </div>

        <div class="chart-month" id="lineMonth"></div>
      </div>

      <div class="chart-number" id="lineValue">
        ${lineData.at(-1)} ${unitWord}
      </div>

      <canvas id="lineMain" height="140"></canvas>
    </div>

    <div class="chart-box" style="flex:1 1 320px;">
      <div class="chart-head">
        <div class="chart-title">Breakdown</div>
      </div>

      <div class="chart-number" id="sliceValue"></div>

      <canvas id="pieMain" height="140"></canvas>
    </div>
  `;

  const lineCanvas = $('lineMain');
  const pieCanvas = $('pieMain');

  if (!lineCanvas || !pieCanvas) {
    console.warn('Pie/line canvases were not created.');
    return;
  }

  const lineContext = lineCanvas.getContext('2d');
  const pieContext = pieCanvas.getContext('2d');

  if (!lineContext || !pieContext) {
    return;
  }

  const labels = buckets.map(bucket => bucket.label);

  let originalColors = [];

  /*
    Create line chart.
  */
  new Chart(lineContext, {
    type: 'line',

    data: {
      labels,

      datasets: [
        {
          label: prettyName(metricName),
          data: [...lineData],
          borderColor: '#000000',
          backgroundColor: '#000000',
          tension: 0.18,
          pointRadius: 0,
          pointHoverRadius: 5
        }
      ]
    },

    options: {
      responsive: true,
      animation: false,

      plugins: {
        legend: {
          display: false
        },

        tooltip: {
          enabled: false
        }
      },

      interaction: {
        mode: 'nearest',
        axis: 'x',
        intersect: false
      },

      scales: {
        y: {
          beginAtZero: true
        }
      },

      onHover(event, elements) {
        if (!elements?.length) return;

        const index = elements[0].index;

        updatePie(index);

        const valueElement = $('lineValue');
        const monthElement = $('lineMonth');

        if (valueElement) {
          valueElement.textContent =
            `${lineData[index]} ${unitWord}`;
        }

        if (monthElement) {
          monthElement.textContent = labels[index];
        }
      }
    }
  });

  /*
    Create empty pie chart.
  */
  pieChart = new Chart(pieContext, {
    type: 'pie',

    data: {
      labels: [],

      datasets: [
        {
          data: [],
          backgroundColor: []
        }
      ]
    },

    options: {
      plugins: {
        legend: {
          position: 'right'
        },

        tooltip: {
          enabled: false
        }
      },

      onHover(event, elements) {
        /*
          Important: pieChart may have been destroyed
          during a dashboard rebuild.
        */
        if (!pieChart) return;

        const box = $('sliceValue');

        if (!elements?.length) {
          pieChart.data.datasets[0].backgroundColor =
            originalColors;

          pieChart.update('none');

          if (box) {
            box.textContent = '';
            box.style.color = '#000000';
          }

          return;
        }

        const index = elements[0].index;

        const label =
          pieChart.data.labels[index];

        const value =
          pieChart.data.datasets[0].data[index];

        pieChart.data.datasets[0].backgroundColor =
          originalColors.map((color, colorIndex) =>
            colorIndex === index
              ? color
              : fadeColor(color)
          );

        pieChart.update('none');

        if (box) {
          box.textContent =
            `${label}: ${value} ${unitWord}`;

          box.style.color =
            originalColors[index] || '#000000';
        }
      }
    }
  });

  /*
    Update pie chart safely.
  */
  function updatePie(index) {
    if (!pieChart) return;
    if (!buckets[index]) return;
    if (!pieChart.data) return;
    if (!pieChart.data.datasets?.length) return;

    const key = buckets[index].key;

    const sliceLabels = [];
    const sliceData = [];
    const sliceColors = [];

    let colorIndex = 1;

    Object.keys(groupCounts || {}).forEach(group => {
      const value =
        groupCounts[group]?.[key] || 0;

      if (!value) return;

      sliceLabels.push(group);
      sliceData.push(value);

      sliceColors.push(
        COLORS[colorIndex % COLORS.length]
      );

      colorIndex++;
    });

    originalColors = [...sliceColors];

    pieChart.data.labels = sliceLabels;

    pieChart.data.datasets[0].data =
      sliceData;

    pieChart.data.datasets[0].backgroundColor =
      sliceColors;

    pieChart.update('none');
  }

  /*
    Initial pie state.
  */
  if (buckets.length && pieChart) {
    updatePie(buckets.length - 1);
  }

  const monthElement = $('lineMonth');

  if (monthElement) {
    monthElement.textContent = labels.at(-1);
  }
}

/* =========================================================
   LARGE COMPARISON CHART
========================================================= */

function initLargeChart() {
  const canvas = $('largeChart');

  if (!canvas) {
    console.warn(
      'largeChart canvas not found in HTML.'
    );
    return;
  }

  const context = canvas.getContext('2d');

  if (!context) return;

  /*
    Destroy existing instance if necessary.
  */
  if (largeChart) {
    try {
      largeChart.destroy();
    } catch (error) {
      console.warn('Could not destroy large chart', error);
    }

    largeChart = null;
  }

  largeChart = new Chart(context, {
    type: 'line',

    data: {
      labels: [],
      datasets: []
    },

    options: {
      responsive: true,

      plugins: {
        legend: {
          position: 'top'
        }
      },

      interaction: {
        mode: 'nearest',
        axis: 'x',
        intersect: false
      },

      scales: {
        y: {
          beginAtZero: true
        }
      }
    }
  });
}

function toggleLargeChart(index) {
  /*
    Defensive checks prevent:
    Cannot read properties of null (reading 'data')
  */
  if (!largeChart) {
    console.warn(
      'Comparison chart is not initialized.'
    );
    return;
  }

  if (!Array.isArray(charts)) return;
  if (!charts[index]) return;

  const sourceChart = charts[index];

  if (!sourceChart.data) return;
  if (!sourceChart.data.datasets?.length) return;

  const sourceDataset =
    sourceChart.data.datasets[0];

  const label = sourceDataset.label;

  const existing =
    largeChart.data.datasets.find(
      dataset => dataset.label === label
    );

  if (existing) {
    largeChart.data.datasets =
      largeChart.data.datasets.filter(
        dataset => dataset.label !== label
      );
  } else {
    largeChart.data.datasets.push({
      label,

      data: [...sourceDataset.data],

      borderColor: sourceDataset.borderColor,
      backgroundColor: sourceDataset.borderColor,

      tension: 0.18,
      pointRadius: 0,
      pointHoverRadius: 4
    });

    if (!largeChart.data.labels.length) {
      largeChart.data.labels =
        [...sourceChart.data.labels];
    }
  }

  const compareSection = $('compareSection');

  if (compareSection) {
    compareSection.style.display =
      largeChart.data.datasets.length
        ? 'block'
        : 'none';
  }

  largeChart.update();

  if (!largeChart.data.datasets.length) {
    largeChart.data.labels = [];
    largeChart.update('none');
  }
}

/*
  Required because your generated checkbox uses:
  onchange="toggleLargeChart(index)"
*/
window.toggleLargeChart = toggleLargeChart;

/* =========================================================
   HOVER SYNCHRONIZATION
========================================================= */

function hover(index, labels, isCount) {
  charts.forEach((chart, chartIndex) => {
    if (!chart) return;
    if (!chart.data?.datasets?.length) return;

    chart.setActiveElements([
      {
        datasetIndex: 0,
        index
      }
    ]);

    chart.update('none');

    const value =
      chart.data.datasets[0].data[index];

    const valueElement = $(`v${chartIndex}`);
    const monthElement = $(`m${chartIndex}`);

    if (valueElement) {
      valueElement.textContent =
        fmt(value, isCount);
    }

    if (monthElement) {
      monthElement.textContent =
        labels[index] || '';
    }
  });
}

function clear(isCount) {
  charts.forEach((chart, index) => {
    if (!chart) return;
    if (!chart.data?.datasets?.length) return;

    chart.setActiveElements([]);
    chart.update('none');

    const values =
      chart.data.datasets[0].data;

    const value = values.at(-1);

    const valueElement = $(`v${index}`);
    const monthElement = $(`m${index}`);

    if (valueElement) {
      valueElement.textContent =
        fmt(value, isCount);
    }

    if (monthElement) {
      monthElement.textContent = '';
    }
  });
}

/* =========================================================
   OPTIONAL SLIDE PANEL CONTROLS
========================================================= */

/*
  These elements do not exist in your supplied HTML.
  Therefore, only initialize them if present.
*/

function setupOptionalPanelControls() {
  const wrap =
    document.querySelector('.panel-wrapper');

  const buttons =
    document.querySelectorAll('.view-toggle button');

  function activatePanel(index) {
    if (!wrap) return;

    wrap.style.transform =
      `translateX(-${index * 33.333}%)`;

    buttons.forEach((button, buttonIndex) => {
      button.classList.toggle(
        'active',
        buttonIndex === index
      );
    });
  }

  const toMain = $('toMain');
  const toStats = $('toStats');
  const toMonthly = $('toMonthly');

  if (toMain) {
    toMain.onclick = () => activatePanel(0);
  }

  if (toStats) {
    toStats.onclick = () => activatePanel(1);
  }

  if (toMonthly) {
    toMonthly.onclick = () => activatePanel(2);
  }
}

/* =========================================================
   INITIALIZATION
========================================================= */

async function initializeDashboard() {
  try {
    /*
      IMPORTANT:
      Initialize largeChart BEFORE build().
      Otherwise comparison controls can fire while
      largeChart is still null.
    */
    initLargeChart();

    setupControls();
    setupOptionalPanelControls();

    await ensureLoaded('cases');

    initDimension();

    build();
  } catch (error) {
    console.error(
      'Dashboard initialization failed:',
      error
    );

    renderNoData(
      'Unable to load dashboard data. Check the Excel files and data folder.'
    );
  }
}

/*
  Because this is a deferred module, DOM is normally ready,
  but this also makes the initialization robust.
*/
if (document.readyState === 'loading') {
  document.addEventListener(
    'DOMContentLoaded',
    initializeDashboard
  );
} else {
  initializeDashboard();
}

/* =========================================================
   EXPORTS
========================================================= */


window.build = build;