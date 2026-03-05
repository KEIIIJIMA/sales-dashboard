const yen = new Intl.NumberFormat("ja-JP", {
  style: "currency",
  currency: "JPY",
  maximumFractionDigits: 0,
});

const monthKeys = ["6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月", "4月", "5月"];
const salesTypes = [
  { key: "total", label: "全体売上", targetColor: "#006d77", actualColor: "#f28e00" },
  { key: "pile", label: "杭売上", targetColor: "#2f5597", actualColor: "#4f81ff" },
  { key: "design", label: "設計売上", targetColor: "#7a3e00", actualColor: "#c46a00" },
];

const fallbackTargetComparisonRows = {
  west: {
    totalTarget: [30600000, 31800000, 31900000, 29500000, 29700000, 31700000, 30900000, 35200000, 27200000, 40500000, 32400000, 28600000],
    pileTarget: [17600000, 18800000, 18900000, 17500000, 17700000, 18700000, 17900000, 21200000, 16200000, 23500000, 19400000, 16600000],
    designTarget: [13000000, 13000000, 13000000, 12000000, 12000000, 13000000, 13000000, 14000000, 11000000, 17000000, 13000000, 12000000],
    totalActual: [21673162, 23386332, 18801467, 19935806, 25832634, 30180043, 16559904, 21308517, 25338263, 0, 0, 0],
    pileActual: [16449886, 17851387, 15063890, 11863133, 17715509, 17713237, 12172500, 14069228, 18124539, 0, 0, 0],
    designActual: [5223276, 5534945, 3737577, 8072673, 8117125, 12466806, 4387404, 7239289, 7213724, 0, 0, 0],
  },
  tama: {
    totalTarget: [56100000, 53900000, 50600000, 49300000, 48700000, 50800000, 52300000, 42200000, 50200000, 50200000, 40700000, 46000000],
    pileTarget: [28100000, 27900000, 25600000, 24300000, 24700000, 25800000, 26300000, 20200000, 25200000, 25200000, 19700000, 23000000],
    designTarget: [28000000, 26000000, 25000000, 25000000, 24000000, 25000000, 26000000, 22000000, 25000000, 25000000, 21000000, 23000000],
    totalActual: [50425255, 50150828, 45255218, 53127853, 41566644, 40540609, 46411453, 41856379, 40436228, 21384, 0, 0],
    pileActual: [25222233, 22760035, 23557703, 24706070, 18363916, 20574346, 20820082, 20522017, 19170330, 0, 0, 0],
    designActual: [25203022, 27390793, 21697515, 28421783, 23202728, 19966263, 25591371, 21334362, 21265898, 21384, 0, 0],
  },
  kanagawa: {
    totalTarget: [32500000, 31600000, 30400000, 30400000, 32900000, 29600000, 27500000, 29800000, 31400000, 32500000, 43100000, 44300000],
    pileTarget: [19300000, 18900000, 17800000, 17800000, 19600000, 17700000, 15200000, 18000000, 19000000, 19300000, 25200000, 26200000],
    designTarget: [13200000, 12700000, 12600000, 12600000, 13300000, 11900000, 12300000, 11800000, 12400000, 13200000, 17900000, 18100000],
    totalActual: [33833813, 33367094, 31594719, 32210585, 36112097, 30446887, 28904126, 27415955, 29048301, 0, 0, 0],
    pileActual: [21067120, 19938347, 19929693, 19016747, 26028214, 17391596, 16804410, 16698851, 20188748, 0, 0, 0],
    designActual: [12766693, 13428747, 11665026, 13193838, 10083883, 13055291, 12099716, 10717104, 8859553, 0, 0, 0],
  },
  shonan: {
    totalTarget: [9200000, 8900000, 8600000, 8600000, 9300000, 8400000, 7700000, 8400000, 8900000, 9200000, 10700000, 11100000],
    pileTarget: [1400000, 1600000, 1200000, 1200000, 1600000, 1300000, 1000000, 1200000, 1300000, 1400000, 1600000, 1200000],
    designTarget: [7800000, 7300000, 7400000, 7400000, 7700000, 7100000, 6700000, 7200000, 7600000, 7800000, 9100000, 9900000],
    totalActual: [9162308, 9688878, 8678349, 9224650, 8850400, 8731390, 14391873, 16460862, 17642178, 0, 0, 0],
    pileActual: [266443, 2610977, 979679, 345205, 1290345, 1924682, 4627380, 7764155, 6895885, 0, 0, 0],
    designActual: [8895865, 7077901, 7698670, 8879445, 7560055, 6806708, 9764493, 8696707, 10746293, 0, 0, 0],
  },
};

function buildFiscalMonthNames(startYear) {
  return monthKeys.map((_, i) => {
    const monthNo = i < 7 ? 6 + i : i - 6;
    const year = i < 7 ? startYear : startYear + 1;
    return `${year}/${monthNo}`;
  });
}

function createMonthly(monthNames, seedTarget, seedActual) {
  return monthNames.map((month, i) => ({
    month,
    target: seedTarget + i * 250000,
    actual: Math.max(0, seedActual + i * 200000),
  }));
}

function createMonthlyFromTargetActual(monthNames, targets, actuals) {
  return monthNames.map((month, i) => ({
    month,
    target: Math.max(0, Math.round(targets[i] || 0)),
    actual: Math.max(0, Math.round(actuals[i] || 0)),
  }));
}

function pipelineSales(monthNames, rows) {
  return {
    total: createMonthlyFromTargetActual(monthNames, rows.totalTarget, rows.totalActual),
    pile: createMonthlyFromTargetActual(monthNames, rows.pileTarget, rows.pileActual),
    design: createMonthlyFromTargetActual(monthNames, rows.designTarget, rows.designActual),
  };
}

function normalizeTermsData() {
  if (window.EXCEL_TERMS && typeof window.EXCEL_TERMS === "object") {
    return window.EXCEL_TERMS;
  }

  const meta = window.EXCEL_FISCAL_META || { termNo: 39, startYear: 2025 };
  const rows = window.EXCEL_TARGET_COMPARISON || fallbackTargetComparisonRows;
  return {
    [String(meta.termNo)]: {
      termNo: meta.termNo,
      startYear: meta.startYear,
      rows,
    },
  };
}

const termsData = normalizeTermsData();
const termKeys = Object.keys(termsData).sort((a, b) => Number(a) - Number(b));

const state = {
  currentTermKey: termKeys[termKeys.length - 1],
  currentOrgKey: "導管部-西部事業所",
};

let baseOrgData = {};
let orgViews = [];

const termSelect = document.getElementById("termSelect");
const orgSelect = document.getElementById("orgSelect");
const salesTypeHint = document.getElementById("salesTypeHint");
const orgPath = document.getElementById("orgPath");
const fiscalInfo = document.getElementById("fiscalInfo");
const annualBreakdownTable = document.getElementById("annualBreakdownTable");
const monthlyTable = document.getElementById("monthlyTable");
const orgSummaryTable = document.getElementById("orgSummaryTable");
const yearRateEl = document.getElementById("yearRate");
const yearBarEl = document.getElementById("yearBar");
const yearRemainingEl = document.getElementById("yearRemaining");
const needPerMonthEl = document.getElementById("needPerMonth");
const chart = document.getElementById("yearChart");
const ctx = chart.getContext("2d");

function currentTerm() {
  return termsData[state.currentTermKey];
}

function currentMonthNames() {
  return buildFiscalMonthNames(currentTerm().startYear);
}

function buildBaseOrgData(termRows, monthNames) {
  return {
    "導管部-西部事業所": { hq: "導管部", unit: "西部事業所", source: "目標比較シート読込", sales: pipelineSales(monthNames, termRows.west) },
    "導管部-多摩事業所": { hq: "導管部", unit: "多摩事業所", source: "目標比較シート読込", sales: pipelineSales(monthNames, termRows.tama) },
    "導管部-神奈川事業所": { hq: "導管部", unit: "神奈川事業所", source: "目標比較シート読込", sales: pipelineSales(monthNames, termRows.kanagawa) },
    "導管部-神奈川事業所（湘南分室）": { hq: "導管部", unit: "神奈川事業所（湘南分室）", source: "目標比較シート読込", sales: pipelineSales(monthNames, termRows.shonan) },
    "設備本部-設計G": { hq: "設備本部", unit: "設計G", source: "手入力サンプル", sales: { total: createMonthly(monthNames, 5100000, 4800000) } },
    "設備本部-設備技術G": { hq: "設備本部", unit: "設備技術G", source: "手入力サンプル", sales: { total: createMonthly(monthNames, 5600000, 5300000) } },
    "設備本部-設備技術G（AMS）": { hq: "設備本部", unit: "設備技術G（AMS）", source: "手入力サンプル", sales: { total: createMonthly(monthNames, 3900000, 3700000) } },
  };
}

function rebuildTermContext() {
  const term = currentTerm();
  const monthNames = currentMonthNames();
  baseOrgData = buildBaseOrgData(term.rows || fallbackTargetComparisonRows, monthNames);
  orgViews = [
    ...Object.keys(baseOrgData).map((key) => ({ key, label: key })),
    { key: "__pipeline_hq__", label: "導管本部" },
    { key: "__equipment_hq__", label: "設備本部" },
    { key: "__company_total__", label: "全社合計" },
  ];

  if (!orgViews.some((v) => v.key === state.currentOrgKey)) {
    state.currentOrgKey = orgViews[0].key;
  }
}

function emptyMonths(monthNames) {
  return monthNames.map((month) => ({ month, target: 0, actual: 0 }));
}

function sumMonthlySeries(monthNames, seriesList) {
  const summed = emptyMonths(monthNames);
  seriesList.forEach((months) => {
    months.forEach((m, i) => {
      summed[i].target += m.target;
      summed[i].actual += m.actual;
    });
  });
  return summed;
}

function buildAggregateOrg(unit, source, selector, includePipelineBreakdown) {
  const monthNames = currentMonthNames();
  const orgs = Object.values(baseOrgData).filter(selector);
  const totalSeries = sumMonthlySeries(monthNames, orgs.map((org) => org.sales.total));
  const sales = { total: totalSeries };

  if (includePipelineBreakdown) {
    const pileSeries = sumMonthlySeries(monthNames, orgs.map((org) => org.sales.pile || emptyMonths(monthNames)));
    const designSeries = sumMonthlySeries(monthNames, orgs.map((org) => org.sales.design || emptyMonths(monthNames)));
    sales.pile = pileSeries;
    sales.design = designSeries;
  }

  return { hq: unit, unit, source, sales };
}

function currentOrg() {
  const key = state.currentOrgKey;
  if (baseOrgData[key]) return baseOrgData[key];
  if (key === "__pipeline_hq__") {
    return buildAggregateOrg("導管本部", "目標比較シート読込（導管部合算）", (org) => org.hq === "導管部", true);
  }
  if (key === "__equipment_hq__") {
    return buildAggregateOrg("設備本部", "手入力サンプル（設備本部合算）", (org) => org.hq === "設備本部", false);
  }
  return buildAggregateOrg("全社合計", "全組織合算", () => true, false);
}

function ratio(actual, target) {
  if (target <= 0) return 0;
  return (actual / target) * 100;
}

function clampPercent(v) {
  return Math.max(0, Math.min(v, 100));
}

function currentMonthIndex() {
  const thisMonthLabel = `${new Date().getMonth() + 1}月`;
  const idx = monthKeys.indexOf(thisMonthLabel);
  return idx >= 0 ? idx : 0;
}

function annualTotals(months) {
  const targetTotal = months.reduce((sum, m) => sum + m.target, 0);
  const actualTotal = months.reduce((sum, m) => sum + m.actual, 0);
  return { targetTotal, actualTotal };
}

function availableTypeDefs(org) {
  return org.hq === "導管部" || org.hq === "導管本部" ? salesTypes : [salesTypes[0]];
}

function getMonths(org, typeKey) {
  return org.sales[typeKey] || [];
}

function renderTermSelect() {
  termSelect.innerHTML = "";
  [...termKeys].sort((a, b) => Number(b) - Number(a)).forEach((key) => {
    const t = termsData[key];
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = `${t.termNo}期 (${t.startYear}年6月〜${t.startYear + 1}年5月)`;
    if (key === state.currentTermKey) opt.selected = true;
    termSelect.appendChild(opt);
  });
}

function renderOrgSelect() {
  orgSelect.innerHTML = "";
  orgViews.forEach((view) => {
    const opt = document.createElement("option");
    opt.value = view.key;
    opt.textContent = view.label;
    if (view.key === state.currentOrgKey) opt.selected = true;
    orgSelect.appendChild(opt);
  });
}

function renderOrgMeta() {
  const org = currentOrg();
  const term = currentTerm();
  const source = org.source ? ` / データ: ${org.source}` : "";
  fiscalInfo.textContent = `${term.termNo}期 (${term.startYear}年6月〜${term.startYear + 1}年5月) / 期を切替して過去期と比較可能`;
  orgPath.textContent = `本部: ${org.hq} / 組織: ${org.unit}${source}`;
  salesTypeHint.textContent =
    org.hq === "導管部" || org.hq === "導管本部"
      ? "導管部は 全体売上・杭売上・設計売上 を同時表示しています。"
      : "設備本部/全社合計は 全体売上を表示しています。";
}

function renderBreakdownTables() {
  const org = currentOrg();
  annualBreakdownTable.innerHTML = "";

  availableTypeDefs(org).forEach((typeDef) => {
    const months = getMonths(org, typeDef.key);
    const year = annualTotals(months);

    const yearRow = document.createElement("tr");
    yearRow.innerHTML = `
      <td>${typeDef.label}</td>
      <td>${yen.format(year.targetTotal)}</td>
      <td>${yen.format(year.actualTotal)}</td>
      <td class="rate-cell">${ratio(year.actualTotal, year.targetTotal).toFixed(1)}%</td>
    `;
    annualBreakdownTable.appendChild(yearRow);
  });
}

function renderMonthlyRows() {
  monthlyTable.innerHTML = "";
  const org = currentOrg();
  const labels = currentMonthNames();

  labels.forEach((monthLabel, index) => {
    const row = document.createElement("tr");
    row.innerHTML = `<td>${monthLabel}</td>`;

    salesTypes.forEach((typeDef) => {
      const months = getMonths(org, typeDef.key);
      if (!months.length) {
        row.innerHTML += "<td>-</td><td>-</td><td class=\"rate-cell\">-</td>";
        return;
      }
      const month = months[index];
      const rate = ratio(month.actual, month.target).toFixed(1);
      row.innerHTML += `
        <td class="amount-cell">${yen.format(month.target)}</td>
        <td class="amount-cell">${yen.format(month.actual)}</td>
        <td class="rate-cell">${rate}%</td>
      `;
    });

    monthlyTable.appendChild(row);
  });
}

function renderOrgSummaryTable() {
  orgSummaryTable.innerHTML = "";
  const hqTotals = {};
  const companyTotals = { targetTotal: 0, actualTotal: 0 };

  Object.values(baseOrgData).forEach((org) => {
    const totals = annualTotals(getMonths(org, "total"));
    if (!hqTotals[org.hq]) hqTotals[org.hq] = { targetTotal: 0, actualTotal: 0 };
    hqTotals[org.hq].targetTotal += totals.targetTotal;
    hqTotals[org.hq].actualTotal += totals.actualTotal;
    companyTotals.targetTotal += totals.targetTotal;
    companyTotals.actualTotal += totals.actualTotal;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${org.hq}</td>
      <td>${org.unit}</td>
      <td>${yen.format(totals.targetTotal)}</td>
      <td>${yen.format(totals.actualTotal)}</td>
      <td class="rate-cell">${ratio(totals.actualTotal, totals.targetTotal).toFixed(1)}%</td>
    `;
    orgSummaryTable.appendChild(tr);
  });

  Object.entries(hqTotals).forEach(([hq, totals]) => {
    const tr = document.createElement("tr");
    tr.className = "hq-total-row";
    tr.innerHTML = `
      <td>${hq}</td>
      <td>${hq}合計</td>
      <td>${yen.format(totals.targetTotal)}</td>
      <td>${yen.format(totals.actualTotal)}</td>
      <td class="rate-cell">${ratio(totals.actualTotal, totals.targetTotal).toFixed(1)}%</td>
    `;
    orgSummaryTable.appendChild(tr);
  });

  const companyRow = document.createElement("tr");
  companyRow.className = "company-total-row";
  companyRow.innerHTML = `
    <td>全社</td>
    <td>全社合計</td>
    <td>${yen.format(companyTotals.targetTotal)}</td>
    <td>${yen.format(companyTotals.actualTotal)}</td>
    <td class="rate-cell">${ratio(companyTotals.actualTotal, companyTotals.targetTotal).toFixed(1)}%</td>
  `;
  orgSummaryTable.appendChild(companyRow);
}

function renderTotalSummary() {
  const org = currentOrg();
  const months = getMonths(org, "total");
  const idx = currentMonthIndex();
  const year = annualTotals(months);

  const yearPct = ratio(year.actualTotal, year.targetTotal);
  yearRateEl.textContent = `${yearPct.toFixed(1)}%`;
  yearBarEl.style.width = `${clampPercent(yearPct)}%`;

  const remaining = Math.max(year.targetTotal - year.actualTotal, 0);
  yearRemainingEl.textContent = yen.format(remaining);

  const monthsLeft = 12 - (idx + 1);
  const needPerMonth = monthsLeft > 0 ? remaining / monthsLeft : remaining;
  needPerMonthEl.textContent = yen.format(needPerMonth);
  needPerMonthEl.classList.toggle("negative", remaining > 0 && monthsLeft === 0);
}

function drawChart() {
  const org = currentOrg();
  const typeDefs = availableTypeDefs(org);
  const series = typeDefs.map((def) => ({
    ...def,
    targets: getMonths(org, def.key).map((m) => m.target),
    actuals: getMonths(org, def.key).map((m) => m.actual),
  }));

  const labels = currentMonthNames();
  const padding = { top: 52, right: 28, bottom: 44, left: 64 };
  const width = chart.width;
  const height = chart.height;
  const areaW = width - padding.left - padding.right;
  const areaH = height - padding.top - padding.bottom;

  const allValues = series.flatMap((s) => [...s.targets, ...s.actuals]);
  const maxY = Math.max(...allValues, 1000000);
  const ticks = 5;

  ctx.clearRect(0, 0, width, height);
  ctx.font = "12px sans-serif";
  ctx.strokeStyle = "#d6e3e2";
  ctx.fillStyle = "#6e8088";

  for (let i = 0; i <= ticks; i += 1) {
    const y = padding.top + (areaH * i) / ticks;
    const value = maxY - (maxY * i) / ticks;
    ctx.beginPath();
    ctx.moveTo(padding.left, y);
    ctx.lineTo(width - padding.right, y);
    ctx.stroke();
    ctx.fillText(`${Math.round(value / 10000)}万`, 6, y + 4);
  }

  const stepX = areaW / (labels.length - 1);

  function drawLine(values, color, dashed) {
    ctx.beginPath();
    values.forEach((v, i) => {
      const x = padding.left + stepX * i;
      const y = padding.top + areaH - (v / maxY) * areaH;
      if (i === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });

    ctx.setLineDash(dashed ? [5, 4] : []);
    ctx.strokeStyle = color;
    ctx.lineWidth = 2;
    ctx.stroke();
    ctx.setLineDash([]);
  }

  series.forEach((s) => {
    drawLine(s.targets, s.targetColor, true);
    drawLine(s.actuals, s.actualColor, false);
  });

  ctx.fillStyle = "#23353a";
  labels.forEach((label, i) => {
    const x = padding.left + stepX * i;
    ctx.fillText(label, x - 18, height - 12);
  });

  let legendX = width - 320;
  series.forEach((s) => {
    ctx.fillStyle = s.targetColor;
    ctx.fillRect(legendX, 14, 10, 10);
    ctx.fillStyle = "#23353a";
    ctx.fillText(`${s.label} 目標`, legendX + 14, 23);
    legendX += 100;
  });

  legendX = width - 320;
  series.forEach((s) => {
    ctx.fillStyle = s.actualColor;
    ctx.fillRect(legendX, 28, 10, 10);
    ctx.fillStyle = "#23353a";
    ctx.fillText(`${s.label} 実績`, legendX + 14, 37);
    legendX += 100;
  });
}

function recalcAndRender() {
  renderTermSelect();
  renderOrgSelect();
  renderOrgMeta();
  renderBreakdownTables();
  renderMonthlyRows();
  renderTotalSummary();
  drawChart();
  renderOrgSummaryTable();
}

termSelect.addEventListener("change", () => {
  state.currentTermKey = termSelect.value;
  rebuildTermContext();
  recalcAndRender();
});

orgSelect.addEventListener("change", () => {
  state.currentOrgKey = orgSelect.value;
  recalcAndRender();
});

rebuildTermContext();
recalcAndRender();
