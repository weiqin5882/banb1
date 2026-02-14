const officialFileInput = document.getElementById("officialFile");
const serviceFileInput = document.getElementById("serviceFile");
const defaultCostInput = document.getElementById("defaultCost");
const analyzeBtn = document.getElementById("analyzeBtn");
const exportBtn = document.getElementById("exportBtn");
const summaryEl = document.getElementById("summary");
const resultBody = document.querySelector("#resultTable tbody");

let latestResultRows = [];
let latestSummary = null;

const headerAliases = {
  orderNo: ["快手订单编号", "订单号", "订单编号", "订单id", "订单ID", "交易单号", "订单流水号"],
  status: ["订单状态", "状态", "交易状态", "发货状态"],
  productName: ["订单商品名称", "商品名称", "产品名称", "商品", "sku名称"],
  revenue: ["商家实收", "支付金额", "订单金额", "销售金额", "销售额", "应收金额", "实付金额", "金额"],
  cost: ["成本", "产品成本", "采购价", "采购成本", "货品成本"],
};

const validStatusKeywords = ["交易成功", "已发货"];

analyzeBtn.addEventListener("click", async () => {
  try {
    const officialFile = officialFileInput.files[0];
    const serviceFile = serviceFileInput.files[0];

    if (!officialFile || !serviceFile) {
      alert("请先选择官方订单表和客服统计表");
      return;
    }

    const [officialRows, serviceRows] = await Promise.all([
      parseFile(officialFile),
      parseFile(serviceFile),
    ]);

    const defaultCost = Number(defaultCostInput.value || 0);
    const officialNormalized = normalizeRows(officialRows, "官方表", defaultCost);
    const serviceNormalized = normalizeRows(serviceRows, "客服表", defaultCost);

    const officialFiltered = filterByStatus(officialNormalized.rows, officialNormalized.hasStatusColumn, "官方表");
    const serviceFiltered = filterByStatus(serviceNormalized.rows, serviceNormalized.hasStatusColumn, "客服表");

    const officialMap = toOrderMap(officialFiltered);
    const serviceMap = toOrderMap(serviceFiltered);

    const allOrderNos = Array.from(new Set([...officialMap.keys(), ...serviceMap.keys()]));

    latestResultRows = allOrderNos.map((orderNo, index) => {
      const official = officialMap.get(orderNo);
      const service = serviceMap.get(orderNo);
      const source = official && service ? "官方+客服" : official ? "官方缺客服" : "客服缺官方";

      const record = official || service;
      const revenue = safeNum(record?.revenue);
      const cost = safeNum(record?.cost);
      const profit = revenue - cost;

      return {
        serial: index + 1,
        orderNo,
        source,
        status: official?.status || service?.status || "",
        productName: mergeProductName(official?.productName, service?.productName),
        revenue,
        cost,
        profit,
        matched: official && service ? "已匹配" : "有漏单",
        isLoss: profit < 0,
        isMissing: !(official && service),
      };
    });

    latestSummary = calcSummary(latestResultRows);
    renderSummary(latestSummary);
    renderTable(latestResultRows);

    exportBtn.disabled = false;
  } catch (err) {
    console.error(err);
    alert(`处理失败：${err.message}`);
  }
});

exportBtn.addEventListener("click", () => {
  if (!latestResultRows.length) return;

  const rows = latestResultRows.map((item) => ({
    类序号: item.serial,
    订单号: item.orderNo,
    来源: item.source,
    状态: item.status,
    产品名称: item.productName,
    销售额: item.revenue,
    成本: item.cost,
    利润: item.profit,
    匹配情况: item.matched,
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "订单比对");

  const red = { font: { color: { rgb: "9C0006" } }, fill: { fgColor: { rgb: "FFC7CE" } } };
  latestResultRows.forEach((r, i) => {
    if (r.isLoss) {
      const rowNum = i + 2;
      ["A", "B", "C", "D", "E", "F", "G", "H", "I"].forEach((col) => {
        const cell = ws[`${col}${rowNum}`];
        if (!cell) return;
        cell.s = { ...(cell.s || {}), ...red };
      });
    }
  });

  ws["!cols"] = [
    { wch: 8 },
    { wch: 22 },
    { wch: 12 },
    { wch: 12 },
    { wch: 36 },
    { wch: 12 },
    { wch: 12 },
    { wch: 12 },
    { wch: 12 },
  ];

  XLSX.writeFile(wb, "订单比对结果.xlsx");
});

async function parseFile(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array" });
  const first = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(first, { defval: "" });
}

function normalizeRows(rows, sourceName, defaultCost) {
  if (!rows.length) {
    return { rows: [], hasStatusColumn: false };
  }

  const keys = Object.keys(rows[0]);
  const mapping = {
    orderNo: findHeader(keys, headerAliases.orderNo),
    status: findHeader(keys, headerAliases.status),
    productName: findHeader(keys, headerAliases.productName),
    revenue: findHeader(keys, headerAliases.revenue),
    cost: findHeader(keys, headerAliases.cost),
  };

  if (!mapping.orderNo) {
    throw new Error(`${sourceName}无法识别“订单号”列，请检查表头名称`);
  }

  const normalizedRows = rows
    .map((row) => {
      const orderNo = String(row[mapping.orderNo] ?? "").trim();
      if (!orderNo) return null;
      const status = mapping.status ? String(row[mapping.status] ?? "").trim() : "";
      return {
        orderNo,
        status,
        productName: String(row[mapping.productName] ?? "").trim(),
        revenue: safeNum(row[mapping.revenue]),
        cost: mapping.cost ? safeNum(row[mapping.cost]) : defaultCost,
      };
    })
    .filter(Boolean);

  return {
    rows: normalizedRows,
    hasStatusColumn: Boolean(mapping.status),
  };
}

function filterByStatus(rows, hasStatusColumn, sourceName) {
  if (!hasStatusColumn) {
    console.warn(`${sourceName}未识别到状态列，默认不过滤状态，使用全部订单参与比对。`);
    return rows;
  }
  return rows.filter((row) => validStatusKeywords.some((k) => row.status.includes(k)));
}

function toOrderMap(rows) {
  const map = new Map();
  rows.forEach((row) => {
    if (!map.has(row.orderNo)) {
      map.set(row.orderNo, row);
    }
  });
  return map;
}

function findHeader(keys, aliases) {
  const lowerKeys = keys.map((k) => ({ raw: k, normalized: normalizeText(k) }));

  for (const alias of aliases) {
    const target = normalizeText(alias);
    const exact = lowerKeys.find((k) => k.normalized === target);
    if (exact) return exact.raw;
  }

  for (const alias of aliases) {
    const target = normalizeText(alias);
    const partial = lowerKeys.find((k) => k.normalized.includes(target) || target.includes(k.normalized));
    if (partial) return partial.raw;
  }

  return "";
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .toLowerCase();
}

function safeNum(val) {
  if (typeof val === "number") return val;
  const cleaned = String(val || "")
    .replace(/[￥,\s]/g, "")
    .trim();
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function mergeProductName(a, b) {
  if (a && b && a !== b) return `${a} / ${b}`;
  return a || b || "";
}

function calcSummary(rows) {
  return {
    compared: rows.length,
    matched: rows.filter((r) => r.matched === "已匹配").length,
    missing: rows.filter((r) => r.isMissing).length,
    totalRevenue: sum(rows.map((r) => r.revenue)),
    totalCost: sum(rows.map((r) => r.cost)),
    totalProfit: sum(rows.map((r) => r.profit)),
    lossOrders: rows.filter((r) => r.isLoss).length,
  };
}

function renderSummary(summary) {
  summaryEl.innerHTML = `
    <div class="summary-grid">
      ${card("参与比对订单", summary.compared)}
      ${card("匹配成功", summary.matched)}
      ${card("漏单", summary.missing)}
      ${card("销售额合计", formatMoney(summary.totalRevenue))}
      ${card("成本合计", formatMoney(summary.totalCost))}
      ${card("总利润", formatMoney(summary.totalProfit))}
      ${card("亏损订单数", summary.lossOrders)}
    </div>
  `;
}

function renderTable(rows) {
  resultBody.innerHTML = "";
  rows.forEach((item) => {
    const tr = document.createElement("tr");
    if (item.isLoss) tr.classList.add("loss-row");
    if (item.isMissing) tr.classList.add("missing-row");

    tr.innerHTML = `
      <td>${item.serial}</td>
      <td>${item.orderNo}</td>
      <td>${item.source}</td>
      <td>${item.status}</td>
      <td>${item.productName}</td>
      <td>${formatMoney(item.revenue)}</td>
      <td>${formatMoney(item.cost)}</td>
      <td>${formatMoney(item.profit)}</td>
      <td>${item.matched}</td>
    `;

    resultBody.appendChild(tr);
  });
}

function card(label, value) {
  return `<div class="card"><div class="label">${label}</div><div class="value">${value}</div></div>`;
}

function sum(arr) {
  return arr.reduce((acc, n) => acc + safeNum(n), 0);
}

function formatMoney(n) {
  return `¥${safeNum(n).toFixed(2)}`;
}
