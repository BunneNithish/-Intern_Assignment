/***************************************
 * CONFIG
 ***************************************/
let selectedDataView = "Orders";      // Orders | Sessions | Calls
let selectedGranularity = "Daily";    // Daily | Weekly | Monthly

/***************************************
 * FILE UPLOAD HANDLER
 ***************************************/
document.getElementById("upload").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    console.log("Sheet Names:", workbook.SheetNames);

    // Load required sheets
    const ordersRaw = XLSX.utils.sheet_to_json(workbook.Sheets["Orders_Raw"]);
    const callsRaw  = XLSX.utils.sheet_to_json(workbook.Sheets["Calls_Raw"]);

    console.log("Orders Raw:", ordersRaw);
    console.log("Calls Raw:", callsRaw);

    // TASK 1 – Aggregate Orders
    const ordersAgg = aggregateData(
      ordersRaw,
      "Orders",
      selectedGranularity
    );

    console.log("Orders Aggregated:", ordersAgg);

    // TASK 2 – Chart
    renderChart(ordersAgg);

    // TASK 3 – Orders ↔ Calls Matching
    const callsAgg = aggregateData(
      callsRaw,
      "Calls",
      selectedGranularity
    );

    const derivedMetrics = matchOrdersWithCalls(ordersAgg, callsAgg);
    console.log("Derived Metrics (Orders vs Calls):", derivedMetrics);
  };

  reader.readAsArrayBuffer(file);
});

/***************************************
 * AGGREGATION LOGIC (TASK 1)
 ***************************************/
function aggregateData(data, dataView, granularity) {
  const result = {};

  data.forEach(row => {
    const entityId = getEntityId(row, dataView);
    const dateObj = getDateValue(row, dataView);
    const timeKey = getTimeKey(dateObj, granularity);

    if (!entityId || !timeKey) return;

    const key = `${entityId}_${timeKey}`;

    if (!result[key]) {
      result[key] = {
        entityId,
        timeKey,
        count: 0
      };
    }

    result[key].count += 1;
  });

  return Object.values(result);
}

/***************************************
 * ENTITY IDENTIFICATION
 ***************************************/
function getEntityId(row, dataView) {
  if (dataView === "Orders") {
    return row["Phone"];
  }
  if (dataView === "Calls") {
    return row["Phone"];
  }
  return null;
}

/***************************************
 * DATE EXTRACTION (EXCEL SERIAL → JS DATE)
 ***************************************/
function getDateValue(row, dataView) {
  let excelDate;

  if (dataView === "Orders") {
    excelDate = row["Order Date"];
  } else if (dataView === "Calls") {
    excelDate = row["Call Date"];
  }

  if (!excelDate) return null;

  return new Date((excelDate - 25569) * 86400 * 1000);
}

/***************************************
 * TIME KEY GENERATION
 ***************************************/
function getTimeKey(date, granularity) {
  if (!(date instanceof Date) || isNaN(date)) return null;

  if (granularity === "Daily") {
    return date.toISOString().slice(0, 10);
  }

  if (granularity === "Monthly") {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  }

  if (granularity === "Weekly") {
    return getISOWeek(date);
  }

  return null;
}

/***************************************
 * ISO WEEK CALCULATION
 ***************************************/
function getISOWeek(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));

  const weekYear = d.getFullYear();
  const week1 = new Date(weekYear, 0, 4);

  const weekNumber =
    1 +
    Math.round(
      ((d - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7
    );

  return `${weekYear}-W${String(weekNumber).padStart(2, "0")}`;
}

/***************************************
 * TASK 2 – CHART LOGIC
 ***************************************/
function groupByTime(aggregatedData) {
  const map = {};

  aggregatedData.forEach(item => {
    map[item.timeKey] = (map[item.timeKey] || 0) + item.count;
  });

  return {
    labels: Object.keys(map),
    values: Object.values(map)
  };
}

let chartInstance = null;

function renderChart(aggregatedData) {
  const ctx = document.getElementById("chart").getContext("2d");
  const { labels, values } = groupByTime(aggregatedData);

  if (chartInstance) chartInstance.destroy();

  chartInstance = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: `Orders (${selectedGranularity})`,
        data: values
      }]
    }
  });
}

/***************************************
 * TASK 3 – CROSS TABLE MATCHING
 ***************************************/
function matchOrdersWithCalls(ordersAgg, callsAgg) {
  const callsMap = {};

  callsAgg.forEach(c => {
    callsMap[`${c.entityId}_${c.timeKey}`] = c.count;
  });

  return ordersAgg.map(o => {
    const key = `${o.entityId}_${o.timeKey}`;
    const calls = callsMap[key] || 0;

    return {
      entityId: o.entityId,
      timeKey: o.timeKey,
      orders: o.count,
      calls,
      ordersPerCall: calls === 0 ? 0 : (o.count / calls)
    };
  });
}
