const express = require("express");
const path = require("path");
const axios = require("axios");
require("dotenv").config();
const { ConfidentialClientApplication } = require("@azure/msal-node");

const app = express();
app.use(express.json());
const PORT = process.env.PORT || 3000;

const formatDate = (value) => {
  if (value === undefined || value === null || value === "" || value === "-")
    return "";

  let raw = value;
  if (typeof raw === "string") raw = raw.trim();

  if (typeof raw === "number" || (!isNaN(raw) && raw !== "")) {
    const num = Number(raw);
    if (num > 0) {
      const date = new Date(Math.round((num - 25569) * 86400 * 1000));
      if (!isNaN(date.getTime())) {
        return date.toLocaleDateString("en-GB", {
          day: "2-digit",
          month: "2-digit",
          year: "numeric",
        });
      }
    }
  }

  if (typeof raw === "string" && (raw.includes("/") || raw.includes("-"))) {
    const parts = raw.replace(/\s+/g, "").split(/[-\/]/);
    if (parts.length === 3) {
      let [p1, p2, p3] = parts;
      let day;
      let month;
      let year;
      if (p1.length === 4) {
        year = p1;
        month = p2;
        day = p3;
      } else {
        day = p1;
        month = p2;
        year = p3;
      }
      if (year.length === 2) year = "20" + year;
      const parsed = new Date(Number(year), Number(month) - 1, Number(day));
      if (!isNaN(parsed.getTime())) {
        return parsed.toLocaleDateString("en-GB", {
          day: "2-digit",
          month: "2-digit",
          year: "numeric",
        });
      }
    }
  }

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return parsed.toLocaleDateString("en-GB", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    });
  }

  return "-";
};

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

const cca = new ConfidentialClientApplication(msalConfig);

async function getAccessToken() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return result.accessToken;
}

async function getSharePointFileContext() {
  const token = await getAccessToken();
  const headers = { Authorization: `Bearer ${token}` };

  const siteRes = await axios.get(
    "https://graph.microsoft.com/v1.0/sites/mlmpackagingmy.sharepoint.com:/sites/FileStorage",
    { headers },
  );
  const siteId = siteRes.data.id;

  const drivesRes = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers },
  );
  const driveId = drivesRes.data.value[0]?.id;
  if (!driveId) throw new Error("Could not find driveId in site drives");

  const fileRes = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/Database.xlsx`,
    { headers },
  );
  const fileId = fileRes.data.id;
  if (!fileId) throw new Error("Could not find fileId for Database.xlsx");

  return { token, headers, siteId, driveId, fileId };
}

async function getTableRowsAsObjects(tableName) {
  const ctx = await getSharePointFileContext();

  // console.log(`[getTableRowsAsObjects] Fetching ${tableName} from SharePoint`);

  const headerRes = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables/${tableName}/headerRowRange`,
    { headers: ctx.headers },
  );
  const headers = headerRes.data.values?.[0] || [];
  // console.log(`[getTableRowsAsObjects] ${tableName} headers:`, headers);
  // console.log(`[getTableRowsAsObjects] ${tableName} header count:`, headers.length);

  const rowsRes = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables/${tableName}/rows`,
    { headers: ctx.headers },
  );
  const rows = rowsRes.data.value || [];
  // console.log(`[getTableRowsAsObjects] ${tableName} has ${rows.length} rows`);
  if (rows.length > 0) {
    // console.log(`[getTableRowsAsObjects] First row structure:`, rows[0]);
    // console.log(`[getTableRowsAsObjects] First row values:`, rows[0].values);
  }

  const result = rows.map((row, idx) => {
    // Handle different response formats from Graph API
    let values = [];

    if (row.values) {
      if (Array.isArray(row.values)) {
        // If values is array of arrays, get first element; if array of values, use directly
        values = Array.isArray(row.values[0]) ? row.values[0] : row.values;
      } else if (typeof row.values === "string") {
        // If values is a single string, split by comma (CSV format)
        values = row.values.split(",");
        // console.log(`[getTableRowsAsObjects] ${tableName} row ${idx} split by comma:`, values);
      }
    }

    // console.log(`[getTableRowsAsObjects] ${tableName} row ${idx} final values:`, values);

    const item = {};
    headers.forEach((h, idx) => {
      item[h] = values[idx] !== undefined ? values[idx] : null;
    });
    return item;
  });

  console.log(
    `[getTableRowsAsObjects] Converted ${tableName} to objects:`,
    result.slice(0, 2),
  );
  return result;
}

async function getTableColumnCount(tableName) {
  const ctx = await getSharePointFileContext();

  try {
    // First try direct table columns count (best fit for table metadata)
    const colsRes = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables/${tableName}/columns`,
      { headers: ctx.headers },
    );
    const columns = colsRes.data.value || [];
    if (columns.length > 0) {
      // console.log(`[getTableColumnCount] ${tableName} columns via /columns:`, columns.length);
      return columns.length;
    }

    // Fallback to headerRowRange if columns endpoint gives empty
    const headerRes = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables/${tableName}/headerRowRange`,
      { headers: ctx.headers },
    );
    const headers = headerRes.data.values?.[0] || [];
    // console.log(`[getTableColumnCount] ${tableName} columns via headerRowRange:`, headers.length);
    return headers.length;
  } catch (error) {
    // console.error(`[getTableColumnCount] Error fetching columns for ${tableName}:`, error.message);
    throw error;
  }
}

async function addTableRow(tableName, values) {
  const ctx = await getSharePointFileContext();
  console.log(
    `[addTableRow] Adding row to ${tableName}:`,
    values.length,
    "columns",
  );

  const resp = await axios.post(
    `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables/${tableName}/rows/add`,
    { values: [values] },
    { headers: ctx.headers },
  );

  // console.log(`[addTableRow] Successfully added to ${tableName}`);
  return resp.data;
}

async function generateDashboardData() {
  // console.log('[generateDashboardData] Starting...');

  // Get file context once
  const fileContext = await getSharePointFileContext();

  // Fetch all rows for BatchListing with numeric index preservation
  const batchRowRes = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${fileContext.driveId}/items/${fileContext.fileId}/workbook/tables('BatchListing')/rows`,
    { headers: fileContext.headers },
  );
  const batchRowsRaw = batchRowRes.data.value || [];

  // console.log('[generateDashboardData] Batch rows count:', batchRowsRaw.length);
  // if (batchRowsRaw.length > 0) {
  //   console.log('[generateDashboardData] First batch row values:', batchRowsRaw[0].values);
  // }

  // Convert to objects for JobListing
  const jobListing = await getTableRowsAsObjects("JobListing");

  // console.log('[generateDashboardData] JobListing:', jobListing);

  // Create jobInfoMap from JobListing for quick lookup
  const jobInfoMap = {};
  jobListing.forEach((job) => {
    const psn = String(job["Product Serial Number"] || job.psn || "").trim();
    if (!psn) return;
    if (!jobInfoMap[psn]) {
      jobInfoMap[psn] = {
        pi: job["PI Number"] || job.PI_Number || job.pi || job.PI_No || "-",
        code: job.salesCode || job.Sales_Code || job.code || "-",
        client: job["Job Name"] || job.Job_Name || job.client || "-",
        orderDate: formatDate(
          job["Order Date"] || job.Order_Date || job.orderDate,
        ),
        priority: job.Priority || "NORMAL",
        deliveryDate: formatDate(
          job["Delivery Date"] || job.Delivery_Date || job.deliveryDate,
        ),
        status: job.Status || "ON SCHEDULE",
      };
    }
  });
  // console.log(jobInfoMap[0]);
  // console.log('[generateDashboardData] jobInfoMap:', jobInfoMap);

  // Process BatchListing to create batches with steps structure
  const batches = [];

  // Use raw rows with numeric indices for step parsing (if available)
  batchRowsRaw.forEach((rawRow, rowIdx) => {
    // Handle different response formats for raw rows
    let values = [];
    if (rawRow.values) {
      if (Array.isArray(rawRow.values)) {
        values = Array.isArray(rawRow.values[0])
          ? rawRow.values[0]
          : rawRow.values;
      } else if (typeof rawRow.values === "string") {
        // If values is a single string, split by comma (CSV format)
        values = rawRow.values.split(",");
        // console.log(`[generateDashboardData] Batch row ${rowIdx} split by comma:`, values);
      }
    }

    // Extract basic info from numeric indices (matching app.js)
    const psn = String(values[0] || "").trim();
    if (!psn || psn.toLowerCase().includes("done")) {
      // console.log('[generateDashboardData] Skipping batch row', rowIdx, '- no PSN or is done');
      return;
    }

    // Parse 12 process steps from columns 6+ (each step is 9 columns)
    let definedSteps = [];
    let ticksFound = 0;
    let activeStepStatus = "";

    const START_COL = 6;
    const BLOCK_SIZE = 9;

    for (let s = 0; s < 12; s++) {
      let base = START_COL + s * BLOCK_SIZE;
      if (base >= values.length) break;

      let pName = String(values[base] || "").trim();

      if (pName !== "" && pName !== "--") {
        let rawTick = values[base + 8];
        let isDone =
          rawTick === true ||
          rawTick === 1 ||
          String(rawTick).toUpperCase() === "TRUE";

        if (isDone) ticksFound++;
        if (!activeStepStatus && !isDone) {
          activeStepStatus = String(values[base + 5] || "").trim();
        }

        // Helper to safely parse dates
        const parseDateValue = (val) => {
          const formatted = formatDate(val);
          if (formatted === "-") return "-";
          return formatted;
        };

        definedSteps.push({
          name: pName,
          expDate: parseDateValue(values[base + 1]) || "-",
          rawExpDate:
            typeof values[base + 1] === "number"
              ? (values[base + 1] - 25569) * 86400000
              : null,
          endDate: parseDateValue(values[base + 2]) || "-",
          duration: values[base + 3] || 0,
          detail: String(values[base + 4] || ""),
          status: String(values[base + 5] || ""),
          remark: String(values[base + 6] || ""),
          revertRemark: String(values[base + 7] || ""),
          isDone: isDone,
          baseCol: base,
        });
      }
    }

    let batchProgress =
      definedSteps.length > 0 ? ticksFound / definedSteps.length : 0;

    const info = jobInfoMap[psn] || {
      pi: "-",
      code: "-",
      client: "-",
      orderDate: "-",
      priority: "NORMAL",
      deliveryDate: "-",
      status: "ON SCHEDULE",
    };

    batches.push({
      row: rowIdx,
      psn: psn,
      batchId: String(values[1] || ""),
      batchDate: values[2]
        ? new Date(values[2]).toLocaleDateString("en-GB", {
            year: "2-digit",
            month: "2-digit",
            day: "2-digit",
          })
        : "-",
      jobName: String(values[3] || ""),
      qty: Number(values[4] || 0),
      progress: definedSteps.length > 0 ? ticksFound / definedSteps.length : 0,
      steps: definedSteps,
      activeStepStatus: activeStepStatus,
      piNumber: info.pi,
      salesCode: info.code,
      clientName: info.client,
      orderDate: info.orderDate,
      priority: info.priority,
      deliveryDate: info.deliveryDate,
      status: values[10] || info.status,
      qtyString: String(values[119] || ""),
    });
  });

  // console.log('[generateDashboardData] Processed batches:', batches);

  // Calculate process averages for workload display
  const processList = [
    "Sheeting",
    "Printing",
    "Lamination",
    "Efluting",
    "Die-cut",
    "Convert",
    "Baseboard",
    "Hotstamping",
    "Packing",
    "Double-Side-Tape",
    "Emboss",
    "Blind Emboss",
    "Gluing",
    "Side Glue",
    "2 Point Glue",
    "Attach Handle",
    "Peeling",
    "Punch Hole",
    "Slitting LF",
    "FlexoSlitting",
    "Spot UV",
    "Texture",
    "Trimming",
    "Varnish",
    "Waterbase",
    "Delivery",
  ];
  const processStats = {};

  // Initialize stats for each process
  processList.forEach((processName) => {
    processStats[processName] = {
      name: processName,
      totalTime: 0,
      count: 0,
      activeCount: 0,
    };
  });

  // Calculate stats from batch steps - ONLY current active step per batch
  batches.forEach((batch) => {
    if (batch.steps && Array.isArray(batch.steps)) {
      const currentStep = batch.steps.find((step) => !step.isDone);
      if (currentStep && processStats[currentStep.name]) {
        const durationDays = Math.max(
          1,
          Math.ceil(Number(currentStep.duration) || 0),
        );
        processStats[currentStep.name].totalTime += durationDays;
        processStats[currentStep.name].count += 1;
        processStats[currentStep.name].activeCount += 1;
      }
    }
  });

  // Calculate averages - CEIL the average time (0.5 = 1 day, 1.5 = 2 days)
  const averages = Object.values(processStats).map((stat) => ({
    name: stat.name,
    avgTime: stat.count > 0 ? Math.ceil(stat.totalTime / stat.count) : 0,
    activeCount: stat.activeCount,
  }));

  // Generate rawCapacity data from COMPLETED batches (isDone = true)
  const machineCapacityMap = {};

  batches.forEach((batch) => {
    if (batch.steps && Array.isArray(batch.steps)) {
      batch.steps.forEach((step) => {
        // Only count COMPLETED steps (isDone = true)
        if (step.isDone && step.endDate && step.endDate !== "-") {
          const machine = step.detail || "GENERAL";
          const dateKey = step.endDate; // Format: D/M/YYYY
          const key = `${dateKey}|${machine}`;

          if (!machineCapacityMap[key]) {
            machineCapacityMap[key] = {
              date: dateKey,
              machine:
                machine === "Ijima" || machine.toLowerCase().includes("ijima")
                  ? "IJIMA"
                  : machine === "Hand" || machine.toLowerCase().includes("hand")
                    ? "HANDSWITCH"
                    : machine === "Out" || machine.toLowerCase().includes("out")
                      ? "OUTSOURCED"
                      : "GENERAL",
              qty: 0,
            };
          }
          machineCapacityMap[key].qty += batch.qty;
        }
      });
    }
  });

  const rawCapacity = Object.values(machineCapacityMap);

  const totalJobs = jobListing.length;
  const totalBatches = batches.length;
  const totalQuantity = batches.reduce(
    (acc, b) => acc + (Number(b.qty) || 0),
    0,
  );

  const completedBatches = batches.filter((b) => {
    const status = String(b.status || "").toLowerCase();
    return status === "completed" || status === "done";
  }).length;

  const result = {
    jobs: batches,
    jobListing: batches, // For compatibility with both old and new frontend
    batchListing: batches, // Also return as batchListing for backward compat
    averages: averages, // For process workload display
    rawCapacity: rawCapacity, // For machine capacity display
    stats: {
      totalJobs,
      totalBatches,
      totalQuantity,
      completedBatches,
      pendingBatches: totalBatches - completedBatches,
    },
  };

  // console.log('[generateDashboardData] Result:', result);
  return result;
}

app.get("/", (req, res) => {
  console.log("[APP] GET / - redirecting to dashboard");
  res.redirect("/dashboard.html");
});
app.use(express.static(path.join(__dirname, "public")));

app.get("/api/test", (req, res) => {
  console.log("[API] GET /api/test");
  res.json({ message: "Node.js backend working" });
});

app.get("/api/getToken", async (req, res) => {
  try {
    console.log("[API] GET /api/getToken - attempting to get token");
    const token = await getAccessToken();
    console.log("[API] GET /api/getToken - success");
    res.json({ accessToken: token });
  } catch (error) {
    console.error("[API] GET /api/getToken - error:", error);
    res.status(500).json({ error: "Failed to get token" });
  }
});

app.get("/api/jobListing", async (req, res) => {
  try {
    console.log("[API] GET /api/jobListing");
    const data = await getTableRowsAsObjects("JobListing");
    res.json(data);
  } catch (error) {
    console.error("[API] GET /api/jobListing - error:", error);
    res.status(500).json({ error: "Failed to load JobListing" });
  }
});

app.get("/api/batchListing", async (req, res) => {
  try {
    console.log("[API] GET /api/batchListing");
    const data = await getTableRowsAsObjects("BatchListing");
    res.json(data);
  } catch (error) {
    console.error("[API] GET /api/batchListing - error:", error);
    res.status(500).json({ error: "Failed to load BatchListing" });
  }
});

app.get("/api/dashboard", async (req, res) => {
  try {
    console.log("[API] GET /api/dashboard - request received");
    const data = await generateDashboardData();
    console.log("[API] GET /api/dashboard - sending response");
    res.json(data);
  } catch (error) {
    console.error("[API] GET /api/dashboard - error:", error);
    res.status(500).json({ error: "Failed to load dashboard data" });
  }
});

app.post("/api/submitData", async (req, res) => {
  try {
    console.log("🚀 API HIT: /api/submitData");

    const data = req.body;
    console.log("📥 Incoming data:", JSON.stringify(data, null, 2));

    let jobListingColumnCount = null;
    let batchListingColumnCount = null;
    let jobRowLength = null;
    let batchRowLength = null;

    // ✅ STEP 1: Validate input
    if (!data || !data.psn || !data.batches) {
      console.error("❌ Validation failed:", data);
      return res.status(400).json({ error: "Invalid submit payload" });
    }
    console.log("✅ Step 1: Validation passed");

    // ✅ STEP 2: Check batches array
    if (!Array.isArray(data.batches)) {
      console.error("❌ batches is not an array:", data.batches);
      return res.status(400).json({ error: "batches must be an array" });
    }
    console.log("✅ Step 2: batches is valid array");

    // ✅ STEP 3: Get existing rows
    console.log("⏳ Fetching existing JobListing...");
    const existing = await getTableRowsAsObjects("JobListing");
    console.log("📊 Existing rows fetched:", existing.length);

    // Debug PSN values
    console.log(
      "🔍 Existing PSNs:",
      existing.map((j) => j.PSN || j.psn),
    );

    // ✅ STEP 4: Duplicate check
    if (
      existing.some(
        (job) =>
          String(job["Product Serial Number"] || job.PSN || job.psn).trim() ===
          String(data.psn).trim(),
      )
    ) {
      console.error("❌ Duplicate PSN found:", data.psn);
      return res.status(409).json({ error: "Duplicate PSN" });
    }
    console.log("✅ Step 4: No duplicate found");

    // ✅ STEP 5: Format dates
    console.log("⏳ Formatting dates...");
    const normalizedOrderDate = formatDate(data.orderDate);
    const normalizedDeliveryDate = formatDate(data.deliveryDate);
    console.log("📅 Formatted dates:", {
      order: normalizedOrderDate,
      delivery: normalizedDeliveryDate,
    });

    // ✅ STEP 6: Prepare job row
    const jobRow = [
      data.psn,
      data.piNumber || "",
      data.salesCode || "",
      data.jobName || "",
      data.jobType || "",
      data.quantity || 0,
      normalizedOrderDate !== "-" ? normalizedOrderDate : "",
      normalizedDeliveryDate !== "-" ? normalizedDeliveryDate : "",
      data.item || "",
      data.priority || "",
      data.status || "ON SCHEDULE", // Default status if not provided
    ];

    // match table columns exactly for JobListing
    jobListingColumnCount = await getTableColumnCount("JobListing");
    if (jobListingColumnCount <= 0) {
      throw new Error(
        `JobListing column count invalid: ${jobListingColumnCount}`,
      );
    }

    if (jobRow.length < jobListingColumnCount) {
      while (jobRow.length < jobListingColumnCount) jobRow.push("");
    } else if (jobRow.length > jobListingColumnCount) {
      jobRow.splice(jobListingColumnCount);
    }

    jobRowLength = jobRow.length;
    console.log(
      "📦 Job row to insert:",
      jobRow,
      `length=${jobRowLength}`,
      `expected=${jobListingColumnCount}`,
    );

    // ✅ STEP 7: Insert job row
    console.log("⏳ Inserting into JobListing...");
    await addTableRow("JobListing", jobRow);
    console.log("✅ Step 7: Job row inserted");

    // ✅ STEP 8: Insert batch rows
    const createDate = formatDate(new Date().toISOString());
    console.log("⏳ Inserting batch rows...");

    // Discover the actual column count for BatchListing table
    batchListingColumnCount = await getTableColumnCount("BatchListing");
    if (batchListingColumnCount <= 0) {
      throw new Error(
        `BatchListing column count invalid: ${batchListingColumnCount}`,
      );
    }
    console.log(`📊 BatchListing table has ${batchListingColumnCount} columns`);

    for (let batchIndex = 0; batchIndex < data.batches.length; batchIndex++) {
      const batch = data.batches[batchIndex];
      const batchId = `${data.psn}-${batchIndex + 1}`;
      console.log("➡️ Processing batch:", batchId);

      const batchQty = batch.batchQty || 0;

      // Build basic batch info (columns 0-5)
      const batchRow = [
        data.psn,
        batchId,
        createDate,
        data.jobName || "",
        batchQty,
        "", // Status column
      ];

      // Build step data block (columns 6-113: 12 steps × 9 columns)
      const stepsData = new Array(108).fill("");
      const BLOCK_SIZE = 9;

      if (batch.steps && Array.isArray(batch.steps)) {
        batch.steps.forEach((step, index) => {
          if (index < 12) {
            const baseIdx = index * BLOCK_SIZE;
            stepsData[baseIdx] = step.processName || "";
            stepsData[baseIdx + 1] = formatDate(step.expDate) || "";
            stepsData[baseIdx + 2] = formatDate(step.endDate) || "";
            stepsData[baseIdx + 3] = step.duration || "";
            stepsData[baseIdx + 4] = step.detail || "";
            stepsData[baseIdx + 5] = step.status || "";
            stepsData[baseIdx + 6] = step.remark || "";
            stepsData[baseIdx + 7] = step.revertRemark || "";
            stepsData[baseIdx + 8] = step.isDone ? "TRUE" : "FALSE";
          }
        });
      }

      // Combine basic + steps = 114 columns
      let finalRow = batchRow.concat(stepsData);

      // Pad/trim to match exact table column count
      if (finalRow.length < batchListingColumnCount) {
        // Pad with empty strings
        const neededPadding = batchListingColumnCount - finalRow.length;
        for (let p = 0; p < neededPadding; p++) {
          finalRow.push("");
        }
      } else if (finalRow.length > batchListingColumnCount) {
        // Trim to exact size
        finalRow = finalRow.slice(0, batchListingColumnCount);
      }

      batchRowLength = finalRow.length;
      console.log(
        `📦 Batch row prepared: ${batchRowLength} columns (table expects ${batchListingColumnCount})`,
      );
      if (batchRowLength !== batchListingColumnCount) {
        const err = new Error(
          `Batch row length mismatch for ${batchId}: finalRow=${batchRowLength}, expected=${batchListingColumnCount}`,
        );
        err.batchListingColumnCount = batchListingColumnCount;
        err.batchRowLength = batchRowLength;
        throw err;
      }

      await addTableRow("BatchListing", finalRow);
      console.log("✅ Batch inserted:", batchId);
    }

    console.log("✅ Step 8: All batch rows inserted");

    // ✅ SUCCESS
    const message = `Success! Job recorded and ${data.batches.length} batch(es) created.`;
    console.log("🎉 SUCCESS:", message);

    res.json({ message });
  } catch (error) {
    console.error("🔥 ERROR OCCURRED");
    console.error("🔥 Error Type:", error.constructor.name);
    console.error("🔥 Error Message:", error.message);

    const errorPayload = {
      errorName: error.constructor.name,
      message: error.message,
      jobListingColumnCount:
        error.jobListingColumnCount || jobListingColumnCount || null,
      batchListingColumnCount:
        error.batchListingColumnCount || batchListingColumnCount || null,
      jobRowLength: error.jobRowLength || jobRowLength || null,
      batchRowLength: error.batchRowLength || batchRowLength || null,
    };

    console.error(
      "👉 Diagnostic Payload:",
      JSON.stringify(errorPayload, null, 2),
    );

    if (error.response) {
      console.error("👉 Graph API Error Status:", error.response.status);
      console.error(
        "👉 Graph API Error Body:",
        JSON.stringify(error.response.data, null, 2),
      );
      errorPayload.graphError = {
        status: error.response.status,
        data: error.response.data,
      };
    }

    console.error("🔥 FINAL RESPONSE:", JSON.stringify(errorPayload, null, 2));

    res.status(500).json({
      error: "Failed to submit data",
      diagnostics: errorPayload,
    });
  }
});

// ============ HELPER FUNCTIONS FOR BATCH UPDATES ============

/**
 * Normalize PSN: trim, lowercase, remove trailing '.0'
 */
function normalizePsn(value) {
  let str = String(value || "")
    .trim()
    .toLowerCase();
  // Remove trailing .0 (common in Excel number conversions)
  str = str.replace(/\.0$/, "");
  return str;
}

/**
 * Fetch a specific row from BatchListing by row number
 * Uses the table rows API to get the row at the specified index
 */
async function getBatchListingRow(tableRowIndex) {
  const ctx = await getSharePointFileContext();

  try {
    const res = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables('BatchListing')/rows`,
      { headers: ctx.headers },
    );

    const rows = res.data.value || [];

    // Check if the index exists in the data rows array
    if (tableRowIndex < 0 || tableRowIndex >= rows.length) {
      throw new Error(
        `Index ${tableRowIndex} out of bounds. Table has ${rows.length} rows.`,
      );
    }

    // Get the row at that specific index
    const row = rows[tableRowIndex];
    const values = row.values?.[0] || [];

    return values;
  } catch (error) {
    console.error(`[getBatchListingRow] Error:`, error.message);
    throw error;
  }
}

/**
 * Convert column index (0-based) to Excel column letter
 * 0=A, 1=B, ..., 25=Z, 26=AA, 27=AB, etc.
 */
function indexToColumnLetter(index) {
  let letter = "";
  while (index >= 0) {
    letter = String.fromCharCode(65 + (index % 26)) + letter;
    index = Math.floor(index / 26) - 1;
  }
  return letter;
}

// Inside your saveMultiBatchUpdate or helper function:

async function updateBatchListingCell(tableRowIndex, colIndex, value) {
  const ctx = await getSharePointFileContext();
  const colLetter = indexToColumnLetter(colIndex);
  const physicalRow = parseInt(tableRowIndex) + 4;

  const cellAddress = `${colLetter}${physicalRow}`;
  const url = `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/worksheets('BatchListing')/range(address='${cellAddress}')`;
  console.log(
    `[Patch] Targeting ${cellAddress} for Table Index ${tableRowIndex}`,
  );
  let formattedValue = value;
  if (value === true) formattedValue = "TRUE";
  if (value === false) formattedValue = "FALSE";

  await axios.patch(
    url,
    {
      values: [[formattedValue]],
    },
    { headers: ctx.headers },
  );
}

/**
 * Update JobListing delivery date by PSN match
 */
async function updateJobListingDeliveryDateByPsn(psn, deliveryDate) {
  const ctx = await getSharePointFileContext();
  const normalizedSearchPsn = normalizePsn(psn);

  try {
    // Get all JobListing rows to find the match
    const rowsRes = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables('JobListing')/rows`,
      { headers: ctx.headers },
    );
    const rows = rowsRes.data.value || [];

    console.log(
      `[updateJobListingDeliveryDateByPsn] Searching ${rows.length} JobListing rows for PSN:`,
      psn,
    );

    for (let i = 0; i < rows.length; i++) {
      const values = rows[i].values?.[0] || [];
      const rowPsn = normalizePsn(values[0]); // Column A (PSN)

      if (rowPsn === normalizedSearchPsn) {
        const excelRowNumber = i + 2;

        console.log(
          `[updateJobListingDeliveryDateByPsn] Found PSN match at table row ${i}, Excel row ${excelRowNumber}`,
        );

        // Use worksheet API endpoint for update
        const endpoint = `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/worksheets('JobListing')/range(address='H${excelRowNumber}')`;

        await axios.patch(
          endpoint,
          { values: [[deliveryDate]] },
          { headers: ctx.headers },
        );

        console.log(
          `[updateJobListingDeliveryDateByPsn] Updated JobListing for PSN ${psn} to delivery date:`,
          deliveryDate,
        );
        return true;
      }
    }

    console.warn(
      `[updateJobListingDeliveryDateByPsn] No matching PSN found in JobListing:`,
      psn,
    );
    return false;
  } catch (error) {
    console.error(
      `[updateJobListingDeliveryDateByPsn] Error updating delivery date:`,
      error.message,
    );
    throw error;
  }
}

/**
 * Generate a new Batch ID based on existing suffixes
 * Example: If J101-1 exists, returns J101-2
 */
async function generateNewBatchId(currentBatchId) {
  // Use the raw data fetcher to get all rows
  const rows = await getTableRowsAsObjects("BatchListing");
  const baseId = String(currentBatchId).split("-")[0];
  let maxSuffix = 0;

  rows.forEach((row) => {
    const bId = String(row.BatchId || row.batchId || "");
    if (bId.startsWith(baseId + "-")) {
      const suffix = parseInt(bId.split("-").pop());
      if (!isNaN(suffix) && suffix > maxSuffix) maxSuffix = suffix;
    }
  });

  return `${baseId}-${maxSuffix + 1}`;
}

/**
 * Main function: saveMultiBatchUpdate
 * Orchestrates batch updates with waterfall logic and splits
 */
function excelToJSDate(serial) {
  if (!serial || isNaN(serial)) return null;
  // If it's already a string date, just return it as a Date object
  if (typeof serial === "string" && serial.includes("-"))
    return new Date(serial);

  // Excel base date is Dec 30, 1899
  const excelEpoch = new Date(1899, 11, 30);
  const jsDate = new Date(excelEpoch.getTime() + serial * 86400000);
  return jsDate;
}

async function saveMultiBatchUpdate(payload) {
  const rowIdx = parseInt(payload.row); // 1-based Excel row number (e.g., 19)
  const BLOCK_SIZE = 9;
  const START_COL = 6;

  try {
    // 1. FETCH CURRENT ROW DATA
    const runningRowData = await getBatchListingRow(rowIdx);
    const psn = String(runningRowData[0] || "").trim();
    const today = new Date();
    const todayISO = today.toISOString().split("T")[0];

    const normalize = (v) =>
      String(v).replace(/\.0$/, "").replace(/\s/g, "").toLowerCase();

    // 2. PARSE QTY MAP
    const existingQtyString = String(runningRowData[119] || "");
    let qtyMap = {};
    if (
      existingQtyString &&
      existingQtyString !== "0" &&
      existingQtyString !== ""
    ) {
      existingQtyString.split("|").forEach((pair) => {
        const parts = pair.split(":");
        if (parts.length === 2) qtyMap[parts[0]] = parts[1];
      });
    }

    const updates = (payload.updates || []).sort(
      (a, b) => a.baseCol - b.baseCol,
    );
    let trackingQty = Number(runningRowData[4]);

    for (const u of updates) {
      const currentProcessName = String(runningRowData[u.baseCol] || "").trim();
      const isDeliveryStep = currentProcessName
        .toLowerCase()
        .includes("delivery");
      const isSplitting =
        u.isDone && u.qty < Number(qtyMap[u.baseCol] || trackingQty);

      // STATUS & REMARKS
      let statusVal = "";
      if (u.isDone) {
        // If it's a delivery step and we are splitting the quantity, mark as Partially Delivered
        if (isDeliveryStep && isSplitting) {
          statusVal = "Partially Delivered";
        } else {
          statusVal = u.isDelayed ? "Delayed" : "Completed";
        }
      } else if (u.isDelayed) {
        statusVal = "Delayed";
      }
      await updateBatchListingCell(rowIdx, u.baseCol + 5, statusVal);
      await updateBatchListingCell(
        rowIdx,
        u.baseCol + 6,
        u.isDelayed ? u.remark || "" : "",
      );

      if (u.detail)
        await updateBatchListingCell(rowIdx, u.baseCol + 4, u.detail);

      if (u.isDone) {
        const currentAvailable = Number(qtyMap[u.baseCol] || trackingQty);

        // --- UPDATED DURATION LOGIC ---
        let prevDateRaw = null;
        for (
          let prevBase = u.baseCol - BLOCK_SIZE;
          prevBase >= START_COL;
          prevBase -= BLOCK_SIZE
        ) {
          if (runningRowData[prevBase + 2]) {
            prevDateRaw = runningRowData[prevBase + 2];
            break;
          }
        }

        // Fallback to Batch Date (Column Index 2) if no previous step found
        if (!prevDateRaw) prevDateRaw = runningRowData[2];

        // CONVERT: Use the helper to handle Serial Numbers vs ISO Strings
        const startDate = excelToJSDate(prevDateRaw) || new Date();

        // Set hours to 0 to ensure we only count full calendar days
        const d1 = new Date(todayISO);
        const d2 = new Date(startDate.toISOString().split("T")[0]);

        const diffTime = d1.getTime() - d2.getTime();
        const diffDays = Math.max(
          1,
          Math.ceil(diffTime / (1000 * 60 * 60 * 24)),
        );

        // Update Sheet
        await updateBatchListingCell(rowIdx, u.baseCol + 2, todayISO);
        await updateBatchListingCell(rowIdx, u.baseCol + 3, diffDays);
        await updateBatchListingCell(rowIdx, u.baseCol + 8, true);

        // Update local memory for sequential steps in the same save
        runningRowData[u.baseCol + 2] = todayISO;
        runningRowData[u.baseCol + 8] = true;
        runningRowData[u.baseCol + 5] = u.isDelayed ? "Delayed" : "Completed";

        // DELIVERY SYNC
        const targetDate = payload.deliveryDate || payload.newDeliveryDate;
        if (isDeliveryStep && targetDate && targetDate !== "KEEP_ORIGINAL") {
          await updateJobListingDeliveryDateByPsn(psn, targetDate);
        }

        // SPLIT LOGIC
        if (isSplitting) {
          const diff = currentAvailable - u.qty;
          // Use a clean snapshot of the row BEFORE current changes for the child
          await createSplitBatchFromWaterfall(
            runningRowData,
            qtyMap,
            diff,
            u.baseCol,
            payload.splitRemark || "",
          );

          trackingQty = u.qty;
          for (let i = 0; i < 12; i++) {
            let futureBase = START_COL + i * BLOCK_SIZE;
            if (futureBase >= u.baseCol) qtyMap[futureBase] = trackingQty;
          }
        } else {
          qtyMap[u.baseCol] = u.qty;
        }
      }
    }

    // FINALIZE
    const newQtyString = Object.keys(qtyMap)
      .sort((a, b) => a - b)
      .map((k) => `${k}:${qtyMap[k]}`)
      .join("|");
    await updateBatchListingCell(rowIdx, 119, newQtyString);
    await updateBatchListingCell(rowIdx, 4, trackingQty);

    return { success: true };
  } catch (error) {
    console.error("[saveMultiBatchUpdate] Failed:", error.message);
    throw error;
  }
}

/**
 * Create a split batch from waterfall
 */
async function createSplitBatchFromWaterfall(
  parentData,
  parentQtyMap,
  diffQty,
  splitAtBase,
  userRemark,
) {
  const BLOCK_SIZE = 9;
  const START_COL = 6;

  // Clone the parent array
  let newRow = [...parentData];

  // Update Identity & Qty
  newRow[1] = await generateNewBatchId(String(parentData[1]));
  newRow[2] = new Date().toISOString().split("T")[0];
  newRow[4] = diffQty;
  newRow[118] = userRemark;

  // Build New Child Qty Map
  let childMap = {};
  for (let i = 0; i < 12; i++) {
    childMap[START_COL + i * BLOCK_SIZE] = diffQty;
  }
  newRow[119] = Object.keys(childMap)
    .sort((a, b) => a - b)
    .map((k) => `${k}:${childMap[k]}`)
    .join("|");

  // RESET Forward Steps for the new split row
  for (let i = 0; i < 12; i++) {
    let base = START_COL + i * BLOCK_SIZE;
    if (base >= splitAtBase) {
      newRow[base + 2] = ""; // End Date
      newRow[base + 3] = ""; // Duration
      newRow[base + 5] = ""; // Status
      newRow[base + 8] = false; // Tick
      newRow[base + 4] = ""; // Detail
      newRow[base + 6] = ""; // Remark
    }
  }

  return await addTableRow("BatchListing", newRow);
}

/**
 * Update process quantities only (without marking as done)
 */
async function updateProcessQtysOnly(rowIdx, qtyMapArray) {
  try {
    console.log(`[updateProcessQtysOnly] Updating row ${rowIdx} quantities`);

    const serialized = qtyMapArray
      .map((obj) => `${obj.baseCol}:${obj.qty}`)
      .join("|");

    // Update Column DP (index 119)
    await updateBatchListingCell(rowIdx, 119, serialized);

    console.log("[updateProcessQtysOnly] Quantities updated successfully");
    return { success: true, message: "Quantities updated" };
  } catch (error) {
    console.error("[updateProcessQtysOnly] Error:", error.message);
    throw error;
  }
}

/**
 * Revert a process step
 */
async function revertProcessStep(rowIdx, baseCol, revertRemark) {
  const ctx = await getSharePointFileContext();
  const START_COL = 6;
  const BLOCK_SIZE = 9;
  const TOTAL_STEPS = 12;

  try {
    if (!revertRemark || revertRemark.trim() === "") {
      throw new Error("Revert remark is mandatory.");
    }

    console.log(
      `[revertProcessStep] Reverting row ${rowIdx}, baseCol ${baseCol}`,
    );

    // Fetch the entire row
    const rowValues = await getBatchListingRow(rowIdx);

    // Iterate through all steps from the target baseCol onwards
    for (let i = 0; i < TOTAL_STEPS; i++) {
      let currentStepBase = START_COL + i * BLOCK_SIZE;

      if (currentStepBase >= baseCol) {
        const pName = rowValues[currentStepBase];
        if (!pName || pName === "" || pName === "--") continue;

        // Check if this step was completed
        const wasDone =
          rowValues[currentStepBase + 8] === true ||
          String(rowValues[currentStepBase + 8]).toUpperCase() === "TRUE";

        // Clear completion data
        await updateBatchListingCell(rowIdx, currentStepBase + 2, ""); // End Date
        await updateBatchListingCell(rowIdx, currentStepBase + 3, ""); // Duration
        await updateBatchListingCell(rowIdx, currentStepBase + 6, ""); // Completion Remark
        await updateBatchListingCell(rowIdx, currentStepBase + 8, false); // Untick

        // Handle remarks/status
        if (currentStepBase === baseCol) {
          // Target step
          await updateBatchListingCell(rowIdx, currentStepBase + 5, "Reverted"); // Status
          await updateBatchListingCell(
            rowIdx,
            currentStepBase + 7,
            revertRemark,
          ); // Revert Remark
        } else if (wasDone) {
          // Sequential completed steps
          await updateBatchListingCell(rowIdx, currentStepBase + 5, ""); // Status
          await updateBatchListingCell(
            rowIdx,
            currentStepBase + 7,
            "Auto-reverted (Sequential)",
          ); // Revert Remark
        } else {
          // Pending steps
          await updateBatchListingCell(rowIdx, currentStepBase + 5, ""); // Status
          await updateBatchListingCell(rowIdx, currentStepBase + 7, ""); // Revert Remark
        }
      }
    }

    console.log("[revertProcessStep] Revert completed successfully");
    return { success: true, message: "Process step reverted" };
  } catch (error) {
    console.error("[revertProcessStep] Error:", error.message);
    throw error;
  }
}

// ============ API ENDPOINTS FOR BATCH UPDATES ============

app.post("/api/updateProcessQtysOnly", async (req, res) => {
  try {
    console.log("[API] POST /api/updateProcessQtysOnly");
    const { row, qtyMapArray } = req.body;

    const result = await updateProcessQtysOnly(row, qtyMapArray);
    res.json(result);
  } catch (error) {
    console.error("[API] POST /api/updateProcessQtysOnly - error:", error);
    res
      .status(500)
      .json({ error: "Failed to update quantities", details: error.message });
  }
});

app.post("/api/saveMultiBatchUpdate", async (req, res) => {
  try {
    console.log("[API] POST /api/saveMultiBatchUpdate");
    const result = await saveMultiBatchUpdate(req.body);
    res.json(result);
  } catch (error) {
    console.error("[API] POST /api/saveMultiBatchUpdate - error:", error);
    res
      .status(500)
      .json({ error: "Failed to save batch update", details: error.message });
  }
});

app.post("/api/revertProcessStep", async (req, res) => {
  try {
    console.log("[API] POST /api/revertProcessStep");
    const { row, baseCol, revertRemark } = req.body;

    const result = await revertProcessStep(row, baseCol, revertRemark);
    res.json(result);
  } catch (error) {
    console.error("[API] POST /api/revertProcessStep - error:", error);
    res
      .status(500)
      .json({ error: "Failed to revert process step", details: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
