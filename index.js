const express = require("express");
const path = require("path");
const axios = require("axios");
require("dotenv").config();
const { ConfidentialClientApplication } = require("@azure/msal-node");

const app = express();
app.use(express.json());
const PORT = process.env.PORT || 3000;

//Date format not string
const toExcelDate = (date) => {
  if (!(date instanceof Date)) return null;

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`; // ✅ Excel safe
};

//Convert various date formats (Excel serial, dd/mm/yyyy, ISO) to consistent dd/mm/yyyy for frontend display
const displayDate = (value) => {
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

const formatDateForExcel = (value) => {
  if (!value || value === "-") return "";

  let raw = typeof value === "string" ? value.trim() : value;
  let date;

  // ✅ 1. Excel serial number (FIXED timezone)
  if (!isNaN(raw) && raw !== "") {
    const num = Number(raw);
    if (num > 0) {
      const utcDate = new Date(Math.round((num - 25569) * 86400 * 1000));
      date = new Date(
        utcDate.getUTCFullYear(),
        utcDate.getUTCMonth(),
        utcDate.getUTCDate(),
      );
    }
  }

  // ✅ 2. Handle DD/MM/YYYY or YYYY-MM-DD
  if (
    !date &&
    typeof raw === "string" &&
    (raw.includes("/") || raw.includes("-"))
  ) {
    const parts = raw.replace(/\s+/g, "").split(/[-\/]/);

    if (parts.length === 3) {
      let [p1, p2, p3] = parts;
      let day, month, year;

      if (Number(p1) > 31) {
        // YYYY-MM-DD
        year = p1;
        month = p2;
        day = p3;
      } else {
        // DD/MM/YYYY
        day = p1;
        month = p2;
        year = p3;
      }

      if (year.length === 2) year = "20" + year;

      date = new Date(Number(year), Number(month) - 1, Number(day));
    }
  }

  // ✅ 3. Fallback
  if (!date) {
    date = new Date(raw);
  }

  if (isNaN(date)) return "";

  // ✅ FINAL: LOCAL FORMAT (NO TIMEZONE SHIFT)
  const pad = (n) => (n < 10 ? "0" + n : n);

  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
};

// Convert date into Excel-safe TEXT format(Use for POST/api/submitData -> (Delivery Format) 'yyyy-MM-dd')
const toExcelDateText = (value) => {
  if (value === undefined || value === null || value === "" || value === "-")
    return "";

  let raw = typeof value === "string" ? value.trim() : value;
  let date = null;

  // ✅ Case 1: Excel serial
  if (!isNaN(raw)) {
    const num = Number(raw);
    if (num > 0) {
      const utcDate = new Date(Math.round((num - 25569) * 86400 * 1000));
      date = new Date(
        utcDate.getUTCFullYear(),
        utcDate.getUTCMonth(),
        utcDate.getUTCDate(),
      );
    }
  }

  // ✅ Case 2: ISO yyyy-mm-dd
  else if (
    typeof raw === "string" &&
    raw.includes("-") &&
    raw.split("-")[0].length === 4
  ) {
    const temp = new Date(raw);
    date = new Date(temp.getFullYear(), temp.getMonth(), temp.getDate());
  }

  // ✅ Case 3: dd/mm/yyyy
  else if (
    typeof raw === "string" &&
    (raw.includes("/") || raw.includes("-"))
  ) {
    const parts = raw.split(/[-\/]/);
    if (parts.length === 3) {
      let [d, m, y] = parts;
      if (y.length === 2) y = "20" + y;
      date = new Date(y, m - 1, d);
    }
  }

  // ✅ Case 4: fallback
  else {
    const temp = new Date(raw);
    date = new Date(temp.getFullYear(), temp.getMonth(), temp.getDate());
  }

  if (!date || isNaN(date.getTime())) return "";

  // ✅ Format yyyy-MM-dd (TEXT for Excel)
  const pad = (n) => (n < 10 ? "0" + n : n);
  return `'${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
};

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

const cca = new ConfidentialClientApplication(msalConfig);

//Authenticate with Microsoft Graph API using Client Credentials
async function getAccessToken() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return result.accessToken;
}

//Get SharePoint file context: siteId, driveId, fileId, and auth headers for API calls
async function getSharePointFileContext() {
  const token = await getAccessToken();
  const headers = { Authorization: `Bearer ${token}` };

  // Get site ID for the SharePoint site containing the file (adjust the path as needed)
  const siteRes = await axios.get(
    "https://graph.microsoft.com/v1.0/sites/mlmpackagingmy.sharepoint.com:/sites/FileStorage",
    { headers },
  );
  const siteId = siteRes.data.id;
  // Get the first drive in the site (adjust if you know the specific drive name or ID)
  const drivesRes = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers },
  );
  const driveId = drivesRes.data.value[0]?.id;
  if (!driveId) throw new Error("Could not find driveId in site drives");

  // Get the file ID for Database.xlsx in the drive (adjust the path if the file is in a subfolder)
  const fileRes = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/Database.xlsx`,
    { headers },
  );

  const fileId = fileRes.data.id;
  if (!fileId) throw new Error("Could not find fileId for Database.xlsx");

  return { token, headers, siteId, driveId, fileId };
}

// Fetch all rows from a specified table and convert to array of objects using header row for keys
async function getTableRowsAsObjects(tableName) {
  const ctx = await getSharePointFileContext();

  // Get header row to use as keys for objects (assumes header row is in first row of the table)
  const headerRes = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables/${tableName}/headerRowRange`,
    { headers: ctx.headers },
  );
  const headers = headerRes.data.values?.[0] || [];

  // Get all rows in the table (this may return values in different formats depending on Graph API response, so we handle multiple cases)
  const rowsRes = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables/${tableName}/rows`,
    { headers: ctx.headers },
  );
  const rows = rowsRes.data.value || [];

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
      }
    }

    // console.log(`[getTableRowsAsObjects] ${tableName} row ${idx} final values:`, values);

    const item = {};
    headers.forEach((h, idx) => {
      item[h] = values[idx] !== undefined ? values[idx] : null;
    });
    return item;
  });

  // console.log(
  //   `[getTableRowsAsObjects] Converted ${tableName} to objects:`,
  //   result.slice(0, 2),
  // );
  return result;
}

// Get the column count of a table by first trying the /columns endpoint, then falling back to headerRowRange if needed (handles different Graph API response formats)
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

    return headers.length;
  } catch (error) {
    throw error;
  }
}

// Add a new row to a specified table with given values (values should be an array matching the table's column order)
async function addTableRow(tableName, values) {
  const ctx = await getSharePointFileContext();
  console.log(
    `[addTableRow] Adding row to ${tableName}:`,
    values.length,
    "columns",
  );

  // Ensure values is an array of strings/numbers and convert any non-string values to strings (Graph API can be sensitive to data types)
  const resp = await axios.post(
    `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables/${tableName}/rows/add`,
    { values: [values] },
    { headers: ctx.headers },
  );

  console.log(`[addTableRow] Successfully added to ${tableName}`);
  return resp.data;
}

async function generateDashboardData() {
  // Get file context once
  const fileContext = await getSharePointFileContext();

  // Fetch all rows for BatchListing with numeric index preservation
  const batchRowRes = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${fileContext.driveId}/items/${fileContext.fileId}/workbook/tables('BatchListing')/rows`,
    { headers: fileContext.headers },
  );
  const batchRowsRaw = batchRowRes.data.value || [];

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
        jobType: job["Job Type"],
        orderDate: displayDate(
          job["Order Date"] || job.Order_Date || job.orderDate,
        ),
        priority: job.Priority || "NORMAL",
        deliveryDate: displayDate(
          job["Delivery Date"] || job.Delivery_Date || job.deliveryDate,
        ),
        status: job.Status || "ON SCHEDULE",
      };
    }
  });

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
      }
    }

    // Extract basic info from numeric indices (matching app.js)
    const psn = String(values[0] || "").trim();
    if (!psn || psn.toLowerCase().includes("done")) {
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
          const formatted = displayDate(val);
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
      jobType: "-",
      orderDate: "-",
      priority: "NORMAL",
      deliveryDate: "-",
      status: "ON SCHEDULE",
    };
    batches.push({
      row: rowIdx,
      psn: psn,
      batchId: String(values[1] || ""),
      batchDate: excelToJSDate(values[2]),
      jobName: String(values[3] || ""),
      qty: Number(values[4] || 0),
      progress: definedSteps.length > 0 ? ticksFound / definedSteps.length : 0,
      steps: definedSteps,
      activeStepStatus: activeStepStatus,
      piNumber: info.pi,
      salesCode: info.code,
      clientName: info.client,
      jobType: info.jobType,
      orderDate: info.orderDate,
      priority: info.priority,
      deliveryDate: info.deliveryDate,
      status: values[10] || info.status,
      splitRemark: String(values[118] || "").trim(),
      qtyString: String(values[119] || ""),
      maxQtyString: String(values[120] || ""),
    });
  });

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
      // ✅ 1. Handle DONE steps
      batch.steps.forEach((step) => {
        if (step.isDone && processStats[step.name]) {
          const durationDays = Math.max(
            1,
            Math.ceil(Number(step.duration) || 0),
          );
          processStats[step.name].totalTime += durationDays;
          processStats[step.name].count += 1;
        }
      });

      // ✅ 2. Handle ONLY current active step
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

  //
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

  // Calculate requirement updates
  const requirementUpdates = [];

  batches.forEach((batch) => {
    if (!batch.steps || batch.steps.length === 0) return;
    // Find current step index (first step without isDone, meaning not done)
    let currentStepIndex = -1;
    for (let i = 0; i < batch.steps.length; i++) {
      if (!batch.steps[i].isDone) {
        currentStepIndex = i;
        break;
      }
    }
    if (currentStepIndex === -1) return; // all steps done

    const currentStep = batch.steps[currentStepIndex];
    const expectedDateStr = currentStep.expDate;
    if (!expectedDateStr || expectedDateStr === "-") return;

    // Parse expected date (dd/mm/yyyy)
    const [day, month, year] = expectedDateStr.split("/").map(Number);
    const expectedDate = new Date(year, month - 1, day);
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const diffTime = expectedDate - today;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    if (diffDays > 2) return; // not urgent

    let duration;
    if (diffDays == 0) {
      duration = `Due today`;
    } else if (diffDays > 0) {
      duration = `${diffDays} days left`;
    } else {
      duration = `Overdue ${Math.abs(diffDays)} days`;
    }

    let lastUpdated = "";
    let prevDurationTime = "";
    if (currentStepIndex > 0) {
      const prevStep = batch.steps[currentStepIndex - 1];
      const prevDurationStr = prevStep.endDate;
      if (!prevDurationStr || prevDurationStr === "-") return;

      // Parse expected date (dd/mm/yyyy)
      const [day, month, year] = prevDurationStr.split("/").map(Number);
      const prevDuration = new Date(year, month - 1, day);

      const prevdiffTime = prevDuration - today;
      const prevdiffDays = Math.ceil(prevdiffTime / (1000 * 60 * 60 * 24));
      prevDurationTime = `${Math.abs(prevdiffDays)}`;
      lastUpdated = `${prevStep.name}-${prevStep.endDate}` || "";
    } else {
      // First step: use batch creation date
      const batchDateStr = batch.batchDate;
      if (!batchDateStr || batchDateStr === "-") return;

      const batchDate = new Date(batchDateStr);
      const batchDiffTime = today - batchDate;
      const batchDiffDays = Math.ceil(batchDiffTime / (1000 * 60 * 60 * 24));
      prevDurationTime = `${batchDiffDays}`;
      const formattedBatchDate = batchDate.toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
      });
      lastUpdated = `Created-${formattedBatchDate}`;
    }

    requirementUpdates.push({
      psn: batch.psn,
      batchId: batch.batchId,
      jobName: batch.jobName,
      quantity: batch.qty,
      currentProcess: currentStep.name,
      duration,
      lastUpdated,
      dueDays: diffDays,
      isDelayed: diffDays < 0,
      prevDurationTime,
    });
  });

  const result = {
    jobs: batches,
    jobListing: batches, // For compatibility with both old and new frontend
    batchListing: batches, // Also return as batchListing for backward compat
    averages: averages, // For process workload display
    rawCapacity: rawCapacity, // For machine capacity display
    requirementUpdates: requirementUpdates, // New requirement update listing
    stats: {
      totalJobs,
      totalBatches,
      totalQuantity,
      completedBatches,
      pendingBatches: totalBatches - completedBatches,
    },
  };

  return result;
}

// ============ EXPRESS ROUTES ============
// Redirect root to dashboard.html
app.get("/", (req, res) => {
  console.log("[APP] GET / - redirecting to dashboard");
  res.redirect("/dashboard.html");
});
// Serve static files from the 'public' directory (make sure to create this and add your frontend files)
app.use(express.static(path.join(__dirname, "public")));

// API route to test server connectivity
app.get("/api/test", (req, res) => {
  console.log("[API] GET /api/test");
  res.json({ message: "Node.js backend working" });
});

// API route to get access token (for testing purposes, not needed by frontend)
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

// API route to get JobListing data
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

// API route to get BatchListing data
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

// Main API route to get dashboard data (combines JobListing and BatchListing with processing)
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

// API route to submit new job and batches
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
    const normalizedOrderDate = formatDateForExcel(data.orderDate);
    const normalizedDeliveryDate = formatDateForExcel(data.deliveryDate);

    console.log("📅 Formatted dates:", {
      order: normalizedOrderDate,
      delivery: normalizedDeliveryDate,
    });

    const formatDeliveryDate = toExcelDateText(data.deliveryDate);

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
      data.status || "ON SCHEDULE",
      formatDeliveryDate || "",
    ];

    jobListingColumnCount = await getTableColumnCount("JobListing");

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
    const createDate = toExcelDate(new Date());
    console.log("⏳ Inserting batch rows...");

    batchListingColumnCount = await getTableColumnCount("BatchListing");
    console.log(`📊 BatchListing table has ${batchListingColumnCount} columns`);

    for (let batchIndex = 0; batchIndex < data.batches.length; batchIndex++) {
      const batch = data.batches[batchIndex];
      const batchId = `${data.psn}-${batchIndex + 1}`;
      console.log("➡️ Processing batch:", batchId);

      const batchQty = batch.batchQty || 0;

      // Basic batch info
      const batchRow = [
        data.psn,
        batchId,
        createDate,
        data.jobName || "",
        batchQty,
        "",
      ];

      // Steps block (108 columns)
      const stepsData = new Array(108).fill("");
      const BLOCK_SIZE = 9;

      for (let j = 0; j < 12; j++) {
        const baseIdx = j * BLOCK_SIZE;
        stepsData[baseIdx + 8] = "FALSE";
      }

      if (batch.steps && Array.isArray(batch.steps)) {
        batch.steps.forEach((step, index) => {
          if (index < 12) {
            const baseIdx = index * BLOCK_SIZE;
            stepsData[baseIdx] = step.processName || "";
            stepsData[baseIdx + 1] = formatDateForExcel(step.expDate) || "";
            stepsData[baseIdx + 2] = formatDateForExcel(step.endDate) || "";
            stepsData[baseIdx + 3] = step.duration || "";
            stepsData[baseIdx + 4] = step.detail || "";
            stepsData[baseIdx + 5] = step.status || "";
            stepsData[baseIdx + 6] = step.remark || "";
            stepsData[baseIdx + 7] = step.revertRemark || "";
            // ✅ overwrite if provided
            stepsData[baseIdx + 8] = step.isDone ? "TRUE" : "FALSE";
          }
        });
      }

      // Combine
      let finalRow = batchRow.concat(stepsData);

      // ✅ NEW DYNAMIC QTY STRING LOGIC
      const START_COL = 6;

      let activeProcessCols = [];

      for (let j = 0; j < 12; j++) {
        let baseIdx = j * BLOCK_SIZE;
        let processName = String(stepsData[baseIdx] || "").trim();

        if (processName !== "" && processName !== "--") {
          let colIndex = START_COL + j * BLOCK_SIZE;
          activeProcessCols.push(colIndex);
        }
      }

      const qtyString = activeProcessCols
        .map((colIdx) => `${colIdx}:${batchQty}`)
        .join("|");

      // Ensure index 120 exists (DP = 119, DQ = 120)
      if (finalRow.length <= 120) {
        const neededPadding = 121 - finalRow.length;
        for (let i = 0; i < neededPadding; i++) {
          finalRow.push("");
        }
      }

      // Set Column DP & DQ
      finalRow[114] = "FALSE";
      finalRow[119] = qtyString;
      finalRow[120] = qtyString;

      // Pad/trim to match table
      if (finalRow.length < batchListingColumnCount) {
        const neededPadding = batchListingColumnCount - finalRow.length;
        for (let p = 0; p < neededPadding; p++) {
          finalRow.push("");
        }
      } else if (finalRow.length > batchListingColumnCount) {
        finalRow = finalRow.slice(0, batchListingColumnCount);
      }

      batchRowLength = finalRow.length;

      if (batchRowLength !== batchListingColumnCount) {
        throw new Error(`Batch row length mismatch for ${batchId}`);
      }

      await addTableRow("BatchListing", finalRow);
      console.log("✅ Batch inserted:", batchId);
    }

    console.log("✅ Step 8: All batch rows inserted");

    res.json({
      message: `Success! Job recorded and ${data.batches.length} batch(es) created.`,
    });
  } catch (error) {
    console.error("🔥 ERROR:", error.message);

    res.status(500).json({
      error: "Failed to submit data",
      message: error.message,
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
 * Updates an entire row in the BatchListing table in one API call.
 * @param {number} tableRowIndex - 0-based index of the row in the table
 * @param {Array} fullRowValues - The complete array of values for that row (all columns)
 */
async function updateBatchListingRow(tableRowIndex, fullRowValues) {
  const ctx = await getSharePointFileContext();
  const physicalRow = parseInt(tableRowIndex) + 4; // Your existing offset logic

  // We target the range from Column A (0) to the end of your data (e.g., Column ZZ or DP)
  const lastColLetter = indexToColumnLetter(fullRowValues.length - 1);
  const rangeAddress = `A${physicalRow}:${lastColLetter}${physicalRow}`;

  const url = `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/worksheets('BatchListing')/range(address='${rangeAddress}')`;

  // Ensure booleans are strings if your Excel logic requires "TRUE"/"FALSE"
  const formattedValues = fullRowValues.map((val) => {
    if (val === true) return "TRUE";
    if (val === false) return "FALSE";
    return val === null || val === undefined ? "" : val;
  });

  await axios.patch(
    url,
    { values: [formattedValues] },
    { headers: ctx.headers },
  );
  console.log(`[Row Update] Successfully patched Row ${physicalRow}`);
}

/**
 * Update JobListing delivery date by PSN match
 */
async function updateJobListingDeliveryDateByPsn(psn, deliveryDate) {
  const ctx = await getSharePointFileContext();
  const normalizedSearchPsn = normalizePsn(psn);

  try {
    const rowsRes = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables('JobListing')/rows`,
      { headers: ctx.headers },
    );
    const rows = rowsRes.data.value || [];

    for (let i = 0; i < rows.length; i++) {
      const values = rows[i].values?.[0] || [];
      const rowPsn = normalizePsn(values[0]);

      if (rowPsn === normalizedSearchPsn) {
        const excelRowNumber = i + 2;

        // --- PREPARE DIFFERENT FORMATS ---
        // Column H gets the display format (dd/mm/yyyy)
        const formatH = formatDateForExcel(deliveryDate);
        // Column L gets the Excel-safe text format ('yyyy-MM-dd)
        const formatL = toExcelDateText(deliveryDate);

        // --- UPDATE COLUMN H ---
        const endpointH = `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/worksheets('JobListing')/range(address='H${excelRowNumber}')`;
        await axios.patch(
          endpointH,
          { values: [[formatH]] },
          { headers: ctx.headers },
        );

        // --- UPDATE COLUMN L ---
        const endpointL = `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/worksheets('JobListing')/range(address='L${excelRowNumber}')`;
        await axios.patch(
          endpointL,
          { values: [[formatL]] },
          { headers: ctx.headers },
        );

        console.log(
          `[Success] PSN ${psn}: H updated to ${formatH}, L updated to ${formatL}`,
        );
        return true;
      }
    }
    return false;
  } catch (error) {
    console.error(`[Error] Update failed:`, error.message);
    throw error;
  }
}

/**
 * Generate a new Batch ID based on existing suffixes
 * Example: If J101-1 exists, returns J101-2
 */
async function generateNewBatchId(currentBatchId, localNewIds = []) {
  // 1. Fetch rows from the server
  const rows = await getTableRowsAsObjects("BatchListing");
  const baseId = String(currentBatchId).split("-")[0];
  let maxSuffix = 0;

  // 2. Check existing rows on the server
  rows.forEach((row) => {
    const bId = String(row.BatchId || row.batchId || "");
    if (bId.startsWith(baseId + "-")) {
      const suffix = parseInt(bId.split("-").pop());
      if (!isNaN(suffix) && suffix > maxSuffix) maxSuffix = suffix;
    }
  });

  // 3. Check IDs we just generated locally in this loop
  localNewIds.forEach((bId) => {
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
  const rowIdx = parseInt(payload.row);
  const START_COL = 6;
  const BLOCK_SIZE = 9;

  const COL_COMPLETED_DATE =115;
  const COL_COMPLETION_STATUS = 114;
  try {
    // 1. Fetch current row
    let runningRowData = await getBatchListingRow(rowIdx);
    const psn = String(runningRowData[0] || "").trim();
    const todayISO = new Date().toISOString().split("T")[0];
    let trackingQty = Number(runningRowData[4] || 0);

    // Parse Maps
    let qtyMap = {};
    let maxQtyMap = {};

    const parseToMap = (str, target) => {
      if (str && str !== "0") {
        str.split("|").forEach((p) => {
          const pts = p.split(":");
          if (pts.length === 2) target[pts[0]] = pts[1];
        });
      }
    };
    parseToMap(String(runningRowData[119] || ""), qtyMap);
    parseToMap(String(runningRowData[120] || ""), maxQtyMap);

    const updates = (payload.updates || []).sort(
      (a, b) => a.baseCol - b.baseCol,
    );

    for (const u of updates) {
      const currentProcessName = String(runningRowData[u.baseCol] || "").trim();
      const currentStepMax = Number(maxQtyMap[u.baseCol] || trackingQty);
      const isDeliveryStep = currentProcessName
        .toLowerCase()
        .includes("delivery");
      const isSplitting = u.isDone && Number(u.qty) < currentStepMax;

      // Update memory array instead of API
      runningRowData[u.baseCol + 5] = u.isDone
        ? isDeliveryStep && isSplitting
          ? "Partially Delivered"
          : u.isDelayed
            ? "Delayed"
            : "Completed"
        : u.isDelayed
          ? "Delayed"
          : "";
      runningRowData[u.baseCol + 6] = u.isDelayed ? u.remark || "" : "";
      if (u.detail) runningRowData[u.baseCol + 4] = u.detail;

      if (u.isDone) {
        // Duration Logic
        let prevDateRaw = null;
        for (
          let pb = u.baseCol - BLOCK_SIZE;
          pb >= START_COL;
          pb -= BLOCK_SIZE
        ) {
          if (runningRowData[pb + 2]) {
            prevDateRaw = runningRowData[pb + 2];
            break;
          }
        }
        if (!prevDateRaw) prevDateRaw = runningRowData[2];

        const startDate = excelToJSDate(prevDateRaw) || new Date();
        const diffDays = Math.max(
          1,
          Math.ceil(
            (new Date(todayISO) -
              new Date(startDate.toISOString().split("T")[0])) /
              86400000,
          ),
        );

        runningRowData[u.baseCol + 2] = todayISO;
        runningRowData[u.baseCol + 3] = diffDays;
        runningRowData[u.baseCol + 8] = true;

        // DELIVERY SYNC
        if (isDeliveryStep) {
          // Mark the entire row as completed in the final status columns
          runningRowData[COL_COMPLETION_STATUS] = true;
          runningRowData[COL_COMPLETED_DATE] = todayISO;

          // Delivery Sync to Job Listing
          const targetDate = payload.deliveryDate || payload.newDeliveryDate;
          if (targetDate && targetDate !== "KEEP_ORIGINAL") {
            await updateJobListingDeliveryDateByPsn(psn, targetDate);
          }
        }

        if (isSplitting) {
          const diff = currentStepMax - Number(u.qty);
          trackingQty = Number(u.qty);
          for (let i = 0; i < 12; i++) {
            let base = START_COL + i * BLOCK_SIZE;
            qtyMap[base] = trackingQty;
            maxQtyMap[base] = trackingQty;
          }
          // Important: Handle the creation of the new split batch row here
          await createSplitBatchFromWaterfall(
            runningRowData,
            diff,
            u.baseCol,
            payload.splitRemark || "Split Batch",
            [],
          );
        } else {
          qtyMap[u.baseCol] = u.qty;
        }
      }
    }

    const finalizeMap = (map) =>
      Object.keys(map)
        .sort((a, b) => a - b)
        .map((k) => `${k}:${map[k]}`)
        .join("|");
    // Update the final tracking columns in the memory array
    runningRowData[4] = trackingQty;
    runningRowData[119] = finalizeMap(qtyMap);
    runningRowData[120] = finalizeMap(maxQtyMap);

    // 2. ONE SINGLE API CALL TO SAVE EVERYTHING
    await updateBatchListingRow(rowIdx, runningRowData);

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
  diffQty,
  splitAtBase,
  userRemark,
  localNewIds,
) {
  const childMaxMap = {};
  const START_COL = 6;
  const BLOCK_SIZE = 9;

  for (let i = 0; i < 12; i++) {
    let base = START_COL + i * BLOCK_SIZE;
    // The child batch's max capacity is the split amount
    childMaxMap[base] = diffQty;
  }

  const childMaxString = Object.keys(childMaxMap)
    .sort((a, b) => Number(a) - Number(b))
    .map((k) => `${k}:${childMaxMap[k]}`)
    .join("|");

  // Clone the parent array
  let newRow = [...parentData];
  const newId = await generateNewBatchId(String(parentData[1]), localNewIds);
  // Update Identity & Qty
  newRow[1] = newId;
  newRow[2] = new Date().toISOString().split("T")[0];
  newRow[4] = diffQty;
  newRow[118] = userRemark;

  newRow[119] = childMaxString;
  newRow[120] = childMaxString;

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
  const response = await addTableRow("BatchListing", newRow);

  // Return the response so saveMultiBatchUpdate can update the localNewIds list
  return response;
}

/**
 * Update process quantities only (without marking as done)
 */
async function updateProcessQtysOnly(rowIdx, qtyMapArray) {
  try {
    const rowValues = await getBatchListingRow(rowIdx);
    const overallQty = Number(rowValues[4] || 0);

    // Parse Max Map (Column 120)
    const maxQtyMapString = String(rowValues[120] || "");
    let maxQtyMap = {};
    if (maxQtyMapString && maxQtyMapString !== "0") {
      maxQtyMapString.split("|").forEach((pair) => {
        const parts = pair.split(":");
        if (parts.length === 2) maxQtyMap[parts[0]] = Number(parts[1]);
      });
    }

    // Parse Existing Current Map (Column 119)
    let currentQtyMap = {};
    const existingQtyStr = String(rowValues[119] || "");
    if (existingQtyStr && existingQtyStr !== "0") {
      existingQtyStr.split("|").forEach((p) => {
        const pts = p.split(":");
        if (pts.length === 2) currentQtyMap[pts[0]] = Number(pts[1]);
      });
    }

    // VALIDATION & MERGE
    for (const item of qtyMapArray) {
      const base = String(item.baseCol);
      const newQty = Number(item.qty);
      const maxAllowed =
        maxQtyMap[base] !== undefined ? maxQtyMap[base] : overallQty;
      const existingQty = currentQtyMap[base];

      if (newQty !== existingQty && newQty > maxAllowed) {
        throw new Error(
          `Validation Failed: Column ${base} requested ${newQty}, but Max is ${maxAllowed}.`,
        );
      }

      // Only validate the items being UPDATED right now
      if (newQty > maxAllowed) {
        throw new Error(
          `Validation Failed: Column ${base} requested ${newQty}, but Max is ${maxAllowed}.`,
        );
      }
      // Update the map in memory
      currentQtyMap[base] = newQty;
    }

    // Serialize back to Column 119
    const serialized = Object.keys(currentQtyMap)
      .sort((a, b) => Number(a) - Number(b))
      .map((k) => `${k}:${currentQtyMap[k]}`)
      .join("|");

    await updateBatchListingCell(rowIdx, 119, serialized);

    return { success: true };
  } catch (error) {
    console.error("[updateProcessQtysOnly] Error:", error.message);
    throw error;
  }
}

/**
 * Revert a process step
 */
async function revertProcessStep(rowIdx, baseCol, revertRemark) {
  try {
    if (!revertRemark || revertRemark.trim() === "")
      throw new Error("Revert remark is mandatory.");

    // 1. Fetch the entire row once
    let rowValues = await getBatchListingRow(rowIdx);

    let currentQtyMap = parseMapString(rowValues[119], rowValues[4]);
    let maxQtyMap = parseMapString(rowValues[120], rowValues[4]);

    // 2. Modify the array in memory
    for (let i = 0; i < 12; i++) {
      let currentStepBase = 6 + i * 9;

      if (currentStepBase >= baseCol) {
        const pName = rowValues[currentStepBase];
        if (!pName || pName === "" || pName === "--") continue;

        const wasDone =
          rowValues[currentStepBase + 8] === true ||
          String(rowValues[currentStepBase + 8]).toUpperCase() === "TRUE";

        // Reset Qty to Max
        currentQtyMap[currentStepBase] =
          maxQtyMap[currentStepBase] || rowValues[4];

        // Clear Step Data in memory array
        rowValues[currentStepBase + 2] = ""; // End Date
        rowValues[currentStepBase + 3] = ""; // Duration
        rowValues[currentStepBase + 6] = ""; // Completion Remark
        rowValues[currentStepBase + 8] = false; // IsDone

        if (currentStepBase === baseCol) {
          rowValues[currentStepBase + 5] = "Reverted";
          rowValues[currentStepBase + 7] = revertRemark;
        } else if (wasDone) {
          rowValues[currentStepBase + 5] = "";
          rowValues[currentStepBase + 7] = "Auto-reverted (Sequential)";
        }
      }
    }

    // 3. Update the Map string in the array
    rowValues[119] = serializeMap(currentQtyMap);

    // 4. SAVE THE WHOLE ROW AT ONCE
    await updateBatchListingRow(rowIdx, rowValues);

    return { success: true, message: "Process step reverted" };
  } catch (error) {
    console.error("[revertProcessStep] Error:", error.message);
    throw error;
  }
}

async function saveDateUpdates(payload) {
  const rowIdx = parseInt(payload.row);

  try {
    let runningRowData = await getBatchListingRow(rowIdx);

    if (payload.updates && Array.isArray(payload.updates)) {
      payload.updates.forEach((u) => {
        const expDateColIndex = u.baseCol + 1;

        // u.newExpDate is "DD/MM/YY" from your frontend
        const parts = u.newExpDate.split("/");
        const day = parseInt(parts[0]);
        const month = parseInt(parts[1]) - 1; // 0-indexed
        const year = parseInt("20" + parts[2]);

        // 1. Create a date object in LOCAL Malaysia time
        // 2. Set the time to 12:00 PM (Noon) to act as a GMT+8 buffer
        const safeDate = new Date(year, month, day, 12, 0, 0);

        // 3. Convert back to your specific string format
        // This ensures the value sent to Excel is a PURE string
        const d = String(safeDate.getDate()).padStart(2, "0");
        const m = String(safeDate.getMonth() + 1).padStart(2, "0");
        const y = String(safeDate.getFullYear()).substring(2);

        const finalString = `${d}/${m}/${y}`;

        // 4. Wrap in a single quote if the backend is still auto-converting
        // This is the "Nuclear Option" to stop Excel from changing the date
        runningRowData[expDateColIndex] = `'${finalString}`;
      });
    }

    await updateBatchListingRow(rowIdx, runningRowData);
    return { success: true };
  } catch (error) {
    console.error("[saveDateUpdates] Failed:", error.message);
    throw error;
  }
}

function serializeMap(map) {
  return Object.keys(map)
    .sort((a, b) => a - b)
    .map((k) => `${k}:${map[k]}`)
    .join("|");
}

function parseMapString(str, fallbackQty) {
  if (!str || str === "0" || str === "" || str === "undefined") {
    return {};
  }
  const map = {};
  str.split("|").forEach((pair) => {
    const [k, v] = pair.split(":");
    if (k) map[k] = v;
  });
  return map;
}

// ============ API ENDPOINTS FOR BATCH UPDATES ============
// Endpoint to update process quantities only (without marking as done)
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

// Endpoint to save batch updates with waterfall logic and splits
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

// Endpoint to revert a process step
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

app.post("/api/updateProcessDates", async (req, res) => {
  try {
    console.log("[API] POST /api/updateProcessDates");
    const result = await saveDateUpdates(req.body);
    res.json(result);
  } catch (error) {
    console.error("[API] Error updating process dates:", error);
    res.status(500).json({
      error: "Failed to update expected dates",
      details: error.message,
    });
  }
});

app.get("/api/admin/repair-all", async (req, res) => {
  try {
    console.log("[Admin] Starting Precise Table Repair...");
    const ctx = await getSharePointFileContext();

    const rowsRes = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${ctx.driveId}/items/${ctx.fileId}/workbook/tables('BatchListing')/rows`,
      { headers: ctx.headers },
    );

    const rows = rowsRes.data.value || [];
    if (rows.length === 0) return res.send("No rows found.");

    let repairCount = 0;
    const START_COL = 6;
    const BLOCK_SIZE = 9;

    for (let i = 0; i < rows.length; i++) {
      const values = rows[i].values?.[0] || [];
      const actualBatchQty = Number(values[4] || 0); // Column E
      if (actualBatchQty <= 0) continue;

      // --- STEP 1: COUNT DEFINED PROCESSES ---
      // We only want to include columns that actually have a process name
      let activeProcessCols = [];
      for (let s = 0; s < 12; s++) {
        let base = START_COL + s * BLOCK_SIZE;
        let pName = String(values[base] || "").trim();

        // If there's a name, this is a real process for this specific batch
        if (pName !== "" && pName !== "--") {
          activeProcessCols.push(base);
        }
      }

      // --- STEP 2: BUILD THE PRECISE STRINGS ---
      // Format: "6:10000|15:10000" only for the columns we found above
      const fixedString = activeProcessCols
        .map((colIdx) => `${colIdx}:${actualBatchQty}`)
        .join("|");

      // --- STEP 3: COMPARE & UPDATE ---
      const existingQtyString = String(values[119] || "");

      // Only update if the string is different (prevents unnecessary API calls)
      // or if it contains values larger than the actual batch qty
      if (
        existingQtyString !== fixedString ||
        existingQtyString.includes("30000")
      ) {
        await updateBatchListingCell(i, 119, fixedString); // Current Qty Map
        await updateBatchListingCell(i, 120, fixedString); // Max Qty Map

        repairCount++;
        console.log(
          `[Fixed] Row ${i}: Found ${activeProcessCols.length} processes. Synced to ${actualBatchQty}`,
        );
      }
    }

    res.send(
      `<h1>Repair Complete</h1><p>Processed ${rows.length} rows. Fixed <b>${repairCount}</b> mismatches.</p>`,
    );
  } catch (error) {
    console.error("Repair Error:", error.message);
    res.status(500).send(error.message);
  }
});

// =========== START SERVER ============
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
