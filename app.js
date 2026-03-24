// const SPREADSHEET_ID = "12zrqlEbQii9XmV6LHAxCDTslpHbOYA98Ueo6bpbsZUU";
// const BATCH_SHEET = "BatchListing";
// // This function handles the routing of your pages
// function doGet(e) {
//   var page = e.parameter.page || "dashboard";
//   var batchId = e.parameter.id || '';         // Get the ID from URL

//   var template = HtmlService.createTemplateFromFile(page);

//   template.autoOpenId = batchId;
//   template.currentPage = page;

//   return template
//     .evaluate()
//     .setTitle("MLM Packaging Web App")
//     .addMetaTag("viewport", "width=device-width, initial-scale=1")
//     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
// }


// function include(filename, currentPage) {
//   var tmp = HtmlService.createTemplateFromFile(filename);
//   tmp.currentPage = currentPage;
//   return tmp.evaluate().getContent();
// }

// /**
//  * RE-DESIGNED BACKEND FOR DESKTOP DASHBOARD
//  * Handles 12 repeating process steps
//  */
// /**
//  * RE-DESIGNED BACKEND
//  * Logic: Progress = (Count of Ticks) / (Count of defined Processes)
//  */


// function getDashboardData() {
//   try {
//     const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

//     // 1. GET DATA FROM JOBLISTING
//     const jobSheet = ss.getSheetByName("JobListing");
//     const jobRaw = jobSheet.getDataRange().getValues();
//     jobRaw.shift();

//     let jobInfoMap = {};
//     jobRaw.forEach((row) => {
//       const psn = String(row[0] || "").trim();
//       if (!psn) return;
//       if (!jobInfoMap[psn]) {
//         jobInfoMap[psn] = {
//           pi: row[1],
//           code: row[2],
//           client: row[3],
//           orderDate: row[6] instanceof Date ? Utilities.formatDate(row[6], "GMT+8", "dd/MM") : row[6],
//           priority: row[9] || "NORMAL",
//           deliveryDate: row[7] instanceof Date ? Utilities.formatDate(row[7], "GMT+8", "dd/MM") : row[7] || "-",
//           status: row[10] || "ON SCHEDULE",
//         };
//       }
//     });

//     // 2. GET DATA FROM BATCHLISTING
//     const batchSheet = ss.getSheetByName("BatchListing");
//     const batchRaw = batchSheet.getDataRange().getValues();
//     const batches = [];

//     for (let r = 16; r < batchRaw.length; r++) {
//       const row = batchRaw[r];
//       const psn = String(row[0] || "").trim();
//       if (!psn || psn.toLowerCase().includes("done")) continue;

//       let definedSteps = [];
//       let ticksFound = 0;
//       let activeStepFound = false;
//       let activeStepStatus = "";

//       const qtyString =
//         row[119] !== undefined && row[119] !== null ? String(row[119]) : "";
//       const START_COL = 6;
//       const BLOCK_SIZE = 9;

//       for (let s = 0; s < 12; s++) {
//         let base = START_COL + s * BLOCK_SIZE;
//         if (base >= row.length) break;
//         let pName = String(row[base] || "").trim();

//         if (pName !== "" && pName !== "--") {
//           // Tick is now at base + 8
//           let isDone =
//             row[base + 8] === true ||
//             String(row[base + 8]).toUpperCase() === "TRUE";

//           if (isDone) ticksFound++;
//           else if (!activeStepFound) {
//             activeStepStatus = String(row[base + 5] || "").trim();
//             activeStepFound = true;
//           }

//           definedSteps.push({
//             name: pName,
//             expDate: row[base + 1]
//               ? Utilities.formatDate(
//                 new Date(row[base + 1]),
//                 "GMT+8",
//                 "dd/MM/yy",
//               )
//               : "-",
//             rawExpDate:
//               row[base + 1] instanceof Date ? row[base + 1].getTime() : null,
//             endDate: row[base + 2]
//               ? Utilities.formatDate(
//                 new Date(row[base + 2]),
//                 "GMT+8",
//                 "dd/MM/yy",
//               )
//               : "-",
//             duration: row[base + 3] || 0,
//             detail: String(row[base + 4] || ""),
//             status: String(row[base + 5] || ""),
//             remark: String(row[base + 6] || ""),
//             revertRemark: String(row[base + 7] || ""),
//             isDone: isDone,
//             baseCol: base,
//           });
//         }
//       }

//       let batchProgress =
//         definedSteps.length > 0 ? ticksFound / definedSteps.length : 0;
//       const info = jobInfoMap[psn] || {
//         pi: "-",
//         code: "-",
//         client: "-",
//         orderDate: "-",
//         priority: "NORMAL",
//         deliveryDate: "-",
//         status: "ON SCHEDULE",
//       };

//       batches.push({
//         row: r + 1,
//         psn: psn,
//         batchId: String(row[1] || ""),
//         // ADDED THIS LINE: Grabs the date from Column C (index 2)
//         batchDate:
//           row[2] instanceof Date
//             ? Utilities.formatDate(row[2], "GMT+8", "dd/MM/yy")
//             : row[2] || "-",
//         jobName: String(row[3] || ""),
//         qty: row[4] || 0,
//         progress: batchProgress,
//         steps: definedSteps,
//         activeStepStatus: activeStepStatus,
//         piNumber: info.pi,
//         salesCode: info.code,
//         clientName: info.client,
//         orderDate: info.orderDate,
//         priority: info.priority,
//         deliveryDate: info.deliveryDate,
//         status: info.status,
//         qtyString: qtyString,
//         splitRemark: String(row[118] || ""),
//       });
//     }

//     // 3. PROCESS STATISTICS (Using the new logic)
//     let processStats = {};
//     let completedTasks = [];

//     batches.forEach(job => {
//       const qty = parseFloat(job.qty) || 0;
//       let foundCurrentLiveStep = false;

//       job.steps.forEach(step => {
//         const name = step.name;
//         const duration = parseFloat(step.duration) || 0;
//         const detail = (step.detail || "").toLowerCase();

//         if (!processStats[name]) {
//           processStats[name] = { totalTime: 0, count: 0, activeCount: 0 };
//         }

//         if (step.isDone) {
//           if (duration > 0) {
//             processStats[name].totalTime += duration;
//             processStats[name].count += 1;
//           }
//           if (step.endDate && step.endDate !== "-") {
//             let machine = "";
//             if (detail.includes("ijima")) machine = "IJIMA";
//             else if (detail.includes("hand switch")) machine = "HANDSWITCH";
//             else if (detail.includes("outsource")) machine = "OUTSOURCED";

//             if (machine) {
//               completedTasks.push({ qty: qty, machine: machine, date: step.endDate });
//             }
//           }
//         }
//         else if (!foundCurrentLiveStep) {
//           processStats[name].activeCount += 1;
//           foundCurrentLiveStep = true;
//         }
//       });
//     });

//     let averageTimes = Object.keys(processStats).map(name => {
//       const s = processStats[name];
//       const avg = s.count > 0 ? s.totalTime / s.count : 0;
//       return {
//         name: name,
//         avgTime: avg < 1 ? avg.toFixed(1) : Math.round(avg),
//         activeCount: s.activeCount
//       };
//     });

//     return {
//       jobs: batches,
//       averages: averageTimes,
//       rawCapacity: completedTasks,
//       today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy")
//     };

//   } catch (err) {
//     console.log("Error: " + err.stack);
//     return { jobs: [], error: err.toString() };
//   }
// }

// function submitData(formData) {
//   const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//   const jobSheet = ss.getSheetByName("JobListing");
//   const batchSheet = ss.getSheetByName("BatchListing");
//   const createDate = new Date();

//   try {
//     // --- 1. DUPLICATE CHECK ---
//     const psnValues = jobSheet.getRange(1, 1, jobSheet.getLastRow(), 1)
//       .getValues()
//       .flat()
//       .map(val => val.toString().trim());

//     const incomingPsn = formData.psn.toString().trim();
//     if (psnValues.includes(incomingPsn)) {
//       return "Duplicate PSN";
//     }

//     // 2. Append to JobListing
//     jobSheet.appendRow([
//       formData.psn, formData.piNumber, formData.salesCode, formData.jobName,
//       formData.jobType, formData.quantity, formData.orderDate,
//       formData.deliveryDate, formData.item, formData.priority
//     ]);

//     // 3. Append to BatchListing
//     formData.batches.forEach((batch) => {
//       let currentBatchId = batch.batchId.toString();
//       if (!currentBatchId.includes("-")) { currentBatchId = currentBatchId + "-1"; }

//       // Basic batch info: PSN, ID, Date, Name, Qty, Status
//       let batchRow = [formData.psn, currentBatchId, createDate, formData.jobName, batch.batchQty, ""];

//       // Initialize 96 columns for steps (12 steps x 8 cols/step)
//       // Note: If you use 9 columns per block, change 96 to 108
//       let stepsData = new Array(108).fill("");

//       batch.steps.forEach((step, index) => {
//         if (index < 12) {
//           let baseIdx = index * 9;
//           stepsData[baseIdx] = step.processName;
//           stepsData[baseIdx + 1] = step.expDate;
//         }
//       });

//       // --- NEW LOGIC: Generate QTY_STRING for Column DP (Index 119) ---
//       // We map the batch quantity to all 12 potential step slots
//       let qtyMap = [];
//       const START_COL = 6; // Column G
//       const BLOCK_SIZE = 9;

//       for (let j = 0; j < 12; j++) {
//         let colIndex = START_COL + (j * BLOCK_SIZE);
//         qtyMap.push(colIndex + ":" + batch.batchQty);
//       }
//       const qtyString = qtyMap.join("|");

//       // We need to ensure batchRow + stepsData reaches Column DP (Col 120)
//       // Batch (6) + Steps (108) = 114 columns. 
//       // We add empty strings to reach index 119 (Column DP)
//       let padding = new Array(119 - (batchRow.length + stepsData.length)).fill("");

//       let finalRow = batchRow.concat(stepsData).concat(padding);
//       finalRow[119] = qtyString; // Set the string at Column DP (120th column)

//       batchSheet.appendRow(finalRow);
//     });

//     updateJobStatus();
//     SpreadsheetApp.flush();

//     return "Success! Job recorded and " + formData.batches.length + " batch(es) created.";
//   } catch (e) {
//     throw new Error(e.toString());
//   }
// }

// /**
//  * Main function to calculate statuses
//  * Compares Delivery Date (Col H) with Today and writes to Col K
//  */
// function updateJobStatus() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getSheetByName("JobListing");
//   if (!sheet) return;

//   const lastRow = sheet.getLastRow();
//   if (lastRow < 2) return;

//   const deliveryDates = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
//   const statusRange = sheet.getRange(2, 11, lastRow - 1, 1);
//   const statusValues = [];

//   const today = new Date();
//   today.setHours(0, 0, 0, 0);

//   for (let i = 0; i < deliveryDates.length; i++) {
//     const dDate = new Date(deliveryDates[i][0]);
//     let status = "";

//     if (deliveryDates[i][0] && !isNaN(dDate.getTime())) {
//       dDate.setHours(0, 0, 0, 0);
//       const diffDays = Math.ceil((dDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));

//       if (diffDays < 0) {
//         status = "OVERDUE";
//       } else if (diffDays <= 7) {
//         status = "ALMOST DUE";
//       } else {
//         status = "ON SCHEDULE";
//       }
//     } else {
//       status = "PENDING DATE";
//     }
//     statusValues.push([status]);
//   }
//   statusRange.setValues(statusValues);
// }

// // Helper function to find PSN and update Delivery Date in JobListing
// function updateJobListingDeliveryDate(sheet, psn, newDate) {
//   const data = sheet.getDataRange().getValues();
//   for (let i = 1; i < data.length; i++) {
//     if (data[i][0] == psn) {
//       // Column A is PSN
//       sheet.getRange(i + 1, 8).setValue(newDate); // Column H is Delivery Date (8)
//       break;
//     }
//   }
//   updateJobStatus();
// }

// function saveMultiBatchUpdate(payload) {
//   const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//   const batchSheet = ss.getSheetByName("BatchListing");
//   const jobSheet = ss.getSheetByName("JobListing");
//   const rowIdx = payload.row;

//   // 1. FRESH DATA FETCH
//   let runningRowData = batchSheet
//     .getRange(rowIdx, 1, 1, batchSheet.getLastColumn())
//     .getValues()[0];

//   const psn = String(runningRowData[0] || "").trim();
//   const today = new Date();
//   const BLOCK_SIZE = 9;
//   const normalize = (v) =>
//     String(v).replace(/\.0$/, "").replace(/\s/g, "").toLowerCase();

//   const existingQtyString = String(runningRowData[119] || "");
//   let qtyMap = {};
//   if (existingQtyString && existingQtyString !== "0") {
//     existingQtyString.split("|").forEach((pair) => {
//       const parts = pair.split(":");
//       if (parts.length === 2) qtyMap[parts[0]] = parts[1];
//     });
//   }

//   const updates = payload.updates.sort((a, b) => a.baseCol - b.baseCol);
//   let trackingQty = Number(runningRowData[4]);

//   updates.forEach((u) => {
//     const currentProcessName = String(runningRowData[u.baseCol] || "").trim();
//     Logger.log(
//       "Row " +
//       rowIdx +
//       " Process Name at Col " +
//       u.baseCol +
//       ": " +
//       currentProcessName,
//     );
//     Logger.log("BaseCol: " + u.baseCol);

//     batchSheet
//       .getRange(rowIdx, u.baseCol + 5 + 1)
//       .setValue(u.isDelayed ? "Delayed" : "");
//     batchSheet
//       .getRange(rowIdx, u.baseCol + 6 + 1)
//       .setValue(u.isDelayed ? u.remark : "");
//     if (u.detail)
//       batchSheet.getRange(rowIdx, u.baseCol + 4 + 1).setValue(u.detail);

//     if (u.isDone) {
//       const currentAvailable = Number(qtyMap[u.baseCol] || trackingQty);
//       let cleanRowSnapshot = [...runningRowData];
//       const targetDate = payload.deliveryDate || payload.newDeliveryDate;

//       // Update local array for subsequent loops
//       runningRowData[u.baseCol + 2] = today;
//       runningRowData[u.baseCol + 8] = true;

//       if (
//         currentProcessName.toLowerCase().includes("delivery") &&
//         targetDate &&
//         targetDate !== "KEEP_ORIGINAL"
//       ) {
//         const jobData = jobSheet.getDataRange().getValues();
//         const deliveryDateObj = new Date(targetDate + "T00:00:00");

//         if (!isNaN(deliveryDateObj.getTime())) {
//           for (let i = 1; i < jobData.length; i++) {
//             // Use the strict normalization from your working code
//             if (normalize(jobData[i][0]) === normalize(psn)) {
//               jobSheet.getRange(i + 1, 8).setValue(deliveryDateObj); // Column H
//               Logger.log("SUCCESS: Updated JobListing for PSN: " + psn);
//               break;
//             }
//           }
//         }
//       }

//       if (u.qty < currentAvailable) {
//         let diff = currentAvailable - u.qty;

//         // CREATE NEW BATCH
//         createSplitBatchFromWaterfall(
//           batchSheet,
//           cleanRowSnapshot,
//           qtyMap, // Send the current map
//           diff,
//           u.baseCol,
//           payload.splitRemark, // The user-typed remark from prompt
//         );

//         // Update Parent tracking
//         trackingQty = u.qty;
//         for (let i = 0; i < 12; i++) {
//           let futureBase = 6 + i * BLOCK_SIZE;
//           if (futureBase >= u.baseCol) qtyMap[futureBase] = trackingQty;
//         }
//       } else {
//         qtyMap[u.baseCol] = u.qty;
//       }

//       batchSheet.getRange(rowIdx, u.baseCol + 2 + 1).setValue(today);
//       batchSheet.getRange(rowIdx, u.baseCol + 8 + 1).setValue(true);
//     }
//   });

//   // 4. FINALIZE & SAVE QTY_STRING
//   // This preserves the high numbers from previous steps because we never
//   // touched keys in qtyMap that were lower than the current u.baseCol
//   const newQtyString = Object.keys(qtyMap)
//     .sort((a, b) => Number(a) - Number(b)) // Keep columns in order
//     .map((k) => k + ":" + qtyMap[k])
//     .join("|");

//   batchSheet.getRange(rowIdx, 120).setValue(newQtyString); // Column DP
//   batchSheet.getRange(rowIdx, 5).setValue(trackingQty); // Column E
// }

// function updateProcessQtysOnly(rowIdx, qtyMapArray) {
//   const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//   const sheet = ss.getSheetByName("BatchListing");

//   const serialized = qtyMapArray
//     .map((obj) => `${obj.baseCol}:${obj.qty}`)
//     .join("|");

//   sheet.getRange(rowIdx, 120).setValue(serialized); // Column DR
//   return "All process quantities synchronized.";
// }

// /**
//  * Helper to create a split batch
//  */

// function createSplitBatchFromWaterfall(
//   sheet,
//   parentData,
//   parentQtyMap,
//   diffQty,
//   splitAtBase,
//   userRemark,
// ) {
//   let newRow = [...parentData];
//   const today = new Date();

//   newRow[1] = generateNewBatchId(sheet, String(parentData[1]));
//   newRow[2] = today;
//   newRow[4] = diffQty; // New smaller quantity

//   // NEW BATCH REMARK -> Column DO -> Index 118
//   newRow[118] = userRemark || "";

//   // 1. CHILD QTY MAP -> Column DP -> Index 119
//   // All steps for this new batch are now capped at its own size
//   let childMap = {};
//   for (let i = 0; i < 12; i++) {
//     let base = 6 + i * 9;
//     childMap[base] = diffQty;
//   }
//   newRow[119] = Object.keys(childMap)
//     .sort((a, b) => a - b)
//     .map((k) => `${k}:${childMap[k]}`)
//     .join("|");

//   // 2. THE RESET
//   const START_COL = 6;
//   const BLOCK_SIZE = 9;
//   for (let i = 0; i < 12; i++) {
//     let colBase = START_COL + i * BLOCK_SIZE;
//     if (colBase >= splitAtBase) {
//       newRow[colBase + 2] = "";
//       newRow[colBase + 3] = "";
//       newRow[colBase + 5] = "";
//       newRow[colBase + 8] = false;
//       newRow[colBase + 4] = "";
//       newRow[colBase + 6] = "";
//     }
//   }

//   sheet.appendRow(newRow);
// }

// /**
//  * Generates a new Batch ID based on existing suffixes.
//  * Example: If J101-1 exists, returns J101-2.
//  */
// function generateNewBatchId(sheet, currentBatchId) {
//   const data = sheet.getDataRange().getValues();
//   const baseId = currentBatchId.split("-")[0]; // Gets "J101" from "J101-1"
//   let maxSuffix = 0;

//   for (let i = 1; i < data.length; i++) {
//     const bId = String(data[i][1]); // Column B
//     if (bId.startsWith(baseId + "-")) {
//       const parts = bId.split("-");
//       const suffix = parseInt(parts[parts.length - 1]);
//       if (!isNaN(suffix) && suffix > maxSuffix) {
//         maxSuffix = suffix;
//       }
//     }
//   }

//   return baseId + "-" + (maxSuffix + 1);
// }

// function revertProcessStep(row, baseCol, revertRemark) {
//   if (!revertRemark || revertRemark.trim() === "") {
//     throw new Error("Revert remark is mandatory.");
//   }

//   try {
//     const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//     const sheet = ss.getSheetByName("BatchListing");

//     const START_COL = 6;
//     const BLOCK_SIZE = 9;
//     const TOTAL_STEPS = 12;

//     // Fetch the entire row once to check "isDone" status of future steps
//     const rowValues = sheet
//       .getRange(row, 1, 1, sheet.getLastColumn())
//       .getValues()[0];

//     for (let i = 0; i < TOTAL_STEPS; i++) {
//       let currentStepBase = START_COL + i * BLOCK_SIZE;

//       if (currentStepBase >= baseCol) {
//         let pName = rowValues[currentStepBase];
//         if (!pName || pName === "" || pName === "--") continue;

//         // Check if this step was currently completed (Tick is at base + 8)
//         let wasDone =
//           rowValues[currentStepBase + 8] === true ||
//           String(rowValues[currentStepBase + 8]).toUpperCase() === "TRUE";

//         // 1. Clear core completion data regardless
//         sheet.getRange(row, currentStepBase + 2 + 1).clearContent(); // End Date
//         sheet.getRange(row, currentStepBase + 3 + 1).clearContent(); // Duration
//         sheet.getRange(row, currentStepBase + 6 + 1).clearContent(); // Completion Remark
//         sheet.getRange(row, currentStepBase + 8 + 1).setValue(false); // Untick

//         // 2. Handle Remarks/Status
//         if (currentStepBase === baseCol) {
//           // The target step
//           sheet.getRange(row, currentStepBase + 5 + 1).setValue("Reverted");
//           sheet.getRange(row, currentStepBase + 7 + 1).setValue(revertRemark);
//         } else if (wasDone) {
//           // Sequential steps that were actually finished
//           sheet.getRange(row, currentStepBase + 5 + 1).clearContent();
//           sheet
//             .getRange(row, currentStepBase + 7 + 1)
//             .setValue("Auto-reverted (Sequential)");
//         } else {
//           // Steps that were already pending: just clear status and remark to be safe
//           sheet.getRange(row, currentStepBase + 5 + 1).clearContent();
//           sheet.getRange(row, currentStepBase + 7 + 1).clearContent();
//         }
//       }
//     }
//     return { success: true };
//   } catch (err) {
//     throw new Error(err.toString());
//   }
// }


// function backfillQtyStrings() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getSheetByName("BatchListing");
//   const data = sheet.getDataRange().getValues();

//   const START_COL = 6; // Column G
//   const BLOCK_SIZE = 9;
//   const QTY_COL_IDX = 4; // Column E (0-based)
//   const QTY_STRING_COL = 120; // Column DR

//   const updates = [];

//   // Skip header row
//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];
//     const currentQty = row[QTY_COL_IDX];
//     const existingString = row[119]; // Column DR (0-based)

//     // Only backfill if the string is empty or "0"
//     if (!existingString || existingString === "0" || existingString === "") {
//       let qtyMap = [];
//       for (let j = 0; j < 12; j++) {
//         let base = START_COL + j * BLOCK_SIZE;
//         qtyMap.push(base + ":" + currentQty);
//       }
//       const serialized = qtyMap.join("|");

//       // Store the range and value to update
//       sheet.getRange(i + 1, QTY_STRING_COL).setValue(serialized);
//     }
//   }

//   console.log("Backfill complete.");
// }

