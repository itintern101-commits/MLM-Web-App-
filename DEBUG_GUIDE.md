# Debug Guide - Node.js + SharePoint Migration

## Console Log Points Added

### Backend (index.js)
1. **[API]** - All API endpoints log request/response
2. **[generateDashboardData]** - Logs JobListing and BatchListing fetch
3. **[getTableRowsAsObjects]** - Logs table headers and row count
4. **[addTableRow]** - Logs when row is added to table
5. **[getSharePointFileContext]** - Logs authentication flow (implicit)

### Frontend (dashboard.html)
1. **Dashboard: Starting data fetch...** - Fetch initiated
2. **Dashboard: Fetched data:** - Raw data from API
3. **Dashboard: jobListing count:** - Count of jobs
4. **Dashboard: batchListing count:** - Count of batches
5. **Dashboard: MASTER_DATA:** - Processed data structure

### Frontend (BatchDetail.html)
1. **BatchDetail: Starting data fetch...** - Fetch initiated
2. **BatchDetail: Fetched data:** - Raw data from API
3. **BatchDetail: jobListing:** - Job listing data
4. **BatchDetail: MASTER_DATA set to:** - Set data reference
5. **renderStats: Processing** - Stats calculation
6. **renderStats: Job [idx]: [data]** - Individual job processing
7. **renderStats: Job has no steps array** - Defensive check for missing steps

## How to Check Data Flow

### 1. Server Console (Terminal)
```
Watch for:
- [API] GET /api/dashboard - request received
- [generateDashboardData] Starting...
- [getTableRowsAsObjects] Fetching JobListing from SharePoint
- Column headers fetched
- Row count
- Data conversion complete
```

### 2. Browser Console (F12)
```
Watch for:
Dashboard:
- "Dashboard: Starting data fetch..."
- "Dashboard: Fetched data:" (check JSON structure)
- "Dashboard: jobListing count: X"
- "Dashboard: batchListing count: Y"

BatchDetail:
- "BatchDetail: Starting data fetch..."
- "BatchDetail: Fetched data:"
- "BatchDetail: MASTER_DATA set to:"
- "renderStats: Processing X items"
```

## Data Structure Expected

### From SharePoint JobListing Table
```json
{
  "PSN": "J001",
  "PI_Number": "PI-123",
  "Sales_Code": "SC-001",
  "Job_Name": "Job Name",
  "Job_Type": "Type",
  "Quantity": 100,
  "Order_Date": "2025-01-01",
  "Delivery_Date": "2025-01-15",
  "Item": "Item Name",
  "Priority": "Normal"
}
```

### From SharePoint BatchListing Table
```json
{
  "PSN": "J001",
  "Batch_ID": "J001-1",
  "Created_Date": "2025-01-01",
  "Job_Name": "Job Name",
  "Quantity": 100,
  "Status": "Active"
}
```

## Common Issues & Solutions

### Issue: "Cannot read properties of undefined (reading 'find')"
- **Cause**: `job.steps` is undefined
- **Solution**: Defensive checks added - only logs jobs with steps
- **Check**: Look for "renderStats: Job has no steps array"

### Issue: "Invalid left-hand side in assignment"
- **Cause**: Google Apps Script template tag `<?=...?>`
- **Status**: ✅ FIXED - Replaced with simple `/BatchDetail.html?id=X`

### Issue: "Failed to fetch dashboard data"
- **Check**: Open browser console → Network tab → see response
- **Server logs**: Check terminal for [API] error messages

## Next Steps

1. Open http://localhost:3000/dashboard.html
2. Open browser console (F12)
3. Check for all debug logs
4. Compare data shape with expected structure
5. Report any data structure mismatches
