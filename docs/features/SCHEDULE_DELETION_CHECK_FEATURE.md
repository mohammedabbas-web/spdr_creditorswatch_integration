# Schedule Deletion Check Feature

## Overview
This feature validates if schedules exist in SimPro API and updates the deletion status in Smartsheet. It provides both manual API triggers and automated scheduler options.

## Environment Variables

### Required Environment Variables

#### Smartsheet Configuration (Reused)
```env
# Smartsheet API Token
SMARTSHEET_ACCESS_TOKEN=<your-smartsheet-token>
```

#### SimPro Configuration (Reused)
```env
# SimPro API Base URL
SIMPRO_BASE_URL=https://api.simpro.com/v1.0

# SimPro Company ID
SIMPRO_COMPANY_ID=<your-company-id>

# SimPro API Access Token
SIMPRO_ACCESS_TOKEN=<your-simpro-token>
```

#### Roofing Schedules - Smartsheet Sheets
```env
# Roofing Schedules Active Sheet ID
# This sheet contains active roofing schedules ("Roofing Schedules from DB")
ROOFING_SCHEDULES_ACTIVE_FROM_DB_SHEET_ID=<your-active-sheet-id>

# Roofing Schedules Archived Sheet ID
# This sheet contains archived/moved past roofing schedules ("Roofing Schedule from DB Move Past")
ROOFING_SCHEDULES_ARCHIVED_FROM_DB_SHEET_ID=<your-archived-sheet-id>
```

### Optional Environment Variables

#### Scheduler Configuration
```env
# Enable/disable the automated scheduler
# Default: false (scheduler is disabled)
ENABLE_SCHEDULE_DELETED_CHECK_SCHEDULER=true

# Cron expression for scheduler frequency
# Format: "minute hour day month dayOfWeek"
# Default: "0 0 * * *" (Daily at midnight)
# Examples:
#   "0 0 * * *"      - Every day at 00:00 (midnight)
#   "0 */6 * * *"    - Every 6 hours (00:00, 06:00, 12:00, 18:00)
#   "0 0 * * 0"      - Every Sunday at 00:00
#   "0 2 * * 1-5"    - Every weekday at 02:00
SCHEDULE_DELETED_CHECK_CRON_EXPRESSION=0 0 * * *
```

## API Endpoints

### Manual Schedule Deletion Check

#### Check All Schedules
```
GET /api/smartsheet/check-schedule-deletion
```

**Description:** Validates all schedules in both configured sheets (active and archived)

**Uses environment variables:** (Only if sheet IDs are NOT provided as parameters)
- `ROOFING_SCHEDULES_ACTIVE_FROM_DB_SHEET_ID`
- `ROOFING_SCHEDULES_ARCHIVED_FROM_DB_SHEET_ID`

**Note:** This endpoint can work WITHOUT environment variables if sheet IDs are passed as URL parameters. However, it's optional to provide them - the scheduler requires them.

**Response:**
```json
{
  "status": true,
  "message": "Schedule deletion check completed successfully",
  "data": {
    "totalSchedulesChecked": 150,
    "deletedSchedulesFound": 5,
    "activeSchedulesFound": 145,
    "updatedRows": [123, 456, 789],
    "erroredSchedules": []
  }
}
```

#### Check Specific Schedule
```
GET /api/smartsheet/check-schedule-deletion?scheduleId=12345
```

**Parameters:**
- `scheduleId` (query): The Schedule ID to validate

**Uses environment variables:** (Only if sheet IDs are NOT provided as parameters)
- `ROOFING_SCHEDULES_ACTIVE_FROM_DB_SHEET_ID`
- `ROOFING_SCHEDULES_ARCHIVED_FROM_DB_SHEET_ID`

**Response:**
```json
{
  "status": true,
  "message": "Schedule deletion check completed successfully",
  "data": {
    "totalSchedulesChecked": 1,
    "deletedSchedulesFound": 0,
    "activeSchedulesFound": 1,
    "updatedRows": [789],
    "erroredSchedules": []
  }
}
```

#### Override Sheet IDs (Advanced)
```
GET /api/smartsheet/check-schedule-deletion?activeSheetId=xxxxx&archivedSheetId=yyyyy
```

**Parameters:**
- `activeSheetId` (query): Override default active sheet ID from environment (ROOFING_SCHEDULES_ACTIVE_FROM_DB_SHEET_ID) - OPTIONAL
- `archivedSheetId` (query): Override default archived sheet ID from environment (ROOFING_SCHEDULES_ARCHIVED_FROM_DB_SHEET_ID) - OPTIONAL
- `scheduleId` (query, optional): Specific schedule to validate

**Note:** If you provide sheet ID parameters, you do NOT need the environment variables for the API endpoint.

## Smartsheet Sheet Structure

The feature expects the following columns in both sheets:

| Column Name | Type | Required | Description |
|---|---|---|---|
| ID-Schedule | Text | Yes | Unique schedule identifier |
| ID-Job | Number | Yes | Job ID from SimPro |
| ID-Section | Number | Yes | Section ID from SimPro |
| ID-CostCentre | Number | Yes | Cost Center ID from SimPro |
| ISDeleted | Text | Yes | Deletion status (Yes/No) |

## How It Works

### Schedule Validation Process

1. **Fetch Sheet Rows**: Retrieves all rows from active and archived schedule sheets
2. **Extract Required Fields**: Gets Schedule ID, Job ID, Section ID, and Cost Center ID from each row
3. **Validate in SimPro**: For each schedule, makes an API call to SimPro:
   - Endpoint: `/jobs/{jobID}/sections/{sectionID}/costCenters/{costCenterID}/schedules/{scheduleID}`
   - Success (200): Schedule exists → `ISDeleted = "No"`
   - Not Found (404) or "not found" error: Schedule deleted → `ISDeleted = "Yes"`
4. **Update Smartsheet**: Batches updates (max 300 rows per batch) and applies to sheets
5. **Log Results**: Outputs detailed summary with success and error information

### Environment Variables Configuration

#### For API Endpoint
- **Optional**: `ROOFING_SCHEDULES_ACTIVE_FROM_DB_SHEET_ID` and `ROOFING_SCHEDULES_ARCHIVED_FROM_DB_SHEET_ID`
  - Only needed if you want to use environment variables
  - Can be bypassed by passing sheet IDs as URL parameters
  - Provides convenience for standard operations

#### For Scheduler
- **Required**: `ROOFING_SCHEDULES_ACTIVE_FROM_DB_SHEET_ID` and `ROOFING_SCHEDULES_ARCHIVED_FROM_DB_SHEET_ID`
  - Scheduler has no URL parameters to pass, so these MUST be set in environment variables
  - Without these, scheduler cannot run

This design allows flexibility:
- **API users** can pass sheet IDs dynamically via URL parameters
- **Scheduler** relies on environment configuration for consistency and automation

### Error Handling

- **Missing Required Columns**: Logs warning and skips the sheet
- **Missing Row Fields**: Logs and skips individual rows
- **API Failures**: Records in `erroredSchedules` array, continues processing
- **Batch Update Failures**: Throws error with details
- **Network Issues**: Handled by axios retry logic (3 retries with exponential backoff)
- **Schedule Not Found (404)**: Correctly marks schedule as deleted

## Usage Examples

### Using cURL

#### Check all schedules
```bash
curl -X GET http://localhost:6001/api/smartsheet/check-schedule-deletion
```

#### Check specific schedule
```bash
curl -X GET "http://localhost:6001/api/smartsheet/check-schedule-deletion?scheduleId=12345"
```

#### Override sheet IDs
```bash
curl -X GET "http://localhost:6001/api/smartsheet/check-schedule-deletion?activeSheetId=123&archivedSheetId=456"
```

### Using Postman

1. Create new GET request
2. URL: `{{baseUrl}}/api/smartsheet/check-schedule-deletion`
3. Params:
   - `scheduleId` (optional): `12345`
   - `activeSheetId` (optional): `123`
   - `archivedSheetId` (optional): `456`
4. Send

### Automated Scheduler

Enable in environment:
```env
ENABLE_SCHEDULE_DELETED_CHECK_SCHEDULER=true
SCHEDULE_DELETED_CHECK_CRON_EXPRESSION=0 0 * * *
ROOFING_SCHEDULES_ACTIVE_FROM_DB_SHEET_ID=<sheet-id>
ROOFING_SCHEDULES_ARCHIVED_FROM_DB_SHEET_ID=<sheet-id>
```

The scheduler will automatically run at the configured time and update deletion status for all schedules.

## Logging and Monitoring

The feature provides detailed console logging at each step:

- 🔍 Validation process start
- 📋 Sheet processing
- 📄 Row processing
- ✅ Active schedules found
- ❌ Deleted schedules found
- 📝 Batch updates
- 📊 Final summary
- ⚠️ Errors and warnings

All logs include timestamps for easy tracking in production logs.

## Performance Considerations

- **Batch Size**: Updates are batched in groups of 300 rows (Smartsheet API limit)
- **Rate Limiting**: SimPro API calls have rate limiting (5 requests per 1000ms)
- **Retry Logic**: Automatic retries for network/server errors (3 attempts)
- **Timeout**: 10-minute timeout per API request
- **Test Mode**: During development, set TEST_MODE=true to process only first 20 rows

## Troubleshooting

### Scheduler Not Running
- Check: `ENABLE_SCHEDULE_DELETED_CHECK_SCHEDULER=true`
- Check: Cron expression is valid
- Check: Application is running in production environment
- Check: Both sheet ID environment variables are set

### No Rows Updated
- Verify required columns exist in Smartsheet: ID-Schedule, ID-Job, ID-Section, ID-CostCentre, ISDeleted
- Check column names match exactly (case-sensitive)
- Verify rows have values in all required columns

### API Errors
- 401: Check SimPro and Smartsheet tokens
- 404: Verify Job/Section/CostCenter IDs are valid
- 429: Rate limit exceeded - wait before retrying

### Missing Environment Variables
- Check `.env.production` or `.env.development` files
- Ensure all required variables are set:
  - `ROOFING_SCHEDULES_ACTIVE_FROM_DB_SHEET_ID`
  - `ROOFING_SCHEDULES_ARCHIVED_FROM_DB_SHEET_ID`
  - `SMARTSHEET_ACCESS_TOKEN`
  - `SIMPRO_BASE_URL`
  - `SIMPRO_COMPANY_ID`
  - `SIMPRO_ACCESS_TOKEN`
  - (Optional) `ENABLE_SCHEDULE_DELETED_CHECK_SCHEDULER`
  - (Optional) `SCHEDULE_DELETED_CHECK_CRON_EXPRESSION`
- Restart application after changing environment variables

## Architecture

The feature is implemented across multiple layers:

### Service Layer
- **SimproScheduleService**: `validateScheduleExistence()` - Validates schedule in SimPro API
- **SmartsheetService**: `checkAndUpdateScheduleDeletionStatus()` - Orchestrates validation and updates

### Controller Layer
- **SmartSheetController**: `checkScheduleDeletionStatus()` - Handles HTTP requests

### Route Layer
- **smartSheetRoutes**: `GET /check-schedule-deletion` - Exposes API endpoint

### Scheduler Layer
- **roofingScheduleDeletedCheckfromDBScheduler**: `executeRoofingScheduleDeletedCheck()` - Automated execution

### Type Layer
- Comprehensive TypeScript types for type safety

### Files Modified/Created
1. `src/services/SimproServices/simproScheduleService.ts` - Schedule validation
2. `src/services/SmartsheetServices/SmartsheetServices.ts` - Deletion check & updates
3. `src/controllers/smartSheetController.ts` - API endpoint handler
4. `src/routes/smartSheetRoutes.ts` - Route registration
5. `src/cron/roofingScheduleDeletedCheckfromDBScheduler.ts` - Scheduler
6. `src/types/smartsheet.types.d.ts` - TypeScript types
7. `src/index.ts` - Scheduler integration

## Future Enhancements

Potential improvements:
- Batch validation of schedules (reduce API calls)
- Webhook notifications on deletion detection
- Custom filtering (by date, job, section)
- Detailed audit trail of changes
- Performance metrics/analytics
- Dashboard for monitoring deletion checks
