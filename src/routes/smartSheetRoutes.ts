import express from 'express';
const router = express.Router();

import { handleSmartSheetWebhook, handleSmartSheetWebhookPost, updateAmountValuesInRoofingWipSheet, updateSuburbDataForSite, checkScheduleDeletionStatus } from '../controllers/smartSheetController';

// To update myob data to smartsheet, this is manual API call, that can be sent from postman or frontend whenever it needs to update the myob data to the smartsheet.
router.get("/webhooks", handleSmartSheetWebhook);

// This route handles the webhook API trigger from smartsheet.
router.post("/webhooks", handleSmartSheetWebhookPost);

router.put('/update-site-suburb',updateSuburbDataForSite);
router.put('/update-amount-values-in-wip', updateAmountValuesInRoofingWipSheet)

// Schedule deletion check API
// GET /api/smartsheet/check-schedule-deletion (checks all schedules)
// GET /api/smartsheet/check-schedule-deletion?scheduleId=12345 (checks specific schedule)
router.get('/check-schedule-deletion', checkScheduleDeletionStatus);


export default router;  // Use default export
