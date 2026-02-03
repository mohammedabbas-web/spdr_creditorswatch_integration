const cron = require('node-cron');
import moment from 'moment';
import { SmartsheetService } from '../services/SmartsheetServices/SmartsheetServices';
import { ScheduleDeletionCheckResultType } from '../types/smartsheet.types';

// Get scheduler configuration from environment variables
const scheduleDeletedCheckCronExpression = process.env.SCHEDULE_DELETED_CHECK_CRON_EXPRESSION || '0 0 * * *'; // Default: Daily at midnight
const isSchedulerEnabled = process.env.ENABLE_SCHEDULE_DELETED_CHECK_SCHEDULER === 'true';

console.log(`🔧 ROOFING SCHEDULE DELETED CHECK (FROM DB) SCHEDULER: Initialized`);
console.log(`   Cron Expression: ${scheduleDeletedCheckCronExpression}`);
console.log(`   Enabled: ${isSchedulerEnabled}`);

/**
 * Executes the roofing schedule deletion check
 * This function queries both roofing schedule sheets and validates each schedule's existence in SimPro
 * Updates the IsDeleted column accordingly
 */
export const executeRoofingScheduleDeletedCheck = async (): Promise<ScheduleDeletionCheckResultType | { status: string; message: string; error: any }> => {
    try {
        const startTime = moment().format('YYYY-MM-DD HH:mm:ss');
        console.log(`\n⏱️  ROOFING SCHEDULE DELETED CHECK (FROM DB) SCHEDULER: Task execution started at ${startTime}`);

        // Call the Smartsheet service to check and update schedule deletion status
        const result: ScheduleDeletionCheckResultType = await SmartsheetService.checkAndUpdateScheduleDeletionStatus();

        const endTime = moment().format('YYYY-MM-DD HH:mm:ss');
        console.log(`✅ ROOFING SCHEDULE DELETED CHECK (FROM DB) SCHEDULER: Task execution completed at ${endTime}`);

        // Log summary
        console.log(`\n📊 ROOFING SCHEDULE DELETED CHECK (FROM DB) SCHEDULER: Summary`);
        console.log(`   Total Schedules Checked: ${result.totalSchedulesChecked}`);
        console.log(`   Deleted Schedules Found: ${result.deletedSchedulesFound}`);
        console.log(`   Active Schedules Found: ${result.activeSchedulesFound}`);
        console.log(`   Rows Updated: ${result.updatedRows.length}`);
        console.log(`   Errors Encountered: ${result.erroredSchedules.length}`);

        if (result.erroredSchedules.length > 0) {
            console.warn(`⚠️  ROOFING SCHEDULE DELETED CHECK (FROM DB) SCHEDULER: Some errors occurred:`);
            result.erroredSchedules.forEach((error, index) => {
                console.warn(`   [${index + 1}] Row ${error.rowId}: ${error.error}`);
            });
        }

        return result;
    } catch (err) {
        const errorTime = moment().format('YYYY-MM-DD HH:mm:ss');
        console.error(`\n❌ ROOFING SCHEDULE DELETED CHECK (FROM DB) SCHEDULER: Error occurred at ${errorTime}`);
        console.error(`   Error Details:`, err);

        // Continue execution even on error - don't break the scheduler
        if (err instanceof Error) {
            console.error(`   Message: ${err.message}`);
            console.error(`   Stack: ${err.stack}`);
        }

        throw {
            message: 'Error in roofing schedule deletion check scheduler',
            error: err,
            timestamp: errorTime,
        };
    }
};

/**
 * Initialize the cron job if enabled in environment
 */
if (isSchedulerEnabled) {
    console.log(`\n🚀 ROOFING SCHEDULE DELETED CHECK (FROM DB) SCHEDULER: Activating cron job with expression: ${scheduleDeletedCheckCronExpression}`);

    cron.schedule(scheduleDeletedCheckCronExpression, async () => {
        try {
            await executeRoofingScheduleDeletedCheck();
        } catch (err) {
            console.error('ROOFING SCHEDULE DELETED CHECK (FROM DB) SCHEDULER: Cron job execution failed', err);
            // Don't throw - let the scheduler continue running for next interval
        }
    });

    console.log(`✅ ROOFING SCHEDULE DELETED CHECK (FROM DB) SCHEDULER: Cron job activated successfully`);
} else {
    console.log(`⚠️  ROOFING SCHEDULE DELETED CHECK (FROM DB) SCHEDULER: Scheduler is disabled. Set ENABLE_SCHEDULE_DELETED_CHECK_SCHEDULER=true to enable`);
}

export default executeRoofingScheduleDeletedCheck;
