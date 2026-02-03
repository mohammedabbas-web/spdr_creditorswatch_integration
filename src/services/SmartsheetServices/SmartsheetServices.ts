import { AxiosError } from "axios";
import axiosSimPRO from "../../config/axiosSimProConfig";
import { CostCenterJobInfo, SimproAccountType, SimproContractorJobType, SimproCostCenter, SimproCostCenterType, SimproJobCostCenterType, SimproJobCostCenterTypeForAmountUpdate, SimproJobType, SimproScheduleType, SimproWebhookType } from "../../types/simpro.types";
import { SimproContractorWorkOrderType, SmartsheetColumnType, SmartsheetSheetRowsType } from "../../types/smartsheet.types";
import { convertSimproContractorDataToSmartsheetFormat, convertSimproContractorJobDataToSmartsheetFormatForUpdate, convertSimprocostCenterDataToSmartsheetFormatForUpdate, convertSimproRoofingDataToSmartsheetFormat, convertSimproScheduleDataToSmartsheetFormat, convertSimproScheduleDataToSmartsheetFormatForUpdate } from "../../utils/transformSimproToSmartsheetHelper";
import { fetchSimproPaginatedData } from "../SimproServices/simproPaginationService";
import { validateScheduleExistence } from "../SimproServices/simproScheduleService";
import { extractLineItemsDataFromContractorJob, splitIntoChunks } from "../../utils/helper";
const SmartsheetClient = require('smartsheet');
const smartSheetAccessToken: string | undefined = process.env.SMARTSHEET_ACCESS_TOKEN;
const smartsheet = SmartsheetClient.createClient({ accessToken: smartSheetAccessToken });
const jobCardV2SheetId = process.env.JOB_CARD_SHEET_V2_ID ? process.env.JOB_CARD_SHEET_V2_ID : "";
const jobCardRoofingDetailSheetId = process.env.JOB_CARD_SHEET_ROOFING_DETAIL_ID ?? "";
const wipJobArchivedSheetId = process.env.WIP_JOB_ARCHIVED_SHEET_ID ?? "";
const jobCardV2MovePastSheetId = process.env.JOB_CARD_V2_MOVE_PAST_SHEET_ID ?? "";
const workOrderLineItemsActiveSheetId = process.env.WORKORDER_LINE_ITEMS_ACTIVE_SHEET_ID ?? "";
const workOrderLineItemsArchivedSheetId = process.env.WORKORDER_LINE_ITEMS_ARCHIVED_SHEET_ID ?? "";

// Roofing Schedules - Dedicated Sheet IDs
const roofingSchedulesActiveFromDbSheetId = process.env.ROOFING_SCHEDULES_ACTIVE_FROM_DB_SHEET_ID ?? "";
const roofingSchedulesArchivedFromDbSheetId = process.env.ROOFING_SCHEDULES_ARCHIVED_FROM_DB_SHEET_ID ?? "";
export class SmartsheetService {


    static async handleAddUpdateScheduleToSmartsheet(webhookData: SimproWebhookType) {
        try {
            let isInvoiceAccountNameRoofing = false;
            const { scheduleID, jobID, sectionID, costCenterID } = webhookData.reference;
            let fetchedChartOfAccounts = await axiosSimPRO.get('/setup/accounts/chartOfAccounts/?pageSize=250&columns=ID,Name,Number');
            let chartOfAccountsArray: SimproAccountType[] = fetchedChartOfAccounts?.data;
            console.log('scheduleID, jobID, sectionID, costCenterID', scheduleID, jobID, sectionID, costCenterID)
            // /api/v1.0/companies/{companyID}/jobs/{jobID}/sections/{sectionID}/costCenters/{costCenterID}/schedules/{scheduleID}
            let simPROScheduleUpdateUrl = `/jobs/${jobID}/sections/${sectionID}/costCenters/${costCenterID}/schedules/${scheduleID}`;
            console.log('simPROScheduleUpdateUrl', simPROScheduleUpdateUrl)
            let individualScheduleResponse = await axiosSimPRO(`${simPROScheduleUpdateUrl}?columns=ID,Staff,Date,Blocks,Notes`)
            let jobForScheduleResponse = await axiosSimPRO(`/jobs/${jobID}?columns=ID,Type,Site,SiteContact,DateIssued,Status,Total,Customer,Name,ProjectManager,CustomFields,Totals`)
            let schedule: SimproScheduleType = individualScheduleResponse?.data;
            console.log('Shceuld Blocks', schedule.Blocks)
            let fetchedJobData: SimproJobType = jobForScheduleResponse?.data;
            let siteId = fetchedJobData?.Site?.ID;
            if (siteId) {
                const siteResponse = await axiosSimPRO.get(`/sites/${siteId}?columns=ID,Name,Address`);
                let siteResponseData = siteResponse.data;
                fetchedJobData.Site = siteResponseData;
            }

            schedule.Job = fetchedJobData;
            if (schedule?.Job?.Customer) {
                const customerId = schedule.Job.Customer.ID?.toString();
                try {
                    // This will always fail
                    const customerResponse = await axiosSimPRO.get(`/customers/${customerId}`)
                    // console.log("customerResponse: ", customerResponse)
                } catch (err: any) {
                    // console.log("Error getting customer", )
                    let endpoint = err?.response?.data?._href;

                    // Extract the part starting from "/customers"
                    const startFromCustomers = endpoint.substring(endpoint.indexOf("/customers"));

                    console.log("startFromCustomers: ", startFromCustomers);

                    // Check if it's "/customers/companies" or "/customers/individuals"
                    if (startFromCustomers.includes("/companies")) {
                        // Handle the case for companies
                        const companyResponse = await axiosSimPRO.get(
                            `${startFromCustomers}?columns=ID,CompanyName,Phone,Address,Email`
                        );
                        console.log("Company Response: ", companyResponse?.data);
                        schedule.Job.Customer = { ...companyResponse.data, Type: "Company" };
                    } else if (startFromCustomers.includes("/individuals")) {
                        // Handle the case for individuals
                        const individualResponse = await axiosSimPRO.get(
                            `${startFromCustomers}?columns=ID,GivenName,FamilyName,Phone,Address,Email`
                        );
                        console.log("Individual Response: ", individualResponse?.data);
                        schedule.Job.Customer = { ...individualResponse.data, Type: "Individual" };
                    } else {
                        console.error("Unknown customer type in the endpoint");
                    }
                }
            }


            // console.log('costCenterDataForSchedule', `/jobCostCenters/?ID=${costCenterID}&columns=ID,Name,Job,Section,CostCenter`)
            const costCenterDataForSchedule = await axiosSimPRO.get(`/jobCostCenters/?ID=${costCenterID}&columns=ID,Name,Job,Section,CostCenter`);
            let setupCostCenterID = costCenterDataForSchedule.data[0]?.CostCenter?.ID;
            let fetchedSetupCostCenterData = await axiosSimPRO.get(`/setup/accounts/costCenters/${setupCostCenterID}?columns=ID,Name,IncomeAccountNo`);
            let setupCostCenterData = fetchedSetupCostCenterData.data;
            console.log('CostCenterId IncomeAccountNo', costCenterID, setupCostCenterData);

            if (setupCostCenterData?.IncomeAccountNo) {
                let incomeAccountName = chartOfAccountsArray?.find(account => account?.Number == setupCostCenterData?.IncomeAccountNo)?.Name;
                console.log("Income Account Name: " + incomeAccountName)
                if (incomeAccountName == "Roofing Income") {
                    isInvoiceAccountNameRoofing = true;
                }
            }

            let costCenterResponse = await axiosSimPRO.get(`jobs/${jobID}/sections/${sectionID}/costCenters/${costCenterID}?columns=Name,ID,Claimed,Total,Totals`);
            
            if (costCenterResponse) {
                console.log('CostCenterResponse Data', costCenterResponse.data)
                schedule.CostCenter = costCenterResponse.data;
            }

            // console.log('IsInvoiceAccountNameRoofing: ' + isInvoiceAccountNameRoofing, costCenterID, scheduleID)
            if (isInvoiceAccountNameRoofing) {
                
                if (jobCardV2SheetId) {
                    console.log("Adding/Updating schedule in Job Card V2 Sheet for schedule ID: ", schedule.ID,"in sheet ID: ", jobCardV2SheetId);
                    const sheetInfo = await smartsheet.sheets.getSheet({ id: jobCardV2SheetId });
                    const columns = sheetInfo.columns;
                    const column = columns.find((col: SmartsheetColumnType) => col.title === "ScheduleID");

                    console.log("schedule column in v2", column)
                    if (!column) {
                        throw {
                            message: "ScheduleID column not found in the sheet",
                            status: 400
                        }
                    }

                    const scheduleIdColumnId = column.id;
                    const existingRows: SmartsheetSheetRowsType[] = sheetInfo.rows;
                    let scheduleDataForSmartsheet: SmartsheetSheetRowsType | undefined;
                    console.log('existingRows', existingRows.length)
                    for (let i = 0; i < existingRows.length; i++) {
                        let currentRow = existingRows[i];
                        const cellData = currentRow.cells.find(
                            (cell: { columnId: string; value: any }) => cell.columnId === scheduleIdColumnId
                        );
                        if (cellData?.value === schedule.ID) {
                            scheduleDataForSmartsheet = currentRow;
                            break;
                        }
                    }

                    console.log('scheduleDataForSmartsheet', scheduleDataForSmartsheet)
                    if (scheduleDataForSmartsheet) {
                        let rowIdMap: { [key: string]: string } = {};
                        rowIdMap = {
                            [schedule.ID.toString()]: scheduleDataForSmartsheet?.id?.toString() || "",
                        };
                        const convertedData = convertSimproScheduleDataToSmartsheetFormatForUpdate([schedule], columns, rowIdMap, 'full');

                        await smartsheet.sheets.updateRow({
                            sheetId: jobCardV2SheetId,
                            body: convertedData,
                        });
                        console.log('Updated row in smartsheet in sheet active ', jobCardV2SheetId)
                    } else {
                        // Add logic to check the row in  move past sheet
                        console.log("Schedule not found in the main sheet, checking Move Past Sheet", schedule.ID, "in sheet ID move past: ", jobCardV2MovePastSheetId)
                        let movePastSheetInfo = await smartsheet.sheets.getSheet({ id: jobCardV2MovePastSheetId });
                        const movePastSheetColumns = movePastSheetInfo.columns;
                        const movePastScheduleColumn = movePastSheetColumns.find((col: SmartsheetColumnType) => col.title === "ScheduleID");
                        const movePastScheduleIdColumnId = movePastScheduleColumn.id;
                        let movePastScheduleDataForSmartsheet: SmartsheetSheetRowsType | undefined;
                        const existingMovePastRows: SmartsheetSheetRowsType[] = movePastSheetInfo.rows;
                        for (let i = 0; i < existingMovePastRows.length; i++) {
                            let currentRow = existingMovePastRows[i];
                            const cellData = currentRow.cells.find(
                                (cell: { columnId: string; value: any }) => cell.columnId === movePastScheduleIdColumnId
                            );
                            if (cellData?.value === schedule.ID) {
                                movePastScheduleDataForSmartsheet = currentRow;
                                break;
                            }
                        }

                        if (movePastScheduleDataForSmartsheet) {
                            let rowIdMap: { [key: string]: string } = {};
                            rowIdMap = {
                                [schedule.ID.toString()]: movePastScheduleDataForSmartsheet?.id?.toString() || "",
                            };
                            const convertedData = convertSimproScheduleDataToSmartsheetFormatForUpdate([schedule], movePastSheetColumns, rowIdMap, 'full');

                            await smartsheet.sheets.updateRow({
                                sheetId: jobCardV2MovePastSheetId,
                                body: convertedData,
                            });
                            console.log('Updated row in smartsheet in Move Past Sheet ', jobCardV2MovePastSheetId)
                        } else {
                            console.log("Schedule not found in Move Past Sheet, adding new row for schedule ", schedule.ID)
                            const convertedDataForSmartsheet = convertSimproScheduleDataToSmartsheetFormat([schedule], columns, 'full');
                            await smartsheet.sheets.addRows({
                                sheetId: jobCardV2SheetId,
                                body: convertedDataForSmartsheet,
                            });
                            console.log('Added row in smartsheet in sheeet', jobCardV2SheetId)
                        }



                    }
                }
            }

            // console.log('schedule', schedule)
        } catch (err) {
            console.log("Error in the update schedule simpro webhook", err);
            throw {
                message: `Error in the update schedule simpro webhook: ${JSON.stringify(err)}`
            }
        }
    }
    

    static async handleDeleteScheduleInSmartsheet(webhookData: SimproWebhookType) {
        try {
            const { scheduleID, jobID, sectionID } = webhookData.reference;
            if (jobCardV2SheetId) {
                const sheetInfo = await smartsheet.sheets.getSheet({ id: jobCardV2SheetId });
                const columns = sheetInfo.columns;
                const scheduleColumn = columns.find((col: SmartsheetColumnType) => col.title === "ScheduleID");
                const scheduleIdColumnId = scheduleColumn.id;

                console.log("schedule column in v2 for delete", scheduleColumn)

                const scheduleCommentColumn = columns.find((col: SmartsheetColumnType) => col.title === "ScheduleComment");
                // console.log("schedule comment column in v2 for delete", scheduleCommentColumn)
                const scheduleCommentColumnId = scheduleCommentColumn.id;
                let scheduleDataForSmartsheet: SmartsheetSheetRowsType | undefined;
                const existingRows: SmartsheetSheetRowsType[] = sheetInfo.rows;

                console.log("Scheudle ID", scheduleID)
                for (let i = 0; i < existingRows.length; i++) {
                    let currentRow = existingRows[i];
                    const cellData = currentRow.cells.find(
                        (cell: { columnId: string; value: any }) => cell.columnId === scheduleIdColumnId
                    );
                    if (cellData?.value === scheduleID) {
                        scheduleDataForSmartsheet = currentRow;
                        break;
                    }
                }

                console.log('scheduleDataForSmartsheet for delete', scheduleDataForSmartsheet)

                if (scheduleDataForSmartsheet) {
                    const rowsToUpdate = [{
                        id: scheduleDataForSmartsheet?.id,
                        cells: [{ columnId: scheduleCommentColumnId, value: "Deleted from Simpro" }],
                    }]

                    await smartsheet.sheets.updateRow({
                        sheetId: jobCardV2SheetId,
                        body: rowsToUpdate,
                    });

                    console.log('delete comment added to the schedule in smartsheet', jobCardV2SheetId)
                } else {
                    // Logic to check the row in move past sheet
                    console.log("Schedule not found in the main sheet, checking Move Past Sheet", scheduleID)
                    let movePastSheetInfo = await smartsheet.sheets.getSheet({ id: jobCardV2MovePastSheetId });
                    const movePastSheetColumns = movePastSheetInfo.columns;
                    const movePastScheduleColumn = movePastSheetColumns.find((col: SmartsheetColumnType) => col.title === "ScheduleID");
                    const movePastScheduleIdColumnId = movePastScheduleColumn.id;
                    const movePastScheduleCommentColumn = movePastSheetColumns.find((col: SmartsheetColumnType) => col.title === "ScheduleComment");
                    const movePastScheduleCommentColumnId = movePastScheduleCommentColumn.id;
                    let movePastScheduleDataForSmartsheet: SmartsheetSheetRowsType | undefined;
                    const existingMovePastRows: SmartsheetSheetRowsType[] = movePastSheetInfo.rows;
                    for (let i = 0; i < existingMovePastRows.length; i++) {
                        let currentRow = existingMovePastRows[i];
                        const cellData = currentRow.cells.find(
                            (cell: { columnId: string; value: any }) => cell.columnId === movePastScheduleIdColumnId
                        );
                        if (cellData?.value === scheduleID) {
                            movePastScheduleDataForSmartsheet = currentRow;
                            break;
                        }
                    }
                    if (movePastScheduleDataForSmartsheet) {
                        const rowsToUpdate = [{
                            id: movePastScheduleDataForSmartsheet?.id,
                            cells: [{ columnId: movePastScheduleCommentColumnId, value: "Deleted from Simpro" }],
                        }]

                        await smartsheet.sheets.updateRow({
                            sheetId: jobCardV2MovePastSheetId,
                            body: rowsToUpdate,
                        });

                        console.log('delete comment added to the schedule in smartsheet in Move Past Sheet', jobCardV2MovePastSheetId)
                    }
                }

            }

        } catch (err) {
            console.log("Error in the delete schedule simpro webhook", err);
            throw {
                message: "Error in the delete schedule simpro webhook"
            }
        }

    }

    static async handleAddUpdateRoofingCostcenterForInvoiceSmartsheet(webhookData: SimproWebhookType) {
        try {
            const { invoiceID } = webhookData.reference;
            const { ID, date_triggered } = webhookData;
            if (invoiceID) {
                // console.log("Invoice ID: ", invoiceID);
                const url = `/invoices/${invoiceID}?columns=ID,Jobs`;
                const invoiceData = await axiosSimPRO.get(url);
                const invoice = invoiceData.data;
                // console.log("Invoice Data: ", invoice);
                let jobIDsForInvoice: number[] = invoice?.Jobs?.map((job: any) => job.ID) || [];

                for (const jobID of jobIDsForInvoice) {
                    if (ID === "invoice.updated") {
                        // implement the tolerance check here 
                        const toleranceMs = 60000; // 1 minute
                        const webhookTriggerDate = new Date(date_triggered);
                        // Fetch job logs
                        const jobLogsResponse = await axiosSimPRO.get(`/logs/jobs/?jobID=${jobID}`);
                        const jobLogs = jobLogsResponse?.data || [];
                        if (jobLogs.length > 0) {
                            // Check if any log falls within ± tolerance
                            const logsWithinTolerance = jobLogs.some((log: any) => {
                                const logDate = new Date(log.DateLogged);
                                const diff = Math.abs(webhookTriggerDate.getTime() - logDate.getTime());
                                return diff <= toleranceMs; // log within 1 min before/after webhook
                            });

                            if (!logsWithinTolerance) {
                                console.log(`Skipping job ${jobID} — no matching log within 1 minute before/after`);
                                continue; // Skip this job if no matching log found
                            }
                        }
                    }

                    console.log("Processing Job ID for invoice id: ", jobID);
                    await SmartsheetService.handleAddUpdateCostcenterRoofingToSmartSheet({
                        ID: webhookData.ID,
                        build: webhookData.build,
                        description: webhookData.description,
                        name: webhookData.name,
                        action: webhookData.action,
                        reference: {
                            companyID: webhookData.reference?.companyID || 0,
                            scheduleID: webhookData.reference?.scheduleID || 0,
                            jobID: jobID,
                            sectionID: webhookData.reference?.sectionID || 0,
                            costCenterID: webhookData.reference?.costCenterID || 0,
                            invoiceID: webhookData.reference?.invoiceID || 0,
                        },
                        date_triggered: webhookData?.date_triggered || new Date().toISOString(),
                    });
                }
            }

        }
        catch (err) {
            console.log("Error in the update roofing cost center simpro webhook", err);
            throw {
                message: "Error in the update roofing cost center simpro webhook"
            }
        }
    }

    static async handleAddUpdateCostcenterRoofingToSmartSheet(webhookData: SimproWebhookType) {
        const { jobID } = webhookData.reference;
        const { date_triggered, ID } = webhookData;

        if (ID !== "job.created") {
            const toleranceMs = 60000; // 1 minute
            const webhookTriggerDate = new Date(date_triggered);

            // Fetch job logs
            const jobLogsResponse = await axiosSimPRO.get(`/logs/jobs/?jobID=${jobID}`);
            const jobLogs = jobLogsResponse?.data || [];

            if (jobLogs.length > 0) {
                // Check if any log falls within ± tolerance
                const logsWithinTolerance = jobLogs.some((log: any) => {
                    const logDate = new Date(log.DateLogged);
                    const diff = Math.abs(webhookTriggerDate.getTime() - logDate.getTime());

                    return diff <= toleranceMs; // log within 1 min before/after webhook
                });

                if (!logsWithinTolerance) {
                    console.log(`Skipping job ${jobID} — no matching log within 1 minute before/after`);
                    return;
                }
            }
        }


        let costCenterDataFromSimpro: SimproJobCostCenterType[] = [];
        const url = `/jobCostCenters/?Job.ID=${jobID}`;
        const costCenters: SimproJobCostCenterType[] = await fetchSimproPaginatedData(url, "ID,CostCenter,Name,Job,Section,DateModified,_href");
        let fetchedChartOfAccounts = await axiosSimPRO.get('/setup/accounts/chartOfAccounts/?pageSize=250&columns=ID,Name,Number');
        let chartOfAccountsArray: SimproAccountType[] = fetchedChartOfAccounts?.data;
        let foundCostCenters = 0;
        const jobDataForCostCentre = await axiosSimPRO.get(`/jobs/${jobID}?columns=ID,Type,Site,SiteContact,DateIssued,Status,Total,Customer,Name,ProjectManager,CustomFields,Totals,Stage`);
        let fetchedJobData: SimproJobType = jobDataForCostCentre?.data;
        for (const jobCostCenter of costCenters) {
            console.dir(jobCostCenter, { depth: null })
            jobCostCenter.Job = fetchedJobData;
            try {
                const ccRecordId = jobCostCenter?.CostCenter?.ID;
                let fetchedSetupCostCenterData = await axiosSimPRO.get(`/setup/accounts/costCenters/${ccRecordId}?columns=ID,Name,IncomeAccountNo`);
                let setupCostCenterData = fetchedSetupCostCenterData.data;
                if (setupCostCenterData?.IncomeAccountNo) {
                    let incomeAccountName = chartOfAccountsArray?.find(account => account?.Number == setupCostCenterData?.IncomeAccountNo)?.Name;
                    console.log("Income Account Name : " + incomeAccountName)
                    if (incomeAccountName == "Roofing Income") {
                        // console.log("Roofing income  1", jobCostCenter?.ID, jobCostCenter?.Job?.ID);
                        try {
                            const jcUrl = jobCostCenter?._href?.substring(jobCostCenter?._href?.indexOf('jobs'), jobCostCenter?._href.length);
                            let costCenterResponse = await axiosSimPRO.get(`${jcUrl}?columns=Name,ID,Claimed,Total,Totals,Site`);
                            if (costCenterResponse) {
                                jobCostCenter.CostCenter = costCenterResponse.data;
                                const siteResponse = await axiosSimPRO.get(`/sites/${costCenterResponse.data?.Site.ID}?columns=ID,Name,Address`);
                                const siteResponseData = siteResponse.data;
                                jobCostCenter.Site = siteResponseData;
                                jobCostCenter.ccRecordId = ccRecordId;
                                foundCostCenters++;
                                costCenterDataFromSimpro.push(jobCostCenter);
                            }
                        } catch (error) {
                            console.log("Error in costCenterFetch : ", error)
                        }
                    }
                }
            } catch (err) {
                if (err instanceof AxiosError) {
                    console.log("Error in fetch Const center from setup");
                    console.log("Error details: ", err.response?.data);

                } else if (err instanceof Error) {
                    console.error("Unexpected error:", err.message);
                } else {
                    // Handle non-Error objects
                    console.error("Non-error rejection:", JSON.stringify(err));
                }
            }
        }
        await SmartsheetService.updateCostcenterRoofingToSmartSheet(jobID, costCenterDataFromSimpro);
        // console.log(`Completed processing for job ${jobID}`);
    }

    static async updateCostcenterRoofingToSmartSheet(
        jobID: number,
        costCenterDataFromSimpro: SimproJobCostCenterType[]
    ) {
        try {
            // console.log('costCenterIdToMarkDeleted', costCenterIdToMarkDeleted);

            if (!jobCardRoofingDetailSheetId) {
                throw new Error("Job Card Roofing Detail Sheet ID is undefined");
            }

            if (!wipJobArchivedSheetId) {
                throw new Error("WIP Job Archived Sheet ID is undefined");
            }

            let archivedJobSheetInfo: any;
            let archivedJobSheetColumns: any;

            const activeJobSheetInfo = await smartsheet.sheets.getSheet({ id: jobCardRoofingDetailSheetId });
            const activeJobSheetColumns = activeJobSheetInfo.columns;
            const costCenterIdColumn = activeJobSheetColumns.find((col: SmartsheetColumnType) => col.title === "Cost_Center.ID");
            const jobIdColumnInActiveJobsSheet = activeJobSheetColumns.find((col: SmartsheetColumnType) => col.title === "JobID");

            if (!costCenterIdColumn) {
                throw new Error("Cost_Center.ID column not found in the sheet");
            }

            const costCenterIdColumnId = costCenterIdColumn.id;
            const jobIdColumnIdInActiveJobsSheet = jobIdColumnInActiveJobsSheet.id;
            const existingRowInActiveJobsSheet: SmartsheetSheetRowsType[] = activeJobSheetInfo.rows;
            const existingCostCenterIdsDataInActiveJobSheet: any[] = existingRowInActiveJobsSheet
                .map((row: SmartsheetSheetRowsType) => {
                    const costCenterId = row.cells.find(cell => cell.columnId === costCenterIdColumnId)?.value;
                    return costCenterId ? { costCenterId: Number(costCenterId), rowId: row.id } : null;

                })
                .filter(Boolean);
            // console.log('existingCostCenterIdsDataInActiveJobSheet', existingCostCenterIdsDataInActiveJobSheet);

            let existingCostCenterIdsDataForJobinActiveJobSheet =
                existingRowInActiveJobsSheet?.filter((row: SmartsheetSheetRowsType) => {
                    const jobIdCell = row.cells.find(
                        (cell) => cell.columnId === jobIdColumnIdInActiveJobsSheet
                    );
                    return jobIdCell?.value != null && jobIdCell.value == jobID; // loose equality
                });


            let costCenterIdNotPresentInSimproResponse: string[] = SmartsheetService.filterTheCostCenterIdNotInSimproResponse(costCenterIdColumnId, existingCostCenterIdsDataForJobinActiveJobSheet, costCenterDataFromSimpro);
            // console.log('costCenterIdNotPresentInSimproResponse', costCenterIdNotPresentInSimproResponse);
            let costCenterIdToBeMarkedAsDeleted: string[] = [];
            if (costCenterIdNotPresentInSimproResponse.length > 0) {
                costCenterIdToBeMarkedAsDeleted = await SmartsheetService.validateCostCentersBatch(costCenterIdNotPresentInSimproResponse);
                const simproCommentColumn = activeJobSheetColumns.find((col: SmartsheetColumnType) => col.title === "SIMPROComment");
                const simproCommentColumnId = simproCommentColumn.id;
                if (costCenterIdToBeMarkedAsDeleted.length > 0) {
                    // console.log('costCenterIdToBeMarkedAsDeleted', costCenterIdToBeMarkedAsDeleted);
                    const rowsIdsToMarkDeleted = existingCostCenterIdsDataInActiveJobSheet
                        .filter(item => costCenterIdToBeMarkedAsDeleted.includes(item.costCenterId))
                        .map(item => item.rowId);

                    // console.log('rowsIdsToMarkDeleted', rowsIdsToMarkDeleted)

                    const chunks = splitIntoChunks(rowsIdsToMarkDeleted, 300);
                    // console.log('Total chunks to update for deletion:', chunks.length);

                    for (const chunk of chunks) {
                        // Prepare rows for batch update
                        const rowsToUpdate = chunk.map(rowId => ({
                            id: rowId,
                            cells: [{ columnId: simproCommentColumnId, value: "Deleted from Simpro" }],
                        }));

                        // Batch update rows
                        await smartsheet.sheets.updateRow({
                            sheetId: jobCardRoofingDetailSheetId,
                            body: rowsToUpdate,
                        });

                        // console.log('WIP mark as delete  in active sheet', chunk.length, 'rows');
                    }
                }
            }


            for (const jobCostCenter of costCenterDataFromSimpro) {
                try {
                    let costCenterRowDataForActiveJobsSheet: SmartsheetSheetRowsType | undefined;
                    let costCenterRowDataForArchivedJobsSheet: SmartsheetSheetRowsType | undefined;

                    for (const element of existingRowInActiveJobsSheet) {
                        const cellData = element.cells.find(
                            (cell: { columnId: string; value: any }) => cell.columnId === costCenterIdColumnId
                        );
                        if (cellData?.value === jobCostCenter.CostCenter.ID) {
                            costCenterRowDataForActiveJobsSheet = element;
                            break;
                        }
                    }

                    if (!costCenterRowDataForActiveJobsSheet) {
                        archivedJobSheetInfo = await smartsheet.sheets.getSheet({ id: wipJobArchivedSheetId });
                        archivedJobSheetColumns = archivedJobSheetInfo.columns;
                        const costCenterIdColumnInArchivedSheet = archivedJobSheetColumns.find((col: SmartsheetColumnType) => col.title === "Cost_Center.ID");
                        const jobIdColumnInArchivedJobsSheet = archivedJobSheetColumns.find((col: SmartsheetColumnType) => col.title === "JobID");
                        if (!costCenterIdColumnInArchivedSheet) {
                            throw new Error("Cost_Center.ID column not found in the Archived Job sheet");
                        }

                        const costCenterIdColumnIdInArchivedSheet = costCenterIdColumnInArchivedSheet.id;
                        const jobIdColumnIdInArchivedJobsSheet = jobIdColumnInArchivedJobsSheet.id;
                        const existingRowInArchivedJobsSheet: SmartsheetSheetRowsType[] = archivedJobSheetInfo.rows;

                        const existingCostCenterIdsDataInArchievedJobSheet: any[] = existingRowInArchivedJobsSheet
                            .map((row: SmartsheetSheetRowsType) => {
                                const costCenterId = row.cells.find(cell => cell.columnId === costCenterIdColumnIdInArchivedSheet)?.value;
                                return costCenterId ? { costCenterId: Number(costCenterId), rowId: row.id } : null;

                            })
                            .filter(Boolean);
                        // console.log('existingCostCenterIdsDataInArchievedJobSheet', existingCostCenterIdsDataInArchievedJobSheet)

                        let existingCostCenterIdsDataForJobinArchivedJobSheet =
                            existingRowInArchivedJobsSheet?.filter((row: SmartsheetSheetRowsType) => {
                                const jobIdCell = row.cells.find(
                                    (cell) => cell.columnId === jobIdColumnIdInArchivedJobsSheet
                                );
                                return jobIdCell?.value != null && jobIdCell.value == jobID; // loose equality
                            });


                        let costCenterIdNotPresentInSimproResponseForArchivedJobSheet: string[] = SmartsheetService.filterTheCostCenterIdNotInSimproResponse(costCenterIdColumnIdInArchivedSheet, existingCostCenterIdsDataForJobinArchivedJobSheet, costCenterDataFromSimpro);
                        // console.log('costCenterIdNotPresentInSimproResponseForArchivedJobSheet', costCenterIdNotPresentInSimproResponseForArchivedJobSheet);
                        let costCenterIdToBeMarkedAsDeletedinArchievedJobSheet: string[] = [];
                        if (costCenterIdNotPresentInSimproResponseForArchivedJobSheet.length > 0) {
                            costCenterIdToBeMarkedAsDeletedinArchievedJobSheet = await SmartsheetService.validateCostCentersBatch(costCenterIdNotPresentInSimproResponseForArchivedJobSheet);
                            const simproCommentColumnInArchivedJobSheet = archivedJobSheetColumns.find((col: SmartsheetColumnType) => col.title === "SIMPROComment");
                            const simproCommentColumnIdInArchivedSheet = simproCommentColumnInArchivedJobSheet.id;
                            if (costCenterIdToBeMarkedAsDeletedinArchievedJobSheet.length > 0) {
                                // console.log('costCenterIdToBeMarkedAsDeletedinArchievedJobSheet', costCenterIdToBeMarkedAsDeletedinArchievedJobSheet);
                                const rowsIdsToMarkDeletedInArchivedSheet = existingCostCenterIdsDataInArchievedJobSheet
                                    .filter(item => costCenterIdToBeMarkedAsDeletedinArchievedJobSheet.includes(item.costCenterId))
                                    .map(item => item.rowId);

                                // console.log('rowsIdsToMarkDeletedInArchivedSheet', rowsIdsToMarkDeletedInArchivedSheet)

                                const chunks = splitIntoChunks(rowsIdsToMarkDeletedInArchivedSheet, 300);
                                // console.log('Total chunks to update for deletion:', chunks.length);

                                for (const chunk of chunks) {
                                    // Prepare rows for batch update
                                    const rowsToUpdate = chunk.map(rowId => ({
                                        id: rowId,
                                        cells: [{ columnId: simproCommentColumnIdInArchivedSheet, value: "Deleted from Simpro" }],
                                    }));

                                    // Batch update rows
                                    await smartsheet.sheets.updateRow({
                                        sheetId: wipJobArchivedSheetId,
                                        body: rowsToUpdate,
                                    });

                                    // console.log('WIP mark as delete  in archeived sheet', chunk.length, 'rows');
                                }
                            }
                        }


                        for (const element of existingRowInArchivedJobsSheet) {
                            const costCenterCellData = element.cells.find(
                                (cell: { columnId: string; value: any }) => cell.columnId === costCenterIdColumnIdInArchivedSheet
                            );
                            if (costCenterCellData?.value === jobCostCenter.CostCenter.ID) {
                                costCenterRowDataForArchivedJobsSheet = element;
                                break;
                            }
                        }
                    }

                    if (costCenterRowDataForActiveJobsSheet) {
                        const rowIdMap = {
                            [jobCostCenter.CostCenter.ID.toString()]: costCenterRowDataForActiveJobsSheet.id?.toString() || "",
                        };

                        const convertedData = convertSimprocostCenterDataToSmartsheetFormatForUpdate(
                            [jobCostCenter],
                            activeJobSheetColumns,
                            rowIdMap,
                            'full'
                        );

                        await smartsheet.sheets.updateRow({
                            sheetId: jobCardRoofingDetailSheetId,
                            body: convertedData,
                        });

                        console.log('✅ Updated row in Smartsheet in Active Jobs Sheet (Sheet ID:', jobCardRoofingDetailSheetId, ')');
                    } else if (costCenterRowDataForArchivedJobsSheet) {
                        const rowIdMap = {
                            [jobCostCenter.CostCenter.ID.toString()]: costCenterRowDataForArchivedJobsSheet.id?.toString() || "",
                        };

                        const convertedData = convertSimprocostCenterDataToSmartsheetFormatForUpdate(
                            [jobCostCenter],
                            archivedJobSheetColumns,
                            rowIdMap,
                            'full'
                        );

                        await smartsheet.sheets.updateRow({
                            sheetId: wipJobArchivedSheetId,
                            body: convertedData,
                        });

                        console.log('✅ Updated row in Smartsheet in Archived Jobs Sheet (Sheet ID:', wipJobArchivedSheetId, ')');
                    } else {
                        const convertedDataForSmartsheet = convertSimproRoofingDataToSmartsheetFormat(
                            [jobCostCenter],
                            activeJobSheetColumns,
                            'full'
                        );

                        await smartsheet.sheets.addRows({
                            sheetId: jobCardRoofingDetailSheetId,
                            body: convertedDataForSmartsheet,
                        });

                        console.log('✅ Added row in Smartsheet (Sheet ID:', jobCardRoofingDetailSheetId, ')');
                    }
                } catch (rowError) {
                    console.error(
                        `❌ Error processing cost center ID ${jobCostCenter?.CostCenter?.ID}:`,
                        rowError
                    );
                }
            }
        } catch (error) {
            console.error("❌ Failed to update cost center roofing in Smartsheet:", error);
            throw error; // rethrow so caller can handle if needed
        }
    }

    static async validateCostCentersBatch(
        costCenterIdNotPresentInSimproResponse: string[],
        chunkSize = 50 // adjust if SimPRO allows fewer
    ): Promise<string[]> {
        if (costCenterIdNotPresentInSimproResponse.length === 0) {
            return [];
        }

        const allResponses: SimproCostCenter[] = [];

        // Split IDs into chunks
        for (let i = 0; i < costCenterIdNotPresentInSimproResponse.length; i += chunkSize) {
            const chunk = costCenterIdNotPresentInSimproResponse.slice(i, i + chunkSize);
            const idQuery = `in(${chunk.join(",")})`;
            const url = `/jobCostCenters/?ID=${idQuery}`;

            // fetchSimproPaginatedData already handles pagination
            const response = await fetchSimproPaginatedData<SimproCostCenter>(url, "ID");
            allResponses.push(...response);
        }

        // Build a set of all IDs returned by SimPRO
        const simproIds = new Set(allResponses.map((r) => r.ID.toString()));

        // Filter out IDs not returned by SimPRO
        return costCenterIdNotPresentInSimproResponse.filter(
            (id) => !simproIds.has(id.toString())
        );
    }

    static filterTheCostCenterIdNotInSimproResponse(costCenterColumnID: string, existingRowInActiveJobsSheet: SmartsheetSheetRowsType[], costCenterDataFromSimpro: SimproJobCostCenterType[]): string[] {
        let costCenterIdNotPresentInSimproResponse: string[] = [];
        if (costCenterColumnID) {
            const simproCostCenterIds = costCenterDataFromSimpro.map(cc => cc.CostCenter.ID);
            for (const element of existingRowInActiveJobsSheet) {
                const cellData = element.cells.find(
                    (cell: { columnId: string; value: any }) => cell.columnId === costCenterColumnID
                );
                if (cellData?.value && !simproCostCenterIds.includes(cellData.value)) {
                    costCenterIdNotPresentInSimproResponse.push(cellData.value);
                }

            }
        }
        return costCenterIdNotPresentInSimproResponse;
    }

    static async handleAddUpdateWorkOrderLineItemsToSmartsheet(webhookData: SimproWebhookType) {
        try {
            let isInvoiceAccountNameRoofing = false;
            const { contractorJobID } = webhookData.reference;
            const contractorJobResponse = await axiosSimPRO.get(`/contractorJobs/${contractorJobID}`);
            const contractorJobData: any = contractorJobResponse?.data;
            const href = contractorJobData._href;
            const regex = /companies\/(\d+)\/jobs\/(\d+)\/sections\/(\d+)\/costCenters\/(\d+)\/contractorJobs\/(\d+)/;
            const matches = href?.match(regex);
            if (!matches) {
                throw {
                    message: "Invalid _href format in contractor job data",
                    status: 400
                }
            }
            const jobID = parseInt(matches[2]);
            const sectionID = parseInt(matches[3]);
            const costCenterID = parseInt(matches[4]);

            const costCenterDataForSchedule = await axiosSimPRO.get(`/jobCostCenters/?ID=${costCenterID}&columns=ID,Name,Job,Section,CostCenter`);
            let setupCostCenterID = costCenterDataForSchedule.data[0]?.CostCenter?.ID;
            let fetchedSetupCostCenterData = await axiosSimPRO.get(`/setup/accounts/costCenters/${setupCostCenterID}?columns=ID,Name,IncomeAccountNo`);
            let setupCostCenterData = fetchedSetupCostCenterData.data;

            let fetchedChartOfAccounts = await axiosSimPRO.get('/setup/accounts/chartOfAccounts/?pageSize=250&columns=ID,Name,Number');
            let chartOfAccountsArray: SimproAccountType[] = fetchedChartOfAccounts?.data;

            if (setupCostCenterData?.IncomeAccountNo) {
                let incomeAccountName = chartOfAccountsArray?.find(account => account?.Number == setupCostCenterData?.IncomeAccountNo)?.Name;
                if (incomeAccountName == "Roofing Income") {
                    isInvoiceAccountNameRoofing = true;
                }
            }

            if (isInvoiceAccountNameRoofing) {
                let contractorWorkOrderResponse = await axiosSimPRO(`/jobs/${jobID}/sections/${sectionID}/costCenters/${costCenterID}/contractorJobs/?columns=ID,Items,Status,DateIssued,Total`);
                let contractorWorkOrderData: SimproContractorJobType[] = contractorWorkOrderResponse?.data;
                let costCenterResponse = await axiosSimPRO.get(`jobs/${jobID}/sections/${sectionID}/costCenters/${costCenterID}?columns=Name,ID,Claimed,Total,Totals,Site`);
                if (!costCenterResponse) {
                    throw new Error("Cost center data not found");
                }
                const costCenterData: SimproCostCenterType = costCenterResponse.data;

                // Gather all current Simpro LineItemIDs
                let allSimproLineItemIDs: string[] = [];
                let allConvertedContractorJobDataArray: SimproContractorWorkOrderType[] = [];
                for (let index = 0; index < contractorWorkOrderData.length; index++) {
                    let contractorWorkOrderDataItem = contractorWorkOrderData[index];
                    let convertedContractorJobDataArray: SimproContractorWorkOrderType[] = extractLineItemsDataFromContractorJob({
                        jobID,
                        contractorJob: contractorWorkOrderDataItem,
                        costCenterData,
                        contractorName: contractorJobData?.Contractor?.Name || ''
                    });
                    allConvertedContractorJobDataArray.push(...convertedContractorJobDataArray);
                    allSimproLineItemIDs.push(
                        ...convertedContractorJobDataArray
                            .map(item => item.LineItemID)
                            .filter((id): id is string | number => id !== undefined && id !== null)
                            .map(id => id.toString())
                    );
                }

                // --- DELETE COMMENT LOGIC FOR ACTIVE SHEET ---
                if (workOrderLineItemsActiveSheetId) {
                    const activeWorkOrdersheetInfo = await smartsheet.sheets.getSheet({ id: workOrderLineItemsActiveSheetId });
                    const columnsForActiveWorkOrderSheet = activeWorkOrdersheetInfo.columns;
                    const activeOrdercolumnForLineItemID = columnsForActiveWorkOrderSheet.find((col: SmartsheetColumnType) => col.title === "LineItemID");
                    const workOrderIdColumn = columnsForActiveWorkOrderSheet.find((col: SmartsheetColumnType) => col.title === "WorkOrderID");
                    const simproCommentColumn = columnsForActiveWorkOrderSheet.find((col: SmartsheetColumnType) => col.title === "SIMPROComment");
                    if (!activeOrdercolumnForLineItemID || !workOrderIdColumn || !simproCommentColumn) {
                        throw {
                            message: "LineItemID, WorkOrderID or SIMPROComment column not found in the sheet",
                            status: 400
                        }
                    }
                    const activeSheetlineItemIdColumnId = activeOrdercolumnForLineItemID.id;
                    const workOrderIdColumnId = workOrderIdColumn.id;
                    const simproCommentColumnId = simproCommentColumn.id;
                    const existingRowsInActiveWoSheet: SmartsheetSheetRowsType[] = activeWorkOrdersheetInfo.rows;

                    // Only consider rows for this WorkOrderID
                    let rowsToMarkDeleted: string[] = [];
                    for (const row of existingRowsInActiveWoSheet) {
                        const lineItemCell = row.cells.find(cell => cell.columnId === activeSheetlineItemIdColumnId);
                        const workOrderCell = row.cells.find(cell => cell.columnId === workOrderIdColumnId);
                        if (
                            workOrderCell?.value == contractorJobID &&
                            lineItemCell?.value &&
                            !allSimproLineItemIDs.includes(lineItemCell.value)
                        ) {
                            if (row.id !== undefined && row.id !== null) {
                                rowsToMarkDeleted.push(row.id.toString());
                            }
                        }
                    }
                    if (rowsToMarkDeleted.length > 0) {
                        const chunks = splitIntoChunks(rowsToMarkDeleted, 300);
                        for (const chunk of chunks) {
                            const rowsToUpdate = chunk.map(rowId => ({
                                id: rowId,
                                cells: [{ columnId: simproCommentColumnId, value: "Deleted from Simpro" }],
                            }));
                            await smartsheet.sheets.updateRow({
                                sheetId: workOrderLineItemsActiveSheetId,
                                body: rowsToUpdate,
                            });
                            console.log('Marked as deleted in active sheet for work order line items:', chunk.length, 'rows');
                        }
                    }
                }

                // --- EXISTING ADD/UPDATE LOGIC ---
                for (let i = 0; i < allConvertedContractorJobDataArray.length; i++) {
                    const currentLineItem = allConvertedContractorJobDataArray[i];
                    if (
                        workOrderLineItemsActiveSheetId &&
                        currentLineItem?.LineItemID !== undefined
                    ) {
                        const activeWorkOrdersheetInfo = await smartsheet.sheets.getSheet({ id: workOrderLineItemsActiveSheetId });
                        const columnsForActiveWorkOrderSheet = activeWorkOrdersheetInfo.columns;
                        const activeOrdercolumnForLineItemID = columnsForActiveWorkOrderSheet.find((col: SmartsheetColumnType) => col.title === "LineItemID");
                        if (!activeOrdercolumnForLineItemID) {
                            throw {
                                message: "LineItemID column not found in the sheet",
                                status: 400
                            }
                        }
                        const activeSheetlineItemIdColumnId = activeOrdercolumnForLineItemID.id;
                        const existingRowsInActiveWoSheet: SmartsheetSheetRowsType[] = activeWorkOrdersheetInfo.rows;
                        let activeWorkOrderItemDataForSmartsheet: SmartsheetSheetRowsType | undefined;
                        for (let i = 0; i < existingRowsInActiveWoSheet.length; i++) {
                            let currentRow = existingRowsInActiveWoSheet[i];
                            const cellData = currentRow.cells.find(
                                (cell: { columnId: string; value: any }) => cell.columnId === activeSheetlineItemIdColumnId
                            );
                            if (cellData?.value === currentLineItem.LineItemID) {
                                activeWorkOrderItemDataForSmartsheet = currentRow;
                                break;
                            }
                        }
                        if (activeWorkOrderItemDataForSmartsheet) {
                            const rowIdMap: { [key: string]: string } = {
                                [currentLineItem.LineItemID.toString()]:
                                    activeWorkOrderItemDataForSmartsheet.id?.toString() || "",
                            };
                            const convertedDataForActiveSheet = convertSimproContractorJobDataToSmartsheetFormatForUpdate([currentLineItem], columnsForActiveWorkOrderSheet, rowIdMap);
                            await smartsheet.sheets.updateRow({
                                sheetId: workOrderLineItemsActiveSheetId,
                                body: convertedDataForActiveSheet,
                            });
                            // console.log('Updated row in smartsheet in sheet for contractor job work order line item ', workOrderLineItemsActiveSheetId)
                        } else {
                            const archivedWorkOrdersheetInfo = await smartsheet.sheets.getSheet({ id: workOrderLineItemsArchivedSheetId });
                            const columnsForArchivedWorkOrderSheet = archivedWorkOrdersheetInfo.columns;
                            const archivedOrdercolumnForLineItemID = columnsForArchivedWorkOrderSheet.find((col: SmartsheetColumnType) => col.title === "LineItemID");
                            if (!archivedOrdercolumnForLineItemID) {
                                throw {
                                    message: "LineItemID column not found in the archived work order sheet",
                                    status: 400
                                }
                            }
                            const archivedSheetlineItemIdColumnId = archivedOrdercolumnForLineItemID.id;
                            const existingRowsInArchivedWoSheet: SmartsheetSheetRowsType[] = archivedWorkOrdersheetInfo.rows;
                            let archivedWorkOrderItemDataForSmartsheet: SmartsheetSheetRowsType | undefined;
                            for (let i = 0; i < existingRowsInArchivedWoSheet.length; i++) {
                                let currentRow = existingRowsInArchivedWoSheet[i];
                                const cellData = currentRow.cells.find(
                                    (cell: { columnId: string; value: any }) => cell.columnId === archivedSheetlineItemIdColumnId
                                );
                                if (cellData?.value === currentLineItem.LineItemID) {
                                    archivedWorkOrderItemDataForSmartsheet = currentRow;
                                    break;
                                }
                            }
                            if (archivedWorkOrderItemDataForSmartsheet) {
                                const rowIdMap: { [key: string]: string } = {
                                    [currentLineItem.LineItemID.toString()]:
                                        archivedWorkOrderItemDataForSmartsheet.id?.toString() || "",
                                };
                                const convertedDataForArchivedSheet = convertSimproContractorJobDataToSmartsheetFormatForUpdate([currentLineItem], columnsForArchivedWorkOrderSheet, rowIdMap);
                                await smartsheet.sheets.updateRow({
                                    sheetId: workOrderLineItemsArchivedSheetId,
                                    body: convertedDataForArchivedSheet,
                                });
                                // console.log('Updated row in smartsheet in archived sheet for contractor job work order line item ', workOrderLineItemsArchivedSheetId)
                            } else {
                                const convertedDataForSmartsheet = convertSimproContractorDataToSmartsheetFormat([currentLineItem], columnsForActiveWorkOrderSheet);
                                await smartsheet.sheets.addRows({
                                    sheetId: workOrderLineItemsActiveSheetId,
                                    body: convertedDataForSmartsheet,
                                });
                                // console.log('Added row in smartsheet in sheeet for contractor job work order line item ', workOrderLineItemsActiveSheetId)
                            }
                        }
                    }
                }
            }
        } catch (err) {
            console.log("Error in the update schedule simpro webhook", err);
            throw {
                message: "Error in the update schedule simpro webhook"
            }
        }
    }

    static async fetchCostCenterDataForGivenCostCenterIds(
        costCenterIdsDataFetchFromSimpro: CostCenterJobInfo[],
    ): Promise<SimproJobCostCenterTypeForAmountUpdate[]> {  // Return type should be an array

        if (costCenterIdsDataFetchFromSimpro.length === 0) {
            return [];
        }

        const allResponses: SimproJobCostCenterTypeForAmountUpdate[] = [];

        for (let i = 0; i < costCenterIdsDataFetchFromSimpro.length; i++) {
            try {
                const costCenterData = costCenterIdsDataFetchFromSimpro[i];
                const costCenterId = costCenterData.costCenterId;
                const sectionId = costCenterData?.sectionId;
                const jobId = costCenterData?.cellDataJobId;

                const jcUrl = `/jobs/${jobId}/sections/${sectionId}/costCenters/${costCenterId}`;

                const costCenterResponse = await axiosSimPRO.get(
                    `${jcUrl}?columns=Name,ID,Claimed,Total,Totals,Site`
                );
                const costCenterResponseData = costCenterResponse?.data;

                const siteResponse = await axiosSimPRO.get(
                    `/sites/${costCenterResponseData?.Site.ID}?columns=ID,Name,Address`
                );
                const siteResponseData = siteResponse.data;

                const sectionResponse = await axiosSimPRO.get(`/jobs/${jobId}/sections/${sectionId}`);
                const sectionResponseData = sectionResponse.data;

                const jobStageResponse = await axiosSimPRO.get(`/jobs/${jobId}?columns=ID,Stage`);
                const jobStageResponseData = jobStageResponse.data;

                const jobCostCenterData: SimproJobCostCenterTypeForAmountUpdate = {
                    CostCenter: costCenterResponseData,
                    Site: siteResponseData,
                    Section: sectionResponseData,
                    JobStage: jobStageResponseData?.Stage || null,
                };

                allResponses.push(jobCostCenterData);

                if ((i + 1) % 50 === 0) {
                    console.log(`Fetched ${i + 1} cost centers out of ${costCenterIdsDataFetchFromSimpro.length}`);
                }


            } catch (err) {
                console.log("Error fetching the data", err)
            }

        }

        return allResponses;
    }

    /**
     * Checks schedule deletion status in SimPro and updates the IsDeleted column in Smartsheet
     * This method iterates through specified sheets, validates each schedule's existence in SimPro,
     * and marks deleted schedules in Smartsheet
     * 
     * @param scheduleIdToCheck - Optional: Specific schedule ID to check. If not provided, checks all schedules in both sheets
     * @param sheetIds - Optional: Specific sheet IDs to check. If not provided, uses environment sheet IDs
     *                   Can include sheetsToProcess array to specify which sheets to process: ['active'] | ['archived'] | ['active', 'archived']
     * @returns Object with operation results and summary
     */
    static async checkAndUpdateScheduleDeletionStatus(
        scheduleIdToCheck?: number | string,
        sheetIds?: { activeSheetId?: string; archivedSheetId?: string; sheetsToProcess?: string[] }
    ) {
        try {
            console.log("🔍 Schedule Deletion Check: Starting validation process...");
            
            // Determine which sheets to process
            const sheetsToProcess = sheetIds?.sheetsToProcess || ['active', 'archived']; // Default: both sheets
            
            // Use provided sheet IDs or fall back to environment variables for roofing schedules
            const activeSheetId = sheetIds?.activeSheetId || roofingSchedulesActiveFromDbSheetId;
            const archivedSheetId = sheetIds?.archivedSheetId || roofingSchedulesArchivedFromDbSheetId;

            // Validate that we have the sheet IDs we need to process
            if (sheetsToProcess.includes('active') && !activeSheetId) {
                throw new Error(
                    "Active sheet ID not configured. Please set ROOFING_SCHEDULES_ACTIVE_FROM_DB_SHEET_ID environment variable or pass activeSheetId parameter"
                );
            }
            
            if (sheetsToProcess.includes('archived') && !archivedSheetId) {
                throw new Error(
                    "Archived sheet ID not configured. Please set ROOFING_SCHEDULES_ARCHIVED_FROM_DB_SHEET_ID environment variable or pass archivedSheetId parameter"
                );
            }

            console.log(`📋 Using sheets: Active=${activeSheetId}, Archived=${archivedSheetId}`);

            const results: {
                totalSchedulesChecked: number;
                deletedSchedulesFound: number;
                activeSchedulesFound: number;
                erroredSchedules: Array<{ rowId: number | string; error: string }>;
                updatedRows: Array<number | string>;
            } = {
                totalSchedulesChecked: 0,
                deletedSchedulesFound: 0,
                activeSchedulesFound: 0,
                erroredSchedules: [],
                updatedRows: [],
            };

            // Process sheets based on sheetsToProcess array
            const sheetsToCheck = [
                { id: activeSheetId, name: "Active Schedule Sheet", type: 'active' },
                { id: archivedSheetId, name: "Archived Schedule Sheet", type: 'archived' },
            ].filter(sheet => sheetsToProcess.includes(sheet.type));

            if (sheetsToCheck.length === 0) {
                throw new Error("No sheets to process. Invalid sheetsToProcess configuration.");
            }

            console.log(`📋 Processing ${sheetsToCheck.length} sheet(s): ${sheetsToCheck.map(s => s.type).join(', ')}`);

            for (const sheetConfig of sheetsToCheck) {
                try {
                    console.log(`\n📄 Processing ${sheetConfig.name} (ID: ${sheetConfig.id})...`);
                    
                    const sheetInfo = await smartsheet.sheets.getSheet({ id: sheetConfig.id });
                    const columns = sheetInfo.columns;
                    const rows = sheetInfo.rows;

                    // Find column IDs for required fields
                    const scheduleIdColumn = columns.find((col: SmartsheetColumnType) => col.title === "ID-Schedule");
                    const jobIdColumn = columns.find((col: SmartsheetColumnType) => col.title === "ID-Job");
                    const sectionIdColumn = columns.find((col: SmartsheetColumnType) => col.title === "ID-Section");
                    const costCenterIdColumn = columns.find((col: SmartsheetColumnType) => col.title === "ID-CostCentre");
                    const isDeletedColumn = columns.find((col: SmartsheetColumnType) => col.title === "ISDeleted");

                    if (!scheduleIdColumn || !jobIdColumn || !sectionIdColumn || !costCenterIdColumn || !isDeletedColumn) {
                        console.warn(
                            `⚠️ ${sheetConfig.name}: Missing required columns. ` +
                            `Found: ID-Schedule=${!!scheduleIdColumn}, ID-Job=${!!jobIdColumn}, ` +
                            `ID-Section=${!!sectionIdColumn}, ID-CostCentre=${!!costCenterIdColumn}, ISDeleted=${!!isDeletedColumn}`
                        );
                        continue;
                    }

                    const scheduleIdColId = scheduleIdColumn.id;
                    const jobIdColId = jobIdColumn.id;
                    const sectionIdColId = sectionIdColumn.id;
                    const costCenterIdColId = costCenterIdColumn.id;
                    const isDeletedColId = isDeletedColumn.id;

                    // Filter rows if specific schedule ID is provided
                    let rowsToProcess = rows;
                    if (scheduleIdToCheck) {
                        rowsToProcess = rows.filter((row: SmartsheetSheetRowsType) => {
                            const scheduleCell = row.cells.find((cell) => cell.columnId === scheduleIdColId);
                            return scheduleCell?.value === scheduleIdToCheck || scheduleCell?.value === scheduleIdToCheck.toString();
                        });
                        console.log(`🎯 Filtered to ${rowsToProcess.length} row(s) for schedule ID: ${scheduleIdToCheck}`);
                    }

                    // TODO: TESTING - Remove this after testing. Limits to first 20 rows only
                    // const TEST_MODE = true; // Set to false to process all rows
                    // if (TEST_MODE) {
                    //     rowsToProcess = rowsToProcess.slice(0, 20);
                    //     console.log(`🧪 TEST MODE ENABLED: Limited to first 20 rows. Processing ${rowsToProcess.length} rows.`);
                    // }

                    // Process each row
                    const rowsToUpdate: any[] = [];
                    for (const row of rowsToProcess) {
                        try {
                            const scheduleCell = row.cells.find((cell: any) => cell.columnId === scheduleIdColId);
                            const jobCell = row.cells.find((cell: any) => cell.columnId === jobIdColId);
                            const sectionCell = row.cells.find((cell: any) => cell.columnId === sectionIdColId);
                            const costCenterCell = row.cells.find((cell: any) => cell.columnId === costCenterIdColId);

                            const scheduleId = scheduleCell?.value;
                            const jobId = jobCell?.value;
                            const sectionId = sectionCell?.value;
                            const costCenterId = costCenterCell?.value;

                            // Skip if any required field is missing
                            if (!scheduleId || !jobId || !sectionId || !costCenterId) {
                                console.log(
                                    `⏭️ Row ${row.id}: Skipping - missing required fields ` +
                                    `(S:${scheduleId}, J:${jobId}, Sec:${sectionId}, CC:${costCenterId})`
                                );
                                continue;
                            }

                            results.totalSchedulesChecked++;

                            // Validate schedule existence in SimPro
                            const validationResult = await validateScheduleExistence(
                                jobId,
                                sectionId,
                                costCenterId,
                                scheduleId
                            );

                            if (!validationResult) {
                                throw new Error(`Validation result is undefined for schedule ${scheduleId}`);
                            }

                            const isDeletedValue = validationResult.exists ? "No" : "Yes";
                            if (!validationResult.exists) {
                                results.deletedSchedulesFound++;
                                console.log(
                                    `❌ Schedule ${scheduleId} marked as DELETED (Job:${jobId}, Section:${sectionId}, CC:${costCenterId})`
                                );
                            } else {
                                results.activeSchedulesFound++;
                                console.log(
                                    `✅ Schedule ${scheduleId} is ACTIVE (Job:${jobId}, Section:${sectionId}, CC:${costCenterId})`
                                );
                            }

                            // Queue row for update
                            rowsToUpdate.push({
                                id: row.id,
                                cells: [{ columnId: isDeletedColId, value: isDeletedValue }],
                            });
                        } catch (err: any) {
                            console.error(`❌ Error processing row ${row.id}:`, err);
                            results.erroredSchedules.push({
                                rowId: row.id,
                                error: err instanceof Error ? err.message : JSON.stringify(err),
                            });
                        }
                    }

                    // Update Smartsheet in batches
                    if (rowsToUpdate.length > 0) {
                        console.log(`📝 Updating ${rowsToUpdate.length} rows in ${sheetConfig.name}...`);
                        const chunks = splitIntoChunks(rowsToUpdate, 300); // Smartsheet API batch limit

                        for (let i = 0; i < chunks.length; i++) {
                            try {
                                await smartsheet.sheets.updateRow({
                                    sheetId: sheetConfig.id,
                                    body: chunks[i],
                                });
                                console.log(
                                    `✅ Updated batch ${i + 1}/${chunks.length} (${chunks[i].length} rows) in ${sheetConfig.name}`
                                );
                                results.updatedRows.push(...chunks[i].map((r) => r.id));
                            } catch (updateErr) {
                                console.error(`❌ Error updating batch ${i + 1} in ${sheetConfig.name}:`, updateErr);
                                throw updateErr;
                            }
                        }
                    } else {
                        console.log(`ℹ️ No rows to update in ${sheetConfig.name}`);
                    }
                } catch (sheetErr) {
                    console.error(`❌ Error processing ${sheetConfig.name}:`, sheetErr);
                    throw sheetErr;
                }
            }

            console.log("\n📊 Summary:");
            console.log(`  Total Schedules Checked: ${results.totalSchedulesChecked}`);
            console.log(`  Deleted Schedules Found: ${results.deletedSchedulesFound}`);
            console.log(`  Active Schedules Found: ${results.activeSchedulesFound}`);
            console.log(`  Rows Updated: ${results.updatedRows.length}`);
            console.log(`  Errors Encountered: ${results.erroredSchedules.length}`);

            return {
                status: "success",
                message: "Schedule deletion check completed",
                ...results,
            };
        } catch (err) {
            console.error("❌ Error in checkAndUpdateScheduleDeletionStatus:", err);
            throw {
                status: "error",
                message: "Error checking schedule deletion status",
                error: err instanceof Error ? err.message : JSON.stringify(err),
            };
        }
    }


}

