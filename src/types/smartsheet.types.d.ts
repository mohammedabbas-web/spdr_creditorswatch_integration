import { Document } from 'mongoose';

export type SmartsheetRowCellType = {
    columnId: number;
    value: string;
};

export type SmartsheetSheetRowsType = {
    cells: SmartsheetCellType[];
    id?: number;
};

export type SmartsheetColumnType = {
    id: number;
    version?: number;
    index?: number;
    title: string;
    type?: string;
    options?: string[];
    locked?: boolean;
    lockedForUser?: boolean;
    validation?: boolean;
    width?: number;
};


// Define the TimeInterval type
export type TimeInterval = {
    start?: Date;
    stop?: Date;
};


export type TaskHourRecord = {
    taskId: string;
    timeIntervals: TimeInterval[];
    createdAt: Date;
    __v?: number;
};

export type TaskData = {
    summary: string;
    totalTimeSpent: string;
};

// Define the ITaskHourRecord type extending Document (to work with Mongoose models)
export type ITaskHourRecord = Document & {
    taskId: string;
    timeIntervals: TimeInterval[];
    createdAt: Date;
    addStartTime(startTime: Date): Promise<ITaskHourRecord>;
    addStopTime(stopTime: Date): Promise<ITaskHourRecord | void>;
    _id?: string;
};

export type SimproScheduleRowObjectType = {
    ScheduleID: string | number | undefined;
    ScheduleType: string | undefined;
    StaffName: string | undefined;
    ScheduleDate: string | Date | undefined;
    StartTime: string | undefined;
    EndTime: string | undefined;
    CustomerName?: string | undefined;
    JobID?: string | number | undefined;
    SiteID?: string | number | undefined;
    SiteName?: string | undefined;
    SiteContact?: string | undefined;
    CostCenterName?: string | undefined;
    Cost_center_ID?: string | number | undefined;
    CustomerPhone?: string | undefined;
    CustomerEmail?: string | undefined;
    JobName?: string | undefined;
    ProjectManager?: string | undefined;
    Zone?: string | number | null | undefined;
    JobTrade?: string | number | null | undefined;
    "Roof Double or Single"?: string | number | null | undefined;
    ScheduleNotes?: string | null | undefined;
    "Percentage Client Invoice Claimed (From Simpro)"?: number | null | undefined;
    Suburb?: string | null | undefined;
    "Job level invoiced Percent"?: number | string | null | undefined;
    "Costcenter level invoiced Percent"?: number | string | null | undefined;
    "CC Yet to invoice"?: number | string | null | undefined;
}


export type SimproQuotationRowObjectType = {
    QuoteID: string | number | undefined;
    QuoteName: string | undefined;
    Status: string | undefined;
    Customer: string | undefined;
    Site: string | undefined;
    "SellPrice(IncTax)": string | undefined;
    SalesPerson: string | number | undefined;
    QuoteStage: string | number | undefined;
    Tags: string | undefined;
    CreatedDate: string | Date | undefined;
    DueDate: string | Date | undefined;
    QuoteChasedBy: string | undefined | null;
    NewQuoteOrVariation: string | undefined | null;
    Priority: string | null | undefined;
    Lead: string | null | undefined;
}

export type SimproLeadRowObjectType = {
    LeadID: string | number | undefined;
    LeadName: string | undefined;
    Status: string | undefined;
    CreatedDate: string | undefined;
    FollowUpDate: string | undefined;
    LatestSchedule: string | number | undefined;
    SalesPerson: string | number | undefined;
    CustomerName: string | undefined;
    SiteName: string | undefined;
    DueDateBy: string | Date | undefined;
    Priority: string | null | undefined;
    Lead: string | null | undefined;
    Tags: string | null | undefined;
}

declare module 'smartsheet' {
    const smartsheet: any;
    export = smartsheet;
}

export type ExistingScheduleType = {
    scheduleId: number | string;
    rowId: number | string;
}

export type ExistingCostCenterType = {
    costCenterId: number | string;
    jobId: number | string;
    jobSectionId: number | string;
    rowId: number | string;
}

export type ExistingQuotationType = {
    quoteId: number | string;
    rowId: number | string;
}

export type ExistingLeadsType = {
    leadId: number | string;
    rowId: number | string;
}

export type SimproJobRoofingDetailType = {
    JobID: string | number | undefined;
    Customer: string | null | undefined;
    "Job.SiteName": string | null | undefined;
    "SectionName": string | null | undefined;
    "Site.Suburb": string | null | undefined;
    "Job.Name": string | null | undefined;
    "Job.Stage": string | null | undefined;
    "Job_Section.ID": string | number | undefined;
    "Cost_Center.ID": string | number | undefined;
    "Cost_Center.Name": string | null | undefined;
    "Remainingamount_Ex.Tax": string | number | undefined;
    "CostCentre_Total_Ex.Tax": string | number | undefined;
    "Remaining_Invoice_Percentage": string | number | undefined;
    "CC_Material_Cost": string | number | undefined;
    "CC_Resource_Cost": string | number | undefined;
}

export type ExistingIncomeType = {
    JobID: number;
    CostCenterID: number;
    SectionID: number;
}

export type SimproContractorWorkOrderType ={
    JobID: string | number | undefined;
    SiteName: string | null | undefined;
    CostCenterID: string | number | undefined;
    CostCenterName: string | null | undefined;
    WorkOrderID: string | number | undefined;
    ContractorName: string | null | undefined;
    ContractorJobStatus: string | null | undefined;
    ContractorJobTotal: string | number | undefined;
    DateIssued: string | null | undefined;
    LineItemType: string | null | undefined;
    LineItemID: string | number | undefined;
    LineItemName : string | null | undefined;
    LineItemQty: string | number | undefined;
    LineItemAmount: string | number | undefined;
}

/**
 * Type for schedule deletion check operation
 * Represents the result of validating a schedule's existence in SimPro
 */
export type ScheduleValidationResultType = {
    exists: boolean;
    scheduleID: number | string;
    staffName: string | null; // Staff name from SimPro if schedule exists
    message: string;
    data?: any;
    errorStatus?: number;
};

/**
 * Type for schedule deletion check service response
 * Contains summary of the deletion check operation
 */
export type ScheduleDeletionCheckResultType = {
    status: string;
    message: string;
    totalSchedulesChecked: number;
    deletedSchedulesFound: number;
    activeSchedulesFound: number;
    updatedRows: (number | string)[];
    erroredSchedules: ScheduleCheckErrorType[];
};

/**
 * Type for errors encountered during schedule deletion check
 */
export type ScheduleCheckErrorType = {
    rowId: number | string;
    error: string;
};

/**
 * Type for schedule row data used in deletion check
 * Combines the key columns needed to validate schedule existence
 */
export type ScheduleCheckRowType = {
    scheduleId: number | string;
    jobId: number | string;
    sectionId: number | string;
    costCenterId: number | string;
    rowId: number | string;
    isDeleted?: string;
};

/**
 * Type for sheet ID override configuration
 * Allows specifying custom sheet IDs instead of using environment variables
 */
export type SheetIdOverrideType = {
    activeSheetId: string;
    archivedSheetId: string;
};