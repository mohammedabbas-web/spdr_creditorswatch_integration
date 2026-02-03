import { AxiosError } from "axios";
import axiosSimPRO from "../../config/axiosSimProConfig";
import { SimproAccountType, SimproCustomerType, SimproJobCostCenterType, SimproJobType, SimproScheduleType } from "../../types/simpro.types";
import { fetchSimproPaginatedData } from "./simproPaginationService";
import moment from "moment";

export const fetchScheduleData = async () => {
    try {
        let fetchedChartOfAccounts = await axiosSimPRO.get('/setup/accounts/chartOfAccounts/?pageSize=250&columns=ID,Name,Number');
        let chartOfAccountsArray: SimproAccountType[] = fetchedChartOfAccounts?.data;
        const allCustomerData: SimproCustomerType[] = await fetchSimproPaginatedData('/customers/', "");
        const simproCustomerMap: { [key: string]: SimproCustomerType } = {};

        if (allCustomerData.length > 0) {
            // Fetching customer data sequentially
            for (const customer of allCustomerData) {
                const hrefForCustomerData = customer?._href;
                if (!hrefForCustomerData) continue;

                let customerUrl = hrefForCustomerData.split('customers')[1];
                if (!customerUrl) continue;

                try {
                    let customerDataForJob: SimproCustomerType;
                    if (customerUrl.includes('companies')) {
                        const customerResponse = await axiosSimPRO.get(`/customers${customerUrl}?columns=ID,CompanyName,Phone,Address,Email`);
                        customerDataForJob = { ...customerResponse.data, Type: "Company" };
                    } else {
                        const customerResponse = await axiosSimPRO.get(`/customers${customerUrl}?columns=ID,GivenName,FamilyName,Phone,Address,Email`);
                        customerDataForJob = { ...customerResponse.data, Type: "Individual" };
                    }
                    if (customerDataForJob?.ID) {
                        simproCustomerMap[customerDataForJob.ID.toString()] = customerDataForJob;
                    }
                } catch (error) {
                    console.error("Error fetching customer data for:", customer.ID, error);
                }
            }
        }

        if (Object.keys(simproCustomerMap).length) {
            const currentDate = moment().subtract(7, 'day').format("YYYY-MM-DD");
            const url = `/schedules/?Type=job&Date=gt(${currentDate})&pageSize=100`;
            let fetchedSimproSchedulesData: SimproScheduleType[] = await fetchSimproPaginatedData(url, "ID,Reference,Staff,Date,Blocks,Notes");
            // console.log('Total schedules fetched: ', fetchedSimproSchedulesData.length)
            let jobIdsToAddArray: number[] = []
            // Fetch job information sequentially
            for (let i = 0; i < fetchedSimproSchedulesData.length; i++) {
                const schedule = fetchedSimproSchedulesData[i];
                console.log('Schedule',schedule)
                let jobIdForSchedule = schedule?.Reference?.split('-')[0];
                let costCenterIdForSchedule = schedule?.Reference?.split('-')[1];
                if (jobIdForSchedule) {
                    try {
                        // console.log("Fetching job for schdule " + jobIdForSchedule + ' at index', i, " of ", fetchedSimproSchedulesData.length)
                        const jobDataForSchedule = await axiosSimPRO.get(`/jobs/${jobIdForSchedule}?columns=ID,Type,Site,SiteContact,DateIssued,Status,Total,Customer,Name,ProjectManager,CustomFields,Totals`);
                        let fetchedJobData: SimproJobType = jobDataForSchedule?.data;
                        let siteId = fetchedJobData?.Site?.ID;
                        if (siteId) {
                            const siteResponse = await axiosSimPRO.get(`/sites/${siteId}?columns=ID,Name,Address`);
                            let siteResponseData = siteResponse.data;
                            fetchedJobData.Site = siteResponseData;
                        }
      
                        schedule.Job = fetchedJobData;
                        if (schedule?.Job?.Customer) {
                            const customerId = schedule.Job.Customer.ID?.toString();
                            if (customerId && simproCustomerMap[customerId]) {
                                schedule.Job.Customer = simproCustomerMap[customerId];
                            }
                        }
                    } catch (error) {
                        console.error(`Error fetching job data for schedule ID: ${jobIdForSchedule}`, error);
                    }
                }

                console.log('costCenterId for Shcduel   ',costCenterIdForSchedule)
                if (costCenterIdForSchedule) {
                    try {
                        const costCenterDataForSchedule = await axiosSimPRO.get(`/jobCostCenters/?ID=${costCenterIdForSchedule}&columns=ID,Name,Job,Section,CostCenter`);
                        let sectionIdForSchedule =
                            Array.isArray(costCenterDataForSchedule?.data)
                                ? costCenterDataForSchedule.data[0]?.Section?.ID
                                : null;

                        let jobIdForScheduleFetched =
                            Array.isArray(costCenterDataForSchedule?.data)
                                ? costCenterDataForSchedule.data[0]?.Job?.ID
                                : null;

                        let setupCostCenterID = costCenterDataForSchedule.data[0]?.CostCenter?.ID;
                        let fetchedSetupCostCenterData = await axiosSimPRO.get(`/setup/accounts/costCenters/${setupCostCenterID}?columns=ID,Name,IncomeAccountNo`);
                        let setupCostCenterData = fetchedSetupCostCenterData.data;
                        if (setupCostCenterData?.IncomeAccountNo) {
                            let incomeAccountName = chartOfAccountsArray?.find(account => account?.Number == setupCostCenterData?.IncomeAccountNo)?.Name;
                            if (incomeAccountName == "Roofing Income") {
                                // console.log('CostCenterId For Roofing Income 3', costCenterIdForSchedule, jobIdForScheduleFetched)
                                jobIdsToAddArray.push(jobIdForScheduleFetched)
                            }
                        }

                        try {
                            let costCenterResponse = await axiosSimPRO.get(`jobs/${jobIdForScheduleFetched}/sections/${sectionIdForSchedule}/costCenters/${costCenterIdForSchedule}?columns=Name,ID,Claimed,Total,Totals`);
                            if (costCenterResponse) {
                                schedule.CostCenter = costCenterResponse.data;
                            }
                        } catch (error) {
                            console.log("Error in costCenterFetch : ", error)
                        }
                    } catch (error) {
                        console.error(`Error fetching cost center data for schedule ID: ${costCenterIdForSchedule}`, error);
                    }
                }


            }

            // console.log('jobsToFilter Length: ', jobIdsToAddArray.length)
            console.log('Current schedule Length: ', fetchedSimproSchedulesData.length)
            fetchedSimproSchedulesData = fetchedSimproSchedulesData.filter(schedule =>
                jobIdsToAddArray.includes(schedule?.Job?.ID ?? -1)
            );
            console.log('Filtered schedule Length: ', fetchedSimproSchedulesData.length)
            return fetchedSimproSchedulesData;
        } else {
            return [];
        }
    } catch (err) {
        if (err instanceof AxiosError) {
            console.log("Error in fetchScheduleData as AxiosError");
            console.log("Error details: ", err.response?.data);
            throw { message: "Something went wrong while fetching schedule data : " + JSON.stringify(err.response) }
        } else {
            console.log("Error in fetchScheduleData as other error");
            console.log("Error details: ", err);
            throw { message: `Internal Server Error in fetching schedule data : ${JSON.stringify(err)}` }
        }
    }
}

export const fetchScheduleMinimal = async () => {
    try {
        const currentDate = moment().subtract(2, 'day').format("YYYY-MM-DD");
        const url = `/schedules/?Type=job&Date=gt(${currentDate})&pageSize=100`;
        let fetchedSimproSchedulesData: SimproScheduleType[] = await fetchSimproPaginatedData(url, "ID,Type,Reference,Staff,Date,Blocks,Notes");
        return fetchedSimproSchedulesData;
    } catch (err) {
        if (err instanceof AxiosError) {
            console.log("Error in fetchScheduleData as AxiosError");
            console.log("Error details: ", err.response?.data);
            throw { message: "Something went wrong while fetching schedule data : " + JSON.stringify(err.response) }
        } else {
            console.log("Error in fetchScheduleData as other error");
            console.log("Error details: ", err);
            throw { message: `Internal Server Error in fetching schedule data : ${JSON.stringify(err)}` }
        }
    }
}


export const fetchCurrentDateScheduleData = async () => {
    try {
        let fetchedChartOfAccounts = await axiosSimPRO.get('/setup/accounts/chartOfAccounts/?pageSize=250&columns=ID,Name,Number');
        let chartOfAccountsArray: SimproAccountType[] = fetchedChartOfAccounts?.data;
        const allCustomerData: SimproCustomerType[] = await fetchSimproPaginatedData('/customers/', "");
        const simproCustomerMap: { [key: string]: SimproCustomerType } = {};

        if (allCustomerData.length > 0) {
            // Fetching customer data sequentially
            for (const customer of allCustomerData) {
                const hrefForCustomerData = customer?._href;
                if (!hrefForCustomerData) continue;

                let customerUrl = hrefForCustomerData.split('customers')[1];
                if (!customerUrl) continue;

                try {
                    let customerDataForJob: SimproCustomerType;
                    if (customerUrl.includes('companies')) {
                        const customerResponse = await axiosSimPRO.get(`/customers${customerUrl}?columns=ID,CompanyName,Phone,Address,Email`);
                        customerDataForJob = { ...customerResponse.data, Type: "Company" };
                    } else {
                        const customerResponse = await axiosSimPRO.get(`/customers${customerUrl}?columns=ID,GivenName,FamilyName,Phone,Address,Email`);
                        customerDataForJob = { ...customerResponse.data, Type: "Individual" };
                    }
                    if (customerDataForJob?.ID) {
                        simproCustomerMap[customerDataForJob.ID.toString()] = customerDataForJob;
                    }
                } catch (error) {
                    console.error("Error fetching customer data for:", customer.ID, error);
                }
            }
        }

        if (Object.keys(simproCustomerMap).length) {
            const currentDate = moment().format("YYYY-MM-DD");
            const url = `/schedules/?Type=job&Date=${currentDate}&pageSize=100`;
            let fetchedSimproSchedulesData: SimproScheduleType[] = await fetchSimproPaginatedData(url, "ID,Type,Reference,Staff,Date,Blocks,Notes");
            let jobIdsToAddArray: number[] = []
            // Fetch job information sequentially
            for (let i = 0; i < fetchedSimproSchedulesData.length; i++) {
                const schedule = fetchedSimproSchedulesData[i];
                let jobIdForSchedule = schedule?.Reference?.split('-')[0];
                let costCenterIdForSchedule = schedule?.Reference?.split('-')[1];
                if (jobIdForSchedule) {
                    try {
                        const jobDataForSchedule = await axiosSimPRO.get(`/jobs/${jobIdForSchedule}?columns=ID,Type,Site,SiteContact,DateIssued,Status,Total,Customer,Name,ProjectManager,CustomFields`);
                        let fetchedJobData: SimproJobType = jobDataForSchedule?.data;
                        let siteId = fetchedJobData?.Site?.ID;
                        if (siteId) {
                            const siteResponse = await axiosSimPRO.get(`/sites/${siteId}?columns=ID,Name,Address`);
                            let siteResponseData = siteResponse.data;
                            fetchedJobData.Site = siteResponseData;
                        }
                        // let jobTradeValue = fetchedJobData?.CustomFields?.find(field => field?.CustomField?.Name === "Job Trade (ie, Plumbing, Drainage, Roofing)")?.Value;
                        // if (jobTradeValue == 'Roofing') {
                        //     jobIdsToAddArray.push(fetchedJobData.ID)
                        // }
                        schedule.Job = fetchedJobData;
                        if (schedule?.Job?.Customer) {
                            const customerId = schedule.Job.Customer.ID?.toString();
                            if (customerId && simproCustomerMap[customerId]) {
                                schedule.Job.Customer = simproCustomerMap[customerId];
                            }
                        }
                    } catch (error) {
                        console.error(`Error fetching job data for schedule ID: ${jobIdForSchedule}`, error);
                    }
                }

                console.log('costCenterIdForSchedule',costCenterIdForSchedule)
                if (costCenterIdForSchedule) {
                    try {
                        const costCenterDataForSchedule = await axiosSimPRO.get(`/jobCostCenters/?ID=${costCenterIdForSchedule}&columns=ID,Name,Job,Section,CostCenter`);
                        let sectionIdForSchedule =
                            Array.isArray(costCenterDataForSchedule?.data)
                                ? costCenterDataForSchedule.data[0]?.Section?.ID
                                : null;

                        let jobIdForScheduleFetched =
                            Array.isArray(costCenterDataForSchedule?.data)
                                ? costCenterDataForSchedule.data[0]?.Job?.ID
                                : null;

                        let setupCostCenterID = costCenterDataForSchedule.data[0]?.CostCenter?.ID;
                        let fetchedSetupCostCenterData = await axiosSimPRO.get(`/setup/accounts/costCenters/${setupCostCenterID}?columns=ID,Name,IncomeAccountNo`);
                        let setupCostCenterData = fetchedSetupCostCenterData.data;
                        if (setupCostCenterData?.IncomeAccountNo) {
                            let incomeAccountName = chartOfAccountsArray?.find(account => account?.Number == setupCostCenterData?.IncomeAccountNo)?.Name;
                            console.log('Income Account Name: ', incomeAccountName);
                            if (incomeAccountName == "Roofing Income") {
                                // console.log('CostCenterId For Roofing Income 2', costCenterIdForSchedule, jobIdForScheduleFetched)
                                jobIdsToAddArray.push(jobIdForScheduleFetched)
                            }
                        }
                        try {
                            let costCenterResponse = await axiosSimPRO.get(`jobs/${jobIdForScheduleFetched}/sections/${sectionIdForSchedule}/costCenters/${costCenterIdForSchedule}?columns=Name,ID,Claimed`);
                            if (costCenterResponse) {
                                schedule.CostCenter = costCenterResponse.data;
                            }
                        } catch (error) {
                            console.log("Error in costCenterFetch : ", error)
                        }
                    } catch (error) {
                        console.error(`Error fetching cost center data for schedule ID: ${costCenterIdForSchedule}`, error);
                    }
                }


            }

            // console.log('jobsToFilter Length: ', jobIdsToAddArray.length)
            // console.log('Current schedule Length: ', fetchedSimproSchedulesData.length)
            fetchedSimproSchedulesData = fetchedSimproSchedulesData.filter(schedule =>
                jobIdsToAddArray.includes(schedule?.Job?.ID ?? -1)
            );
            // console.log('Filtered schedule Length: ', fetchedSimproSchedulesData.length)
            return fetchedSimproSchedulesData;
        } else {
            return [];
        }
    } catch (err) {
        if (err instanceof AxiosError) {
            console.log("Error in fetchScheduleData as AxiosError");
            console.log("Error details: ", err.response?.data);
            throw { message: "Something went wrong while fetching schedule data : " + JSON.stringify(err.response) }
        } else {
            console.log("Error in fetchScheduleData as other error");
            console.log("Error details: ", err);
            throw { message: `Internal Server Error in fetching schedule data : ${JSON.stringify(err)}` }
        }
    }
}

export const fetchNextDateScheduleData = async () => {
    try {
        let fetchedChartOfAccounts = await axiosSimPRO.get('/setup/accounts/chartOfAccounts/?pageSize=250&columns=ID,Name,Number');
        let chartOfAccountsArray: SimproAccountType[] = fetchedChartOfAccounts?.data;
        const allCustomerData: SimproCustomerType[] = await fetchSimproPaginatedData('/customers/', "");
        const simproCustomerMap: { [key: string]: SimproCustomerType } = {};

        if (allCustomerData.length > 0) {
            // Fetching customer data sequentially
            for (const customer of allCustomerData) {
                const hrefForCustomerData = customer?._href;
                if (!hrefForCustomerData) continue;

                let customerUrl = hrefForCustomerData.split('customers')[1];
                if (!customerUrl) continue;

                try {
                    let customerDataForJob: SimproCustomerType;
                    if (customerUrl.includes('companies')) {
                        const customerResponse = await axiosSimPRO.get(`/customers${customerUrl}?columns=ID,CompanyName,Phone,Address,Email`);
                        customerDataForJob = { ...customerResponse.data, Type: "Company" };
                    } else {
                        const customerResponse = await axiosSimPRO.get(`/customers${customerUrl}?columns=ID,GivenName,FamilyName,Phone,Address,Email`);
                        customerDataForJob = { ...customerResponse.data, Type: "Individual" };
                    }
                    if (customerDataForJob?.ID) {
                        simproCustomerMap[customerDataForJob.ID.toString()] = customerDataForJob;
                    }
                } catch (error) {
                    console.error("Error fetching customer data for:", customer.ID, error);
                }
            }
        }

        if (Object.keys(simproCustomerMap).length) {
            const nextDate = moment().add(1, 'days').format("YYYY-MM-DD");
            const url = `/schedules/?Type=job&Date=${nextDate}&pageSize=100`;
            let fetchedSimproSchedulesData: SimproScheduleType[] = await fetchSimproPaginatedData(url, "ID,Type,Reference,Staff,Date,Blocks,Notes");
            let jobIdsToAddArray: number[] = []
            // Fetch job information sequentially
            for (let i = 0; i < fetchedSimproSchedulesData.length; i++) {
                const schedule = fetchedSimproSchedulesData[i];
                let jobIdForSchedule = schedule?.Reference?.split('-')[0];
                let costCenterIdForSchedule = schedule?.Reference?.split('-')[1];
                if (jobIdForSchedule) {
                    try {
                        const jobDataForSchedule = await axiosSimPRO.get(`/jobs/${jobIdForSchedule}?columns=ID,Type,Site,SiteContact,DateIssued,Status,Total,Customer,Name,ProjectManager,CustomFields`);
                        let fetchedJobData: SimproJobType = jobDataForSchedule?.data;
                        let siteId = fetchedJobData?.Site?.ID;
                        if (siteId) {
                            const siteResponse = await axiosSimPRO.get(`/sites/${siteId}?columns=ID,Name,Address`);
                            let siteResponseData = siteResponse.data;
                            fetchedJobData.Site = siteResponseData;
                        }
                        // let jobTradeValue = fetchedJobData?.CustomFields?.find(field => field?.CustomField?.Name === "Job Trade (ie, Plumbing, Drainage, Roofing)")?.Value;
                        // if (jobTradeValue == 'Roofing') {
                        //     jobIdsToAddArray.push(fetchedJobData.ID)
                        // }
                        schedule.Job = fetchedJobData;
                        if (schedule?.Job?.Customer) {
                            const customerId = schedule.Job.Customer.ID?.toString();
                            if (customerId && simproCustomerMap[customerId]) {
                                schedule.Job.Customer = simproCustomerMap[customerId];
                            }
                        }
                    } catch (error) {
                        console.error(`Error fetching job data for schedule ID: ${jobIdForSchedule}`, error);
                    }
                }

                if (costCenterIdForSchedule) {
                    try {
                        const costCenterDataForSchedule = await axiosSimPRO.get(`/jobCostCenters/?ID=${costCenterIdForSchedule}&columns=ID,Name,Job,Section,CostCenter`);
                        let sectionIdForSchedule =
                            Array.isArray(costCenterDataForSchedule?.data)
                                ? costCenterDataForSchedule.data[0]?.Section?.ID
                                : null;

                        let jobIdForScheduleFetched =
                            Array.isArray(costCenterDataForSchedule?.data)
                                ? costCenterDataForSchedule.data[0]?.Job?.ID
                                : null;

                        let setupCostCenterID = costCenterDataForSchedule.data[0]?.CostCenter?.ID;
                        let fetchedSetupCostCenterData = await axiosSimPRO.get(`/setup/accounts/costCenters/${setupCostCenterID}?columns=ID,Name,IncomeAccountNo`);
                        let setupCostCenterData = fetchedSetupCostCenterData.data;
                        if (setupCostCenterData?.IncomeAccountNo) {
                            let incomeAccountName = chartOfAccountsArray?.find(account => account?.Number == setupCostCenterData?.IncomeAccountNo)?.Name;
                            if (incomeAccountName == "Roofing Income") {
                                // console.log('CostCenterId For Roofing Income 1', costCenterIdForSchedule, jobIdForScheduleFetched)
                                jobIdsToAddArray.push(jobIdForScheduleFetched)
                            }
                        }
                        try {
                            let costCenterResponse = await axiosSimPRO.get(`jobs/${jobIdForScheduleFetched}/sections/${sectionIdForSchedule}/costCenters/${costCenterIdForSchedule}?columns=Name,ID,Claimed`);
                            if (costCenterResponse) {
                                schedule.CostCenter = costCenterResponse.data;
                            }
                        } catch (error) {
                            console.log("Error in costCenterFetch : ", error)
                        }
                    } catch (error) {
                        console.error(`Error fetching cost center data for schedule ID: ${costCenterIdForSchedule}`, error);
                    }
                }


            }

            // console.log('jobsToFilter Length: ', jobIdsToAddArray.length)
            // console.log('Current schedule Length: ', fetchedSimproSchedulesData.length)
            fetchedSimproSchedulesData = fetchedSimproSchedulesData.filter(schedule =>
                jobIdsToAddArray.includes(schedule?.Job?.ID ?? -1)
            );
            // console.log('Filtered schedule Length: ', fetchedSimproSchedulesData.length)
            return fetchedSimproSchedulesData;
        } else {
            return [];
        }
    } catch (err) {
        if (err instanceof AxiosError) {
            console.log("Error in fetchScheduleData as AxiosError");
            console.log("Error details: ", err.response?.data);
            throw { message: "Something went wrong while fetching schedule data : " + JSON.stringify(err.response) }
        } else {
            console.log("Error in fetchScheduleData as other error");
            console.log("Error details: ", err);
            throw { message: `Internal Server Error in fetching schedule data : ${JSON.stringify(err)}` }
        }
    }
}

/**
 * Validates if a schedule exists in SimPro API
 * Makes an API call to fetch the schedule and checks if it returns successfully
 * 
 * @param jobID - The Job ID
 * @param sectionID - The Section ID
 * @param costCenterID - The Job Cost Center ID
 * @param scheduleID - The Schedule ID to validate
 * @returns Object with status and details about schedule existence
 * 
 * @example
 * const result = await validateScheduleExistence(123, 456, 789, 101112);
 * // Returns: { exists: true, scheduleID: 101112, message: "Schedule found" }
 * // Or: { exists: false, scheduleID: 101112, message: "Schedule not found or deleted" }
 */
export const validateScheduleExistence = async (
    jobID: number | string,
    sectionID: number | string,
    costCenterID: number | string,
    scheduleID: number | string
) => {
    try {
        const apiUrl = `/jobs/${jobID}/sections/${sectionID}/costCenters/${costCenterID}/schedules/${scheduleID}`;
        console.log(`Validating schedule existence: ${apiUrl}`);
        
        const response = await axiosSimPRO.get(`${apiUrl}?columns=ID,Staff,Date,Blocks,Notes`);
        
        if (response?.data?.ID) {
            console.log(`Schedule ${scheduleID} validation successful - Schedule exists`);
            return {
                exists: true,
                scheduleID: scheduleID,
                message: "Schedule found in SimPro",
                data: response.data
            };
        }
    } catch (err) {
        if (err instanceof AxiosError) {
            console.log(`Schedule ${scheduleID} validation error:`, err.response?.status);
            console.log("Error details:", err.response?.data);
            
            const errors = err?.response?.data?.errors;
            
            // Check if it's a 404 or "Schedule not found" error
            if (
                err.response?.status === 404 ||
                (Array.isArray(errors) && errors.some((e: any) => 
                    e?.message?.includes("not found") || 
                    e?.message?.includes("Schedule not found")
                ))
            ) {
                console.log(`Schedule ${scheduleID} - Schedule not found or deleted in SimPro`);
                return {
                    exists: false,
                    scheduleID: scheduleID,
                    message: "Schedule not found or deleted from SimPro",
                    errorStatus: err.response?.status
                };
            }
            
            // For other API errors, throw
            console.error(`Unexpected error validating schedule ${scheduleID}:`, err.response?.data);
            throw {
                message: "Error validating schedule existence",
                scheduleID: scheduleID,
                error: err.response?.data,
                status: err.response?.status
            };
        } else {
            // For non-Axios errors
            console.error(`Unexpected error validating schedule ${scheduleID}:`, err);
            throw {
                message: "Unexpected error validating schedule existence",
                scheduleID: scheduleID,
                error: err
            };
        }
    }
}

