import express, { Request, Response } from 'express';
import dotenv from 'dotenv';
import mongoose from 'mongoose';
dotenv.config({ path: `.env.${process.env.NODE_ENV}` });
const app = express();
app.use(express.json({ limit: '10mb' }));
import simproRoutes from './routes/simproRoute';
import creditorsWatchRoutes from './routes/creditorsWatchRoutes';
import smartSheetRoutes from './routes/smartSheetRoutes';
import redisRoutes from './routes/redisRoutes';
import { ExpressAdapter } from '@bull-board/express';
import { createBullBoard } from '@bull-board/api';
import { BullAdapter } from '@bull-board/api/bullAdapter';
import { simproWebhookQueue } from './queues/queue';

const serverAdapter = new ExpressAdapter();
serverAdapter.setBasePath('/admin/queues');

createBullBoard({
    queues: [new BullAdapter(simproWebhookQueue)], // Add your queues here
    serverAdapter,
});


const PORT: number = parseInt(process.env.PORT as string, 10) || 6001;

console.log("ENV PATH", `.env.${process.env.NODE_ENV}`)

if (process.env.NODE_ENV === 'production') {
    const cronJobs = [
        './cron/createUpdateContactsDataScheduler',
        './cron/createUpdateInvoiceCreditNoteScheduler',
        './cron/deleteDataScheduler',
        './cron/updateLateFeeScheduler',
        './cron/roofingScheduleDeletedCheckfromDBScheduler',
        // './cron/taskWorkingHourScheduler',
        // './cron/jobCardScheduler',
        // './cron/jobCardMinimalUpdateScheduler',
        // './cron/ongoingQuotationsAndLeadsScheduler',
        './cron/failedJobsEmailReporter'
    ];
    cronJobs.forEach(job => {
        require(job);
    });
}

// For local Development
// if (process.env.NODE_ENV === 'development') {
//     const cronJobs = [
//         './cron/ongoingQuotationsAndLeadsScheduler'
//     ];
//     cronJobs.forEach(job => {
//         require(job);
//     });
// }


app.use('/admin/queues', serverAdapter.getRouter());
app.use('/api/smartsheet', smartSheetRoutes);
app.use('/api/creditorswatch', creditorsWatchRoutes);
app.use('/api/simpro', simproRoutes);
app.use('/api/redis', redisRoutes);


app.get('/', (req: Request, res: Response) => {
    res.send('SPDR Server is running!!');
});

mongoose.connect(process.env.DB_URL as string).then(() => {
    console.log('MongoDB Connected...');
}).catch((error) => {
    console.log(error);
})

app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
