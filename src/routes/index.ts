import express,  { Request, Response } from 'express'
import GenExcel from '../Tools/CreateExcelFile'
import { ExcelReport } from '../types/interfaces' 


class Router {
    public router: any

    constructor() {
        this.router = express.Router()
    }

    preparedRoutes() {
        this.router.get('/', (req: Request, res: Response) => {

            //test data
            const data: Array<ExcelReport> = [
                {
                    date: 'Jan 21, 2019',
                    plate_no: 'AWER-23',
                    // model: 'honda',
                    w_nw: 'w',
                    repair_type: 'EM',
                    bay_number: '89',
                    km_check_up: '10',
                    series: '1K/5K FPM',
                    as_c: 'WITH',
                    as_details: 'Nitro',
                    aj_c: 'None/ EM ITEM',
                    aj_details: 'BRAKE PADS',
                    customer_arrival: '7:00',
                    appointment_time: '7:10',
                    cat_guard: '7:20',
                    cat_receptionist: '8:32',
                    ticketing_time: '7:50',
                    customer_wt: '0:17',
                    sa: 'Cedrick so Cool',
                    sa_rt_s: '8:10',
                    sa_rt_e: '8:30',
                    reception_lead_time: '0:10',
                    reception_idle_time: '0:16',
                    technician_cit: '8:56',
                    technician_idle_time_0: '0:18',
                    technician_idle_time_1: '0:28',
                    technician_idle_time_2: '0:58',
                    repair_s: '11:27',
                    repair_e: '12:11',
                    repair_lead_time: '0:18',
                    repair_idle_time: '0:17',
                    carwash_s: '12:20',
                    carwash_e: '14:11',
                    carwash_lead_time: '0:40',
                    release_date: '7/15/2019',
                    car_out_time: '16:00',
                    total_em_lt: '2:06',
                    total_lt: '2:20',
                    promised_time: '13:00',
                    otd: 'On Time Delivery',
                    em60_ach: 'NOT EM60',
                    remarks: 'remarks here'
                }       
            ]

            const filename: string = 'testName'
            const createExcelReport: any = new GenExcel(data, filename)  
            createExcelReport.saveNew()
                .then((file: string) => {
                    console.log(file)
                    res.status(200).download(file)
                })
                .catch((error: string) => {
                    console.log(error)
                    res.status(404)
                })
        })
        return this.router
    }

}

export default new Router()