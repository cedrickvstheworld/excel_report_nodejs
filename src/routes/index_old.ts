// import express,  { Request, Response } from 'express'
// import Excel from 'exceljs'

// class Router {
//     public router: any
//     public workbook: any
//     public DRCS: any

//     constructor() {
//         this.router = express.Router()
//         this.workbook = new Excel.Workbook()
//         this.setWorkBookProps()
//     }

//     setWorkBookProps() {
//         this.workbook.creator = 'Jed'
//         this.workbook.lastModifiedBy = 'Jed'
//         this.workbook.created = new Date()
//         this.workbook.modified = new Date()
//         this.workbook.lastPrinted = new Date()
//         this.workbook.views = [
//             {
//               x: 0, y: 0, width: 10000, height: 20000,
//               firstSheet: 0, activeTab: 1, visibility: 'visible'
//             }
//           ]
//         this.DRCS = this.workbook.addWorksheet('Daily RO CHASE SHEET')
//     }

//     // test
//     setUpWorkSheet() {
//     }

//     // configure column names
//     setColumnNames(columnNames: Array<any>) {
//         const colNames: any = []
//         columnNames.forEach((i) => {
//             colNames.push(
//                 {header: i.name , key: i.key, width: i.width}
//             )
//         })
//         this.DRCS.columns = colNames
        
//         const repairCell = this.DRCS.getRow(1).getCell('repair')
//         repairCell.columns = [
//             {name: 'Start', key: 'start', width: 10},
//             {name: 'End', key: 'end', width: 10},
//         ]
//     }
    
//     initializeExcelFile() {
//         this.workbook.xlsx.writeFile('testFile.xlsx')
//         .then(() => {
//             console.log('success write')
//         })
//         .catch((error: string) => {
//             console.log(error)
//         })
//     }

//     preparedRoutes() {
//         this.router.get('/', (req: Request, res: Response) => {
            

//             this.setColumnNames([
//                 {name: 'Count', key: 'count', width: 20},
//                 {name: 'Date', key: 'date', width: 20},
//                 {name: 'Plate No.', key: 'plate no', width: 20},
//                 {name: 'Model', key: 'model', width: 20},
//                 {name: 'W/ NW', key: 'w/ nw', width: 20},
//                 {name: 'Repeair Type', key: 'repair type', width: 20},
//                 {name: 'Bay Number', key: 'bay number', width: 20},
//                 {name: 'KM Check-up', key: 'km check-up', width: 20},
//                 {name: 'Series', key: 'series', width: 20},
//                 {name: 'Ancillary Service', key: 'ancillary service', width: 20},
//                 {name: 'Additional Job', key: 'additional job', width: 20},
//                 {name: 'Customer Arrival', key: 'customer arrival', width: 20},
//                 {name: 'Appointment Time', key: 'appointment time', width: 20},
//                 {name: 'Customer Arrival Time', key: 'customer arrival time', width: 26},
//                 {name: 'Ticket Time', key: 'ticket time', width: 20},
//                 {name: 'Customer Waiting Time', key: 'customer waiting time', width: 26},
//                 {name: 'Service Advisor', key: 'service advisor', width: 20},
//                 {name: 'SA RECEPTION TIME', key: 'sa reception time', width: 26},
//                 {name: 'LEAD TIME', key: 'lead time', width: 20},
//                 {name: 'IDLE TIME', key: 'idle time a', width: 20},
//                 {name: 'Technician Clock-in Time', key: 'technician clock-in time', width: 30},
//                 {name: 'IDLE TIME', key: 'idle time b', width: 20},
//                 {name: 'Repair', key: 'repair', width: 20, height: 20},
//                 // {name: '', key: 'idle time b', width: 20},
//                 // {name: 'IDLE TIME', key: 'idle time b', width: 20},
//                 // {name: 'IDLE TIME', key: 'idle time b', width: 20},
//                 // {name: 'IDLE TIME', key: 'idle time b', width: 20},
//                 // {name: 'IDLE TIME', key: 'idle time b', width: 20},
//                 // {name: 'IDLE TIME', key: 'idle time b', width: 20},
//                 // {name: 'IDLE TIME', key: 'idle time b', width: 20},
//                 // {name: 'IDLE TIME', key: 'idle time b', width: 20},

                
//             ])

//             this.initializeExcelFile()
            
//             res.send('test')
//         })
//         return this.router
//     }

// }

// export default new Router()