import Excel from 'exceljs'
import fs from 'fs'
import { ExcelReport } from '../types/interfaces' 


export default class GenExcel {
    public workbook: any
    public worksheet: any
    public startRowMain: number
    private data: Array<ExcelReport>
    public filename: string
    public folder_name: string

    constructor(data: Array<ExcelReport>, filename: string) {
        this.workbook = new Excel.Workbook()
        this.startRowMain = 9
        this.data = data
        this.folder_name = 'excel-reports'
        this.filename = filename.replace(/ /g, '')
    }

    public async createFile() {
        await this.workbook.xlsx.readFile('excel_template.xlsx')
            .then(() => {
                
                this.worksheet = this.workbook.getWorksheet(1)
                this.populateMain()
            })
    }

    // approach one - lazy type ####################################################################
    // private populateMain() {
    //     let rowToModify = this.startRowMain
    //     this.data.forEach((rowData) => {
    //         let currentRow = this.worksheet.getRow(rowToModify)

    //         let currentCol = 3

    //         for (let key in rowData) {
    //             if (rowData.hasOwnProperty(key)) {
    //                 currentRow.getCell(currentCol).value = rowData[key]
    //                 currentCol++
    //             }
    //         }
    //         currentRow.commit()
    //         rowToModify++
    //     })
    // }
    // end of section ################################################################################

    // approach two - hardcoded ######################################################################
    private populateMain() {
        let rowToModify = this.startRowMain
        this.data.forEach((rowData) => {
            let currentRow = this.worksheet.getRow(rowToModify)

            const {
                date,
                plate_no,
                model,
                w_nw,
                repair_type,
                bay_number,
                km_check_up,
                series,
                as_c,
                as_details,
                aj_c,
                aj_details,
                customer_arrival,
                appointment_time,
                cat_guard,
                cat_receptionist,
                ticketing_time,
                customer_wt,
                sa,
                sa_rt_s,
                sa_rt_e,
                reception_lead_time,
                reception_idle_time,
                technician_cit,
                technician_idle_time_0,
                technician_idle_time_1,
                technician_idle_time_2,
                repair_s,
                repair_e,
                repair_lead_time,
                repair_idle_time,
                carwash_s,
                carwash_e,
                carwash_lead_time,
                release_date,
                car_out_time,
                total_em_lt,
                total_lt,
                promised_time,
                otd,
                em60_ach,
                remarks
            } = rowData

            currentRow.getCell(3).value = date
            currentRow.getCell(4).value = plate_no
            currentRow.getCell(5).value = model
            currentRow.getCell(6).value = w_nw
            currentRow.getCell(7).value = repair_type
            currentRow.getCell(8).value = bay_number
            currentRow.getCell(9).value = km_check_up
            currentRow.getCell(10).value = series
            currentRow.getCell(11).value = as_c
            currentRow.getCell(12).value = as_details
            currentRow.getCell(13).value = aj_c
            currentRow.getCell(14).value = aj_details
            currentRow.getCell(15).value = customer_arrival
            currentRow.getCell(16).value = appointment_time
            currentRow.getCell(17).value = cat_guard
            currentRow.getCell(18).value = cat_receptionist
            currentRow.getCell(19).value = ticketing_time
            currentRow.getCell(20).value = customer_wt
            currentRow.getCell(21).value = sa
            currentRow.getCell(22).value = sa_rt_s
            currentRow.getCell(23).value = sa_rt_e
            currentRow.getCell(24).value = reception_lead_time
            currentRow.getCell(25).value = reception_idle_time
            currentRow.getCell(26).value = technician_cit
            currentRow.getCell(27).value = technician_idle_time_0
            currentRow.getCell(28).value = technician_idle_time_1
            currentRow.getCell(29).value = technician_idle_time_2
            currentRow.getCell(30).value = repair_s
            currentRow.getCell(31).value = repair_e
            currentRow.getCell(32).value = repair_lead_time
            currentRow.getCell(33).value = repair_idle_time
            currentRow.getCell(34).value = carwash_s
            currentRow.getCell(35).value = carwash_e
            currentRow.getCell(36).value = carwash_lead_time
            currentRow.getCell(37).value = release_date
            currentRow.getCell(38).value = car_out_time
            currentRow.getCell(39).value = total_em_lt
            currentRow.getCell(40).value = total_lt
            currentRow.getCell(41).value = promised_time
            currentRow.getCell(42).value = otd
            currentRow.getCell(43).value = em60_ach
            currentRow.getCell(44).value = remarks
            currentRow.commit()
            rowToModify++
        })
    }
    // end of section ############################################################################

    public async saveNew() {
        const file: string = `./${this.folder_name}/${this.filename}.xlsx`
        await this.createFile()
        .then(() => {
            fs.access(this.folder_name, (error) => {
                if (error && error.code === 'ENOENT') {
                    fs.mkdir(this.folder_name, () => {})
                }
            })
            this.workbook.xlsx.writeFile(file)
        })
        
        return Promise.resolve(file) 

    }

}