import express from 'express'
import { Application } from 'express'

import indexRoute from './routes/index'

class App {
    public app: Application
    public PORT: (string | number)

    constructor() {
        this.app = express()
        this.PORT = 8000
        this.app.use('', indexRoute.preparedRoutes())
    }

    public createServer() {
        this.app.listen(8000, (): void => {
                console.log(`listening to port ${this.PORT} . . .`)
        })
    }
}


const Main = new App()
Main.createServer()