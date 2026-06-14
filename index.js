import express from 'express'
import cors from 'cors'
import dotenv from 'dotenv'
import usersRouter from './routes/users.js'
import protocolsRouter from './routes/protocols.js'
import exportRouter from './routes/export.js'
import settingsRouter from './routes/settings.js'
dotenv.config()

const app = express()
app.use(cors({ origin: 'http://localhost:5173' }))
app.use(express.json())

app.use('/api/settings', settingsRouter)
app.get('/health', (req, res) => res.json({ status: 'ok' }))
app.use('/api/users', usersRouter)
app.use('/api/protocols', protocolsRouter)
app.use('/api/export', exportRouter)

app.listen(process.env.PORT, () => {
  console.log(`Backend running on port ${process.env.PORT}`)
})