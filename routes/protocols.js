import express from 'express'
import { supabase } from '../lib/supabase.js'

const router = express.Router()

// ── AIRCRAFT ──────────────────────────────────────────────────────────────────

router.get('/aircraft', async (req, res) => {
  const { data, error } = await supabase
    .from('aircraft_types').select('*').order('name')
  if (error) return res.status(400).json({ error: error.message })
  res.json(data)
})

router.post('/aircraft', async (req, res) => {
  const { name, engine } = req.body
  const { data, error } = await supabase
    .from('aircraft_types').insert({ name, engine }).select().single()
  if (error) return res.status(400).json({ error: error.message })
  res.json(data)
})

router.delete('/aircraft/:id', async (req, res) => {
  const { error } = await supabase
    .from('aircraft_types').delete().eq('id', req.params.id)
  if (error) return res.status(400).json({ error: error.message })
  res.json({ success: true })
})

// ── TASKS (before /:id to avoid route conflict) ───────────────────────────────

router.put('/tasks/:id', async (req, res) => {
  const { ata, operation_performed, maintenance_record_ref, time_duration, order, is_active } = req.body
  const { data, error } = await supabase
    .from('tasks')
    .update({ ata, operation_performed, maintenance_record_ref, time_duration, order, is_active })
    .eq('id', req.params.id)
    .select().single()
  if (error) return res.status(400).json({ error: error.message })
  res.json(data)
})

router.delete('/tasks/:id', async (req, res) => {
  const { error } = await supabase
    .from('tasks').delete().eq('id', req.params.id)
  if (error) return res.status(400).json({ error: error.message })
  res.json({ success: true })
})

// ── PROTOCOLS ─────────────────────────────────────────────────────────────────

router.get('/', async (req, res) => {
  const { data, error } = await supabase
    .from('protocols')
    .select('*, aircraft_types(name)')
    .order('created_at', { ascending: false })
  if (error) return res.status(400).json({ error: error.message })
  res.json(data)
})

router.post('/', async (req, res) => {
  const { aircraft_type_id, protocol_type, created_by, auto_fill } = req.body
  const { data, error } = await supabase
    .from('protocols')
    .insert({
      aircraft_type_id, protocol_type, created_by, auto_fill: auto_fill || false
    })
    .select('*, aircraft_types(name)')
    .single()
  if (error) return res.status(400).json({ error: error.message })
  res.json(data)
})

router.put('/:id', async (req, res) => {
  const { aircraft_type_id, protocol_type, auto_fill } = req.body
  const { data, error } = await supabase
    .from('protocols')
    .update({
      aircraft_type_id, protocol_type, auto_fill: auto_fill || false
    })
    .eq('id', req.params.id)
    .select('*, aircraft_types(name)')
    .single()
  if (error) return res.status(400).json({ error: error.message })
  res.json(data)
})

router.delete('/:id', async (req, res) => {
  const { error } = await supabase
    .from('protocols').delete().eq('id', req.params.id)
  if (error) return res.status(400).json({ error: error.message })
  res.json({ success: true })
})

// ── PROTOCOL TASKS ────────────────────────────────────────────────────────────

router.get('/:id/tasks', async (req, res) => {
  const { data, error } = await supabase
    .from('tasks')
    .select('*')
    .eq('protocol_id', req.params.id)
    .order('order')
  if (error) return res.status(400).json({ error: error.message })
  res.json(data)
})

router.post('/:id/tasks', async (req, res) => {
  const { ata, operation_performed, maintenance_record_ref, time_duration, order } = req.body
  const { data, error } = await supabase
    .from('tasks')
    .insert({
      protocol_id: req.params.id,
      ata, operation_performed,
      maintenance_record_ref, time_duration, order
    })
    .select().single()
  if (error) return res.status(400).json({ error: error.message })
  res.json(data)
})

export default router