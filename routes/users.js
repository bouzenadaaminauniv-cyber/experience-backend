import express from 'express'
import { supabase } from '../lib/supabase.js'

const router = express.Router()

router.post('/', async (req, res) => {
  const { email, password, role } = req.body

  const { data: authData, error: authError } = await supabase.auth.admin.createUser({
    email,
    password: password || 'Technics2024!',
    email_confirm: true,
    user_metadata: { role }
  })

  if (authError) return res.status(400).json({ error: authError.message })

  const { error: profileError } = await supabase
    .from('users')
    .update({ role: role || 'user' })
    .eq('id', authData.user.id)

  if (profileError) return res.status(400).json({ error: profileError.message })

  res.json({ success: true })
})

router.put('/:id/password', async (req, res) => {
  const { password } = req.body
  const { error } = await supabase.auth.admin.updateUserById(req.params.id, { password })
  if (error) return res.status(400).json({ error: error.message })
  res.json({ success: true })
})

router.delete('/:id', async (req, res) => {
  const { error } = await supabase.auth.admin.deleteUser(req.params.id)
  if (error) return res.status(400).json({ error: error.message })
  res.json({ success: true })
})

router.put('/:id/profile', async (req, res) => {
  const { name, surname, date_of_birth, place_of_birth, aht_id, authorization_nr, scope_of_authorization, role } = req.body
  const { error } = await supabase
    .from('users')
    .update({ name, surname, date_of_birth: date_of_birth || null, place_of_birth, aht_id, authorization_nr, scope_of_authorization, role })
    .eq('id', req.params.id)
  if (error) return res.status(400).json({ error: error.message })
  res.json({ success: true })
})
export default router