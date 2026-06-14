import express from 'express';
import { supabase } from '../lib/supabase.js';

const router = express.Router();

// GET /settings/:key
router.get('/:key', async (req, res) => {
  const { key } = req.params;
  const { data, error } = await supabase
    .from('settings')
    .select('value')
    .eq('key', key)
    .single();

  if (error) return res.status(500).json({ error: error.message });
  res.json({ value: data?.value || null });
});

export default router;