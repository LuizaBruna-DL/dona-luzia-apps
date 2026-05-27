import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://rgsgjqnlvyossuykemjl.supabase.co'
const supabaseKey = 'sb_publishable_uC6SF60PfuFnZfDDLRo2Wg_m7l2FoBW'

export const supabase = createClient(supabaseUrl, supabaseKey)