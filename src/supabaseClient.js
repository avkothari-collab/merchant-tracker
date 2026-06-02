import { createClient } from "@supabase/supabase-js";

// Your project's connection details (the publishable key is safe to ship in the browser).
const SUPABASE_URL = "https://gcuyehjqazqpzlanlobp.supabase.co";
const SUPABASE_KEY = "sb_publishable_dnqks5hHA5lsZh8xaj8s3g_HI5-jKcX";

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
