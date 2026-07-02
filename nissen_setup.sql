-- ============================================================
-- Nissen 代購訂單系統 — Supabase SQL Setup
-- 請在 Supabase > SQL Editor 執行此 SQL
-- ============================================================

-- 1. nissen_recommended 推薦商品表
CREATE TABLE IF NOT EXISTS nissen_recommended (
  id          SERIAL PRIMARY KEY,
  name        TEXT NOT NULL,
  description TEXT    DEFAULT '',
  image_url   TEXT    DEFAULT '',
  nissen_url  TEXT    DEFAULT '',
  price_hkd   NUMERIC DEFAULT 0,
  price_twd   NUMERIC DEFAULT 0,
  price_jpy   NUMERIC DEFAULT 0,
  sort_order  INTEGER DEFAULT 0,
  active      BOOLEAN DEFAULT true,
  created_at  TIMESTAMPTZ DEFAULT NOW(),
  updated_at  TIMESTAMPTZ DEFAULT NOW()
);

-- 2. nissen_orders 訂單表
CREATE TABLE IF NOT EXISTS nissen_orders (
  id             SERIAL PRIMARY KEY,
  order_no       TEXT UNIQUE NOT NULL,
  region         TEXT NOT NULL,             -- hk | tw
  customer_name  TEXT NOT NULL,
  customer_email TEXT NOT NULL,
  customer_phone TEXT    DEFAULT '',
  items          JSONB   NOT NULL,          -- [{type, name, desc, qty, jpy, local, url, color, size, image_url}]
  remark         TEXT    DEFAULT '',
  status         TEXT    DEFAULT 'pending', -- pending | paid | confirmed | cancelled
  created_at     TIMESTAMPTZ DEFAULT NOW(),
  confirmed_at   TIMESTAMPTZ
);

-- 3. RLS
ALTER TABLE nissen_recommended ENABLE ROW LEVEL SECURITY;
ALTER TABLE nissen_orders      ENABLE ROW LEVEL SECURITY;

CREATE POLICY "nissen_rec_read"   ON nissen_recommended FOR SELECT USING (true);
CREATE POLICY "nissen_rec_write"  ON nissen_recommended FOR ALL    USING (true) WITH CHECK (true);
CREATE POLICY "nissen_orders_all" ON nissen_orders      FOR ALL    USING (true) WITH CHECK (true);
