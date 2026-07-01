-- ============================================================
-- 現貨搶購 Flash Stock System — Supabase SQL Setup
-- 請在 Supabase > SQL Editor 執行此 SQL
-- ============================================================

-- 1. flash_stock 商品表
CREATE TABLE IF NOT EXISTS flash_stock (
  id            SERIAL PRIMARY KEY,
  slug          TEXT UNIQUE NOT NULL,
  name          TEXT NOT NULL,
  description   TEXT    DEFAULT '',
  image_url     TEXT    DEFAULT '',
  price_hkd     NUMERIC DEFAULT 0,
  price_twd     NUMERIC DEFAULT 0,
  price_jpy     NUMERIC DEFAULT 0,
  qty_total     INTEGER DEFAULT 1,
  qty_available INTEGER DEFAULT 1,
  status        TEXT    DEFAULT 'active',   -- active | sold_out | hidden
  created_at    TIMESTAMPTZ DEFAULT NOW(),
  updated_at    TIMESTAMPTZ DEFAULT NOW()
);

-- 2. flash_orders 訂單表
CREATE TABLE IF NOT EXISTS flash_orders (
  id             SERIAL PRIMARY KEY,
  item_id        INTEGER REFERENCES flash_stock(id),
  token          TEXT UNIQUE NOT NULL,
  region         TEXT NOT NULL,              -- hk | tw
  status         TEXT DEFAULT 'claimed',     -- claimed | submitted | paid | cancelled
  customer_name  TEXT,
  phone          TEXT,
  email          TEXT,
  pay_method     TEXT,
  remark         TEXT,
  claimed_at     TIMESTAMPTZ DEFAULT NOW(),
  submitted_at   TIMESTAMPTZ,
  paid_at        TIMESTAMPTZ
);

-- 3. RLS（Row Level Security）
ALTER TABLE flash_stock  ENABLE ROW LEVEL SECURITY;
ALTER TABLE flash_orders ENABLE ROW LEVEL SECURITY;

-- 公開讀取商品（前端搶購頁用）
CREATE POLICY "flash_stock_read"  ON flash_stock  FOR SELECT USING (true);
-- Admin 寫入（前端 admin 有密碼保護）
CREATE POLICY "flash_stock_write" ON flash_stock  FOR ALL    USING (true) WITH CHECK (true);
-- 訂單全部操作（grab RPC 用 SECURITY DEFINER，submit/paid 用 anon key）
CREATE POLICY "flash_orders_all"  ON flash_orders FOR ALL    USING (true) WITH CHECK (true);

-- 4. 原子搶購 RPC function（防止兩人同時搶到同一件）
CREATE OR REPLACE FUNCTION grab_flash_item(p_slug TEXT, p_region TEXT)
RETURNS JSON
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  v_token   TEXT;
  v_item_id INTEGER;
BEGIN
  -- 原子扣減：只有 qty_available > 0 且 status='active' 才成功
  UPDATE flash_stock
  SET
    qty_available = qty_available - 1,
    status        = CASE WHEN (qty_available - 1) <= 0 THEN 'sold_out' ELSE 'active' END,
    updated_at    = NOW()
  WHERE slug = p_slug
    AND qty_available > 0
    AND status = 'active'
  RETURNING id INTO v_item_id;

  IF NOT FOUND THEN
    RETURN json_build_object('success', false, 'error', 'sold_out');
  END IF;

  -- 生成 32 位 hex token（作為訂單憑證）
  v_token := encode(gen_random_bytes(16), 'hex');

  INSERT INTO flash_orders (item_id, token, region, status, claimed_at)
  VALUES (v_item_id, v_token, p_region, 'claimed', NOW());

  RETURN json_build_object('success', true, 'token', v_token, 'item_id', v_item_id);
END;
$$;

-- 授權 anon 角色呼叫 RPC
GRANT EXECUTE ON FUNCTION grab_flash_item TO anon;
