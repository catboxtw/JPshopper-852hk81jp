-- ============================================================
-- 收緊 RLS — Supabase SQL Setup（喺 SQL Editor 行一次）
-- ------------------------------------------------------------
-- 之前每張表都係 FOR ALL USING (true) WITH CHECK (true)，
-- 而 anon key 一定會出現喺前台原始碼度（前台要用嚟讀商品）。
-- 兩樣加埋 = 任何人揭開個網頁原始碼，就讀晒所有訂單（姓名／電話／
-- email／取貨地址），仲改得刪得，連匯率同價錢都改到。
--
-- 而家嘅分工：
--   anon（客人瀏覽器）  只讀得商品目錄、只落得單，讀唔到人哋張單
--   service_role（後台） 由 GAS 驗完密碼先派落嚟，本身繞過 RLS
--   查自己張單           行下面幾個 SECURITY DEFINER function，要對得住身份先俾
--
-- ⚠️ 行完之後即刻試：落單、查單、後台。有問題就 rollback 返最底嗰段。
-- ============================================================

-- ── 1. 目錄類：公開讀得，但唔准 anon 寫 ──────────────────
-- （後台改嘢用 service_role，唔受 RLS 限制）

DROP POLICY IF EXISTS "app_settings_write"  ON app_settings;
DROP POLICY IF EXISTS "flash_stock_write"   ON flash_stock;
DROP POLICY IF EXISTS "nissen_rec_write"    ON nissen_recommended;

DROP POLICY IF EXISTS "app_settings_read"   ON app_settings;
CREATE POLICY "app_settings_read" ON app_settings       FOR SELECT USING (true);
DROP POLICY IF EXISTS "flash_stock_read"    ON flash_stock;
CREATE POLICY "flash_stock_read"  ON flash_stock        FOR SELECT USING (true);
DROP POLICY IF EXISTS "nissen_rec_read"     ON nissen_recommended;
CREATE POLICY "nissen_rec_read"   ON nissen_recommended FOR SELECT USING (true);

-- 活動商品表。呢三張本身可能未開過 RLS —— 未開即係任何人改得，
-- 所以照開，然後只俾讀。
ALTER TABLE events        ENABLE ROW LEVEL SECURITY;
ALTER TABLE products      ENABLE ROW LEVEL SECURITY;
ALTER TABLE product_subs  ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "events_read"       ON events;
CREATE POLICY "events_read"       ON events       FOR SELECT USING (true);
DROP POLICY IF EXISTS "products_read"     ON products;
CREATE POLICY "products_read"     ON products     FOR SELECT USING (true);
DROP POLICY IF EXISTS "product_subs_read" ON product_subs;
CREATE POLICY "product_subs_read" ON product_subs FOR SELECT USING (true);

-- ── 2. 訂單類：客人只落得單，讀唔到、改唔到 ──────────────
DROP POLICY IF EXISTS "nissen_orders_all"    ON nissen_orders;
DROP POLICY IF EXISTS "nissen_orders_insert" ON nissen_orders;
CREATE POLICY "nissen_orders_insert" ON nissen_orders FOR INSERT WITH CHECK (true);
-- 冇 SELECT／UPDATE／DELETE policy ＝ anon 一律做唔到

-- flash 訂單由頭到尾行 function（搶貨、查單、交資料），所以 anon 對呢張表零權限
DROP POLICY IF EXISTS "flash_orders_all" ON flash_orders;

-- ── 3. 客人查自己張單 ────────────────────────────────────
-- SECURITY DEFINER ＝ 呢個 function 用表主人嘅身份行，繞得過上面啲 policy，
-- 但佢只會回一行、而且要你答得出身份先回。

-- 單號係 NK-YYYYMMDD-XXXX，尾嗰四個字隨機 —— 得四個字元係撞得中嘅，
-- 所以唔可以當佢係密碼，要再對 email 或者電話尾四位。
CREATE OR REPLACE FUNCTION get_nissen_order(p_order_no TEXT, p_contact TEXT)
RETURNS SETOF nissen_orders
LANGUAGE sql SECURITY DEFINER SET search_path = public AS $$
  SELECT * FROM nissen_orders
   WHERE order_no = p_order_no
     AND (
       lower(customer_email) = lower(btrim(p_contact))
       OR (length(regexp_replace(p_contact, '\D', '', 'g')) >= 4
           AND right(regexp_replace(customer_phone, '\D', '', 'g'), 4)
             = right(regexp_replace(p_contact,     '\D', '', 'g'), 4))
     )
   LIMIT 1;
$$;

-- flash 嗰個參考編號係 32 位隨機 hex 嘅頭 8 位，本身已經夠難撞，
-- 所以唔使再問多樣嘢。但一定要夠長先俾查，唔可以打兩個字掃晒成張表。
CREATE OR REPLACE FUNCTION get_flash_order(p_ref TEXT)
RETURNS TABLE (token TEXT, item_id INTEGER, region TEXT, status TEXT,
               customer_name TEXT, pay_method TEXT, remark TEXT,
               claimed_at TIMESTAMPTZ, submitted_at TIMESTAMPTZ, paid_at TIMESTAMPTZ)
LANGUAGE sql SECURITY DEFINER SET search_path = public AS $$
  SELECT o.token, o.item_id, o.region, o.status,
         o.customer_name, o.pay_method, o.remark,
         o.claimed_at, o.submitted_at, o.paid_at
    FROM flash_orders o
   WHERE length(regexp_replace(lower(p_ref), '[^a-f0-9]', '', 'g')) >= 8
     AND o.token LIKE regexp_replace(lower(p_ref), '[^a-f0-9]', '', 'g') || '%'
   LIMIT 1;
$$;

-- 搶到貨之後填資料。要有成條 token 先改到，而且改咗一次就唔再收
-- （唔係就有人可以拎住條 token 不停改人哋張單）。
CREATE OR REPLACE FUNCTION submit_flash_order(
  p_token TEXT, p_name TEXT, p_phone TEXT, p_email TEXT,
  p_pay_method TEXT, p_remark TEXT)
RETURNS JSON
LANGUAGE plpgsql SECURITY DEFINER SET search_path = public AS $$
DECLARE v_n INTEGER;
BEGIN
  UPDATE flash_orders
     SET customer_name = p_name, phone = p_phone, email = p_email,
         pay_method = p_pay_method, remark = p_remark,
         status = 'submitted', submitted_at = NOW()
   WHERE token = p_token AND status = 'claimed';
  GET DIAGNOSTICS v_n = ROW_COUNT;
  IF v_n = 0 THEN
    RETURN json_build_object('success', false, 'message', '搵唔到呢張單，或者已經交咗');
  END IF;
  RETURN json_build_object('success', true);
END;
$$;

REVOKE ALL ON FUNCTION get_nissen_order(TEXT, TEXT)                        FROM PUBLIC;
REVOKE ALL ON FUNCTION get_flash_order(TEXT)                               FROM PUBLIC;
REVOKE ALL ON FUNCTION submit_flash_order(TEXT,TEXT,TEXT,TEXT,TEXT,TEXT)   FROM PUBLIC;
GRANT EXECUTE ON FUNCTION get_nissen_order(TEXT, TEXT)                      TO anon, authenticated;
GRANT EXECUTE ON FUNCTION get_flash_order(TEXT)                             TO anon, authenticated;
GRANT EXECUTE ON FUNCTION submit_flash_order(TEXT,TEXT,TEXT,TEXT,TEXT,TEXT) TO anon, authenticated;

-- ── 出事嘅話，行返呢段就即刻返到舊樣（但即係又全開返，唔好留喺度）──
-- CREATE POLICY "nissen_orders_all" ON nissen_orders FOR ALL USING (true) WITH CHECK (true);
-- CREATE POLICY "flash_orders_all"  ON flash_orders  FOR ALL USING (true) WITH CHECK (true);
-- CREATE POLICY "app_settings_write" ON app_settings FOR ALL USING (true) WITH CHECK (true);
-- CREATE POLICY "flash_stock_write"  ON flash_stock  FOR ALL USING (true) WITH CHECK (true);
-- CREATE POLICY "nissen_rec_write"   ON nissen_recommended FOR ALL USING (true) WITH CHECK (true);
