-- ============================================================
-- 網站設定（匯率等）— Supabase SQL Setup
-- 請在 Supabase > SQL Editor 執行一次
-- 之後喺 admin 後台「⚙️ 匯率設定」就可以自己改，唔使再掂 code
-- ============================================================

CREATE TABLE IF NOT EXISTS app_settings (
  key        TEXT PRIMARY KEY,
  value      TEXT NOT NULL,
  note       TEXT DEFAULT '',
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

ALTER TABLE app_settings ENABLE ROW LEVEL SECURITY;

-- 前台要讀匯率嚟報價，所以公開讀取
CREATE POLICY "app_settings_read"  ON app_settings FOR SELECT USING (true);
-- Admin 修改（admin 頁面有密碼保護）
CREATE POLICY "app_settings_write" ON app_settings FOR ALL    USING (true) WITH CHECK (true);

-- 預設值（Nissen 同 ZOZOTOWN 共用同一組匯率）
INSERT INTO app_settings (key, value, note) VALUES
  ('rate_hk',      '0.057', '日幣→港幣匯率（稅込価格換算）'),
  ('rate_tw',      '0.24',  '日幣→台幣匯率（稅込価格換算）'),
  ('ship_hk_50g',  '4.5',   '到港運費 HK$／50g'),
  ('ship_tw_50g',  '20',    '到台運費 NT$／50g')
ON CONFLICT (key) DO NOTHING;
