-- ============================================
-- SEED: Default admin user + config
-- Password: admin123 (bcrypt hash)
-- ============================================

INSERT IGNORE INTO users (username, password_hash, full_name, role, language) VALUES
('admin', '$2a$10$N9qo8uLOickgx2ZMRZoMye.IjqQ0Oj0aBl1dHUVN/k15VhYOW7EUa', 'Administrator', 'admin', 'en');

INSERT IGNORE INTO app_config (config_key, config_value, description) VALUES
('general.language', '"en"', 'Default language'),
('general.dateFormat', '"dd/MM/yyyy"', 'Date format'),
('general.theme', '"light"', 'UI theme (light/dark)'),
('stats.defaultTestTypes', '["P_FINAL","P_PROD"]', 'TEST types for sigma calculations'),
('stats.refreshInterval', '300', 'Auto-refresh interval (seconds)'),
('export.companyName', '"Hanna Instruments"', 'Company name for reports');
