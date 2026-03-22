-- ============================================
-- 005: BUFFER PRODUCTION + CORRECTION MATERIALS
-- ============================================
CREATE TABLE IF NOT EXISTS buffer_production (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    recipe_code     VARCHAR(20),
    ph_value        DECIMAL(5,2) NOT NULL,
    product_codes   TEXT,
    lot_number      VARCHAR(20) NOT NULL,
    production_date DATE,
    quantity_kg     DECIMAL(10,2),
    first_qc_failed VARCHAR(100),
    cm_description  VARCHAR(255),
    cm_code         VARCHAR(20),
    cm_grams        DECIMAL(10,4),
    cm_percentage   DECIMAL(12,8),
    source_filename VARCHAR(255),
    imported_at     DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_buf_date ON buffer_production(production_date);
CREATE INDEX idx_buf_ph ON buffer_production(ph_value);

CREATE TABLE IF NOT EXISTS correction_materials (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    cm_code         VARCHAR(20) NOT NULL UNIQUE,
    cm_name         VARCHAR(255),
    cas_number      VARCHAR(30),
    used_in_ph      TEXT
);
