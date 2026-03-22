-- ============================================
-- 006: PREPARATION BATCHES
-- ============================================
CREATE TABLE IF NOT EXISTS preparation_batches (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    recipe_code     VARCHAR(30) NOT NULL,
    batch_type      ENUM('CP','SOL') NOT NULL,
    description     VARCHAR(255),
    production_line VARCHAR(20),
    revision        DECIMAL(4,2),
    expiry_years    VARCHAR(10),
    density         DECIMAL(8,4),
    preparation_date DATE,
    batch_number    TINYINT,
    planned_week    VARCHAR(10),
    actual_week     VARCHAR(10),
    planning_reference VARCHAR(50),
    operator        VARCHAR(50),
    exp_date        VARCHAR(10),
    mix_lot_number  INT,
    source_filename VARCHAR(255),
    imported_at     DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_prep_recipe ON preparation_batches(recipe_code);
CREATE INDEX idx_prep_date ON preparation_batches(preparation_date);
CREATE INDEX idx_prep_week ON preparation_batches(actual_week);

CREATE TABLE IF NOT EXISTS preparation_hanna_codes (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    batch_id        INT NOT NULL,
    hanna_code      VARCHAR(30) NOT NULL,
    product_name    VARCHAR(255),
    volume_weight   DECIMAL(10,2),
    unit            VARCHAR(10),
    qty_to_produce  INT,
    lot_number      INT,
    FOREIGN KEY (batch_id) REFERENCES preparation_batches(id) ON DELETE CASCADE
);
