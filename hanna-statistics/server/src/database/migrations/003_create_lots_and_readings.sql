-- ============================================
-- 003: PRODUCTION LOTS + QC READINGS
-- ============================================
CREATE TABLE IF NOT EXISTS production_lots (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    hanna_code_id   INT NOT NULL,
    lot_number      VARCHAR(20) NOT NULL,
    lot_sequence    INT,
    preparation_week VARCHAR(20),
    first_qc_date   DATE,
    source_filename VARCHAR(255),
    imported_at     DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (hanna_code_id) REFERENCES hanna_codes(id),
    UNIQUE KEY uk_lot (hanna_code_id, lot_number)
);

CREATE INDEX idx_lot_hc ON production_lots(hanna_code_id);
CREATE INDEX idx_lot_date ON production_lots(first_qc_date);

CREATE TABLE IF NOT EXISTS qc_readings (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    lot_id          INT NOT NULL,
    std_number      TINYINT NOT NULL,
    std_value       DECIMAL(10,4),
    test_number     INT,
    test_type       ENUM('VALID','OLD_A','OLD_B','OLD_C','OLD_D','P_FINAL','P_PROD') NOT NULL,
    qc_date         DATE,
    qc_time         TIME,
    prod_date       DATE,
    prod_time       TIME,
    prod_operator   VARCHAR(50),
    head_number     TINYINT,
    meter1_reading  DECIMAL(10,4),
    meter2_reading  DECIMAL(10,4),
    meter3_reading  DECIMAL(10,4),
    meter4_reading  DECIMAL(10,4),
    spectr_abs      DECIMAL(10,6),
    ph1             DECIMAL(6,3),
    ph2             DECIMAL(6,3),
    ph3             DECIMAL(6,3),
    turbidity       DECIMAL(10,4),
    weight_mg       DECIMAL(10,2),
    reagent_set     TINYINT,
    qc_operator     VARCHAR(50),
    correction      VARCHAR(100),
    note            TEXT,
    FOREIGN KEY (lot_id) REFERENCES production_lots(id) ON DELETE CASCADE
);

CREATE INDEX idx_qc_lot ON qc_readings(lot_id);
CREATE INDEX idx_qc_date ON qc_readings(qc_date);
CREATE INDEX idx_qc_type ON qc_readings(test_type);
