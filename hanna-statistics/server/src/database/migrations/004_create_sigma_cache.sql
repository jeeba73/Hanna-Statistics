-- ============================================
-- 004: SIGMA CACHE, RUNNING AVERAGES, CONTROL CHART LIMITS
-- ============================================
CREATE TABLE IF NOT EXISTS lot_sigma_distribution (
    id                  INT AUTO_INCREMENT PRIMARY KEY,
    lot_id              INT NOT NULL,
    std_number          TINYINT NOT NULL,
    total_tests         INT DEFAULT 0,
    count_within_1sigma INT DEFAULT 0,
    pct_within_1sigma   DECIMAL(6,2),
    count_1to2_sigma    INT DEFAULT 0,
    pct_1to2_sigma      DECIMAL(6,2),
    count_2to3_sigma    INT DEFAULT 0,
    pct_2to3_sigma      DECIMAL(6,2),
    count_beyond_3sigma INT DEFAULT 0,
    pct_beyond_3sigma   DECIMAL(6,2),
    calculated_at       DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (lot_id) REFERENCES production_lots(id) ON DELETE CASCADE,
    UNIQUE KEY uk_lsd (lot_id, std_number)
);

CREATE TABLE IF NOT EXISTS lot_running_averages (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    lot_id          INT NOT NULL,
    std_number      TINYINT NOT NULL,
    running_avg     DECIMAL(10,6),
    calculated_at   DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (lot_id) REFERENCES production_lots(id) ON DELETE CASCADE,
    UNIQUE KEY uk_lra (lot_id, std_number)
);

CREATE TABLE IF NOT EXISTS control_chart_limits (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    hanna_code_id   INT NOT NULL,
    std_number      TINYINT NOT NULL,
    base_value      DECIMAL(10,6),
    sigma_1_low     DECIMAL(10,6),
    sigma_1_high    DECIMAL(10,6),
    sigma_2_low     DECIMAL(10,6),
    sigma_2_high    DECIMAL(10,6),
    sigma_3_low     DECIMAL(10,6),
    sigma_3_high    DECIMAL(10,6),
    FOREIGN KEY (hanna_code_id) REFERENCES hanna_codes(id) ON DELETE CASCADE,
    UNIQUE KEY uk_ccl (hanna_code_id, std_number)
);
