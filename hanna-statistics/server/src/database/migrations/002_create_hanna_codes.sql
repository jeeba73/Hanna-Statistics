-- ============================================
-- 002: HANNA CODES (Product registry)
-- ============================================
CREATE TABLE IF NOT EXISTS hanna_codes (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    sfg_code        VARCHAR(30) NOT NULL UNIQUE,
    description     VARCHAR(255),
    parameter_formula VARCHAR(50),
    recipe          VARCHAR(50),
    production_line VARCHAR(50),
    qc_department   VARCHAR(50),
    registration_book VARCHAR(50),
    qc_type         VARCHAR(50),
    product_type    ENUM('REAGENT','BUFFER','OTHER') NOT NULL DEFAULT 'REAGENT',
    ref_weight_mg   DECIMAL(10,2),
    ref_weight_min_mg DECIMAL(10,2),
    ref_weight_max_mg DECIMAL(10,2),
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS product_standards (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    hanna_code_id   INT NOT NULL,
    std_number      TINYINT NOT NULL,
    std_value       DECIMAL(10,4) NOT NULL,
    sigma_value     DECIMAL(10,6),
    tolerance_fixed DECIMAL(10,4),
    tolerance_operator ENUM('AND','OR'),
    tolerance_percent DECIMAL(10,4),
    qc_restriction  VARCHAR(50),
    ph_value        DECIMAL(6,3),
    ph_min          DECIMAL(6,3),
    ph_max          DECIMAL(6,3),
    FOREIGN KEY (hanna_code_id) REFERENCES hanna_codes(id) ON DELETE CASCADE,
    UNIQUE KEY uk_hc_std (hanna_code_id, std_number)
);
