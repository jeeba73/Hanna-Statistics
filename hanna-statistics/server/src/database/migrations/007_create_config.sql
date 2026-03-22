-- ============================================
-- 007: CONFIG + IMPORT LOG
-- ============================================
CREATE TABLE IF NOT EXISTS app_config (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    config_key      VARCHAR(100) UNIQUE NOT NULL,
    config_value    JSON NOT NULL,
    description     VARCHAR(255),
    updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS import_log (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    filename        VARCHAR(255) NOT NULL,
    file_type       ENUM('xlsx','csv','json') NOT NULL,
    module          ENUM('REAGENTI_QC','BUFFER_PRODUCTION','PREPARATION_LIST','OTHER'),
    rows_imported   INT DEFAULT 0,
    rows_skipped    INT DEFAULT 0,
    rows_errors     INT DEFAULT 0,
    status          ENUM('pending','processing','completed','failed') DEFAULT 'pending',
    error_details   JSON,
    user_id         INT,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    completed_at    DATETIME,
    FOREIGN KEY (user_id) REFERENCES users(id)
);
