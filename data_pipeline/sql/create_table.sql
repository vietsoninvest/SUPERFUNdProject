DROP TABLE IF EXISTS fund_data CASCADE;
DROP TABLE IF EXISTS data_load_log CASCADE;

-- Create main fund_data table
CREATE TABLE fund_data (
    id BIGSERIAL PRIMARY KEY,
    effective_date DATE,
    fund_name VARCHAR(255) NOT NULL,
    option_name VARCHAR(255),
    asset_class_name VARCHAR(100),
    int_ext VARCHAR(20), -- Internal/External
    investment_item_name TEXT,
    currency VARCHAR(10),
    stock_id VARCHAR(100),
    listed_country VARCHAR(100),
    units_held NUMERIC(20,4),
    ownership_percentage NUMERIC(8,4),
    address TEXT,
    value_aud NUMERIC(20,2), 
    weighting NUMERIC(8,4),
    
    -- Audit columns
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);


-- Create indexes for performance optimization
CREATE INDEX idx_fund_data_effective_date ON fund_data(effective_date);
CREATE INDEX idx_fund_data_fund_name ON fund_data(fund_name);
CREATE INDEX idx_fund_data_asset_class ON fund_data(asset_class_name);

-- Create audit log table
CREATE TABLE data_load_log (
    id BIGSERIAL PRIMARY KEY,
    load_timestamp TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    file_name VARCHAR(255) NOT NULL,
    rows_loaded INTEGER DEFAULT 0,
    rows_failed INTEGER DEFAULT 0,
    status VARCHAR(20) DEFAULT 'RUNNING', -- 'RUNNING', 'SUCCESS', 'FAILED', 'PARTIAL'
    error_message TEXT,
    load_duration_seconds INTEGER,
    user_name VARCHAR(100) DEFAULT CURRENT_USER
);

-- Create a summary view for easy analysis
CREATE VIEW v_fund_summary AS
SELECT 
    fund_name,
    asset_class_name,
    COUNT(*) as total_investments,
    SUM(value_aud) as total_value_aud,
    AVG(value_aud) as avg_value_aud,
    AVG(weighting) as avg_weighting,
    AVG(ownership_percentage) as avg_ownership_percentage,
    COUNT(DISTINCT investment_item_name) as unique_investments
FROM fund_data 
WHERE fund_name IS NOT NULL
GROUP BY fund_name, asset_class_name
ORDER BY total_value_aud DESC;

-- Add table and column comments for documentation
COMMENT ON TABLE fund_data IS 'Main table storing all data';
COMMENT ON COLUMN fund_data.effective_date IS 'Date when the portfolio is effective';
COMMENT ON COLUMN fund_data.fund_name IS 'Name of the super fund';
COMMENT ON COLUMN fund_data.option_name IS 'Investment option within the fund';
COMMENT ON COLUMN fund_data.asset_class_name IS 'Asset class category (e.g., Equity, Fixed Income, etc.)';
COMMENT ON COLUMN fund_data.int_ext IS 'Internal or External investment classification';
COMMENT ON COLUMN fund_data.investment_item_name IS 'Specific name/description of the investment';
COMMENT ON COLUMN fund_data.currency IS 'Currency code for the investment';
COMMENT ON COLUMN fund_data.stock_id IS 'Stock identifier or ticker symbol';
COMMENT ON COLUMN fund_data.listed_country IS 'Country where the investment is listed';
COMMENT ON COLUMN fund_data.units_held IS 'Number of units/shares held';
COMMENT ON COLUMN fund_data.ownership_percentage IS 'Percentage ownership (0-1)';
COMMENT ON COLUMN fund_data.value_aud IS 'Market value in Australian Dollars';
COMMENT ON COLUMN fund_data.weighting IS 'Portfolio weighting (0-1)';

COMMENT ON TABLE data_load_log IS 'Audit trail for data loading operations';
COMMENT ON VIEW v_fund_summary IS 'A summary view of key metrics';