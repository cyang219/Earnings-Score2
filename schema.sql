-- ============================================================
-- ES Updater: Supabase schema
-- Run once in the Supabase SQL Editor before first use
-- ============================================================

CREATE TABLE IF NOT EXISTS tickers (
    ticker     TEXT PRIMARY KEY,   -- e.g. "AVGO US"
    index_name TEXT,               -- e.g. "SPX"
    currency   TEXT,               -- e.g. "USD"
    sector     TEXT                -- e.g. "Semiconductors & Semiconductor"
);

CREATE TABLE IF NOT EXISTS quarters (
    id      BIGSERIAL PRIMARY KEY,
    ticker  TEXT NOT NULL REFERENCES tickers(ticker) ON DELETE CASCADE,
    quarter TEXT NOT NULL,         -- e.g. "26Q1"
    date    DATE,
    UNIQUE (ticker, quarter)
);

-- One wide row per quarter containing all fixed signals:
--   read-through, bull, bear, and the 3 common themes.
-- Each signal has an integer value (-1/0/1) and a combined description (text — rationale).
CREATE TABLE IF NOT EXISTS quarter_signals (
    quarter_id                     BIGINT PRIMARY KEY REFERENCES quarters(id) ON DELETE CASCADE,

    read_through_signal            SMALLINT,
    read_through_description       TEXT,

    bull_signal                    SMALLINT,
    bull_description               TEXT,

    bear_signal                    SMALLINT,
    bear_description               TEXT,

    -- Common theme 1: Top-line Growth
    topline_mgmt_signal            SMALLINT,
    topline_mgmt_description       TEXT,
    topline_analyst_signal         SMALLINT,
    topline_analyst_description    TEXT,

    -- Common theme 2: Bottom-line Expansion/Contraction
    bottomline_mgmt_signal         SMALLINT,
    bottomline_mgmt_description    TEXT,
    bottomline_analyst_signal      SMALLINT,
    bottomline_analyst_description TEXT,

    -- Common theme 3: Financial Guidance
    guidance_mgmt_signal           SMALLINT,
    guidance_mgmt_description      TEXT,
    guidance_analyst_signal        SMALLINT,
    guidance_analyst_description   TEXT
);

-- One row per (quarter, theme) for AI-discovered variable themes.
-- rank 1-3 reflects discussion volume ranking in that quarter.
CREATE TABLE IF NOT EXISTS variable_themes (
    id               BIGSERIAL PRIMARY KEY,
    quarter_id       BIGINT NOT NULL REFERENCES quarters(id) ON DELETE CASCADE,
    theme_name       TEXT NOT NULL,
    rank             INTEGER,
    mgmt_signal         SMALLINT,
    mgmt_description    TEXT,
    analyst_signal      SMALLINT,
    analyst_description TEXT,
    UNIQUE (quarter_id, theme_name)
);

-- One row per (ticker, quarter) for Bloomberg data.
-- All fields stored as JSONB so new Bloomberg fields require no schema change.
CREATE TABLE IF NOT EXISTS external_data (
    id      BIGSERIAL PRIMARY KEY,
    ticker  TEXT NOT NULL REFERENCES tickers(ticker) ON DELETE CASCADE,
    quarter TEXT NOT NULL,
    data    JSONB NOT NULL DEFAULT '{}',
    UNIQUE (ticker, quarter)
);
