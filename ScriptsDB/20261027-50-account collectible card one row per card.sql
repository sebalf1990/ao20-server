-- Plan 24.001 (D2): one row per card, so two characters of the same account can add cards without
-- overwriting each other (the old table stored the whole collection as a 1024-byte blob per account).
-- Old table "account_collectible_cards" is kept (read-only, no longer written) until the plan closes.
CREATE TABLE IF NOT EXISTS "account_collectible_card" (
    "account_id" INTEGER NOT NULL,
    "card_index" INTEGER NOT NULL,
    "quantity" INTEGER NOT NULL,
    PRIMARY KEY ("account_id", "card_index"),
    FOREIGN KEY (account_id) REFERENCES account(id) ON DELETE CASCADE ON UPDATE CASCADE
);

-- Blob byte k (0-based) holds the quantity of card k+1 (SaveUserAccountCollectibleCards wrote
-- QuantityData(i - 1) = quantities(i), i = 1..1024). hex() gives two hex digits per byte, so byte k
-- is substr(hex(blob), 2*k+1, 2) = substr(hex(blob), 2*n-1, 2) for card n. Pure SQL, no extension functions.
WITH RECURSIVE n(i) AS (
    SELECT 1 UNION ALL SELECT i + 1 FROM n WHERE i < 1024
),
b AS (
    SELECT c.account_id AS account_id, n.i AS card_index, hex(c.card_quantities) AS h
    FROM account_collectible_cards c, n
    WHERE length(c.card_quantities) >= n.i
),
q AS (
    SELECT account_id, card_index,
           (instr('0123456789ABCDEF', substr(h, 2 * card_index - 1, 1)) - 1) * 16
           + (instr('0123456789ABCDEF', substr(h, 2 * card_index, 1)) - 1) AS quantity
    FROM b
)
INSERT OR IGNORE INTO account_collectible_card (account_id, card_index, quantity)
SELECT account_id, card_index, quantity FROM q WHERE quantity > 0;
