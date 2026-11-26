-- Plan 2026-11-26-borrar-personaje: player-initiated character deletion.
-- delete_requested_at: epoch seconds when the web asked to delete the character (NULL = no pending request).
--   The server sets it (op mark_char_delete) and clears it when the owner enters the game (that cancels the deletion).
-- deleted_at: epoch seconds when the character was logically deleted (user.deleted = 1).
ALTER TABLE user ADD COLUMN delete_requested_at INTEGER NULL;
ALTER TABLE user ADD COLUMN deleted_at INTEGER NULL;
