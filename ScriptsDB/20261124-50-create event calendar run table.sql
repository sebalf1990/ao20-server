-- Plan 2026-11-24-revision-eventos-automaticos, stage 2: one row per (turno, date) of the automatic events
-- calendar. Any row for a (turno, date) means "that turno was already handled today" (opened, finished,
-- cancelled or skipped), and it survives a server restart.
-- estado: ABIERTO, TERMINADO, CANCELADO, OMITIDO. fecha is "yyyy-mm-dd" (the day the turno STARTED).
CREATE TABLE IF NOT EXISTS "event_calendar_run" (
    "turno_key" TEXT NOT NULL,
    "fecha" TEXT NOT NULL,
    "variante" TEXT NOT NULL DEFAULT '',
    "estado" TEXT NOT NULL,
    "intentos" INTEGER NOT NULL DEFAULT 0,
    "detalle" TEXT NOT NULL DEFAULT '',
    "actualizado" TEXT NOT NULL DEFAULT '',
    PRIMARY KEY ("turno_key", "fecha")
);
