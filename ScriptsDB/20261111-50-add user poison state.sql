-- Plan revision-motor-elemental D7: parametros reales de los venenos activos (Menor/Hemo/Neuro) para
-- recrearlos igual al reloguear. Texto compacto, vacio = sin foto (personajes viejos cargan como antes).
ALTER TABLE user ADD COLUMN poison_state TEXT NOT NULL DEFAULT '';
