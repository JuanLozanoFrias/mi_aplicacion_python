-- Ordenes de produccion (base: vBI_850).
-- Usa ? para IdCia (Weston=1, WBR=5, TEKOAM=6).
SELECT DISTINCT
  b850_op_numero            AS OpNumero,
  b850_op_fecha_elaboracion AS FechaDocto,
  b850_op_desc_estado       AS Estado,
  b850_notas                AS Notas,
  b850_op_docto_referencia1 AS OpReferencia1,
  b850_op_docto_referencia2 AS OpReferencia2
FROM dbo.vBI_850
WHERE b850_op_co_id = RIGHT('000' + CAST(? AS varchar(3)), 3)
  AND b850_op_fecha_elaboracion >= DATEADD(MONTH, -2, CAST(GETDATE() AS date));
