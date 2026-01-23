from __future__ import annotations

import sys
from typing import Optional

import pyodbc

from data.siesa.siesa_uno_client import SiesaConfig, SiesaUNOClient

PROD_DB = "UNOEE"
TEST_DB = "PRUEBAWESTON"

TEMPLATE_OPM = 50759
ITEM_ID = 9044  # 0009044
COMP_REF = "129006"  # TRAPOS
DEFAULT_REF = "CALVO"
DEFAULT_USER = "scaicedo"


def _as_int(value):
    return int(value) if value is not None else None


def _client(db_name: str) -> SiesaUNOClient:
    cfg = SiesaConfig(auth="credman", cred_target="CalvoSiesaUNOEE", cred_user="sa", database=db_name)
    return SiesaUNOClient(cfg)


def _item_id_from_ref(test_client: SiesaUNOClient, ref: str) -> int:
    sql = "SELECT TOP 1 f120_id FROM dbo.t120_mc_items WHERE f120_referencia = ?;"
    df = test_client.fetch_df(sql, params=[ref])
    if df.empty:
        raise RuntimeError(f"No se encontro item con referencia {ref} en {TEST_DB}.")
    return int(df.iloc[0]["f120_id"])


def _fetch_template():
    prod_client = _client(PROD_DB)

    header_sql = """
    SELECT f850_id_cia, f850_id_co, f850_id_tipo_docto, f850_id_fecha,
           f850_id_grupo_clase_docto, f850_id_clase_docto, f850_id_clase_op,
           f850_ind_tipo_op, f850_ind_multiples_items, f850_ind_consolida_comp_oper,
           f850_ind_requiere_lm, f850_ind_genera_ordenes_comp, f850_ind_genera_misma_orden,
           f850_ind_genera_todos_niveles, f850_ind_genera_solo_faltantes, f850_ind_metodo_lista_op,
           f850_ind_controla_tep, f850_ind_genera_consumos_tep, f850_ind_genera_entradas_tep,
           f850_id_clase_op_generar, f850_ind_confirmar_al_aprobar, f850_ind_distribucion_costos,
           f850_ind_devolucion_comp, f850_ind_estado, f850_fecha_cumplida,
           f850_rowid_tercero_planif, f850_id_instalacion, f850_rowid_op_padre,
           f850_ind_transmitido, f850_ind_impresion, f850_nro_impresiones,
           f850_usuario_creacion,
           f850_ind_posdeduccion, f850_ind_pedido_venta, f850_rowid_pv_docto,
           f850_ind_posdeduccion_tep, f850_ind_reg_incons_posd, f850_ind_lote_automatico,
           f850_ind_controlar_cant_ext, f850_ind_incluir_operacion, f850_ind_entrega_estandar,
           f850_ind_valida_consumo_tot, f850_ind_liq_tep_estandar, f850_ind_no_liq_tep
    FROM dbo.t850_mf_op_docto
    WHERE f850_id_tipo_docto='OPM' AND f850_consec_docto = ?;
    """

    item_sql = """
    SELECT f851_id_instalacion, f851_id_fecha, f851_ind_estado, f851_fecha_cumplida,
           f851_ind_automatico, f851_id_metodo_lista_mater, f851_id_metodo_ruta,
           f851_fecha_terminacion, f851_fecha_inicio, f851_id_tipo_inv_serv,
           f851_ind_tipo_item, f851_id_unidad_medida, f851_factor,
           f851_cant_planeada_base, f851_cant_ordenada_base, f851_cant_completa_base,
           f851_cant_desechos_base, f851_cant_rechazos_base,
           f851_cant_planeada_1, f851_cant_ordenada_1, f851_cant_completa_1,
           f851_cant_desechos_1, f851_cant_rechazos_1,
           f851_cant_parcial_base, f851_ind_controla_secuencia, f851_porc_rendimiento,
           f851_notas, f851_id_lote
    FROM dbo.t851_mf_op_docto_item
    WHERE f851_rowid_op_docto IN (
        SELECT f850_rowid
        FROM dbo.t850_mf_op_docto
        WHERE f850_id_tipo_docto='OPM' AND f850_consec_docto=?
    );
    """

    item_ids_sql = """
    SELECT TOP 1 b.f150_id AS bodega_id, r.f808_id AS ruta_id
    FROM dbo.t851_mf_op_docto_item i
    LEFT JOIN dbo.t150_mc_bodegas b ON i.f851_rowid_bodega = b.f150_rowid
    LEFT JOIN dbo.t808_mf_rutas r ON i.f851_rowid_ruta = r.f808_rowid
    WHERE i.f851_rowid_op_docto IN (
        SELECT f850_rowid
        FROM dbo.t850_mf_op_docto
        WHERE f850_id_tipo_docto='OPM' AND f850_consec_docto=?
    );
    """

    comp_sql = """
    SELECT TOP 1 c.f860_numero_operacion, c.f860_rowid_ctrabajo,
           b.f150_id AS bodega_id, c.f860_id_instalacion, c.f860_id_unidad_medida,
           c.f860_ind_manual, c.f860_factor,
           c.f860_cant_requerida_base, c.f860_cant_requerida_1, c.f860_cant_requerida_2,
           c.f860_cant_desperdicio_base, c.f860_fecha_requerida, c.f860_notas,
           c.f860_rowid_item_ext_sustituido, c.f860_codigo_sustitucion,
           c.f860_cant_equiv_sustitucion, c.f860_rowid_movto_entidad, c.f860_ind_cambio_cantidad
    FROM dbo.t860_mf_op_componentes c
    JOIN dbo.t121_mc_items_extensiones e ON c.f860_rowid_item_ext_componente = e.f121_rowid
    JOIN dbo.t120_mc_items it ON e.f121_rowid_item = it.f120_rowid
    LEFT JOIN dbo.t150_mc_bodegas b ON c.f860_rowid_bodega = b.f150_rowid
    WHERE c.f860_rowid_op_docto IN (
        SELECT f850_rowid
        FROM dbo.t850_mf_op_docto
        WHERE f850_id_tipo_docto='OPM' AND f850_consec_docto=?
    )
    AND it.f120_referencia = ?;
    """

    header_df = prod_client.fetch_df(header_sql, params=[TEMPLATE_OPM])
    item_df = prod_client.fetch_df(item_sql, params=[TEMPLATE_OPM])
    item_ids_df = prod_client.fetch_df(item_ids_sql, params=[TEMPLATE_OPM])
    comp_df = prod_client.fetch_df(comp_sql, params=[TEMPLATE_OPM, COMP_REF])

    if header_df.empty or item_df.empty or item_ids_df.empty or comp_df.empty:
        raise RuntimeError("No se pudo leer la plantilla desde UNOEE.")

    return header_df.iloc[0], item_df.iloc[0], item_ids_df.iloc[0], comp_df.iloc[0]


def _resolve_rowids(
    test_client: SiesaUNOClient,
    user: str,
    item_bodega_id: int,
    comp_bodega_id: int,
    ruta_id: int,
) -> dict:
    resolve_sql = """
    SELECT
      (SELECT TOP 1 f552_rowid FROM dbo.t552_ss_usuarios WHERE UPPER(f552_nombre)=UPPER(?)) AS rowid_usuario,
      (SELECT TOP 1 f121_rowid FROM dbo.t121_mc_items_extensiones e JOIN dbo.t120_mc_items i ON e.f121_rowid_item = i.f120_rowid WHERE i.f120_id = ?) AS rowid_item_ext_padre,
      (SELECT TOP 1 f121_rowid FROM dbo.t121_mc_items_extensiones e JOIN dbo.t120_mc_items i ON e.f121_rowid_item = i.f120_rowid WHERE i.f120_id = ?) AS rowid_item_ext_comp,
      (SELECT TOP 1 f150_rowid FROM dbo.t150_mc_bodegas WHERE f150_id = ?) AS rowid_bodega_item,
      (SELECT TOP 1 f150_rowid FROM dbo.t150_mc_bodegas WHERE f150_id = ?) AS rowid_bodega_comp,
      (SELECT TOP 1 f808_rowid FROM dbo.t808_mf_rutas WHERE f808_id = ?) AS rowid_ruta;
    """

    row = test_client.fetch_df(
        resolve_sql,
        params=[
            user,
            ITEM_ID,
            _item_id_from_ref(test_client, COMP_REF),
            item_bodega_id,
            comp_bodega_id,
            ruta_id,
        ],
    ).iloc[0]

    if row.isnull().any():
        raise RuntimeError(f"No se pudo resolver rowids en {TEST_DB}: {row.to_dict()}")

    return {
        "rowid_usuario": int(row["rowid_usuario"]),
        "rowid_item_ext_padre": int(row["rowid_item_ext_padre"]),
        "rowid_item_ext_comp": int(row["rowid_item_ext_comp"]),
        "rowid_bodega_item": int(row["rowid_bodega_item"]),
        "rowid_bodega_comp": int(row["rowid_bodega_comp"]),
        "rowid_ruta": int(row["rowid_ruta"]),
    }


def _next_consec(cur) -> int:
    cur.execute(
        """
        SELECT MAX(f850_consec_docto)
        FROM dbo.t850_mf_op_docto
        WHERE f850_id_tipo_docto='OPM' AND f850_id_co='001'
          AND f850_consec_docto BETWEEN 49000 AND 60000;
        """
    )
    row = cur.fetchone()
    base = row[0] if row and row[0] is not None else None

    if base is None:
        cur.execute(
            """
            SELECT f022_cons_proximo
            FROM dbo.t022_mm_consecutivos
            WHERE f022_id_cia=1 AND f022_id_co='001' AND f022_id_tipo_docto='OPM';
            """
        )
        row = cur.fetchone()
        base = row[0] if row and row[0] else None

    if base is None:
        raise RuntimeError("No se pudo determinar el consecutivo.")

    next_consec = int(base) + 1
    while True:
        cur.execute(
            """
            SELECT COUNT(*)
            FROM dbo.t850_mf_op_docto
            WHERE f850_id_tipo_docto='OPM' AND f850_id_co='001' AND f850_consec_docto=?;
            """,
            next_consec,
        )
        if cur.fetchone()[0] == 0:
            break
        next_consec += 1

    return next_consec


def create_opm(ref_value: str = DEFAULT_REF, note_value: Optional[str] = None, user: str = DEFAULT_USER) -> None:
    if note_value is None:
        note_value = ref_value

    header, item, item_ids, comp = _fetch_template()
    test_client = _client(TEST_DB)

    with pyodbc.connect(test_client._conn_str()) as conn:
        conn.autocommit = False
        cur = conn.cursor()
        cur.execute("SELECT DB_NAME()")
        db_name = cur.fetchone()[0]
        if db_name.upper() != TEST_DB.upper():
            raise RuntimeError(f"DB inesperada: {db_name}")

        next_consec = _next_consec(cur)
        rowids = _resolve_rowids(test_client, user, int(item_ids["bodega_id"]), int(comp["bodega_id"]), int(item_ids["ruta_id"]))

        header_sql = """
        DECLARE @p_retorno smallint, @p_ts datetime, @p_rowid int;
        EXEC dbo.sp_mf_op_docto_eventos
            @p_retorno=@p_retorno OUTPUT,
            @p_ts=@p_ts OUTPUT,
            @p_rowid=@p_rowid OUTPUT,
            @p_opcion=?,
            @p_id_cia=?,
            @p_id_co=?,
            @p_id_tipo_docto=?,
            @p_consec_docto=?,
            @p_id_fecha=?,
            @p_id_grupo_clase_docto=?,
            @p_id_clase_docto=?,
            @p_id_clase_op=?,
            @p_ind_tipo_op=?,
            @p_ind_multiples_items=?,
            @p_ind_consolida_comp_oper=?,
            @p_ind_requiere_lm=?,
            @p_ind_genera_ordenes_comp=?,
            @p_ind_genera_misma_orden=?,
            @p_ind_genera_todos_niveles=?,
            @p_ind_genera_solo_faltantes=?,
            @p_ind_metodo_lista_op=?,
            @p_ind_controla_tep=?,
            @p_ind_genera_consumos_tep=?,
            @p_ind_genera_entradas_tep=?,
            @p_id_clase_op_generar=?,
            @p_ind_confirmar_al_aprobar=?,
            @p_ind_distribucion_costos=?,
            @p_ind_devolucion_comp=?,
            @p_ind_estado=?,
            @p_fecha_cumplida=?,
            @p_rowid_tercero_planif=?,
            @p_id_instalacion=?,
            @p_rowid_op_padre=?,
            @p_referencia_1=?,
            @p_referencia_2=?,
            @p_referencia_3=?,
            @p_ind_transmitido=?,
            @p_ind_impresion=?,
            @p_nro_impresiones=?,
            @p_usuario=?,
            @p_rowid_usuario=?,
            @p_notas=?,
            @p_ind_posdeduccion=?,
            @p_ind_pedido_venta=?,
            @p_rowid_pv_docto=?,
            @p_ind_posdeduccion_tep=?,
            @p_ind_reg_incons_posd=?,
            @p_ind_lote_automatico=?,
            @p_ind_controlar_cant_ext=?,
            @p_ind_incluir_operacion=?,
            @p_ind_entrega_estandar=?,
            @p_ind_valida_consumo_tot=?,
            @p_ind_liq_tep_estandar=?,
            @p_ind_no_liq_tep=?;
        SELECT @p_retorno AS retorno, @p_ts AS ts, @p_rowid AS rowid;
        """

        header_params = [
            0,
            int(header["f850_id_cia"]),
            header["f850_id_co"],
            header["f850_id_tipo_docto"],
            next_consec,
            header["f850_id_fecha"],
            int(header["f850_id_grupo_clase_docto"]),
            int(header["f850_id_clase_docto"]),
            header["f850_id_clase_op"],
            int(header["f850_ind_tipo_op"]),
            int(header["f850_ind_multiples_items"]),
            int(header["f850_ind_consolida_comp_oper"]),
            int(header["f850_ind_requiere_lm"]),
            int(header["f850_ind_genera_ordenes_comp"]),
            int(header["f850_ind_genera_misma_orden"]),
            int(header["f850_ind_genera_todos_niveles"]),
            int(header["f850_ind_genera_solo_faltantes"]),
            int(header["f850_ind_metodo_lista_op"]),
            int(header["f850_ind_controla_tep"]),
            int(header["f850_ind_genera_consumos_tep"]),
            int(header["f850_ind_genera_entradas_tep"]),
            header["f850_id_clase_op_generar"],
            int(header["f850_ind_confirmar_al_aprobar"]),
            int(header["f850_ind_distribucion_costos"]),
            int(header["f850_ind_devolucion_comp"]),
            int(header["f850_ind_estado"]),
            header["f850_fecha_cumplida"],
            int(header["f850_rowid_tercero_planif"]),
            header["f850_id_instalacion"],
            header["f850_rowid_op_padre"],
            ref_value,
            ref_value,
            ref_value,
            int(header["f850_ind_transmitido"]),
            int(header["f850_ind_impresion"]),
            int(header["f850_nro_impresiones"]),
            user,
            rowids["rowid_usuario"],
            note_value,
            int(header["f850_ind_posdeduccion"]),
            int(header["f850_ind_pedido_venta"]),
            header["f850_rowid_pv_docto"],
            int(header["f850_ind_posdeduccion_tep"]),
            int(header["f850_ind_reg_incons_posd"]),
            int(header["f850_ind_lote_automatico"]),
            int(header["f850_ind_controlar_cant_ext"]),
            int(header["f850_ind_incluir_operacion"]),
            int(header["f850_ind_entrega_estandar"]),
            int(header["f850_ind_valida_consumo_tot"]),
            int(header["f850_ind_liq_tep_estandar"]),
            int(header["f850_ind_no_liq_tep"]),
        ]

        cur.execute(header_sql, header_params)
        retorno, _, rowid_op = cur.fetchone()
        if retorno not in (0, None):
            conn.rollback()
            raise RuntimeError(f"Error creando encabezado: retorno={retorno}")

        item_sql = """
        DECLARE @p_ts datetime, @p_rowid int;
        EXEC dbo.sp_mf_op_movto_eventos
            @p_ts=@p_ts OUTPUT,
            @p_rowid=@p_rowid OUTPUT,
            @p_opcion=?,
            @p_id_cia=?,
            @p_rowid_op_docto=?,
            @p_rowid_item_ext_padre=?,
            @p_rowid_bodega=?,
            @p_id_instalacion=?,
            @p_id_fecha=?,
            @p_ind_estado=?,
            @p_fecha_cumplida=?,
            @p_ind_automatico=?,
            @p_id_metodo_lista_mater=?,
            @p_rowid_ruta=?,
            @p_id_metodo_ruta=?,
            @p_fecha_terminacion=?,
            @p_fecha_inicio=?,
            @p_id_tipo_inv_serv=?,
            @p_ind_tipo_item=?,
            @p_id_unidad_medida=?,
            @p_factor=?,
            @p_cant_planeada_base=?,
            @p_cant_ordenada_base=?,
            @p_cant_completa_base=?,
            @p_cant_desechos_base=?,
            @p_cant_rechazos_base=?,
            @p_cant_planeada_1=?,
            @p_cant_ordenada_1=?,
            @p_cant_completa_1=?,
            @p_cant_desechos_1=?,
            @p_cant_rechazos_1=?,
            @p_cant_parcial_base=?,
            @p_ind_controla_secuencia=?,
            @p_porc_rendimiento=?,
            @p_rowid_bodega_componentes=?,
            @p_notas=?,
            @p_diferente=?,
            @p_ind_condicionar=?,
            @p_estado_condicionar=?,
            @p_id_lote=?,
            @p_ind_afectar_items=?,
            @p_seg_selec_item_por_nivel=?,
            @p_rowid_pv_movto=?;
        SELECT @p_rowid AS rowid;
        """

        item_params = [
            0,
            int(header["f850_id_cia"]),
            rowid_op,
            rowids["rowid_item_ext_padre"],
            rowids["rowid_bodega_item"],
            item["f851_id_instalacion"],
            item["f851_id_fecha"],
            int(item["f851_ind_estado"]),
            item["f851_fecha_cumplida"],
            int(item["f851_ind_automatico"]),
            item["f851_id_metodo_lista_mater"],
            rowids["rowid_ruta"],
            item["f851_id_metodo_ruta"],
            item["f851_fecha_terminacion"],
            item["f851_fecha_inicio"],
            item["f851_id_tipo_inv_serv"],
            int(item["f851_ind_tipo_item"]),
            item["f851_id_unidad_medida"],
            float(item["f851_factor"]),
            float(item["f851_cant_planeada_base"]),
            float(item["f851_cant_ordenada_base"]),
            float(item["f851_cant_completa_base"]),
            float(item["f851_cant_desechos_base"]),
            float(item["f851_cant_rechazos_base"]),
            float(item["f851_cant_planeada_1"]),
            float(item["f851_cant_ordenada_1"]),
            float(item["f851_cant_completa_1"]),
            float(item["f851_cant_desechos_1"]),
            float(item["f851_cant_rechazos_1"]),
            float(item["f851_cant_parcial_base"]),
            int(item["f851_ind_controla_secuencia"]),
            float(item["f851_porc_rendimiento"]),
            None,
            item["f851_notas"],
            0,
            0,
            0,
            item["f851_id_lote"],
            0,
            0,
            None,
        ]

        cur.execute(item_sql, item_params)
        rowid_op_item = cur.fetchone()[0]

        comp_sql = """
        DECLARE @p_ts datetime, @p_rowid int;
        EXEC dbo.sp_mf_op_comp_eventos
            @p_ts=@p_ts OUTPUT,
            @p_rowid=@p_rowid OUTPUT,
            @p_opcion=?,
            @p_id_cia=?,
            @p_rowid_op_docto_item=?,
            @p_rowid_op_docto=?,
            @p_rowid_item_ext_padre=?,
            @p_rowid_item_ext_componente=?,
            @p_numero_operacion=?,
            @p_rowid_ctrabajo=?,
            @p_rowid_bodega=?,
            @p_id_instalacion=?,
            @p_id_unidad_medida=?,
            @p_ind_manual=?,
            @p_factor=?,
            @p_cant_requerida_base=?,
            @p_cant_comprometida_base=?,
            @p_cant_consumida_base=?,
            @p_cant_requerida_1=?,
            @p_cant_comprometida_1=?,
            @p_cant_consumida_1=?,
            @p_cant_requerida_2=?,
            @p_cant_comprometida_2=?,
            @p_cant_consumida_2=?,
            @p_cant_desperdicio_base=?,
            @p_fecha_requerida=?,
            @p_notas=?,
            @p_usuario=?,
            @p_rowid_item_ext_sustituido=?,
            @p_codigo_sustitucion=?,
            @p_cant_equiv_sustitucion=?,
            @p_permiso_costos=?,
            @p_rowid_movto_entidad=?,
            @p_ind_cambio_cantidad=?;
        SELECT @p_rowid AS rowid;
        """

        comp_params = [
            0,
            int(header["f850_id_cia"]),
            rowid_op_item,
            rowid_op,
            rowids["rowid_item_ext_padre"],
            rowids["rowid_item_ext_comp"],
            int(comp["f860_numero_operacion"]),
            comp["f860_rowid_ctrabajo"],
            rowids["rowid_bodega_comp"],
            comp["f860_id_instalacion"],
            comp["f860_id_unidad_medida"],
            int(comp["f860_ind_manual"]),
            float(comp["f860_factor"]),
            float(comp["f860_cant_requerida_base"]),
            0.0,
            0.0,
            float(comp["f860_cant_requerida_1"]),
            0.0,
            0.0,
            float(comp["f860_cant_requerida_2"]),
            0.0,
            0.0,
            float(comp["f860_cant_desperdicio_base"]),
            comp["f860_fecha_requerida"],
            comp["f860_notas"],
            user,
            comp["f860_rowid_item_ext_sustituido"],
            int(comp["f860_codigo_sustitucion"]),
            float(comp["f860_cant_equiv_sustitucion"]),
            0,
            comp["f860_rowid_movto_entidad"],
            int(comp["f860_ind_cambio_cantidad"]),
        ]

        cur.execute(comp_sql, comp_params)
        _ = cur.fetchone()

        conn.commit()
        print(f"CREATED: OPM {next_consec} en {db_name}")


def annul_opm(opm: int, user: str = DEFAULT_USER, force_zero_consumption: bool = False) -> None:
    test_client = _client(TEST_DB)
    with pyodbc.connect(test_client._conn_str()) as conn:
        cur = conn.cursor()
        cur.execute("SELECT DB_NAME()")
        db_name = cur.fetchone()[0]
        if db_name.upper() != TEST_DB.upper():
            raise RuntimeError(f"DB inesperada: {db_name}")

        cur.execute(
            """
            SELECT f850_rowid
            FROM dbo.t850_mf_op_docto
            WHERE f850_id_tipo_docto='OPM' AND f850_consec_docto=?;
            """,
            opm,
        )
        row = cur.fetchone()
        if not row:
            print(f"NO_OPM: {opm}")
            return
        rowid_op = int(row[0])

        if force_zero_consumption:
            cur.execute(
                """
                SELECT f860_rowid, f860_id_cia, f860_rowid_op_docto_item, f860_rowid_op_docto,
                       f860_rowid_item_ext_padre, f860_rowid_item_ext_componente,
                       f860_numero_operacion, f860_rowid_ctrabajo, f860_rowid_bodega,
                       f860_id_instalacion, f860_id_unidad_medida, f860_ind_manual, f860_factor,
                       f860_cant_requerida_base, f860_cant_comprometida_base, f860_cant_consumida_base,
                       f860_cant_requerida_1, f860_cant_comprometida_1, f860_cant_consumida_1,
                       f860_cant_requerida_2, f860_cant_comprometida_2, f860_cant_consumida_2,
                       f860_cant_desperdicio_base, f860_fecha_requerida, f860_notas,
                       f860_rowid_item_ext_sustituido, f860_codigo_sustitucion, f860_cant_equiv_sustitucion,
                       f860_rowid_movto_entidad, f860_ind_cambio_cantidad
                FROM dbo.t860_mf_op_componentes
                WHERE f860_rowid_op_docto=?;
                """,
                rowid_op,
            )
            cols = [d[0] for d in cur.description]
            comps = [dict(zip(cols, row)) for row in cur.fetchall()]

            for comp in comps:
                update_sql = """
                DECLARE @p_ts datetime, @p_rowid int;
                SET @p_rowid = ?;
                EXEC dbo.sp_mf_op_comp_eventos
                    @p_ts=@p_ts OUTPUT,
                    @p_rowid=@p_rowid OUTPUT,
                    @p_opcion=?,
                    @p_id_cia=?,
                    @p_rowid_op_docto_item=?,
                    @p_rowid_op_docto=?,
                    @p_rowid_item_ext_padre=?,
                    @p_rowid_item_ext_componente=?,
                    @p_numero_operacion=?,
                    @p_rowid_ctrabajo=?,
                    @p_rowid_bodega=?,
                    @p_id_instalacion=?,
                    @p_id_unidad_medida=?,
                    @p_ind_manual=?,
                    @p_factor=?,
                    @p_cant_requerida_base=?,
                    @p_cant_comprometida_base=?,
                    @p_cant_consumida_base=?,
                    @p_cant_requerida_1=?,
                    @p_cant_comprometida_1=?,
                    @p_cant_consumida_1=?,
                    @p_cant_requerida_2=?,
                    @p_cant_comprometida_2=?,
                    @p_cant_consumida_2=?,
                    @p_cant_desperdicio_base=?,
                    @p_fecha_requerida=?,
                    @p_notas=?,
                    @p_usuario=?,
                    @p_rowid_item_ext_sustituido=?,
                    @p_codigo_sustitucion=?,
                    @p_cant_equiv_sustitucion=?,
                    @p_permiso_costos=?,
                    @p_rowid_movto_entidad=?,
                    @p_ind_cambio_cantidad=?;
                """
                cur.execute(
                    update_sql,
                    _as_int(comp["f860_rowid"]),
                    1,
                    _as_int(comp["f860_id_cia"]),
                    _as_int(comp["f860_rowid_op_docto_item"]),
                    _as_int(comp["f860_rowid_op_docto"]),
                    _as_int(comp["f860_rowid_item_ext_padre"]),
                    _as_int(comp["f860_rowid_item_ext_componente"]),
                    _as_int(comp["f860_numero_operacion"]),
                    _as_int(comp["f860_rowid_ctrabajo"]),
                    _as_int(comp["f860_rowid_bodega"]),
                    comp["f860_id_instalacion"],
                    comp["f860_id_unidad_medida"],
                    _as_int(comp["f860_ind_manual"]),
                    comp["f860_factor"],
                    comp["f860_cant_requerida_base"],
                    comp["f860_cant_comprometida_base"],
                    0.0,
                    comp["f860_cant_requerida_1"],
                    comp["f860_cant_comprometida_1"],
                    0.0,
                    comp["f860_cant_requerida_2"],
                    comp["f860_cant_comprometida_2"],
                    0.0,
                    comp["f860_cant_desperdicio_base"],
                    comp["f860_fecha_requerida"],
                    comp["f860_notas"],
                    user,
                    _as_int(comp["f860_rowid_item_ext_sustituido"]),
                    _as_int(comp["f860_codigo_sustitucion"]),
                    comp["f860_cant_equiv_sustitucion"],
                    0,
                    _as_int(comp["f860_rowid_movto_entidad"]),
                    _as_int(comp["f860_ind_cambio_cantidad"]),
                )

        val_sql = """
        DECLARE @p_error int;
        EXEC dbo.sp_mf_op_validar_anular
            @p_error=@p_error OUTPUT,
            @p_rowid_op_docto=?,
            @p_usuario=?,
            @p_ind_seg_oc=?;
        SELECT @p_error AS error;
        """
        cur.execute(val_sql, rowid_op, user, 0)
        val_err = cur.fetchone()[0]
        if val_err not in (0, None):
            conn.rollback()
            raise RuntimeError(f"Validacion anular fallo: {val_err}")

        ann_sql = """
        DECLARE @p_error int;
        EXEC dbo.sp_mf_op_anular
            @p_error=@p_error OUTPUT,
            @p_rowid_op_docto=?,
            @p_usuario=?;
        SELECT @p_error AS error;
        """
        cur.execute(ann_sql, rowid_op, user)
        ann_err = cur.fetchone()[0]
        if ann_err not in (0, None):
            conn.rollback()
            raise RuntimeError(f"Anulacion fallo: {ann_err}")

        conn.commit()
        print(f"ANULADO: OPM {opm} en {db_name}")


def _usage() -> None:
    print("Usage:")
    print("  python tools\\siesa_pruebaweston_opm.py create [REF] [NOTA]")
    print("  python tools\\siesa_pruebaweston_opm.py annul <OPM> [--force-zero-consumption]")


def main() -> None:
    if len(sys.argv) < 2:
        _usage()
        return

    cmd = sys.argv[1].lower()

    if cmd == "create":
        ref = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_REF
        note = sys.argv[3] if len(sys.argv) > 3 else ref
        create_opm(ref, note, DEFAULT_USER)
        return

    if cmd == "annul":
        if len(sys.argv) < 3:
            _usage()
            return
        opm = int(sys.argv[2])
        force = "--force-zero-consumption" in sys.argv[3:]
        annul_opm(opm, DEFAULT_USER, force)
        return

    _usage()


if __name__ == "__main__":
    main()
