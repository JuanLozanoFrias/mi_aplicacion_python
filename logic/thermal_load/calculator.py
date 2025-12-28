from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from math import ceil
from pathlib import Path
from typing import Dict, List, Optional, Tuple

DATA_DIR = Path("data/cuartos_industriales")


def _load_json(name: str):
    path = DATA_DIR / name
    with path.open(encoding="utf-8") as f:
        return json.load(f)


# ---------------------------- MODELOS ---------------------------- #


@dataclass
class RoomInputs:
    nombre: str = "CUARTO"
    largo_m: float = 0.0
    ancho_m: float = 0.0
    altura_m: float = 0.0
    perfil_id: str = ""
    puertas: int = 1

    # características térmicas
    T_internal_C: float = 0.0
    T_ext_front_C: float = 30.0
    T_ext_back_C: float = 30.0
    T_ext_right_C: float = 30.0
    T_ext_left_C: float = 30.0
    T_ext_roof_C: float = 30.0
    ground_temp_C: float = 13.0
    wall_transfer_factor: float = 0.96

    # aislamiento
    insulation_type: str = ""
    insulation_thickness_in: float = 4.0

    # infiltración
    outside_air_temp_C: float = 30.0
    outside_RH: float = 0.6
    inside_RH: Optional[float] = None
    air_changes_24h_override: Optional[float] = None
    use_factor: float = 2.0
    infiltration_enabled: bool = True

    # producto
    product_name: Optional[str] = None
    product_mass_kg: float = 0.0
    product_Tin_C: Optional[float] = None
    product_Tout_C: float = 0.0
    product_cycle_h: float = 24.0
    product_method: str = "enthalpy"  # enthalpy | cp_latent
    product_packaging_multiplier: float = 1.05

    # internas
    lighting_W: float = 0.0
    lighting_hours: float = 18.0
    motors_W: float = 0.0
    motors_hours: float = 24.0
    forklift_hp: float = 0.0
    forklift_hours: float = 2.0
    people_count: float = 0.0
    people_hours: float = 0.0
    people_btuh: float = 1350.0
    defrost_W: float = 0.0
    defrost_count: float = 0.0
    defrost_duration_min: float = 0.0
    defrost_fraction_to_room: float = 1.0

    # extras
    run_hours_supp: float = 20.0


@dataclass
class ComponentBreakdown:
    transmission_btuh: float = 0.0
    infiltration_btuh: float = 0.0
    infiltration_sensible_btuh: float = 0.0
    infiltration_latent_btuh: float = 0.0
    lighting_btuh: float = 0.0
    motors_btuh: float = 0.0
    forklift_btuh: float = 0.0
    people_btuh: float = 0.0
    defrost_btuh: float = 0.0
    product_btuh: float = 0.0

    @property
    def internal_btuh(self):
        return (
            self.lighting_btuh
            + self.motors_btuh
            + self.forklift_btuh
            + self.people_btuh
            + self.defrost_btuh
        )

    @property
    def total_btuh(self):
        return (
            self.transmission_btuh
            + self.infiltration_btuh
            + self.internal_btuh
            + self.product_btuh
        )

    def as_percentages(self):
        tot = self.total_btuh or 1.0
        return {
            "transmission_pct": self.transmission_btuh / tot,
            "infiltration_pct": self.infiltration_btuh / tot,
            "internal_pct": self.internal_btuh / tot,
            "product_pct": self.product_btuh / tot,
        }


@dataclass
class RoomResult:
    inputs: RoomInputs
    components: ComponentBreakdown
    total_btuh: float
    total_kw: float
    total_tr: float
    percentages: Dict[str, float] = field(default_factory=dict)
    product: Optional["ProductBreakdown"] = None
    # accesos directos opcionales
    ref_btu_cycle: Optional[float] = None
    cong_btu_cycle: Optional[float] = None
    post_btu_cycle: Optional[float] = None
    infiltration_ach_24h: Optional[float] = None
    infiltration_cfm: Optional[float] = None
    infiltration_sensible_only: bool = False
    infiltration_notes: Optional[str] = None


@dataclass
class ProductBreakdown:
    ref_btu_cycle: float = 0.0
    cong_btu_cycle: float = 0.0
    post_btu_cycle: float = 0.0
    total_btu_cycle: float = 0.0
    tf_c: float = 0.0
    water_frac: float = 0.0
    cp_above_kjkgk: float = 0.0
    cp_below_kjkgk: float = 0.0


@dataclass
class ProjectResult:
    rooms: List[RoomResult]
    total_btuh: float
    total_kw: float
    total_tr: float


# ---------------------------- CALCULADORA ---------------------------- #


class ThermalLoadCalculator:
    def __init__(self):
        self.insulation = _load_json("insulation_k_factors.json")["k_factor_by_insulation"]
        self.air_changes_tbl = _load_json("air_changes_24h_by_volume_ft3.json")["table"]
        foods_json = _load_json("foods_thermal_properties.json")
        self.foods = foods_json.get("by_name", {})
        self.default_profiles = {
            p["id"]: p for p in _load_json("thermal_load_default_profiles.json")["profiles"]
        }

    # utilidades
    def _profile_defaults(self, perfil_id: str) -> Dict:
        prof = self.default_profiles.get(perfil_id)
        if not prof:
            prof = next(iter(self.default_profiles.values()))
        return prof["defaults"]

    def _air_changes(self, volume_ft3: float, freezing: bool) -> float:
        for row in self.air_changes_tbl:
            if volume_ft3 <= row["max_volume_ft3"]:
                return row["air_changes_24h_freezing"] if freezing else row["air_changes_24h_refrigeration"]
        last = self.air_changes_tbl[-1]
        return last["air_changes_24h_freezing"] if freezing else last["air_changes_24h_refrigeration"]

    def _u_factor(self, insulation_type: str, thickness_in: float) -> float:
        k = self.insulation.get(insulation_type, 0.16)  # en unidades Excel
        if thickness_in <= 0:
            thickness_in = 1.0
        return k / thickness_in

    def _saturation_pressure_pa(self, t_C: float) -> float:
        # Tetens approximation
        from math import exp
        return 610.94 * exp((17.625 * t_C) / (t_C + 243.04))

    def _humidity_ratio(self, t_C: float, rh: float, pressure_pa: float = 101325.0) -> float:
        rh = max(0.0, min(1.0, rh))
        psat = self._saturation_pressure_pa(t_C)
        pv = rh * psat
        return 0.621945 * pv / max(1.0, pressure_pa - pv)

    def _air_enthalpy_btu_lb(self, t_C: float, w: float) -> float:
        # h (kJ/kg_da) = 1.006*T + w*(2501 + 1.86*T); convert to Btu/lb_da
        h_kjkg = 1.006 * t_C + w * (2501.0 + 1.86 * t_C)
        return h_kjkg * 0.429922614

    def _infiltration_btuh(self, room: RoomInputs, volume_ft3: float, Tint_C: float, run_hours: float, apply_use_factor: bool = True):
        """
        Devuelve (total_btuh, sensible_btuh, latent_btuh, ach_used, cfm_used, sensible_only_flag, notes).
        Fallback sensible-only si falta RH.
        """
        Tint_F = Tint_C * 9/5 + 32
        Text_C = room.outside_air_temp_C
        Text_F = Text_C * 9/5 + 32
        ach = room.air_changes_24h_override
        freezing = Tint_C < -5
        mode = "override"
        if ach is None or ach <= 0:
            mode = "auto"
            ach = self._air_changes(volume_ft3, freezing)
            doors_factor = 1.0 + (room.puertas or 0) * 0.15
            ach *= doors_factor
        cfm = (ach/24.0) * (volume_ft3/60.0)

        sensible_btuh = 1.08 * cfm * (Text_F - Tint_F)
        total_btuh = sensible_btuh
        latent_btuh = 0.0
        sensible_only = False
        notes = ""

        outside_rh = room.outside_RH
        inside_rh = room.inside_RH
        if inside_rh is None:
            inside_rh = 0.9 if Tint_C < -5 else 0.85
            notes = f"default_inside_RH={inside_rh}"

        if outside_rh is None or outside_rh <= 0:
            # log deshabilitado
            sensible_only = True
        else:
            try:
                w_out = self._humidity_ratio(Text_C, outside_rh)
                w_in = self._humidity_ratio(Tint_C, inside_rh)
                h_out = self._air_enthalpy_btu_lb(Text_C, w_out)
                h_in = self._air_enthalpy_btu_lb(Tint_C, w_in)
                mass_flow_lb_da_per_h = cfm * 60.0 * 0.075
                total_btuh = mass_flow_lb_da_per_h * (h_out - h_in)
                latent_btuh = max(0.0, total_btuh - sensible_btuh)
            except Exception:
                # log deshabilitado
                sensible_only = True
                total_btuh = sensible_btuh
                latent_btuh = 0.0

        if total_btuh < 0:
            total_btuh = 0.0
            sensible_btuh = max(0.0, sensible_btuh)
            latent_btuh = 0.0

        if apply_use_factor:
            factor = room.use_factor or 1.0
            total_btuh *= factor
            sensible_btuh *= factor
            latent_btuh *= factor

        scale_hours = (run_hours or 24.0) / 24.0
        total_btuh *= scale_hours
        sensible_btuh *= scale_hours
        latent_btuh *= scale_hours

        # log deshabilitado
        return total_btuh, sensible_btuh, latent_btuh, ach, cfm, sensible_only, notes

    def _get_product_data(self, name: Optional[str]) -> Optional[Dict]:
        if not name:
            return None
        prod = self.foods.get(name)
        if prod:
            return prod
        target = (name or "").lower()
        for key, val in self.foods.items():
            if key.lower() == target:
                return val
        # log deshabilitado
        return None

    def _compute_product_breakdown_btu_cycle(
        self,
        product: Optional[Dict],
        Tin_C: float,
        Tout_C: float,
        mass_kg: float,
        packaging_multiplier: float = 1.0,
    ) -> Tuple[float, float, float, float, float, float, float]:
        """
        Devuelve (ref, cong, post, tf_c, water_frac, cp_above_kjkgk, cp_below_kjkgk) en BTU/ciclo.
        """
        if mass_kg <= 0:
            return 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0

        pkg_mult = packaging_multiplier if packaging_multiplier and packaging_multiplier > 0 else 1.0
        mass_lb = mass_kg * pkg_mult * 2.2046226218

        tf_c = 0.0
        used_default_tf = False
        if product:
            if "freezing_point_C" in product:
                tf_c = product.get("freezing_point_C") or 0.0
            elif "freezing_point_F" in product:
                tf_c = ((product.get("freezing_point_F") or 32.0) - 32.0) * 5 / 9
            else:
                used_default_tf = True
        else:
            used_default_tf = True
        if used_default_tf:
            name = product.get("name") if isinstance(product, dict) else None
            # log deshabilitado
        composition = product.get("composition_pct") if product else {}

        def _pct(key: str, default: float = 0.0) -> float:
            try:
                val = composition.get(key, default)
                return max(0.0, float(val))
            except Exception:
                return default

        water_frac = min(_pct("moisture", 70.0) / 100.0, 1.0)
        prot = _pct("protein") / 100.0
        fat = _pct("fat") / 100.0
        carbo = _pct("carbohydrate") / 100.0
        fiber = _pct("fiber") / 100.0
        ash = _pct("ash") / 100.0
        carbs_total = carbo + fiber

        cp_above_kjkgk_formula = 4.186 * water_frac + 1.7 * prot + 1.9 * fat + 1.55 * carbs_total + 1.42 * ash
        cp_below_kjkgk_formula = 2.05 * water_frac + 1.7 * prot + 1.9 * fat + 1.55 * carbs_total + 1.42 * ash

        cp_above_btu_lbF = None
        cp_below_btu_lbF = None
        if product:
            cp_above_btu_lbF = product.get("cp_above_Btu_lbF")
            cp_below_btu_lbF = product.get("cp_below_Btu_lbF")

        if cp_above_btu_lbF is None:
            cp_above_kjkgk = cp_above_kjkgk_formula
            cp_above_btu_lbF = cp_above_kjkgk * 0.2388458966
        else:
            cp_above_kjkgk = cp_above_btu_lbF / 0.2388458966

        if cp_below_btu_lbF is None:
            cp_below_kjkgk = cp_below_kjkgk_formula
            cp_below_btu_lbF = cp_below_kjkgk * 0.2388458966
        else:
            cp_below_kjkgk = cp_below_btu_lbF / 0.2388458966

        latent_btulb = None
        if product:
            latent_btulb = product.get("latent_heat_Btu_lb")
        if latent_btulb is None:
            latent_btulb = 333.55 * 0.429922614  # kJ/kg to Btu/lb

        ref_btu = 0.0
        cong_btu = 0.0
        post_btu = 0.0

        if Tout_C >= tf_c:
            # Trayectoria sin congelar
            dT_ref_C = max(Tin_C - Tout_C, 0.0)
            if dT_ref_C > 0:
                ref_btu = mass_lb * cp_above_btu_lbF * (dT_ref_C * 1.8)
        else:
            # Se congela
            if Tin_C > tf_c:
                ref_btu = mass_lb * cp_above_btu_lbF * ((Tin_C - tf_c) * 1.8)
                cong_btu = mass_lb * latent_btulb * water_frac
                start_post_C = tf_c
            else:
                start_post_C = Tin_C
            dT_post_C = start_post_C - Tout_C
            if dT_post_C > 0:
                post_btu = mass_lb * cp_below_btu_lbF * (dT_post_C * 1.8)

        return ref_btu, cong_btu, post_btu, tf_c, water_frac, cp_above_kjkgk, cp_below_kjkgk

    def _product_btuh(self, room: RoomInputs) -> float:
        if room.product_mass_kg <= 0 or room.product_cycle_h <= 0:
            return 0.0
        prod = self._get_product_data(room.product_name)
        Tin = room.product_Tin_C if room.product_Tin_C is not None else (room.T_internal_C or 0.0)
        Tout = room.product_Tout_C if room.product_Tout_C is not None else Tin
        ref, cong, post, *_ = self._compute_product_breakdown_btu_cycle(
            prod, Tin, Tout, room.product_mass_kg, room.product_packaging_multiplier
        )
        total_cycle = ref + cong + post
        if room.product_cycle_h <= 0:
            return 0.0
        return max(0.0, total_cycle / room.product_cycle_h)
    def compute_room(self, room: RoomInputs, safety_factor: float = 1.1) -> RoomResult:
        # defaults del perfil
        defaults = self._profile_defaults(room.perfil_id)
        run_hours = room.run_hours_supp or defaults.get("run_hours_supp", 20)
        Tint_C = room.T_internal_C if room.T_internal_C != 0 else defaults.get("Tint_C", 0)
        # geometria
        L_ft = room.largo_m * 3.28084
        W_ft = room.ancho_m * 3.28084
        H_ft = room.altura_m * 3.28084
        volume_ft3 = L_ft * W_ft * H_ft
        A_front = W_ft * H_ft
        A_back = W_ft * H_ft
        A_left = L_ft * H_ft
        A_right = L_ft * H_ft
        A_roof = L_ft * W_ft
        A_floor = L_ft * W_ft

        U = self._u_factor(room.insulation_type, room.insulation_thickness_in)

        def face_btuh(area_ft2: float, text_c: float) -> float:
            dT_F = (text_c - Tint_C) * 9/5
            return U * area_ft2 * max(0.0, dT_F)

        trans_btuh = (
            face_btuh(A_front, room.T_ext_front_C)
            + face_btuh(A_back, room.T_ext_back_C)
            + face_btuh(A_left, room.T_ext_left_C)
            + face_btuh(A_right, room.T_ext_right_C)
            + face_btuh(A_roof, room.T_ext_roof_C)
            + face_btuh(A_floor, room.ground_temp_C)
        )
        trans_btuh = trans_btuh * room.wall_transfer_factor * run_hours / 24.0

        if room.infiltration_enabled is False:
            infil_btuh = infil_sensible_btuh = infil_latent_btuh = 0.0
            ach_used = 0.0
            cfm_used = 0.0
            infil_sensible_only = False
            infil_notes = "Desactivada por usuario"
            # log deshabilitado
        else:
            infil_btuh, infil_sensible_btuh, infil_latent_btuh, ach_used, cfm_used, infil_sensible_only, infil_notes = self._infiltration_btuh(room, volume_ft3, Tint_C, run_hours, apply_use_factor=True)

        # internas
        factor_use = room.use_factor or 1.0
        # Aplicar factor de uso tambien a internas en modo detallado
        lighting_btuh = room.lighting_W * 3.412 * (room.lighting_hours / 24.0) * factor_use
        motors_btuh = room.motors_W * 3.412 * (room.motors_hours / 24.0) * factor_use
        forklift_btuh = room.forklift_hp * 2544.43 * (room.forklift_hours / 24.0) * factor_use
        people_btuh = room.people_count * room.people_btuh * (room.people_hours / 24.0) * factor_use
        defrost_btuh = room.defrost_W * (room.defrost_duration_min/60.0) * room.defrost_count * (room.defrost_fraction_to_room) / 24.0 * factor_use

        # producto
        product_btuh = 0.0
        product_breakdown = None
        Tin_prod = room.product_Tin_C if room.product_Tin_C is not None else (room.T_internal_C or 0.0)
        Tout_prod = room.product_Tout_C if room.product_Tout_C is not None else Tin_prod
        if room.product_mass_kg > 0 and room.product_cycle_h > 0:
            method = (room.product_method or "cp_latent").lower()
            if method != "cp_latent":
                # log deshabilitado
                pass
            prod_data = self._get_product_data(room.product_name)
            ref_btu, cong_btu, post_btu, tf_c, water_frac, cp_above_kjkgk, cp_below_kjkgk = self._compute_product_breakdown_btu_cycle(
                prod_data, Tin_prod, Tout_prod, room.product_mass_kg, room.product_packaging_multiplier
            )
            total_cycle_btu = ref_btu + cong_btu + post_btu
            product_btuh = total_cycle_btu / room.product_cycle_h if room.product_cycle_h > 0 else 0.0
            product_breakdown = ProductBreakdown(
                ref_btu_cycle=ref_btu,
                cong_btu_cycle=cong_btu,
                post_btu_cycle=post_btu,
                total_btu_cycle=total_cycle_btu,
                tf_c=tf_c,
                water_frac=water_frac,
                cp_above_kjkgk=cp_above_kjkgk,
                cp_below_kjkgk=cp_below_kjkgk,
            )

        comp = ComponentBreakdown(
            transmission_btuh=trans_btuh,
            infiltration_btuh=infil_btuh,
            infiltration_sensible_btuh=infil_sensible_btuh,
            infiltration_latent_btuh=infil_latent_btuh,
            lighting_btuh=lighting_btuh,
            motors_btuh=motors_btuh,
            forklift_btuh=forklift_btuh,
            people_btuh=people_btuh,
            defrost_btuh=defrost_btuh,
            product_btuh=product_btuh,
        )
        total_btuh = comp.total_btuh * safety_factor
        total_kw = total_btuh / 3412.142
        total_tr = total_btuh / 12000.0
        # log deshabilitado
        return RoomResult(
            inputs=room,
            components=comp,
            total_btuh=total_btuh,
            total_kw=total_kw,
            total_tr=total_tr,
            percentages=comp.as_percentages(),
            product=product_breakdown,
            ref_btu_cycle=product_breakdown.ref_btu_cycle if product_breakdown else None,
            cong_btu_cycle=product_breakdown.cong_btu_cycle if product_breakdown else None,
            post_btu_cycle=product_breakdown.post_btu_cycle if product_breakdown else None,
            infiltration_ach_24h=ach_used,
            infiltration_cfm=cfm_used,
            infiltration_sensible_only=infil_sensible_only,
            infiltration_notes=infil_notes,
        )

    def compute_project(self, rooms: List[RoomInputs], safety_factor: float) -> ProjectResult:
        res_rooms = [self.compute_room(r, safety_factor) for r in rooms]
        tot_btuh = sum(r.total_btuh for r in res_rooms)
        return ProjectResult(
            rooms=res_rooms,
            total_btuh=tot_btuh,
            total_kw=tot_btuh / 3412.142,
            total_tr=tot_btuh / 12000.0,
        )

    # alias para compatibilidad UI
    def compute(self, rooms: List[RoomInputs], safety_factor: float) -> ProjectResult:
        return self.compute_project(rooms, safety_factor)

    # ---------------- GOLDEN CASE ---------------- #
    def run_golden_case(self) -> Tuple[RoomResult, Dict]:
        data = _load_json("golden_case_cuarto2_devoluciones.json")
        sel = data["selected"]
        rdat = sel["room"]
        op = sel["operation"]
        prod = sel["product"]
        room = RoomInputs(
            nombre="GOLDEN",
            largo_m=rdat["L_m"],
            ancho_m=rdat["W_m"],
            altura_m=rdat["H_m"],
            perfil_id="freezer_general",
            T_internal_C=rdat["T_internal_C"],
            T_ext_front_C=rdat["T_ext_front_C"],
            T_ext_back_C=rdat["T_ext_back_C"],
            T_ext_right_C=rdat["T_ext_right_C"],
            T_ext_left_C=rdat["T_ext_left_C"],
            T_ext_roof_C=rdat["T_ext_roof_C"],
            insulation_type=rdat["insulation_type"],
            insulation_thickness_in=rdat["insulation_thickness_in"],
            lighting_W=op["lighting_W"],
            motors_W=op["motors_W"],
            forklift_hp=op["forklift_hp"],
            defrost_W=op["defrost_W"],
            people_count=op["people_count"],
            people_hours=op["people_hours"],
            outside_RH=op["outside_RH"],
            air_changes_24h_override=op["air_changes_24h"],
            use_factor=op["use_factor"],
            outside_air_temp_C=(op["outside_temp_F_for_infil"] - 32) * 5/9,
            product_name=prod["name"],
            product_Tin_C=prod.get("Tin_C"),
            product_Tout_C=prod["Tout_C"],
            product_mass_kg=prod["mass_kg"],
            product_cycle_h=prod["cycle_h"],
        )
        res = self.compute_room(room, safety_factor=1.0)
        expected = sel["key_results"]
        return res, expected
