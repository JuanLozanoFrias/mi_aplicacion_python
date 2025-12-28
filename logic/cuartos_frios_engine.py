from __future__ import annotations

from dataclasses import dataclass
from math import sqrt
from pathlib import Path
from typing import Dict, List, Literal, Optional, Tuple
import json
import pandas as pd

Family = Literal["frontal", "dual"]


# --------------------------------------------------------------------------- DTOs
@dataclass(frozen=True)
class ColdRoomInputs:
    length_m: float
    width_m: float
    height_m: float
    usage: str
    n_evaporators: Optional[int] = None  # None => Auto (1..4)
    safety_factor: float | None = None
    family_override: Literal["auto", "frontal", "dual", "frontal_wef", "frontal_wefm"] | None = None


@dataclass(frozen=True)
class ColdRoomResult:
    valid: bool
    messages: List[str]
    length_ft: int | None = None
    width_ft: int | None = None
    height_ft: int | None = None
    height_bucket_ft: int | None = None
    height_factor: float | None = None
    load_btu_hr: float | None = None
    tevap_f: float | None = None
    tcam_f: float | None = None
    evap_family: Family | None = None
    evap_model: str | None = None
    evap_capacity_btu_hr: float | None = None
    load_per_evap_btu_hr: float | None = None
    fit_ok: bool | None = None
    fit_msg: str | None = None
    tiro_ok: bool | None = None
    tiro_msg: str | None = None
    n_requested: int | None = None
    n_used: int | None = None
    auto_note: str | None = None


# --------------------------------------------------------------------------- Motor
class ColdRoomEngine:
    """
    Motor de cálculo para cuartos fríos usando JSON unificado y fit-check WEF/WEFM/WED.
    """

    def __init__(self, data_path: str | Path):
        p = Path(data_path)
        if not p.exists():
            raise FileNotFoundError(f"No se encontró el archivo de datos: {p}")

        if p.suffix.lower() == ".json":
            blob = json.loads(p.read_text(encoding="utf-8"))
            self.config = blob["config"]
            self.usage_profiles = blob["usage_profiles"]
            self.capacity_tables = {k: pd.DataFrame(v) for k, v in blob["capacity_tables"].items()}
            for df in self.capacity_tables.values():
                df.index = df.index.astype(int)
                df.columns = [int(c) for c in df.columns]
            self.evap_dual = pd.DataFrame(blob["evaporadores_dual"])
            self.evap_frontal = pd.DataFrame(blob["evaporadores_frontal"])
            for df in (self.evap_dual, self.evap_frontal):
                df.index = df.index.astype(str)
                df.columns = [int(c) for c in df.columns]
        else:
            base = p
            self.config = json.loads((base / "config.json").read_text(encoding="utf-8"))
            self.usage_profiles = json.loads((base / "usage_profiles.json").read_text(encoding="utf-8"))
            self.capacity_tables = {}
            for name in ["carnes", "lacteos", "frutas", "cc", "helado"]:
                df = pd.read_csv(base / f"capacity_{name}.csv", index_col=0)
                df.index = df.index.astype(int)
                df.columns = [int(c) for c in df.columns]
                self.capacity_tables[name.upper()] = df
            self.evap_dual = self._load_evaps(base / "evaporadores_dual_btuhr.csv")
            self.evap_frontal = self._load_evaps(base / "evaporadores_frontal_btuhr.csv")

        # Fit-check y WEFM opcional
        assets = Path("data/cuartos_frios/wefm_fitcheck_assets")
        if not assets.exists():
            assets = Path("data/wefm_fitcheck_assets")  # legacy ubicación
        self.evap_wefm = None
        self.evap_dims = None
        self.meta = None
        if assets.exists():
            try:
                self.evap_wefm = self._load_evaps(assets / "evaporadores_frontal_wefm_btuhr.csv")
                self.evap_dims = pd.read_csv(assets / "evaporadores_dimensiones_mm.csv")
                self.meta = json.loads((assets / "evaporadores_meta.json").read_text(encoding="utf-8"))
            except Exception:
                self.evap_wefm = None
                self.evap_dims = None
                self.meta = None

        self.tevap_columns = sorted([int(x) for x in self.evap_dual.columns])

    @staticmethod
    def _load_evaps(path: Path) -> pd.DataFrame:
        df = pd.read_csv(path, index_col=0)
        df.index = df.index.astype(str)
        df.columns = [int(c) for c in df.columns]
        return df

    # ---------------- reglas utilitarias ----------------
    @staticmethod
    def m_to_ft_round(m: float) -> int:
        return int(round(m * 3.28))

    def clamp_ft(self, ft: int) -> int:
        return max(self.config["dimension_limits_ft"]["min"], ft)

    def validate_dimension_ft(self, ft: int) -> bool:
        lim = self.config["dimension_limits_ft"]
        return lim["min"] <= ft <= lim["max"]

    def height_bucket(self, height_m: float) -> Tuple[int | None, int | None]:
        h_ft_round = self.m_to_ft_round(height_m)
        if h_ft_round <= self.config["height_bucket_thresholds_ft_rounded"]["8_max"]:
            return h_ft_round, 8
        if h_ft_round <= self.config["height_bucket_thresholds_ft_rounded"]["10_max"]:
            return h_ft_round, 10
        if h_ft_round <= self.config["height_bucket_thresholds_ft_rounded"]["12_max"]:
            return h_ft_round, 12
        return h_ft_round, None

    def height_factor(self, bucket: int) -> float:
        return float(self.config["height_factors"][str(bucket)])

    def pick_usage_sheet(self, usage: str) -> str:
        u = usage.strip().upper()
        if "HELADO" in u:
            return "HELADO"
        if "CONGEL" in u or u in ("CC", "COMIDA CONGELADA"):
            return "CC"
        if "CARNE" in u:
            return "CARNES"
        if "LACT" in u:
            return "LACTEOS"
        if "FRU" in u or "FRUTA" in u or "FRUVER" in u:
            return "FRUTAS"
        if "PROCESO" in u:
            return "PROCESO"
        return "FRUTAS"

    def floor_tevap_column(self, tevap_f: float, cols: Optional[List[int]] = None) -> int:
        use_cols = cols if cols is not None else self.tevap_columns
        eligible = [c for c in use_cols if c <= tevap_f]
        return eligible[-1] if eligible else use_cols[0]

    def _fits_room(self, model: str, L_mm: float, W_mm: float, H_mm: float) -> bool:
        if self.evap_dims is None or self.meta is None:
            return True
        row = self.evap_dims[self.evap_dims["modelo"] == model]
        if row.empty:
            return True
        r = row.iloc[0]
        X, B, H = float(r["X_mm"]), float(r["B_mm"]), float(r["H_mm"])
        allow_rotate = bool(self.meta["reglas"]["fitcheck"].get("allow_rotate", True))
        validate_height = bool(self.meta["reglas"]["fitcheck"].get("validate_height", True))
        fits_base = (X <= L_mm and B <= W_mm) or (allow_rotate and X <= W_mm and B <= L_mm)
        fits_height = (not validate_height) or (H <= H_mm)
        return fits_base and fits_height

    def select_evaporator(
        self, family: Family, table: pd.DataFrame, tevap_f: float, load_per_evap: float,
        L_mm: float, W_mm: float, H_mm: float
    ) -> Tuple[str, float, bool, str]:
        rule_min_frac = float(self.config["selection_rules"]["min_load_fraction_for_placeholder"])
        allow_over = float(self.config["selection_rules"]["allow_overload_multiplier"])
        table_cols = [int(c) for c in table.columns]
        col = self.floor_tevap_column(tevap_f, table_cols)

        models = list(table.index)
        smallest_cap = float(table.loc[models[0], col])
        if load_per_evap <= rule_min_frac * smallest_cap:
            placeholder = self.config["placeholder_models"][family]
            return placeholder, smallest_cap, True, "Carga muy baja; placeholder"

        for m in models:
            if not self._fits_room(m, L_mm, W_mm, H_mm):
                continue
            cap = float(table.loc[m, col])
            if load_per_evap / allow_over <= cap:
                label = f"{m} - {'FRONTAL' if family=='frontal' else 'DUAL'}"
                return label, cap, True, ""
        return "REVISE LA CANTIDAD DE EVAPORADORES", float(table.loc[models[-1], col]), False, "No se encontró modelo que quepa"

    # ---------------- cálculo principal ----------------
    def compute(self, inp: ColdRoomInputs) -> ColdRoomResult:
        msgs: List[str] = []

        if inp.n_evaporators is not None and inp.n_evaporators <= 0:
            return ColdRoomResult(valid=False, messages=["n_evaporators debe ser >= 1"])

        L_ft = self.clamp_ft(self.m_to_ft_round(inp.length_m))
        W_ft = self.clamp_ft(self.m_to_ft_round(inp.width_m))
        H_ft, bucket = self.height_bucket(inp.height_m)

        height_msg = None
        if bucket is None:
            # fuerza a 12 ft para no cortar
            height_msg = "Altura fuera de rango (8/10/12 ft); se usó 12 ft."
            bucket = 12
        height_factor = self.height_factor(bucket)
        safety = float(inp.safety_factor) if inp.safety_factor is not None else float(self.config["safety_factor_default"])

        sheet = self.pick_usage_sheet(inp.usage)
        table = self.capacity_tables.get(sheet)
        if table is None:
            msgs.append(f"No hay tabla de capacidad para '{sheet}'. Se usa FRUTAS.")
            table = self.capacity_tables["FRUTAS"]

        # Carga base
        if self.validate_dimension_ft(L_ft) and self.validate_dimension_ft(W_ft):
            kbtuh = float(table.loc[W_ft, L_ft])
            load_btu = kbtuh * 1000.0 * height_factor * safety
        else:
            side_eq = int(round(sqrt(L_ft * W_ft)))
            side_eq = max(self.config["dimension_limits_ft"]["min"], min(side_eq, self.config["dimension_limits_ft"]["max"]))
            kbtuh = float(table.loc[side_eq, side_eq])
            load_btu = kbtuh * 1000.0 * height_factor * safety
            msgs.append("Dimensiones fuera de 4..42 ft; se usó lado equivalente (diagonales).")

        prof_key = inp.usage.strip().upper()
        if prof_key in ("CC",):
            prof_key = "COMIDA CONGELADA"
        if prof_key == "FRUTAS":
            prof_key = "FRUVER"
        prof = self.usage_profiles.get(prof_key)
        if not prof:
            msgs.append(f"No se encontró perfil de uso para '{inp.usage}', se usa 'FRUVER'.")
            prof = self.usage_profiles["FRUVER"]
        tevap_f = float(prof["tevap_f"])
        tcam_f = float(prof["tcam_f"])

        # Selección de familia
        forced_tab = None
        if inp.family_override and inp.family_override != "auto":
            if inp.family_override == "dual":
                family = "dual"
            else:
                family = "frontal"
                if inp.family_override == "frontal_wef":
                    forced_tab = self.evap_frontal
                elif inp.family_override == "frontal_wefm":
                    forced_tab = self.evap_wefm if self.evap_wefm is not None else self.evap_frontal
        else:
            family_lookup = {
                "HELADO": "HELADOS",
                "CC": "COMIDA CONGELADA",
                "FRUTAS": "FRUVER",
                "CARNES": "CARNES",
                "LACTEOS": "LACTEOS",
                "PROCESO": "PROCESO",
            }
            family_key = family_lookup.get(sheet, sheet)
            family = self.config["family_by_usage"].get(family_key, "frontal")
        family = "frontal" if family is None else family

        frontal_table = self.evap_frontal
        if family == "frontal" and self.evap_wefm is not None:
            umbral = float(self.meta["reglas"]["frontal_altura_umbral_m"]) if self.meta else 3.0
            frontal_table = self.evap_wefm if inp.height_m > umbral else self.evap_frontal
        if forced_tab is not None:
            frontal_table = forced_tab

        L_mm = inp.length_m * 1000.0
        W_mm = inp.width_m * 1000.0
        H_mm = inp.height_m * 1000.0

        def tiro_lim(fam: str, tab) -> float:
            if fam == "frontal" and tab is self.evap_wefm:
                return 14.0
            return 8.0

        CLEARANCE = 400.0  # mm de holgura por lado

        def eval_n(n: int, fam: str, tab):
            per_evap = load_btu / float(n)
            model, cap, fit_ok, fit_msg = self.select_evaporator(
                fam, tab, tevap_f, per_evap, L_mm + CLEARANCE, W_mm + CLEARANCE, H_mm
            )
            if fam == "frontal" and (not fit_ok or "REVISE" in (model or "")) and self.evap_wefm is not None:
                tab = self.evap_wefm
                model, cap, fit_ok, fit_msg = self.select_evaporator(
                    "frontal", tab, tevap_f, per_evap, L_mm + CLEARANCE, W_mm + CLEARANCE, H_mm
                )
            seg_m = max(inp.length_m, inp.width_m)
            tlim = tiro_lim(fam, tab)
            tiro_ok = seg_m <= tlim
            tiro_msg = "" if tiro_ok else f"Tiro requerido {seg_m:.1f} m > tiro permitido {tlim:.1f} m"
            util = per_evap / cap if cap else 0.0
            valid = bool(fit_ok and tiro_ok and cap and model)
            return {
                "n": n,
                "model": model,
                "cap": cap,
                "fit_ok": fit_ok,
                "fit_msg": fit_msg,
                "tiro_ok": tiro_ok,
                "tiro_msg": tiro_msg,
                "util": util,
                "per_evap": per_evap,
                "tab": tab,
                "valid": valid,
            }

        table_evap = frontal_table if family == "frontal" else self.evap_dual
        auto_note = ""
        if inp.n_evaporators is None:
            candidates = [eval_n(n, family, table_evap) for n in range(1, 5)]
            valids = [c for c in candidates if c["valid"]]
            if valids:
                small_room = (
                    inp.length_m <= 3.0
                    and inp.width_m <= 3.0
                    and inp.height_m <= 3.0
                )

                # evitamos N=4 salvo que no haya otra opción válida
                valids_pref = [c for c in valids if c["n"] <= 3] or valids

                def score(c):
                    s = 0.0
                    # Preferimos utilización en 70-90 %
                    if 0.70 <= c["util"] <= 0.90:
                        s += 1.0
                    else:
                        s -= abs(c["util"] - 0.80)
                    # Preferencia por N=1 en cuartos pequeños; en general ligera preferencia por N=2
                    if small_room and c["n"] == 1:
                        s += 0.20
                    if c["n"] == 2:
                        s += 0.15
                    elif c["n"] == 3:
                        s += 0.05
                    elif c["n"] == 4:
                        s -= 0.30  # penalizamos fuerte N=4
                    return s

                valids_pref.sort(key=score, reverse=True)
                chosen = valids_pref[0]
                auto_note = f"AUTO eligió N={chosen['n']} (util {chosen['util']*100:.1f}%)."
            else:
                msgs.append("No hay combinación válida (capacidad/tiro/cabida); revise N o dimensiones.")
                chosen = candidates[0] if candidates else {"n": None, "model": None, "cap": 0, "fit_ok": False, "fit_msg": "Sin candidatos", "tiro_ok": False, "tiro_msg": "", "per_evap": 0}
        else:
            chosen = eval_n(inp.n_evaporators, family, table_evap)
            if not chosen["valid"]:
                msgs.append("N seleccionado no cumple capacidad/tiro/cabida. Pruebe Auto.")

        evap_model = chosen["model"]
        evap_cap = chosen["cap"]
        fit_ok = chosen["fit_ok"]
        fit_msg = chosen["fit_msg"]
        tiro_ok = chosen["tiro_ok"]
        tiro_msg = chosen["tiro_msg"]
        load_per_evap = chosen["per_evap"]
        n_used = chosen["n"]

        if height_msg:
            msgs.append(height_msg)

        return ColdRoomResult(
            valid=True,
            messages=msgs,
            length_ft=L_ft,
            width_ft=W_ft,
            height_ft=H_ft,
            height_bucket_ft=bucket,
            height_factor=height_factor,
            load_btu_hr=load_btu,
            tevap_f=tevap_f,
            tcam_f=tcam_f,
            evap_family=family,
            evap_model=evap_model,
            evap_capacity_btu_hr=evap_cap,
            load_per_evap_btu_hr=load_per_evap,
            fit_ok=fit_ok,
            fit_msg=fit_msg,
            tiro_ok=tiro_ok,
            tiro_msg=tiro_msg,
            n_requested=inp.n_evaporators,
            n_used=n_used,
            auto_note=auto_note,
        )
