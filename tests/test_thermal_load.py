import math

from logic.thermal_load.calculator import ThermalLoadCalculator


def test_golden_case_tolerancia():
    calc = ThermalLoadCalculator()
    res, expected = calc.run_golden_case()
    btuh = res.total_btuh
    target = expected["total_Btuh_excel"]
    tol = target * 0.05  # 5%
    assert math.isclose(btuh, target, abs_tol=tol), f"Btuh {btuh} vs {target}"
    comp = res.components
    # Checks suaves por componente (solo que no sean cero)
    assert comp.transmission_btuh > 0
    assert comp.infiltration_btuh > 0
    assert comp.internal_btuh > 0
