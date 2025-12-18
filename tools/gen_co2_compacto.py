from pathlib import Path
import json
import pandas as pd


def main() -> None:
    base = Path("data") / "basedatos.xlsx"
    if not base.exists():
        raise SystemExit(f"No encontré {base}")

    df = pd.read_excel(base, sheet_name="OPCIONES CO2", header=None, dtype=str).fillna("")

    items = []
    for i in range(len(df.index)):
        q = str(df.iat[i, 0] or "").strip()
        if not q:
            continue
        b = str(df.iat[i, 1] or "").strip()
        c = str(df.iat[i, 2] or "").strip()

        # Detectar tipo
        if "#" in b or "#" in c:
            items.append({"pregunta": q, "tipo": "spin", "opciones": [0, 1, 2, 3, 4, 5]})
            continue

        raw = []
        for s in (b, c):
            if s:
                raw += [p.strip() for p in s.replace("|", ",").replace("/", ",").split(",") if p.strip()]

        if not raw:
            items.append({"pregunta": q, "tipo": "radio", "opciones": ["SI", "NO"]})
        else:
            seen, opts = set(), []
            for x in raw:
                up = x.upper()
                if up not in seen:
                    seen.add(up)
                    opts.append(x)
            up_opts = [o.upper() for o in opts]
            if up_opts in (["SI", "NO"], ["NO", "SI"]):
                items.append({"pregunta": q, "tipo": "radio", "opciones": ["SI", "NO"]})
            else:
                items.append({"pregunta": q, "tipo": "combo", "opciones": opts})

    out_path = Path("data") / "preguntas_opciones_co2.json"
    payload = {"version": 1, "items": items}
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Generado {out_path} con {len(items)} ítems")


if __name__ == "__main__":
    main()
