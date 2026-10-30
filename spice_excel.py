from pathlib import Path
from datetime import datetime

import re
import shutil
import sys

import pandas as pd
import yaml


REPO = Path(__file__).resolve().parent
BASE_CONFIG = REPO / "base_config.yaml"


def load_paths():
    with open(BASE_CONFIG, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    pcfg = cfg["path_configs"]

    project_root = (
        REPO
        / "data"
        / pcfg["data_folder_name"]
        / pcfg["project_name"]
    )

    input_root = project_root / "input"
    global_root = input_root / "global_input"

    scenario_name = pcfg["input_scenario_name"]
    scenario_root = input_root / scenario_name

    return (
        project_root,
        input_root,
        global_root,
        scenario_root,
        scenario_name,
    )


def prepare_for_excel(df):
    """
    Evita que Excel convierta TRUE/FALSE en booleanos localizados
    como VERDADERO/FALSO.
    """
    df = df.copy()

    def normalize_value(value):
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"

        if isinstance(value, str):
            cleaned = value.strip()

            if cleaned.upper() in {"TRUE", "VERDADERO"}:
                return "TRUE"

            if cleaned.upper() in {"FALSE", "FALSO"}:
                return "FALSE"

        return value

    return df.map(normalize_value)


def prepare_for_csv(df):
    """
    Normaliza booleanos provenientes de Excel al formato esperado
    por SPICE: TRUE / FALSE.
    """
    df = df.copy()

    def normalize_value(value):
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"

        if isinstance(value, str):
            cleaned = value.strip()

            if cleaned.upper() in {"TRUE", "VERDADERO"}:
                return "TRUE"

            if cleaned.upper() in {"FALSE", "FALSO"}:
                return "FALSE"

        return value

    return df.map(normalize_value)


def safe_sheet_name(csv_file, root, used):
    rel = csv_file.relative_to(root).with_suffix("")

    name = rel.as_posix().replace("/", "_")

    name = re.sub(r"[\[\]:*?/\\]", "_", name)

    name = name[:31] or "Sheet"

    base = name
    i = 2

    while name in used:
        suffix = f"_{i}"
        name = base[:31 - len(suffix)] + suffix
        i += 1

    used.add(name)

    return name


def write_workbook(csv_root, workbook, relative_base):
    csv_files = (
        sorted(csv_root.rglob("*.csv"))
        if csv_root.exists()
        else []
    )

    if not csv_files:
        print(f"Sin CSV en: {csv_root}")
        return 0

    mapping = []
    used = set()

    with pd.ExcelWriter(
        workbook,
        engine="openpyxl",
    ) as writer:

        for csv_file in csv_files:
            rel = csv_file.relative_to(relative_base)

            sheet = safe_sheet_name(
                csv_file,
                csv_root,
                used,
            )

            print(f"  Leyendo {rel} ...")

            df = pd.read_csv(
                csv_file,
                sep=",",
                low_memory=False,
            )

            # Mantener TRUE/FALSE como texto
            df = prepare_for_excel(df)

            df.to_excel(
                writer,
                sheet_name=sheet,
                index=False,
            )

            mapping.append({
                "sheet": sheet,
                "relative_csv_path": rel.as_posix(),
                "rows": len(df),
                "columns": len(df.columns),
            })

        pd.DataFrame(mapping).to_excel(
            writer,
            sheet_name="_SPICE_MAP",
            index=False,
        )

        writer.book["_SPICE_MAP"].sheet_state = "hidden"

    print(f"  -> Creado: {workbook.name}")

    return len(csv_files)


def export_to_excel():
    (
        project_root,
        input_root,
        global_root,
        scenario_root,
        scenario_name,
    ) = load_paths()

    print("\nCreando libros de Excel...\n")

    total = 0

    total += write_workbook(
        global_root,
        project_root / "GLOBAL_INPUT.xlsx",
        input_root,
    )

    for sector in [
        "power",
        "industry",
        "transport",
    ]:
        sector_root = scenario_root / sector

        workbook = (
            project_root
            / f"{scenario_name}_{sector.upper()}.xlsx"
        )

        total += write_workbook(
            sector_root,
            workbook,
            input_root,
        )

    print("\nListo.")
    print(f"CSV incluidos en total: {total}")

    print(
        f"\nLos libros están en:\n"
        f"{project_root}"
    )

    print(
        "\nEdita los .xlsx normalmente en Excel, "
        "guárdalos y ciérralos."
    )

    print("Después ejecuta:")
    print("python spice_excel.py import")


def import_workbook(
    workbook,
    input_root,
    backup_root,
):
    if not workbook.exists():
        print(
            f"  No existe {workbook.name}; "
            f"se omite."
        )
        return 0

    mapping = pd.read_excel(
        workbook,
        sheet_name="_SPICE_MAP",
        dtype=str,
    )

    required = {
        "sheet",
        "relative_csv_path",
    }

    if not required.issubset(mapping.columns):
        raise ValueError(
            f"{workbook.name} no contiene "
            f"la hoja de mapeo esperada."
        )

    count = 0

    for _, row in mapping.iterrows():
        sheet = row["sheet"]
        rel = Path(row["relative_csv_path"])

        csv_file = input_root / rel

        if csv_file.exists():
            backup_file = backup_root / rel

            backup_file.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            shutil.copy2(
                csv_file,
                backup_file,
            )

        print(
            f"  {workbook.name}: "
            f"{sheet} -> {rel}"
        )

        df = pd.read_excel(
            workbook,
            sheet_name=sheet,
        )

        # Recuperar TRUE/FALSE correctamente
        df = prepare_for_csv(df)

        csv_file.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        df.to_csv(
            csv_file,
            index=False,
            sep=",",
            lineterminator="\n",
        )

        count += 1

    return count


def import_to_csv():
    (
        project_root,
        input_root,
        global_root,
        scenario_root,
        scenario_name,
    ) = load_paths()

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    backup_root = (
        project_root
        / f"_csv_backup_{timestamp}"
    )

    workbooks = [
        project_root / "GLOBAL_INPUT.xlsx",
        project_root
        / f"{scenario_name}_POWER.xlsx",
        project_root
        / f"{scenario_name}_INDUSTRY.xlsx",
        project_root
        / f"{scenario_name}_TRANSPORT.xlsx",
    ]

    print(
        "\nActualizando CSV "
        "desde Excel...\n"
    )

    total = 0

    for workbook in workbooks:
        total += import_workbook(
            workbook,
            input_root,
            backup_root,
        )

    if total == 0:
        print(
            "No encontré libros "
            "para importar."
        )
        return

    print("\nListo.")
    print(
        f"CSV actualizados: {total}"
    )

    print(
        "Todos fueron escritos "
        "con separador coma (,)."
    )

    print(
        f"\nBackup de los CSV anteriores:\n"
        f"{backup_root}"
    )


def main():
    if (
        len(sys.argv) != 2
        or sys.argv[1].lower()
        not in {"export", "import"}
    ):
        print("\nUso:")
        print(
            "  python spice_excel.py export"
        )
        print(
            "  python spice_excel.py import"
        )
        return

    mode = sys.argv[1].lower()

    if mode == "export":
        export_to_excel()
    else:
        import_to_csv()


if __name__ == "__main__":
    main()