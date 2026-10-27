import os
import pandas as pd

def convert_bom(input_csv="bom.csv", output_xls=False):
    """
    Konvertiert eine KiCad/Standard-BOM-CSV in das JLCPCB-BOM-Format
    mit einem benutzerdefinierten Präfix für den Dateinamen.
    """
    if not os.path.exists(input_csv):
        raise FileNotFoundError(f"Die Datei '{input_csv}' wurde nicht gefunden.")

    # 1. Präfix per interaktiver Benutzereingabe abfragen
    prefix = input("Bitte Präfix für den Dateinamen eingeben: ").strip()
    
    # Dateinamen zusammensetzen (z. B. 'MeinProjekt_bom.xlsx' oder Standard, falls leer)
    base_name = f"{prefix}_bom" if prefix else "bom_converted"

    # 2. CSV einlesen (BOM / UTF-8 sicherstellen)
    df = pd.read_csv(input_csv, encoding="utf-8-sig")
    df.columns = df.columns.str.strip()

    # 3. Ziel-DataFrame nach Template-Vorlage aufbauen
    template_df = pd.DataFrame()
    template_df["Comment"] = df["Value"]
    template_df["Designator"] = df["Designator"]
    template_df["Footprint"] = df["Footprint"]
    
    if "LCSC Part #" in df.columns:
        template_df["JLCPCB Part#(optional)"] = df["LCSC Part #"].fillna("")
    else:
        template_df["JLCPCB Part#(optional)"] = ""

    sheet_name = "JLCPCB_BOM Template"

    # 4. Speichern
    if output_xls:
        # Als .xls exportieren (erfordert: pip install xlwt)
        target_file = f"{base_name}.xls"
        with pd.ExcelWriter(target_file, engine="xlwt") as writer:
            template_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Erfolgreich gespeichert: {target_file}")
    else:
        # Als .xlsx exportieren (erfordert: pip install openpyxl)
        target_file = f"{base_name}.xlsx"
        with pd.ExcelWriter(target_file, engine="openpyxl") as writer:
            template_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Erfolgreich gespeichert: {target_file}")


if __name__ == "__main__":
    convert_bom("bom.csv", output_xls=False)