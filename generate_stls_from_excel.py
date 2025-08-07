
import pandas as pd
import os
import subprocess
# Created by Abbas Madhwala
# === YOUR FILE PATHS please change based on your file locations === 
excel_path = r"C:\Users\PumpI\Desktop\Shaper Cover Maker\ShaperCoverValues.xlsx"
scad_template_path = r"C:\Users\PumpI\Desktop\Shaper Cover Maker\template_bowtie.scad"
scad_output_folder = r"C:\Users\PumpI\Desktop\Shaper Cover Maker\generated_scad"
stl_output_folder = r"C:\Users\PumpI\Desktop\Shaper Cover Maker\output_stls"
openscad_exe = r"C:\Users\PumpI\Desktop\OpenSCAD (Nightly)\openscad.exe"

# === MAKE FOLDERS IF NEEDED ===
os.makedirs(scad_output_folder, exist_ok=True)
os.makedirs(stl_output_folder, exist_ok=True)

# === READ EXCEL ===
df = pd.read_excel(excel_path)
df.columns = df.columns.str.strip()  # clean up column names

print("Columns loaded:", df.columns.tolist())

# === LOAD SCAD TEMPLATE ===
with open(scad_template_path, 'r') as f:
    template_code = f.read()

# === LOOP THROUGH EXCEL ROWS ===
for index, row in df.iterrows():
    hole_diameter = row["Hole Diameter"]
    wt_number = str(row["Wt-#"])
    det_number = str(row["DET-#"])
    

    # Replace placeholders
    code = template_code.format(
        hole_diameter=hole_diameter,
        wt_number=wt_number,
        det_number=det_number,
        
    )

    # Write .scad file
    scad_filename = f"cover_{index + 1}.scad"
    scad_path = os.path.join(scad_output_folder, scad_filename)
    with open(scad_path, "w") as f:
        f.write(code)

    # Generate .stl
    stl_filename = f"cover_{index + 1}.stl"
    stl_path = os.path.join(stl_output_folder, stl_filename)

    subprocess.run([openscad_exe, "-o", stl_path, scad_path])

print("✅ Bowtie STL generation complete.")
