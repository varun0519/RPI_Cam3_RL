#!/usr/bin/env python3
"""
Automated KiCad Project Fabrication Outputs Exporter

Exports schematic, PCB, BOM, Gerbers, drill files, assembly docs, and 3D visualizations from a KiCad project.
Generates a README.md summarizing all outputs.

Usage:
    python export_fab_outputs.py
    (Run in the directory containing the .kicad_pro file)
Requirements:
    - kicad-cli installed and in PATH
    - Python 3.x
    - pandas library
"""

import os
import time
import sys
import logging
import subprocess
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

# -------------------- Logging Setup --------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)

# -------------------- Dependency Checks --------------------
def check_dependencies():
    try:
        subprocess.run(["kicad-cli", "--version"], check=True, stdout=subprocess.PIPE)
    except Exception:
        logger.error("kicad-cli is not installed or not in PATH.")
        sys.exit(1)
    try:
        import pandas
    except ImportError:
        logger.error("pandas is not installed. Please install it with 'pip install pandas'.")
        sys.exit(1)

# -------------------- Utility Functions --------------------
def get_project_name(project_dir=None):
    """
    Returns the project name by searching for a .kicad_pro file in the given directory.
    If not found, returns None.
    """
    if project_dir is None:
        project_dir = os.getcwd()
    for fname in os.listdir(project_dir):
        if fname.endswith(".kicad_pro"):
            return os.path.splitext(fname)[0]
    return None


PROJECT_NAME = get_project_name()
if PROJECT_NAME is None:
    print("❌ No .kicad_pro file found in the current directory.")
    sys.exit(1)

PROJECT_DIR = os.getcwd()
OUTPUT_DIR = os.path.join(PROJECT_DIR, "output")


def ensure_dir(directory):
    os.makedirs(directory, exist_ok=True)

def run_command(command, output_file=None):
    logger.info(f"Running: {' '.join(command) if isinstance(command, list) else command}")
    try:
        subprocess.run(command, check=True, shell=isinstance(command, str))
        if output_file and not os.path.exists(output_file):
            logger.error(f"Expected output file not found: {output_file}")
            return False
        logger.info("Done")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"Command failed: {e}")
        return False

# -------------------- Export Functions --------------------
def generate_schematic_pdf(sch_file, output_dir, project_name):
    out = os.path.join(output_dir, f"{project_name}_schematic.pdf")
    cmd = ["kicad-cli", "sch", "export", "pdf", sch_file, "--output", out]
    return run_command(cmd, out)

def export_schematic_svg(sch_file, output_dir, project_name):
    out = os.path.join(output_dir, f"{project_name}_schematic.svg")
    cmd = ["kicad-cli", "sch", "export", "svg", sch_file, "--output", out]
    return run_command(cmd, out)

def run_erc(sch_file, output_dir):
    out = os.path.join(output_dir, "erc_report.txt")
    cmd = ["kicad-cli", "sch", "erc", sch_file, "--output", out]
    return run_command(cmd, out)

def run_drc(pcb_file, output_dir):
    out = os.path.join(output_dir, "drc_report.txt")
    cmd = ["kicad-cli", "pcb", "drc", pcb_file, "--output", out]
    return run_command(cmd, out)

def export_gerbers(pcb_file, output_dir, project_name):
    gerbers_dir = os.path.join(output_dir, "Gerbers")
    ensure_dir(gerbers_dir)
    cmd = [
        "kicad-cli", "pcb", "export", "gerbers", pcb_file,
        "--board-plot-params",
        "--layers", "F.Cu,In1.Cu,In2.Cu,B.Cu,F.Mask,B.Mask,F.Paste,B.Paste,F.Silkscreen,B.Silkscreen,Edge.Cuts",
        "--output", gerbers_dir
    ]
    # Output is a directory, just check directory exists
    return run_command(cmd)

def export_drill(pcb_file, output_dir):
    gerbers_dir = os.path.join(output_dir, "Gerbers")
    ensure_dir(gerbers_dir)
    cmd = ["kicad-cli", "pcb", "export", "drill", pcb_file, "--output", gerbers_dir]
    return run_command(cmd)

def export_position(pcb_file, output_dir, project_name):
    out = os.path.join(output_dir, "Gerbers", f"{project_name}-all-pos.csv")
    cmd = [
        "kicad-cli", "pcb", "export", "pos", pcb_file,
        "--format", "csv", "--units", "mm", "--exclude-dnp",
        "--output", out
    ]
    return run_command(cmd, out)

def export_bom(sch_file, output_dir, project_name):
    csv_path = os.path.join(output_dir, f"{project_name}_BOM.csv")
    xlsx_path = os.path.join(output_dir, f"{project_name}_BOM.xlsx")
    cmd = [
        "kicad-cli", "sch", "export", "bom", sch_file,
        "--group-by", "Value,Footprint,${DNP}",
        "--ref-range-delimiter", "",
        "--fields", "${ITEM_NUMBER},Reference,Value,Footprint,Description,${QUANTITY},${DNP},MPN,SKU,Link",
        "--output", csv_path
    ]
    success = run_command(cmd, csv_path)
    if success and os.path.exists(csv_path):
        df = pd.read_csv(csv_path)
        df.to_excel(xlsx_path, index=False)
        logger.info(f"Converted CSV BOM to Excel: {xlsx_path}")
        return True
    else:
        logger.error(f"BOM CSV not found for conversion: {csv_path}")
        return False

def export_top_assembly(pcb_file, output_dir, project_name):
    out = os.path.join(output_dir, f"{project_name}_Top_Assembly.pdf")
    cmd = [
        "kicad-cli", "pcb", "export", "pdf", pcb_file,
        "--layers", "F.Mask,F.Silkscreen,Edge.Cuts",
        "--black-and-white", "--output", out
    ]
    return run_command(cmd, out)

def export_bottom_assembly(pcb_file, output_dir, project_name):
    out = os.path.join(output_dir, f"{project_name}_Bottom_Assembly.pdf")
    cmd = [
        "kicad-cli", "pcb", "export", "pdf", pcb_file,
        "--layers", "B.Mask,B.Silkscreen,Edge.Cuts",
        "--black-and-white", "--mirror", "--output", out
    ]
    return run_command(cmd, out)

def take_3d_screenshots(pcb_file, output_dir, project_name):
    top = os.path.join(output_dir, f"{project_name}_3D_top.png")
    bottom = os.path.join(output_dir, f"{project_name}_3D_bottom.png")
    persp = os.path.join(output_dir, f"{project_name}_3D_perspective.png")
    cmds = [
        (["kicad-cli", "pcb", "render", pcb_file, "--side", "top", "--output", top], top),
        (["kicad-cli", "pcb", "render", pcb_file, "--side", "bottom", "--output", bottom], bottom),
        (["kicad-cli", "pcb", "render", pcb_file, "--side", "top", "--perspective", "--rotate", "315,0,0", "--output", persp], persp)
    ]
    results = []
    for cmd, outfile in cmds:
        results.append(run_command(cmd, outfile))
    return all(results)

# -------------------- README Generation --------------------
def write_readme(project_dir, output_dir, project_name):
    readme_path = os.path.join(project_dir, "README.md")
    content = f"""# {project_name}

## 🚀 How to Use Automation Script to generate fabrication outputs

### 🔧 Command Line (Terminal)

Run the following command **from the project directory** (where `.kicad_pro` file is located):

```bash
make
```
To generate zip of Gerbers for fabrication:
```bash
make release
```

Or run the script directly with Python:
```bash
python3 scripts/export_fab_outputs.py
```

### 3D view
#### Top View
![]({os.path.relpath(os.path.join(output_dir, project_name + '_3D_top.png'), project_dir)})
#### Bottom View
![]({os.path.relpath(os.path.join(output_dir, project_name + '_3D_bottom.png'), project_dir)})
#### Perspective View
![]({os.path.relpath(os.path.join(output_dir, project_name + '_3D_perspective.png'), project_dir)})

### Schematic
[{project_name}_schematic.pdf]({os.path.relpath(os.path.join(output_dir, project_name + '_schematic.pdf'), project_dir)})

### BOM
[{project_name}_BOM.xlsx]({os.path.relpath(os.path.join(output_dir, project_name + '_BOM.xlsx'), project_dir)})

### Top Assembly
[{project_name}_Top_Assembly.pdf]({os.path.relpath(os.path.join(output_dir, project_name + '_Top_Assembly.pdf'), project_dir)})

### Bottom Assembly
[{project_name}_Bottom_Assembly.pdf]({os.path.relpath(os.path.join(output_dir, project_name + '_Bottom_Assembly.pdf'), project_dir)})

### Interactive BOM
Download and open [ibom.html](bom/ibom.html) in browser
"""
    with open(readme_path, "w") as f:
        f.write(content)
    logger.info(f"README.md generated at {readme_path}")

# -------------------- Main Orchestration --------------------
def main():
    start_time = time.time()
    logger.info(f"Starting export for project: {PROJECT_NAME}")
    check_dependencies()

    ensure_dir(OUTPUT_DIR)

    sch_file = os.path.join(PROJECT_DIR, f"{PROJECT_NAME}.kicad_sch")
    pcb_file = os.path.join(PROJECT_DIR, f"{PROJECT_NAME}.kicad_pcb")

    if not os.path.isfile(sch_file):
        logger.error(f"Schematic file not found: {sch_file}")
        sys.exit(1)

    if not os.path.isfile(pcb_file):
        logger.error(f"PCB file not found: {pcb_file}")
        sys.exit(1)

    tasks = [
        (generate_schematic_pdf, (sch_file, OUTPUT_DIR, PROJECT_NAME)),
        (export_schematic_svg, (sch_file, OUTPUT_DIR, PROJECT_NAME)),
        (run_erc, (sch_file, OUTPUT_DIR)),
        (run_drc, (pcb_file, OUTPUT_DIR)),
        (export_bom, (sch_file, OUTPUT_DIR, PROJECT_NAME)),
        (export_gerbers, (pcb_file, OUTPUT_DIR, PROJECT_NAME)),
        (export_drill, (pcb_file, OUTPUT_DIR)),
        (export_position, (pcb_file, OUTPUT_DIR, PROJECT_NAME)),
        (export_top_assembly, (pcb_file, OUTPUT_DIR, PROJECT_NAME)),
        (export_bottom_assembly, (pcb_file, OUTPUT_DIR, PROJECT_NAME)),
        (take_3d_screenshots, (pcb_file, OUTPUT_DIR, PROJECT_NAME)),
    ]

    with ThreadPoolExecutor(max_workers=4) as executor:
        future_to_task = {executor.submit(func, *params): func.__name__ for func, params in tasks}
        for future in as_completed(future_to_task):
            task_name = future_to_task[future]
            try:
                result = future.result()
                if result:
                    logger.info(f"{task_name}: Success")
                else:
                    logger.error(f"{task_name}: Failed")
            except Exception as exc:
                logger.error(f"{task_name} generated an exception: {exc}")

    write_readme(PROJECT_DIR, OUTPUT_DIR, PROJECT_NAME)
    elapsed = time.time() - start_time
    logger.info(f"All tasks completed in {elapsed:.2f} seconds.")


if __name__ == "__main__":
    main()