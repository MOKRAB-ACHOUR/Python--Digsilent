import os
import sys

# ============================================================
# PowerFactory 2024 + Python 3.12 configuration
# ============================================================
PF_PATH = r"C:\Program Files\DIgSILENT\PowerFactory 2024"

os.environ["PATH"] = PF_PATH + os.pathsep + os.environ.get("PATH", "")
sys.path.insert(0, os.path.join(PF_PATH, "Python", "3.12"))

# ============================================================
# Output folder
# ============================================================
output_dir = r"C:\Users\mokrab\Downloads\resultados"
os.makedirs(output_dir, exist_ok=True)
output_file = os.path.join(output_dir, "resultados.txt")


def print_and_log(message, file):
    print(message)
    file.write(message + "\n")


f = open(output_file, "w", encoding="utf-8")

# ============================================================
# Import PowerFactory
# ============================================================
try:
    import powerfactory as pf

    print_and_log("powerfactory module imported successfully", f)
except ImportError as e:
    print_and_log(f"Error importing powerfactory: {e}", f)
    f.close()
    sys.exit(1)

# ============================================================
# Initialize PowerFactory (must be CLOSED before running)
# ============================================================
print_and_log("Trying to start PowerFactory...", f)

app = None
try:
    # This is the recommended way for external scripts
    app = pf.GetApplicationExt()  # starts PowerFactory
except pf.ExitError as e:
    print_and_log(f"GetApplicationExt failed with ExitError: {e}", f)
    print_and_log(
        "Possible causes: licence problem, CodeMeter not running, or wrong user.", f
    )
    f.close()
    raise

if app is None:
    print_and_log("app is None – PowerFactory could not be started", f)
    f.close()
    raise Exception("No se pudo inicializar PowerFactory")

print_and_log("PowerFactory started successfully!", f)

# Optional: show the GUI
try:
    app.Show()
    print_and_log("GUI shown", f)
except Exception:
    print_and_log("Could not show GUI (running in engine mode)", f)

# ============================================================
# Activate project
# ============================================================
project_name = "Nine-bus System"
print_and_log(f"Activating project: {project_name}", f)

ierr = app.ActivateProject(project_name)
project = app.GetActiveProject()

if not project:
    print_and_log(f"No se pudo activar el proyecto: {project_name}", f)
    f.close()
    sys.exit()

print_and_log(f"Proyecto '{project_name}' activado con éxito", f)

# ============================================================
# Find the grid
# ============================================================
grid_name = "Nine-bus System"
grids = app.GetCalcRelevantObjects("*.ElmNet")
grid = None

for g in grids:
    if grid_name in g.loc_name:
        grid = g
        break

if not grid:
    print_and_log(f"No se encontró la red: {grid_name}", f)
    f.close()
    sys.exit()

print_and_log(f"Red '{grid_name}' encontrada con éxito", f)

# ============================================================
# Execute Load Flow
# ============================================================
load_flow = app.GetFromStudyCase("ComLdf")
load_flow.iopt_net = 0
nerr = load_flow.Execute()

if nerr != 0:
    print_and_log(f"El flujo de carga no convergió para la red: {grid_name}", f)
    f.close()
    sys.exit()

print_and_log("Flujo de carga ejecutado correctamente", f)

# ============================================================
# Total losses
# ============================================================
summary_command = app.GetFromStudyCase("ComSh")
summary_command.iopt_cmd = 2
summary_command.Execute()

summary_grid = app.GetSummaryGrid()

if summary_grid:
    perdidas_totales_P = summary_grid.GetAttribute("m:LossP")
    perdidas_totales_Q = summary_grid.GetAttribute("m:LossQ")
    print_and_log(f"[Inicial] Pérdidas totales activas: {perdidas_totales_P:.2f} MW", f)
    print_and_log(
        f"[Inicial] Pérdidas totales reactivas: {perdidas_totales_Q:.2f} MVAr", f
    )
else:
    print_and_log("No se encontró una red activa para obtener pérdidas iniciales.", f)

# ============================================================
# Line losses and loading
# ============================================================
lineas_red = [
    line for line in grid.GetContents("*.ElmLne") if line.GetAttribute("outserv") == 0
]

if lineas_red:
    print_and_log("\nPérdidas y cargabilidad de todas las líneas en servicio:", f)
    for line in lineas_red:
        losses = line.GetAttribute("c:Losses")
        loading = line.GetAttribute("c:loading")
        msg = f"  Línea {line.loc_name}: Pérdidas {losses:.2f} kW, Cargabilidad {loading:.2f} %"
        print_and_log(msg, f)
else:
    print_and_log("No se encontraron líneas de transmisión en servicio.", f)

# ============================================================
# End
# ============================================================
f.close()
print("Script finalizado. Resultados guardados en:", output_file)

# Optional: close PowerFactory when finished
# del app
