# Python--Digsilent
Welcome to the Python scripting for DIgSILENT. This documentation will teach you how to leverage Python for automating tasks in DIgSILENT PowerFactory.
Python Scripting for DIgSILENT PowerFactory: Learn how to automate simulations, perform grid analysis, and enhance power system workflows using Python.
# Python Scripting in DIgSILENT


## Table of Contents

1. [Connect to PowerFactory Application](#connect-to-powerfactory-application)
2. [Activate Project](#activate-project)

---

## Connect to PowerFactory Application

Before you can perform any operations within DIgSILENT PowerFactory using Python, you need to establish a connection to the PowerFactory application. This connection allows your Python script to interact with PowerFactory's objects and functionalities.

### Steps to Connect:

1. **Import the PowerFactory Module**: PowerFactory provides a Python module that facilitates interaction with its environment. Begin by importing this module in your script.
   ```python
    try:
    import powerfactory as pf
    print_and_log("powerfactory module imported successfully", f)
except ImportError as e:
    print_and_log(f"Error importing powerfactory: {e}", f)
    f.close()
    sys.exit(1)
   ```
2. **Initialize the PowerFactory Application**: Use the GetApplication() method to initialize and obtain a reference to the PowerFactory application instance.

   ```python
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
   ```

### Example:

Here's a complete example that connects to the PowerFactory application and verifies the connection:

```python
import sys
sys.path.append(r"C:\\Program Files\\DIgSILENT\\PowerFactory 2022 SP2\\Python\\3.8")
import powerfactory

app = powerfactory.GetApplication()

if not app:
    raise Exception("Could not connect to PowerFactory application.")
else:
    print("Connected to PowerFactory successfully.")
```

## Activate Project

The `ActivateProject()` method in the Python API is used to activate a specific project within PowerFactory. Once activated, the project becomes the context for all subsequent actions like running simulations, managing scenarios, or executing scripts.

```python
project_name = "Nine-bus System"
print_and_log(f"Activating project: {project_name}", f)
ierr = app.ActivateProject(project_name)
project = app.GetActiveProject()
if not project:
    print_and_log(f"No se pudo activar el proyecto: {project_name}", f)
    f.close()
    sys.exit()

print_and_log(f"Proyecto '{project_name}' activado con éxito", f)
```

### Parameters:

- `project_name`: Name of the project to be activated.

### Returns:

- The method returns the activated project object. If the project is not found or cannot be activated, it returns `None`.

### Find the grid:

```python
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
```
