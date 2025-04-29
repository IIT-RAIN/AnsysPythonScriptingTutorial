### File: main.py
```python
import os
import datetime
from setup_log import setup_log
from load_excel_data import load_excel_data
from apply_loads import apply_loads
from run_simulation import run_simulation
from extract_max_stress import extract_max_stress
from write_excel_results import write_excel_results
from plot_results import plot_results

# --- Paths ---
input_excel_path = r"C:\path\to\input_loads.xlsx"
output_excel_path = r"C:\path\to\output_results.xlsx"
output_plot_path = r"C:\path\to\stress_plot.png"
output_log_path = r"C:\path\to\analysis_log.txt"
joint_surfaces_selection = "JointSurfaces"

# 1. Initialize log
t_log = setup_log(output_log_path)

# 2. Read load cases
load_cases = load_excel_data(input_excel_path, t_log)
if not load_cases:
    t_log.close()
    raise Exception("No load cases found.")

# 3. Get model and analysis
model = ExtAPI.DataModel.Project.Model
analysis = model.Analyses[0]

# 4. Named selection
ns_objs = ExtAPI.DataModel.Tree.GetObjectsByName(joint_surfaces_selection)
if not ns_objs:
    t_log.close()
    raise Exception(f"Named selection '{joint_surfaces_selection}' not found.")
joint_ns = ns_objs[0]

# 5. Add loads and stress result
with ExtAPI.DataModel.Tree.Suspend():
    force_load = analysis.AddForce()
    force_load.Name = "Joint Force"
    force_load.Location = joint_ns
    force_load.DefineBy = LoadDefineBy.Components

    moment_load = analysis.AddMoment()
    moment_load.Name = "Joint Moment"
    moment_load.Location = joint_ns
    moment_load.DefineBy = LoadDefineBy.Components

stress_res = None
for res in analysis.Solution.Children:
    if res.ObjectType == Ansys.ACT.Automation.Mechanical.Results.EquivalentStress:
        stress_res = res
        break
if stress_res is None:
    stress_res = analysis.Solution.AddEquivalentStress()
    stress_res.Name = "Equivalent Stress"

t_log.write("Setup completed. Beginning simulations...\n")

# 6. Loop cases
max_stresses = []
for idx, loads in enumerate(load_cases, 1):
    t_log.write(f"Case {idx}: {loads} ... ")
    apply_loads(force_load, moment_load, loads)
    if not run_simulation(analysis, t_log):
        t_log.write("Solve failed.\n")
        max_stresses.append(None)
        continue
    max_val = extract_max_stress(stress_res)
    max_stresses.append(max_val)
    t_log.write(f"Max Stress = {max_val/1e6:.2f} MPa\n")

# 7. Write results and plot
write_excel_results(load_cases, max_stresses, output_excel_path, t_log)
plot_results(max_stresses, output_plot_path, t_log)

t_log.write(f"End Time: {datetime.datetime.now()}\n")
t_log.write("Automation completed.\n")
t_log.close()