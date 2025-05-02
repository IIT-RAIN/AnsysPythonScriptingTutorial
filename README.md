This automation toolkit is designed for users of ANSYS Mechanical Workbench who need to perform a large number of static structural simulations. 
Specifically, multiple load cases are applied to a common geometry, such as a robotic joint or mechanical assembly. 
Instead of manually changing loads and solving each case one by one, this Python script automates the entire process.

Using the ANSYS scripting API (ExtAPI) along with Microsoft Excel COM interop, the script reads a list of force and moment inputs from an Excel file, applies them in sequence, solves each case, and logs the maximum von Mises stress results.
Output includes:
  -A structured Excel file containing the input and resulting stress for each case
  -A line plot visualising stress trends across load cases
  -A detailed log of all simulation steps, useful for troubleshooting and verification

This project is modularised for clarity and maintainability. Each major task (e.g. reading data, applying loads, solving, writing output) is encapsulated in its own script.
The main.py script coordinates these modules to execute the complete automation workflow.

While this project is tailored to ANSYS users familiar with Mechanical scripting and structural simulation workflows, the structure can be adapted by anyone comfortable with Python automation.
Below, the main script is explained, along with all the functions used in it.





The <<main.py>> file is the entry point. It:
	-Defines all file paths (input loads, output results, plot image, and log).
	-Initialises a log via setup_log, so directory creation and header writing happen in one place.
	-Loads the Excel data using load_excel_data, yielding a list of (Fx, Fy, Fz, Mx, My) tuples.
	-Fetches the ANSYS model and analysis objects from the Workbench API (ExtAPI).
	-Locates the “JointSurfaces” named selection for scoping loads.
	-Adds Force and Moment objects (component‐defined) and ensures an Equivalent Stress result is ready.
	-Loops over every load case:
	-Applies the current load via apply_loads.
	-Calls run_simulation; if it fails, logs "Solve failed" and records None.
	-On success, extracts the max von Mises stress with extract_max_stress, logs it, and appends it.
	-Writes out the Excel results with write_excel_results.
	-Plots stress vs. case index using plot_results.
	-Closes the log with an end timestamp.

FUNCTIONS:

<<setup_log.py>> summarises all log-file setup:
	-Directory creation: Ensures the log’s parent folder exists.
	-File opening: Opens analysis_log.txt for writing (overwriting any old log).
	-Header writing: Stamps the start time and log path.
	Return value: an open file handle (log_file) that every other module writes into.

<<load_excel_data.py>> handles Excel input via COM interop:
	-References Microsoft.Office.Interop.Excel to drive Excel invisibly.
	-Opens the workbook, picks the first worksheet, and finds the last used row in column A.
	-Iterates rows 1 – last, reading five columns: Fx, Fy, Fz, Mx, My.
	-Filters out any row missing a value.
	-Closes/quits Excel and releases COM objects to avoid open background Excel processes.
	-Logs the count of load cases read.
	Returns a list of tuples (5 floats).

<<apply_loads.py>> handles the repetitive work of pushing values into ANSYS load objects:
	-Accepts a Force object and a Moment object from the analysis.
	-Takes one load tuple (Fx, Fy, Fz, Mx, My).
	-Calls SetDiscreteValue(0, Quantity(value, unit)) on each X/Y/Z component.
	Because all five components are set in one place, adding Mz (or other components), which is not included in this case, is very easy.

<<run_simulation.py>> Keeps the solve‐and‐error logic in one spot:
	-Calls analysis.Solve() inside a try/except.
	-On success, returns True.
	-On solver exceptions, logs the error message, and returns False.
	Because the function gives back a True or a False, the loop can easily check whether to continue or skip, without messy logic.


<<extract_max_stress.py>>: After a successful solve, this module:
	-Calls stress_result.EvaluateAllResults() to refresh the result data.
	-Reads stress_result.Maximum, which holds the peak von Mises stress in Pascals.
	-Returns that numeric value for downstream logging, plotting, and Excel output.

<<write_excel_results.py>> writes all cases and stresses into a new Excel workbook:
	-Reuses the Excel COM interop technique to create a fresh workbook.
	-Writes a header row (Case, Fx, Fy, Fz, Mx, My, Max Von Mises).
	-Iterates through load_cases and max_stress_values in parallel:
		-Case index in column A.
		-Input values in columns B–F.
		-Stress (or "SolveFailed") in column G.
	-Saves and closes the workbook, releases COM objects, and logs the output path.

<<plot_results.py>> generates a line chart of stress vs. load‐case index:
	-Uses System.Windows.Forms.DataVisualization.Charting to build a Chart object.
	-Adds a single Series of type Line, iterating non‐None stresses (converted to MPa).
	-Configures axis titles (“Load Case” and “Max von Mises Stress (MPa)”).
	-Saves the chart to the specified PNG path and logs success or any plotting errors.
 
