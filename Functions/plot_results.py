import clr
from System.Windows.Forms.DataVisualization.Charting import Chart, ChartArea, Series, SeriesChartType, ChartImageFormat
from System.Drawing import Color

def plot_results(max_stress_values, output_plot_path, log_file):
    """
    Generate and save a line chart of max von Mises stress vs load case index.
    """
    try:
        chart = Chart()
        chart.Width, chart.Height = 600, 400
        chart.ChartAreas.Add(ChartArea("MainArea"))
        series = Series("Max von Mises Stress")
        series.ChartType = SeriesChartType.Line
        series.BorderWidth = 2

        for i, stress in enumerate(max_stress_values, start=1):
            if stress is not None:
                series.Points.AddXY(i, stress / 1e6)

        chart.Series.Add(series)
        chart.ChartAreas["MainArea"].AxisX.Title = "Load Case"
        chart.ChartAreas["MainArea"].AxisY.Title = "Max von Mises Stress (MPa)"
        chart.SaveImage(output_plot_path, ChartImageFormat.Png)
        log_file.write(f"Stress plot saved to: {output_plot_path}\n")
    except Exception as err:
        log_file.write(f"ERROR: Plotting failed. Exception: {err}\n")
        raise