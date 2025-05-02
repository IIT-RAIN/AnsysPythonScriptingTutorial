def run_simulation(analysis, log_file):
    """
    Solve the analysis and return True if successful.
    """
    try:
        analysis.Solve()
        return True
    except Exception as err:
        log_file.write(f"ERROR: Solve failed. Exception: {err}\n")
        return False