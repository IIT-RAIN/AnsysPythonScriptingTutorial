def extract_max_stress(stress_result):
    """
    Evaluate and return maximum von Mises stress (Pa).
    """
    stress_result.EvaluateAllResults()
    return stress_result.Maximum