from Ansys.ACT.Automation.Mechanical import LoadDefineBy
from Ansys.ACT.Automation.Mechanical import Quantity

def apply_loads(force_obj, moment_obj, loads):
    """
    Apply force and moment components to the given load objects.
    """
    fx, fy, fz, mx, my = loads
    force_obj.XComponent.Output.SetDiscreteValue(0, Quantity(fx, "N"))
    force_obj.YComponent.Output.SetDiscreteValue(0, Quantity(fy, "N"))
    force_obj.ZComponent.Output.SetDiscreteValue(0, Quantity(fz, "N"))
    moment_obj.XComponent.Output.SetDiscreteValue(0, Quantity(mx, "N*m"))
    moment_obj.YComponent.Output.SetDiscreteValue(0, Quantity(my, "N*m"))