import sys
try:
    from PyQt6.QtCharts import QChart
    print("PyQt6.QtCharts IS available.")
except ImportError:
    print("PyQt6.QtCharts IS NOT available.")
try:
    import matplotlib
    print("matplotlib IS available.")
except ImportError:
    print("matplotlib IS NOT available.")
