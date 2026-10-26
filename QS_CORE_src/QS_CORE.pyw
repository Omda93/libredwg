"""QS CORE launcher - double-click to start (no console window). Same program as dwg_measure_extractor_v6.py."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    import dwg_measure_extractor_v6 as qs_core
    qs_core.main()
