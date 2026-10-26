"""Print the unit-related header variables of a DXF (diagnostic)."""
import sys

import ezdxf

doc = ezdxf.readfile(sys.argv[1])
for k in ("$INSUNITS", "$MEASUREMENT", "$LUNITS", "$ACADVER"):
    print(k, doc.header.get(k))
