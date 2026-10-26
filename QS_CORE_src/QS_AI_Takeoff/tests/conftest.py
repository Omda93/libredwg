import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

FIX_RULES = os.path.join(HERE, "fixtures", "project_rules.json")
FIX_LAYERS = os.path.join(HERE, "fixtures", "layer_mapping.json")


@pytest.fixture(autouse=True, scope="session")
def _isolated_config():
    """Tests use a frozen config, never the user's live config/ files (the GUI layer picker edits those).
    Session scope so module-scoped fixtures also see it."""
    from services import project_service as ps
    orig = ps.Project.load.__func__

    def load(cls, rules_path=FIX_RULES, layers_path=FIX_LAYERS):
        return orig(cls, rules_path, layers_path)

    mp = pytest.MonkeyPatch()
    mp.setattr(ps.Project, "load", classmethod(load))
    yield
    mp.undo()
