import importlib.util
from pathlib import Path
import sys

import pytest


@pytest.fixture
def plugin_root() -> Path:
    return Path(__file__).parents[1]


@pytest.fixture
def fixtures(plugin_root: Path) -> Path:
    return plugin_root / "tests" / "fixtures"


@pytest.fixture
def syllabus_module(plugin_root: Path):
    path = plugin_root / "skills" / "course-syllabus-compiler" / "scripts" / "syllabus.py"
    spec = importlib.util.spec_from_file_location("course_syllabus_script", path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module
