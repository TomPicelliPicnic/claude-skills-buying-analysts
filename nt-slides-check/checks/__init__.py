import importlib
import pkgutil
import pathlib

from core.check_template import CheckTemplate


def load_checks() -> list[CheckTemplate]:
    checks = []
    pkg_dir = pathlib.Path(__file__).parent
    for info in pkgutil.iter_modules([str(pkg_dir)]):
        if not info.name.startswith("check_"):
            continue
        mod = importlib.import_module(f"checks.{info.name}")
        for obj in vars(mod).values():
            if isinstance(obj, type) and issubclass(obj, CheckTemplate) and obj is not CheckTemplate:
                checks.append(obj())
    return sorted(checks, key=lambda c: c.id)
