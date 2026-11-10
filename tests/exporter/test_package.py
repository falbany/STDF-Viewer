def test_package_importable():
    import deps.exporter as pkg

    assert pkg.__doc__ is not None


def test_package_has_no_openpyxl_imports():
    """Guard: the runtime package renders through xlsxwriter only —
    openpyxl survives as a DEV dependency (test read-back) and must never
    be imported from deps/exporter (docstring mentions are fine)."""
    import re
    from pathlib import Path

    package_dir = Path(__file__).resolve().parents[2] / "deps" / "exporter"
    import_line = re.compile(r"^\s*(import openpyxl|from openpyxl)", re.M)
    offenders = []
    for path in sorted(package_dir.glob("*.py")):
        source = path.read_text(encoding="utf-8")
        match = import_line.search(source)
        if match:
            offenders.append(f"{path.name}: {match.group(0).strip()}")
    assert offenders == []
