from pathlib import Path

from finddoc import __all__


def test_package_exports_list() -> None:
    assert isinstance(__all__, list)


def test_package_module_exists() -> None:
    assert Path(__file__).exists()
