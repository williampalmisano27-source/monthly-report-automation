import os
from pathlib import Path
import subprocess
import sys


def run_cli(*args):
    env = dict(os.environ, PYTHONPATH=str(Path(__file__).resolve().parents[1] / "src"))
    return subprocess.run([sys.executable, "-m", "report_rollover.cli", *map(str, args)],
                          capture_output=True, text=True, env=env)


def test_cli_success_and_repeat_refusal(tmp_path, make_report):
    source = tmp_path / "input"
    source.mkdir()
    make_report(source / "report.xlsx")
    output = tmp_path / "out"
    result = run_cli("--input", source, "--output", output)
    assert result.returncode == 0, result.stderr
    assert (output / "report.xlsx").exists()
    assert "1" in result.stdout
    repeat = run_cli("--input", source, "--output", output)
    assert repeat.returncode == 2
    assert "Traceback" not in repeat.stderr


def test_cli_invalid_workbook_is_actionable(tmp_path):
    source = tmp_path / "input"
    source.mkdir()
    (source / "broken.xlsx").write_bytes(b"broken")
    result = run_cli("--input", source, "--output", tmp_path / "out")
    assert result.returncode == 2
    assert "broken.xlsx" in result.stderr
    assert "Traceback" not in result.stderr


def test_cli_missing_arguments():
    result = run_cli()
    assert result.returncode == 2
    assert "--input" in result.stderr


import pytest


@pytest.mark.parametrize("kind", ["style_index", "shared_string", "font_size"])
def test_cli_structural_corruption_is_validation_error(tmp_path, make_report, corrupt_workbook, kind):
    source = tmp_path / "input"
    source.mkdir()
    file = make_report(source / "malformed.xlsx")
    corrupt_workbook(file, kind)
    result = run_cli("--input", source, "--output", tmp_path / "out")
    assert result.returncode == 2, result.stderr
    assert "malformed.xlsx" in result.stderr
    assert "Traceback" not in result.stderr
    assert not (tmp_path / "out").exists()
