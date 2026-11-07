from pathlib import Path

from pnp2pdf import cli


def _cards(card_folder):
    return card_folder(names=["1.png", "2.png"])


def _pdf_in(folder):
    return list(Path(folder).rglob("*.pdf"))


def test_cli_builds_pdf_without_stdout(capsys, tmp_path, card_folder):
    folder = _cards(card_folder)
    out = tmp_path / "out"

    code = cli.main(["--cards", str(folder), "--out", str(out),
                     "--card", "62,5x88,7", "--paper", "210x297",
                     "--grid", "2x2"])

    captured = capsys.readouterr()
    assert code == 0
    assert captured.out == ""
    assert _pdf_in(out)


def test_decimal_comma_and_times_sign_parse():
    assert cli._parse_size("62,5×88,7", "Card width", "Card height") == (62.5, 88.7)
    assert cli._parse_size("210x297", "Paper width", "Paper height") == (210.0, 297.0)
    assert cli._parse_grid("2x3") == (2, 3)


def test_bad_values_exit_two_without_traceback(capsys, tmp_path, card_folder):
    folder = _cards(card_folder)
    out = tmp_path / "out"
    base = ["--cards", str(folder), "--out", str(out)]

    assert cli.main(base + ["--grid", "3.0x3"]) == 2
    first = capsys.readouterr()
    assert "error:" in first.err
    assert "Traceback" not in first.err

    assert cli.main(base + ["--card", "abc"]) == 2
    second = capsys.readouterr()
    assert "error:" in second.err
    assert "Traceback" not in second.err


def test_missing_out_exits_two(capsys, tmp_path, card_folder):
    folder = _cards(card_folder)

    code = cli.main(["--cards", str(folder)])

    assert code == 2
    assert "error:" in capsys.readouterr().err


def test_quiet_suppresses_progress(capsys, tmp_path, card_folder):
    folder = _cards(card_folder)
    base = ["--cards", str(folder), "--out", str(tmp_path / "out")]

    assert cli.main(base + ["--quiet"]) == 0
    assert capsys.readouterr().err == ""

    assert cli.main(base) == 0
    assert "Page" in capsys.readouterr().err


def test_version_prints_and_exits_zero(capsys):
    code = cli.main(["--version"])

    assert code == 0
    assert capsys.readouterr().out.startswith("pnp2pdf ")
