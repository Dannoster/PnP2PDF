def test_missing_file_returns_defaults(pnp, tmp_path):
    result = pnp.load_save(str(tmp_path / "missing.json"))

    assert result == pnp.DEFAULT_SETTINGS


def test_partial_file_is_merged(pnp, tmp_path):
    path = tmp_path / "settings.json"
    path.write_text('{"card_width": 70}', encoding="utf-8")

    result = pnp.load_save(str(path))

    assert result["card_width"] == 70
    for key in pnp.DEFAULT_SETTINGS:
        assert key in result


def test_unknown_keys_ignored(pnp, tmp_path):
    path = tmp_path / "settings.json"
    path.write_text('{"dpi": 600}', encoding="utf-8")

    result = pnp.load_save(str(path))

    assert "dpi" not in result


def test_corrupt_file_returns_defaults(pnp, tmp_path):
    path = tmp_path / "settings.json"
    path.write_text("not json{{", encoding="utf-8")

    result = pnp.load_save(str(path))

    assert result == pnp.DEFAULT_SETTINGS


def test_roundtrip(pnp, tmp_path):
    path = tmp_path / "settings.json"
    settings = {"rows": 4, "columns": 2, "card_width": 70.0}

    pnp.create_save(str(path), settings)

    result = pnp.load_save(str(path))
    for key, value in settings.items():
        assert result[key] == value
