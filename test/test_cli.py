import pytest

from ymaps_excel_export.models import RunResult
from ymaps_excel_export.cli import main


def test_cli_exits_1_when_error(monkeypatch, capsys, st_base):
    import ymaps_excel_export.cli as cli_mod

    # Settings.from_env -> возвращаем st
    monkeypatch.setattr(cli_mod.Settings, "from_env", classmethod(lambda cls: st_base))

    # run -> возвращаем ошибку
    import ymaps_excel_export.pipeline as pipe_mod
    monkeypatch.setattr(cli_mod, "run", lambda st: RunResult(companies=[], request_meta={"error": "boom", "saved": None, "rows": 0}))

    with pytest.raises(SystemExit) as e:
        main()
    assert e.value.code == 1

    out = capsys.readouterr().out
    assert "ERROR:" in out
    assert "boom" in out


def test_cli_not_exits_when_ok(monkeypatch, capsys, st_base):
    import ymaps_excel_export.cli as cli_mod
    monkeypatch.setattr(cli_mod.Settings, "from_env", classmethod(lambda cls: st_base))
    monkeypatch.setattr(cli_mod, "run", lambda st: RunResult(companies=[], request_meta={"saved": "x.xlsx", "rows": 0}))

    # main() не должен падать
    main()
    out = capsys.readouterr().out
    assert "saved:" in out
