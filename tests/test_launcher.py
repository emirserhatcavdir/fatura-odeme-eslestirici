import os
import shutil
import socket
import subprocess
import sys
from unittest.mock import patch

import pytest

import baslat


def test_occupied_port_keeps_existing_listener_and_shows_alternative(capsys):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as listener:
        listener.bind(("127.0.0.1", 0))
        listener.listen()
        port = listener.getsockname()[1]
        with patch("baslat.subprocess.run") as runner:
            assert baslat.main([str(port)]) == 1
        runner.assert_not_called()
        assert listener.getsockname()[1] == port
        output = capsys.readouterr().out
        assert "Hiçbir süreç sonlandırılmadı" in output
        assert ".\\baslat.bat 8502" in output


def test_available_alternative_port_uses_same_python_and_loopback_only():
    with patch("baslat.port_available", return_value=True), patch("baslat.subprocess.run") as runner:
        runner.return_value.returncode = 0
        assert baslat.main(["8502"]) == 0
    command = runner.call_args.args[0]
    assert command[:4] == [sys.executable, "-m", "streamlit", "run"]
    assert command[4] == str(baslat.ROOT / "sap.py")
    assert command[command.index("--server.address") + 1] == "127.0.0.1"
    assert command[command.index("--server.port") + 1] == "8502"
    assert runner.call_args.kwargs["cwd"] == baslat.ROOT


def test_missing_streamlit_returns_error_without_installing(capsys):
    with patch("baslat.find_spec", return_value=None), patch("baslat.subprocess.run") as runner:
        assert baslat.main([]) == 1
    runner.assert_not_called()
    assert "Streamlit bulunamadı" in capsys.readouterr().out


@pytest.mark.parametrize("port", ["abc", "0", "65536"])
def test_invalid_port_rejected_before_launch(port):
    with patch("baslat.subprocess.run") as runner, pytest.raises(SystemExit) as error:
        baslat.main([port])
    assert error.value.code == 2
    runner.assert_not_called()


@pytest.mark.skipif(sys.platform != "win32", reason="Gerçek Windows .bat davranışı")
def test_batch_missing_venv_from_other_directory_shows_message_and_pause(tmp_path):
    project = tmp_path / "proje bosluk ve & isaret"
    project.mkdir()
    batch = project / "baslat.bat"
    shutil.copyfile(baslat.ROOT / "baslat.bat", batch)
    result = subprocess.run(
        f'"{os.environ.get("COMSPEC", "cmd.exe")}" /d /s /c ""{batch}""',
        cwd=tmp_path, input="\n", capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=15,
    )
    assert result.returncode == 1
    assert "Sanal ortam bulunamadi" in result.stdout
    assert "README.md" in result.stdout
    assert "pause" in batch.read_text(encoding="utf-8").lower()


def test_process_start_failure_is_reported(capsys):
    with patch("baslat.port_available", return_value=True), patch("baslat.subprocess.run", side_effect=OSError("test failure")):
        assert baslat.main(["8502"]) == 1
    assert "Sunucu başlatılamadı" in capsys.readouterr().out
