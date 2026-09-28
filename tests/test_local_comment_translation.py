from __future__ import annotations

import importlib.util
import subprocess
import sys
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "local_comment_translation", ROOT / "scripts" / "local_comment_translation.py"
)
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class FakeProcess:
    def __init__(self, exit_code: int | None = None):
        self.exit_code = exit_code
        self.terminated = False
        self.killed = False

    def poll(self):
        return self.exit_code

    def terminate(self):
        self.terminated = True
        self.exit_code = 0

    def wait(self, timeout=None):
        return self.exit_code

    def kill(self):
        self.killed = True
        self.exit_code = 1


class LocalCommentTranslationPathTests(unittest.TestCase):
    def tearDown(self):
        MODULE._stop_server()

    def _lock_mocks(self):
        def acquire():
            MODULE._PATH_MUTEX = "held-mutex"

        def release():
            MODULE._PATH_MUTEX = None

        return mock.patch.object(MODULE, "_acquire_path_mutex", side_effect=acquire), \
            mock.patch.object(MODULE, "_release_path_mutex", side_effect=release)

    def _prepare_startup_aliases(self, server, model):
        MODULE._PATH_MAPPINGS[:] = ["Z:"]
        MODULE._PATH_MUTEX = "held-mutex"
        return Path("Z:/llama-server.exe"), Path("Y:/QwenPaw-Flash-2B-Q4_K_M.gguf")

    @staticmethod
    def _release_test_mutex():
        MODULE._PATH_MUTEX = None

    def test_unicode_runtime_and_model_directories_get_ascii_aliases(self):
        server = Path("D:/验证材料/llama/llama-server.exe")
        model = Path("D:/验证材料/模型/QwenPaw-Flash-2B-Q4_K_M.gguf")
        completed = subprocess.CompletedProcess([], 0, stdout="", stderr="")
        with mock.patch.object(MODULE.os, "name", "nt"), \
                mock.patch.object(MODULE, "_acquire_path_mutex"), \
                mock.patch.object(MODULE, "_get_logical_drives", return_value=0), \
                mock.patch.object(MODULE.subprocess, "run", return_value=completed) as run:
            alias_server, alias_model = MODULE._ascii_runtime_paths(server, model)

        self.assertEqual(str(alias_server), "Z:\\llama-server.exe")
        self.assertEqual(str(alias_model), "Y:\\QwenPaw-Flash-2B-Q4_K_M.gguf")
        self.assertTrue(str(alias_server).isascii())
        self.assertTrue(str(alias_model).isascii())
        self.assertEqual([call.args[0][1] for call in run.call_args_list], ["Z:", "Y:"])
        self.assertEqual(MODULE._PATH_MAPPINGS, ["Z:", "Y:"])

    def test_mapping_failure_releases_prior_mapping_and_mutex(self):
        server = Path("D:/验证材料/llama/llama-server.exe")
        model = Path("D:/验证材料/模型/QwenPaw-Flash-2B-Q4_K_M.gguf")
        succeeded = subprocess.CompletedProcess([], 0, stdout="", stderr="")
        failed = subprocess.CompletedProcess([], 1, stdout="busy", stderr="")

        def create_first_then_fail(*args, **kwargs):
            return succeeded if args[0][1] == "Z:" else failed

        used = sum(1 << (ord(letter) - ord("A")) for letter in "DEFGHIJKLMNOPQRSTUVWX")
        with mock.patch.object(MODULE.os, "name", "nt"), \
                mock.patch.object(MODULE, "_acquire_path_mutex"), \
                mock.patch.object(MODULE, "_get_logical_drives", return_value=used), \
                mock.patch.object(MODULE.subprocess, "run", side_effect=create_first_then_fail) as run:
            with self.assertRaisesRegex(RuntimeError, "临时路径映射"):
                MODULE._ascii_runtime_paths(server, model)

        self.assertEqual(MODULE._PATH_MAPPINGS, [])
        self.assertIn(["subst", "Z:", "/D"], [call.args[0] for call in run.call_args_list])

    def test_drive_enumeration_failure_releases_mutex(self):
        server = Path("D:/验证材料/llama/llama-server.exe")
        model = Path("D:/验证材料/模型/QwenPaw-Flash-2B-Q4_K_M.gguf")
        acquire, release = self._lock_mocks()
        with mock.patch.object(MODULE.os, "name", "nt"), acquire, release as release_mutex, \
                mock.patch.object(MODULE, "_get_logical_drives", side_effect=OSError("drive query failed")):
            with self.assertRaisesRegex(OSError, "drive query failed"):
                MODULE._ascii_runtime_paths(server, model)

        release_mutex.assert_called_once_with()
        self.assertIsNone(MODULE._PATH_MUTEX)
        self.assertEqual(MODULE._PATH_MAPPINGS, [])

    def test_subst_timeout_removes_prior_mapping_and_releases_mutex(self):
        server = Path("D:/验证材料/llama/llama-server.exe")
        model = Path("D:/验证材料/模型/QwenPaw-Flash-2B-Q4_K_M.gguf")
        succeeded = subprocess.CompletedProcess([], 0, stdout="", stderr="")
        used = sum(1 << (ord(letter) - ord("A")) for letter in "DEFGHIJKLMNOPQRSTUVWX")

        def map_first_then_timeout(*args, **kwargs):
            command = args[0]
            if command[1] == "Z:" and command[2] != "/D":
                return succeeded
            if command[1] == "Y:":
                raise subprocess.TimeoutExpired(command, 10)
            return succeeded

        acquire, release = self._lock_mocks()
        with mock.patch.object(MODULE.os, "name", "nt"), acquire, release as release_mutex, \
                mock.patch.object(MODULE, "_get_logical_drives", return_value=used), \
                mock.patch.object(MODULE.subprocess, "run", side_effect=map_first_then_timeout) as run:
            with self.assertRaises(subprocess.TimeoutExpired):
                MODULE._ascii_runtime_paths(server, model)

        release_mutex.assert_called_once_with()
        self.assertIn(["subst", "Z:", "/D"], [call.args[0] for call in run.call_args_list])
        self.assertIsNone(MODULE._PATH_MUTEX)
        self.assertEqual(MODULE._PATH_MAPPINGS, [])

    def test_subst_start_error_releases_mutex(self):
        server = Path("D:/验证材料/llama/llama-server.exe")
        model = Path("D:/验证材料/模型/QwenPaw-Flash-2B-Q4_K_M.gguf")
        used = sum(1 << (ord(letter) - ord("A")) for letter in "DEFGHIJKLMNOPQRSTUVWXY")
        acquire, release = self._lock_mocks()
        with mock.patch.object(MODULE.os, "name", "nt"), acquire, release as release_mutex, \
                mock.patch.object(MODULE, "_get_logical_drives", return_value=used), \
                mock.patch.object(MODULE.subprocess, "run", side_effect=OSError("subst could not start")):
            with self.assertRaisesRegex(OSError, "subst could not start"):
                MODULE._ascii_runtime_paths(server, model)

        release_mutex.assert_called_once_with()
        self.assertIsNone(MODULE._PATH_MUTEX)
        self.assertEqual(MODULE._PATH_MAPPINGS, [])

    def test_startup_arguments_keep_model_and_quality_configuration(self):
        server_path = Path("Z:/llama-server.exe")
        model_path = Path("Y:/QwenPaw-Flash-2B-Q4_K_M.gguf")
        process = FakeProcess()
        response = mock.MagicMock()
        response.__enter__.return_value.status = 200
        with mock.patch.object(MODULE, "_runtime", return_value=(Path("原目录/llama-server.exe"), Path("模型/QwenPaw-Flash-2B-Q4_K_M.gguf"))), \
                mock.patch.object(MODULE, "_ascii_runtime_paths", return_value=(server_path, model_path)), \
                mock.patch.object(MODULE.subprocess, "Popen", return_value=process) as popen, \
                mock.patch.object(MODULE.urllib.request, "urlopen", return_value=response):
            url = MODULE._ensure_server(Path("app"))

        args, kwargs = popen.call_args
        self.assertEqual(args[0][:3], [str(server_path), "--model", str(model_path)])
        self.assertEqual(kwargs["cwd"], str(server_path.parent))
        self.assertIn("--ctx-size", args[0])
        self.assertEqual(args[0][args[0].index("--ctx-size") + 1], "4096")
        self.assertEqual(args[0][args[0].index("--n-gpu-layers") + 1], "99")
        self.assertEqual(url.split("/")[2].split(":")[0], "127.0.0.1")

    def test_startup_failure_reports_server_stderr_and_cleans_mapping(self):
        failed_process = FakeProcess(exit_code=1)
        completed = subprocess.CompletedProcess([], 0, stdout="", stderr="")

        def start_process(*args, **kwargs):
            kwargs["stderr"].write(b"model path could not be opened")
            return failed_process

        with mock.patch.object(MODULE, "_runtime", return_value=(Path("server.exe"), Path("model.gguf"))), \
                mock.patch.object(MODULE, "_ascii_runtime_paths", return_value=(Path("Z:/server.exe"), Path("Z:/model.gguf"))), \
                mock.patch.object(MODULE.subprocess, "Popen", side_effect=start_process), \
                mock.patch.object(MODULE.subprocess, "run", return_value=completed) as run:
            MODULE._PATH_MAPPINGS[:] = ["Z:"]
            with self.assertRaisesRegex(RuntimeError, "model path could not be opened"):
                MODULE._ensure_server(Path("app"))

        self.assertIn(["subst", "Z:", "/D"], [call.args[0] for call in run.call_args_list])
        self.assertIsNone(MODULE._SERVER)
        self.assertIsNone(MODULE._SERVER_STDERR)

    def test_socket_bind_failure_releases_successful_path_mapping(self):
        server = Path("原目录/llama-server.exe")
        model = Path("模型/QwenPaw-Flash-2B-Q4_K_M.gguf")
        socket_context = mock.MagicMock()
        socket_object = socket_context.__enter__.return_value
        bind_error = OSError("port bind failed")
        socket_object.bind.side_effect = bind_error
        completed = subprocess.CompletedProcess([], 0, stdout="", stderr="")
        with mock.patch.object(MODULE, "_runtime", return_value=(server, model)), \
                mock.patch.object(MODULE, "_ascii_runtime_paths", side_effect=self._prepare_startup_aliases), \
                mock.patch.object(MODULE.socket, "socket", return_value=socket_context), \
                mock.patch.object(MODULE, "_release_path_mutex", side_effect=self._release_test_mutex) as release_mutex, \
                mock.patch.object(MODULE.subprocess, "run", return_value=completed) as run:
            with self.assertRaises(OSError) as raised:
                MODULE._ensure_server(Path("app"))

        self.assertIs(raised.exception, bind_error)
        self.assertEqual(release_mutex.call_count, 2)  # 初始清理 + 启动准备失败回滚
        self.assertIn(["subst", "Z:", "/D"], [call.args[0] for call in run.call_args_list])
        self.assertEqual(MODULE._PATH_MAPPINGS, [])
        self.assertIsNone(MODULE._PATH_MUTEX)
        self.assertIsNone(MODULE._SERVER_STDERR)

    def test_stderr_file_creation_failure_releases_successful_path_mapping(self):
        server = Path("原目录/llama-server.exe")
        model = Path("模型/QwenPaw-Flash-2B-Q4_K_M.gguf")
        socket_context = mock.MagicMock()
        socket_object = socket_context.__enter__.return_value
        socket_object.getsockname.return_value = ("127.0.0.1", 12345)
        file_error = OSError("temporary log creation failed")
        completed = subprocess.CompletedProcess([], 0, stdout="", stderr="")
        with mock.patch.object(MODULE, "_runtime", return_value=(server, model)), \
                mock.patch.object(MODULE, "_ascii_runtime_paths", side_effect=self._prepare_startup_aliases), \
                mock.patch.object(MODULE.socket, "socket", return_value=socket_context), \
                mock.patch.object(MODULE.tempfile, "TemporaryFile", side_effect=file_error), \
                mock.patch.object(MODULE, "_release_path_mutex", side_effect=self._release_test_mutex) as release_mutex, \
                mock.patch.object(MODULE.subprocess, "run", return_value=completed) as run:
            with self.assertRaises(OSError) as raised:
                MODULE._ensure_server(Path("app"))

        self.assertIs(raised.exception, file_error)
        self.assertEqual(release_mutex.call_count, 2)  # 初始清理 + 启动准备失败回滚
        self.assertIn(["subst", "Z:", "/D"], [call.args[0] for call in run.call_args_list])
        self.assertEqual(MODULE._PATH_MAPPINGS, [])
        self.assertIsNone(MODULE._PATH_MUTEX)
        self.assertIsNone(MODULE._SERVER_STDERR)

    def test_normal_stop_terminates_owned_server_and_releases_mapping(self):
        process = FakeProcess()
        completed = subprocess.CompletedProcess([], 0, stdout="", stderr="")
        MODULE._SERVER = process
        MODULE._BASE_URL = "http://127.0.0.1:1234"
        MODULE._PATH_MAPPINGS[:] = ["Z:"]
        with mock.patch.object(MODULE.subprocess, "run", return_value=completed) as run:
            MODULE._stop_server()

        self.assertTrue(process.terminated)
        self.assertEqual(run.call_args.args[0], ["subst", "Z:", "/D"])
        self.assertEqual(MODULE._PATH_MAPPINGS, [])
        self.assertIsNone(MODULE._BASE_URL)


if __name__ == "__main__":
    unittest.main()
