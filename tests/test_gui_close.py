import threading
import unittest
from contextlib import ExitStack
from unittest.mock import Mock, patch

import GUI as app


class CloseWindowTests(unittest.TestCase):
    def setUp(self):
        self.patches = ExitStack()
        self.addCleanup(self.patches.close)
        self.root = Mock()
        self.patches.enter_context(patch.object(app, "root", self.root, create=True))
        self.patches.enter_context(patch.object(app, "merge_thread", None, create=True))
        self.patches.enter_context(patch.object(app, "selected_files", ["a.xlsx", "b.xlsx"]))
        for name in ("merge_button", "select_button", "delete_button", "up_button", "down_button"):
            self.patches.enter_context(patch.object(app, name, Mock(), create=True))
        self.messages = self.patches.enter_context(patch.object(app, "messagebox"))
        self.dialogs = self.patches.enter_context(patch.object(app, "filedialog"))
        self.dialogs.asksaveasfilename.return_value = "merged.xlsx"
        self.patches.enter_context(patch.object(app, "log_message"))

    def test_close_before_any_merge_destroys_window_without_prompt(self):
        app.on_close()

        self.root.destroy.assert_called_once_with()
        self.assertEqual(self.messages.mock_calls, [])

    def check_worker_lifecycle(self, fail=False):
        entered = threading.Event()
        release = threading.Event()
        workers = []

        def merge_worker(*args):
            workers.append(threading.current_thread())
            entered.set()
            if not release.wait(5):
                raise TimeoutError("Test worker was not released")
            if fail:
                raise RuntimeError("Simulated worker failure")

        with patch.object(app.merge_logic, "merge_excel_files", side_effect=merge_worker), \
                patch.object(threading, "excepthook") as exception_hook:
            try:
                app.start_merge_thread()
                self.assertTrue(entered.wait(5), "Merge worker did not start")
                # Both repeated close attempts must leave the merge running.
                app.on_close()
                app.on_close()

                self.root.destroy.assert_not_called()
                self.assertEqual(self.messages.showwarning.call_count, 2)
                self.messages.showwarning.assert_called_with(
                    "병합 진행 중",
                    "아직 병합이 진행 중이므로 종료할 수 없습니다.\n"
                    "작업이 완료된 후 종료해 주세요.",
                    parent=self.root,
                )
            finally:
                release.set()
                for worker in workers:
                    worker.join(timeout=5)

            self.assertFalse(workers[0].is_alive())
            if fail:
                exception_hook.assert_called_once()
            else:
                exception_hook.assert_not_called()
            self.messages.reset_mock()

            # Finishing or failing must allow closing without a lingering busy flag.
            app.on_close()

        self.root.destroy.assert_called_once_with()
        self.assertEqual(self.messages.mock_calls, [])

    def test_close_is_blocked_during_merge_and_allowed_after_completion(self):
        self.check_worker_lifecycle()

    def test_close_is_allowed_after_worker_failure(self):
        self.check_worker_lifecycle(fail=True)

    def test_aborted_start_does_not_block_closing(self):
        cases = (
            ([], "merged.xlsx"),
            (["a.xlsx"], "merged.xlsx"),
            (["a.xlsx", "b.xlsx"], ""),
            (["a.xlsx", "b.xlsx"], "a.xlsx"),
        )
        for files, output in cases:
            with self.subTest(files=files, output=output), \
                    patch.object(app, "selected_files", files), \
                    patch.object(app.threading, "Thread") as thread_class:
                self.dialogs.asksaveasfilename.return_value = output
                self.root.reset_mock()
                app.start_merge_thread()
                thread_class.assert_not_called()
                self.messages.reset_mock()

                app.on_close()

                self.root.destroy.assert_called_once_with()
                self.assertEqual(self.messages.mock_calls, [])

    def test_window_close_event_is_connected_to_handler(self):
        with ExitStack() as widgets:
            widgets.enter_context(patch.object(app.tk, "Tk", return_value=self.root))
            for name in ("Frame", "Button", "Listbox", "Label"):
                widgets.enter_context(patch.object(app.tk, name))
            widgets.enter_context(patch.object(app.scrolledtext, "ScrolledText"))

            app.GUI()

        self.root.protocol.assert_called_once_with(
            "WM_DELETE_WINDOW", getattr(app, "on_close", None)
        )


if __name__ == "__main__":
    unittest.main()
