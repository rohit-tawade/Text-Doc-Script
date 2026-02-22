import re
from pathlib import Path

from kivy.app import App
from kivy.clock import Clock
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.utils import platform

from converter import convert_text_to_pdf


ANDROID_DOWNLOAD_DIR = Path("/storage/emulated/0/Download")
DEFAULT_FILENAME = "resume.pdf"


def sanitize_filename(filename):
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", (filename or "").strip())
    cleaned = re.sub(r"\s+", "_", cleaned).strip("._ ")
    if not cleaned:
        cleaned = "resume"
    if not cleaned.lower().endswith(".pdf"):
        cleaned = f"{cleaned}.pdf"
    return cleaned


class ResumePdfApp(App):
    def build(self):
        self.title = "Resume PDF Generator"

        root = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(8))

        info = Label(
            text="Paste resume text, enter a file name, then tap Generate PDF.",
            size_hint_y=None,
            height=dp(28),
            halign="left",
            valign="middle",
        )
        info.bind(size=lambda instance, _: setattr(instance, "text_size", (instance.width, None)))
        root.add_widget(info)

        self.resume_input = TextInput(
            hint_text="Paste resume text here...",
            multiline=True,
            size_hint=(1, 1),
        )
        root.add_widget(self.resume_input)

        filename_row = BoxLayout(orientation="horizontal", spacing=dp(8), size_hint_y=None, height=dp(44))
        filename_label = Label(text="Filename:", size_hint_x=None, width=dp(90), halign="left", valign="middle")
        filename_label.bind(size=lambda instance, _: setattr(instance, "text_size", (instance.width, None)))
        filename_row.add_widget(filename_label)
        self.filename_input = TextInput(text=DEFAULT_FILENAME, multiline=False)
        filename_row.add_widget(self.filename_input)
        root.add_widget(filename_row)

        buttons = BoxLayout(orientation="horizontal", spacing=dp(8), size_hint_y=None, height=dp(48))
        generate_btn = Button(text="Generate PDF")
        generate_btn.bind(on_release=self.on_generate_pdf)
        clear_btn = Button(text="Clear")
        clear_btn.bind(on_release=self.on_clear)
        buttons.add_widget(generate_btn)
        buttons.add_widget(clear_btn)
        root.add_widget(buttons)

        self.status_label = Label(
            text="Status: Ready",
            size_hint_y=None,
            height=dp(36),
            halign="left",
            valign="middle",
        )
        self.status_label.bind(size=lambda instance, _: setattr(instance, "text_size", (instance.width, None)))
        root.add_widget(self.status_label)

        return root

    def set_status(self, message):
        self.status_label.text = f"Status: {message}"

    def get_download_directory(self):
        if platform == "android":
            try:
                from android.storage import primary_external_storage_path

                return Path(primary_external_storage_path()) / "Download"
            except Exception:
                pass
            return ANDROID_DOWNLOAD_DIR
        return Path.home() / "Downloads"

    def on_generate_pdf(self, _instance):
        resume_text = self.resume_input.text.strip()
        if not resume_text:
            self.set_status("Failed - Resume text is empty.")
            return

        filename = sanitize_filename(self.filename_input.text)
        self.set_status("Checking storage permission...")
        self.ensure_storage_permission(
            lambda granted: self._after_permission(granted, resume_text, filename)
        )

    def _after_permission(self, granted, resume_text, filename):
        if not granted:
            self.set_status("Failed - Storage permission denied.")
            return
        self.set_status("Generating PDF...")
        Clock.schedule_once(lambda _dt: self.generate_pdf(resume_text, filename), 0)

    def generate_pdf(self, resume_text, filename):
        target_dir = self.get_download_directory()
        target_path = target_dir / filename
        try:
            target_dir.mkdir(parents=True, exist_ok=True)
            result = convert_text_to_pdf(resume_text, target_path, source_hint=filename)
            self.set_status(f"Success - Saved to {result}")
        except Exception as exc:
            self.set_status(f"Failed - {exc}")

    def ensure_storage_permission(self, callback):
        if platform != "android":
            callback(True)
            return

        try:
            from android.permissions import Permission, check_permission, request_permissions
        except Exception:
            callback(True)
            return

        required_permissions = [Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE]
        permissions = list(required_permissions)
        if hasattr(Permission, "MANAGE_EXTERNAL_STORAGE"):
            permissions.append(Permission.MANAGE_EXTERNAL_STORAGE)
        missing = [permission for permission in permissions if not check_permission(permission)]
        if not missing:
            callback(True)
            return

        def handle_permission_result(_permissions, grants):
            granted = all(check_permission(permission) for permission in required_permissions)
            Clock.schedule_once(lambda _dt: callback(granted), 0)

        request_permissions(missing, handle_permission_result)

    def on_clear(self, _instance):
        self.resume_input.text = ""
        self.filename_input.text = DEFAULT_FILENAME
        self.set_status("Cleared.")


if __name__ == "__main__":
    ResumePdfApp().run()
