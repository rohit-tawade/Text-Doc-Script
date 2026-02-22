import re
from pathlib import Path
import shutil

from kivy.app import App
from kivy.clock import Clock
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.utils import platform

from converter import convert_text_to_pdf, parse_resume


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


def sanitize_foldername(folder_name):
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", (folder_name or "").strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned or "Generated Resumes"


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

    def on_start(self):
        # Android dangerous permissions are requested at runtime (not install time).
        self.set_status("Checking permissions...")
        self.ensure_storage_permission(self._on_initial_permission_check)

    def _on_initial_permission_check(self, granted):
        if platform != "android":
            self.set_status("Ready")
            return
        if granted:
            self.set_status("Ready")
        else:
            self.set_status("Storage permission denied. You can still retry on Generate PDF.")

    def set_status(self, message):
        message_text = str(message).replace("\n", " ").strip()
        if len(message_text) > 180:
            message_text = message_text[:177] + "..."
        self.status_label.text = f"Status: {message_text}"

    def get_download_directory(self):
        if platform == "android":
            try:
                from android.storage import primary_external_storage_path

                return Path(primary_external_storage_path()) / "Download"
            except Exception:
                pass
            return ANDROID_DOWNLOAD_DIR
        return Path.home() / "Downloads"

    def extract_candidate_folder_name(self, resume_text):
        try:
            data = parse_resume(str(resume_text).splitlines())
            candidate_name = (data.get("name") or "").strip()
        except Exception:
            candidate_name = ""
        if not candidate_name:
            for line in str(resume_text).splitlines():
                stripped = line.strip()
                if stripped:
                    candidate_name = stripped
                    break
        return sanitize_foldername(candidate_name or "Generated Resumes")

    def get_android_sdk_int(self):
        if platform != "android":
            return 0
        try:
            from jnius import autoclass

            return int(autoclass("android.os.Build$VERSION").SDK_INT)
        except Exception:
            return 0

    def save_pdf_to_android_downloads(self, source_pdf_path, filename, folder_name=None):
        """Save a generated PDF into the public Downloads collection using MediaStore."""
        folder_name = sanitize_foldername(folder_name or "")
        if platform != "android":
            target = self.get_download_directory() / folder_name / filename
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(str(source_pdf_path), str(target))
            return target

        from jnius import autoclass

        PythonActivity = autoclass("org.kivy.android.PythonActivity")
        ContentValues = autoclass("android.content.ContentValues")
        Integer = autoclass("java.lang.Integer")
        MediaStoreDownloads = autoclass("android.provider.MediaStore$Downloads")
        MediaColumns = autoclass("android.provider.MediaStore$MediaColumns")

        activity = PythonActivity.mActivity
        resolver = activity.getContentResolver()

        def content_values_put(values_obj, key, value):
            """Box Python primitives so pyjnius resolves ContentValues.put() overloads."""
            if isinstance(value, bool):
                # pyjnius can map Python bool directly to java.lang.Boolean overload
                values_obj.put(key, bool(value))
            elif isinstance(value, int):
                values_obj.put(key, Integer.valueOf(int(value)))
            elif isinstance(value, float):
                values_obj.put(key, float(value))
            else:
                values_obj.put(key, value)

        values = ContentValues()
        content_values_put(values, MediaColumns.DISPLAY_NAME, filename)
        content_values_put(values, MediaColumns.MIME_TYPE, "application/pdf")
        if self.get_android_sdk_int() >= 29:
            content_values_put(values, MediaColumns.RELATIVE_PATH, f"Download/{folder_name}")
            content_values_put(values, MediaColumns.IS_PENDING, 1)

        uri = resolver.insert(MediaStoreDownloads.EXTERNAL_CONTENT_URI, values)
        if uri is None:
            raise RuntimeError("Could not create file in Android Downloads")

        output_stream = None
        try:
            output_stream = resolver.openOutputStream(uri)
            if output_stream is None:
                raise RuntimeError("Could not open Android Downloads output stream")
            with open(source_pdf_path, "rb") as pdf_file:
                while True:
                    chunk = pdf_file.read(65536)
                    if not chunk:
                        break
                    try:
                        output_stream.write(chunk)
                    except Exception:
                        output_stream.write(bytearray(chunk))
            output_stream.flush()
        finally:
            if output_stream is not None:
                output_stream.close()

        if self.get_android_sdk_int() >= 29:
            done_values = ContentValues()
            content_values_put(done_values, MediaColumns.IS_PENDING, 0)
            resolver.update(uri, done_values, None, None)

        return ANDROID_DOWNLOAD_DIR / folder_name / filename

    def get_app_pdf_temp_dir(self):
        temp_dir = Path(self.user_data_dir) / "generated_pdfs"
        temp_dir.mkdir(parents=True, exist_ok=True)
        return temp_dir

    def on_generate_pdf(self, _instance):
        resume_text = self.resume_input.text.strip()
        if not resume_text:
            self.set_status("Failed - Resume text is empty.")
            return

        filename = sanitize_filename(self.filename_input.text)
        folder_name = self.extract_candidate_folder_name(resume_text)
        self.set_status("Checking storage permission...")
        self.ensure_storage_permission(
            lambda granted: self._after_permission(granted, resume_text, filename, folder_name)
        )

    def _after_permission(self, granted, resume_text, filename, folder_name):
        if not granted:
            self.set_status("Failed - Storage permission denied.")
            return
        self.set_status("Generating PDF...")
        Clock.schedule_once(lambda _dt: self.generate_pdf(resume_text, filename, folder_name), 0)

    def generate_pdf(self, resume_text, filename, folder_name):
        try:
            if platform == "android" and self.get_android_sdk_int() >= 29:
                temp_folder = self.get_app_pdf_temp_dir() / folder_name
                temp_folder.mkdir(parents=True, exist_ok=True)
                temp_pdf_path = temp_folder / filename
                generated_pdf = convert_text_to_pdf(resume_text, temp_pdf_path, source_hint=filename)
                result = self.save_pdf_to_android_downloads(generated_pdf, filename, folder_name=folder_name)
            else:
                target_dir = self.get_download_directory() / folder_name
                target_path = target_dir / filename
                target_dir.mkdir(parents=True, exist_ok=True)
                result = convert_text_to_pdf(resume_text, target_path, source_hint=filename)
            self.set_status(f"Success - Saved to {result}")
        except Exception as exc:
            print(f"PDF generation/save failed: {exc}")
            self.set_status(f"Failed - {exc}")

    def ensure_storage_permission(self, callback):
        if platform != "android":
            callback(True)
            return

        # Android 10+ can write to the Downloads collection via MediaStore without
        # legacy READ/WRITE external storage runtime permissions.
        if self.get_android_sdk_int() >= 29:
            callback(True)
            return

        try:
            from android.permissions import Permission, check_permission, request_permissions
        except Exception:
            callback(True)
            return

        required_permissions = [Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE]
        permissions = list(required_permissions)
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
