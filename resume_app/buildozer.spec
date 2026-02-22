[app]
title = Resume PDF Generator
package.name = resumepdf
package.domain = org.textdocscript
source.dir = .
source.include_exts = py,txt,png,jpg,jpeg,kv,atlas,jinja
source.exclude_dirs = .git,__pycache__,build,dist,.venv,venv,build-env
source.exclude_patterns = *.pyc,*.pyo,*.tmp
version = 1.0.3
requirements = python3,kivy,pyjnius,setuptools
orientation = portrait
fullscreen = 0

# Android
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 24
android.ndk = 25b
android.archs = arm64-v8a
android.allow_backup = True
android.enable_androidx = True
android.accept_sdk_license = True
android.add_manifest_application_attributes = android:requestLegacyExternalStorage="true"

# Keep app fully offline; no internet permission requested.
presplash.color = #FFFFFF

log_level = 2
warn_on_root = 1

[buildozer]
log_level = 2
warn_on_root = 1
