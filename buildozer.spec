[app]
title = Mi App Kivy
package.name = miakivy
package.domain = org.abdon
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,c,so
version = 0.1

# CORREGIDO: Quitamos python3==3.11 (dejamos python3) y quitamos miakivy (ya se incluye solo)
requirements = python3,kivy,openpyxl,et_xmlfile,pillow,requests,urllib3,idna,certifi,charset-normalizer==3.3.2

p4a.branch = master
p4a.local_recipes = p4a-recipes
orientation = portrait

# CORREGIDO: Dejamos solo arm64-v8a para evitar errores de compilación cruzada en 32 bits
android.archs = arm64-v8a

android.api = 33
android.minapi = 24
android.ndk = 25b
android.skip_update = False

[buildozer]
log_level = 2
warn_on_root = 1
android.accept_sdk_license = True