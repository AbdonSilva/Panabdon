from pythonforandroid.recipe import CExtensionRecipe

class MiaKivyRecipe(CExtensionRecipe):
    version = "0.1"
    url = ""
    name = "miakivy"
    # Módulos de C protegidos por Cython
    generated_sources = [
        "actividades.c",
        "diagnostico.c",
        "encabezado4.c",
        "generador_rp66_nuevo.c",
        "suscripcion.c",
        "vistas_diagnostico.c"
    ]
    call_hostpython_via_targetpython = False

recipe = MiaKivyRecipe()