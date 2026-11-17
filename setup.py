from setuptools import setup, Extension

setup(
    name="miakivy",
    version="0.1",
    ext_modules=[
        Extension("actividades", ["actividades.c"]),
        Extension("diagnostico", ["diagnostico.c"]),
        Extension("encabezado4", ["encabezado4.c"]),
        Extension("generador_rp66_nuevo", ["generador_rp66_nuevo.c"]),
        Extension("suscripcion", ["suscripcion.c"]),
        Extension("vistas_diagnostico", ["vistas_diagnostico.c"]),
    ],
)