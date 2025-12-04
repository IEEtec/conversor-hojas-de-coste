from __future__ import annotations

from os import listdir
import re

from openpyxl import Workbook, load_workbook
from pkg_resources import working_set

from parse_file import ParseFile
from join_ot import juntar_ot_en_workbook

from cyclopts import App, Parameter
from typing import Annotated


app = App(
    name="Convesor OT",
    version="1.0.0",
    help='Convierte hojas de coste de proyectos en formato específico a otro formato. El uso normal es ejecutar el comando "todos" en la carpeta donde se encuentran los archivos a convertir. \n\nUsar "todos -h" para ver la ayuda de dicho comando.',
)


@app.command(name="todos")
def convierte_ficheros_directorio_actual(
    errores: Annotated[
        bool,
        Parameter(name=["--errores", "-e"]),
    ] = False,
    verboso: Annotated[
        bool,
        Parameter(name=["--verboso", "-v"]),
    ] = False,
) -> None:
    """Lee todos los archivos de la carpeta actual que sigan el formato "xxxxxxventa.txt" y "xxxxxxcoste.txt" y los convierte en csv para poder verlos en Excel.

    Parameters
    ----------
    errores: bool
        Muestra todos los errores, no solo los importantes.
    verboso: bool
        Muestra información detallada del proceso.
    """

    # Recupera todos los archivos de la carpeta actual

    ficheros: list[str] = listdir(".")

    # Guardamos que ficheros ya han sido ya meditos en la lista de procesar con una array de booleans

    ficheros_procesados = [False] * len(ficheros)

    print(ficheros)

    class FicherosProyecto:
        def __init__(self, archivo_venta: str, archivo_coste: str) -> None:
            if len(archivo_venta) > 6:
                self.proyecto = archivo_venta[:6]
            elif len(archivo_coste) > 6:
                self.proyecto = archivo_coste[:6]
            else:
                self.proyecto = ""

            self.archivo_venta = archivo_venta
            self.archivo_coste = archivo_coste

    archivos_a_procesar: dict[str, FicherosProyecto] = {}

    # Busca archivos con el formato "xxxxxxventa.txt", donde x son dígitos

    for i, fichero in enumerate(ficheros):
        if (
            len(fichero) == 15
            and fichero.endswith(".txt")
            and fichero[:6].isdigit()
            and fichero[6:11] == "venta"
        ):
            proyecto = FicherosProyecto(fichero, "")
            archivos_a_procesar[proyecto.proyecto] = proyecto
            ficheros_procesados[i] = True

    # Busca archivos con el formato "xxxxxxcoste.txt",
    for fichero in ficheros:
        if (
            len(fichero) == 15
            and fichero.endswith(".txt")
            and fichero[:6].isdigit()
            and fichero[6:11] == "coste"
        ):
            proyecto_id = fichero[:6]
            if proyecto_id in archivos_a_procesar:
                archivos_a_procesar[proyecto_id].archivo_coste = fichero
            else:
                proyecto = FicherosProyecto("", fichero)
                archivos_a_procesar[proyecto.proyecto] = proyecto

    # Muestra los archivos encontrados para cada proyecto y pide confirmación para procesarlos

    for proyecto_id, archivos in archivos_a_procesar.items():
        print(f"Proyecto: {proyecto_id}")
        print(f"  Archivo de venta: {archivos.archivo_venta}")
        print(f"  Archivo de coste: {archivos.archivo_coste}")

    confirmar = input("¿Desea procesar estos archivos? (s/n): ")
    if confirmar.lower() != "s":
        print("Proceso cancelado por el usuario.")
        exit()

    # Procesa los archivos confirmados

    for proyecto_id, archivos in archivos_a_procesar.items():
        print(f"Procesando proyecto {proyecto_id}...")
        ParseFile(
            archivos.archivo_coste,
            archivos.archivo_venta,
            proyecto_id,
            errores,
            verboso,
        )
        print(f"Proyecto {proyecto_id} procesado.")

    print("Procesamiento completado.")


@app.command(name="juntar")
def juntar_ficheros_ot(
    errores: Annotated[
        bool,
        Parameter(name=["--errores", "-e"]),
    ] = False,
    verboso: Annotated[
        bool,
        Parameter(name=["--verboso", "-v"]),
    ] = False,
) -> None:
    """Lee todos los archivos de la carpeta actual que sean una ot y los junta en un solo archivo.

    Parameters
    ----------
    errores: bool
        Muestra todos los errores, no solo los importantes.
    verboso: bool
        Muestra información detallada del proceso.
    """

    todos_los_ficheros: list[str] = listdir(".")

    # Los ficheros de OT tienen el nombre xxxxxx_de ddmmaaaa_a_ddmmaaaa.xlsx

    ficheros_de_ot: list[str] = []

    for fichero in todos_los_ficheros:
        match_ot = re.match(r"^(\d{6})_de_(\d{8})_a_(\d{8})\.xlsx$", fichero)
        if match_ot:
            ficheros_de_ot.append(fichero)
            print(f"Fichero OT encontrado: {fichero}")

    # Use the first file  as the base for the output

    wb = load_workbook(ficheros_de_ot.pop(0))

    # Crea un nuevo fichero

    ws_materiales = wb["Materiales"]

    if ws_materiales is None:
        raise ValueError("No se pudo crear la hoja de cálculo.")

    ws_mano_obra = wb["Mano de Obra"]

    if ws_mano_obra is None:
        raise ValueError("No se pudo crear la hoja de cálculo.")

    for fichero_ot in ficheros_de_ot:
        print(f"Procesando fichero OT: {fichero_ot}")
        juntar_ot_en_workbook(
            ws_materiales,
            ws_mano_obra,
            fichero_ot,
            errores,
            verboso,
        )

    nombre_fichero_salida = "OT_juntadas.xlsx"
    wb.save(nombre_fichero_salida)

    print(f"Fichero OT juntado guardado como: {nombre_fichero_salida}")


if __name__ == "__main__":
    app()
