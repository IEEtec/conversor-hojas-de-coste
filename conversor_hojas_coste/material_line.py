from __future__ import annotations
from typing import Optional
import datetime


class MaterialLine:
    def __init__(self, numero_linea: int):
        self.line_number = numero_linea
        self.Referencia = ""
        self.Descripcion = ""
        self.FechaStr = ""
        self.Fecha: datetime.date | None = None
        self.CantidadStr = ""
        self.Cantidad: float = 0.0
        self.PrecioUnitarioCosteStr = ""
        self.PrecioUnitarioCoste: float = 0.0
        self.PrecioUnitarioVentaStr = ""
        self.PrecioUnitarioVenta: float = 0.0
        self.ImporteTotalCosteStr = ""
        self.ImporteTotalCoste: float = 0.0
        self.ImporteTotalVentaStr = ""
        self.ImporteTotalVenta: float = 0.0

    def __str__(self):
        return f"{self.Referencia} | {self.Descripcion} | {self.FechaStr} | {self.CantidadStr} | {self.PrecioUnitarioCosteStr} | {self.PrecioUnitarioVentaStr} | {self.ImporteTotalCosteStr} | {self.ImporteTotalVentaStr}"

    # Crear excepcion personalizada para errores de parseo, error de parseo o incongruencia entre coste y venta

    class NotValidError(Exception):
        def __init__(self, message: str):
            super().__init__(message)

    class IncongruentError(Exception):
        def __init__(self, message: str):
            super().__init__(message)

    @staticmethod
    def parse(
        linea_guia: str,
        linea_coste: str,
        linea_venta: str,
        numero_linea: int,
        use_coste: bool,
        use_venta: bool,
    ) -> Optional[MaterialLine]:

        # Extract the fields from the line
        # PHO3208100       BORNA CONEIXION PIT-1,5/S                03/05/2024       50,00           0,403           20,15

        material_line = MaterialLine(numero_linea)

        # The are constant character positions, so we can use slicing to extract the fields:
        REFERENCIA_END = 16
        DESCRIPCION_END = 57
        FECHA_END = 68
        CANTIDAD_END = 80
        PRECIO_END = 96
        IMPORTE_END = 112

        material_line.Referencia = linea_guia[0:REFERENCIA_END].strip()
        material_line.Descripcion = linea_guia[
            REFERENCIA_END:DESCRIPCION_END
        ].strip()
        material_line.FechaStr = linea_guia[DESCRIPCION_END:FECHA_END].strip()
        material_line.CantidadStr = linea_guia[FECHA_END:CANTIDAD_END].strip()
        if use_coste:
            material_line.PrecioUnitarioCosteStr = linea_coste[
                CANTIDAD_END:PRECIO_END
            ].strip()
            material_line.ImporteTotalCosteStr = linea_coste[
                PRECIO_END:IMPORTE_END
            ].strip()
        if use_venta:
            material_line.PrecioUnitarioVentaStr = linea_venta[
                CANTIDAD_END:PRECIO_END
            ].strip()
            material_line.ImporteTotalVentaStr = linea_venta[
                PRECIO_END:IMPORTE_END
            ].strip()

        # Checkear que los campos tienen un formato correcto, ya sean numeros o fechas

        if material_line.FechaStr == "":
            raise MaterialLine.NotValidError(
                f"Error Linea invalida - Fecha vacia en linea {numero_linea + 1}: {material_line.FechaStr}"
            )

        try:
            material_line.Fecha = datetime.datetime.strptime(
                material_line.FechaStr, "%d/%m/%Y"
            ).date()
        except ValueError:
            raise MaterialLine.NotValidError(
                f"Error Linea invalida - Fecha con formato incorrecto en linea {numero_linea + 1}: {material_line.FechaStr}"
            )

        if material_line.CantidadStr == "":
            raise MaterialLine.NotValidError(
                f"Error Linea invalida - Cantidad vacia en linea {numero_linea + 1}: {material_line.CantidadStr}"
            )

        try:
            material_line.Cantidad = float(
                material_line.CantidadStr.replace(".", "").replace(",", ".")
            )
        except ValueError:
            raise MaterialLine.NotValidError(
                f"Error Linea invalida - Cantidad con formato incorrecto en linea {numero_linea + 1}: {material_line.CantidadStr}"
            )

        if use_coste:
            if material_line.PrecioUnitarioCosteStr == "":
                raise MaterialLine.NotValidError(
                    f"Error Linea invalida - PrecioUnitarioCoste vacio en linea {numero_linea + 1}: {material_line.PrecioUnitarioCosteStr}"
                )
            try:
                material_line.PrecioUnitarioCoste = float(
                    material_line.PrecioUnitarioCosteStr.replace(
                        ".", ""
                    ).replace(",", ".")
                )
            except ValueError:
                raise MaterialLine.NotValidError(
                    f"Error Linea invalida - PrecioUnitarioCoste con formato incorrecto en linea {numero_linea + 1}: {material_line.PrecioUnitarioCosteStr}"
                )

            if material_line.ImporteTotalCosteStr == "":
                raise MaterialLine.NotValidError(
                    f"Error Linea invalida - ImporteTotalCoste vacio en linea {numero_linea + 1}: {material_line.ImporteTotalCosteStr}"
                )
            try:
                material_line.ImporteTotalCoste = float(
                    material_line.ImporteTotalCosteStr.replace(
                        ".", ""
                    ).replace(",", ".")
                )
            except ValueError:
                raise MaterialLine.NotValidError(
                    f"Error Linea invalida - ImporteTotalCoste con formato incorrecto en linea {numero_linea + 1}: {material_line.ImporteTotalCosteStr}"
                )

        if use_venta:
            if material_line.PrecioUnitarioVentaStr == "":
                raise MaterialLine.NotValidError(
                    f"Error Linea invalida - PrecioUnitarioVenta vacio en linea {numero_linea + 1}: {material_line.PrecioUnitarioVentaStr}"
                )
            try:
                material_line.PrecioUnitarioVenta = float(
                    material_line.PrecioUnitarioVentaStr.replace(
                        ".", ""
                    ).replace(",", ".")
                )
            except ValueError:
                raise MaterialLine.NotValidError(
                    f"Error Linea invalida - PrecioUnitarioVenta con formato incorrecto en linea {numero_linea + 1}: {material_line.PrecioUnitarioVentaStr}"
                )

            if material_line.ImporteTotalVentaStr == "":
                raise MaterialLine.NotValidError(
                    f"Error Linea invalida - ImporteTotalVenta vacio en linea {numero_linea + 1}: {material_line.ImporteTotalVentaStr}"
                )
            try:
                material_line.ImporteTotalVenta = float(
                    material_line.ImporteTotalVentaStr.replace(
                        ".", ""
                    ).replace(",", ".")
                )
            except ValueError:
                raise MaterialLine.NotValidError(
                    f"Error Linea invalida - ImporteTotalVenta con formato incorrecto en linea {numero_linea + 1}: {material_line.ImporteTotalVentaStr}"
                )

        # Check si la linea de venta es consistente con la linea de coste (mismo referencia, descripcion, fecha, cantidad)
        if use_venta and use_coste:

            for field_name, cost_value, sell_value in [
                (
                    "Referencia",
                    linea_coste[0:REFERENCIA_END].strip(),
                    linea_venta[0:REFERENCIA_END].strip(),
                ),
                (
                    "Descripcion",
                    linea_coste[REFERENCIA_END:DESCRIPCION_END].strip(),
                    linea_venta[REFERENCIA_END:DESCRIPCION_END].strip(),
                ),
                (
                    "Fecha",
                    linea_coste[DESCRIPCION_END:FECHA_END].strip(),
                    linea_venta[DESCRIPCION_END:FECHA_END].strip(),
                ),
                (
                    "Cantidad",
                    linea_coste[FECHA_END:CANTIDAD_END].strip(),
                    linea_venta[FECHA_END:CANTIDAD_END].strip(),
                ),
            ]:
                if cost_value != sell_value:
                    raise MaterialLine.IncongruentError(
                        f"Error Incongruencia - {field_name} incongruente en linea {numero_linea + 1}: coste '{cost_value}' vs venta '{sell_value}'"
                    )

        return material_line
