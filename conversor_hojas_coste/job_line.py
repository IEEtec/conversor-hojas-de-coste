from __future__ import annotations
import datetime


class JobLine:
    def __init__(
        self,
        line_number: int,
    ):
        self.line_number = line_number
        self.OperacionId = ""
        self.Operacion = ""
        self.FechaStr = ""
        self.Fecha: datetime.date | None = None
        self.OperarioId = ""
        self.OperarioNombre = ""
        self.CantidadStr = ""
        self.Cantidad: float = 0.0
        self.PrecioUnitarioCosteStr = ""
        self.PrecioUnitarioCoste: float = 0.0
        self.DietasCosteStr = ""
        self.DietasCoste: float = 0.0
        self.DesplazamientoCosteStr = ""
        self.DesplazamientoCoste: float = 0.0
        self.ImporteTotalCosteStr = ""
        self.ImporteTotalCoste: float = 0.0
        self.PrecioUnitarioVentaStr = ""
        self.PrecioUnitarioVenta: float = 0.0
        self.DietasVentaStr = ""
        self.DietasVenta: float = 0.0
        self.DesplazamientoVentaStr = ""
        self.DesplazamientoVenta: float = 0.0
        self.ImporteTotalVentaStr = ""
        self.ImporteTotalVenta: float = 0.0

    def __str__(self):
        return f"{self.Operacion} | {self.FechaStr} | {self.OperarioId} | {self.OperarioNombre} | {self.CantidadStr} | {self.PrecioUnitarioCosteStr} | {self.DietasCosteStr} | {self.DesplazamientoCosteStr} | {self.DietasVentaStr} | {self.DesplazamientoVentaStr} | {self.ImporteTotalCosteStr} | {self.ImporteTotalVentaStr}"

    class NotValidError(Exception):
        def __init__(self, message: str):
            super().__init__(message)

    class IncongruentError(Exception):
        def __init__(self, message: str):
            super().__init__(message)

    @staticmethod
    def parse(
        line_guide: str,
        line_cost: str,
        line_sell: str,
        line_number: int,
        use_coste: bool,
        use_venta: bool,
    ) -> JobLine:

        # Extract the fields from the line
        # Operación                     Fecha      Operario                        Cant.  Precio   Dietas Despl.    Importe

        #   1 ARMADO DE CUADRO DE DINA 20/08/2024  109 GARCIA RODRIGUEZ, KEVIN     3,00   30,00     0,00   0,00      90,00

        job_line = JobLine(line_number)

        # Los campos tienen un ancho de caracteres constante, asi que podemos usar slicing para extraer los campos:
        OPERACION_ID_END = 4
        OPERACION_END = 29
        FECHA_END = 40
        OPERARIO_ID_END = 45
        OPERARIO_NOMBRE_END = 70
        CANTIDAD_END = 78
        PRECIO_END = 86
        DIETRAS_END = 95
        DESPLAZAMIENTO_END = 102
        IMPORTE_END = 113

        job_line.OperacionId = line_guide[0:OPERACION_ID_END].strip()
        job_line.Operacion = line_guide[OPERACION_ID_END:OPERACION_END].strip()
        job_line.FechaStr = line_guide[OPERACION_END:FECHA_END].strip()
        job_line.OperarioId = line_guide[FECHA_END:OPERARIO_ID_END].strip()
        job_line.OperarioNombre = line_guide[
            OPERARIO_ID_END:OPERARIO_NOMBRE_END
        ].strip()
        job_line.CantidadStr = line_guide[
            OPERARIO_NOMBRE_END:CANTIDAD_END
        ].strip()
        job_line.PrecioUnitarioCosteStr = line_cost[
            CANTIDAD_END:PRECIO_END
        ].strip()
        job_line.DietasCosteStr = line_guide[PRECIO_END:DIETRAS_END].strip()
        job_line.DesplazamientoCosteStr = line_guide[
            DIETRAS_END:DESPLAZAMIENTO_END
        ].strip()
        job_line.ImporteTotalCosteStr = line_cost[
            DESPLAZAMIENTO_END:IMPORTE_END
        ].strip()
        job_line.PrecioUnitarioVentaStr = line_sell[
            CANTIDAD_END:PRECIO_END
        ].strip()
        job_line.DietasVentaStr = line_sell[PRECIO_END:DIETRAS_END].strip()
        job_line.DesplazamientoVentaStr = line_sell[
            DIETRAS_END:DESPLAZAMIENTO_END
        ].strip()
        job_line.ImporteTotalVentaStr = line_sell[
            DESPLAZAMIENTO_END:IMPORTE_END
        ].strip()

        # Checkear que los campos tienen un formato correcto, ya sean numeros o fechas

        if job_line.FechaStr == "":
            raise JobLine.NotValidError(
                f"Error Linea invalida - Fecha vacia en linea {line_number + 1}: {job_line.FechaStr}"
            )

        try:
            job_line.Fecha = datetime.datetime.strptime(
                job_line.FechaStr, "%d/%m/%Y"
            ).date()
        except ValueError:
            raise JobLine.NotValidError(
                f"Error Linea invalida - Fecha con formato incorrecto en linea {line_number + 1}: {job_line.FechaStr}"
            )

        if job_line.CantidadStr == "":
            raise JobLine.NotValidError(
                f"Error Linea invalida - Cantidad vacia en linea {line_number + 1}: {job_line.CantidadStr}"
            )

        try:
            job_line.Cantidad = float(
                job_line.CantidadStr.replace(".", "").replace(",", ".")
            )
        except ValueError:
            raise JobLine.NotValidError(
                f"Error Linea invalida - Cantidad con formato incorrecto en linea {line_number + 1}: {job_line.CantidadStr}"
            )

        if use_coste:
            if job_line.PrecioUnitarioCosteStr == "":
                raise JobLine.NotValidError(
                    f"Error Linea invalida - PrecioUnitarioCoste vacio en linea {line_number + 1}: {job_line.PrecioUnitarioCosteStr}"
                )
            try:
                job_line.PrecioUnitarioCoste = float(
                    job_line.PrecioUnitarioCosteStr.replace(".", "").replace(
                        ",", "."
                    )
                )
            except ValueError:
                raise JobLine.NotValidError(
                    f"Error Linea invalida - PrecioUnitarioCoste con formato incorrecto en linea {line_number + 1}: {job_line.PrecioUnitarioCosteStr}"
                )

            if job_line.ImporteTotalCosteStr == "":
                raise JobLine.NotValidError(
                    f"Error Linea invalida - ImporteTotalCoste vacio en linea {line_number + 1}: {job_line.ImporteTotalCosteStr}"
                )
            try:
                job_line.ImporteTotalCoste = float(
                    job_line.ImporteTotalCosteStr.replace(".", "").replace(
                        ",", "."
                    )
                )
            except ValueError:
                raise JobLine.NotValidError(
                    f"Error Linea invalida - ImporteTotalCoste con formato incorrecto en linea {line_number + 1}: {job_line.ImporteTotalCosteStr}"
                )

            if job_line.DietasCosteStr == "":
                raise JobLine.NotValidError(
                    f"Error Linea invalida - DietasCoste vacio en linea {line_number + 1}: {job_line.DietasCosteStr}"
                )
            try:
                job_line.DietasCoste = float(
                    job_line.DietasCosteStr.replace(".", "").replace(",", ".")
                )
            except ValueError:
                raise JobLine.NotValidError(
                    f"Error Linea invalida - DietasCoste con formato incorrecto en linea {line_number + 1}: {job_line.DietasCosteStr}"
                )

            if job_line.DesplazamientoCosteStr == "":
                raise JobLine.NotValidError(
                    f"Error Linea invalida - DesplazamientoCoste vacio en linea {line_number + 1}: {job_line.DesplazamientoCosteStr}"
                )
            try:
                job_line.DesplazamientoCoste = float(
                    job_line.DesplazamientoCosteStr.replace(".", "").replace(
                        ",", "."
                    )
                )
            except ValueError:
                raise JobLine.NotValidError(
                    f"Error Linea invalida - DesplazamientoCoste con formato incorrecto en linea {line_number + 1}: {job_line.DesplazamientoCosteStr}"
                )

        if use_venta:
            if job_line.PrecioUnitarioVentaStr == "":
                raise JobLine.NotValidError(
                    f"Error Linea invalida - PrecioUnitarioVenta vacio en linea {line_number + 1}: {job_line.PrecioUnitarioVentaStr}"
                )
            try:
                job_line.PrecioUnitarioVenta = float(
                    job_line.PrecioUnitarioVentaStr.replace(".", "").replace(
                        ",", "."
                    )
                )
            except ValueError:
                raise JobLine.NotValidError(
                    f"Error Linea invalida - PrecioUnitarioVenta con formato incorrecto en linea {line_number + 1}: {job_line.PrecioUnitarioVentaStr}"
                )

            if job_line.ImporteTotalVentaStr == "":
                raise JobLine.NotValidError(
                    f"Error Linea invalida - ImporteTotalVenta vacio en linea {line_number + 1}: {job_line.ImporteTotalVentaStr}"
                )
            try:
                job_line.ImporteTotalVenta = float(
                    job_line.ImporteTotalVentaStr.replace(".", "").replace(
                        ",", "."
                    )
                )
            except ValueError:
                raise JobLine.NotValidError(
                    f"Error Linea invalida - ImporteTotalVenta con formato incorrecto en linea {line_number + 1}: {job_line.ImporteTotalVentaStr}"
                )

            if job_line.DietasVentaStr == "":
                raise JobLine.NotValidError(
                    f"Error Linea invalida - DietasVenta vacio en linea {line_number + 1}: {job_line.DietasVentaStr}"
                )
            try:
                job_line.DietasVenta = float(
                    job_line.DietasVentaStr.replace(".", "").replace(",", ".")
                )
            except ValueError:
                raise JobLine.NotValidError(
                    f"Error Linea invalida - DietasVenta con formato incorrecto en linea {line_number + 1}: {job_line.DietasVentaStr}"
                )

            if job_line.DesplazamientoVentaStr == "":
                raise JobLine.NotValidError(
                    f"Error Linea invalida - DesplazamientoVenta vacio en linea {line_number + 1}: {job_line.DesplazamientoVentaStr}"
                )
            try:
                job_line.DesplazamientoVenta = float(
                    job_line.DesplazamientoVentaStr.replace(".", "").replace(
                        ",", "."
                    )
                )
            except ValueError:
                raise JobLine.NotValidError(
                    f"Error Linea invalida - DesplazamientoVenta con formato incorrecto en linea {line_number + 1}: {job_line.DesplazamientoVentaStr}"
                )

        # Check si la linea de venta es consistente con la linea de coste (mismo referencia, descripcion, fecha, cantidad)

        if use_venta and use_coste:

            for field_name, cost_value, sell_value in [
                (
                    "OperacionId",
                    line_cost[0:OPERACION_ID_END].strip(),
                    line_sell[0:OPERACION_ID_END].strip(),
                ),
                (
                    "Operacion",
                    line_cost[OPERACION_ID_END:OPERACION_END].strip(),
                    line_sell[OPERACION_ID_END:OPERACION_END].strip(),
                ),
                (
                    "Fecha",
                    line_cost[OPERACION_END:FECHA_END].strip(),
                    line_sell[OPERACION_END:FECHA_END].strip(),
                ),
                (
                    "Operario_Id",
                    line_cost[FECHA_END:OPERARIO_ID_END].strip(),
                    line_sell[FECHA_END:OPERARIO_ID_END].strip(),
                ),
                (
                    "Operario_Nombre",
                    line_cost[OPERARIO_ID_END:OPERARIO_NOMBRE_END].strip(),
                    line_sell[OPERARIO_ID_END:OPERARIO_NOMBRE_END].strip(),
                ),
                (
                    "Cantidad",
                    line_cost[OPERARIO_NOMBRE_END:CANTIDAD_END].strip(),
                    line_sell[OPERARIO_NOMBRE_END:CANTIDAD_END].strip(),
                ),
            ]:
                if cost_value != sell_value:
                    raise JobLine.IncongruentError(
                        f"Error Incongruencia - {field_name} no congruente entre coste y venta en linea {line_number + 1}: {cost_value} != {sell_value}"
                    )

        # Return a JobLine object
        return job_line
