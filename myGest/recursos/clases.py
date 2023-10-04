
# clase para el almacenaje de facturas de compra y venta
class Factura:
    def __init__(self, fecha, iva, total):
        self.fecha = fecha
        self.iva = iva
        self.total = total


# clase para el almacenaje de datos de empleados
class Empleado:
    def __init__(self, nombre, categoria, periodo_liquidacion, salario, incentivos, pagas_extra, numero_cuenta):
        self.nombre = nombre
        self.categoria = categoria
        self.periodo_liquidacion = periodo_liquidacion
        self.salario = salario
        self.incentivos = incentivos
        self.pagas_extra = pagas_extra
        self.numero_cuenta = numero_cuenta
