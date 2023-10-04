import datetime
from openpyxl import load_workbook
from recursos import gestor
from recursos.clases import Factura

__author__ = 'REBECA GONZÁLEZ BALADO'

lista_facturas = []


# carga los datos de las hojas de ventas a una lista de facturas
def leer_hojas_ventas():
    try:
        libro = load_workbook(filename='../datos/facturas/facturas-emitidas.xlsx', data_only=True)

        # leemos los valores de todas las hojas
        for hoja in libro.worksheets:
            fecha = hoja['H3'].value
            iva_cobrado = hoja['F48'].value
            total_sin_iva = hoja['F47'].value

            # si los valores de la fila son correctos, la grabamos en la lista
            if (type(fecha) is datetime.datetime) and ((type(iva_cobrado) is int) or (type(iva_cobrado) is float)) \
                    and ((type(total_sin_iva) is int) or (type(total_sin_iva) is float)):
                lista_facturas.append(Factura(fecha, iva_cobrado, total_sin_iva))

        gestor.log('Se han leído ' + str(len(lista_facturas)) + ' facturas de venta.')

    except FileNotFoundError:
        gestor.log('Error: No se ha encontrado el fichero de ventas.')
    except ValueError:
        gestor.log('Error: Los datos del fichero de ventas no son válidos.')


def main():
    leer_hojas_ventas()
    gestor.generar_facturas_procesadas(lista_facturas, 'ventas')


if __name__ == '__main__':
    main()
