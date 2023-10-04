import datetime
from openpyxl import load_workbook
from recursos import gestor
from recursos.clases import Factura

__author__ = 'REBECA GONZÁLEZ BALADO'

lista_facturas = []


# carga los datos de la hoja de compras a una lista de facturas
def leer_hoja_compras():
    try:
        libro = load_workbook(filename='../datos/facturas/compras.xlsx', data_only=True)
        hoja = libro.active

        fila = 1
        while True:
            # leemos los valores de la hoja
            fecha = hoja['A' + str(fila)].value
            iva_pagado = hoja['F' + str(fila)].value
            total_sin_iva = hoja['E' + str(fila)].value

            # si la fila esta vacia, paramos la lectura
            if (fecha is None) and (iva_pagado == 0) and (total_sin_iva is None):
                break

            # si los valores de la fila son correctos, la grabamos en la lista y cambiamos de fila
            if (type(fecha) is datetime.datetime) and ((type(iva_pagado) is int) or (type(iva_pagado) is float)) \
                    and ((type(total_sin_iva) is int) or (type(total_sin_iva) is float)):
                lista_facturas.append(Factura(fecha, iva_pagado, total_sin_iva))
            fila += 1

        gestor.log('Se han leído ' + str(len(lista_facturas)) + ' facturas de compra.')

    except FileNotFoundError:
        gestor.log('Error: No se ha encontrado el fichero de compras.')
    except ValueError:
        gestor.log('Error: Los datos del fichero de compras no son válidos.')


def main():
    leer_hoja_compras()
    gestor.generar_facturas_procesadas(lista_facturas, 'compras')


if __name__ == '__main__':
    main()
