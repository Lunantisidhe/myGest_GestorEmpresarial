from recursos import gestor

__author__ = 'REBECA GONZÁLEZ BALADO'

global iva_ventas, iva_compras


# lee las ventas procesadas y calcula el iva cobrado
def calcular_iva_ventas():
    try:
        global iva_ventas
        iva_ventas = gestor.suma_columna_procesados('../datos/facturas/ventas-procesadas.xlsx', 2, 'C')
        gestor.log('Se han calculado ' + str(round(iva_ventas, 2)) + '€ de IVA cobrado.')

    except FileNotFoundError:
        gestor.log('Error: No se ha encontrado el fichero de ventas procesadas.')
    except TypeError:
        gestor.log('Error: Los datos del fichero de ventas procesadas no son válidos.')


# lee las compras procesadas y calcula el iva pagado
def calcular_iva_compras():
    try:
        global iva_compras
        iva_compras = gestor.suma_columna_procesados('../datos/facturas/compras-procesadas.xlsx', 2, 'C')
        gestor.log('Se han calculado ' + str(round(iva_compras, 2)) + '€ de IVA pagado.')

    except FileNotFoundError:
        gestor.log('Error: No se ha encontrado el fichero de compras procesadas.')
    except TypeError:
        gestor.log('Error: Los datos del fichero de compras procesadas no son válidos.')


# graba los datos calculados en la declaracion del iva
def grabar_declaracion_iva():
    try:
        if (iva_ventas is not None) and (iva_compras is not None):
            gestor.grabar_informe_mensual('../datos/facturas/declaracion-iva.xlsx', 'c d e g h i k l m o p q',
                                          {2: iva_ventas, 3: iva_compras})
            gestor.log('Se ha grabado la declaración del IVA.')

        else:
            gestor.log('Error: Existen datos a grabar no válidos.')

    except FileNotFoundError:
        gestor.log('Error: No se ha encontrado el fichero de declaración del IVA.')
    except NameError:
        gestor.log('Error: Faltan datos para grabar el fichero de declaración del IVA.')


def main():
    calcular_iva_ventas()
    calcular_iva_compras()
    grabar_declaracion_iva()


if __name__ == '__main__':
    main()
