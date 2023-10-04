import os
from recursos import gestor

__author__ = 'REBECA GONZÁLEZ BALADO'

global ingresos, gastos_personal, gastos_compra


# lee las ventas procesadas y calcula los ingresos
def calcular_ingresos():
    try:
        global ingresos
        ingresos = gestor.suma_columna_procesados('../datos/facturas/ventas-procesadas.xlsx', 2, 'B')

        if ingresos is not None:
            gestor.log('Se han calculado ' + str(ingresos) + '€ de ingresos.')
        else:
            gestor.log('Error: Datos de ventas procesadas no válidos.')

    except FileNotFoundError:
        gestor.log('Error: No se ha encontrado el fichero de ventas procesadas.')


# lee las compras procesadas y calcula los gastos de compra
def calcular_gastos_compra():
    try:
        global gastos_compra
        gastos_compra = gestor.suma_columna_procesados('../datos/facturas/compras-procesadas.xlsx', 2, 'B')

        if gastos_compra is not None:
            gestor.log('Se han calculado ' + str(gastos_compra) + '€ de gastos en compras.')
        else:
            gestor.log('Error: Datos de compras procesadas no válidos.')

    except FileNotFoundError:
        gestor.log('Error: No se ha encontrado el fichero de compras procesadas.')


# lee las nominas y calcula los gastos de personal
def calcular_gastos_personal():
    global gastos_personal
    gastos_personal = 0

    for fichero in os.listdir('../datos/nominas'):
        if fichero != 'plantilla.xlsx':
            try:
                coste_empresa = gestor.calcular_gasto_nomina(fichero)

                # calculamos el gasto mensual, sumamos el gasto y cambiamos de fichero
                gastos_personal += (coste_empresa / 12)
                gestor.log('Añadidos ' + str(round(coste_empresa, 2)) + '€ al costo total.')

            except TypeError:
                gestor.log('Error: Valor no válido.')

    gestor.nwlog('Se han calculado ' + str(round(gastos_personal, 2)) + '€ de gastos en personal.')


# graba los datos calculados en el informe de gerencia
def grabar_informe_generancia():
    try:
        if (ingresos is not None) and (gastos_personal is not None) and (gastos_compra is not None):
            gestor.grabar_informe_mensual('../datos/informe_gerencia/informeGerencia.xlsx', 'c d e f g h i j k l m n',
                                          {6: ingresos, 3: gastos_personal, 4: gastos_compra})
            gestor.log('Se ha grabado el informe de gerencia.')

        else:
            gestor.log('Error: Existen datos a grabar no válidos.')

    except FileNotFoundError:
        gestor.log('Error: No se ha encontrado el fichero de informe de gerencia.')
    except NameError:
        gestor.log('Error: Faltan datos para grabar el informe de gerencia.')


def main():
    calcular_gastos_personal()
    calcular_gastos_compra()
    calcular_ingresos()
    grabar_informe_generancia()


if __name__ == '__main__':
    main()
