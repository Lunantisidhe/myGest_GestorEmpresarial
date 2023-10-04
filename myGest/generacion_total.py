from recursos import gestor
import compras
import nominas
import ventas
import declaracion_iva
import informe_gerencia

__author__ = 'REBECA GONZÁLEZ BALADO'


def main():
    gestor.log('Ejecutando compras.py...')
    compras.main()

    gestor.nwlog('Ejecutando ventas.py...')
    ventas.main()

    gestor.nwlog('Ejecutando nominas.py...')
    nominas.main()

    gestor.nwlog('Ejecutando declaracion_iva.py...')
    declaracion_iva.main()

    gestor.nwlog('Ejecutando informe_gerencia.py...')
    informe_gerencia.main()


if __name__ == '__main__':
    main()
