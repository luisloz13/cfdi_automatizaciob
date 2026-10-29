"""Guarda la CIEC de un cliente en el Administrador de credenciales de Windows.

Uso:  python guardar_credencial.py XAXX010101000
"""
import getpass
import sys

import keyring

if len(sys.argv) != 2:
    sys.exit("Uso: python guardar_credencial.py <RFC>")
rfc = sys.argv[1].upper()
keyring.set_password("cfdi_sat", rfc, getpass.getpass(f"CIEC de {rfc}: "))
print("Guardada.")
