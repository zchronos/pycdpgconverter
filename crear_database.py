# -*- coding: utf-8 -*-
import sqlite3

# Llena los datos
cuenta = " "      # Ejemplo: 303-1234567-0-02
razon_social = " "
abreviatura = " "        # Es un alias, pueden ser las siglas de la empresa


conn = sqlite3.connect('database.db')
c = conn.cursor()

# Create Tables
c.execute('''CREATE TABLE empresas
             (cuenta text, razon_social text, abreviatura text)''')

c.execute('''CREATE TABLE clientes
             (codigo text, ape_pat text, ape_mat text, nombres text)''')

# Larger example that inserts many records at a time
c.execute("INSERT INTO empresas VALUES ('" + cuenta + "', '"+ razon_social + "', '" + abreviatura + "')")

conn.commit()
conn.close()