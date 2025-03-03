import random
import sqlite3
import openpyxl
import logging
logging.basicConfig(filename='bingo.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def generar_carton():
    logging.info('Generando carton...')
    carton = []
    for columna in range(5):
        numeros = random.sample(range(1 + columna * 15, 16 + columna * 15), 5)
        carton.append(numeros)
    logging.info('Carton generado.')
    return carton

def crear_base_datos():
    logging.info('Creando base de datos...')
    conexion = sqlite3.connect('bingo.db')
    cursor = conexion.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cartones (
            id INTEGER PRIMARY KEY,
            b1 INTEGER, b2 INTEGER, b3 INTEGER, b4 INTEGER, b5 INTEGER,
            i1 INTEGER, i2 INTEGER, i3 INTEGER, i4 INTEGER, i5 INTEGER,
            n1 INTEGER, n2 INTEGER, n3 INTEGER, n4 INTEGER, n5 INTEGER,
            g1 INTEGER, g2 INTEGER, g3 INTEGER, g4 INTEGER, g5 INTEGER,
            o1 INTEGER, o2 INTEGER, o3 INTEGER, o4 INTEGER, o5 INTEGER
        )
    ''')
    conexion.commit()
    conexion.close()
    logging.info('Base de datos creada.')

def guardar_carton(carton):
    logging.info('Guardando carton...')
    conexion = sqlite3.connect('bingo.db')
    cursor = conexion.cursor()
    cursor.execute('''
        INSERT INTO cartones (
            b1, b2, b3, b4, b5,
            i1, i2, i3, i4, i5,
            n1, n2, n3, n4, n5,
            g1, g2, g3, g4, g5,
            o1, o2, o3, o4, o5
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        carton[0][0], carton[0][1], carton[0][2], carton[0][3], carton[0][4],
        carton[1][0], carton[1][1], carton[1][2], carton[1][3], carton[1][4],
        carton[2][0], carton[2][1], carton[2][2], carton[2][3], carton[2][4],
        carton[3][0], carton[3][1], carton[3][2], carton[3][3], carton[3][4],
        carton[4][0], carton[4][1], carton[4][2], carton[4][3], carton[4][4]
    ))
    conexion.commit()
    conexion.close()
    logging.info('Carton guardado.')

def generar_y_guardar_cartones(cantidad):
    logging.info('Generando y guardando cartones...')
    crear_base_datos()
    for _ in range(cantidad):
        carton = generar_carton()
        guardar_carton(carton)
    logging.info(f'{cantidad} cartones generados y guardados en la base de datos.')

def exportar_cartones_a_excel(cantidad):
    logging.info('Exportando cartones a Excel...')
    conexion = sqlite3.connect('bingo.db')
    cursor = conexion.cursor()
    cursor.execute("SELECT * FROM cartones LIMIT ?", (cantidad,))
    cartones = cursor.fetchall()
    conexion.close()

    libro_excel = openpyxl.Workbook()
    hoja = libro_excel.active
    hoja.append(["ID", "B1", "B2", "B3", "B4", "B5", "I1", "I2", "I3", "I4", "I5", "N1", "N2", "N3", "N4", "N5", "G1", "G2", "G3", "G4", "G5", "O1", "O2", "O3", "O4", "O5"])

    for carton in cartones:
        hoja.append(carton)

    libro_excel.save("cartones_bingo.xlsx")
    logging.info(f'{cantidad} cartones exportados a cartones_bingo.xlsx')

