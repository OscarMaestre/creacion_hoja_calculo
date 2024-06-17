#!/usr/bin/python3
from xlsxwriter import Workbook,utility
from random import choice


PRIMERA_FILA_NOMBRE_ALUMNO=5
COLUMNA_NOMBRE_ALUMNOS=0
LETRA_COLUMNA_NOMBRE_ALUMNOS="A"

COLUMNA_DNI_ALUMNOS=COLUMNA_NOMBRE_ALUMNOS+1
LETRA_COLUMNA_DNI_ALUMNOS="B"

MAX_ALUMNOS=30
TEXTO_AZAR="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
DNI_AZAR="0123456789"

MAX_EXAMENES=4
MAX_TEMAS=7
NOMBRE_HOJA_ALUMNADO='Alumnado'
REF_HOJA_ALUMNADO="$"+NOMBRE_HOJA_ALUMNADO

def get_cadena_azar(cadena, longitud):
    texto=cadena
    devolver=""
    for i in range(1,longitud):
        devolver=devolver+choice(list(texto))
    return devolver

def anadir_encabezado_en_hoja(hoja):
    hoja.write(PRIMERA_FILA_NOMBRE_ALUMNO, COLUMNA_NOMBRE_ALUMNOS, "Nombre")
    hoja.write(PRIMERA_FILA_NOMBRE_ALUMNO, COLUMNA_DNI_ALUMNOS, "DNI")

def crear_hoja_alumnos(libro):
    hoja = libro.add_worksheet(NOMBRE_HOJA_ALUMNADO)
    
    for i in range(1, MAX_ALUMNOS):
        fila=(i+PRIMERA_FILA_NOMBRE_ALUMNO)
        anadir_encabezado_en_hoja(hoja)

        nombre_azar=get_cadena_azar(TEXTO_AZAR, 20)
        hoja.write(fila, COLUMNA_NOMBRE_ALUMNOS, nombre_azar)
        
        dni_azar=get_cadena_azar(DNI_AZAR, 8)
        hoja.write(fila, COLUMNA_DNI_ALUMNOS, dni_azar)
    
def construir_referencia (nombre_hoja, num_fila, letra_columna):
    referencia="="+nombre_hoja+"!"+letra_columna+str(num_fila)
    return referencia


def crear_hojas_examenes(libro):
    for num_tema in range(1, MAX_TEMAS):
        for num_examen in range(1, MAX_EXAMENES):
            nombre=f"Examen {num_examen} tema {num_tema}"
            hoja=libro.add_worksheet(nombre)
            anadir_encabezado_en_hoja(hoja)
            for i in range(1, MAX_ALUMNOS):
                fila=(i+1+PRIMERA_FILA_NOMBRE_ALUMNO)
                
                celda_nombre_alumno="A"+str(fila)
                print(celda_nombre_alumno)
                ref_nombre_alumno=construir_referencia(NOMBRE_HOJA_ALUMNADO,
                                                       fila, LETRA_COLUMNA_NOMBRE_ALUMNOS)
                print(ref_nombre_alumno)
                hoja.write_formula(fila, COLUMNA_NOMBRE_ALUMNOS, ref_nombre_alumno)

                celda_dni_alumno="B"+str(fila)
                ref_dni_alumno=construir_referencia(NOMBRE_HOJA_ALUMNADO,
                                                    fila,
                                                    LETRA_COLUMNA_DNI_ALUMNOS)
                print(ref_dni_alumno)
                hoja.write(fila, COLUMNA_DNI_ALUMNOS, ref_dni_alumno)

def crear_hoja(nombre_archivo="alumnos.xls"):

    with Workbook(nombre_archivo) as libro:
        crear_hoja_alumnos(libro)        
        crear_hojas_examenes(libro)
        

if __name__=="__main__":
    crear_hoja()