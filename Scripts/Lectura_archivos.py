import pandas as pd 
import os

def main():
    df = leer_archivos()
    df = agregar_filtros(df)

    
    #visualizar_datos(df)
    exportar_datos(df)

def leer_archivos():
    print("Leyendo archivo")
    

    input_cols =  [0,2,3,4,8]

    path = "C:/Users/rocha/OneDrive/Documents/Udemy/Python/Ejercicio/Input/"
    filename = input("Ingresar el nombre del archivo: ") + ".xlsx"
    fullpath = os.path.join(path,filename)

    df = pd.read_excel(fullpath,
                    sheet_name="Aplicaciones y tecnologías",
                    header=0,
                    usecols=input_cols) 

    return df

def agregar_filtros(df):

    df = df[df["Ambiente"] == "Producción"]
    df = df.drop_duplicates(subset ="Equipo")
    print("Agregando filtros por comando")

    cod_adpp = input("Ingresar codigo de aplicacion en filtrar: ")
    df = df[df["Código de aplicación"] == cod_adpp]
    
    #Condition por pestaña Host
    df = df[df["Equipo"].str.startswith('p')]
    #Condition por pestaña Cache
    df = df[df["Tecnología"].str.contains('Cache')]

    return df

def exportar_datos(df):
    print("Exportando archivo proceso")

    df.to_excel("C:/Users/rocha/OneDrive/Documents/Udemy/Python/Ejercicio/Output/Exportado.xlsx", 
                index = False,
                sheet_name= "Host")
    df.to_excel("C:/Users/rocha/OneDrive/Documents/Udemy/Python/Ejercicio/Output/Exportado.xlsx", 
                index = False,
                sheet_name= "Cache")
    

if __name__ == "__main__":
    main()
    input("\tPROCESO FINALIZADO. Presionar Enter para salir")


#   def visualizar_datos(df):
#   print("Visualizando los primeros 5 registros")
#
#   df_cols = df.columns
#
#   for col in df_cols:
#       print(df[col].head(5))