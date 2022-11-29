import openpyxl
import pandas as pd 
import os

def main():
    pestanas = ["Cache","Host","Infra"]
    filename = input("Ingresar el nombre del archivo: ") + ".xlsx"
    write_excel(filename,pestanas)

def leer_archivos(filename):
    print("Leyendo archivo")
    input_cols =  [0,2,3,4,8]
    path = "C:/Proyectos/CoberturaMonitoreo/test-panda/Input"
    fullpath = os.path.join(path,filename)
    df = pd.read_excel(fullpath,
                    sheet_name="Aplicaciones y tecnologías",
                    header=0,
                    usecols=input_cols) 

    return df

def agregar_filtros(df,cod_adpp,filtro):

    df = df[df["Ambiente"] == "Producción"]
    df = df.drop_duplicates(subset ="Equipo")
    print("Agregando filtros por comando")
    df = df[df["Código de aplicación"] == cod_adpp]
    
    print("dataframe")
    print (df)

    match filtro:
        case "Cache":
            print("switch cache")
            return df[df["Tecnología"].str.contains(filtro,na=False)]
        case "Host":
            return df[df["Equipo"].str.startswith('p')]
        # case "Infra":
        #     return df

    return df



def write_excel(filename,pestanas):
    print("Escribiendo")
    df = leer_archivos(filename)
    cod_adpp = input("Ingresar codigo de aplicacion en filtrar: ")
    path = "C:/Proyectos/CoberturaMonitoreo/test-panda/Output/Exportado.xlsx"
    with pd.ExcelWriter(path) as writer:
        for x in pestanas:
            dataframe = agregar_filtros(df,cod_adpp,x) 
            dataframe.to_excel(writer, sheet_name=x)
        



if __name__ == "__main__":
    main()
    input("\tPROCESO FINALIZADO. Presionar Enter para salir")

