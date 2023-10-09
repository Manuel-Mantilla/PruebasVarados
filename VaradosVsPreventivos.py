import pandas as pd
import os

def read_files(path:str) -> list:
    file_names = []
    try:
        file_names = os.listdir(path)
    except Exception as ex:
        print(f"There was an exception: {ex}")

    return file_names

def create_dataframe(filename:str, sheet:str, columns:list[str]) -> pd.DataFrame:
    dataframe = pd.DataFrame()
    try:
        dataframe = pd.read_excel(filename, sheet_name=sheet, usecols=columns)
    except Exception as ex:
        print(f"There was an exception: {ex}")
    return dataframe

def main(path:str):
    files = read_files(path=path)
    varados = pd.DataFrame()
    preventivos = pd.DataFrame()
    for file in files:
        if file == "Copia Varados editar.xlsx":
            columns = ["Fecha DPV", "BUS - FLOTA", "FRANJA HORA", "AUX NOVEDAD"]
            sheet = "V_CORTO"
            full_path = path + "/" + file
            varados = create_dataframe(full_path, sheet=sheet, columns=columns)
        if file == "PMCEXP.xlsx":
            columns = ["FECHA EJECUCION", "ID", "ORDEN DE TRABAJO", "ESTADO"]
            sheet = "BD_MP"
            full_path = path + "/" + file
            preventivos = create_dataframe(full_path, sheet=sheet, columns=columns)
    if preventivos.empty | varados.empty:
        print("One or both dataframes are empty")
    else:
        preventivos_ejecutados = preventivos[preventivos["ESTADO"] == "EJECUTADO"]
        preventivos_ejecutados.loc[:,"ID"] = preventivos_ejecutados.loc[:,"ID"].str.upper()
        preventivos_ejecutados = preventivos_ejecutados.sort_values(by=["ID", "FECHA EJECUCION"] ,ascending=[True, True]).reset_index(drop=True)

        varados["OT CRUCE"] = 0

        for i in range(0, len(varados)-1):
            fecha = varados.iloc[i]["Fecha DPV"]
            bus = varados.iloc[i]["BUS - FLOTA"]
            resultado = preventivos_ejecutados[(preventivos_ejecutados["ID"] == bus) & (preventivos_ejecutados["FECHA EJECUCION"] >= fecha)].reset_index(drop=True).copy()
            try:
                varados.loc[i, "OT CRUCE"] = resultado.loc[0, "ORDEN DE TRABAJO"]
            except:
                continue

    print(varados.sample(50))

if __name__ == "__main__":
    path = "C:/Users/Administrador/Documents/MANUEL/PROGRAMACION/PYTHON/Pruebas_CEXP/Varados"
    main(path=path)