import pandas as pd
import numpy as np


worksheet = pd.read_excel(r'PSO20220223.xls', sheet_name='Arcos Comerciales y T.Recorrido', engine="xlrd")
# print(worksheet)

worksheet['trayecto'] = pd.to_numeric(worksheet['trayecto'])

subconjunto_tropt = worksheet[
    ["linea", "sentido", "tipodia", "trayecto", "tr_optimo"]]
subconjunto_trmin = worksheet[
    ["linea", "sentido", "tipodia", "trayecto", "tr_minimo"]]
subconjunto_trmax = worksheet[
    ["linea", "sentido", "tipodia", "trayecto", "tr_maximo", "nodo_1", "nodo_2"]]


tropt = subconjunto_tropt.groupby(["linea", "sentido", "tipodia", "trayecto"]).min()
trmin = subconjunto_trmin.groupby(["linea", "sentido", "tipodia", "trayecto"]).min()
trmax = subconjunto_trmax.groupby(["linea", "sentido", "tipodia", "trayecto"]).max()

shot_tropt = tropt.join(trmin, how="inner").join(trmax, how='inner')
shot_tropt = shot_tropt.rename(columns={"trayecto":"sublinea"})
shot_tropt.to_csv('solution.csv')
print(str(shot_tropt))
