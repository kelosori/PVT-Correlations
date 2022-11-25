#%% Import libraries and functions
import xlwings as xw
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from PVT-Correlations.Model.Funciones import Bo
from PVT-Correlations.Model.Funciones import Pb
from PVT-Correlations.Model.Funciones import Rs
from PVT-Correlations.Model.Funciones import uo

#%% Create sheet, variable and data names

# Names for sheets
SHEET_SUMMARY = "Datos"
SHEET_RESULTS = "Resultados"

# Name of columns for distribution definitions
VARIABLES = "Variables"
VALORES = "Valores"
PARAMETROS = "Parametros"
CORRELACION = "Correlacion"

# Name of Data
STOC_VALUES = "df_bo_calculator"
# Result cells # Call range cells from MS Excel
BO_STANDING = "Bo_Standing"
BO_AL_MARHOUN = "Bo_Al_Marhoun"
RS_STANDING = "Rs_Standing"
RS_AL_MARHOUN = "Rs_Al_Marhoun"
PB_STANDING = "Pb_Standing"
PB_AL_MARHOUN = "Pb_Al_Marhoun"
UO_BEAL = "uo_Beal"
UO_GLASO = "uo_Glaso"
VALUES = "Valores"
CORRELACION_S = "Standing"
CORRELACION_AL = "AL_Marhoun"
CORRELACION_B = "Beggs & Robinson"
CORRELACION_G = "Glaso"

#%%
################################
def main():
    wb=xw.Book.caller()
    sheet=wb.sheets[SHEET_SUMMARY]

    df_TVD = sheet[STOC_VALUES].options(pd.DataFrame, index=False, expand="table").value
    input_col_names = df_TVD["Valores"].to_list()
    Rs_value, Yo_value, Yg_value, T_value, P_value, API_value, BASIO, BASIO2 = tuple(
        input_col_names
    )
    inpt_idx = [CORRELACION_S, CORRELACION_AL]
    result_Bo = {}
    result_Pb = {}
    result_Rs = {}
    for col in inpt_idx:
        print(col)
        result_Bo[col] = Bo(col, Rs_value, Yg_value, Yo_value, T_value)
        result_Pb[col] = Pb(col, Rs_value, Yg_value, T_value, API_value, Yo_value)
        result_Rs[col] = Rs(col, P_value, API_value, T_value, Yg_value, Yo_value)
    inpt_idx2 = [CORRELACION_B, CORRELACION_G]

    result_uo = {}
    for col2 in inpt_idx2:
        result_uo[col2] = uo(col2, API_value, T_value)

    PVT_summary_results = [
            result_Bo[CORRELACION_S],
            result_Bo[CORRELACION_AL],
            result_Pb[CORRELACION_S],
            result_Pb[CORRELACION_AL],
            result_Rs[CORRELACION_S],
            result_Rs[CORRELACION_AL],
            result_uo[CORRELACION_B],
            result_uo[CORRELACION_G],
    ]
    sheet[BO_STANDING].options(transpose=True).value = PVT_summary_results
    print(PVT_summary_results)


if __name__ == "__main__":
    xw.Book("Control.xlsm").set_mock_caller()
    main()