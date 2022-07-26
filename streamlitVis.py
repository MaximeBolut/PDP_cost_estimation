import sys
from streamlit import cli as stcli
import streamlit as st
import pandas as pd
import numpy as np
import max_functions as mf
# streamlit ça permet d'afficher en local un page web interactive sur laquelle on visualise les resultats etc


def main():
  
    # st.header() c'est le titre pour la page streamlit
    st.title("LEAD TIME, FTE and COST ESTIMATION")

    # Get the list of sheet in the excel doc (each sheet = one model)
    xl = pd.ExcelFile("PDP Surrogate Testing.xlsx")
    sheet = xl.sheet_names

    # Make it a droplist to select from (with streamlit)
    sheet_selected = st.selectbox(
        'Select a model to do the estimate:',
        (sheet))
    #st.write('You selected:', sheet_selected)

    # load the selected sheet (model)
    df = pd.read_excel("PDP Surrogate Testing.xlsx",
                       sheet_name=sheet_selected)
    df = df.drop(columns=df.columns[8:26])

    # selection du niveau de convergence (on veut un int, 0 pr CL1, 1 pr CL2 etc...)
    # st.radio c'est les boutons pour l'affichage de streamlit
    conv = ['CL1', 'CL2', 'CL3', 'CL4']
    convergence = st.radio(
        "choose the convergence level: ",
        (conv))
    convergence = conv.index(convergence)

    cplx = ["min", "25% scope", "50% scope", "75% scope", "100% full scope"]
    complexity = st.radio(
        "choose the complexity level: ",
        (cplx))
    complexity = cplx.index(complexity)

    # Given the convergence level and complexity a lead time and FTE number can be calculated
    nominal = np.array(df.iloc[8:10, 5])
    selected_conv = np.array(df.iloc[8:10, convergence+2])
    factor = selected_conv/nominal
    res = factor*np.array(df.iloc[14:16, complexity+2])

    st.header("Resulting estimation")

    st.write("Estimated lead time (weeks): ", res[0])
    st.write(" Estimated number of FTE:", res[1])

    # From the estimated FTE we can find the estimated cost
    # fte is the line with the FTEcost
    fte = np.array(df.iloc[19, 2:8])

    # We get the FTEcost that is the closest to the FTE estimate
    col_val = mf.findClosest(fte, res[1])
    col = np.where(fte == col_val)
    col = col[0][0]+2

    # We print the corresponding cost with the detail
    cost = df.iloc[19:24, [1, col]]
    st.write(
        "The cost estimate for the closest corresponding number of estimated FTE is as follow: ")
    st.table(cost)  # static table but nicer visually
    # st.dataframe(cost)  # column can be sorted in ascending or descending order

    st.header("Actual number of FTE available and its impact on cost and lead time")
    actual_fte = st.number_input('Enter the actual Number of FTE available :', min_value=1,
                                 max_value=300, value=5, step=1)

    # We get the FTEcost that is the closest to the FTE entered
    col_val = mf.findClosest(fte, actual_fte)
    col = np.where(fte == col_val)
    col = col[0][0]+2
    # We print the corresponding cost with the detail
    cost = df.iloc[19:24, [1, col]]
    st.write(
        "The cost estimate for the closest corresponding number of actual FTE is as follow: ")
    st.table(cost)  # static table but nicer visually
    # st.dataframe(cost)  # column can be sorted in ascending or descending order

    ratio = actual_fte/res[1]
    leadtime_impact_line = mf.findClosest([0.5, 0.75, 1, 1.5, 2, 3], ratio)
    leadtime_impact_index = [0.5, 0.75, 1,
                             1.5, 2, 3].index(leadtime_impact_line)
    nominal = df.iloc[30, 2]
    leadtime_impact_factor = df.iloc[28:34, 2]/nominal
    st.write("Lead time will be:",
             (leadtime_impact_factor.iloc[leadtime_impact_index])*res[0])


# boot code pour lancer automatiquement streamlit (sinon faut faire dans un terminal: streamlit run 'mon_pgm.py')
if __name__ == '__main__':
    if st._is_running_with_streamlit:
        main()
    else:
        sys.argv = ["streamlit", "run", sys.argv[0]]
        sys.exit(stcli.main())
