# WF_case_competition_2019

Requires Gurobi optimization software to run.

The case competition is described in 2019 WFU Business Analytics Case Competition - CASE.pdf, and the data I was provided are in 2019 WFU BA Case Competition CaseData.xlsx and City Coordinates.xlsx.

WFcase_final_costs.py uses Gurobi to determine the optimal production and distribution for each of 260 scenarios, using the number of distribution plants, whether we can sell current distribution plants, and number of days we can promise delivery time.

This creates two files, melted.csv, which is used for Tableau, and short_melted.csv, which is used for Best_Model.xlsx.

Final Viz.twbx allows you to visualize each scenario aginst one another, and how the optimization changes based off of constraints.

WFU Solutions.xlsx allows you to set constraints and variables, so an optimization model can be selected.

WFcase_optimal.py allows you to set variables and create Best_Model.xlsx, which is similar to the original data file, but contains the optimal production and distribution plan.
