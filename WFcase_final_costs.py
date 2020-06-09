from gurobipy import *
import pandas as pd

xls = "2019 WFU BA Case Competition CaseData.xlsx"

Forecasted_Demand = pd.read_excel(xls, "Forecasted Demand")
Annual_Plant_Capacity = pd.read_excel(xls, "Annual Plant Capacity")
inbound_freight_costs = pd.read_excel(xls, "Inbound Freight Costs")
Handling_Charges = pd.read_excel(xls, "Handling Charges")
Outbound_Ground_Cost = pd.read_excel(xls, "Outbound Ground Cost")
Transit_Time_Ground = pd.read_excel(xls, "Transit Time (Ground)")
Next_Day_Air_Cost = pd.read_excel(xls, "Next Day Air Cost")


def opt_model(min_time, dc_number, fixed_den_pitt, fixed_plant_const):
    m = Model("X")
    m.ModelSense = GRB.MINIMIZE

    # Plant Capacity
    plant_prod = list()
    plant_const = list()
    for i in range(len(Annual_Plant_Capacity)):
        plant_prod.append(m.addVar(vtype=GRB.INTEGER, name=Annual_Plant_Capacity.iloc[i, 0]))
        plant_const.append(m.addConstr(plant_prod[i], GRB.LESS_EQUAL, Annual_Plant_Capacity.iloc[i, 1] * (2 - fixed_plant_const)))

    m.update()

    # Plant to DC
    plant_to_dc_amt = list()
    for i in range(len(inbound_freight_costs)):
        plant_to_dc_amt.append(list())
        plant_to_dc_amt[len(plant_to_dc_amt)-1].append(m.addVar(vtype=GRB.INTEGER, name="Camden, "+inbound_freight_costs.iloc[i, 0]))
        plant_to_dc_amt[len(plant_to_dc_amt)-1].append(m.addVar(vtype=GRB.INTEGER, name="Modesto, "+inbound_freight_costs.iloc[i, 0]))

    plant_to_dc_const = [0, 0]
    for i in plant_to_dc_amt:
        plant_to_dc_const[0] += i[0]
        plant_to_dc_const[1] += i[1]

    m.addConstr(plant_to_dc_const[0], GRB.LESS_EQUAL, plant_prod[0])
    m.addConstr(plant_to_dc_const[1], GRB.LESS_EQUAL, plant_prod[1])

    m.update()

    # Choose DCs
    dc_choice = list()
    for i in range(len(inbound_freight_costs)):
        if fixed_den_pitt:
            if inbound_freight_costs.iloc[i, 0] in ["Denver", "Pittsburgh"]:
                dc_choice.append(1)
            else:
                dc_choice.append(m.addVar(vtype=GRB.BINARY, name=inbound_freight_costs.iloc[i, 0]+"Bin"))
        else:
            dc_choice.append(m.addVar(vtype=GRB.BINARY, name=inbound_freight_costs.iloc[i, 0]+"Bin"))

    total_number_dcs = sum(dc_choice)
    m.addConstr(total_number_dcs, GRB.LESS_EQUAL, dc_number)

    m.update()

    # DC to Customers
    dc_to_cust_amt = list()
    dc_to_cust_const = list()
    for i in range(len(Outbound_Ground_Cost)):
        dc_to_cust_amt.append(list())
        for j in range(1, len(Outbound_Ground_Cost.columns)):
            dc_to_cust_amt[len(dc_to_cust_amt)-1].append(m.addVar(vtype=GRB.INTEGER, name=str(Outbound_Ground_Cost.columns[j])+","+str(Outbound_Ground_Cost.iloc[i, 0])))
        dc_to_cust_const.append(m.addConstr(sum(dc_to_cust_amt[len(dc_to_cust_amt)-1]), GRB.EQUAL, Forecasted_Demand.iloc[i, 1]))

    dc_amt_const = list()
    for i in range(len(dc_to_cust_amt[0])):
        temp = 0
        for j in range(len(dc_to_cust_amt)):
            temp += dc_to_cust_amt[j][i]
        dc_amt_const.append(m.addConstr(temp, GRB.LESS_EQUAL, sum(plant_to_dc_amt[i]) * dc_choice[i]))
    
    m.update()

    # Cost
    OBJ = 0
    # To Plants and Handling
    for i in range(len(plant_to_dc_amt)):
        OBJ += plant_to_dc_amt[i][0] * (inbound_freight_costs.iloc[i, 1] + Handling_Charges.iloc[i, 1])+plant_to_dc_amt[i][1] * (inbound_freight_costs.iloc[i, 2] + Handling_Charges.iloc[i, 1])
    
    # To Customers
    for i in range(len(dc_to_cust_amt)):
        for j in range(len(dc_to_cust_amt[i])):
            if Transit_Time_Ground.iloc[i, j+1] <= min_time:
                OBJ += dc_to_cust_amt[i][j]*Outbound_Ground_Cost.iloc[i, j+1]
            else:
                OBJ += dc_to_cust_amt[i][j]*Next_Day_Air_Cost.iloc[i, j+1]

    m.update()
    m.setObjective(OBJ)
    m.update()

    m.optimize()
    
    optimal_plant_to_dc_amt = list()
    for i in range(len(inbound_freight_costs)):
        optimal_plant_to_dc_amt.append(list())
        optimal_plant_to_dc_amt[len(optimal_plant_to_dc_amt)-1].append(inbound_freight_costs.iloc[i, 0])
        optimal_plant_to_dc_amt[len(optimal_plant_to_dc_amt)-1].append(plant_to_dc_amt[i][0].X)
        optimal_plant_to_dc_amt[len(optimal_plant_to_dc_amt)-1].append(plant_to_dc_amt[i][1].X)
   
    camden_prod = 0
    modesto_prod = 0
    for i in optimal_plant_to_dc_amt:
        camden_prod += i[1]
        modesto_prod += i[2]
        
    return [m.objVal, camden_prod, modesto_prod, optimal_plant_to_dc_amt, fixed_den_pitt, fixed_plant_const]


def fill_solutions_df_2_10():
    to_be_solutions_df_any = pd.DataFrame(index=[5, 4, 3, 2, 1], columns=range(2, 16))
    to_be_solutions_df_den_pitt = pd.DataFrame(index=[5, 4, 3, 2, 1], columns=range(2, 16))
    to_be_solutions_df_any_no_plant_const = pd.DataFrame(index=[5, 4, 3, 2, 1], columns=range(2, 16))
    to_be_solutions_df_den_pitt_no_plant_const = pd.DataFrame(index=[5, 4, 3, 2, 1], columns=range(2, 16))
    min_times = [5, 4, 3, 2, 1]
    for i in range(len(min_times)):
        for j in range(2, 16):
            to_be_solutions_df_any.at[min_times[i], j] = opt_model(min_times[i], j, False, True)
            to_be_solutions_df_den_pitt.at[min_times[i], j] = opt_model(min_times[i], j, True, True)
            to_be_solutions_df_any_no_plant_const.at[min_times[i], j] = opt_model(min_times[i], j, False, False)
            to_be_solutions_df_den_pitt_no_plant_const.at[min_times[i], j] = opt_model(min_times[i], j, True, False)
    return to_be_solutions_df_any, to_be_solutions_df_den_pitt, to_be_solutions_df_any_no_plant_const, to_be_solutions_df_den_pitt_no_plant_const


'''
Create Data for Tableau (melted.csv)
'''


def create_results_for_tableau_csv():
    results = fill_solutions_df_2_10()

    melted = pd.DataFrame(columns=["min_time", "Number of DCs", "DC", "From Camden", "From Modesto", "cost"])

    for h in results:
        for i in [5, 4, 3, 2, 1]:
            for j in range(2, 16):
                min_time = i
                no_dc = j
                c0st = h.loc[i, j][0]
                fixed_den_pitt = h.loc[i, j][4]
                fixed_plant_const = h.loc[i, j][5]
                opt_plant_to_dc = h.loc[i, j][3]
                for k in opt_plant_to_dc:
                    dc = k[0]
                    from_camden = k[1]
                    from_modesto = k[2]
                    data = [min_time, no_dc, dc, from_camden, from_modesto, fixed_den_pitt, fixed_plant_const, c0st]
                    df_temp = pd.DataFrame(columns=["min_time", "Number of DCs", "DC", "From Camden", "From Modesto", "fixed_den_pitt", "fixed_plant_const", "cost"], index=[1])
                    for l in range(len(data)):
                        df_temp.iloc[0, l] = data[l]
                    melted = melted.append(df_temp)

    melted.to_csv('melted.csv')
    create_results_for_excel(melted)


'''
Create Data for Excel to Find the Best Optimal Model (short_melted)
'''


def create_results_for_excel(melted):

    short_melted = pd.DataFrame(columns=["min_time", "Number of DCs", "Camden", "Modesto", "fixed_den_pitt", "fixed_plant_const", "cost", "Pitt?"])
    for i in range(280):
        camden = 0
        modesto = 0
        pitt = False
        for j in range(15):
            camden += melted.iloc[i*15+j, 3]
            modesto += melted.iloc[i*15+j, 4]
        min_time = melted.iloc[i*15, 7]
        no_dcs = melted.iloc[i*15, 5]
        fixed_den_pitt = melted.iloc[i*15, 1]
        fixed_plant_const = melted.iloc[i*15, 2]
        cost = melted.iloc[i*15, 6]
        if melted.iloc[i*15+14, 3]+melted.iloc[i*15+14, 4] > 0:
            pitt = True
        data = [min_time, no_dcs, camden, modesto, fixed_den_pitt, fixed_plant_const, cost, pitt]
        df_temp = pd.DataFrame(columns=["min_time", "Number of DCs", "Camden", "Modesto", "fixed_den_pitt", "fixed_plant_const", "cost", "Pitt?"], index=[1])
        for l in range(len(data)):
            df_temp.iloc[0, l] = data[l]
        short_melted = short_melted.append(df_temp)

    short_melted.to_csv('short_melted.csv')


if __name__ == "__main__":
    create_results_for_tableau_csv()

# QED
