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
    
'''
Describe Best Model Giving Excel Constants
'''


def best_opt_model(min_time, dc_number, fixed_den_pitt, fixed_plant_const):
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
        plant_to_dc_amt[len(plant_to_dc_amt)-1].append(m.addVar(vtype=GRB.INTEGER, name="Camden, " + inbound_freight_costs.iloc[i, 0]))
        plant_to_dc_amt[len(plant_to_dc_amt)-1].append(m.addVar(vtype=GRB.INTEGER, name="Modesto, " + inbound_freight_costs.iloc[i, 0]))

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
        dc_amt_const.append(m.addConstr(temp, GRB.LESS_EQUAL, sum(plant_to_dc_amt[i])*dc_choice[i]))
    
    m.update()

    # Cost
    OBJ = 0
    # To Plants and Handling
    for i in range(len(plant_to_dc_amt)):
        OBJ += plant_to_dc_amt[i][0]*(inbound_freight_costs.iloc[i, 1]+Handling_Charges.iloc[i, 1])+plant_to_dc_amt[i][1]*(inbound_freight_costs.iloc[i, 2]+Handling_Charges.iloc[i, 1])
    
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
    
    opt_ground = pd.DataFrame(columns=Outbound_Ground_Cost.columns)
    opt_fly = pd.DataFrame(columns=Outbound_Ground_Cost.columns)
    for i in range(len(Outbound_Ground_Cost)):
        opt_ground = opt_ground.append(pd.DataFrame(columns=Outbound_Ground_Cost.columns, index=[1]))
        opt_fly = opt_fly.append(pd.DataFrame(columns=Outbound_Ground_Cost.columns, index=[1]))
        opt_ground.iloc[i, 0] = Outbound_Ground_Cost.iloc[i, 0]
        opt_fly.iloc[i, 0] = Outbound_Ground_Cost.iloc[i, 0]
        for j in range(15):
            if Transit_Time_Ground.iloc[i, j+1] <= min_time:
                opt_ground.iloc[i, j+1] = dc_to_cust_amt[i][j].X
                opt_fly.iloc[i, j+1] = 0
            if Transit_Time_Ground.iloc[i, j+1] > min_time:
                opt_ground.iloc[i, j+1] = 0
                opt_fly.iloc[i, j+1] = dc_to_cust_amt[i][j].X

    return [m.objVal, camden_prod, modesto_prod, optimal_plant_to_dc_amt, opt_ground, opt_fly]


def best_model_gen(best_max_time, best_dc_number, best_fixed_den_pitt, best_fixed_plant_const):
    best_model = best_opt_model(best_max_time, best_dc_number, True, best_fixed_plant_const)

    best_constants = pd.DataFrame([best_max_time, best_dc_number, best_fixed_den_pitt, best_fixed_plant_const], index=["Max Delivery Time", "DC Number", "Fixed Denver/Pittsburgh", "Fixed Plant Constraint"], columns=["Values"])
    best_objval = pd.DataFrame(best_model[0], columns=["Cost"], index=["Optimal"])
    best_factory_prod = pd.DataFrame([best_model[1], best_model[2]], index=["Camden", "Modesto"], columns=["Production (Pounds)"])
    best_optimal_plant_to_dc_amt = pd.DataFrame(best_model[3], columns=["DC", "Camden", "Modesto"])
    best_opt_ground = best_model[4]
    best_opt_fly = best_model[5]

    writer = pd.ExcelWriter("Best_Model.xlsx")

    best_constants.to_excel(writer, sheet_name="Decisions")
    best_objval.to_excel(writer, sheet_name="Cost", index=False)
    best_factory_prod.to_excel(writer, sheet_name="Plant Production")
    best_optimal_plant_to_dc_amt.to_excel(writer, sheet_name="Plant to DC Amount", index=False)
    best_opt_ground.to_excel(writer, sheet_name="DC to Customer Ground", index=False)
    best_opt_fly.to_excel(writer, sheet_name="DC to Customer Air", index=False)

    writer.save()


run_best_max_time = 3
run_best_dc_number = 2
run_best_fixed_den_pitt = True
run_best_fixed_plant_const = True

if __name__ == "__main__":
    best_model_gen(run_best_max_time, run_best_dc_number, run_best_fixed_den_pitt, run_best_fixed_plant_const)

# QED
