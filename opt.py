import gurobipy as gp
from gurobipy import GRB
import openpyxl
import numpy as np

alloy_number = 4.7
widgets_number = 3
alloy_p = 0.02
widgets_p = 0.15
widgets_d = 0.12
plants = range(5)
retails = range(8)
warehouses = range(4)
years = range(10)

alloy_price = [alloy_p*1.03**i for i in years]

widgets_price = [widgets_p*1.03**i for i in years]
widgets_discount = [widgets_d*1.03**i for i in years]

plant_cap = [16000,12000,14000,20000,13000]
plant_capacity = [plant_cap for i in years]

operatingcost = [420,380,460,280,340]
operating_cost = [[operatingcost[j]*1.03**i for j in plants] for i in years]


reopening = [190,150,160,100,130]
reopening_cst = [[reopening[j]*1.03**i for j in plants] for i in years]
reopening_cost = np.array(reopening_cst) + np.array(operating_cost)

construction = [2000, 1600, 1800, 900, 1500]
construction_cost = [[construction[j]*1.03**i for j in plants] for i in years]

shutdowm = [170,120,130,80,110]
shutdow_cost = [[shutdowm[j]*1.03**i for j in plants] for i in years]

#### read in data
book = openpyxl.load_workbook("cost assumption.xlsx")

### Demand data
demand_name = book["Demand"]
demand = []
for i in list(demand_name.columns)[1:11]:
    demand_year_data = []
    for cell in i[1:9]:
        demand_year_data.append(cell.value)
    demand.append(demand_year_data)


### Plant shipping data
planttowh_name = book['shipping_cost_W']
planttowh = []
planttowarehouse = []

for i in list(planttowh_name.rows)[1:6]:
    planttowh_year_data = []
    for cell in i[1:5]:
        planttowh_year_data.append(cell.value)
    planttowh.append(planttowh_year_data)

for i in years:
    data = []
    for j in plants:
        data2 = []
        for k in warehouses:
            data2.append(planttowh[j][k]*(1.03)**i)
        data.append(data2)
    planttowarehouse.append(data)


### Warehouse shipping data
whtoretail_name = book['shipping_cost_R']
whtoretail = []
waretoretail = []

for i in list(whtoretail_name.rows)[1:5]:
    whtoretail_year_data = []
    for cell in i[1:9]:
        whtoretail_year_data.append(cell.value)
    whtoretail.append(whtoretail_year_data)

for i in years:
    data3 = []
    for j in warehouses:
        data4 = []
        for k in retails:
            data4.append(whtoretail[j][k]*(1.03)**i)
        data3.append(data4)
    waretoretail.append(data3)

m = gp.Model("Min_cost")

#### Variables

##Binary
open = m.addVars(years,plants,vtype=GRB.BINARY,obj=operating_cost,name= 'operating')
reopen = m.addVars(years,plants,vtype = GRB.BINARY,obj = reopening_cost,name ="open")
shut = m.addVars(years, plants,vtype= GRB.BINARY,obj=shutdow_cost,name='shut')
construc = m.addVars(years,plants,vtype=GRB.BINARY,obj= construction_cost, name= 'construction')

##Continuous
transport_ptow = m.addVars(years,plants,warehouses,vtype=GRB.CONTINUOUS,obj = planttowarehouse ,name ="trans_ptow")
transport_wtor = m.addVars(years,warehouses,retails,vtype=GRB.CONTINUOUS,obj= waretoretail,name=" trans_wtor")
wvars = m.addVars(warehouses,years, vtype =GRB.CONTINUOUS,name = 'wvar')


#### Constraints


m.addConstrs((transport_wtor.sum(year,'*',retail) >= demand[year][retail]
              for retail in retails for year in years ), name="trans_ptow")
# Demand constraint

m.addConstrs((gp.quicksum(transport_ptow.select(year,plant,'*')) <= plant_capacity[year][plant]*open[(year,plant)]
              for plant in plants for year in years), name = 'capacity limitation')
# Plant Capacity constraint

m.addConstrs((construc.sum('*',plant) <= 1 for plant in plants),'construction')
# Construction status constraint

m.addConstrs((transport_ptow.sum(year,'*',warehouse) <= 12000
              for warehouse in warehouses for year in years), name= "warehouse capacity flows in")
# Number of flugels flow in constraint

m.addConstrs((gp.quicksum(transport_wtor.select(year,warehouse,'*')) <= 12000
              for warehouse in warehouses for year in years), name= "warehouse capacity flows out")
# Number of flugels flow out constraint

m.addConstrs((gp.quicksum(transport_ptow.select(year,plant,'*')) <= 60000/alloy_number
              for plant in plants for year in years), name= "alloy")
# Alloy constraint

m.addConstrs((reopen[year,plant]+shut[year,plant]+open[year,plant] == 1
              for plant in plants for year in years), name= "status")

# Plant status constraint

m.addConstrs((wvars.sum(warehouse,year) <= 12000 for year in years for warehouse in warehouses), "inventory constraint")
# Inventory constraint

m.addConstrs(((wvars[warehouse,year-1]+wvars[warehouse,year])/2 <= 4000 for warehouse in warehouses for year in years if year >=1))
# Average inventory constraint

m.addConstrs((transport_ptow.sum(year,'*',warehouse)+wvars[warehouse,year] == transport_wtor.sum(year,warehouse,'*')+wvars[warehouse,year-1]
              for warehouse in warehouses for year in years if year >= 1), "Balance Node")
# For second year and ever after, the balance node

m.addConstrs((transport_ptow.sum(year,'*',warehouse) == transport_wtor.sum(year,warehouse,'*') + wvars[warehouse,year]
              for warehouse in warehouses for year in years if year == 0), "Balance Node")
# For first year, the balance node

m.addConstrs((wvars[warehouse,year] >=0 for warehouse in warehouses for year in years),name='rreaffes')
# Non-negative for inventory

#### Obejective function

obj = (

### Shipping cost
        gp.quicksum(transport_wtor[year,warehouse,retail]*waretoretail[year][warehouse][retail]
                     for warehouse in warehouses for retail in retails for year in years)
        +gp.quicksum(transport_ptow[year,plant,warehouse]*planttowarehouse[year][plant][warehouse]
                     for plant in plants for warehouse in warehouses for year in years)
### Plant cost
        +gp.quicksum(reopen[(year, plant)] * reopening_cost[year][plant]
                     for plant in plants for year in years)
        +gp.quicksum(shut[(year, plant)]* shutdow_cost[year][plant]
                     for plant in plants for year in years)
        +gp.quicksum(open[(year, plant)]* operating_cost[year][plant]
                     for plant in plants for year in years)
### Material fee
        +gp.quicksum(transport_ptow[year,plant,warehouse]*alloy_number*alloy_price[year]
                     for warehouse in warehouses for year in years for plant in plants)
        +gp.quicksum(transport_ptow[year,plant,warehouse]*widgets_number*widgets_price[year]
                     for warehouse in warehouses for year in years for plant in plants)
)

m.setObjective(obj,GRB.MINIMIZE)

m.optimize()

print('\nTotal Costs: %g' % m.objVal)

# Solution (Variable Values):
print('SOLUTION:')

for o in years:
    for i in plants:
        if open[o,i].x > 0.99 or reopen[o,i].x > 0.99:
            print ( 'Plant %s open in year %g' % ((i + 1),(o + 1)) )
            for j in warehouses:
                if transport_ptow[o, i, j].x > -1:
                    print ( 'Transport %g units to warehouse %s in year %s' % ((transport_ptow[o, i, j].x, (j + 1),(o+1))) )
        else:
            print ( 'Plant %s closed in year %g' % ((i + 1),(o + 1)))

for o in years:
    for j in warehouses:
            for k in retails:
                if transport_wtor[o,j,k].x >0 :
                    print ('From warehouse %s transport %s units to retails %g in year %s' % ((j +1),(transport_wtor[o,j,k].x),(k + 1),(o+1)))