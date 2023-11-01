# -*- coding: utf-8 -*-
"""
Spyder Editor

author: Astha 12/28/2022
"""
# for all systems within the Biogenic synthesis the following will be calculated
#system_energy_demand         (kwh/ g GM)
#system_total_cost          ($/ g GM)
#system_LCA_impact          (impact/ g GM)


# systems and respective major input include:
    #LCD dismantling, crushing, and slurry preparation
   # process- phytoextraction
#process- pyrolysis and graphitization
#energy demand
#manual labor
#part lifetime and cost
#emmisions
# LCA inputs (includes factors for all TRACI impact categories)


import numpy as np
import pandas as pd
import lhs



#%% LCC/LCA modeling inputs

# # outputs of interest
output_perc_mid = 50
output_perc_low = 5
output_perc_high = 95

# general parameters - import spreadsheet tabs as dataframes
general_assumptions = pd.read_excel('assumptions_biomass.xlsx', sheet_name = 'General', index_col = 'Parameter')
design_assumptions = pd.read_excel('assumptions_biomass.xlsx', sheet_name = 'Design', index_col='Parameter')
LCA_assumptions = pd.read_excel('assumptions_biomass.xlsx', sheet_name = 'LCA', index_col='Parameter')

# number of Monte Carlo runs
n_samples = int(general_assumptions.loc['n_samples','expected'])

# create empty datasets to eventually store data for sensitivity analysis (Spearman's coefficients)
correlation_distributions = np.full((n_samples, n_samples), np.nan)
correlation_parameters = np.full((n_samples, 1), np.nan)
correlation_parameters = correlation_parameters.tolist()


# general inputs
expected_lifetime,correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['expected_lifetime'], correlation_distributions, correlation_parameters, n_samples)
discount_rate, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['discount_rate'], correlation_distributions, correlation_parameters, n_samples)
g_GM_yield, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['g_GM_yield'], correlation_distributions, correlation_parameters, n_samples)
cycle_lifetime, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['cycle_lifetime'], correlation_distributions, correlation_parameters, n_samples)


#LCD dismantling, crushing, and slurry preparation
pulp_density_waste_slurry,correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['pulp_density_waste_slurry'], correlation_distributions, correlation_parameters, n_samples)
indium_concentration_LCD,correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['indium_concentration_LCD'], correlation_distributions, correlation_parameters, n_samples)
indium_in_biomass,correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['indium_in_biomass'], correlation_distributions, correlation_parameters, n_samples)
graphite_synthesis,correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['graphite_synthesis'], correlation_distributions, correlation_parameters, n_samples)
screen_required,correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['screen_required'], correlation_distributions, correlation_parameters, n_samples)
screen_required_for_1g_graphitic_material,correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['screen_required_for_1g_graphitic_material'], correlation_distributions, correlation_parameters, n_samples)
time_crushing_screen, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['time_crushing_screen'], correlation_distributions, correlation_parameters, n_samples)
kg_screen_crushed, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['kg_screen_crushed'], correlation_distributions, correlation_parameters, n_samples)
crushing_time_per_1g_graphitic_material, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['crushing_time_per_1g_graphitic_material'], correlation_distributions, correlation_parameters, n_samples)
acid_required, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['acid_required'], correlation_distributions, correlation_parameters, n_samples)
acid_required_for_1g_graphitic_material, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['acid_required_for_1g_graphitic_material'], correlation_distributions, correlation_parameters, n_samples)
biomass_loss_during_pyrolysis,correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['biomass_loss_during_pyrolysis'], correlation_distributions, correlation_parameters, n_samples)


# process- phytoextraction
drying_time, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['drying_time'], correlation_distributions, correlation_parameters, n_samples)


#process- pyrolysis and graphitization
pyrolysis, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['pyrolysis'], correlation_distributions, correlation_parameters, n_samples)
milling_homegenization, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['milling_homegenization'], correlation_distributions, correlation_parameters, n_samples)
nitrogen_for_pyrolysis, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['nitrogen_for_pyrolysis'], correlation_distributions, correlation_parameters, n_samples)
nitrogen_for_pyrolysis_total, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['nitrogen_for_pyrolysis_total'], correlation_distributions, correlation_parameters, n_samples)


#energy demand
drying_energy, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['drying_energy'], correlation_distributions, correlation_parameters, n_samples)
crushing_energy, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['crushing_energy'], correlation_distributions, correlation_parameters, n_samples)
pyrolysis_energy, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['pyrolysis_energy'], correlation_distributions, correlation_parameters, n_samples)
stirrer_energy, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['stirrer_energy'], correlation_distributions, correlation_parameters, n_samples)
vacuum_oven_energy, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['vacuum_oven_energy'], correlation_distributions, correlation_parameters, n_samples)
crushing_LCD_time, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['crushing_LCD_time'], correlation_distributions, correlation_parameters, n_samples)
slurry_preparation_time, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['slurry_preparation_time'], correlation_distributions, correlation_parameters, n_samples)
plant_exposure_time, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['plant_exposure_time'], correlation_distributions, correlation_parameters, n_samples)


#manual labor
dismantling_time, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['dismantling_time'], correlation_distributions, correlation_parameters, n_samples)
GM_milling_time, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['GM_milling_time'], correlation_distributions, correlation_parameters, n_samples)
harvesting_time, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['harvesting_time'], correlation_distributions, correlation_parameters, n_samples)
graphitization_time, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['graphitization_time'], correlation_distributions, correlation_parameters, n_samples)
plant_rewater_time, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['plant_rewater_time'], correlation_distributions, correlation_parameters, n_samples)

#part lifetime and cost
nitric_acid, correlation_distributions, correlation_parameters = lhs.lhs_distribution(design_assumptions.loc['nitric_acid'], correlation_distributions, correlation_parameters, n_samples)
nitrogen_tank, correlation_distributions, correlation_parameters = lhs.lhs_distribution(design_assumptions.loc['nitrogen_tank'], correlation_distributions, correlation_parameters, n_samples)
biomass, correlation_distributions, correlation_parameters = lhs.lhs_distribution(design_assumptions.loc['biomass'], correlation_distributions, correlation_parameters, n_samples)
electricity_cost, correlation_distributions, correlation_parameters = lhs.lhs_distribution(design_assumptions.loc['electricity_cost'], correlation_distributions, correlation_parameters, n_samples)
tube_furnace, correlation_distributions, correlation_parameters = lhs.lhs_distribution(design_assumptions.loc['tube_furnace'], correlation_distributions, correlation_parameters, n_samples)
frequency_tube_furnace_maintenance, correlation_distributions, correlation_parameters = lhs.lhs_distribution(design_assumptions.loc['frequency_tube_furnace_maintenance'], correlation_distributions, correlation_parameters, n_samples)
manual_labor, correlation_distributions, correlation_parameters = lhs.lhs_distribution(design_assumptions.loc['manual_labor'], correlation_distributions, correlation_parameters, n_samples)
waring_commercial_grinder, correlation_distributions, correlation_parameters = lhs.lhs_distribution(design_assumptions.loc['waring_commercial_grinder'], correlation_distributions, correlation_parameters, n_samples)

#emmisions
CO2, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['CO2'], correlation_distributions, correlation_parameters, n_samples)
Methane, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['Methane'], correlation_distributions, correlation_parameters, n_samples)
H2, correlation_distributions, correlation_parameters = lhs.lhs_distribution(general_assumptions.loc['H2'], correlation_distributions, correlation_parameters, n_samples)

# LCA inputs (includes factors for all TRACI impact categories)
electricity_traci = np.reshape(LCA_assumptions.loc['electricity_med_voltage',:].iloc[2:].to_numpy(dtype=float), (1,-1))
nitric_acid_traci = np.reshape(LCA_assumptions.loc['nitric_acid',:].iloc[2:].to_numpy(dtype=float), (1,-1)) 
carbon_dioxide_traci = np.reshape(LCA_assumptions.loc['carbon_dioxide',:].iloc[2:].to_numpy(dtype=float), (1,-1)) 
nitrogen_traci = np.reshape(LCA_assumptions.loc['nitrogen',:].iloc[2:].to_numpy(dtype=float), (1,-1)) 
heating_element = np.reshape(LCA_assumptions.loc['heating_element',:].iloc[2:].to_numpy(dtype=float), (1,-1)) 
insulation = np.reshape(LCA_assumptions.loc['insulation',:].iloc[2:].to_numpy(dtype=float), (1,-1)) 
tube = np.reshape(LCA_assumptions.loc['tube',:].iloc[2:].to_numpy(dtype=float), (1,-1)) 
exhaust_system = np.reshape(LCA_assumptions.loc['exhaust_system',:].iloc[2:].to_numpy(dtype=float), (1,-1)) 


#%% run design 

#%% LCD dismantling, crushing, and slurry preparation

graphite_synthesis= 1-(biomass_loss_during_pyrolysis/100)                   #g GM/ g fresh biomass
indium_in_1g_graphitic_material = indium_in_biomass/ graphite_synthesis             #mg indium / g GM
screen_required_for_1g_graphitic_material = (pulp_density_waste_slurry* indium_in_1g_graphitic_material )/indium_concentration_LCD      #kg screen
biomass_for_1g_graphitic_material = 1/graphite_synthesis                  #g fresh biomass / g GM
screen_required = (pulp_density_waste_slurry/indium_concentration_LCD) *1000     #kg screen/g indium extarction
crushing_time_per_1g_graphitic_material = (time_crushing_screen/kg_screen_crushed)*screen_required_for_1g_graphitic_material
acid_required_for_1g_graphitic_material= (acid_required/1000000)*indium_in_1g_graphitic_material        #L acid required to synthesize 1g GM and unit conversion mg to kg


#%% phytoextraction

biomass_for_1g_graphitic_material= 1/graphite_synthesis                 #g biomass required for synthesis for 1 g GM


#%%  pyrolysis

nitrogen_for_pyrolysis_1_cycle = nitrogen_for_pyrolysis * 60 * pyrolysis          # L N2 required for pyrolysis for 1 cycle



#%% Energy demand

crushing_energy_cost_1_cycle = crushing_energy * crushing_LCD_time * electricity_cost             #$ for 1 cycle
pyrolysis_energy_cost_1_cycle = pyrolysis_energy * graphitization_time * electricity_cost         #$ for 1 cycle
stirrer_energy_cost_1_cycle = stirrer_energy * slurry_preparation_time * electricity_cost         #$ for 1 cycle
vacuum_oven_energy_cost_1_cycle = vacuum_oven_energy * drying_time * electricity_cost             #$ for 1 cycle

crushing_energy_demand_1_cycle = crushing_energy * crushing_LCD_time              #kwh for 1 cycle
pyrolysis_energy_demand_1_cycle = pyrolysis_energy * graphitization_time          #kwh for 1 cycle
stirrer_energy_demand_1_cycle = stirrer_energy * slurry_preparation_time          #kwh for 1 cycle
vacuum_oven_energy_demand_1_cycle = vacuum_oven_energy * drying_time             #kwh for 1 cycle

#unit process energy demand

#LCD dimantling and slurry preparation
LCD_slurry_energy_demand = crushing_energy_demand_1_cycle + stirrer_energy_demand_1_cycle
print ( "LCD dismantling and slurry preparation energy demand", LCD_slurry_energy_demand)

#phytoextraction
print ("phytoextraction energy demand",vacuum_oven_energy_demand_1_cycle)

#graphitization
print ("graphitization energy demand", pyrolysis_energy_demand_1_cycle)

total_energy_demand_1_cycle = crushing_energy_demand_1_cycle + pyrolysis_energy_demand_1_cycle + stirrer_energy_demand_1_cycle + vacuum_oven_energy_demand_1_cycle
total_energy_demand = total_energy_demand_1_cycle/ g_GM_yield
print ("total energy demand",total_energy_demand)

#total energy demand
total_energy_cost_1_cycle =  crushing_energy_cost_1_cycle +pyrolysis_energy_cost_1_cycle + stirrer_energy_cost_1_cycle + vacuum_oven_energy_cost_1_cycle
print ("total energy cost",total_energy_cost_1_cycle)

total_energy_cost = total_energy_cost_1_cycle / g_GM_yield # $ / g GM



#%% manual labor cost
dismantling_manual_labor_1_cycle = dismantling_time * manual_labor      #$ / 1 cycle
GM_milling_labor_1_cycle = GM_milling_time * manual_labor               #$ / 1 cycle
harvesting_labor_1_cycle = harvesting_time * manual_labor               #$ / 1 cycle
graphitization_labor_1_cycle = graphitization_time * manual_labor       #$ / 1 cycle
plant_rewater_labor_1_cycle = plant_rewater_time * manual_labor         #$ / 1 cycle

#unit process labor cost

#LCD dimantling and slurry preparation
print ( "LCD dismantling and slurry preparation labor cost", dismantling_manual_labor_1_cycle)

#phytoextraction
print ("phytoextraction labor cost",plant_rewater_labor_1_cycle + harvesting_labor_1_cycle)

#graphitization
print ("graphitization labor cost", GM_milling_labor_1_cycle + graphitization_labor_1_cycle)

#total energy demand
total_labor_cost = (dismantling_manual_labor_1_cycle + GM_milling_labor_1_cycle + harvesting_labor_1_cycle + graphitization_labor_1_cycle + plant_rewater_labor_1_cycle) / g_GM_yield # $ / g GM
print ("total labor cost",total_labor_cost)



#%% part and lifetime cost

# one time cost
machinery_cost = tube_furnace + waring_commercial_grinder # $    
maintainance_cost_per_year = frequency_tube_furnace_maintenance # $ 
biomass_cost= biomass * biomass_for_1g_graphitic_material # $ /g GM

#chemical cost
acid_cost = nitric_acid*1000 * acid_required_for_1g_graphitic_material # $ / g GM
nitrogen_cost = nitrogen_for_pyrolysis_1_cycle * nitrogen_tank / g_GM_yield # $ / 1 g GM

print ("chemical cost", acid_cost + nitrogen_cost)

print ("raw material (biomass) cost",biomass_cost)

total_capital_cost = machinery_cost  # $ over lifetime
print ("total capital cost", total_capital_cost)

captial_annualized_cost = total_capital_cost * ((discount_rate * (1+discount_rate)**expected_lifetime) / ((1+discount_rate)**expected_lifetime - 1)) # $ / yr

captial_cost_normalized = captial_annualized_cost * cycle_lifetime / g_GM_yield # $ / g GM

operation_cost_normalized = acid_cost + nitrogen_cost + biomass_cost # $ / g GM
print ("operation cost normalized ", operation_cost_normalized )

total_normalized_cost = total_energy_cost + total_labor_cost + captial_cost_normalized + operation_cost_normalized




#%%LCA

#electricity

crushing_electricity_LCA = crushing_energy_demand_1_cycle * electricity_traci /g_GM_yield # impact / g GM
pyrolysis_electricity_LCA = pyrolysis_energy_demand_1_cycle * electricity_traci/g_GM_yield # impact / g GM
stirrer_electricity_LCA = stirrer_energy_demand_1_cycle * electricity_traci/g_GM_yield # impact / g GM
vacuum_oven_electricity_LCA = vacuum_oven_energy_demand_1_cycle * electricity_traci/g_GM_yield # impact / g GM
total_electricity_LCA = ( crushing_electricity_LCA + pyrolysis_electricity_LCA + stirrer_electricity_LCA + vacuum_oven_electricity_LCA) / g_GM_yield # impacts / g GM

print ("total impact of biosynthesis process due to electricity", total_electricity_LCA)

#emmisions, gas, or chemicals consumed
CO2_impact_LCA = carbon_dioxide_traci * (biomass_for_1g_graphitic_material * CO2/1000)  # impact / g GM
nitric_acid_impact_LCA = nitric_acid_traci/660.98 * acid_required_for_1g_graphitic_material # impact / g GM ; 660.98 mL nitric acid is 1 kg nitric acid; to convert in traci unit
nitrogen_impact_LCA = nitrogen_traci * nitrogen_for_pyrolysis_total / g_GM_yield # impact / g GM

total_impact_furnace_production = tube + exhaust_system + insulation + heating_element
total_impact_furnace_production_annualized = total_impact_furnace_production /expected_lifetime #annualizing the impact of the furnace production
total_impact_furnace_production_normalized = total_impact_furnace_production_annualized * cycle_lifetime / g_GM_yield

print ("CO2 impact", CO2_impact_LCA)
print ("nitrogen imapct", nitrogen_impact_LCA)
print ("chemical impact", nitric_acid_impact_LCA)


total_operation_impact = nitric_acid_impact_LCA + nitrogen_impact_LCA # impact / g GM

total_normalized_impacts = total_electricity_LCA + total_operation_impact + CO2_impact_LCA +total_impact_furnace_production_normalized # impact / g GM



#%% System definition 

unit_processes = ('Total','LCD dismantling, crushing, and slurry preparation', 'phytoextraction', 'pyrolysis', 'Energy demand', 'manual labor cost', 'part and lifetime cost', 'LCA')


system_costs = (total_normalized_cost, total_energy_cost, total_labor_cost, captial_cost_normalized, operation_cost_normalized) # $ / g GM
system_cost_label = ('Total_cost', 'Total energy cost', 'Total labor cost', 'Total capital cost', 'Total operation cost')

system_LCA = (total_normalized_impacts, total_electricity_LCA, total_operation_impact, CO2_impact_LCA, total_impact_furnace_production_normalized) # impact / g GM
system_LCA_label = ('Total_impact', 'Total energy impact', 'Total operation impact', 'Total direct emissions impact', 'total_impact_furnace_production_normalized')

#%%sensitivity analysis

writer = pd.ExcelWriter('GM_synthesis_outputs.xlsx', engine='xlsxwriter')


df_cost = pd.DataFrame({k:v.flatten() for k,v in zip(system_cost_label, system_costs)})

df_impact_GWP = pd.DataFrame({k:v.T[1].flatten() for k,v in zip(system_LCA_label, system_LCA)})

df_impact_AP = pd.DataFrame({k:v.T[0].flatten() for k,v in zip(system_LCA_label, system_LCA)})
df_impact_EFW = pd.DataFrame({k:v.T[2].flatten() for k,v in zip(system_LCA_label, system_LCA)})
df_impact_EP = pd.DataFrame({k:v.T[3].flatten() for k,v in zip(system_LCA_label, system_LCA)})
df_impact_HTC = pd.DataFrame({k:v.T[4].flatten() for k,v in zip(system_LCA_label, system_LCA)})
df_impact_HTNC = pd.DataFrame({k:v.T[5].flatten() for k,v in zip(system_LCA_label, system_LCA)})
df_impact_ODP = pd.DataFrame({k:v.T[6].flatten() for k,v in zip(system_LCA_label, system_LCA)})
df_impact_PMPF = pd.DataFrame({k:v.T[7].flatten() for k,v in zip(system_LCA_label, system_LCA)})
df_impact_MIR = pd.DataFrame({k:v.T[8].flatten() for k,v in zip(system_LCA_label, system_LCA)})





df_cost.to_excel(writer, sheet_name = 'cost analysis')
df_impact_GWP.to_excel(writer, sheet_name = 'LCA GWP analsyis')
df_impact_AP.to_excel(writer, sheet_name = 'LCA AP analsyis')
df_impact_EFW.to_excel(writer, sheet_name = 'LCA EFW analsyis')
df_impact_EP.to_excel(writer, sheet_name = 'LCA EP analsyis')
df_impact_HTC.to_excel(writer, sheet_name = 'LCA HTC analsyis')
df_impact_HTNC.to_excel(writer, sheet_name = 'LCA HTNC analsyis')
df_impact_ODP.to_excel(writer, sheet_name = 'LCA ODP analsyis')
df_impact_PMPF.to_excel(writer, sheet_name = 'LCA PMPF analsyis')
df_impact_MIR.to_excel(writer, sheet_name = 'LCA MIR analsyis')






#%% inputs to pandas dataframe  

all_inputs_name = ('expected_lifetime', 'discount_rate', 'g_GM_yield', 'cycle_lifetime','pulp_density_waste_slurry','indium_concentration_LCD','indium_in_biomass', 
                   'graphite_synthesis','screen_required', 'screen_required_for_1g_graphitic_material', 'time_crushing_screen', 'kg_screen_crushed', 'crushing_time_per_1g_graphitic_material', 
                   'acid_required', 'acid_required_for_1g_graphitic_material', 'biomass_loss_during_pyrolysis', 'drying_time', 'pyrolysis', 'milling_homegenization', 'nitrogen_for_pyrolysis', 
                   'nitrogen_for_pyrolysis_total', 'drying_energy', 'crushing_energy', 'pyrolysis_energy', 'stirrer_energy', 'vacuum_oven_energy', 'crushing_LCD_time', 'slurry_preparation_time', 
                   'plant_exposure_time', 'dismantling_time', 'GM_milling_time', 'harvesting_time', 'graphitization_time', 
                   'plant_rewater_time', 'nitric_acid', 'nitrogen_tank', 'biomass', 'electricity_cost', 'tube_furnace', 'frequency_tube_furnace_maintenance', 'manual_labor', 'waring_commercial_grinder')

all_inputs = (expected_lifetime, discount_rate, g_GM_yield, cycle_lifetime, pulp_density_waste_slurry, indium_concentration_LCD, indium_in_biomass, 
              graphite_synthesis, screen_required, screen_required_for_1g_graphitic_material, time_crushing_screen, kg_screen_crushed, crushing_time_per_1g_graphitic_material, 
              acid_required, acid_required_for_1g_graphitic_material, biomass_loss_during_pyrolysis, drying_time, pyrolysis, milling_homegenization, nitrogen_for_pyrolysis, 
              nitrogen_for_pyrolysis_total, drying_energy, crushing_energy, pyrolysis_energy, stirrer_energy, vacuum_oven_energy, crushing_LCD_time, slurry_preparation_time, 
              plant_exposure_time, dismantling_time, GM_milling_time, harvesting_time, graphitization_time, 
              plant_rewater_time, nitric_acid, nitrogen_tank, biomass, electricity_cost, tube_furnace, frequency_tube_furnace_maintenance, manual_labor, waring_commercial_grinder)

dfinputs = pd.DataFrame({k:v.flatten() for k,v in zip(all_inputs_name, all_inputs)})


df_cost_spearmans = (dfinputs.corrwith(df_cost.Total_cost, method='spearman'))
df_GWP_spearmans = (dfinputs.corrwith(df_impact_GWP.Total_impact, method='spearman'))
df_AP_spearmans = (dfinputs.corrwith(df_impact_AP.Total_impact, method='spearman'))
df_EFW_spearmans = (dfinputs.corrwith(df_impact_EFW.Total_impact, method='spearman'))
df_EP_spearmans = (dfinputs.corrwith(df_impact_EP.Total_impact, method='spearman'))
df_HTC_spearmans = (dfinputs.corrwith(df_impact_HTC.Total_impact, method='spearman'))
df_HTNC_spearmans = (dfinputs.corrwith(df_impact_HTNC.Total_impact, method='spearman'))
df_ODP_spearmans = (dfinputs.corrwith(df_impact_ODP.Total_impact, method='spearman'))
df_PMPF_spearmans = (dfinputs.corrwith(df_impact_PMPF.Total_impact, method='spearman'))
df_MIR_spearmans = (dfinputs.corrwith(df_impact_MIR.Total_impact, method='spearman'))



dfinputs.to_excel(writer, sheet_name= 'inputs')
df_cost_spearmans.to_excel(writer, sheet_name= 'cost_spearman')
df_GWP_spearmans.to_excel(writer, sheet_name= 'GHG_spearman')
df_AP_spearmans.to_excel(writer, sheet_name= 'AP_spearman')
df_EFW_spearmans.to_excel(writer, sheet_name= 'EFW_spearman')
df_EP_spearmans.to_excel(writer, sheet_name= 'EP_spearman')
df_HTC_spearmans.to_excel(writer, sheet_name= 'HTC_spearman')
df_HTNC_spearmans.to_excel(writer, sheet_name= 'HTNC_spearman')
df_ODP_spearmans.to_excel(writer, sheet_name= 'ODP_spearman')
df_PMPF_spearmans.to_excel(writer, sheet_name= 'PMPF_spearman')
df_MIR_spearmans.to_excel(writer, sheet_name= 'MIR_spearman')


#%% ouput data to excel file
writer.save()



