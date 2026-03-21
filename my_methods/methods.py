import pandas as pd
import matplotlib.pyplot as plt
import os

def create_study_case(app, study_name):
    study_case_folder = app.GetProjectFolder("study")
    new_study = study_case_folder.CreateObject("IntCase", study_name)
    new_study.Activate()

    network_data_folder = app.GetProjectFolder("netdat")
    grid = network_data_folder.GetContents("*.ElmNet")[0]
    grid.Activate()


def create_operational_scenario(app, scenario_name):
    scenario_folder = app.GetProjectFolder("scen")
    new_scenario = scenario_folder.CreateObject("IntScenario", scenario_name)
    new_scenario.Activate()


def create_simulation_events(app, event_type, event_name, event_time, event_target, event_value, event_variable):

    study = app.GetActiveStudyCase()
    event_folder = study.GetContents("*.IntEvt")

    if not event_folder:
        event_folder = study.CreateObject("IntEvt", "events_folder")
    
    event_folder = study.GetContents("*.IntEvt")[0]
    event = event_folder.CreateObject(event_type, event_name)

    event.time = event_time
    event.p_target = event_target
    event.variable = event_variable
    event.value = event_value


def create_fault_events(app, event_type, event_name, event_time, event_target, event_action):

    study = app.GetActiveStudyCase()
    event_folder = study.GetContents("*.IntEvt")

    if not event_folder:
        event_folder = study.CreateObject("IntEvt", "events_folder")
    
    event_folder = study.GetContents("*.IntEvt")[0]
    event = event_folder.CreateObject(event_type, event_name)

    event.time = event_time
    event.p_target = event_target
    event.i_shc = (event_action) 

def create_variable_selection (app, result_file_name, element_to_spectates, pf_variable_names:list):
    study_case = app.GetActiveStudyCase()
    elmres = study_case.CreateObject("ElmRes",result_file_name)
    element = element_to_spectates

    for variable_name in pf_variable_names:
        elmres.AddVariable(element, variable_name)
 
    elmres.Load()
    return (elmres)


def run_dynamic_simulation(app, pf_simulation_type, simulation_time, pf_result_file):

    initial_conditions = app.GetFromStudyCase('ComInc')
    initial_conditions.iopt_sim = pf_simulation_type
    initial_conditions.p_resvar = pf_result_file
    initial_conditions.Execute()

    dynamic_simulation = app.GetFromStudyCase("ComSim")
    dynamic_simulation.tstop = simulation_time
    dynamic_simulation.Execute()

def export_simulation_results_csv(app, pf_result_file, file_path, file_name):

    study_case = app.GetActiveStudyCase()
    export = study_case.CreateObject("ComRes","Export_res")
    export.pResult = pf_result_file
    export.iopt_exp = 6
    export.f_name = os.path.join(file_path, file_name)
    export.iopt_sep = 1
    export.iopt_head = 1
    export.Execute()
        

    
