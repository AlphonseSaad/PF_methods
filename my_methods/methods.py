

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

def create_variable_selection (result_file_name, element_to_spectates, pf_variable_name):
    
    study_case = app.GetActiveStudyCase()
    elmres = study_case.CreateObject("ElmRes",result_file_name)
    element = element_to_spectates
    variable_name = pf_variable_name
    elmres.AddVariable(element, variable_name)
    elmres.Load()
    return (elmres)


def run_dynamic_simulation(pf_simulation_type, simulation_time, pf_result_file):

    initial_conditions = app.GetFromStudyCase('ComInc')
    initial_conditions.iopt_sim = pf_simulation_type
    initial_conditions.p_resvar = pf_result_file
    initial_conditions.Execute()

    dynamic_simulation = app.GetFromStudyCase("ComSim")
    dynamic_simulation.tstop = simulation_time
    dynamic_simulation.Execute()
        

    
