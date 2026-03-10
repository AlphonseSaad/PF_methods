

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
    
        

    