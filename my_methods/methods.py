import pandas as pd
import matplotlib.pyplot as plt
import os

def roll_back(app, version_name):
    project = app.GetActiveProject()
    version = None
    for v in project.GetVersions():
        if v.loc_name == version_name:
            version = v
            break
    project.Deactivate()
    v.Rollback()
    project.Activate()

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

def create_variable_selection (app, result_file_name, element_to_spectates:list, pf_variable_names:list):
    study_case = app.GetActiveStudyCase()
    elmres = study_case.CreateObject("ElmRes",result_file_name)
    
    for element in element_to_spectates:
        variable_name = pf_variable_names
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
    return (export.f_name)

def create_plots(app, file_path, plot_groups=None, start_time=None, end_time=None):
    """
    Reads CSV and plots data. 
    plot_groups: to specify how many plots are needed
    e.x:
    groups = [
    ['Voltage, Magnitude in p.u.', 'Electrical Frequency in p.u.'], 
    ['Electrical Frequency in p.u.']
    ]
    start_time/end_time: specify the range in seconds to plot.
    """
    df = pd.read_csv(file_path, skiprows=1)
    df = df.apply(pd.to_numeric, errors='coerce')
    df = df.dropna(how='all')
    
    time_col = 'Time in s'

    if start_time is not None:
        df = df[df[time_col] >= start_time]
    
    if end_time is not None:
        df = df[df[time_col] <= end_time]
    
    if df.empty:
        print("Warning: No data found within the specified time range.")
        return

    if plot_groups is None:
        plot_groups = [[col] for col in df.columns if col != time_col]

    for group in plot_groups:
        plt.figure(figsize=(10, 5))
        
        for col in group:
            if col in df.columns:
                plt.plot(df[time_col], df[col], label=col)
            else:
                print(f"Warning: Column '{col}' not found in file.")
        
        plt.title(f"Data Plot: {', '.join(group)}")
        plt.xlabel('Time (s)')
        plt.ylabel('Value (p.u.)')
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.legend()
        plt.tight_layout()
        
    plt.show()

# TODO clean code: elemenate copied code
def task_automate(app, study_cases):

    for study_name, study_data in study_cases.items():

        create_study_case(app, study_name)

        fault_3ph_events = [event for event in study_data["events"] if event["event_type"] == "EvtShc"]

        if fault_3ph_events:
            for event in study_data["events"]:
                target_obj = app.GetCalcRelevantObjects(event["event_target_query"])[0]
                target_elm = app.GetCalcRelevantObjects(event["variables_target_query"])[0]
                
                create_fault_events(
                    app,
                    event_type=event["event_type"],
                    event_name=event["event_name"],
                    event_time=event["event_time"],
                    event_target=target_obj,
                    event_action= event["event_action"]  
                )
                # TODO douple res file created.
                res_file = create_variable_selection(
                    app,
                    result_file_name=event["result_file_name"],
                    element_to_spectates=target_elm,
                    pf_variable_names=event["plot_variables"],
                    )
                
                run_dynamic_simulation(
                app,
                pf_simulation_type=event["simulation_type"],
                simulation_time=event["simulation_time"],
                pf_result_file=res_file
                )

                exported_file = export_simulation_results_csv(
                    app,
                    pf_result_file=res_file,
                    file_path=event["exported_file_path"],
                    file_name=event["exported_file_name"],
                    )
                
                create_plots(
                    app,
                    file_path=exported_file, 
                    plot_groups=event["plot_groups"],
                    start_time=event["plot_start_time"],
                    end_time=event["plot_end_time"],
                    )
                
        else:
            for event in study_data["events"]:
                target_obj = app.GetCalcRelevantObjects(event["event_target_query"])[0]
                target_elm = app.GetCalcRelevantObjects(event["variables_target_query"])[0]
                create_simulation_events(
                    app,
                    event_type=event["event_type"],
                    event_name=event["event_name"],
                    event_time=event["event_time"],
                    event_target=target_obj,
                    event_value=event["event_value"],
                    event_variable=event["event_variable"],
                )
                res_file = create_variable_selection(
                    app,
                    result_file_name=event["result_file_name"],
                    element_to_spectates=target_elm,
                    pf_variable_names=event["plot_variables"],
                    )
                
                run_dynamic_simulation(
                app,
                pf_simulation_type=event["simulation_type"],
                simulation_time=event["simulation_time"],
                pf_result_file=res_file
                )

                exported_file = export_simulation_results_csv(
                    app,
                    pf_result_file=res_file,
                    file_path=event["exported_file_path"],
                    file_name=event["exported_file_name"],
                    )
                
                create_plots(
                    app,
                    file_path=exported_file, 
                    plot_groups=event["plot_groups"],
                    start_time=event["plot_start_time"],
                    end_time=event["plot_end_time"],
                    )
