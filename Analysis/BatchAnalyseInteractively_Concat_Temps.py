import os, re
import pandas as pd

temp_re = '_[^2-9]+_(\D+)_Temporary.csv$'
region_re = '_[CP]?\d+_([A-Za-z ]+)_'

measures = []
measure_results = {}
analysis_runs = [x for x in os.listdir() if '.' not in x]
for run in analysis_runs:
    temp_path = os.path.join(run, 'Temporary_Files')
    concat_path = os.path.join(run, 'Concatenated_Files')
    temporary_files = [os.path.join(temp_path, x) for x in os.listdir(temp_path) if 'Temporary' in x]
    for file in temporary_files:
        measure = re.findall(temp_re, file)[0]
        measures.append(measure) if measure not in measures else None
    for measure in measures:
        measure_results[measure] = []
    for measure in measures:
        for file in [x for x in temporary_files if measure in x]:
            region = re.findall(region_re, file)[-1]
            try:
                temp_df = pd.read_csv(file)
                temp_df = temp_df.drop(labels=[0], axis=0)
            except:
                continue
            temp_df['region'] = region
            measure_results[measure].append(temp_df)

    for measure, results in measure_results.items():
        if len(results) > 0:
            measure_df = pd.concat(results)
            measure_df.to_csv(os.path.join(concat_path, measure+'.csv'), index=False)